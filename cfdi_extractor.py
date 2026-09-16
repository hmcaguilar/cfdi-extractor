"""
CFDI XML Extractor - Extrae datos de facturas electrónicas (Mexico CFDI 4.0)
y genera reportes en Excel con validaciones.

Uso:
    pip install -r requirements.txt
    python cfdi_extractor.py --input "facturas" --output Reporte_Facturas.xlsx --rfc AAA010101AAA --verificar-sat
"""

import os
import sys
import time
import argparse
import logging
from pathlib import Path
from typing import List, Dict, Tuple, Optional
import xml.etree.ElementTree as ET

import pandas as pd
from openpyxl import load_workbook

try:
    import requests
except ImportError:  # requests es opcional: solo se usa con --verificar-sat
    requests = None


# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


# Namespaces CFDI 4.0
NAMESPACES = {
    'cfdi': 'http://www.sat.gob.mx/cfd/4',
    'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'
}

# Webservice público del SAT para consultar el estado (Vigente/Cancelado)
# de un CFDI. No requiere autenticación. El XML de la factura NO se
# actualiza cuando se cancela, así que esta es la única forma confiable
# de saber si una factura sigue vigente.
SAT_WS_URL = 'https://consultaqr.facturaelectronica.sat.gob.mx/ConsultaCFDIService.svc'
SAT_SOAP_ACTION = 'http://tempuri.org/IConsultaCFDIService/Consulta'
SAT_RESPONSE_NS = {
    'a': 'http://schemas.datacontract.org/2004/07/Sat.Cfdi.Negocio.ConsultaCfdi.Servicio'
}


class CFDIExtractor:
    """Extrae datos de facturas CFDI en XML a DataFrame."""
    
    def __init__(self, input_folder: str):
        """
        Inicializa el extractor.
        
        Args:
            input_folder: Ruta carpeta con archivos XML
        """
        self.input_folder = Path(input_folder)
        if not self.input_folder.exists():
            raise ValueError(f"Carpeta no existe: {input_folder}")

        self.data = []
        self.errors = []
        self.recuperados = []
        self.total_encontrados = 0
    
    def extract_xml_files(self) -> List[Path]:
        """Obtiene lista de archivos XML en la carpeta."""
        xml_files = list(self.input_folder.glob("*.xml"))
        logger.info(f"Se encontraron {len(xml_files)} archivos XML")
        return xml_files
    
    def _extract_uuid(self, root) -> str:
        """Extrae UUID del timbre fiscal."""
        try:
            complemento = root.find('cfdi:Complemento', NAMESPACES)
            if complemento is None:
                return None
            
            timbre_fiscal = complemento.find('tfd:TimbreFiscalDigital', NAMESPACES)
            if timbre_fiscal is None:
                return None
            
            return timbre_fiscal.attrib.get('UUID')
        except Exception as e:
            logger.warning(f"Error extrayendo UUID: {e}")
            return None
    
    def _extract_concepto(self, concepto, comprobante_data: Dict) -> Dict:
        """Extrae datos de un concepto (línea de factura)."""
        return {
            **comprobante_data,
            "ClaveProdServ": concepto.attrib.get('ClaveProdServ'),
            "Cantidad": float(concepto.attrib.get('Cantidad', 0)),
            "ClaveUnidad": concepto.attrib.get('ClaveUnidad'),
            "Unidad": concepto.attrib.get('Unidad'),
            "Descripcion": concepto.attrib.get('Descripcion'),
            "ValorUnitario": float(concepto.attrib.get('ValorUnitario', 0)),
            "Importe": float(concepto.attrib.get('Importe', 0)),
        }
    
    def _parse_xml_robusto(self, xml_path: Path):
        """
        Intenta parsear el XML normalmente. Si falla con "not well-formed",
        casi siempre es porque el archivo declara encoding="UTF-8" en su
        primera línea pero en realidad viene guardado en Windows-1252 /
        Latin-1 — muy común en sistemas de facturación más viejos, sobre
        todo cuando el nombre del emisor/receptor tiene acentos, Ñ, o
        símbolos como $. En vez de descartarlo de una vez, se reintenta
        reinterpretando el encoding antes de darlo por perdido.

        Returns:
            (root, nota): nota es None si parseó a la primera, o un texto
            explicando que se recuperó reinterpretando el encoding.
        """
        try:
            tree = ET.parse(xml_path)
            return tree.getroot(), None
        except ET.ParseError as error_original:
            raw = xml_path.read_bytes()
            if not raw.strip():
                raise error_original  # archivo vacío, no hay nada que recuperar

            for encoding_alterno in ('cp1252', 'latin-1'):
                try:
                    texto = raw.decode(encoding_alterno)
                    root = ET.fromstring(texto.encode('utf-8'))
                    return root, (
                        f"El archivo declara UTF-8 pero venía codificado en "
                        f"{encoding_alterno}; se recuperó reinterpretándolo "
                        f"(error original: {error_original})"
                    )
                except Exception:
                    continue
            raise error_original

    def extract_from_file(self, xml_path: Path) -> bool:
        """
        Extrae datos de un archivo XML.

        Args:
            xml_path: Ruta al archivo XML

        Returns:
            True si fue exitoso, False si hubo error
        """
        try:
            root, nota_recuperacion = self._parse_xml_robusto(xml_path)

            # Datos del comprobante
            comprobante = root.attrib
            
            comprobante_data = {
                "Archivo": xml_path.name,
                "Version": comprobante.get('Version'),
                "Serie": comprobante.get('Serie'),
                "Folio": comprobante.get('Folio'),
                "Fecha": pd.to_datetime(comprobante.get('Fecha')),
                "SubTotal": float(comprobante.get('SubTotal', 0)),
                "Moneda": comprobante.get('Moneda'),
                "Total": float(comprobante.get('Total', 0)),
                "TipoDeComprobante": comprobante.get('TipoDeComprobante'),
                "MetodoPago": comprobante.get('MetodoPago'),
                "LugarExpedicion": comprobante.get('LugarExpedicion'),
            }
            
            # Timbre fiscal
            uuid = self._extract_uuid(root)
            comprobante_data["UUID"] = uuid
            
            # Emisor
            emisor = root.find('cfdi:Emisor', NAMESPACES)
            if emisor is not None:
                comprobante_data.update({
                    "RfcEmisor": emisor.attrib.get('Rfc'),
                    "NombreEmisor": emisor.attrib.get('Nombre'),
                    "RegimenFiscal": emisor.attrib.get('RegimenFiscal'),
                })
            
            # Receptor
            receptor = root.find('cfdi:Receptor', NAMESPACES)
            if receptor is not None:
                comprobante_data.update({
                    "RfcReceptor": receptor.attrib.get('Rfc'),
                    "NombreReceptor": receptor.attrib.get('Nombre'),
                    "UsoCFDI": receptor.attrib.get('UsoCFDI'),
                })
            
            # Conceptos
            conceptos = root.findall('cfdi:Conceptos/cfdi:Concepto', NAMESPACES)

            if not conceptos:
                motivo = (
                    "Sin nodo <Conceptos>: no parece un CFDI de factura normal "
                    "(puede ser un acuse de cancelación, complemento de pago u "
                    "otro tipo de XML mezclado en la carpeta)"
                )
                logger.warning(f"{motivo}: {xml_path.name}")
                self.errors.append(f"{xml_path.name} - {motivo}")
                return False

            for concepto in conceptos:
                concepto_data = self._extract_concepto(concepto, comprobante_data)
                self.data.append(concepto_data)

            if nota_recuperacion:
                logger.warning(f"⚠ Recuperado con otro encoding: {xml_path.name} — {nota_recuperacion}")
                self.recuperados.append(f"{xml_path.name} - {nota_recuperacion}")

            logger.info(f"✓ Extraído: {xml_path.name} ({len(conceptos)} conceptos)")
            return True
            
        except ET.ParseError as e:
            self.errors.append(f"{xml_path.name} - Parse Error: {str(e)}")
            logger.error(f"Error XML en {xml_path.name}: {e}")
            return False
        except Exception as e:
            self.errors.append(f"{xml_path.name} - {str(e)}")
            logger.error(f"Error procesando {xml_path.name}: {e}")
            return False
    
    def process_all(self) -> pd.DataFrame:
        """Procesa todos los archivos XML."""
        xml_files = self.extract_xml_files()
        self.total_encontrados = len(xml_files)

        for xml_file in xml_files:
            self.extract_from_file(xml_file)

        if not self.data:
            logger.error("No se extrajeron datos de los archivos XML")
            return pd.DataFrame()

        df = pd.DataFrame(self.data)

        # Ordenar por fecha
        if 'Fecha' in df.columns:
            df = df.sort_values('Fecha')

        archivos_ok = df['Archivo'].nunique()
        logger.info(f"Total de registros extraídos: {len(df)}")
        logger.info(
            f"Archivos encontrados: {self.total_encontrados} | "
            f"Procesados OK: {archivos_ok} | Con error: {len(self.errors)} "
            f"(encontrados debe ser = procesados OK + con error: "
            f"{self.total_encontrados} {'=' if self.total_encontrados == archivos_ok + len(self.errors) else '≠'} "
            f"{archivos_ok + len(self.errors)})"
        )
        if self.recuperados:
            logger.info(f"De los procesados OK, {len(self.recuperados)} se recuperaron de un problema de codificación")

        if self.errors:
            logger.warning("Errores encontrados:")
            for error in self.errors:
                logger.warning(f"  - {error}")

        return df


class CFDIValidator:
    """Valida datos de facturas."""
    
    def __init__(self, df: pd.DataFrame):
        """
        Inicializa validador.
        
        Args:
            df: DataFrame con datos de facturas
        """
        self.df = df.copy()
        self.validation_errors = []
    
    def validate_rfc_receptor(self, rfc_esperado: str) -> pd.DataFrame:
        """
        Valida que todas las facturas sean para un RFC específico.

        La comparación se normaliza (mayúsculas, sin espacios) para que
        pequeñas diferencias de formato no generen falsos positivos.

        Args:
            rfc_esperado: RFC del receptor esperado

        Returns:
            DataFrame con facturas que no coinciden
        """
        if 'RfcReceptor' not in self.df.columns:
            return pd.DataFrame()

        rfc_normalizado = rfc_esperado.strip().upper()
        rfc_receptor_normalizado = (
            self.df['RfcReceptor'].astype(str).str.strip().str.upper()
        )

        no_coinciden = self.df[rfc_receptor_normalizado != rfc_normalizado]

        if len(no_coinciden) > 0:
            logger.warning(
                f"⚠️  {len(no_coinciden)} líneas no corresponden al RFC {rfc_esperado}"
            )
            self.validation_errors.append(
                f"RFC mismatch: {len(no_coinciden)} líneas"
            )
        else:
            logger.info(f"✓ Todas las facturas corresponden a {rfc_esperado}")

        return no_coinciden

    def validate_duplicates(self) -> pd.DataFrame:
        """
        Identifica facturas (no líneas de concepto) duplicadas por UUID.

        Una factura con varios conceptos genera varias filas con el mismo
        UUID; eso NO es un duplicado. Un duplicado real es el mismo UUID
        (folio fiscal) apareciendo en más de un archivo XML distinto.
        """
        if 'UUID' not in self.df.columns or 'Archivo' not in self.df.columns:
            return pd.DataFrame()

        # Colapsar a una fila por factura (UUID + archivo de origen)
        facturas = self.df[['UUID', 'Archivo']].drop_duplicates()

        # UUIDs que aparecen en más de un archivo distinto
        conteo_archivos = facturas.groupby('UUID')['Archivo'].nunique()
        uuids_duplicados = conteo_archivos[conteo_archivos > 1].index

        duplicados = facturas[facturas['UUID'].isin(uuids_duplicados)].sort_values('UUID')

        if len(duplicados) > 0:
            logger.warning(
                f"⚠️  Se encontraron {len(uuids_duplicados)} UUID(s) repetidos "
                f"en {len(duplicados)} archivos distintos"
            )
            self.validation_errors.append(
                f"Duplicados: {len(uuids_duplicados)} UUID(s) en {len(duplicados)} archivos"
            )
        else:
            logger.info("✓ No hay facturas duplicadas (mismo UUID en distintos archivos)")

        return duplicados
    
    def create_summary(self) -> pd.DataFrame:
        """
        Crea resumen agregado por proveedor/fecha.
        
        Returns:
            DataFrame con resumen
        """
        agg_cols = {col: col for col in self.df.columns 
                   if col in ['Total', 'Cantidad', 'Importe', 'SubTotal']}
        
        if not agg_cols:
            return pd.DataFrame()
        
        # Convertir strings numéricos a float si es necesario
        for col in agg_cols:
            if self.df[col].dtype == 'object':
                self.df[col] = self.df[col].str.replace(r'[\$,]', '', regex=True).astype(float)
        
        agg_dict = {
            'Total': 'first',
            'Cantidad': 'sum',
            'Importe': 'sum' if 'Importe' in self.df.columns else 'first',
        }
        columnas_finales = ['PROVEEDOR', 'FECHA', 'UUID', 'RFC_RECEPTOR',
                            'TOTAL', 'CANTIDAD_TOTAL', 'IMPORTE_TOTAL']

        if 'Moneda' in self.df.columns:
            agg_dict['Moneda'] = 'first'
            columnas_finales.append('MONEDA')

        resumen = self.df.groupby(
            ['NombreEmisor', 'Fecha', 'UUID', 'RfcReceptor']
        ).agg(agg_dict).reset_index()

        resumen.columns = columnas_finales

        logger.info(f"Resumen: {len(resumen)} líneas de proveedor")
        return resumen

    def create_provider_summary(self) -> pd.DataFrame:
        """
        Agrupa las facturas por proveedor (RFC emisor) y calcula, por
        proveedor: número de facturas, importe total facturado y el
        rango de fechas (primera y última factura) que cubren.

        Cuenta a nivel de factura (no de línea de concepto), para no
        inflar el número de facturas ni el importe cuando una factura
        tiene varios conceptos.

        Returns:
            DataFrame con un renglón por proveedor
        """
        columnas_necesarias = {'Archivo', 'RfcEmisor', 'NombreEmisor', 'Total', 'Fecha'}
        faltantes = columnas_necesarias - set(self.df.columns)
        if faltantes:
            logger.warning(f"No se puede generar resumen por proveedor, faltan columnas: {faltantes}")
            return pd.DataFrame()

        # Una fila por factura (no por concepto)
        facturas = self.df[list(columnas_necesarias)].drop_duplicates(subset=['Archivo'])

        resumen = facturas.groupby(['RfcEmisor', 'NombreEmisor']).agg(
            NUM_FACTURAS=('Archivo', 'nunique'),
            IMPORTE_TOTAL=('Total', 'sum'),
            FECHA_PRIMERA=('Fecha', 'min'),
            FECHA_ULTIMA=('Fecha', 'max'),
        ).reset_index()

        resumen = resumen.rename(columns={'RfcEmisor': 'RFC_EMISOR', 'NombreEmisor': 'PROVEEDOR'})
        resumen['FECHA_PRIMERA'] = pd.to_datetime(resumen['FECHA_PRIMERA']).dt.date
        resumen['FECHA_ULTIMA'] = pd.to_datetime(resumen['FECHA_ULTIMA']).dt.date
        resumen = resumen[['PROVEEDOR', 'RFC_EMISOR', 'NUM_FACTURAS', 'IMPORTE_TOTAL',
                           'FECHA_PRIMERA', 'FECHA_ULTIMA']]
        resumen = resumen.sort_values('IMPORTE_TOTAL', ascending=False).reset_index(drop=True)

        logger.info(f"Resumen por proveedor: {len(resumen)} proveedores")
        return resumen


# ---------------------------------------------------------------------------
# Verificación de estado ante el SAT (Vigente / Cancelado)
# ---------------------------------------------------------------------------
#
# El XML de una factura NO cambia cuando el emisor la cancela ante el SAT:
# el archivo que tienes en disco sigue viéndose "normal". La única forma
# confiable de saber si sigue vigente es consultar el webservice público
# del SAT con los datos del comprobante (RFC emisor, RFC receptor, Total,
# UUID). Por eso esto es opcional (--verificar-sat): requiere internet y
# tarda más porque es una consulta por factura.

def _parsear_respuesta_sat(xml_content: bytes) -> Dict:
    """Extrae Estado/EsCancelable/EstatusCancelacion de la respuesta SOAP del SAT."""
    root = ET.fromstring(xml_content)
    return {
        'Estado_SAT': root.findtext('.//a:Estado', namespaces=SAT_RESPONSE_NS) or 'DESCONOCIDO',
        'EsCancelable': root.findtext('.//a:EsCancelable', namespaces=SAT_RESPONSE_NS),
        'EstatusCancelacion': root.findtext('.//a:EstatusCancelacion', namespaces=SAT_RESPONSE_NS) or None,
        'CodigoEstatus_SAT': root.findtext('.//a:CodigoEstatus', namespaces=SAT_RESPONSE_NS),
    }


def consultar_estado_sat(rfc_emisor: str, rfc_receptor: str, total: float,
                        uuid: str, timeout: int = 15, reintentos: int = 2,
                        espera_reintento: float = 2.0) -> Dict:
    """
    Consulta el webservice público del SAT para un CFDI puntual.

    Nunca lanza excepción hacia afuera: si no hay internet, hay timeout,
    o el SAT no responde, regresa Estado_SAT='ERROR_CONSULTA' para que
    el resto del proceso pueda seguir sin interrumpirse.

    En una corrida larga (miles de facturas, una consulta HTTPS por cada
    una) es normal que de vez en cuando una sola consulta falle por un
    problema pasajero de red (por ejemplo, el DNS no resuelve el dominio
    del SAT justo en ese instante). Para no marcar esas facturas como
    'no verificable' por una falla momentánea, se reintenta un par de
    veces antes de darse por vencido.
    """
    if requests is None:
        return {
            'Estado_SAT': 'ERROR_CONSULTA', 'EsCancelable': None,
            'EstatusCancelacion': None,
            'CodigoEstatus_SAT': "Falta instalar 'requests' (pip install requests)",
        }
    if not uuid or not rfc_emisor or not rfc_receptor:
        return {
            'Estado_SAT': 'ERROR_CONSULTA', 'EsCancelable': None,
            'EstatusCancelacion': None, 'CodigoEstatus_SAT': 'Datos incompletos para consultar',
        }

    expresion = f"?re={rfc_emisor}&rr={rfc_receptor}&tt={total:.6f}&id={uuid}"
    body = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" '
        'xmlns:tem="http://tempuri.org/">'
        '<soapenv:Header/><soapenv:Body><tem:Consulta>'
        f'<tem:expresionImpresa><![CDATA[{expresion}]]></tem:expresionImpresa>'
        '</tem:Consulta></soapenv:Body></soapenv:Envelope>'
    )
    headers = {
        'Content-Type': 'text/xml; charset=utf-8',
        'SOAPAction': SAT_SOAP_ACTION,
    }

    intentos_totales = max(1, reintentos + 1)
    ultimo_error = None
    for intento in range(1, intentos_totales + 1):
        try:
            resp = requests.post(SAT_WS_URL, data=body.encode('utf-8'), headers=headers, timeout=timeout)
            resp.raise_for_status()
            return _parsear_respuesta_sat(resp.content)
        except Exception as e:
            ultimo_error = e
            if intento < intentos_totales:
                logger.warning(
                    f"Intento {intento}/{intentos_totales} fallido para UUID {uuid} "
                    f"({e}); reintentando en {espera_reintento}s..."
                )
                time.sleep(espera_reintento)

    logger.warning(
        f"No se pudo consultar el SAT para UUID {uuid} tras {intentos_totales} "
        f"intento(s): {ultimo_error}"
    )
    return {
        'Estado_SAT': 'ERROR_CONSULTA', 'EsCancelable': None,
        'EstatusCancelacion': None, 'CodigoEstatus_SAT': str(ultimo_error)[:150],
    }


def verificar_facturas_sat(df: pd.DataFrame, delay: float = 0.3,
                          timeout: int = 15) -> pd.DataFrame:
    """
    Consulta el estado ante el SAT (Vigente/Cancelado) de cada factura
    única en df (una consulta por Archivo/UUID, no por línea de concepto).

    Args:
        df: DataFrame de detalle (una o más filas por factura)
        delay: segundos de espera entre consultas, para no saturar al SAT
        timeout: segundos de espera por consulta antes de darla por fallida

    Returns:
        DataFrame con Archivo, UUID, Estado_SAT, EsCancelable,
        EstatusCancelacion, CodigoEstatus_SAT (una fila por factura)
    """
    columnas_necesarias = {'Archivo', 'UUID', 'RfcEmisor', 'RfcReceptor', 'Total'}
    faltantes = columnas_necesarias - set(df.columns)
    if faltantes:
        logger.warning(f"No se puede verificar contra el SAT, faltan columnas: {faltantes}")
        return pd.DataFrame()

    facturas = df[list(columnas_necesarias)].drop_duplicates(subset=['Archivo'])
    total_facturas = len(facturas)

    logger.info(f"Consultando estado ante el SAT de {total_facturas} facturas (esto puede tardar)...")

    resultados = []
    for i, (_, factura) in enumerate(facturas.iterrows(), start=1):
        resultado = consultar_estado_sat(
            rfc_emisor=factura['RfcEmisor'],
            rfc_receptor=factura['RfcReceptor'],
            total=factura['Total'],
            uuid=factura['UUID'],
            timeout=timeout,
        )
        resultado['Archivo'] = factura['Archivo']
        resultado['UUID'] = factura['UUID']
        resultados.append(resultado)

        if i % 10 == 0 or i == total_facturas:
            logger.info(f"  Consultadas {i}/{total_facturas} facturas ante el SAT")

        if i < total_facturas:
            time.sleep(delay)

    df_estado = pd.DataFrame(resultados)

    n_cancelado = (df_estado['Estado_SAT'] == 'Cancelado').sum()
    n_vigente = (df_estado['Estado_SAT'] == 'Vigente').sum()
    n_error = (df_estado['Estado_SAT'] == 'ERROR_CONSULTA').sum()

    if n_cancelado > 0:
        logger.warning(f"⚠️  {n_cancelado} factura(s) CANCELADA(S) ante el SAT")
    logger.info(f"✓ Vigentes: {n_vigente} | Canceladas: {n_cancelado} | No verificadas: {n_error}")

    return df_estado


def export_to_excel(df_detalle: pd.DataFrame, df_resumen: pd.DataFrame,
                   output_path: str, rfc_esperado: str = None,
                   no_coinciden: pd.DataFrame = None,
                   duplicados: pd.DataFrame = None,
                   moneda_esperada: str = 'MXN',
                   estado_sat: pd.DataFrame = None,
                   df_proveedores: pd.DataFrame = None,
                   total_archivos_encontrados: int = None,
                   archivos_no_procesados: List[str] = None,
                   archivos_recuperados: List[str] = None):
    """
    Exporta DataFrames a Excel con formato y con las validaciones
    (RFC receptor, duplicados, moneda, estado ante el SAT) visibles
    directamente en el archivo, no solo en el log de consola.

    Args:
        df_detalle: DataFrame con datos detallados
        df_resumen: DataFrame con resumen
        output_path: Ruta del archivo Excel
        rfc_esperado: RFC contra el que se validó (opcional)
        no_coinciden: filas de df_detalle cuyo RfcReceptor no coincide (opcional)
        duplicados: filas (UUID, Archivo) detectadas como factura duplicada (opcional)
        moneda_esperada: moneda contra la que se compara para avisar si hay otra (default MXN)
        estado_sat: resultado de verificar_facturas_sat() (opcional, requiere --verificar-sat)
        df_proveedores: resultado de CFDIValidator.create_provider_summary() (opcional)
        total_archivos_encontrados: cuántos .xml había en la carpeta (CFDIExtractor.total_encontrados)
        archivos_no_procesados: CFDIExtractor.errors — lista de "archivo.xml - motivo"
        archivos_recuperados: CFDIExtractor.recuperados — lista de "archivo.xml - motivo"
    """
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    df_detalle = df_detalle.copy()
    df_resumen = df_resumen.copy() if df_resumen is not None else pd.DataFrame()

    archivos_no_coinciden = set()
    if no_coinciden is not None and not no_coinciden.empty and 'Archivo' in no_coinciden.columns:
        archivos_no_coinciden = set(no_coinciden['Archivo'])

    archivos_duplicados = set()
    if duplicados is not None and not duplicados.empty and 'Archivo' in duplicados.columns:
        archivos_duplicados = set(duplicados['Archivo'])

    if rfc_esperado:
        df_detalle['RFC_Valido'] = df_detalle['Archivo'].apply(
            lambda a: 'NO' if a in archivos_no_coinciden else 'SI'
        )
    df_detalle['Factura_Duplicada'] = df_detalle['Archivo'].apply(
        lambda a: 'SI' if a in archivos_duplicados else 'NO'
    )

    # --- Moneda: no se "invalida", solo se hace visible cuál es (MXN, USD, EUR...) ---
    archivos_moneda_distinta = set()
    desglose_moneda = pd.DataFrame()
    if 'Moneda' in df_detalle.columns:
        moneda_norm = df_detalle['Moneda'].astype(str).str.strip().str.upper()
        if moneda_esperada:
            archivos_moneda_distinta = set(
                df_detalle.loc[moneda_norm != moneda_esperada.strip().upper(), 'Archivo']
            )
        desglose_moneda = (
            df_detalle.assign(_moneda_norm=moneda_norm)
            .drop_duplicates(subset=['Archivo'])
            .groupby('_moneda_norm')['Archivo'].nunique()
            .reset_index()
            .rename(columns={'_moneda_norm': 'Moneda', 'Archivo': 'Facturas'})
            .sort_values('Facturas', ascending=False)
        )

    # --- Estado ante el SAT (Vigente/Cancelado), si se consultó ---
    # Solo se consulta el SAT para facturas con el RFC receptor correcto (ver main()),
    # así que las de RFC distinto quedan sin Estado_SAT tras el merge; se etiquetan
    # explícitamente en vez de dejarlas en blanco, para que quede claro que no es un
    # error de consulta sino que ni se intentaron verificar.
    archivos_cancelados = set()
    archivos_error_sat = set()
    if estado_sat is not None and not estado_sat.empty:
        cols_sat = [c for c in ['Archivo', 'Estado_SAT', 'EsCancelable',
                                 'EstatusCancelacion', 'CodigoEstatus_SAT'] if c in estado_sat.columns]
        df_detalle = df_detalle.merge(estado_sat[cols_sat], on='Archivo', how='left')
        if not df_resumen.empty and 'UUID' in df_resumen.columns and 'UUID' in estado_sat.columns:
            cols_sat_resumen = [c for c in ['UUID', 'Estado_SAT'] if c in estado_sat.columns]
            df_resumen = df_resumen.merge(
                estado_sat[cols_sat_resumen].drop_duplicates(subset=['UUID']),
                on='UUID', how='left'
            )

        if archivos_no_coinciden:
            etiqueta_omitido = 'NO_VERIFICADO (RFC receptor no coincide)'
            mask_omitido = df_detalle['Archivo'].isin(archivos_no_coinciden) & df_detalle['Estado_SAT'].isna()
            df_detalle.loc[mask_omitido, 'Estado_SAT'] = etiqueta_omitido
            if not df_resumen.empty and 'Estado_SAT' in df_resumen.columns and 'RFC_RECEPTOR' in df_resumen.columns:
                mask_omitido_resumen = (
                    (df_resumen['RFC_RECEPTOR'].astype(str).str.strip().str.upper()
                     != (rfc_esperado or '').strip().upper())
                    & df_resumen['Estado_SAT'].isna()
                )
                df_resumen.loc[mask_omitido_resumen, 'Estado_SAT'] = etiqueta_omitido

        archivos_cancelados = set(estado_sat.loc[estado_sat['Estado_SAT'] == 'Cancelado', 'Archivo'])
        archivos_error_sat = set(estado_sat.loc[estado_sat['Estado_SAT'] == 'ERROR_CONSULTA', 'Archivo'])

    # --- Archivos que ni siquiera se lograron procesar (parse error, sin
    # Conceptos, etc.) — antes solo se veían en el log de consola ---
    def _parsear_lista_motivos(items):
        filas = []
        for item in items or []:
            if ' - ' in item:
                archivo, motivo = item.split(' - ', 1)
            else:
                archivo, motivo = item, ''
            filas.append({'Archivo': archivo, 'Motivo': motivo})
        return pd.DataFrame(filas)

    df_no_procesados = _parsear_lista_motivos(archivos_no_procesados)
    df_recuperados = _parsear_lista_motivos(archivos_recuperados)

    # Hoja VALIDACION: resumen de auditoría, legible sin ver la consola
    total_facturas = df_detalle['Archivo'].nunique()
    filas_validacion = []
    if total_archivos_encontrados is not None:
        filas_validacion.append({
            'Chequeo': 'Archivos .xml encontrados en la carpeta',
            'Resultado': total_archivos_encontrados,
        })
        filas_validacion.append({
            'Chequeo': 'De esos, NO se pudieron procesar (ver hoja ARCHIVOS_NO_PROCESADOS)',
            'Resultado': len(df_no_procesados),
        })
    filas_validacion.append(
        {'Chequeo': 'Total de archivos XML procesados', 'Resultado': total_facturas}
    )
    if rfc_esperado:
        filas_validacion.append({
            'Chequeo': f'Facturas con RFC receptor = {rfc_esperado}',
            'Resultado': total_facturas - len(archivos_no_coinciden),
        })
        filas_validacion.append({
            'Chequeo': 'Facturas con RFC receptor DISTINTO (revisar)',
            'Resultado': len(archivos_no_coinciden),
        })
    filas_validacion.append({
        'Chequeo': 'Facturas con UUID duplicado en otro archivo (revisar)',
        'Resultado': len(archivos_duplicados),
    })
    if 'Moneda' in df_detalle.columns and moneda_esperada:
        filas_validacion.append({
            'Chequeo': f'Facturas con moneda distinta de {moneda_esperada.upper()} (revisar tipo de cambio)',
            'Resultado': len(archivos_moneda_distinta),
        })
    if estado_sat is not None and not estado_sat.empty:
        if archivos_no_coinciden:
            filas_validacion.append({
                'Chequeo': 'Facturas omitidas de la verificación SAT (RFC receptor no coincide)',
                'Resultado': len(archivos_no_coinciden),
            })
        filas_validacion.append({
            'Chequeo': 'De las facturas con RFC correcto: CANCELADAS ante el SAT (revisar)',
            'Resultado': len(archivos_cancelados),
        })
        filas_validacion.append({
            'Chequeo': 'De las facturas con RFC correcto: no se pudieron verificar ante el SAT',
            'Resultado': len(archivos_error_sat),
        })
    if not df_recuperados.empty:
        filas_validacion.append({
            'Chequeo': 'De los procesados, venían con codificación distinta a UTF-8 (recuperados)',
            'Resultado': len(df_recuperados),
        })
    df_validacion = pd.DataFrame(filas_validacion)

    detalle_problemas = pd.DataFrame()
    archivos_problema = archivos_no_coinciden | archivos_duplicados | archivos_cancelados
    if archivos_problema:
        cols = [c for c in ['Archivo', 'NombreEmisor', 'RfcEmisor', 'RfcReceptor',
                             'UUID', 'Total', 'Moneda', 'RFC_Valido', 'Factura_Duplicada',
                             'Estado_SAT']
                if c in df_detalle.columns]
        detalle_problemas = (
            df_detalle[df_detalle['Archivo'].isin(archivos_problema)][cols]
            .drop_duplicates(subset=['Archivo'])
            .reset_index(drop=True)
        )

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Hoja resumen (primero, es lo que se quiere ver de entrada)
        if not df_resumen.empty:
            df_resumen.to_excel(writer, sheet_name='RESUMEN', index=False)

        # Hoja por proveedor: # de facturas, importe total y rango de fechas
        if df_proveedores is not None and not df_proveedores.empty:
            df_proveedores.to_excel(writer, sheet_name='PROVEEDORES', index=False)

        # Hoja validación
        df_validacion.to_excel(writer, sheet_name='VALIDACION', index=False)
        fila_siguiente = len(df_validacion) + 3
        if not desglose_moneda.empty:
            desglose_moneda.to_excel(
                writer, sheet_name='VALIDACION', index=False, startrow=fila_siguiente
            )
            fila_siguiente += len(desglose_moneda) + 3
        if not detalle_problemas.empty:
            detalle_problemas.to_excel(
                writer, sheet_name='VALIDACION', index=False, startrow=fila_siguiente
            )

        # Hoja detalle
        df_detalle.to_excel(writer, sheet_name='DETALLE', index=False)

        # Hoja de archivos que ni siquiera se lograron procesar, y de los
        # que sí se procesaron pero venían con codificación rara
        if not df_no_procesados.empty:
            df_no_procesados.to_excel(writer, sheet_name='ARCHIVOS_NO_PROCESADOS', index=False)
            fila_siguiente_np = len(df_no_procesados) + 3
            if not df_recuperados.empty:
                pd.DataFrame([{'Archivo': 'RECUPERADOS (sí se procesaron, solo un aviso):', 'Motivo': ''}]).to_excel(
                    writer, sheet_name='ARCHIVOS_NO_PROCESADOS', index=False,
                    header=False, startrow=fila_siguiente_np
                )
                df_recuperados.to_excel(
                    writer, sheet_name='ARCHIVOS_NO_PROCESADOS', index=False,
                    startrow=fila_siguiente_np + 1
                )
        elif not df_recuperados.empty:
            df_recuperados.to_excel(writer, sheet_name='ARCHIVOS_NO_PROCESADOS', index=False)

        # --- Formato ---
        header_font = Font(bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='305496', end_color='305496', fill_type='solid')
        fill_rfc_dup = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')  # rosa
        fill_cancelada = PatternFill(start_color='FF6666', end_color='FF6666', fill_type='solid')  # rojo fuerte
        fill_moneda = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')  # amarillo
        fill_sin_verificar = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')  # gris

        for sheet_name in writer.book.sheetnames:
            ws = writer.book[sheet_name]
            if ws.max_row == 0:
                continue
            for cell in ws[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center')
            ws.freeze_panes = 'A2'
            if ws.max_row > 1:
                ws.auto_filter.ref = ws.dimensions
            # Ancho de columna aproximado según contenido
            for col_cells in ws.columns:
                length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
                col_letter = get_column_letter(col_cells[0].column)
                ws.column_dimensions[col_letter].width = min(max(length + 2, 10), 45)

        # Resaltar en DETALLE las filas con problema de validación.
        # Prioridad: cancelada (rojo fuerte) > RFC/duplicado (rosa) > moneda distinta (amarillo)
        ws_detalle = writer.book['DETALLE']
        headers = [c.value for c in ws_detalle[1]]
        col_archivo = headers.index('Archivo') + 1
        col_rfc = headers.index('RFC_Valido') + 1 if 'RFC_Valido' in headers else None
        col_dup = headers.index('Factura_Duplicada') + 1 if 'Factura_Duplicada' in headers else None
        for row in ws_detalle.iter_rows(min_row=2):
            archivo = row[col_archivo - 1].value
            es_rfc_dup = (
                (col_rfc and row[col_rfc - 1].value == 'NO') or
                (col_dup and row[col_dup - 1].value == 'SI')
            )
            if archivo in archivos_cancelados:
                fill = fill_cancelada
            elif es_rfc_dup:
                fill = fill_rfc_dup
            elif archivo in archivos_moneda_distinta:
                fill = fill_moneda
            elif archivo in archivos_error_sat:
                fill = fill_sin_verificar
            else:
                fill = None
            if fill:
                for cell in row:
                    cell.fill = fill

    logger.info(f"✓ Archivo exportado: {output_path}")


def main():
    """Función principal."""
    parser = argparse.ArgumentParser(
        description='Extrae datos de facturas CFDI en XML a Excel',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Ejemplos de uso:
  python cfdi_extractor.py --input ./facturas --output resultado.xlsx
  python cfdi_extractor.py --input ./facturas --output resultado.xlsx --rfc ABC123456789
        '''
    )
    
    parser.add_argument(
        '--input', '-i',
        required=True,
        help='Carpeta con archivos XML'
    )
    parser.add_argument(
        '--output', '-o',
        default='output.xlsx',
        help='Archivo Excel de salida (default: output.xlsx)'
    )
    parser.add_argument(
        '--rfc', '-r',
        help='RFC del receptor para validación (opcional)'
    )
    parser.add_argument(
        '--no-summary',
        action='store_true',
        help='No generar hoja de resumen'
    )
    parser.add_argument(
        '--moneda',
        default='MXN',
        help='Moneda esperada; las facturas con otra moneda se marcan para revisar (default: MXN)'
    )
    parser.add_argument(
        '--verificar-sat',
        action='store_true',
        help=(
            'Consulta el webservice público del SAT para saber si cada factura '
            'sigue Vigente o está Cancelada. Requiere internet y tarda más '
            '(una consulta por factura).'
        )
    )
    parser.add_argument(
        '--sat-delay',
        type=float,
        default=0.3,
        help='Segundos de espera entre consultas al SAT, para no saturarlo (default: 0.3)'
    )

    args = parser.parse_args()
    
    try:
        # Extraer
        logger.info("=" * 60)
        logger.info("INICIANDO EXTRACCIÓN DE CFDI")
        logger.info("=" * 60)
        
        extractor = CFDIExtractor(args.input)
        df = extractor.process_all()
        
        if df.empty:
            logger.error("No se extrajeron datos. Abortando.")
            return 1
        
        # Validar
        logger.info("\n" + "=" * 60)
        logger.info("VALIDACIÓN")
        logger.info("=" * 60)
        
        validator = CFDIValidator(df)

        no_coinciden = pd.DataFrame()
        if args.rfc:
            no_coinciden = validator.validate_rfc_receptor(args.rfc)
            if len(no_coinciden) > 0:
                logger.warning("\nFacturas que no corresponden:")
                logger.warning(no_coinciden[['NombreEmisor', 'RfcReceptor', 'Total']].to_string())

        duplicados = validator.validate_duplicates()

        # Verificación de cancelación ante el SAT (opcional, requiere internet).
        # Solo se consulta el SAT para las facturas con el RFC receptor correcto;
        # las que no corresponden (RFC distinto) se omiten, no vale la pena
        # gastar la consulta en una factura que de entrada no es tuya.
        estado_sat = pd.DataFrame()
        if args.verificar_sat:
            logger.info("\n" + "=" * 60)
            logger.info("VERIFICANDO ESTADO ANTE EL SAT (Vigente/Cancelado)")
            logger.info("=" * 60)

            if args.rfc and len(no_coinciden) > 0:
                archivos_rfc_incorrecto = set(no_coinciden['Archivo'])
                df_para_sat = df[~df['Archivo'].isin(archivos_rfc_incorrecto)]
                logger.info(
                    f"Se omiten {len(archivos_rfc_incorrecto)} factura(s) con RFC receptor "
                    f"distinto de {args.rfc}; solo se consulta el SAT para las que sí corresponden."
                )
            else:
                df_para_sat = df

            estado_sat = verificar_facturas_sat(df_para_sat, delay=args.sat_delay)

        # Generar resumen
        if not args.no_summary:
            resumen = validator.create_summary()
        else:
            resumen = pd.DataFrame()

        # Resumen por proveedor: # de facturas, importe total, rango de fechas
        proveedores = validator.create_provider_summary()

        # Exportar
        logger.info("\n" + "=" * 60)
        logger.info("EXPORTANDO A EXCEL")
        logger.info("=" * 60)

        export_to_excel(
            df, resumen, args.output,
            rfc_esperado=args.rfc,
            no_coinciden=no_coinciden,
            duplicados=duplicados,
            moneda_esperada=args.moneda,
            estado_sat=estado_sat,
            df_proveedores=proveedores,
            total_archivos_encontrados=extractor.total_encontrados,
            archivos_no_procesados=extractor.errors,
            archivos_recuperados=extractor.recuperados,
        )
        
        logger.info("\n✓ PROCESO COMPLETADO EXITOSAMENTE")
        return 0
        
    except Exception as e:
        logger.error(f"Error crítico: {e}", exc_info=True)
        return 1


if __name__ == '__main__':
    sys.exit(main())
