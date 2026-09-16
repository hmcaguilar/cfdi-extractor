# CFDI Extractor

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

**Extrae datos de comprobantes fiscal digital de Mexico (CFDI 4.0) en formato XML y genera reportes en Excel con validaciones automáticas.**

## 🎯 Problema que resuelve

Procesar cientos de facturas XML manualmente es:
- ⏰ Tedioso y consume horas de trabajo
- ❌ Propenso a errores en transcripción
- 🔍 Difícil de auditar

Este proyecto **automatiza completamente** la extracción y validación de facturas.

## ✨ Características

- ✅ **Extracción masiva** de datos CFDI 4.0 (namespaces correctos)
- ✅ **Validación automática** de RFC del receptor
- ✅ **Detección de duplicados** por UUID
- ✅ **Resumen agregado** por proveedor/fecha
- ✅ **Logging estructurado** para auditoria
- ✅ **Manejo robusto de errores** (continúa aunque falle un XML)
- ✅ **Exportación a Excel** con múltiples hojas

## 📦 Instalación

### Requisitos previos
- Python 3.8+
- pip

### Instalación rápida

```bash
# Clonar el repositorio
git clone https://github.com/tu-usuario/cfdi-extractor.git
cd cfdi-extractor

# Instalar dependencias
pip install -r requirements.txt
```

## 🚀 Uso rápido

```bash
# Extracción simple
python cfdi_extractor.py --input ./facturas --output resultado.xlsx

# Con validación de RFC
python cfdi_extractor.py --input ./facturas --output resultado.xlsx --rfc ABC123456789

# Sin generar resumen
python cfdi_extractor.py --input ./facturas --output resultado.xlsx --no-summary
```

## 📖 Ejemplos detallados

### Ejemplo 1: Procesar carpeta de facturas

```bash
python cfdi_extractor.py \
  --input "C:/Documentos/Facturas" \
  --output "resultados/reporte_enero.xlsx"
```

**Salida esperada:**
```
2024-01-15 10:30:45 - INFO - Se encontraron 150 archivos XML
2024-01-15 10:30:45 - INFO - ✓ Extraído: factura_001.xml (3 conceptos)
...
2024-01-15 10:31:12 - INFO - Total de registros extraídos: 450
2024-01-15 10:31:12 - INFO - ✓ Archivo exportado: resultados/reporte_enero.xlsx
```

### Ejemplo 2: Validar que todas las facturas sean para tu empresa

```bash
python cfdi_extractor.py \
  --input "./facturas" \
  --output "reporte.xlsx" \
  --rfc Tu_RFC
```

Si hay facturas que no corresponden:
```
2024-01-15 10:31:10 - WARNING - ⚠️ 5 facturas no corresponden al RFC AAA010101AAA
```

Se incluirán en una columna `validation_error` para que revises manualmente.

### Ejemplo 3: Validar moneda y cancelación ante el SAT

```bash
python cfdi_extractor.py \
  --input "./facturas" \
  --output "reporte.xlsx" \
  --rfc AAA010101AAA \
  --moneda MXN \
  --verificar-sat
```

- `--moneda` no rechaza nada, solo marca (columna `Moneda` + hoja VALIDACION) las
  facturas que no estén en la moneda esperada, para que revises el tipo de cambio.
- `--verificar-sat` consulta el webservice público del SAT (uno por factura) para
  saber si cada CFDI sigue **Vigente** o está **Cancelado**. Esto es necesario porque
  el XML no cambia cuando el emisor cancela la factura — la única forma confiable de
  saberlo es preguntarle al SAT directamente. Requiere internet y tarda más (por
  default hay una pausa de 0.3s entre consultas, ajustable con `--sat-delay`).
- Si usas `--rfc` junto con `--verificar-sat`, las facturas cuyo RFC receptor **no**
  coincide se omiten de la consulta al SAT (no tiene caso gastar una consulta en una
  factura que de entrada no es tuya). En DETALLE/RESUMEN esas filas quedan con
  `Estado_SAT = NO_VERIFICADO (RFC receptor no coincide)`, y VALIDACION cuenta por
  separado cuántas se omitieron vs. cuántas de las que sí corresponden están
  vigentes/canceladas. Sin `--rfc`, se verifican todas.

### Ejemplo 4: Python script

```python
from cfdi_extractor import CFDIExtractor, CFDIValidator, export_to_excel

# Extraer
extractor = CFDIExtractor('./facturas')
df = extractor.process_all()

# Validar
validator = CFDIValidator(df)
no_coinciden = validator.validate_rfc_receptor('AAA010101AAA')
resumen = validator.create_summary()
proveedores = validator.create_provider_summary()

# Exportar
export_to_excel(df, resumen, 'output.xlsx', df_proveedores=proveedores)
```

## 📊 Estructura del Excel generado

### Hoja "DETALLE"
Contiene todas las líneas extraídas con:

| Columna | Descripción |
|---------|------------|
| `Archivo` | Nombre del XML original |
| `NombreEmisor` | Empresa que emite la factura |
| `RfcEmisor` | RFC del emisor |
| `RfcReceptor` | RFC de quien recibe |
| `Fecha` | Fecha de la factura |
| `UUID` | Folio fiscal único |
| `Total` | Monto total |
| `Descripcion` | Concepto/producto |
| `Cantidad` | Unidades |
| ... | 20+ columnas más |

### Hoja "RESUMEN"
Resumen agregado (una fila por factura):

| PROVEEDOR | FECHA | UUID | TOTAL | MONEDA | CANTIDAD_TOTAL |
|-----------|-------|------|-------|--------|---|
| PROVEEDOR EJEMPLO UNO SA DE CV | 2024-01-13 | a1b2c3d4-... | $10,520.00 | MXN | 400.0 |
| PROVEEDOR XYZ | 2024-01-14 | e5f6a7b8-... | $5,250.50 | MXN | 150.0 |

### Hoja "PROVEEDORES"
Una fila por proveedor (agrupado por RFC emisor), pensada para ver de un vistazo con
quién se concentra el gasto: número de facturas, importe total facturado y el rango
de fechas que cubren. Cuenta a nivel de factura, no de línea de concepto, para no
inflar el número ni el importe. Ordenado de mayor a menor importe.

| PROVEEDOR | RFC_EMISOR | NUM_FACTURAS | IMPORTE_TOTAL | FECHA_PRIMERA | FECHA_ULTIMA |
|-----------|-----------|---|---|---|---|
| CONSTRUCTORA EJEMPLO SA DE CV | CEJ010101AA1 | 11 | $850,000.00 | 2023-08-04 | 2023-12-07 |
| PROVEEDOR PERSONA FISICA EJEMPLO | BBB020202BBB | 5 | $320,500.00 | 2023-08-08 | 2023-09-13 |

### Hoja "VALIDACION"
Resumen de auditoría legible sin ver la consola: cuántas facturas coinciden con el
RFC esperado, cuántas están duplicadas (mismo UUID en más de un archivo), desglose
por moneda y, si usaste `--verificar-sat`, cuántas están Vigentes/Canceladas ante el
SAT. Debajo trae el detalle de los archivos con algo que revisar.

En la hoja DETALLE, las filas con problema se resaltan: rojo fuerte si la factura
está cancelada ante el SAT, rosa si el RFC no coincide o está duplicada, amarillo si
la moneda es distinta a la esperada, gris si no se pudo verificar contra el SAT.

## 🔧 Opciones avanzadas

### Variables de entorno

```bash
# Log level (DEBUG, INFO, WARNING, ERROR)
export LOG_LEVEL=DEBUG
python cfdi_extractor.py --input ./facturas
```

### Configuración con YAML (futuro)

```yaml
# config.yaml
input_folder: "./facturas"
output_file: "reporte.xlsx"
validaciones:
  rfc_receptor: "AAA010101AAA"
  check_duplicates: true
  check_moneda: "MXN"
```

## 🐛 Troubleshooting

### Error: "Carpeta no existe"
```
ValueError: Carpeta no existe: ./facturas
```
**Solución:** Verifica que la ruta sea correcta
```bash
# En Windows
python cfdi_extractor.py --input "C:\Users\tu_usuario\Documentos\Facturas"

# En Linux/Mac
python cfdi_extractor.py --input ~/Documentos/Facturas
```

### Error: "No se encontraron conceptos"
Significa que el XML tiene estructura válida pero no contiene líneas (`Conceptos`). Esto es raro pero ocurre con facturas anuladas.

**Solución:** Revisa el XML manualmente o filtra estos archivos antes de procesarlos.

### Los montos aparecen en formato texto
Esto ocurre si hay símbolos de $ o comas.

**Solución:** Se corrige automáticamente al abrirlo en Excel. Si necesitas números puros en Python:
```python
df['Total'] = df['Total'].str.replace('[\$,]', '', regex=True).astype(float)
```

## 📋 Casos de uso

| Caso | Solución |
|------|----------|
| **Despacho contable** con 500 facturas/mes | Procesa todo en 2 minutos |
| **Auditoría** de RFC receptor | Valida automáticamente |
| **Reportes de proveedores** | Hoja RESUMEN lista para análisis |
| **Conciliación contable** | Extrae UUID para comparar con SAT |

## 🛠️ Desarrollo local

```bash
# Clonar
git clone https://github.com/tu-usuario/cfdi-extractor.git
cd cfdi-extractor

# Crear virtualenv
python -m venv venv
source venv/bin/activate  # En Windows: venv\Scripts\activate

# Instalar en modo desarrollo
pip install -e ".[dev]"

# Ejecutar tests
pytest tests/

# Linter
black cfdi_extractor.py
```

## 📝 Estructura del proyecto

```
cfdi-extractor/
├── cfdi_extractor.py          # Script principal
├── requirements.txt           # Dependencias
├── README.md                  # Este archivo
├── LICENSE                    # MIT License
├── examples/
│   ├── sample_invoice.xml     # Ejemplo CFDI
│   └── sample_config.yaml     # Configuración ejemplo
└── tests/
    └── test_extractor.py      # Tests unitarios
```

## 📄 License

MIT License - ver [LICENSE](LICENSE)

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Por favor:

1. Fork el proyecto
2. Crea una rama (`git checkout -b feature/mejora`)
3. Commit cambios (`git commit -m 'Agrega mejora'`)
4. Push (`git push origin feature/mejora`)
5. Abre un Pull Request

## ⚠️ Disclaimers

- Este proyecto **NO sustituye asesoría fiscal/contable**
- Valida **estructura** del XML pero NO validez fiscal ante SAT
- Uso bajo **tu responsabilidad**

## 📧 Contacto & Soporte

- **Issues:** Crea un issue en GitHub
- **Preguntas:** Discussions en GitHub

---

**Hecho con ❤️ por alguien que odia procesar XMLs manualmente**

### ¿Te salvó tiempo? Dale una ⭐ al repo
