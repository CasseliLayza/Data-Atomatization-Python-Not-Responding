![logo](asset/img/notResponding.png)

# 📊 Automatización de Reporte Not Responding - Genesys Cloud

Proyecto de automatización en **Python** para consultar eventos **Not Responding** desde **Genesys Cloud Analytics API**, generar archivos **CSV** y **gráficas por hora del día**, y poblar automáticamente una plantilla de **Excel** preparada para reportería operativa.

---

## 🎯 Objetivo

Automatizar el flujo completo de reportería de **Not Responding**:

1. Consultar por API las conversaciones de Genesys Cloud filtradas por `tNotResponding`.
2. Generar archivos base en CSV:
   - por fecha y hora
   - por hora del día
   - inconsistencias

3. Generar gráficas PNG.
4. Poblar automáticamente una plantilla Excel en la hoja:
   - `RESUMEN POR RANGO DE HORA`

5. Refrescar el Excel para actualizar tablas, conexiones y visualizaciones.
6. Ejecutar el proceso para un día específico o para un rango completo de fechas.

---

## 🏗️ Arquitectura del flujo

### 1️⃣ Generación de base diaria

Script principal que consulta Genesys Cloud y genera:

- CSV por `conversationStart`
- CSV por `tNotResponding emitDate`
- CSV de inconsistencias
- gráficas PNG

### 2️⃣ Populate de Excel

Script que toma como fuente el CSV **por hora del día** y llena la hoja:

- `RESUMEN POR RANGO DE HORA`

Este proceso se realiza usando **Excel COM** (`pywin32`) para no romper:

- slicers
- segmentadores
- tablas dinámicas
- validaciones
- conexiones del workbook

### 3️⃣ Runner por rango

Script orquestador que:

- procesa fechas en rango
- genera la base por cada día
- ejecuta el populate sin refresh por cada fecha
- opcionalmente hace un único refresh final

---

## 📁 Estructura del proyecto

```text
.
├── notRespondingV5.py
├── populate_nr_excel_v2.py
├── run_nr_range.py
├── requirements.txt
├── README.md
├── LICENSE
├── Not-Responding-Report.xlsx
│
├── outputConversationStart/
├── outputEvent/
├── outputDebug/
├── outputCharts/
│
└── Reporte_Final.xlsx
```

---

## ⚙️ Scripts principales

## 1) 🧠 `notRespondingV5.py`

Consulta Genesys Cloud para una fecha dada y genera los archivos base.

### ✅ Funcionalidades

- autenticación OAuth Client Credentials
- consulta paginada a:
  - `/api/v2/analytics/conversations/details/query`

- filtro por:
  - `tNotResponding exists`
  - `direction inbound/outbound`

- generación de:
  - CSV por fecha y hora
  - CSV por hora del día
  - CSV de inconsistencias
  - gráficas PNG

### 📤 Salidas

#### `outputConversationStart`

- `conversation_start_por_fecha_hora_<fecha>_<runid>.csv`
- `conversation_start_por_hora_del_dia_<fecha>_<runid>.csv`

#### `outputEvent`

- `evento_tnotresponding_por_fecha_hora_<fecha>_<runid>.csv`
- `evento_tnotresponding_por_hora_del_dia_<fecha>_<runid>.csv`

#### `outputDebug`

- `inconsistencias_tnotresponding_<fecha>_<runid>.csv`

#### `outputCharts`

- `grafica_conversation_start_<fecha>_<runid>.png`
- `grafica_evento_tnotresponding_<fecha>_<runid>.png`

### ▶️ Ejecución

#### Fecha específica

```bash
python notRespondingV5.py 2026-04-21
```

#### Sin parámetro

Usa la fecha actual en hora Perú:

```bash
python notRespondingV5.py
```

---

## 2) 📘 `populate_nr_excel_v2.py`

Toma el CSV **por hora del día** y pobla la hoja del Excel:

- `RESUMEN POR RANGO DE HORA`

### ✅ Características

- usa `pywin32` y automatización COM
- no modifica la lógica del dashboard manualmente
- agrega días nuevos al archivo acumulado
- evita duplicar fechas
- opcionalmente reemplaza una fecha ya existente
- opcionalmente ejecuta `RefreshAll`

### 📌 Fuente recomendada

Por defecto se recomienda usar:

- `conversation_start_por_hora_del_dia`

porque es el insumo más estable para la tabla horaria del Excel.

### ▶️ Ejecución

#### Poblar una fecha

```bash
python populate_nr_excel_v2.py --workbook "Not-Responding-Report.xlsx" --date 2026-04-21 --mode conversation --output "Reporte_Final.xlsx"
```

#### Reemplazar una fecha ya existente

```bash
python populate_nr_excel_v2.py --workbook "Not-Responding-Report.xlsx" --date 2026-04-21 --mode conversation --output "Reporte_Final.xlsx" --replace-date
```

#### Ejecutar refresh final

```bash
python populate_nr_excel_v2.py --workbook "Not-Responding-Report.xlsx" --date 2026-04-21 --mode conversation --output "Reporte_Final.xlsx" --refresh
```

#### Refresh con mayor espera

```bash
python populate_nr_excel_v2.py --workbook "Not-Responding-Report.xlsx" --date 2026-04-21 --mode conversation --output "Reporte_Final.xlsx" --refresh --refresh-wait 10
```

### 🧩 Parámetros principales

- `--workbook`: plantilla Excel base
- `--date`: fecha objetivo `YYYY-MM-DD`
- `--mode`: `conversation` o `event`
- `--csv`: ruta explícita a un CSV
- `--output`: archivo Excel de salida/acumulado
- `--replace-date`: reemplaza la fecha si ya existe
- `--refresh`: ejecuta `Actualizar todo`
- `--refresh-wait`: segundos de espera luego de `RefreshAll`

---

## 3) 🚀 `run_nr_range.py`

Runner para procesar un rango completo de fechas.

### ✅ Qué hace

Por cada día del rango:

1. verifica si ya existe el CSV diario
2. si no existe, ejecuta `notRespondingV5.py`
3. ejecuta `populate_nr_excel_v2.py` sin refresh

Al final:

4. opcionalmente ejecuta un único refresh final

### ▶️ Ejecución

#### Rango simple

```bash
python run_nr_range.py --start-date 2026-01-01 --end-date 2026-04-21
```

#### Con refresh final único

```bash
python run_nr_range.py --start-date 2026-01-01 --end-date 2026-04-21 --final-refresh
```

#### Usando modo event

```bash
python run_nr_range.py --start-date 2026-01-01 --end-date 2026-04-21 --mode event
```

### 🧩 Parámetros principales

- `--start-date`
- `--end-date`
- `--mode`
- `--generator-script`
- `--populate-script`
- `--workbook-template`
- `--output-workbook`
- `--log-file`
- `--refresh-wait`
- `--final-refresh`

---

## 📦 Requisitos

### 🐍 Python

Recomendado:

- Python 3.10 o superior

### 📚 Dependencias

Instalar desde `requirements.txt`:

```bash
pip install -r requirements.txt
```

### `requirements.txt`

```txt
requests>=2.31.0
matplotlib>=3.8.0
pywin32>=306
```

---

## 🔐 Variables de entorno

El script generador usa estas variables:

```bash
export GENESYS_CLIENT_ID="tu_client_id"
export GENESYS_CLIENT_SECRET="tu_client_secret"
export GENESYS_REGION="mypurecloud.com"
```

En Windows / Git Bash:

```bash
export GENESYS_CLIENT_ID="tu_client_id"
export GENESYS_CLIENT_SECRET="tu_client_secret"
export GENESYS_REGION="mypurecloud.com"
```

---

## 🧮 Lógica de datos

### 📍 Fuente de consulta

Se usa el endpoint:

- `POST /api/v2/analytics/conversations/details/query`

con estos criterios principales:

- `metric = tNotResponding`
- `operator = exists`
- `direction = inbound/outbound`

### 1️⃣ `conversationStart`

Cuenta la conversación según la hora en que inició.

### 2️⃣ `emitDate` de `tNotResponding`

Cuenta el evento real según la hora exacta en que ocurrió el `Not Responding`.

### ✅ Recomendación operativa

Para poblar el Excel horario:

- usar `conversation_start_por_hora_del_dia`

Para análisis más fino:

- revisar también el CSV de `event`

---

## ⚠️ Consideraciones importantes

### 1. 📘 Excel COM

Se usa `pywin32` para manipular el Excel real porque `openpyxl` puede alterar o romper:

- slicers
- tablas dinámicas
- referencias internas
- conexiones
- comportamiento del dashboard

### 2. 🔄 Refresh del Excel

El refresh puede ser costoso si se ejecuta por cada día.
Por eso se recomienda:

- carga histórica: **sin refresh diario**
- refresh: **solo una vez al final**

### 3. 📅 Duplicados por fecha

El populate evita insertar una fecha dos veces.
Si necesitas recalcular una fecha ya cargada, usa:

```bash
--replace-date
```

### 4. 📂 Archivo acumulado

El Excel de salida se va reutilizando entre ejecuciones:

- si no existe, se crea desde la plantilla
- si ya existe, se sigue poblando

---

## 🛠️ Flujo recomendado de uso

### 📚 Carga histórica

#### Paso 1

Procesar rango:

```bash
python run_nr_range.py --start-date 2026-01-01 --end-date 2026-04-21 --final-refresh
```

#### Resultado

- genera todos los CSV necesarios
- llena el Excel acumulado
- hace un refresh final único

### 📆 Operación diaria

#### Paso 1

Generar base del día:

```bash
python notRespondingV5.py
```

#### Paso 2

Poblar Excel:

```bash
python populate_nr_excel_v2.py --workbook "Not-Responding-Report.xlsx" --date 2026-04-21 --mode conversation --output "Reporte_Final.xlsx" --refresh
```

---

## 🧯 Troubleshooting

### ❌ Error: `ModuleNotFoundError`

Instala dependencias:

```bash
pip install -r requirements.txt
```

### ❌ Error: Excel no puede abrir el archivo

Asegúrate de que:

- la extensión de salida coincida con la plantilla
- si la plantilla es `.xlsx`, la salida también sea `.xlsx`

### ❌ Error: `PermissionError`

Probablemente el archivo Excel o CSV está abierto.
Cierra Excel antes de ejecutar.

### ❌ Error: no encuentra CSV para la fecha

Asegúrate de haber generado antes la base para esa fecha:

```bash
python notRespondingV5.py 2026-04-21
```

### ❌ El dashboard no refleja cambios

Usa el populate con:

```bash
--refresh
```

o el runner con:

```bash
--final-refresh
```

---

## 🌱 Mejoras futuras sugeridas

- soporte para rango de fechas directamente en el generador
- exportación consolidada mensual
- logging más detallado por archivo
- notificación de días fallidos
- validación adicional del contenido del CSV
- selección de tabla Excel por nombre en vez de solo por hoja

---

### 📄 Licencia

Este proyecto está bajo la licencia **MIT**. Puedes consultar más detalles en el archivo `LICENSE`.

---

### 📬 Contacto

Para dudas, sugerencias o contribuciones, puedes escribir a:

📧 **[casseli.layza@gmail.com](mailto:casseli.layza@gmail.com)**

🔗 [LinkedIn](https://www.linkedin.com/in/casseli-layza/)
🔗 [GitHub](https://github.com/CasseliLayza)

💡 **Desarrollado por Casseli Layza como parte de una automatización de reportería con Python, Genesys Cloud y Excel.**

**_💚 ¡Gracias por revisar este proyecto!... Powered by Casse 🌟📚🚀...!!_**

### Derechos Reservados

```markdown
© 2026 Casse. Todos los derechos reservados.
```
