<div align="center">

# 📊 Excel MCP Server

**Análisis rápido y eficiente de hojas de cálculo mediante operaciones atómicas, diseñado específicamente para agentes de IA**

[🇬🇧 English](../../README.md) • [🇷🇺 Русский](../ru/README.md) • [🇨🇳 中文](../zh/README.md) • 🇪🇸 Español • [🇯🇵 日本語](../ja/README.md) • [🇧🇷 Português](../pt/README.md)

Hecho con ❤️ por [@Jwadow](https://github.com/jwadow)

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![MCP](https://img.shields.io/badge/MCP-Compatible-green.svg)](https://modelcontextprotocol.io)
[![Sponsor](https://img.shields.io/badge/💖_Apoyar-Desarrollo-ff69b4)](#-apoya-el-proyecto)

**Analiza hojas de cálculo Excel con tu agente de IA mediante operaciones atómicas — sin volcar datos en el contexto**

*Funciona con OpenCode, Claude Code, Codex app, Cursor, Cline, Roo Code, Kilo Code y otros agentes de IA compatibles con MCP*

[Por qué existe esto](#-por-qué-existe-esto) • [Qué puede hacer tu agente](#-qué-puede-hacer-tu-agente) • [Instalación y configuración](#%EF%B8%8F-instalación-y-configuración) • [Herramientas disponibles](#%EF%B8%8F-herramientas-disponibles) • [💖 Donar](#-apoya-el-proyecto)

</div>

---

## 🤨 Por qué existe esto

**El problema:** La mayoría de herramientas Excel para IA vuelcan los datos crudos de la hoja de cálculo en el contexto del agente. Esto satura la ventana de contexto, ralentiza todo, y la IA aún puede calcular mal o confundirse en conjuntos de datos grandes.

**Este proyecto:** Piensa en SQL para Excel. Tu agente de IA compone operaciones atómicas (`filter_and_count`, `aggregate`, `group_by`) y obtiene resultados precisos — no miles de filas.

El agente analiza datos **sin verlos**. Los resultados llegan como números, fórmulas e insights.

> *"Esto es como trabajar con una base de datos mediante SQL, no arrastrando todo a la memoria."*
> — Agente de IA después de analizar una hoja de cálculo en producción

### 🔌 ¿Qué es MCP?

[Model Context Protocol](https://modelcontextprotocol.io) es un estándar abierto que permite a los agentes de IA usar herramientas externas.

Este proyecto es una de esas herramientas. Cuando conectas este servidor a tu agente de IA (OpenCode, Claude Code, Codex app, Cursor, Cline, Roo Code, Kilo Code, etc.), tu agente obtiene un montón de comandos nuevos para trabajar con archivos Excel — filtrado, conteo, agregación, análisis.

**La ventaja clave:** Tu IA no carga miles de filas de hojas de cálculo en su memoria. En su lugar, hace preguntas específicas y obtiene respuestas precisas. Más rápido, más preciso, sin desbordamiento de contexto.

---

## 💬 Lo que dicen los agentes de IA

Feedback real de agentes de IA que usaron este servidor MCP en producción:

> *"Analicé 34,211 filas sin cargar datos en el contexto. Cada operación devuelve solo el resultado — conteo, suma, promedio. El contexto se mantiene limpio. Las operaciones se ejecutan en 25-45ms independientemente del tamaño del archivo."*

> *"Esto es SQL para Excel. Consultas, filtros, agregación — sin volcar datos en el contexto. Herramienta sólida para tareas analíticas."*

> *"El sistema de filtros maneja bien la lógica compleja. Grupos AND/OR anidados, 12 operadores, condiciones ilimitadas. Construí una clasificación multicategoría sin escribir código."*

> *"Las operaciones por lotes son eficientes. Una llamada a `filter_and_count_batch` en lugar de múltiples solicitudes separadas. El archivo se carga una vez, se aplican todos los filtros, los resultados llegan juntos."*

*Sí, los agentes ahora escriben reseñas. Estas son reflexiones reales de agentes de IA analizando datos de hojas de cálculo del mundo real. Bienvenido a 2026.*

---

## 🚀 Qué puede hacer tu agente

Una vez conectado, tu agente de IA obtiene un montón de herramientas especializadas para analizar datos tabulares. El agente recibe solo consultas precisas y resultados confiables.

### 📊 Exploración de datos
- **Inspeccionar archivos** - estructura, hojas, columnas, tipos de datos (detecta automáticamente encabezados desordenados)
- **Perfilar columnas** - estadísticas, conteos de nulos, valores principales, calidad de datos en una llamada
- **Buscar datos** - buscar en múltiples hojas, localizar columnas en cualquier lugar

### 🔍 Filtrado y consultas
- **12 operadores de filtro** - `==`, `!=`, `>`, `<`, `>=`, `<=`, `in`, `not_in`, `contains`, `startswith`, `endswith`, `regex`
- **Lógica compleja** - grupos AND/OR anidados, operador NOT, condiciones ilimitadas
- **Operaciones por lotes** - clasificar datos en múltiples categorías en una solicitud (6x más rápido)
- **Análisis de superposición** - diagramas de Venn, conteos de intersección, operaciones de conjuntos

### 📈 Agregación y análisis
- **8 funciones de agregación** - sum, mean, median, min, max, std, var, count
- **Agrupar por** - tablas dinámicas con múltiples columnas de agrupación
- **Análisis estadístico** - correlaciones (Pearson/Spearman/Kendall), detección de valores atípicos (IQR/Z-score)
- **Series temporales** - crecimiento período a período, promedios móviles, totales acumulados

### 🏆 Operaciones avanzadas
- **Clasificación** - top-N, bottom-N, clasificación por percentiles (con soporte de agrupación)
- **Columnas calculadas** - expresiones aritméticas entre columnas
- **Validación de datos** - encontrar duplicados, valores nulos, verificaciones de calidad de datos
- **Comparación de hojas** - diferencias entre versiones, encontrar cambios

### ⚡ Características de rendimiento
- **Operaciones atómicas** - resultados en 20-50ms, sin importar el tamaño del archivo
- **Caché inteligente** - archivo cargado una vez, reutilizado para todas las operaciones
- **Filas de muestra** - vista previa de datos filtrados sin recuperación completa
- **Protección de contexto** - límites inteligentes previenen el desbordamiento del contexto de IA

### 📋 Integración con Excel
- **Generación de fórmulas** - cada resultado incluye fórmula Excel para actualizaciones dinámicas
- **Salida TSV** - copiar-pegar resultados directamente en Excel
- **Soporte legacy** - funciona con archivos .xls antiguos (Excel 97-2003)
- **Multi-hoja** - analizar múltiples hojas en un archivo

**Ejemplos de consultas que tu agente ahora puede manejar:**
- *"Muéstrame los 10 mejores clientes por ingresos"*
- *"Encuentra todos los pedidos del Q4 donde el monto > $1000"*
- *"Calcula el crecimiento mes a mes para cada categoría de producto"*
- *"¿Qué clientes son VIP y activos? (análisis de superposición)"*
- *"Encuentra duplicados en la columna Email"*

## ⚙️ Instalación y configuración

### Requisitos previos

**Python 3.10 o superior** — [Descargar aquí](https://www.python.org/downloads/)

### Paso 1: Clonar repositorio

```bash
git clone https://github.com/jwadow/mcp-excel.git
cd mcp-excel
```

*¿No tienes Git? Haz clic en "Code" → "Download ZIP" en la parte superior de esta página del repositorio, extrae y abre la terminal en esa carpeta.*

### Paso 2: Elegir método de instalación

<details>
<summary><b>🎯 Opción A: Poetry (Recomendado)</b></summary>

Poetry es un gestor de dependencias moderno de Python (reemplaza pip+venv+requirements.txt).
[Instálalo](https://python-poetry.org/docs/#installation): `pip install poetry` o `pipx install poetry`

**Instalar dependencias:**
```bash
poetry install
```

**Configurar tu agente de IA:**

Añade esto a tu configuración MCP (config JSON):
```json
{
  "mcpServers": {
    "excel": {
      "command": "poetry",
      "args": ["run", "python", "-m", "mcp_excel.main"],
      "cwd": "C:/path/to/mcp-excel"
    }
  }
}
```

**Importante:** Reemplaza `C:/path/to/mcp-excel` con la ruta real al repositorio clonado.

</details>

<details>
<summary><b>📦 Opción B: pip con entorno virtual</b></summary>

**Instalar dependencias:**
```bash
# Windows
python -m venv venv
venv\Scripts\activate
pip install -e .

# Linux/Mac
python -m venv venv
source venv/bin/activate
pip install -e .
```

**Encontrar ruta de Python en venv:**
```bash
# Windows
where python

# Linux/Mac
which python
```

**Configurar tu agente de IA:**

Añade esto a tu configuración MCP (config JSON):
```json
{
  "mcpServers": {
    "excel": {
      "command": "C:/path/to/mcp-excel/venv/Scripts/python.exe",
      "args": ["-m", "mcp_excel.main"],
      "cwd": "C:/path/to/mcp-excel"
    }
  }
}
```

**Importante:**
- Reemplaza `C:/path/to/mcp-excel/venv/Scripts/python.exe` con la ruta real del comando `where python`
- En Linux/Mac usa la ruta de `which python` (ej. `/path/to/mcp-excel/venv/bin/python`)

</details>

<details>
<summary><b>🐍 Opción C: Python del sistema (No recomendado)</b></summary>

**Instalar dependencias globalmente:**
```bash
pip install "mcp>=1.1.0" "pandas>=2.2.0" "pydantic>=2.10.0" "xlrd>=2.0.1" "openpyxl>=3.1.0" "psutil>=6.1.0" "python-dateutil>=2.9.0"
```

**Configurar tu agente de IA:**
```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["-m", "mcp_excel.main"],
      "cwd": "C:/path/to/mcp-excel"
    }
  }
}
```

⚠️ **Advertencia:** Esto contamina tu entorno Python global. Usa Poetry o venv en su lugar.

</details>

### Paso 3: Verificar instalación

Reinicia tu agente de IA y prueba:
```
"Analiza el archivo Excel en C:/Users/TuNombre/Documents/test.xlsx"
```

Si funciona - ¡listo! Si no, verifica:
- La ruta al repositorio es correcta en `cwd`
- La ruta de Python es correcta en `command` (para método pip)
- Todas las dependencias están instaladas

### Agentes de IA compatibles

Funciona con cualquier agente de IA compatible con MCP.

⚠️ **Importante:** Este es un servidor MCP. Se ejecuta automáticamente cuando tu agente de IA lo necesita. No lo ejecutes manualmente en la terminal.

## 💡 Uso

Después de la configuración, reinicia tu agente de IA y pídele que analice archivos Excel:

```
"Analiza el archivo Excel en C:/Users/TuNombre/Documents/sales.xls"
"Muéstrame los 10 mejores clientes por ingresos de sales.xlsx"
"Encuentra duplicados en la columna 'Email' en contacts.xlsx"
"Calcula el crecimiento mes a mes de revenue.xls"
```

## 🛠️ Herramientas disponibles

<details>
<summary><b>📋 Referencia completa de herramientas (25 herramientas) - Haz clic para expandir</b></summary>

### 📊 Inspección de archivos (5 herramientas)

#### `inspect_file`
Obtener vista general de la estructura del archivo - hojas, dimensiones, formato.
**Usar para:** Exploración inicial del archivo, descubrimiento de hojas, validación de formato
**Devuelve:** Lista de hojas, conteos de filas/columnas, metadatos del archivo

#### `get_sheet_info`
Análisis detallado de hoja con detección automática de encabezados.
**Usar para:** Entender estructura de datos, tipos de columnas, vista previa de muestras
**Devuelve:** Nombres/tipos de columnas, conteo de filas, datos de muestra (3 filas), info de detección de encabezados

#### `get_column_names`
Enumeración rápida de columnas sin cargar datos completos.
**Usar para:** Validación de esquema, construcción de filtros, verificación de disponibilidad de columnas
**Devuelve:** Lista de nombres de columnas, conteo de columnas

#### `get_data_profile`
Perfilado completo de columnas - tipos, estadísticas, nulos, valores principales.
**Usar para:** Exploración inicial de datos, evaluación de calidad, análisis de distribución
**Devuelve:** Por columna: tipo, % nulos, conteo único, estadísticas (numérico), top N valores
**Eficiencia:** Reemplaza 10+ llamadas separadas (get_column_stats + get_value_counts + find_nulls)

#### `find_column`
Localizar columna en múltiples hojas.
**Usar para:** Navegación multi-hoja, descubrimiento de datos, análisis entre hojas
**Devuelve:** Lista de hojas con ubicaciones de columnas, índices, conteos de filas (sin distinción de mayúsculas)

---

### 📥 Recuperación de datos (3 herramientas)

#### `get_unique_values`
Extraer valores únicos de una columna.
**Usar para:** Exploración de datos, construcción de filtros, descubrimiento de valores distintos, verificaciones de calidad de datos
**Devuelve:** Lista de valores únicos, conteo, bandera de truncamiento (si se excede el límite)
**Límite predeterminado:** 100 valores

#### `get_value_counts`
Análisis de frecuencia - top N valores más comunes.
**Usar para:** Análisis de distribución, identificar categorías dominantes, detección de desequilibrio de datos
**Devuelve:** Diccionario valor → conteo, conteo total, salida TSV
**Predeterminado:** Top 10 valores

#### `filter_and_get_rows`
Recuperar filas filtradas con paginación.
**Usar para:** Extracción de datos, inspección de muestras, análisis detallado, exportación
**Devuelve:** Filas filtradas (lista de diccionarios), conteo total, salida TSV
**Paginación:** Soporte de limit/offset

---

### 🔍 Filtrado y conteo (3 herramientas)

#### `filter_and_count`
Contar filas que coinciden con condiciones con 14 operadores.
**Operadores:** `==`, `!=`, `>`, `<`, `>=`, `<=`, `in`, `not_in`, `contains`, `startswith`, `endswith`, `regex`, `is_null`, `is_not_null`
**Lógica:** Grupos AND/OR anidados, operador NOT, condiciones ilimitadas
**Usar para:** Clasificación, segmentación, validación de datos, conteo de categorías
**Devuelve:** Conteo + fórmula Excel (COUNTIFS), filas de muestra opcionales

#### `filter_and_count_batch`
Clasificar datos en múltiples categorías en una llamada (6x más rápido).
**Usar para:** Clasificación multicategoría, segmentación de mercado, control de calidad
**Devuelve:** Conteo + fórmula por categoría, tabla TSV para Excel
**Eficiencia:** Carga archivo una vez, aplica todos los filtros, devuelve todos los resultados

#### `analyze_overlap`
Análisis de diagrama de Venn - intersecciones, uniones, zonas exclusivas.
**Usar para:** Análisis de superposición, oportunidades de venta cruzada, verificaciones de consistencia de datos
**Devuelve:** Conteos de conjuntos, intersecciones por pares (A ∩ B), unión, datos de Venn (2-3 conjuntos)
**Ejemplos:** Clientes VIP Y activos, superposiciones de categorías de productos, pedidos completados SIN fecha de finalización

---

### 📈 Agregación y análisis (2 herramientas)

#### `aggregate`
Realizar agregación con filtros opcionales (8 operaciones).
**Operaciones:** `sum`, `mean`, `median`, `min`, `max`, `std`, `var`, `count`
**Usar para:** Totales, promedios, valores mín/máx, resúmenes estadísticos, agregaciones condicionales, cálculos de KPI
**Devuelve:** Valor agregado + fórmula Excel (SUMIF, AVERAGEIF, etc.)
**Especial:** Autoconversión de números almacenados como texto a numérico

#### `group_by`
Tabla dinámica con agrupación de múltiples columnas.
**Usar para:** Análisis de categorías, agrupación jerárquica, ventas por región/producto
**Devuelve:** Datos agrupados con valores agregados, salida TSV
**Soporta:** Múltiples columnas de agrupación, todas las 8 operaciones de agregación

---

### 📊 Estadísticas (3 herramientas)

#### `get_column_stats`
Resumen estadístico - conteo, media, mediana, desviación estándar, cuartiles.
**Usar para:** Análisis de distribución, perfilado de datos, preparación para detección de valores atípicos
**Devuelve:** Estadísticas completas (min, max, mean, median, std, Q1, Q3), conteo de nulos, salida TSV

#### `correlate`
Matriz de correlación entre 2+ columnas.
**Métodos:** Pearson (lineal), Spearman (basado en rangos), Kendall (basado en rangos)
**Usar para:** Análisis de relaciones, dependencia de variables, selección de características
**Devuelve:** Matriz de correlación (-1 a 1), salida TSV

#### `detect_outliers`
Detección de anomalías usando método IQR o Z-score.
**Métodos:** IQR (robusto), Z-score (asume distribución normal)
**Usar para:** Detección de fraude, errores de sensores, calidad de datos, identificación de valores inusuales
**Devuelve:** Filas atípicas con índices, conteo, método/umbral usado

---

### ✅ Validación de datos (2 herramientas)

#### `find_duplicates`
Detectar filas duplicadas por columnas especificadas.
**Usar para:** Calidad de datos, planificación de deduplicación, verificaciones de integridad
**Devuelve:** Todas las filas duplicadas (incluida la primera aparición), conteo, índices
**Nota:** Usa `duplicated(keep=False)` para marcar todos los duplicados

#### `find_nulls`
Encontrar valores nulos/vacíos con estadísticas detalladas.
**Usar para:** Verificaciones de completitud, análisis de valores faltantes, limpieza de datos
**Devuelve:** Por columna: conteo de nulos, porcentaje, índices (primeros 100)
**Nota:** Los marcadores de posición (".", "-") NO son nulos - usa operadores `==` o `in`

---

### 🔄 Operaciones multi-hoja (2 herramientas)

#### `search_across_sheets`
Buscar valor en todas las hojas.
**Usar para:** Búsqueda entre hojas, seguimiento de valores, ubicación de datos
**Devuelve:** Lista de hojas con conteos de coincidencias, coincidencias totales
**Soporta:** Valores numéricos y de cadena

#### `compare_sheets`
Diferencia entre dos hojas usando columna clave.
**Usar para:** Comparación de versiones, detección de cambios, conciliación, pistas de auditoría
**Devuelve:** Filas con diferencias, estado (only_in_sheet1/sheet2/different_values), comparación lado a lado

---

### 📅 Series temporales (3 herramientas)

#### `calculate_period_change`
Análisis de crecimiento período a período.
**Períodos:** month, quarter, year
**Usar para:** Análisis de tendencias, seguimiento de crecimiento, comparación estacional, análisis interanual
**Devuelve:** Períodos con valores, cambios absolutos/porcentuales, fórmula Excel

#### `calculate_running_total`
Suma acumulativa con agrupación opcional.
**Usar para:** Análisis acumulativo, seguimiento de progreso, cálculos de saldo, flujo de caja
**Devuelve:** Filas con totales acumulados, fórmula Excel (SUM($B$2:B2))
**Soporta:** Agrupación (el total acumulado se reinicia por grupo)

#### `calculate_moving_average`
Suavizado con tamaño de ventana especificado.
**Usar para:** Detección de tendencias, reducción de ruido, identificación de patrones
**Devuelve:** Filas con promedios móviles, fórmula Excel (AVERAGE(B1:B7))
**Ejemplos:** Promedio móvil de 7 días, suavizado de precio de acciones de 30 días

---

### 🏆 Operaciones avanzadas (2 herramientas)

#### `rank_rows`
Clasificar por valor de columna con filtrado top-N.
**Direcciones:** desc (más alto primero), asc (más bajo primero)
**Usar para:** Tablas de clasificación, análisis top/bottom, clasificación por percentiles
**Devuelve:** Filas clasificadas con números de rango, fórmula Excel (RANK)
**Soporta:** Filtrado top-N, clasificación dentro de grupos

#### `calculate_expression`
Expresiones aritméticas entre columnas.
**Operaciones:** `+`, `-`, `*`, `/`, paréntesis
**Usar para:** Métricas derivadas, cálculos financieros, análisis de ratios, cálculos de KPI
**Devuelve:** Valores calculados, fórmula Excel (ej. =A2*B2)
**Ejemplos:** Ingresos = Precio * Cantidad, Margen = (Ingresos - Costo) / Ingresos

</details>

## 🗺️ Hoja de ruta

### 📁 Soporte de formatos de archivo

**Actualmente soportado:**
- ✅ **XLS** - Excel 97-2003 (solo lectura)
- ✅ **XLSX** - Excel 2007+ (solo lectura)

**Planificado:**
- 🔜 **XLSM** - Excel con soporte de macros
- 🔜 **CSV** - Valores separados por comas
- 🔜 **TSV** - Valores separados por tabulaciones
- 🔜 **ODS** - Hoja de cálculo OpenDocument
- 🔜 **Parquet** - Formato de almacenamiento columnar

### 🚀 Características

- **Operaciones de escritura** - Modificar archivos de hojas de cálculo (crear columnas calculadas, actualizar valores)
- **Modo de transporte SSE** - Eventos enviados por servidor para acceso remoto
- **Generación avanzada de fórmulas** - Fórmulas Excel más complejas con funciones anidadas
- **Exportación de datos** - Exportar resultados filtrados/agregados a nuevos archivos

---

## 📜 Licencia

Este proyecto está licenciado bajo la **GNU Affero General Public License v3.0 (AGPL-3.0)**.

Esto significa:
- ✅ Puedes usar, modificar y distribuir este software
- ✅ Puedes usarlo con fines comerciales
- ⚠️ **Debes divulgar el código fuente** cuando distribuyas el software
- ⚠️ **El uso en red es distribución** — si ejecutas una versión modificada en un servidor y permites que otros interactúen con ella, debes hacer disponible el código fuente
- ⚠️ Las modificaciones deben publicarse bajo la misma licencia

Consulta el archivo [LICENSE](../../LICENSE) para el texto completo de la licencia.

### ¿Por qué AGPL-3.0?

AGPL-3.0 asegura que las mejoras a este software beneficien a toda la comunidad. Si modificas este servidor y lo despliegas como servicio, debes compartir tus mejoras con tus usuarios.

---

## 💖 Apoya el proyecto

<div align="center">

<img src="https://raw.githubusercontent.com/Tarikul-Islam-Anik/Animated-Fluent-Emojis/master/Emojis/Smilies/Smiling%20Face%20with%20Hearts.png" alt="Love" width="80" />

**¡Si este proyecto te ahorró tiempo o dinero, considera apoyarlo!**

Cada contribución ayuda a mantener el proyecto vivo y en crecimiento

<br>

### 🤑 Donar

[**☕ Donación única**](https://app.lava.top/jwadow?tabId=donate) • [**💎 Apoyo mensual**](https://app.lava.top/jwadow?tabId=subscriptions)

<br>

### 🪙 O envía cripto

| Moneda | Red | Dirección |
|:--------:|:-------:|:--------|
| **USDT** | TRC20 | `TSVtgRc9pkC1UgcbVeijBHjFmpkYHDRu26` |
| **BTC** | Bitcoin | `12GZqxqpcBsqJ4Vf1YreLqwoMGvzBPgJq6` |
| **ETH** | Ethereum | `0xc86eab3bba3bbaf4eb5b5fff8586f1460f1fd395` |
| **SOL** | Solana | `9amykF7KibZmdaw66a1oqYJyi75fRqgdsqnG66AK3jvh` |
| **TON** | TON | `UQBVh8T1H3GI7gd7b-_PPNnxHYYxptrcCVf3qQk5v41h3QTM` |

</div>

---

## 🤝 Contribuir

¡Las contribuciones son bienvenidas! Por favor asegúrate de que:

1. Todas las dependencias sean compatibles con AGPL
2. El código siga el estilo existente
3. Se incluyan pruebas para nuevas características
4. La documentación esté actualizada

Para problemas, errores o contribuciones, por favor abre un issue en GitHub.

---

## 💬 ¿Necesitas ayuda?

¿Tienes preguntas? ¿Encontraste un error? ¿Tienes una idea para una característica? ¡Estamos aquí para ayudar!

**👉 [Abrir un Issue en GitHub](https://github.com/jwadow/mcp-excel/issues/new)**

Ya sea que estés atascado con la instalación, encontraste algo roto o simplemente quieres sugerir una mejora — GitHub Issues es el lugar. No te preocupes si eres nuevo en GitHub, solo haz clic en el enlace de arriba y describe tu situación. Lo resolveremos juntos.

---

<div align="center">

**[⬆ Volver arriba](#-excel-mcp-server)**

</div>
