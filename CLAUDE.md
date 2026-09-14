# CLAUDE.md — Modelo de Visualizaciones

> Este archivo es para **vos, Claude**. Léelo antes de tocar cualquier cosa de este proyecto.
> El usuario te lo pasa para que entiendas el contexto y el flujo sin tener que adivinar.

## Qué es este proyecto

Un **template/scaffold** para generar reportes HTML interactivos a partir de un Excel. No es un proyecto productivo: es un esqueleto que el usuario adapta a su caso (sus columnas, sus métricas, sus gráficos). El objetivo es que un usuario pueda dejar caer un Excel en `/Input`, correr un comando, y obtener un HTML listo para compartir.

## Filosofía del flujo

Inspirado en el patrón del proyecto hermano `Facturacion Macro`:

```
Input/*.xlsx  →  Logica/main.py  →  Template/*.html (Jinja2)  →  Output/*.html
```

Reglas:

- **Una sola fuente de datos**: un Excel en `/Input`. Nada de scraping, APIs, ni bases de datos.
- **Output autocontenido**: un solo archivo HTML, abrible con doble click, sin servidor.
- **Plotly.js vía CDN** para los gráficos (no se empaqueta JS local).
- **Jinja2** para el render: el HTML es un template con `{{ variables }}` y `{% for %}`.
- **pandas** para todo lo de datos (carga, limpieza, agregaciones).

## Estructura

```
modelo_visualizaciones/
├── Input/             # ÚNICA entrada — Excel del usuario
├── Logica/main.py     # ÚNICO script — pipeline completo
├── Template/*.html    # templates Jinja2 (layout + Plotly)
├── Output/            # HTML generados (no editar a mano)
├── requirements.txt
├── README.md          # para humanos
└── CLAUDE.md          # este archivo
```

## Stack

- Python 3.10+
- `pandas` — manejo de datos
- `openpyxl` — lectura del Excel
- `Jinja2` — render de templates
- `Plotly.js` (vía CDN, cargado en el navegador) — gráficos

No hay backend. No hay base de datos. No hay React. Resistí la tentación de meter cosas; el valor de este modelo es que es chico y entendible de un vistazo.

## Cómo está organizado `main.py`

Cuatro pasos lineales, en este orden:

1. **`load_excel(path)`** — lee el Excel, normaliza nombres de columnas según `COLUMN_MAP`, tipa fechas/números, descarta NaN críticos.
2. **`build_context(df)`** — toma el DataFrame y devuelve un `dict` con todo lo que el template necesita: KPIs, JSON de series, tabla. **Toda la lógica de negocio vive acá.**
3. **`render(ctx, template_name)`** — carga el template Jinja y le pasa el contexto.
4. **`save_output(html)`** — guarda en `/Output`, versionando por fecha (`_v2`, `_v3`...).

## Convenciones críticas (NO violar)

- **Nombres de carpetas en español con mayúscula inicial**: `Input/`, `Logica/`, `Template/`, `Output/`. Es deliberado, no las renombres a `input/` o `templates/`.
- **El usuario sólo toca `/Input`**. Todo lo demás es del proyecto. Si necesitás que ponga algo en otro lado, pensá si en vez de eso podés leerlo de `/Input`.
- **Datos para Plotly se pasan como JSON-string** vía Jinja con `| safe`. Ver `serie_temporal_json` y `por_categoria_json` en `main.py` y `reporte.html`. Si agregás un gráfico nuevo, seguí el mismo patrón.
- **No hardcodees rutas absolutas**. Usá `BASE_DIR = Path(__file__).resolve().parent.parent` y derivá todo de ahí.
- **El HTML de salida es autocontenido**. Nada de archivos `.css` o `.js` separados. Todo inline o vía CDN.
- **Versionado del output por fecha**: si ya existe `reporte_AAAA-MM-DD.html`, generar `_v2`, `_v3`, etc. (lógica ya implementada en `save_output`).

## Cuando el usuario te pida adaptar el modelo a sus datos

Hacé este checklist mental:

1. **Pediles el Excel** (o los nombres reales de las columnas). No adivines.
2. **Ajustá `COLUMN_MAP`** en `main.py` para mapear sus nombres a los nombres internos (`fecha`, `categoria`, `subcategoria`, `valor`).
3. **Ajustá `build_context()`** si las métricas que quieren son distintas (sumas, promedios, ratios, top-N, etc.). Esta función es donde vive la lógica específica del caso.
4. **Ajustá `reporte.html`** si el layout o los gráficos son distintos. Para un gráfico nuevo: agregá un `<div id="chart-X">`, exportá los datos como JSON desde Python, e instanciá `Plotly.newPlot(...)` al final.
5. **Probá end-to-end**: `python Logica/main.py` y abrí el HTML resultante.

## Cuando el usuario quiera múltiples reportes (multi-output)

Patrón sugerido (el proyecto hermano `Facturacion Macro` lo hace así):

- Un `REPORTS_REGISTRY` dict en `main.py` con la config de cada reporte (nombre, template, filtro de datos, carpeta de salida).
- Un loop que para cada entrada del registry: filtra el df, construye contexto, renderiza, guarda.
- Templates separados en `/Template` (uno por tipo de reporte) o un template parametrizado.

No metas esto si no lo piden — es overkill para el caso simple.

## Errores típicos a evitar

- Querer "modernizar" usando React/Vite/Next.js. **No.** El valor es que es Python + Jinja + Plotly, sin build step.
- Pasar a Plotly un dict de Python directamente sin `json.dumps`. Rompe en serialización de fechas/numpy.
- Olvidarse del `| safe` en Jinja al inyectar JSON (queda HTML-escapeado y rompe el JS).
- Hardcodear paths con `\\` o `C:\...`. Usá `pathlib.Path` siempre.
- Cargar todo el Excel en memoria si es enorme. Si el usuario tiene >1M filas, sugerí pre-agregar.

## Si te piden algo fuera de scope

- "Quiero que el reporte se actualice solo" → no es un servidor; sugiere correr `main.py` desde una tarea programada (Windows Task Scheduler / cron).
- "Quiero login" → no es el modelo correcto. Sugerí Streamlit o una app web propiamente dicha.
- "Quiero datos en vivo de una API" → agregá un módulo `Logica/fetch.py` que se ejecute antes y deje un Excel en `/Input`. Mantenés el contrato del flujo.

## Comando de verificación rápida

```bash
cd modelo_visualizaciones
pip install -r requirements.txt
python Logica/main.py
# debe imprimir "✅ Reporte generado: Output/reporte_<fecha>.html"
```

Si esto no funciona, algo está roto. Diagnosticá antes de seguir.
