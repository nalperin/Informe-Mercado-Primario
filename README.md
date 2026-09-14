# Modelo de Visualizaciones

Plantilla mínima y reutilizable para generar **reportes HTML interactivos** a partir de un Excel.
La idea: vos pones tus datos en `/Input`, corres un script, y obtenes un HTML autocontenido en `/Output` listo para compartir.

## Flujo

```
Input (Excel)  →  Logica/main.py (pandas)  →  Template (Jinja2 + Plotly.js)  →  Output (HTML)
```

```
modelo_visualizaciones/
├── Input/             # Pone aca tu Excel de entrada
│   └── datos.xlsx     # Archivo de ejemplo (datos sinteticos)
├── Logica/
│   └── main.py        # Carga, limpia, agrega y renderiza
├── Template/
│   └── reporte.html   # Layout + gráficos (Jinja2 + Plotly.js)
├── Output/            # Reportes generados (HTML versionados por fecha)
├── requirements.txt
├── README.md          # Este archivo (para humanos)
└── CLAUDE.md          # Contexto e instrucciones para Claude
```

## Setup

```bash
# 1. (opcional) crear venv
python -m venv venv
venv\Scripts\activate          # Windows
# source venv/bin/activate     # macOS / Linux

# 2. instalar dependencias
pip install -r requirements.txt
```

## Uso

```bash
# Con el Excel de ejemplo:
python Logica/main.py

# Con tu propio archivo:
python Logica/main.py --input "Input/mi_archivo.xlsx" --sheet "Hoja1"
```

El reporte se genera en `Output/reporte_AAAA-MM-DD.html`.
Si ya existe uno del mismo día, agrega `_v2`, `_v3`, etc.

## Cómo adaptarlo a tus datos

El modelo asume un Excel con estas columnas (renombrables):

| Columna esperada | Tipo     | Descripción                                |
|------------------|----------|---------------------------------------------|
| `Fecha`          | fecha    | Fecha de la observación                     |
| `Categoria`      | texto    | Agrupador principal                         |
| `Subcategoria`   | texto    | Agrupador secundario                        |
| `Valor`          | numérico | Métrica a sumar/promediar                   |

Si tu Excel usa otros nombres, ajusta `COLUMN_MAP` en `Logica/main.py`:

```python
COLUMN_MAP = {
    "fecha":        "Fecha de operación",   # nombre real en tu Excel
    "categoria":    "Area",
    "subcategoria": "Tipo",
    "valor":        "Importe USD",
}
```

## Qué genera el reporte

- Tarjetas de **KPIs** (total, promedio, # registros, # categorías).
- Gráfico de **serie temporal mensual** (Plotly line chart).
- Gráfico de **total por categoría** (Plotly bar chart).
- **Tabla** con los últimos 20 registros.

Todo en un solo `.html` autocontenido. Plotly se carga vía CDN (la primera vez requiere internet; luego el navegador lo cachea).

## Cómo extenderlo

1. **Agregar un gráfico nuevo** → en `build_context()` (en `main.py`) calculá la serie y exportá un nuevo JSON; en `reporte.html` agregá un `<div>` y un `Plotly.newPlot(...)` consumiéndolo.
2. **Cambiar el layout/colores** → editá `reporte.html` (todo el CSS está inline al inicio).
3. **Múltiples reportes** → duplicá `reporte.html` (ej: `reporte_ventas.html`) y corré `python Logica/main.py --template reporte_ventas.html --prefix ventas`.
4. **Filtros interactivos** → Plotly soporta `updatemenus` y `sliders`; alternativa: vanilla JS leyendo los datasets ya inyectados.

## Tips

- Si tu Excel tiene varias hojas, usá `--sheet "Nombre"` o `--sheet 1`.
- El HTML pesa típicamente 50–500 KB. Si crece mucho, considerá samplear o agregar antes de inyectar.
- Para enviar el reporte por mail, adjuntá el `.html` directamente: se abre con doble click.
