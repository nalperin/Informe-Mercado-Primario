"""
MODELO DE VISUALIZACIONES — main.py
====================================
Flujo genérico para generar reportes HTML interactivos a partir de un Excel.

Pipeline:
    /Input/*.xlsx   →   pandas (limpieza + agregaciones)   →
    /Template/*.html (Jinja2)   →   /Output/reporte_<fecha>.html

Uso:
    python Logica/main.py
    python Logica/main.py --input "Input/mis_datos.xlsx"
    python Logica/main.py --sheet "Hoja1"

El reporte de salida es un HTML autocontenido (Plotly.js vía CDN, no requiere
servidor) que el usuario puede abrir en cualquier navegador.

----
COMO ADAPTARLO A TUS DATOS
----
1. Coloca tu Excel en /Input/.
2. Ajusta `COLUMN_MAP` para que apunte a las columnas reales de tu Excel.
3. Ajusta `build_context()` para construir las series/tablas/KPIs que quieras.
4. Ajusta `Template/reporte.html` para el layout visual que quieras.
"""

from __future__ import annotations

import argparse
import base64
import json
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape

# ============================================================
# HELPER: fecha en español
# ============================================================
MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

def fecha_es():
    now = datetime.now()
    return f"{MESES_ES[now.month]} {now.year}"


# ============================================================
# PATHS — relativos a la raíz del proyecto (/modelo_visualizaciones)
# ============================================================
BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "Input"
TEMPLATE_DIR = BASE_DIR / "Template"
OUTPUT_DIR = BASE_DIR / "Output"

DEFAULT_TEMPLATE = "reporte.html"


def find_latest_excel(directory: Path) -> Path:
    """Devuelve el .xlsx más reciente en la carpeta. Si no hay ninguno, lanza error."""
    excels = sorted(directory.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
    if not excels:
        raise FileNotFoundError(f"No hay archivos .xlsx en {directory}")
    return excels[0]


# ============================================================
# CONFIG DE DATOS — adaptar a tu Excel
# ============================================================
# Mapeo: <nombre amigable> -> <nombre real en el Excel>.
# Si tu Excel ya tiene los nombres "limpios", podés dejarlos iguales.
COLUMN_MAP = {
    "fecha": "Fecha Colocación",
    "categoria": "Régimen de emisión",
    "subcategoria": "Tipo de Tasa",
    "valor": "Monto en $ equivalentes",
}


# ============================================================
# CARGA Y LIMPIEZA DE DATOS
# ============================================================
def load_excel(path: Path, sheet_name=0) -> pd.DataFrame:
    """Carga el Excel y normaliza columnas según COLUMN_MAP."""
    if not path.exists():
        raise FileNotFoundError(f"No se encontró el Excel: {path}")

    # Usar header en fila 7 (índice 6) si el archivo es de ONs
    df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl", header=6)

    # Renombrar al esquema interno (amigable). Solo renombra las columnas presentes.
    rename_map = {real: friendly for friendly, real in COLUMN_MAP.items() if real in df.columns}
    df = df.rename(columns=rename_map)

    missing = [k for k in COLUMN_MAP if k not in df.columns]
    if missing:
        print(f"[WARNING] Columnas faltantes (revisar COLUMN_MAP): {missing}")
        print(f"   Columnas disponibles en el Excel: {list(df.columns)}")

    # Tipados básicos: fecha y valor numérico
    if "fecha" in df.columns:
        df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    if "valor" in df.columns:
        df["valor"] = pd.to_numeric(df["valor"], errors="coerce")

    # Solo requerimos fecha para no descartar filas con monto vacío (ej: emisiones en moneda nominal sin ARS)
    df = df.dropna(subset=[c for c in ("fecha",) if c in df.columns])
    return df


# ============================================================
# CONSTRUCCIÓN DEL CONTEXTO PARA EL TEMPLATE
# ============================================================
def _safe_num(x):
    """Convierte a número o devuelve None si no es posible."""
    try:
        v = float(x)
        if pd.isna(v):
            return None
        return v
    except (TypeError, ValueError):
        return None


def _safe_str(x):
    """Convierte a string o devuelve cadena vacía si NaN."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    return str(x).strip()


def _normalize_moneda(x):
    """Normaliza variantes de moneda — unifica todas las formas de 'Dólar Linked'."""
    s = _safe_str(x)
    if not s:
        return ""
    s_lower = (s.lower()
               .replace("á", "a").replace("é", "e").replace("í", "i")
               .replace("ó", "o").replace("ú", "u"))
    if "linked" in s_lower:
        return "Dólar Linked"
    return s


def _format_cell(val):
    """Formatea un valor de celda para JSON-serialización legible."""
    if val is None:
        return ""
    if isinstance(val, float) and pd.isna(val):
        return ""
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, float):
        # Redondear a 4 decimales para evitar JSONs gigantes con flotantes precisos
        return round(val, 4)
    return val


def build_context(df: pd.DataFrame) -> dict:
    """Convierte el DataFrame en un dict que el template Jinja consume.

    Genera el dataset completo + listas de valores únicos para los filtros.
    Toda la lógica de KPIs/gráficos se calcula client-side en JavaScript.
    """
    if df.empty:
        return {
            "titulo": "Informe Mercado Primario (sin datos)",
            "fecha_generacion": fecha_es(),
            "logo_b64": "",
            "dataset_json": "[]",
            "f_years_json": "[]",
            "f_emisores_json": "[]",
            "f_monedas_json": "[]",
            "f_legislaciones_json": "[]",
            "f_monedas_pago_json": "[]",
            "f_estados_json": "[]",
            "tabla_columnas_json": "[]",
            "tabla_filas_json": "[]",
        }

    # Convertir nombres internos de vuelta a nombres originales (para la tabla full)
    columnas_originales = [COLUMN_MAP.get(c, c) for c in df.columns]

    # --- Detectar columnas con nombres complejos por substring ---
    col_fecha_emis = None
    col_monto_nominal = None
    for c in df.columns:
        cl = str(c).lower()
        if col_fecha_emis is None and "emisi" in cl and "liquidaci" in cl:
            col_fecha_emis = c
        if col_monto_nominal is None and "monto nominal" in cl:
            col_monto_nominal = c

    # --- Construir dataset enriquecido + tabla full (orden alineado por idx) ---
    dataset = []
    tabla_filas = []
    for i, (_, row) in enumerate(df.iterrows()):
        fecha = row.get("fecha")
        if pd.isna(fecha):
            continue

        # Excluir emisiones previas a 2020 (no se muestran en el reporte)
        if int(fecha.year) < 2020:
            continue

        # Fecha de Emisión y Liquidación (para gráficos mensuales en moneda de emisión)
        year_emis = None
        month_emis = None
        if col_fecha_emis is not None:
            fe = pd.to_datetime(row.get(col_fecha_emis), errors="coerce")
            if pd.notna(fe):
                year_emis = int(fe.year)
                month_emis = int(fe.month)

        idx = len(dataset)  # índice secuencial post-skip
        item = {
            "idx": idx,
            "year": int(fecha.year),
            "month": int(fecha.month),
            "year_emis": year_emis,
            "month_emis": month_emis,
            "fecha": fecha.strftime("%Y-%m-%d"),
            "emisor": _safe_str(row.get("Sociedad")),
            "moneda": _normalize_moneda(row.get("Moneda")),
            "regimen": _safe_str(row.get("categoria")),
            "tipo_tasa": _safe_str(row.get("subcategoria")),
            "monto_ars": _safe_num(row.get("valor")),
            "monto_usd": _safe_num(row.get("Monto en USD equivalentes")),
            "monto_nominal": _safe_num(row.get(col_monto_nominal)) if col_monto_nominal else None,
            "tir": _safe_num(row.get("TIR inicial")),
            "tna": _safe_num(row.get("TNA inicial")),
            "plazo_meses": _safe_num(row.get("Plazo (meses)")) or _safe_num(row.get("Plazo")),
            "ticker": _safe_str(row.get("Ticker")),
            "estado": _safe_str(row.get("Estado")),
            "serie": _safe_str(row.get("Serie/Clase")),
            "legislacion": _safe_str(row.get("Ley")),
            "moneda_pago": _safe_str(row.get("Pago")),
        }
        dataset.append(item)

        # Fila completa con todas las columnas del Excel (en orden original)
        fila_full = [_format_cell(row[col]) for col in df.columns]
        tabla_filas.append(fila_full)

    # --- Listas para filtros ---
    f_years = sorted({d["year"] for d in dataset}, reverse=True)
    f_emisores = sorted({d["emisor"] for d in dataset if d["emisor"]})
    f_monedas = sorted({d["moneda"] for d in dataset if d["moneda"]})
    f_legislaciones = sorted({d["legislacion"] for d in dataset if d["legislacion"]})
    f_monedas_pago = sorted({d["moneda_pago"] for d in dataset if d["moneda_pago"]})
    f_estados = sorted({d["estado"] for d in dataset if d["estado"]})

    # --- Logo en base64 para autocontener el HTML ---
    logo_path = Path(r"C:\Users\nalperin\OneDrive - One618 Financial Services S.A.U\Escritorio\Logos\ONE\one618 logo - blanco.png")
    logo_b64 = ""
    if logo_path.exists():
        logo_b64 = base64.b64encode(logo_path.read_bytes()).decode("ascii")

    return {
        "titulo": "Informe Mercado Primario",
        "fecha_generacion": fecha_es(),
        "n_emisiones": len(dataset),
        "logo_b64": logo_b64,
        "dataset_json": json.dumps(dataset),
        "f_years_json": json.dumps(f_years),
        "f_emisores_json": json.dumps(f_emisores),
        "f_monedas_json": json.dumps(f_monedas),
        "f_legislaciones_json": json.dumps(f_legislaciones),
        "f_monedas_pago_json": json.dumps(f_monedas_pago),
        "f_estados_json": json.dumps(f_estados),
        "tabla_columnas_json": json.dumps(columnas_originales),
        "tabla_filas_json": json.dumps(tabla_filas, default=str),
    }


# ============================================================
# RENDER
# ============================================================
def render(context: dict, template_name: str = DEFAULT_TEMPLATE) -> str:
    env = Environment(
        loader=FileSystemLoader(str(TEMPLATE_DIR)),
        autoescape=select_autoescape(["html", "xml"]),
    )
    template = env.get_template(template_name)
    return template.render(**context)


def save_output(html: str, prefix: str = "reporte") -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    # Archivo versionado por fecha (historial)
    fname = f"{prefix}_{datetime.now().strftime('%Y-%m-%d')}.html"
    out_path = OUTPUT_DIR / fname
    i = 2
    while out_path.exists():
        out_path = OUTPUT_DIR / f"{prefix}_{datetime.now().strftime('%Y-%m-%d')}_v{i}.html"
        i += 1
    out_path.write_text(html, encoding="utf-8")

    # Archivo fijo "informe_ultimo.html" — siempre apunta al más reciente
    ultimo_path = OUTPUT_DIR / "informe_ultimo.html"
    ultimo_path.write_text(html, encoding="utf-8")

    return out_path


# ============================================================
# CLI
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="Modelo de Visualizaciones — generador de reportes HTML")
    parser.add_argument("--input", default=None,
                        help="Ruta al Excel (default: el .xlsx más reciente en la carpeta Input/)")
    parser.add_argument("--sheet", default=0, help="Hoja del Excel (nombre o índice, default: 0)")
    parser.add_argument("--template", default=DEFAULT_TEMPLATE, help=f"Template a usar (default: {DEFAULT_TEMPLATE})")
    parser.add_argument("--prefix", default="reporte", help="Prefijo del archivo de salida")
    args = parser.parse_args()

    print("=" * 60)
    print("  Modelo de Visualizaciones")
    print("=" * 60)

    # Si no se especificó --input, tomar el .xlsx más reciente de Input/
    if args.input is None:
        excel_path = find_latest_excel(INPUT_DIR)
        print(f"\nInput:    {excel_path.name}  (mas reciente en Input/)")
    else:
        excel_path = Path(args.input)
        print(f"\nInput:    {excel_path}")
    print(f"Template: {args.template}")

    # 1) Cargar
    df = load_excel(excel_path, sheet_name=args.sheet)
    print(f"   {len(df):,} filas cargadas")

    # 2) Construir contexto
    ctx = build_context(df)

    # 3) Renderizar
    html = render(ctx, template_name=args.template)

    # 4) Guardar
    out_path = save_output(html, prefix=args.prefix)
    print(f"\n[OK] Reporte generado: {out_path}")
    print(f"   Tamaño: {len(html) / 1024:.1f} KB")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except FileNotFoundError as e:
        print(f"\n❌ {e}", file=sys.stderr)
        sys.exit(1)
