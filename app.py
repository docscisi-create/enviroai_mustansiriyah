"""
EnviroAI — Air Quality & Weather Analysis System
نظام متكامل لرصد وتحليل جودة الهواء والطقس
"""

# ══════════════════════════════════════════════
#  IMPORTS
# ══════════════════════════════════════════════
import base64
import io
import os
import json
import uuid

import chardet
import dash
from dash import dcc, html, Input, Output, State, no_update
import dash_bootstrap_components as dbc
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from scipy import stats
from scipy.ndimage import gaussian_filter1d

import arabic_reshaper as _ar_reshaper
from bidi.algorithm import get_display as _bidi_display
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table,
    TableStyle, HRFlowable,
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors as rl_colors
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import warnings
warnings.filterwarnings("ignore")


# ══════════════════════════════════════════════
#  SERVER-SIDE SESSION CACHE
# ══════════════════════════════════════════════
_CACHE: dict = {}   # session_id -> {"df": df, "ftype": ftype}


# ══════════════════════════════════════════════
#  AQI SCALE CONSTANTS
# ══════════════════════════════════════════════
AQI_LEVELS = [
    {"range": (0,   50),  "color": "#00E400", "ar": "جيد",               "tc": "#000"},
    {"range": (51,  100), "color": "#FFFF00", "ar": "معتدل",             "tc": "#000"},
    {"range": (101, 150), "color": "#FF7E00", "ar": "غير صحي للحساسين", "tc": "#000"},
    {"range": (151, 200), "color": "#FF0000", "ar": "غير صحي",           "tc": "#fff"},
    {"range": (201, 300), "color": "#8F3F97", "ar": "غير صحي جداً",      "tc": "#fff"},
    {"range": (301, 500), "color": "#7E0023", "ar": "خطير",              "tc": "#fff"},
]

RISK_COLORS = {
    "آمن":    "#00E400",
    "متوسط":  "#FFFF00",
    "تحذير":  "#FF7E00",
    "حرج":    "#FF0000",
    "طوارئ":  "#8F3F97",
}

FACES = ["😊", "🙂", "😐", "😷", "😷", "😷"]

SCALE_DESCS = [
    "جودة الهواء مرضية وتشكل خطرًا قليلاً.",
    "يجب على الأفراد الحساسين تجنب الأنشطة الخارجية المكثفة.",
    "الجمهور والحساسون معرضون لخطر مشاكل تنفسية.",
    "يزداد احتمال حدوث آثار ضارة على الجمهور.",
    "سيتأثر عموم الجمهور. يُنصح بالبقاء في المنازل.",
    "خطير جداً. يجب على الجميع البقاء في الداخل.",
]

REQUIRED_XL = {
    "Date", "Time", "Temperature", "Humidity",
    "AQI", "CO", "SMOKE", "RiskIndex", "RiskLevel",
}

_PDF_FONT_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),  "Arabic.ttf"
)
_PDF_FONTS_READY = False


# ══════════════════════════════════════════════
#  AQI HELPERS
# ══════════════════════════════════════════════
def get_lv(v):
    v = max(0, float(v))
    for lv in AQI_LEVELS:
        if lv["range"][0] <= v <= lv["range"][1]:
            return lv
    return AQI_LEVELS[-1]


def aqi_color(v):
    try:
        return get_lv(v)["color"]
    except Exception:
        return "#ccc"


# ══════════════════════════════════════════════
#  FILE PARSING
# ══════════════════════════════════════════════
def _add_derived(df: pd.DataFrame) -> pd.DataFrame:
    df["Date"]     = pd.to_datetime(df["Date"], errors="coerce")
    df             = df.dropna(subset=["Date"])
    df["DateStr"]  = df["Date"].dt.strftime("%Y-%m-%d")
    df["DayLabel"] = df["Date"].dt.strftime("%-d %b")

    if "Time" in df.columns:
        df["Hour"] = pd.to_datetime(
            df["Time"].astype(str), errors="coerce"
        ).dt.hour
    else:
        df["Hour"] = df["Date"].dt.hour

    if "RiskLevel" not in df.columns:
        def _rl(v):
            if v <= 50:  return "آمن"
            if v <= 100: return "متوسط"
            if v <= 150: return "تحذير"
            if v <= 200: return "حرج"
            return "طوارئ"
        df["RiskLevel"] = df["AQI"].apply(
            lambda v: _rl(float(v)) if pd.notna(v) else "متوسط"
        )
    return df


def parse_file(contents: str, filename: str):
    """Parse uploaded file and return (df, error_msg, ftype)."""
    _, b64 = contents.split(",", 1)
    raw    = base64.b64decode(b64)
    fname  = filename.lower()

    # ── Excel ────────────────────────────────────────────────────────
    if fname.endswith((".xlsx", ".xls", ".xlsm")):
        try:
            sheets = pd.read_excel(io.BytesIO(raw), sheet_name=None)
        except Exception as e:
            return None, f"تعذّر قراءة Excel: {e}", "excel"

        for _, sdf in sheets.items():
            if REQUIRED_XL.issubset(set(sdf.columns)):
                return _add_derived(sdf.copy()), "", "excel"

        return (
            None,
            "لم يُعثر على الأعمدة المطلوبة في ملف Excel.\n"
            "الأعمدة المطلوبة: Date, Time, Temperature, Humidity, "
            "AQI, CO, SMOKE, RiskIndex, RiskLevel",
            "excel",
        )

    # ── CSV ──────────────────────────────────────────────────────────
    try:
        enc  = chardet.detect(raw)["encoding"] or "latin-1"
        text = raw.decode(enc, errors="replace")
        tmp  = pd.read_csv(io.StringIO(text), header=None, on_bad_lines="skip")

        # Detect header row
        hrow = 0
        for i, row in tmp.iterrows():
            vals = [str(v).lower() for v in row if pd.notna(v)]
            if any("date" in v or "time" in v for v in vals):
                hrow = i
                break

        df   = pd.read_csv(io.StringIO(text), header=None,
                           skiprows=hrow + 1, on_bad_lines="skip")
        cols = [
            str(c).strip() if pd.notna(c) else f"c{i}"
            for i, c in enumerate(tmp.iloc[hrow].tolist()[: len(df.columns)])
        ]
        df.columns = cols

        # Auto-rename columns
        rename = {}
        for c in df.columns:
            cl = c.lower()
            if ("date" in cl or "time" in cl) and "Date" not in rename.values():
                rename[c] = "Date"
            elif "aqi" in cl and "AQI" not in rename.values():
                rename[c] = "AQI"
            elif "2.5" in cl or "pm2.5" in cl or "pm 2.5" in cl:
                rename[c] = "PM25"
            elif "pm 10" in cl or "pm10" in cl:
                rename[c] = "PM10"
            elif ("pm 1" in cl or "pm1" in cl) and "PM1" not in rename.values():
                rename[c] = "PM1"
            elif "temp" in cl and "Temperature" not in rename.values():
                rename[c] = "Temperature"
            elif "hum" in cl and "Humidity" not in rename.values():
                rename[c] = "Humidity"
            elif "dew" in cl:
                rename[c] = "DewPoint"

        df = df.rename(columns=rename)

        if "AQI" not in df.columns or "Date" not in df.columns:
            return None, "ملف CSV لا يحتوي أعمدة AQI أو Date المطلوبة.", "csv"

        for c in ["AQI", "PM25", "PM10", "PM1", "Temperature", "Humidity"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")

        if "CO"        not in df.columns: df["CO"]        = np.nan
        if "SMOKE"     not in df.columns: df["SMOKE"]     = df.get("PM25", np.nan)
        if "RiskIndex" not in df.columns:
            df["RiskIndex"] = df["AQI"].apply(
                lambda v: float(v) * 0.9 if pd.notna(v) else np.nan
            )

        df = df[df["AQI"].notna() & (df["AQI"] > 0)]
        return _add_derived(df), "", "csv"

    except Exception as e:
        return None, f"تعذّر قراءة CSV: {e}", "csv"


def daily_stats(df: pd.DataFrame) -> pd.DataFrame:
    agg = dict(
        DayLabel =("DayLabel", "first"),
        AQI_mean =("AQI",      "mean"),
        AQI_max  =("AQI",      "max"),
    )
    for col, key in [
        ("Temperature", "Temp_mean"),
        ("Humidity",    "Hum_mean"),
        ("CO",          "CO_mean"),
        ("SMOKE",       "SMOKE_mean"),
        ("PM25",        "PM25_mean"),
        ("PM10",        "PM10_mean"),
    ]:
        if col in df.columns:
            agg[key] = (col, "mean")

    return df.groupby("DateStr").agg(**agg).round(2).reset_index()


# ══════════════════════════════════════════════
#  CHART HELPERS
# ══════════════════════════════════════════════
def _base(title: str) -> dict:
    return dict(
        title=dict(
            text=title,
            font=dict(size=13, color="#1a1a2e"),
            x=0.5,
            xanchor="center",
        ),
        paper_bgcolor="white",
        plot_bgcolor="#f8f9fa",
        font=dict(family="Cairo,Arial", color="#333"),
        margin=dict(t=55, b=45, l=55, r=35),
        legend=dict(
            bgcolor="rgba(255,255,255,.9)",
            bordercolor="#ddd",
            borderwidth=1,
        ),
        hovermode="x unified",
        xaxis=dict(gridcolor="#ebebeb", tickangle=-30),
        yaxis=dict(gridcolor="#ebebeb"),
    )


def ch_aqi_bar(daily: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=daily["DayLabel"],
        y=daily["AQI_mean"],
        marker_color=[aqi_color(v) for v in daily["AQI_mean"]],
        marker_line_color="rgba(0,0,0,.1)",
        marker_line_width=1,
        text=daily["AQI_mean"].round(0).astype(int),
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>AQI: %{y:.0f}<extra></extra>",
    ))
    for y, lbl, col in [
        (50,  "جيد 50",       "#00E400"),
        (100, "معتدل 100",    "#b8a000"),
        (150, "غير صحي 150", "#FF7E00"),
        (200, "خطر 200",      "#FF0000"),
        (300, "طوارئ 300",    "#8F3F97"),
    ]:
        fig.add_hline(
            y=y, line_dash="dot", line_color=col,
            line_width=1.3, opacity=0.8,
            annotation_text=lbl, annotation_position="right",
            annotation_font_size=9, annotation_font_color=col,
        )
    fig.update_layout(
        **_base("متوسط AQI اليومي  |  Daily Average AQI"),
        yaxis_title="AQI",
        yaxis_range=[0, max(daily["AQI_mean"].max() * 1.25, 230)],
    )
    return fig


def ch_temp_hum(daily: pd.DataFrame) -> go.Figure:
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    if "Temp_mean" in daily.columns:
        fig.add_trace(go.Scatter(
            x=daily["DayLabel"], y=daily["Temp_mean"],
            mode="lines+markers", name="حرارة (°م)",
            line=dict(color="#e67e22", width=2.5),
            marker=dict(size=7),
        ), secondary_y=False)
    if "Hum_mean" in daily.columns:
        fig.add_trace(go.Bar(
            x=daily["DayLabel"], y=daily["Hum_mean"],
            name="رطوبة (%)", opacity=0.28, marker_color="#3498db",
        ), secondary_y=True)
    fig.update_yaxes(
        title_text="°م", secondary_y=False,
        title_font_color="#e67e22", gridcolor="#ebebeb",
    )
    fig.update_yaxes(
        title_text="%", secondary_y=True,
        title_font_color="#3498db", showgrid=False,
    )
    fig.update_layout(**_base("درجة الحرارة والرطوبة  |  Temperature & Humidity"))
    return fig


def ch_pm(daily: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    if "PM25_mean" in daily.columns:
        fig.add_trace(go.Scatter(
            x=daily["DayLabel"], y=daily["PM25_mean"],
            mode="lines+markers", name="PM 2.5 (μg/m³)",
            line=dict(color="#e74c3c", width=2.5), marker=dict(size=7),
        ))
    if "PM10_mean" in daily.columns:
        fig.add_trace(go.Scatter(
            x=daily["DayLabel"], y=daily["PM10_mean"],
            mode="lines+markers", name="PM 10 (μg/m³)",
            line=dict(color="#8F3F97", width=2, dash="dash"),
            marker=dict(size=6),
        ))
    fig.add_hline(
        y=15, line_dash="dot", line_color="#e74c3c", line_width=1.2,
        annotation_text="WHO PM2.5 ≤15",
        annotation_font_size=9, annotation_font_color="#e74c3c",
    )
    fig.add_hline(
        y=45, line_dash="dot", line_color="#8F3F97", line_width=1.2,
        annotation_text="WHO PM10 ≤45",
        annotation_font_size=9, annotation_font_color="#8F3F97",
    )
    fig.update_layout(
        **_base("جسيمات الغبار  |  Particulate Matter (PM)"),
        yaxis_title="μg/m³",
    )
    return fig


def ch_pollutants(daily: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    if "CO_mean" in daily.columns and daily["CO_mean"].notna().any():
        fig.add_trace(go.Bar(
            x=daily["DayLabel"], y=daily["CO_mean"],
            name="CO (ppm)", marker_color="#8F3F97",
        ))
    smoke_col = "SMOKE_mean" if "SMOKE_mean" in daily.columns else "PM25_mean"
    if smoke_col in daily.columns and daily[smoke_col].notna().any():
        fig.add_trace(go.Bar(
            x=daily["DayLabel"], y=daily[smoke_col],
            name="SMOKE / PM2.5", marker_color="#c0392b",
        ))
    fig.update_layout(**_base("الملوثات  |  Pollutants"), barmode="group")
    return fig


def ch_donut(df: pd.DataFrame) -> go.Figure:
    counts = df["RiskLevel"].value_counts()
    order  = ["آمن", "متوسط", "تحذير", "حرج", "طوارئ"]
    vals   = [counts.get(k, 0) for k in order]
    fig = go.Figure(go.Pie(
        labels=order, values=vals, hole=0.58,
        marker=dict(
            colors=[RISK_COLORS[k] for k in order],
            line=dict(color="#fff", width=2),
        ),
        hovertemplate="<b>%{label}</b><br>%{value:,} قراءة  %{percent}<extra></extra>",
    ))
    fig.add_annotation(
        text=f"<b>{sum(vals):,}</b><br>قراءة",
        x=0.5, y=0.5, showarrow=False,
        font=dict(size=14, color="#333"),
    )
    fig.update_layout(**_base("توزيع مستويات الخطر  |  Risk Distribution"))
    return fig


def ch_hourly(df: pd.DataFrame) -> go.Figure:
    h = (
        df.dropna(subset=["Hour"])
        .groupby("Hour")["AQI"]
        .mean()
        .reindex(range(24))
        .reset_index()
    )
    h.columns = ["Hour", "AQI"]
    fig = go.Figure(go.Bar(
        x=h["Hour"],
        y=h["AQI"].fillna(0),
        marker_color=[
            aqi_color(v) if pd.notna(v) else "#ddd" for v in h["AQI"]
        ],
        hovertemplate="الساعة %{x}:00<br>AQI: %{y:.0f}<extra></extra>",
    ))
    fig.add_hline(
        y=100, line_dash="dot", line_color="#FF7E00", line_width=1.5,
        annotation_text="100",
        annotation_font_color="#FF7E00", annotation_font_size=9,
    )
    lyt = _base("متوسط AQI لكل ساعة  |  Hourly AQI Pattern")
    lyt["xaxis"] = dict(
        tickmode="array",
        tickvals=[0, 3, 6, 9, 12, 15, 18, 21],
        ticktext=["12AM", "3AM", "6AM", "9AM", "12PM", "3PM", "6PM", "9PM"],
        tickangle=0,
        gridcolor="#ebebeb",
    )
    fig.update_layout(**lyt)
    return fig


def ch_trend(daily: pd.DataFrame) -> go.Figure:
    aqi   = daily["AQI_mean"].values
    x     = np.arange(len(aqi))
    slope, intercept, r, p, se = stats.linregress(x, aqi)
    trend = slope * x + intercept

    xf = np.arange(len(aqi), len(aqi) + 7)
    yf = np.clip(slope * xf + intercept, 0, 500)
    ci = 1.96 * se * np.sqrt(
        1 + 1 / len(x) + (xf - x.mean()) ** 2 / ((x - x.mean()) ** 2).sum()
    )
    smooth = gaussian_filter1d(aqi, sigma=1.0)

    fig = go.Figure()

    for lo, hi, col in [
        (0,   50,  "#00E400"),
        (50,  100, "#FFFF00"),
        (100, 150, "#FF7E00"),
        (150, 200, "#FF0000"),
        (200, 300, "#8F3F97"),
        (300, 500, "#7E0023"),
    ]:
        fig.add_hrect(y0=lo, y1=hi, fillcolor=col, opacity=0.07, line_width=0)

    fig.add_trace(go.Scatter(
        x=list(daily["DayLabel"]), y=smooth,
        mode="lines", name="AQI منعّم",
        line=dict(color="#3498db", width=2.5),
    ))
    fig.add_trace(go.Scatter(
        x=list(daily["DayLabel"]), y=aqi,
        mode="markers", name="AQI فعلي",
        marker=dict(
            color=[aqi_color(v) for v in aqi],
            size=11,
            line=dict(color="white", width=1.5),
        ),
    ))

    all_x = list(daily["DayLabel"]) + [f"+{j+1}" for j in range(7)]
    fig.add_trace(go.Scatter(
        x=all_x, y=np.append(trend, yf),
        mode="lines",
        name=f"اتجاه ({slope:+.1f}/يوم)",
        line=dict(color="#FF7E00", width=2, dash="dash"),
    ))
    fig.add_trace(go.Scatter(
        x=[f"+{j+1}" for j in range(7)], y=yf,
        mode="markers+text",
        marker=dict(
            color=[aqi_color(v) for v in yf],
            size=13, symbol="diamond",
            line=dict(color="white", width=1.5),
        ),
        text=[f"{v:.0f}" for v in yf],
        textposition="top center",
        textfont=dict(size=9),
        name="تنبؤ 7 أيام",
    ))
    fig.add_trace(go.Scatter(
        x=[f"+{j+1}" for j in range(7)] + [f"+{j+1}" for j in range(6, -1, -1)],
        y=list(np.clip(yf + ci, 0, 600)) + list(np.clip(yf - ci, 0, 600)),
        fill="toself",
        fillcolor="rgba(143,63,151,.15)",
        line=dict(color="rgba(0,0,0,0)"),
        name="فترة ثقة 95%",
    ))
    fig.add_vline(x=len(aqi) - 0.5, line_dash="dot", line_color="#aaa", line_width=1)
    fig.update_layout(
        **_base(f"تحليل الاتجاه والتنبؤ  |  R²={r**2:.3f}  p={p:.3f}"),
        xaxis_tickangle=-30,
    )
    return fig


# ══════════════════════════════════════════════
#  UI HELPERS
# ══════════════════════════════════════════════
def kpi_row(df: pd.DataFrame, ftype: str):
    aqi_m = df["AQI"].mean()
    lv    = get_lv(aqi_m)
    items = [
        ("📊", "إجمالي القراءات",  f"{len(df):,}",             "#2c5f9e"),
        ("📅", "أيام الرصد",       str(df["DateStr"].nunique()), "#27ae60"),
        ("🌡️", "متوسط الحرارة",
         f"{df['Temperature'].mean():.1f}°م"
         if "Temperature" in df.columns and df["Temperature"].notna().any()
         else "—",
         "#e67e22"),
        ("💨", "متوسط AQI",       f"{aqi_m:.1f}",
         lv["color"] if lv["color"] != "#FFFF00" else "#b8a000"),
        ("⚠️", "أعلى AQI مُسجّل", f"{df['AQI'].max():.0f}",    "#c0392b"),
        ("📁", "نوع الملف",
         "CSV (AirLink)" if ftype == "csv" else "Excel",         "#555"),
    ]
    cols = []
    for icon, label, val, color in items:
        cols.append(dbc.Col(
            dbc.Card(dbc.CardBody([
                html.Div(icon, style={"fontSize": "1.6rem", "textAlign": "center", "lineHeight": "1"}),
                html.Div(val,  style={
                    "fontWeight": "700", "fontSize": "1.15rem",
                    "color": color, "textAlign": "center", "margin": "4px 0 2px",
                }),
                html.Div(label, style={"fontSize": "0.68rem", "color": "#777", "textAlign": "center"}),
            ], style={"padding": "10px 6px"}),
            style={
                "borderRadius": "10px",
                "border": f"2px solid {color}22",
                "backgroundColor": "#fff",
                "height": "100%",
            }),
            width=2,
        ))
    return dbc.Row(cols, className="g-2 mb-3")


def aqi_guide_card():
    rows = []
    for i, lv in enumerate(AQI_LEVELS):
        lo, hi = lv["range"]
        rows.append(dbc.Row([
            dbc.Col(html.Div(FACES[i], style={
                "background": lv["color"], "borderRadius": "10px",
                "fontSize": "2rem", "textAlign": "center",
                "padding": "5px 0", "minWidth": "56px", "lineHeight": "1.2",
            }), width="auto"),
            dbc.Col([
                html.Span(f"{lo}–{hi}  ", style={"fontWeight": "700", "fontSize": ".9rem"}),
                html.Span(lv["ar"],       style={"fontWeight": "700", "fontSize": ".9rem", "color": "#222"}),
                html.Br(),
                html.Small(SCALE_DESCS[i], style={"color": "#666"}),
            ]),
        ], align="center", className="py-2",
           style={"borderBottom": "1px solid #eee", "margin": "0", "paddingRight": "8px"}))
    return dbc.Card([
        dbc.CardHeader(
            html.B("دليل مؤشر جودة الهواء  |  AQI Reference Guide",
                   className="d-block text-center"),
            style={"background": "#eef2ff", "borderBottom": "2px solid #2c5f9e"},
        ),
        dbc.CardBody(rows, style={"padding": "0 8px"}),
    ], style={"borderRadius": "12px", "overflow": "hidden"})


def welcome_screen():
    return dbc.Container([
        dbc.Row(dbc.Col([
            html.Div([
                html.Div("🌍", style={"fontSize": "4rem", "textAlign": "center"}),
                html.H3("نظام EnviroAI لتحليل جودة الهواء",
                        className="text-center fw-bold mt-2",
                        style={"color": "#1a3a6b"}),
                html.P("Air Quality & Weather Analysis System",
                       className="text-center text-muted mb-4"),
            ]),
            dbc.Card([
                dbc.CardHeader(
                    html.B("📋 كيفية الاستخدام  |  How to Use"),
                    style={"background": "#eef2ff", "borderBottom": "2px solid #2c5f9e"},
                ),
                dbc.CardBody([
                    dbc.Row([
                        dbc.Col([
                            html.H6("📤 الخطوة 1 — رفع الملف", style={"color": "#1a3a6b"}),
                            html.P("اسحب وأفلت ملف البيانات في المربع أعلاه، أو انقر لاختياره من جهازك."),
                            html.Hr(),
                            html.H6("📊 الخطوة 2 — استعراض التحليل", style={"color": "#1a3a6b"}),
                            html.P("تظهر لوحة التحكم تلقائياً بعد رفع الملف مع 6 تبويبات تفاعلية."),
                            html.Hr(),
                            html.H6("🗂️ أنواع الملفات المدعومة", style={"color": "#1a3a6b", "marginBottom": "8px"}),
                            dbc.Badge("✅ Excel (.xlsx)  — بيانات المحطة الرئيسية",
                                      color="primary", className="me-2 mb-2 d-block"),
                            dbc.Badge("✅ CSV — تصدير AirLink للمحطة المتنقلة",
                                      color="success", className="mb-1 d-block"),
                        ], width=5),
                        dbc.Col([
                            html.H6("📈 ماذا يعرض النظام؟", style={"color": "#1a3a6b"}),
                            html.Ul([
                                html.Li("📊 نظرة عامة — مؤشرات AQI اليومي + درجة الحرارة + الملوثات"),
                                html.Li("🌫️ PM2.5 / PM10 — جسيمات الغبار مع حدود منظمة الصحة العالمية"),
                                html.Li("⏰ نمط ساعي — كيف يتغير AQI خلال 24 ساعة"),
                                html.Li("⚠️ مستوى الخطر — توزيع الفئات (آمن / تحذير / حرج ...)"),
                                html.Li("📈 تنبؤ — تحليل الاتجاه + توقعات 7 أيام"),
                                html.Li("📋 دليل AQI — شرح المقياس الرسمي بالألوان"),
                            ], style={"fontSize": ".88rem", "paddingRight": "1.2rem"}),
                        ], width=7),
                    ]),
                ]),
            ], className="mb-4", style={"borderRadius": "12px"}),
            html.H6("🎨 مقياس AQI الرسمي:", className="mb-3", style={"color": "#1a3a6b"}),
            dbc.Row([
                dbc.Col(html.Div([
                    html.Div(FACES[i], style={
                        "background": lv["color"], "color": lv["tc"],
                        "borderRadius": "10px", "textAlign": "center",
                        "padding": "10px 4px", "fontSize": "1.4rem", "lineHeight": "1",
                    }),
                    html.Div(
                        f"{lv['range'][0]}–{lv['range'][1]}",
                        style={
                            "textAlign": "center", "fontSize": ".72rem",
                            "fontWeight": "600",
                            "color": lv["color"] if lv["color"] != "#FFFF00" else "#7a6a00",
                            "marginTop": "4px",
                        },
                    ),
                    html.Div(lv["ar"], style={"textAlign": "center", "fontSize": ".7rem", "color": "#555"}),
                ]), width=2)
                for i, lv in enumerate(AQI_LEVELS)
            ], className="g-2"),
        ], width={"size": 10, "offset": 1}), className="my-4"),
    ], fluid=True)


# ══════════════════════════════════════════════
#  PDF REPORT GENERATOR
# ══════════════════════════════════════════════
def _pdf_ensure_fonts():
    global _PDF_FONTS_READY
    if not _PDF_FONTS_READY:
        pdfmetrics.registerFont(TTFont("Ar", _PDF_FONT_PATH))
        _PDF_FONTS_READY = True


def _a(text: str) -> str:
    """Arabic reshape + bidi for PDF."""
    return _bidi_display(_ar_reshaper.reshape(str(text)))


def _ps(name: str, **kw) -> ParagraphStyle:
    base = dict(
        fontName="Ar", fontSize=10, alignment=TA_RIGHT,
        leading=20, textColor=rl_colors.HexColor("#222"),
    )
    base.update(kw)
    return ParagraphStyle(name, **base)


def build_pdf_report(
    df: pd.DataFrame,
    daily: pd.DataFrame,
    ftype: str,
    filename: str = "بيانات",
) -> bytes:
    _pdf_ensure_fonts()
    from scipy import stats as _sp

    # ── Compute statistics ───────────────────────────────────────────
    aqi    = df["AQI"]
    mu     = aqi.mean();  mx = aqi.max();  mn = aqi.min()
    std    = aqi.std();   cv = std / mu * 100
    n_days = df["DateStr"].nunique()
    n_read = len(df)
    d0     = df["DateStr"].min()
    d1     = df["DateStr"].max()
    total  = len(df)
    cnt    = df["RiskLevel"].value_counts()
    pct    = lambda k: cnt.get(k, 0) / total * 100

    x      = np.arange(len(daily));  y = daily["AQI_mean"].values
    slope, intercept, r, p_val, _ = _sp.linregress(x, y)
    tdir   = "تصاعدياً" if slope > 0.5 else ("تنازلياً" if slope < -0.5 else "مستقراً نسبياً")
    tsig   = ("ذو دلالة إحصائية (p<0.05)"
              if p_val < 0.05
              else "غير ذي دلالة إحصائية (p≥0.05)")

    hrly   = df.dropna(subset=["Hour"]).groupby("Hour")["AQI"].mean()
    ph     = int(hrly.idxmax());  lh = int(hrly.idxmin())
    fmth   = lambda h: f"{'12' if h == 12 else h % 12 or h}:00 {'م' if h >= 12 else 'ص'}"

    # AQI level colour
    if   mu <= 50:  lvl="جيد";              lvc=rl_colors.HexColor("#00b300"); rc_=rl_colors.HexColor("#27ae60")
    elif mu <= 100: lvl="معتدل";            lvc=rl_colors.HexColor("#b8a000"); rc_=rl_colors.HexColor("#b8a000")
    elif mu <= 150: lvl="غير صحي للحساسين"; lvc=rl_colors.HexColor("#FF7E00"); rc_=rl_colors.HexColor("#FF7E00")
    elif mu <= 200: lvl="غير صحي";          lvc=rl_colors.HexColor("#FF0000"); rc_=rl_colors.HexColor("#c0392b")
    else:           lvl="خطير";             lvc=rl_colors.HexColor("#8F3F97"); rc_=rl_colors.HexColor("#8F3F97")

    # Conclusion text
    if   mu <= 50:
        concl = "تُعدّ جودة الهواء خلال فترة الدراسة جيدة وآمنة للجمهور بصفة عامة."
        recc  = "مواصلة المراقبة الدورية والحفاظ على المستويات الحالية."
    elif mu <= 100:
        concl = "جودة الهواء مقبولة غير أن بعض الملوثات قد تشكّل خطراً على الفئات الحساسة."
        recc  = "ينصح الأفراد الحساسون بتقليل الأنشطة الخارجية في ساعات الذروة."
    elif mu <= 150:
        concl = "جودة الهواء غير صحية للفئات الحساسة وتتطلب إجراءات وقائية."
        recc  = "يُوصى بتقليل التعرض للهواء الخارجي وارتداء الكمامات في ساعات الذروة."
    else:
        concl = "جودة الهواء غير صحية وتستدعي تدخلاً عاجلاً من الجهات المختصة."
        recc  = "يُوصى بالإعلان عن تنبيه بيئي وتجنّب الأنشطة الخارجية والتنسيق مع الجهات المختصة."

    # ── Colours ──────────────────────────────────────────────────────
    C_BLUE  = rl_colors.HexColor("#1a3a6b")
    C_LBLUE = rl_colors.HexColor("#2c5f9e")
    C_BG    = rl_colors.HexColor("#eef2ff")
    C_LINE  = rl_colors.HexColor("#c8d8f8")
    C_RED   = rl_colors.HexColor("#c0392b")
    C_GREEN = rl_colors.HexColor("#27ae60")
    C_PURP  = rl_colors.HexColor("#8F3F97")

    # ── Paragraph styles ─────────────────────────────────────────────
    ST = _ps("title", fontSize=18, alignment=TA_CENTER, textColor=C_BLUE,  leading=28)
    SS = _ps("sub",   fontSize=10, alignment=TA_CENTER, textColor=C_LBLUE, leading=16)
    SM = _ps("mini",  fontSize=8,  alignment=TA_CENTER, textColor=rl_colors.HexColor("#888"), leading=14)
    SH = _ps("head",  fontSize=12, textColor=C_BLUE,   leading=24, spaceBefore=10, spaceAfter=4)
    SB = _ps("body",  fontSize=10, leading=22, spaceAfter=4)
    SF = _ps("foot",  fontSize=7.5, alignment=TA_CENTER, textColor=rl_colors.HexColor("#888"), leading=13)

    def HR():
        return HRFlowable(width="100%", thickness=0.8, color=C_LINE,
                          spaceAfter=4, spaceBefore=4)

    # ── Build story ───────────────────────────────────────────────────
    buf  = io.BytesIO()
    doc  = SimpleDocTemplate(
        buf, pagesize=A4,
        rightMargin=2 * cm, leftMargin=2 * cm,
        topMargin=1.8 * cm, bottomMargin=1.8 * cm,
        title="التقرير التحليلي — EnviroAI",
    )
    story = []

    # Header banner
    hdr = Table([[
        Paragraph(_a("مركز رصد جودة الهواء والبيئة\nقسم مراقبة نوعية الهواء"),
                  _ps("hh", fontSize=10, alignment=TA_RIGHT, textColor=rl_colors.white, leading=18)),
        Paragraph(_a("EnviroAI"),
                  _ps("hh2", fontSize=16, alignment=TA_CENTER, textColor=rl_colors.white, leading=24)),
    ]], colWidths=[11 * cm, 5.5 * cm])
    hdr.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), C_BLUE),
        ("TOPPADDING",    (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING",   (0, 0), (-1, -1), 12),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story += [
        hdr,
        Spacer(1, 0.4 * cm),
        Paragraph(_a("التقرير التحليلي الآلي"), ST),
        Paragraph(_a("Air Quality Automated Analytical Report"), SS),
        Paragraph(_a(f"الملف: {filename}"), SM),
        Spacer(1, 0.3 * cm),
        HR(),
    ]

    # Section 1 — Overview
    story.append(Paragraph(_a("١. نظرة عامة على البيانات"), SH))
    story.append(Paragraph(_a(
        f"يشتمل هذا التقرير على تحليل بيانات جودة الهواء المرصودة خلال الفترة من {d0} إلى {d1}، "
        f"بإجمالي {n_read:,} قراءة موزّعة على {n_days} يوماً من الرصد المستمر. "
        f"مصدر البيانات: {'محطة AirLink المتنقلة (CSV)' if ftype == 'csv' else 'المحطة الثابتة (Excel)'}."
    ), SB))
    story.append(HR())

    # Section 2 — Descriptive stats
    story.append(Paragraph(_a("٢. الإحصاء الوصفي لمؤشر جودة الهواء (AQI)"), SH))
    tbl2 = Table([
        [_a("المتوسط"), _a("الحد الأقصى"), _a("الحد الأدنى"), _a("الانحراف المعياري"), _a("معامل التباين")],
        [f"{mu:.1f}", f"{mx:.0f}", f"{mn:.0f}", f"{std:.1f}", f"{cv:.1f}%"],
        [_a(lvl), _a(get_lv(mx)["ar"]), _a(get_lv(mn)["ar"]), "", ""],
    ], colWidths=[3.3 * cm] * 5)
    tbl2.setStyle(TableStyle([
        ("FONTNAME",      (0, 0), (-1, -1), "Ar"),
        ("FONTSIZE",      (0, 0), (-1,  0), 9),
        ("FONTSIZE",      (0, 1), (-1,  1), 14),
        ("FONTSIZE",      (0, 2), (-1,  2), 8),
        ("BACKGROUND",    (0, 0), (-1,  0), C_BG),
        ("TEXTCOLOR",     (0, 0), (-1,  0), C_BLUE),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",          (0, 0), (-1, -1), 0.5, C_LINE),
        ("TOPPADDING",    (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("TEXTCOLOR",     (0, 1), (0,   1), lvc),
        ("TEXTCOLOR",     (1, 1), (1,   1), C_RED),
        ("TEXTCOLOR",     (2, 1), (2,   1), C_GREEN),
        ("TEXTCOLOR",     (3, 1), (3,   1), C_PURP),
    ]))
    story.append(tbl2)
    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph(_a(
        f"يبلغ المتوسط الحسابي {mu:.1f} ضمن فئة «{lvl}». "
        f"الانحراف المعياري {std:.1f} ومعامل التباين {cv:.1f}%، مما يشير إلى "
        + ("تذبذب واضح." if cv > 40 else "استقرار نسبي.")
    ), SB))
    story.append(HR())

    # Section 3 — Risk distribution
    story.append(Paragraph(_a("٣. توزيع مستويات الخطر"), SH))
    risk_rows = [[_a("المستوى"), _a("القراءات"), _a("النسبة٪")]]
    for nm in ["آمن", "متوسط", "تحذير", "حرج", "طوارئ"]:
        risk_rows.append([_a(nm), f"{cnt.get(nm, 0):,}", f"{pct(nm):.1f}%"])
    tbl3 = Table(risk_rows, colWidths=[4 * cm, 4 * cm, 4 * cm])
    tbl3.setStyle(TableStyle([
        ("FONTNAME",      (0, 0), (-1, -1), "Ar"),
        ("FONTSIZE",      (0, 0), (-1, -1), 10),
        ("BACKGROUND",    (0, 0), (-1,  0), C_BG),
        ("TEXTCOLOR",     (0, 0), (-1,  0), C_BLUE),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",          (0, 0), (-1, -1), 0.5, C_LINE),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [rl_colors.white, C_BG]),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(tbl3)
    story.append(Spacer(1, 0.25 * cm))
    story.append(Paragraph(_a(
        f"نسبة القراءات الخطرة (حرج+طوارئ): {pct('حرج') + pct('طوارئ'):.1f}%  |  "
        f"نسبة الآمنة: {pct('آمن'):.1f}%."
    ), SB))
    if "PM25" in df.columns and df["PM25"].notna().any():
        pm25m  = df["PM25"].mean()
        pmex   = (df["PM25"] > 15).sum()
        story.append(Paragraph(_a(
            f"جسيمات PM2.5: متوسط {pm25m:.1f} μg/m³، تجاوزت حد منظمة الصحة العالمية "
            f"في {pmex:,} قراءة ({pmex / total * 100:.1f}% من الوقت)."
        ), SB))
    story.append(HR())

    # Section 4 — Trend
    story.append(Paragraph(_a("٤. تحليل الاتجاه الزمني (OLS Linear Regression)"), SH))
    story.append(Paragraph(_a(
        f"جرى تطبيق نموذج الانحدار الخطي البسيط OLS على المتوسطات اليومية. "
        f"النتيجة: الاتجاه {tdir} بمعدل {abs(slope):.2f} وحدة/يوم. "
        f"R² = {r**2:.3f}، قيمة p = {p_val:.4f}، النموذج {tsig}."
    ), SB))
    story.append(Paragraph(_a(
        f"ذروة التلوث عند {fmth(ph)} | أدنى قيمة عند {fmth(lh)}. "
        "يعكس هذا النمط تأثير الأنشطة البشرية والظروف الجوية على تراكم الملوثات."
    ), SB))
    ols_t = Table([
        [_a("المعامل"),              _a("القيمة")],
        [_a("الميل (slope)"),        f"{slope:.4f}"],
        [_a("التقاطع (intercept)"),  f"{intercept:.2f}"],
        [_a("معامل الارتباط r"),     f"{r:.4f}"],
        [_a("معامل التحديد R²"),     f"{r**2:.4f}"],
        [_a("p-value"),               f"{p_val:.4f}"],
    ], colWidths=[8 * cm, 4.5 * cm])
    ols_t.setStyle(TableStyle([
        ("FONTNAME",      (0, 0), (-1, -1), "Ar"),
        ("FONTSIZE",      (0, 0), (-1, -1), 9),
        ("BACKGROUND",    (0, 0), (-1,  0), C_BG),
        ("TEXTCOLOR",     (0, 0), (-1,  0), C_BLUE),
        ("ALIGN",         (0, 0), (0,  -1), "RIGHT"),
        ("ALIGN",         (1, 0), (1,  -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",          (0, 0), (-1, -1), 0.5, C_LINE),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [rl_colors.white, C_BG]),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(Spacer(1, 0.2 * cm))
    story.append(ols_t)
    story.append(HR())

    # Section 5 — Conclusion
    story.append(Paragraph(_a("٥. الخلاصة والتوصيات"), SH))
    ct = Table([
        [Paragraph(_a("الخلاصة: " + concl),
                   _ps("cx", fontSize=10, leading=20, textColor=rl_colors.white))],
        [Paragraph(_a("التوصية: " + recc),
                   _ps("cy", fontSize=10, leading=20, textColor=rl_colors.white))],
    ], colWidths=[16.5 * cm])
    ct.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1,  0), rc_),
        ("BACKGROUND",    (0, 1), (-1,  1), rc_),
        ("FONTNAME",      (0, 0), (-1, -1), "Ar"),
        ("TOPPADDING",    (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING",   (0, 0), (-1, -1), 14),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 14),
        ("LINEBELOW",     (0, 0), (-1,  0), 1, rl_colors.white),
    ]))
    story.append(ct)
    story.append(Spacer(1, 0.5 * cm))
    story.append(HRFlowable(width="100%", thickness=0.8, color=C_LINE))
    story.append(Paragraph(_a(
        "الخوارزميات: الإحصاء الوصفي (Descriptive Statistics) — "
        "الانحدار الخطي OLS — التحليل الساعي (Hourly Pattern Analysis)."
    ), SF))
    story.append(Paragraph(_a(
        "EnviroAI — مركز رصد جودة الهواء والبيئة — قسم مراقبة نوعية الهواء"
    ), _ps("ff", fontSize=7, alignment=TA_CENTER,
           textColor=rl_colors.HexColor("#aaa"), leading=12)))

    doc.build(story)
    return buf.getvalue()


# ══════════════════════════════════════════════
#  HTML REPORT CARD (shown inside dashboard)
# ══════════════════════════════════════════════
def generate_report(df: pd.DataFrame, daily: pd.DataFrame, ftype: str):
    from scipy import stats as sp

    aqi        = df["AQI"]
    aqi_mean   = aqi.mean()
    aqi_max    = aqi.max()
    aqi_min    = aqi.min()
    aqi_std    = aqi.std()
    aqi_cv     = aqi_std / aqi_mean * 100
    n_days     = df["DateStr"].nunique()
    n_readings = len(df)
    date_start = df["DateStr"].min()
    date_end   = df["DateStr"].max()
    lv_mean    = get_lv(aqi_mean)
    total      = len(df)
    counts     = df["RiskLevel"].value_counts()
    pct        = lambda k: counts.get(k, 0) / total * 100

    # Trend
    x = np.arange(len(daily));  y = daily["AQI_mean"].values
    slope, intercept, r, p_val, _ = sp.linregress(x, y)
    trend_dir = "تصاعدياً" if slope > 0.5 else ("تنازلياً" if slope < -0.5 else "مستقراً نسبياً")
    trend_sig = ("ذو دلالة إحصائية (p < 0.05)"
                 if p_val < 0.05
                 else "غير ذي دلالة إحصائية (p ≥ 0.05)")

    hourly = df.dropna(subset=["Hour"]).groupby("Hour")["AQI"].mean()
    peak_h = int(hourly.idxmax());  low_h = int(hourly.idxmin())
    fmt_h  = lambda h: f"{'12' if h == 12 else h % 12 or 12}:00 {'م' if h >= 12 else 'ص'}"

    # PM items
    pm_items = []
    if "PM25" in df.columns and df["PM25"].notna().any():
        pm25_m    = df["PM25"].mean()
        pm_exceed = (df["PM25"] > 15).sum()
        pm_pct    = pm_exceed / total * 100
        pm_items.append(html.Li([
            html.B("جسيمات PM2.5: "),
            f"بلغ متوسطها {pm25_m:.1f} μg/m³، وتجاوزت الحد الآمن لمنظمة الصحة العالمية "
            f"(15 μg/m³) في {pm_exceed:,} قراءة ({pm_pct:.1f}٪ من الوقت).",
        ]))

    # Conclusion
    if aqi_mean <= 50:
        conclusion = "تُعدّ جودة الهواء خلال فترة الدراسة جيدة وآمنة للجمهور بصفة عامة."
        rec        = "مواصلة المراقبة الدورية والحفاظ على المستويات الحالية."
        rec_color  = "success"
    elif aqi_mean <= 100:
        conclusion = "جودة الهواء مقبولة، غير أن بعض الملوثات قد تشكّل خطراً محدوداً على الفئات الحساسة."
        rec        = "ينصح الأفراد الحساسون بتقليل الأنشطة الخارجية المجهِدة في ساعات الذروة."
        rec_color  = "warning"
    elif aqi_mean <= 150:
        conclusion = "جودة الهواء غير صحية للفئات الحساسة وتتطلب اتخاذ إجراءات وقائية."
        rec        = "يُوصى بتقليل التعرض للهواء الخارجي وارتداء الكمامات في ساعات الذروة."
        rec_color  = "warning"
    else:
        conclusion = "جودة الهواء غير صحية وتستدعي تدخلاً عاجلاً من الجهات المختصة."
        rec        = "يُوصى بالإعلان عن تنبيه بيئي وتجنّب الأنشطة الخارجية والتنسيق مع الجهات المختصة."
        rec_color  = "danger"

    lv_color = lv_mean["color"] if lv_mean["color"] != "#FFFF00" else "#b8a000"

    return dbc.Card([
        dbc.CardHeader(dbc.Row([
            dbc.Col(html.Div("📋", style={"fontSize": "1.5rem"}), width="auto"),
            dbc.Col([
                html.H5("التقرير التحليلي الآلي", className="mb-0",
                        style={"color": "#1a3a6b", "fontWeight": "700"}),
                html.Small("تم توليده تلقائياً بواسطة نظام EnviroAI",
                           style={"color": "#666"}),
            ]),
        ], align="center"),
        style={"background": "#eef2ff", "borderBottom": "2px solid #2c5f9e", "padding": "12px 16px"}),

        dbc.CardBody([

            # 1 — Overview
            html.H6("١. نظرة عامة على البيانات", className="fw-bold mb-2",
                    style={"color": "#2c5f9e", "borderBottom": "1px solid #dee2e6", "paddingBottom": "4px"}),
            html.P([
                "يشتمل هذا التقرير على تحليل بيانات جودة الهواء المرصودة خلال الفترة من ",
                html.B(date_start), " إلى ", html.B(date_end),
                f"، بإجمالي ", html.B(f"{n_readings:,} قراءة"),
                f" موزّعة على ", html.B(f"{n_days} يوماً"), " من الرصد المستمر.",
                f" مصدر البيانات: "
                f"{'محطة AirLink المتنقلة (CSV)' if ftype == 'csv' else 'المحطة الثابتة (Excel)'}.",
            ], style={"lineHeight": "1.9", "textAlign": "justify"}),

            # 2 — Descriptive Stats
            html.H6("٢. الإحصاء الوصفي لمؤشر AQI", className="fw-bold mb-2 mt-3",
                    style={"color": "#2c5f9e", "borderBottom": "1px solid #dee2e6", "paddingBottom": "4px"}),
            dbc.Row([
                dbc.Col(dbc.Card(dbc.CardBody([
                    html.Div("المتوسط الحسابي", style={"fontSize": ".7rem", "color": "#888"}),
                    html.Div(f"{aqi_mean:.1f}", style={"fontSize": "1.4rem", "fontWeight": "700", "color": lv_color}),
                    html.Div(lv_mean["ar"],     style={"fontSize": ".68rem", "color": "#555"}),
                ], style={"textAlign": "center", "padding": "8px"}),
                style={"borderRadius": "10px", "border": f"2px solid {lv_mean['color']}55"}), width=3),

                dbc.Col(dbc.Card(dbc.CardBody([
                    html.Div("الحد الأقصى",  style={"fontSize": ".7rem", "color": "#888"}),
                    html.Div(f"{aqi_max:.0f}", style={"fontSize": "1.4rem", "fontWeight": "700", "color": "#c0392b"}),
                    html.Div(get_lv(aqi_max)["ar"], style={"fontSize": ".68rem", "color": "#555"}),
                ], style={"textAlign": "center", "padding": "8px"}),
                style={"borderRadius": "10px", "border": "2px solid #c0392b55"}), width=3),

                dbc.Col(dbc.Card(dbc.CardBody([
                    html.Div("الحد الأدنى",  style={"fontSize": ".7rem", "color": "#888"}),
                    html.Div(f"{aqi_min:.0f}", style={"fontSize": "1.4rem", "fontWeight": "700", "color": "#27ae60"}),
                    html.Div(get_lv(aqi_min)["ar"], style={"fontSize": ".68rem", "color": "#555"}),
                ], style={"textAlign": "center", "padding": "8px"}),
                style={"borderRadius": "10px", "border": "2px solid #27ae6055"}), width=3),

                dbc.Col(dbc.Card(dbc.CardBody([
                    html.Div("الانحراف المعياري", style={"fontSize": ".7rem", "color": "#888"}),
                    html.Div(f"{aqi_std:.1f}", style={"fontSize": "1.4rem", "fontWeight": "700", "color": "#8F3F97"}),
                    html.Div(f"CV = {aqi_cv:.1f}٪", style={"fontSize": ".68rem", "color": "#555"}),
                ], style={"textAlign": "center", "padding": "8px"}),
                style={"borderRadius": "10px", "border": "2px solid #8F3F9755"}), width=3),
            ], className="g-2 mb-2"),

            html.P([
                "يبلغ المتوسط الحسابي لمؤشر جودة الهواء ",
                html.B(f"{aqi_mean:.1f}"),
                " وهو يقع ضمن فئة ",
                html.B(f"«{lv_mean['ar']}»", style={"color": lv_color}),
                f". بلغ الانحراف المعياري {aqi_std:.1f} وحدة بمعامل تباين {aqi_cv:.1f}٪، مما يشير إلى ",
                "تذبذب واضح في مستويات التلوث خلال فترة الرصد." if aqi_cv > 40
                else "استقرار نسبي في مستويات التلوث خلال فترة الرصد.",
            ], style={"lineHeight": "1.9", "textAlign": "justify"}),

            # 3 — Risk levels
            html.H6("٣. توزيع مستويات الخطر", className="fw-bold mb-2 mt-3",
                    style={"color": "#2c5f9e", "borderBottom": "1px solid #dee2e6", "paddingBottom": "4px"}),
            html.Ul([
                html.Li([html.B("آمن: "),   f"{pct('آمن'):.1f}٪ من إجمالي قراءات الرصد."]),
                html.Li([html.B("متوسط: "), f"{pct('متوسط'):.1f}٪."]),
                html.Li([html.B("تحذير: "), f"{pct('تحذير'):.1f}٪ (غير صحي للفئات الحساسة)."]),
                html.Li([html.B("حرج: "),   f"{pct('حرج'):.1f}٪ (غير صحي للجمهور العام)."]),
                html.Li([html.B("طوارئ: "), f"{pct('طوارئ'):.1f}٪ (خطير جداً)."]),
            ] + pm_items, style={"lineHeight": "2.0", "paddingRight": "1.2rem"}),
            html.P([
                "بلغت نسبة القراءات ذات المستوى الخطر (حرج + طوارئ) ",
                html.B(f"{pct('حرج') + pct('طوارئ'):.1f}٪", style={"color": "#c0392b"}),
                " من إجمالي فترة الرصد.",
            ], style={"lineHeight": "1.9", "textAlign": "justify"}),

            # 4 — Trend
            html.H6("٤. تحليل الاتجاه الزمني (الانحدار الخطي البسيط — OLS)", className="fw-bold mb-2 mt-3",
                    style={"color": "#2c5f9e", "borderBottom": "1px solid #dee2e6", "paddingBottom": "4px"}),
            html.P([
                "جرى تطبيق نموذج الانحدار الخطي البسيط (Ordinary Least Squares) على المتوسطات اليومية. "
                "أظهرت النتائج أن مؤشر AQI يسير ",
                html.B(trend_dir, style={"color": "#FF7E00"}),
                " بمعدل تغيّر يبلغ ",
                html.B(f"{abs(slope):.2f} وحدة/يوم"),
                f"، ومعامل تحديد R² = {r**2:.3f}، والنموذج {trend_sig}.",
            ], className="mb-2", style={"lineHeight": "1.9", "textAlign": "justify"}),
            html.P([
                "كشف التحليل الساعي أن ذروة التلوث تبلغ أوجها عند ",
                html.B(fmt_h(peak_h)),
                "، وتنخفض إلى أدناها عند ",
                html.B(fmt_h(low_h)),
                ". يعكس هذا النمط تأثير الأنشطة البشرية والظروف الجوية على تراكم الملوثات.",
            ], style={"lineHeight": "1.9", "textAlign": "justify"}),

            # 5 — Conclusion
            html.H6("٥. الخلاصة والتوصيات", className="fw-bold mb-2 mt-3",
                    style={"color": "#2c5f9e", "borderBottom": "1px solid #dee2e6", "paddingBottom": "4px"}),
            dbc.Alert([
                html.P(conclusion, className="mb-1 fw-bold"),
                html.P([html.B("التوصية: "), rec], className="mb-0"),
            ], color=rec_color, style={"borderRadius": "10px"}),

            html.Hr(style={"margin": "12px 0 6px"}),
            html.P(
                "⚙️ الخوارزميات المستخدمة: الإحصاء الوصفي (Descriptive Statistics) — "
                "الانحدار الخطي البسيط OLS (Simple Linear Regression) — "
                "التحليل الساعي لأنماط التلوث (Hourly Pattern Analysis). "
                "جميع الحسابات مبنية على بيانات الرصد الميداني المُدخَلة.",
                style={"color": "#888", "fontSize": ".74rem",
                       "textAlign": "center", "marginBottom": "0"},
            ),

        ], style={"padding": "20px"}),
    ], style={
        "borderRadius": "14px",
        "border": "1px solid #c8d8f8",
        "marginBottom": "20px",
        "boxShadow": "0 2px 12px rgba(28,58,107,.08)",
    })


# ══════════════════════════════════════════════
#  DASHBOARD BUILDER
# ══════════════════════════════════════════════
def build_dashboard(session_id: str):
    if session_id not in _CACHE:
        return welcome_screen()

    data  = _CACHE[session_id]
    df    = data["df"]
    ftype = data["ftype"]
    daily = daily_stats(df)
    has_pm = "PM25_mean" in daily.columns or "PM10_mean" in daily.columns

    G = lambda fig: dcc.Graph(
        figure=fig,
        config={"displayModeBar": False},
        style={"height": "420px"},
    )

    # Tab: Overview
    overview = dbc.Container([
        kpi_row(df, ftype),
        dbc.Row([
            dbc.Col(G(ch_aqi_bar(daily)),    width=12, className="mb-3"),
            dbc.Col(G(ch_temp_hum(daily)),   width=6),
            dbc.Col(G(ch_pollutants(daily)), width=6),
        ]),
    ], fluid=True)

    # Tab: Risk
    rc = df["RiskLevel"].value_counts().reset_index()
    rc.columns = ["المستوى", "العدد"]
    rc["النسبة%"] = (rc["العدد"] / len(df) * 100).round(1).astype(str) + "%"
    danger = df[df["RiskLevel"].isin(["حرج", "طوارئ"])].shape[0]
    risk_tab = dbc.Container([
        dbc.Row([
            dbc.Col(G(ch_donut(df)), width=5),
            dbc.Col([
                dbc.Table.from_dataframe(
                    rc, striped=True, bordered=True, hover=True,
                    style={"fontSize": ".88rem", "textAlign": "center"},
                    className="mt-3",
                ),
                dbc.Alert(
                    f"⚠️ القراءات الخطرة (حرج+طوارئ): {danger:,}",
                    color="warning", className="mt-2 py-2",
                ),
            ], width=7),
        ]),
    ], fluid=True)

    # Assemble tabs
    tabs_list = [
        dbc.Tab(overview, label="📊 نظرة عامة", tab_id="overview"),
    ]
    if has_pm:
        tabs_list.append(dbc.Tab(
            dbc.Container(dbc.Row(dbc.Col(G(ch_pm(daily)), width=12)), fluid=True),
            label="🌫️ PM2.5/PM10", tab_id="pm",
        ))
    tabs_list += [
        dbc.Tab(
            dbc.Container(dbc.Row(dbc.Col(G(ch_hourly(df)), width=12)), fluid=True),
            label="⏰ نمط ساعي", tab_id="hourly",
        ),
        dbc.Tab(risk_tab, label="⚠️ مستوى الخطر", tab_id="risk"),
        dbc.Tab(
            dbc.Container([
                dbc.Row(dbc.Col(G(ch_trend(daily)), width=12)),
                dbc.Row(dbc.Col(dbc.Alert([
                    html.B("ملاحظة: "),
                    "التنبؤ مبني على الانحدار الخطي — للتوجيه فقط وليس تنبؤاً جوياً رسمياً.",
                ], color="info", className="mt-2"), width={"size": 10, "offset": 1})),
            ], fluid=True),
            label="📈 تنبؤ واتجاه", tab_id="forecast",
        ),
        dbc.Tab(
            dbc.Container(
                dbc.Row(
                    dbc.Col(aqi_guide_card(), width={"size": 8, "offset": 2}),
                    className="my-3",
                ),
                fluid=True,
            ),
            label="📋 دليل AQI", tab_id="guide",
        ),
    ]

    report_card = generate_report(df, daily, ftype)

    dl_btn = dbc.Row(dbc.Col(
        dbc.Button([
            "⬇️  تنزيل التقرير التحليلي (PDF)",
        ], id="btn-pdf", color="primary", size="md", className="mb-3",
           style={"fontWeight": "700", "borderRadius": "10px", "padding": "10px 24px"}),
        width="auto", className="d-flex justify-content-start",
    ))

    return dbc.Container([
        html.Hr(style={"margin": "8px 0 14px"}),
        report_card,
        dl_btn,
        dbc.Tabs(tabs_list, active_tab="overview", style={"fontWeight": "600"}),
    ], fluid=True, style={"paddingBottom": "50px"})


# ══════════════════════════════════════════════
#  APP LAYOUT
# ══════════════════════════════════════════════
app = dash.Dash(
    __name__,
    external_stylesheets=[
        dbc.themes.BOOTSTRAP,
        "https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap",
    ],
    title="EnviroAI — نظام رصد جودة الهواء",
    suppress_callback_exceptions=True,
)
server = app.server

HEADER = dbc.Navbar(
    dbc.Container(dbc.Row([
        dbc.Col(html.Div([
            html.Span("🌍 EnviroAI",
                      style={"color": "white", "fontWeight": "700", "fontSize": "1.2rem"}),
            html.Br(),
            html.Span("Air Quality & Weather Analysis System",
                      style={"color": "#a8c8ff", "fontSize": ".68rem"}),
        ]), width="auto"),
        dbc.Col(html.Div([
            html.Div("مركز رصد جودة الهواء والبيئة",
                     style={"color": "white", "fontWeight": "600",
                            "textAlign": "right", "fontSize": ".85rem"}),
            html.Div(
                "قسم مراقبة نوعية الهواء والمؤشرات البيئية",
                style={"color": "#a8c8ff", "textAlign": "right", "fontSize": ".68rem"},
            ),
        ])),
    ], align="center", justify="between"), fluid=True),
    color="#1a3a6b", dark=True,
    style={"padding": "8px 0", "direction": "rtl", "borderBottom": "3px solid #2c5f9e"},
)

UPLOAD = dbc.Container([
    dbc.Row(dbc.Col(
        dcc.Upload(
            id="upload",
            children=html.Div([
                html.Div("📁", style={"fontSize": "2.5rem", "lineHeight": "1"}),
                html.Div("اسحب وأفلت ملف البيانات هنا",
                         style={"fontWeight": "700", "fontSize": "1rem", "marginTop": "6px"}),
                html.Div("يدعم Excel (.xlsx) و CSV (تصدير AirLink)",
                         style={"color": "#555", "fontSize": ".82rem", "marginTop": "3px"}),
                html.Div(
                    "Date/Time, AQI, Temperature, Humidity, PM2.5, PM10, CO, SMOKE ...",
                    style={"color": "#999", "fontSize": ".72rem", "marginTop": "2px"},
                ),
            ], style={"textAlign": "center", "padding": "20px 10px"}),
            style={
                "border": "2px dashed #2c5f9e",
                "borderRadius": "14px",
                "background": "#f0f5ff",
                "cursor": "pointer",
            },
            multiple=False,
        ),
        width={"size": 8, "offset": 2},
    ), className="mt-3 mb-1"),
    dbc.Row(dbc.Col(html.Div(id="upload-msg"), width={"size": 8, "offset": 2})),
], fluid=True)

app.layout = html.Div([
    HEADER,
    UPLOAD,
    dcc.Store(id="session-id"),
    dcc.Download(id="download-pdf"),
    html.Div(id="page-body", children=welcome_screen()),
], style={
    "fontFamily": "Cairo,Arial,sans-serif",
    "direction": "rtl",
    "background": "#f5f7fb",
    "minHeight": "100vh",
})


# ══════════════════════════════════════════════
#  CALLBACKS
# ══════════════════════════════════════════════

# ── 1. Generate / persist session ID ────────────────────────────────
@app.callback(
    Output("session-id", "data"),
    Input("session-id", "data"),
    prevent_initial_call=False,
)
def init_session(sid):
    if sid:
        return sid
    new_id = str(uuid.uuid4())
    print(f"[EnviroAI] New session created: {new_id}")
    return new_id


# ── 2. Upload → parse → cache → render dashboard ────────────────────
@app.callback(
    Output("upload-msg",  "children"),
    Output("page-body",   "children"),
    Input("upload",       "contents"),
    State("upload",       "filename"),
    State("session-id",   "data"),
    prevent_initial_call=True,
)
def on_upload(contents, filename, session_id):
    if not contents:
        return no_update, no_update

    # Guarantee a session key even if the Store hasn't fired yet
    sid = session_id or str(uuid.uuid4())

    df, err, ftype = parse_file(contents, filename)
    if err:
        msg = dbc.Alert([html.B("❌  "), err], color="danger", className="mt-2 py-2")
        return msg, no_update

    _CACHE[sid] = {"df": df, "ftype": ftype}
    print(f"[EnviroAI] Session {sid} cached — {len(df):,} rows, ftype={ftype}")

    rows, days = len(df), df["DateStr"].nunique()
    icon        = "📄" if ftype == "csv" else "📊"
    badge_color = "success" if ftype == "csv" else "primary"
    msg = dbc.Alert([
        html.B(f"{icon} {filename}  "),
        f"— {rows:,} قراءة  |  {days} يوم  |  "
        f"نوع: {'CSV (AirLink)' if ftype == 'csv' else 'Excel'}",
    ], color=badge_color, className="mt-2 py-2")

    return msg, build_dashboard(sid)


# ── 3. PDF download ──────────────────────────────────────────────────
@app.callback(
    Output("download-pdf", "data"),
    Output("upload-msg",   "children", allow_duplicate=True),
    Input("btn-pdf",       "n_clicks"),
    State("session-id",    "data"),
    prevent_initial_call=True,
)
def download_pdf(n_clicks, session_id):
    if not n_clicks:
        return no_update, no_update

    print(f"[EnviroAI] PDF requested — session={session_id}, "
          f"in_cache={session_id in _CACHE}, "
          f"cache_keys={list(_CACHE.keys())}")

    if not session_id or session_id not in _CACHE:
        err = dbc.Alert(
            [html.B("❌ "), "انتهت الجلسة أو لم يُرفع ملف بعد. يرجى رفع الملف مجدداً."],
            color="danger", className="mt-2 py-2",
        )
        return no_update, err

    data      = _CACHE[session_id]
    df        = data["df"]
    ftype     = data["ftype"]
    daily     = daily_stats(df)
    fname     = (
        f"تقرير_جودة_الهواء_"
        f"{df['DateStr'].min()}_to_{df['DateStr'].max()}.pdf"
    )
    pdf_bytes = build_pdf_report(df, daily, ftype, fname)
    return dcc.send_bytes(pdf_bytes, fname), no_update


# ══════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
