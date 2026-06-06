"""
EnviroAI — Air Quality & Weather Analysis System
كلية العلوم / الجامعة المستنصرية
بالتعاون مع وزارة البيئة / مديرية البيئة الحضرية
قسم مراقبة نوعية الهواء والضوضاء

Dash web application — Railway deployment
"""

import base64
import io
import os

import dash
from dash import dcc, html, Input, Output, State, callback_context
import dash_bootstrap_components as dbc
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from scipy import stats
from scipy.ndimage import gaussian_filter1d

# ─────────────────────────────────────────────
#  AQI OFFICIAL SCALE
# ─────────────────────────────────────────────
AQI_LEVELS = [
    {"range": (0,   50),  "color": "#00E400", "label_ar": "جيد",                     "label_en": "Good"},
    {"range": (51,  100), "color": "#FFFF00", "label_ar": "معتدل",                   "label_en": "Moderate"},
    {"range": (101, 150), "color": "#FF7E00", "label_ar": "غير صحي للحساسين",        "label_en": "Unhealthy (Sensitive)"},
    {"range": (151, 200), "color": "#FF0000", "label_ar": "غير صحي",                 "label_en": "Unhealthy"},
    {"range": (201, 300), "color": "#8F3F97", "label_ar": "غير صحي جداً",            "label_en": "Very Unhealthy"},
    {"range": (301, 500), "color": "#7E0023", "label_ar": "خطير",                    "label_en": "Hazardous"},
]

RISK_COLORS = {
    "آمن":    "#00E400",
    "متوسط":  "#FFFF00",
    "تحذير":  "#FF7E00",
    "حرج":    "#FF0000",
    "طوارئ":  "#8F3F97",
}

def get_aqi_level(v):
    for lv in AQI_LEVELS:
        if lv["range"][0] <= v <= lv["range"][1]:
            return lv
    return AQI_LEVELS[-1]

def aqi_color(v):
    return get_aqi_level(v)["color"]

# ─────────────────────────────────────────────
#  DATA LOADING
# ─────────────────────────────────────────────
REQUIRED = {"Date","Time","Temperature","Humidity","AQI","CO","SMOKE","RiskIndex","RiskLevel"}

def parse_excel(contents: str, filename: str) -> tuple[pd.DataFrame | None, str]:
    """Decode base64 upload and return (dataframe, error_msg)."""
    try:
        content_type, content_string = contents.split(",")
        decoded = base64.b64decode(content_string)
        sheets = pd.read_excel(io.BytesIO(decoded), sheet_name=None)
    except Exception as e:
        return None, f"تعذّر قراءة الملف: {e}"

    df = None
    for name, sdf in sheets.items():
        if REQUIRED.issubset(set(sdf.columns)):
            df = sdf.copy()
            break

    if df is None:
        return None, f"لم يُعثر على الأعمدة المطلوبة في: {filename}"

    df["Date"]    = pd.to_datetime(df["Date"], errors="coerce")
    df            = df.dropna(subset=["Date"])
    df["DateStr"] = df["Date"].dt.strftime("%Y-%m-%d")
    df["DayLabel"]= df["Date"].dt.strftime("%-d %b")
    df["Hour"]    = pd.to_datetime(df["Time"].astype(str), errors="coerce").dt.hour
    return df, ""

def daily_stats(df: pd.DataFrame) -> pd.DataFrame:
    return df.groupby("DateStr").agg(
        DayLabel  =("DayLabel",    "first"),
        AQI_mean  =("AQI",         "mean"),
        AQI_max   =("AQI",         "max"),
        Temp_mean =("Temperature", "mean"),
        Temp_max  =("Temperature", "max"),
        Hum_mean  =("Humidity",    "mean"),
        CO_mean   =("CO",          "mean"),
        CO_max    =("CO",          "max"),
        SMOKE_mean=("SMOKE",       "mean"),
        Risk_mean =("RiskIndex",   "mean"),
    ).round(2).reset_index()

# ─────────────────────────────────────────────
#  CHART BUILDERS
# ─────────────────────────────────────────────
def build_aqi_bar(daily: pd.DataFrame) -> go.Figure:
    colors = [aqi_color(v) for v in daily["AQI_mean"]]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=daily["DayLabel"], y=daily["AQI_mean"],
        marker_color=colors,
        marker_line_color="rgba(0,0,0,0.15)", marker_line_width=1,
        text=daily["AQI_mean"].round(0).astype(int),
        textposition="outside", name="AQI",
        hovertemplate="<b>%{x}</b><br>AQI: %{y:.0f}<extra></extra>",
    ))
    for thresh, lbl, col in [
        (50, "جيد / Good",             "#00E400"),
        (100,"معتدل / Moderate",        "#FFFF00"),
        (150,"غير صحي (حساسين)",        "#FF7E00"),
        (200,"غير صحي / Unhealthy",     "#FF0000"),
        (300,"غير صحي جداً",            "#8F3F97"),
    ]:
        fig.add_hline(y=thresh, line_dash="dot", line_color=col,
                      line_width=1.2, opacity=0.7,
                      annotation_text=lbl, annotation_position="right",
                      annotation_font_size=9, annotation_font_color=col)
    fig.update_layout(
        **_layout("متوسط مؤشر AQI اليومي  |  Daily AQI Average"),
        yaxis_title="AQI", xaxis_title="",
        yaxis_range=[0, max(daily["AQI_mean"].max() * 1.22, 320)],
    )
    return fig

def build_temp_humidity(daily: pd.DataFrame) -> go.Figure:
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Scatter(
        x=daily["DayLabel"], y=daily["Temp_mean"],
        mode="lines+markers", name="درجة الحرارة (°م)",
        line=dict(color="#FF7E00", width=2.5),
        marker=dict(size=7),
        hovertemplate="%{y:.1f} °C<extra>حرارة</extra>",
    ), secondary_y=False)
    fig.add_trace(go.Bar(
        x=daily["DayLabel"], y=daily["Hum_mean"],
        name="الرطوبة (%)", opacity=0.35,
        marker_color="#3498db",
        hovertemplate="%{y:.1f}%<extra>رطوبة</extra>",
    ), secondary_y=True)
    fig.update_yaxes(title_text="°C", secondary_y=False, title_font_color="#FF7E00")
    fig.update_yaxes(title_text="%", secondary_y=True, title_font_color="#3498db")
    fig.update_layout(**_layout("درجة الحرارة والرطوبة  |  Temp & Humidity"))
    return fig

def build_co_smoke(daily: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=daily["DayLabel"], y=daily["CO_mean"],
        name="CO (ppm)", marker_color="#8F3F97",
        marker_line_color="rgba(0,0,0,0.1)", marker_line_width=0.8,
        hovertemplate="CO: %{y:.0f} ppm<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        x=daily["DayLabel"], y=daily["SMOKE_mean"],
        name="SMOKE", marker_color="#c0392b",
        marker_line_color="rgba(0,0,0,0.1)", marker_line_width=0.8,
        hovertemplate="SMOKE: %{y:.0f}<extra></extra>",
    ))
    fig.update_layout(**_layout("CO والدخان اليومي  |  CO & Smoke"), barmode="group")
    return fig

def build_risk_donut(df: pd.DataFrame) -> go.Figure:
    counts = df["RiskLevel"].value_counts()
    order  = ["آمن","متوسط","تحذير","حرج","طوارئ"]
    vals   = [counts.get(k, 0) for k in order]
    colors = [RISK_COLORS[k] for k in order]
    fig = go.Figure(go.Pie(
        labels=order, values=vals, hole=0.55,
        marker=dict(colors=colors, line=dict(color="#ffffff", width=2)),
        textfont_size=13,
        hovertemplate="<b>%{label}</b><br>%{value:,} قراءة<br>%{percent}<extra></extra>",
    ))
    total = sum(vals)
    fig.add_annotation(text=f"<b>{total:,}</b><br>قراءة",
                       x=0.5, y=0.5, showarrow=False,
                       font=dict(size=14, color="#333"))
    fig.update_layout(**_layout("توزيع مستويات الخطر  |  Risk Distribution"))
    return fig

def build_hourly_pattern(df: pd.DataFrame) -> go.Figure:
    hourly = df.dropna(subset=["Hour"]).groupby("Hour")["AQI"].mean().reindex(range(24), fill_value=None).reset_index()
    hourly.columns = ["Hour","AQI"]
    colors = [aqi_color(v) if pd.notna(v) else "#cccccc" for v in hourly["AQI"]]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=hourly["Hour"], y=hourly["AQI"].fillna(0),
        marker_color=colors,
        marker_line_color="rgba(0,0,0,0.1)", marker_line_width=0.5,
        name="AQI ساعي",
        hovertemplate="الساعة %{x}:00<br>AQI: %{y:.0f}<extra></extra>",
    ))
    fig.add_hline(y=200, line_dash="dot", line_color="#FF0000",
                  line_width=1, annotation_text="200 — خطر",
                  annotation_font_color="#FF0000", annotation_font_size=9)
    fig.update_layout(
        **_layout("متوسط AQI لكل ساعة  |  Hourly AQI Pattern"),
        xaxis=dict(tickmode="array",
                   tickvals=[0,3,6,9,12,15,18,21],
                   ticktext=["12AM","3AM","6AM","9AM","12PM","3PM","6PM","9PM"]),
    )
    return fig

def build_trend_forecast(daily: pd.DataFrame) -> go.Figure:
    aqi = daily["AQI_mean"].values
    x   = np.arange(len(aqi))
    slope, intercept, r_val, p_val, se = stats.linregress(x, aqi)
    trend = slope * x + intercept

    x_fut = np.arange(len(aqi), len(aqi) + 7)
    y_fut = slope * x_fut + intercept
    ci    = 1.96 * se * np.sqrt(1 + 1/len(x) + (x_fut - x.mean())**2 / ((x - x.mean())**2).sum())

    smooth = gaussian_filter1d(aqi, sigma=1.2)
    all_labels = list(daily["DayLabel"]) + [f"يوم+{j+1}" for j in range(7)]

    fig = go.Figure()
    # AQI bands background
    for lo, hi, col in [(0,50,"#00E400"),(50,100,"#FFFF00"),(100,150,"#FF7E00"),
                         (150,200,"#FF0000"),(200,300,"#8F3F97"),(300,500,"#7E0023")]:
        fig.add_hrect(y0=lo, y1=hi, fillcolor=col, opacity=0.07, line_width=0)

    # historical
    fig.add_trace(go.Scatter(
        x=list(daily["DayLabel"]), y=smooth,
        mode="lines", name="AQI (منعّم)", line=dict(color="#3498db", width=2.5),
    ))
    fig.add_trace(go.Scatter(
        x=list(daily["DayLabel"]), y=aqi,
        mode="markers", name="AQI فعلي",
        marker=dict(color=[aqi_color(v) for v in aqi], size=10,
                    line=dict(color="white", width=1.5)),
    ))
    # trend line extended
    all_x = list(daily["DayLabel"]) + [f"يوم+{j+1}" for j in range(7)]
    fig.add_trace(go.Scatter(
        x=all_x, y=np.append(trend, y_fut),
        mode="lines", name=f"اتجاه (+{slope:.1f}/يوم)",
        line=dict(color="#FF7E00", width=2, dash="dash"),
    ))
    # forecast scatter
    fig.add_trace(go.Scatter(
        x=[f"يوم+{j+1}" for j in range(7)], y=np.clip(y_fut, 0, 500),
        mode="markers+text",
        marker=dict(color=[aqi_color(max(0,v)) for v in y_fut], size=13,
                    symbol="diamond", line=dict(color="white", width=1.5)),
        text=[f"{v:.0f}" for v in y_fut],
        textposition="top center", textfont=dict(size=9),
        name="تنبؤ 7 أيام",
    ))
    # CI band
    fig.add_trace(go.Scatter(
        x=[f"يوم+{j+1}" for j in range(7)] + [f"يوم+{j+1}" for j in range(6,-1,-1)],
        y=list(np.clip(y_fut + ci, 0, 600)) + list(np.clip(y_fut - ci, 0, 600)),
        fill="toself", fillcolor="rgba(143,63,151,0.15)",
        line=dict(color="rgba(0,0,0,0)"), name="فترة ثقة 95%",
    ))
    # divider
    fig.add_vline(x=len(aqi) - 0.5, line_dash="dot",
                  line_color="#888", line_width=1)
    fig.add_annotation(x=len(aqi) - 0.7, y=max(aqi) * 1.05,
                       text="◀ تاريخي", showarrow=False,
                       font=dict(color="#888", size=9))
    fig.add_annotation(x=len(aqi) + 0.2, y=max(aqi) * 1.05,
                       text="تنبؤ ▶", showarrow=False,
                       font=dict(color="#8F3F97", size=9))

    r2 = r_val**2
    fig.update_layout(
        **_layout(f"تحليل الاتجاه والتنبؤ  |  Trend & Forecast   R²={r2:.3f}  p={p_val:.3f}"),
        xaxis_tickangle=-30,
    )
    return fig

def build_scatter_temp_aqi(df: pd.DataFrame) -> go.Figure:
    sample = df.sample(min(2000, len(df)), random_state=42)
    fig = px.scatter(
        sample, x="Temperature", y="AQI",
        color="AQI", color_continuous_scale=[
            [0.00, "#00E400"],[0.10, "#00E400"],
            [0.10, "#FFFF00"],[0.20, "#FFFF00"],
            [0.20, "#FF7E00"],[0.30, "#FF7E00"],
            [0.30, "#FF0000"],[0.40, "#FF0000"],
            [0.40, "#8F3F97"],[0.60, "#8F3F97"],
            [0.60, "#7E0023"],[1.00, "#7E0023"],
        ],
        range_color=[0, 500],
        opacity=0.6, size_max=6,
        labels={"Temperature":"درجة الحرارة (°م)", "AQI":"AQI"},
        hover_data={"RiskLevel": True, "Humidity": True},
    )
    fig.update_layout(**_layout("الحرارة مقابل AQI  |  Temperature vs AQI"))
    return fig

def _layout(title: str) -> dict:
    return dict(
        title=dict(text=title, font=dict(size=13, color="#1a1a2e"), x=0.5, xanchor="center"),
        paper_bgcolor="white", plot_bgcolor="#f8f9fa",
        font=dict(family="Cairo, Arial", color="#333"),
        margin=dict(t=50, b=40, l=50, r=30),
        legend=dict(bgcolor="rgba(255,255,255,0.8)", bordercolor="#ddd", borderwidth=1),
        hovermode="x unified",
        xaxis=dict(gridcolor="#e8e8e8", showgrid=True),
        yaxis=dict(gridcolor="#e8e8e8", showgrid=True),
    )

# ─────────────────────────────────────────────
#  AQI SCALE CARD  (رموز الوجوه)
# ─────────────────────────────────────────────
FACE_EMOJIS = ["😊","🙂","😐","😷","😷","😷"]
SCALE_DATA = [
    ("0 – 50",   "جيد",                            "Good",                "جودة الهواء مرضية وتشكل خطرًا قليلاً أو منعدماً.",       "#00E400","#000"),
    ("51 – 100", "معتدل",                          "Moderate",            "يجب على الأفراد الحساسين تجنب الأنشطة الخارجية المكثفة.","#FFFF00","#000"),
    ("101 – 150","غير صحي للحساسين",              "Unhealthy (Sensitive)","الجمهور والحساسون معرضون لخطر مشاكل تنفسية.",            "#FF7E00","#000"),
    ("151 – 200","غير صحي",                        "Unhealthy",           "يزداد احتمال حدوث آثار ضارة على الجمهور.",               "#FF0000","#fff"),
    ("201 – 300","غير صحي جداً",                   "Very Unhealthy",      "سيتأثر عموم الجمهور. يُنصح بالبقاء في المنازل.",         "#8F3F97","#fff"),
    ("301 – 500","خطير",                            "Hazardous",           "خطير جداً. يجب على الجميع تجنب الرياضة والبقاء في الداخل.","#7E0023","#fff"),
]

def aqi_scale_card():
    rows = []
    for i, (rng, ar, en, desc, bg, tc) in enumerate(SCALE_DATA):
        rows.append(
            dbc.Row([
                dbc.Col(
                    html.Div(FACE_EMOJIS[i], style={
                        "backgroundColor": bg,
                        "borderRadius": "10px",
                        "fontSize": "2.2rem",
                        "textAlign": "center",
                        "padding": "6px 4px",
                        "lineHeight": "1.1",
                        "minWidth": "62px",
                    }),
                    width="auto",
                ),
                dbc.Col([
                    html.Div([
                        html.Span(rng, style={"fontWeight":"700","fontSize":"1rem","marginLeft":"8px"}),
                        html.Span(ar,  style={"fontWeight":"700","fontSize":"1rem","color":"#222"}),
                    ]),
                    html.Div(desc, style={"fontSize":"0.78rem","color":"#555","marginTop":"2px"}),
                ], style={"paddingRight":"4px"}),
            ], align="center",
               style={
                   "borderBottom": "1px solid #eee",
                   "padding": "8px 10px",
                   "margin": "0",
               })
        )
    return dbc.Card([
        dbc.CardHeader(
            html.H6("دليل مؤشر جودة الهواء  |  AQI Reference Guide",
                    className="mb-0 text-center fw-bold"),
            style={"backgroundColor":"#f0f4ff","borderBottom":"2px solid #2c5f9e"}
        ),
        dbc.CardBody(rows, style={"padding":"0"}),
    ], style={"borderRadius":"12px","overflow":"hidden","border":"1px solid #dde"})

# ─────────────────────────────────────────────
#  STAT CARDS
# ─────────────────────────────────────────────
def stat_cards(df):
    aqi_mean = df["AQI"].mean()
    lv = get_aqi_level(aqi_mean)
    cards = [
        ("📊", f"{len(df):,}", "إجمالي القراءات", "#2c5f9e"),
        ("🌡️", f"{df['Temperature'].mean():.1f}°م", "متوسط الحرارة", "#e67e22"),
        ("💨", f"{aqi_mean:.1f}", f"متوسط AQI — {lv['label_ar']}", lv["color"] if lv["color"] != "#FFFF00" else "#b8a000"),
        ("⚠️", f"{df['AQI'].max():.0f}", "أعلى AQI مُسجّل", "#c0392b"),
        ("🏭", f"{df['CO'].mean():.0f}", "متوسط CO (ppm)", "#8F3F97"),
        ("📅", f"{df['DateStr'].nunique()}", "أيام الرصد", "#27ae60"),
    ]
    return dbc.Row([
        dbc.Col(dbc.Card([
            dbc.CardBody([
                html.Div(icon, style={"fontSize":"1.6rem","textAlign":"center"}),
                html.H4(val, className="text-center fw-bold mb-0",
                        style={"color": color, "fontSize":"1.3rem"}),
                html.P(label, className="text-center text-muted mb-0",
                       style={"fontSize":"0.72rem"}),
            ], style={"padding":"10px 6px"}),
        ], style={"borderRadius":"10px","border":f"1px solid {color}22",
                  "backgroundColor":"#fff","height":"100%"}),
        width=2) for icon, val, label, color in cards
    ], className="g-2 mb-3")

# ─────────────────────────────────────────────
#  DASH APP
# ─────────────────────────────────────────────
app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.BOOTSTRAP,
                           "https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap"],
    title="EnviroAI — الجامعة المستنصرية",
    suppress_callback_exceptions=True,
)
server = app.server   # ← gunicorn entry point

# ── HEADER ──────────────────────────────────────────────────────────────────
header = dbc.Navbar(
    dbc.Container([
        html.Div([
            html.H5("🌍 EnviroAI", className="mb-0 fw-bold",
                    style={"color":"white","fontSize":"1.25rem"}),
            html.Small("Air Quality & Weather Analysis System",
                       style={"color":"#a8c8ff","display":"block","fontSize":"0.72rem"}),
        ]),
        html.Div([
            html.Div("كلية العلوم / الجامعة المستنصرية",
                     style={"color":"white","fontWeight":"600","textAlign":"right","fontSize":"0.85rem"}),
            html.Div("بالتعاون مع وزارة البيئة / مديرية البيئة الحضرية",
                     style={"color":"#a8c8ff","textAlign":"right","fontSize":"0.72rem"}),
            html.Div("قسم مراقبة نوعية الهواء والضوضاء",
                     style={"color":"#a8c8ff","textAlign":"right","fontSize":"0.70rem"}),
        ]),
    ], fluid=True, style={"display":"flex","justifyContent":"space-between","alignItems":"center"}),
    color="#1a3a6b", dark=True, style={"padding":"10px 0","direction":"rtl"},
)

# ── UPLOAD AREA ─────────────────────────────────────────────────────────────
upload_section = dbc.Container([
    dbc.Row(dbc.Col(
        dcc.Upload(
            id="upload-data",
            children=html.Div([
                html.Div("📁", style={"fontSize":"2.5rem"}),
                html.Div("اسحب وأفلت ملف Excel هنا", style={"fontWeight":"600","fontSize":"1rem"}),
                html.Div("أو انقر لاختيار الملف (.xlsx)", style={"color":"#888","fontSize":"0.82rem"}),
                html.Div("يدعم أي ملف يحتوي الأعمدة: Date, Time, Temperature, Humidity, AQI, CO, SMOKE, RiskIndex, RiskLevel",
                         style={"color":"#aaa","fontSize":"0.72rem","marginTop":"4px"}),
            ], style={"textAlign":"center","padding":"16px"}),
            style={
                "border":"2px dashed #2c5f9e","borderRadius":"12px",
                "backgroundColor":"#f0f5ff","cursor":"pointer",
                "transition":"border-color 0.2s",
            },
            multiple=False,
        ),
        width={"size":8,"offset":2},
    ), className="my-3"),
    dbc.Row(dbc.Col(html.Div(id="upload-status"), width={"size":8,"offset":2})),
], fluid=True)

# ── MAIN LAYOUT ─────────────────────────────────────────────────────────────
app.layout = html.Div([
    header,
    upload_section,
    html.Div(id="dashboard-content"),
    dcc.Store(id="stored-data"),
], style={"fontFamily":"Cairo, Arial, sans-serif","direction":"rtl","backgroundColor":"#f5f7fb"})


# ─────────────────────────────────────────────
#  CALLBACKS
# ─────────────────────────────────────────────
@app.callback(
    Output("stored-data",    "data"),
    Output("upload-status",  "children"),
    Input("upload-data",     "contents"),
    State("upload-data",     "filename"),
    prevent_initial_call=True,
)
def store_upload(contents, filename):
    if contents is None:
        return None, ""
    df, err = parse_excel(contents, filename)
    if err:
        return None, dbc.Alert(f"❌ {err}", color="danger", className="mt-2")
    status = dbc.Alert(
        [html.B(f"✅ تم تحميل: {filename}  "),
         f"— {len(df):,} قراءة  |  {df['DateStr'].nunique()} يوم رصد"],
        color="success", className="mt-2",
    )
    return df.to_json(date_format="iso", orient="split"), status


@app.callback(
    Output("dashboard-content", "children"),
    Input("stored-data", "data"),
    prevent_initial_call=True,
)
def render_dashboard(json_data):
    if not json_data:
        return ""
    df = pd.read_json(json_data, orient="split")
    df["Date"]    = pd.to_datetime(df["Date"], errors="coerce")
    df["DateStr"] = df["Date"].dt.strftime("%Y-%m-%d")
    df["DayLabel"]= df["Date"].dt.strftime("%-d %b")
    df["Hour"]    = pd.to_datetime(df["Time"].astype(str), errors="coerce").dt.hour
    daily = daily_stats(df)

    tabs = dbc.Tabs([
        dbc.Tab(label="📊 نظرة عامة",   tab_id="overview"),
        dbc.Tab(label="💨 جودة الهواء", tab_id="aqi"),
        dbc.Tab(label="⚠️ مستوى الخطر", tab_id="risk"),
        dbc.Tab(label="📈 تنبؤ واتجاه", tab_id="forecast"),
        dbc.Tab(label="📋 دليل AQI",    tab_id="guide"),
    ], id="main-tabs", active_tab="overview",
       style={"marginBottom":"16px"})

    # ── Overview tab ──
    tab_overview = dbc.Container([
        stat_cards(df),
        dbc.Row([
            dbc.Col(dcc.Graph(figure=build_aqi_bar(daily),    config={"displayModeBar":False}), width=12, className="mb-3"),
            dbc.Col(dcc.Graph(figure=build_temp_humidity(daily), config={"displayModeBar":False}), width=6),
            dbc.Col(dcc.Graph(figure=build_co_smoke(daily),   config={"displayModeBar":False}), width=6),
        ]),
    ], fluid=True)

    # ── AQI tab ──
    tab_aqi = dbc.Container([
        dbc.Row([
            dbc.Col(dcc.Graph(figure=build_hourly_pattern(df), config={"displayModeBar":False}), width=6),
            dbc.Col(dcc.Graph(figure=build_scatter_temp_aqi(df), config={"displayModeBar":False}), width=6),
        ]),
    ], fluid=True)

    # ── Risk tab ──
    risk_counts = df["RiskLevel"].value_counts().reset_index()
    risk_counts.columns = ["المستوى","العدد"]
    risk_counts["النسبة"] = (risk_counts["العدد"] / len(df) * 100).round(1).astype(str) + "%"
    risk_counts["اللون"]   = risk_counts["المستوى"].map(RISK_COLORS)

    tab_risk = dbc.Container([
        dbc.Row([
            dbc.Col(dcc.Graph(figure=build_risk_donut(df), config={"displayModeBar":False}), width=5),
            dbc.Col([
                dbc.Table.from_dataframe(
                    risk_counts[["المستوى","العدد","النسبة"]],
                    striped=True, bordered=True, hover=True,
                    style={"fontSize":"0.88rem","textAlign":"center"},
                    className="mt-3",
                ),
                html.P("* إجمالي القراءات الخطرة (حرج + طوارئ): "
                       f"{df[df['RiskLevel'].isin(['حرج','طوارئ'])].shape[0]:,}",
                       className="text-muted mt-2", style={"fontSize":"0.8rem"}),
            ], width=7),
        ]),
    ], fluid=True)

    # ── Forecast tab ──
    tab_forecast = dbc.Container([
        dbc.Row([
            dbc.Col(dcc.Graph(figure=build_trend_forecast(daily), config={"displayModeBar":False}), width=12),
        ]),
        dbc.Row(dbc.Col(dbc.Alert([
            html.B("ملاحظة: "),
            "التنبؤ مبني على الانحدار الخطي للبيانات المتاحة. "
            "الأرقام تقديرية للتوجيه فقط وليست تنبؤاً جوياً رسمياً.",
        ], color="info", className="mt-2"), width={"size":10,"offset":1})),
    ], fluid=True)

    # ── Guide tab ──
    tab_guide = dbc.Container([
        dbc.Row(dbc.Col(aqi_scale_card(), width={"size":8,"offset":2}), className="my-3"),
    ], fluid=True)

    content_map = {
        "overview": tab_overview,
        "aqi":      tab_aqi,
        "risk":     tab_risk,
        "forecast": tab_forecast,
        "guide":    tab_guide,
    }

    return dbc.Container([
        html.Hr(style={"margin":"8px 0"}),
        tabs,
        html.Div(id="tab-content-area"),
        dcc.Store(id="df-store-local", data=json_data),
    ], fluid=True, style={"paddingBottom":"40px"})


@app.callback(
    Output("tab-content-area", "children"),
    Input("main-tabs",      "active_tab"),
    State("df-store-local", "data"),
    prevent_initial_call=True,
)
def switch_tab(active_tab, json_data):
    if not json_data:
        return ""
    df = pd.read_json(json_data, orient="split")
    df["Date"]    = pd.to_datetime(df["Date"], errors="coerce")
    df["DateStr"] = df["Date"].dt.strftime("%Y-%m-%d")
    df["DayLabel"]= df["Date"].dt.strftime("%-d %b")
    df["Hour"]    = pd.to_datetime(df["Time"].astype(str), errors="coerce").dt.hour
    daily = daily_stats(df)

    if active_tab == "overview":
        return dbc.Container([
            stat_cards(df),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=build_aqi_bar(daily),    config={"displayModeBar":False}), width=12, className="mb-3"),
                dbc.Col(dcc.Graph(figure=build_temp_humidity(daily), config={"displayModeBar":False}), width=6),
                dbc.Col(dcc.Graph(figure=build_co_smoke(daily),   config={"displayModeBar":False}), width=6),
            ]),
        ], fluid=True)

    if active_tab == "aqi":
        return dbc.Container([
            dbc.Row([
                dbc.Col(dcc.Graph(figure=build_hourly_pattern(df),   config={"displayModeBar":False}), width=6),
                dbc.Col(dcc.Graph(figure=build_scatter_temp_aqi(df), config={"displayModeBar":False}), width=6),
            ]),
        ], fluid=True)

    if active_tab == "risk":
        risk_counts = df["RiskLevel"].value_counts().reset_index()
        risk_counts.columns = ["المستوى","العدد"]
        risk_counts["النسبة"] = (risk_counts["العدد"] / len(df) * 100).round(1).astype(str) + "%"
        return dbc.Container([
            dbc.Row([
                dbc.Col(dcc.Graph(figure=build_risk_donut(df), config={"displayModeBar":False}), width=5),
                dbc.Col([
                    dbc.Table.from_dataframe(
                        risk_counts[["المستوى","العدد","النسبة"]],
                        striped=True, bordered=True, hover=True,
                        style={"fontSize":"0.88rem","textAlign":"center"}, className="mt-3",
                    ),
                    html.P(f"القراءات الخطرة (حرج + طوارئ): {df[df['RiskLevel'].isin(['حرج','طوارئ'])].shape[0]:,}",
                           className="text-muted mt-2", style={"fontSize":"0.8rem"}),
                ], width=7),
            ]),
        ], fluid=True)

    if active_tab == "forecast":
        return dbc.Container([
            dbc.Row(dbc.Col(dcc.Graph(figure=build_trend_forecast(daily), config={"displayModeBar":False}), width=12)),
            dbc.Row(dbc.Col(dbc.Alert([html.B("ملاحظة: "),
                "التنبؤ مبني على الانحدار الخطي — للتوجيه فقط."], color="info"), width={"size":10,"offset":1})),
        ], fluid=True)

    if active_tab == "guide":
        return dbc.Container([
            dbc.Row(dbc.Col(aqi_scale_card(), width={"size":8,"offset":2}), className="my-3"),
        ], fluid=True)

    return ""


# ─────────────────────────────────────────────
#  RUN
# ─────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
