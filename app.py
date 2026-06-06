"""
EnviroAI — Air Quality & Weather Analysis System
كلية العلوم / الجامعة المستنصرية
بالتعاون مع وزارة البيئة / مديرية البيئة الحضرية
قسم مراقبة نوعية الهواء والضوضاء
Dash web app — supports Excel (.xlsx) AND CSV (AirLink format)
"""

import base64, io, os
import chardet
import dash
from dash import dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from scipy import stats
from scipy.ndimage import gaussian_filter1d
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
#  AQI OFFICIAL SCALE
# ─────────────────────────────────────────────
AQI_LEVELS = [
    {"range":(0,50),   "color":"#00E400","label_ar":"جيد",               "label_en":"Good"},
    {"range":(51,100), "color":"#FFFF00","label_ar":"معتدل",             "label_en":"Moderate"},
    {"range":(101,150),"color":"#FF7E00","label_ar":"غير صحي للحساسين",  "label_en":"Unhealthy (Sensitive)"},
    {"range":(151,200),"color":"#FF0000","label_ar":"غير صحي",           "label_en":"Unhealthy"},
    {"range":(201,300),"color":"#8F3F97","label_ar":"غير صحي جداً",      "label_en":"Very Unhealthy"},
    {"range":(301,500),"color":"#7E0023","label_ar":"خطير",              "label_en":"Hazardous"},
]
RISK_COLORS = {"آمن":"#00E400","متوسط":"#FFFF00","تحذير":"#FF7E00","حرج":"#FF0000","طوارئ":"#8F3F97"}
FACE_EMOJIS = ["😊","🙂","😐","😷","😷","😷"]
SCALE_DATA  = [
    ("0 – 50",   "جيد",               "Good",                  "جودة الهواء مرضية وتشكل خطرًا قليلاً. يوصى بتهوية منزلك.",              "#00E400","#000"),
    ("51 – 100", "معتدل",             "Moderate",              "يجب على الأفراد الحساسين تجنب الأنشطة الخارجية المكثفة.",               "#FFFF00","#000"),
    ("101 – 150","غير صحي للحساسين","Unhealthy (Sensitive)",  "الجمهور والحساسون معرضون لخطر مشاكل تنفسية.",                          "#FF7E00","#000"),
    ("151 – 200","غير صحي",          "Unhealthy",             "يزداد احتمال حدوث آثار ضارة على الجمهور.",                             "#FF0000","#fff"),
    ("201 – 300","غير صحي جداً",     "Very Unhealthy",        "سيتأثر عموم الجمهور. يُنصح بالبقاء في المنازل.",                       "#8F3F97","#fff"),
    ("301 – 500","خطير",             "Hazardous",             "خطير جداً. يجب على الجميع تجنب الرياضة والبقاء في الداخل.",            "#7E0023","#fff"),
]

def get_aqi_level(v):
    for lv in AQI_LEVELS:
        if lv["range"][0] <= v <= lv["range"][1]: return lv
    return AQI_LEVELS[-1]

def aqi_color(v): return get_aqi_level(max(0, float(v)))["color"]

def aqi_badge(v):
    lv = get_aqi_level(max(0, float(v)))
    tc = "#000" if lv["color"] in ("#00E400","#FFFF00","#FF7E00") else "#fff"
    return html.Span(lv["label_ar"], style={
        "backgroundColor": lv["color"], "color": tc,
        "padding":"2px 10px","borderRadius":"12px",
        "fontSize":"0.75rem","fontWeight":"600"})

# ─────────────────────────────────────────────
#  FILE PARSING — Excel AND CSV
# ─────────────────────────────────────────────
REQUIRED_EXCEL = {"Date","Time","Temperature","Humidity","AQI","CO","SMOKE","RiskIndex","RiskLevel"}

def _enrich(df: pd.DataFrame) -> pd.DataFrame:
    """Add computed columns common to both file types."""
    df["Date"]     = pd.to_datetime(df["Date"], errors="coerce")
    df             = df.dropna(subset=["Date"])
    df["DateStr"]  = df["Date"].dt.strftime("%Y-%m-%d")
    df["DayLabel"] = df["Date"].dt.strftime("%-d %b")
    df["Hour"]     = pd.to_datetime(df["Time"].astype(str), errors="coerce").dt.hour \
                     if "Time" in df.columns else df["Date"].dt.hour
    # AQI level columns
    df["AQI_Color"]  = df["AQI"].apply(lambda v: aqi_color(v) if pd.notna(v) else "#ccc")
    df["AQI_Label"]  = df["AQI"].apply(lambda v: get_aqi_level(max(0,float(v)))["label_ar"] if pd.notna(v) else "")
    # RiskLevel fallback from AQI
    if "RiskLevel" not in df.columns:
        def _risk(v):
            if v <= 50:   return "آمن"
            if v <= 100:  return "متوسط"
            if v <= 150:  return "تحذير"
            if v <= 200:  return "حرج"
            return "طوارئ"
        df["RiskLevel"] = df["AQI"].apply(lambda v: _risk(float(v)) if pd.notna(v) else "متوسط")
    return df

def parse_excel(contents: str, filename: str):
    try:
        _, content_string = contents.split(",")
        decoded = base64.b64decode(content_string)
        sheets  = pd.read_excel(io.BytesIO(decoded), sheet_name=None)
    except Exception as e:
        return None, f"تعذّر قراءة Excel: {e}", "excel"
    df = None
    for _, sdf in sheets.items():
        if REQUIRED_EXCEL.issubset(set(sdf.columns)):
            df = sdf.copy(); break
    if df is None:
        return None, "لم يُعثر على الأعمدة المطلوبة في ملف Excel.", "excel"
    return _enrich(df), "", "excel"

def parse_csv(contents: str, filename: str):
    """Parse AirLink-style CSV with 5 header rows."""
    try:
        _, content_string = contents.split(",")
        raw  = base64.b64decode(content_string)
        enc  = chardet.detect(raw)["encoding"] or "latin-1"
        text = raw.decode(enc, errors="replace")
        buf  = io.StringIO(text)

        # Read without header first to detect data start row
        tmp = pd.read_csv(io.StringIO(text), header=None, on_bad_lines="skip")

        # Find the row that looks like column headers (contains 'Date')
        header_row = 0
        for i, row in tmp.iterrows():
            vals = [str(v).strip() for v in row if pd.notna(v)]
            if any("date" in v.lower() or "time" in v.lower() for v in vals):
                header_row = i
                break

        df = pd.read_csv(io.StringIO(text), header=None, skiprows=header_row+1,
                         on_bad_lines="skip")
        # Use header row as column names
        col_row = tmp.iloc[header_row].tolist()
        df.columns = [str(c).strip() if pd.notna(c) else f"col_{i}"
                      for i, c in enumerate(col_row[:len(df.columns)])]

    except Exception as e:
        return None, f"تعذّر قراءة CSV: {e}", "csv"

    # ── Column mapping ──────────────────────────────────────────────────
    col_map = {}
    for col in df.columns:
        cl = col.lower()
        if "date" in cl or "time" in cl:                     col_map[col] = "Date"
        elif "aqi" in cl:                                    col_map[col] = "AQI"
        elif "pm 2.5" in cl or "pm2.5" in cl or "2.5" in cl:col_map[col] = "PM25"
        elif "pm 10" in cl  or "pm10"  in cl or "pm 1" in cl and "0" in cl: col_map[col] = "PM10"
        elif "pm 1" in cl   or "pm1"   in cl:               col_map[col] = "PM1"
        elif "temp" in cl:                                   col_map[col] = "Temperature"
        elif "hum" in cl:                                    col_map[col] = "Humidity"
        elif "dew" in cl:                                    col_map[col] = "DewPoint"
    df = df.rename(columns=col_map)

    if "AQI" not in df.columns or "Date" not in df.columns:
        return None, "لم يُعثر على أعمدة AQI / Date في ملف CSV.", "csv"

    for c in ["AQI","PM25","PM10","PM1","Temperature","Humidity"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Add missing columns with defaults
    if "CO"        not in df.columns: df["CO"]    = np.nan
    if "SMOKE"     not in df.columns: df["SMOKE"] = df.get("PM25", np.nan)
    if "RiskIndex" not in df.columns:
        df["RiskIndex"] = df["AQI"].apply(lambda v: float(v)*0.9 if pd.notna(v) else np.nan)
    if "Time"      not in df.columns: df["Time"]  = df["Date"]

    df = df[df["AQI"].notna() & (df["AQI"] > 0)]
    return _enrich(df), "", "csv"

def parse_file(contents: str, filename: str):
    fname = filename.lower()
    if fname.endswith(".csv"):
        return parse_csv(contents, filename)
    elif fname.endswith((".xlsx",".xls",".xlsm")):
        return parse_excel(contents, filename)
    else:
        # try Excel first, then CSV
        df, err, ftype = parse_excel(contents, filename)
        if df is not None: return df, err, ftype
        return parse_csv(contents, filename)

# ─────────────────────────────────────────────
#  AGGREGATION
# ─────────────────────────────────────────────
def daily_stats(df: pd.DataFrame) -> pd.DataFrame:
    agg_dict = dict(
        DayLabel   =("DayLabel",    "first"),
        AQI_mean   =("AQI",         "mean"),
        AQI_max    =("AQI",         "max"),
        Temp_mean  =("Temperature", "mean"),
        Hum_mean   =("Humidity",    "mean"),
    )
    if "CO"    in df.columns: agg_dict["CO_mean"]    = ("CO",    "mean")
    if "SMOKE" in df.columns: agg_dict["SMOKE_mean"] = ("SMOKE", "mean")
    if "PM25"  in df.columns: agg_dict["PM25_mean"]  = ("PM25",  "mean")
    if "PM10"  in df.columns: agg_dict["PM10_mean"]  = ("PM10",  "mean")
    return df.groupby("DateStr").agg(**agg_dict).round(2).reset_index()

# ─────────────────────────────────────────────
#  CHART BUILDERS
# ─────────────────────────────────────────────
def _layout(title):
    return dict(
        title=dict(text=title, font=dict(size=13,color="#1a1a2e"), x=0.5, xanchor="center"),
        paper_bgcolor="white", plot_bgcolor="#f8f9fa",
        font=dict(family="Cairo, Arial", color="#333"),
        margin=dict(t=52,b=42,l=50,r=30),
        legend=dict(bgcolor="rgba(255,255,255,0.85)",bordercolor="#ddd",borderwidth=1),
        hovermode="x unified",
        xaxis=dict(gridcolor="#e8e8e8"),
        yaxis=dict(gridcolor="#e8e8e8"),
    )

def build_aqi_bar(daily):
    colors = [aqi_color(v) for v in daily["AQI_mean"]]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=daily["DayLabel"], y=daily["AQI_mean"],
        marker_color=colors, marker_line_color="rgba(0,0,0,0.12)",
        marker_line_width=1,
        text=daily["AQI_mean"].round(0).astype(int),
        textposition="outside", name="AQI",
        hovertemplate="<b>%{x}</b><br>AQI: %{y:.0f}<extra></extra>",
    ))
    for thresh, lbl, col in [(50,"Good","#00E400"),(100,"Moderate","#FFFF00"),
                              (150,"Unhealthy*","#FF7E00"),(200,"Unhealthy","#FF0000"),
                              (300,"Very Unhealthy","#8F3F97")]:
        fig.add_hline(y=thresh, line_dash="dot", line_color=col, line_width=1.2,
                      opacity=0.75, annotation_text=lbl,
                      annotation_font_size=9, annotation_font_color=col)
    fig.update_layout(**_layout("متوسط مؤشر AQI اليومي  |  Daily AQI"),
                      yaxis_title="AQI",
                      yaxis_range=[0, max(daily["AQI_mean"].max()*1.22, 220)])
    return fig

def build_temp_humidity(daily):
    fig = make_subplots(specs=[[{"secondary_y":True}]])
    fig.add_trace(go.Scatter(x=daily["DayLabel"], y=daily["Temp_mean"],
        mode="lines+markers", name="حرارة (°م)",
        line=dict(color="#FF7E00",width=2.5), marker=dict(size=7),
        hovertemplate="%{y:.1f}°م<extra></extra>"), secondary_y=False)
    fig.add_trace(go.Bar(x=daily["DayLabel"], y=daily["Hum_mean"],
        name="رطوبة (%)", opacity=0.3, marker_color="#3498db",
        hovertemplate="%{y:.1f}%<extra></extra>"), secondary_y=True)
    fig.update_yaxes(title_text="°م", secondary_y=False, title_font_color="#FF7E00")
    fig.update_yaxes(title_text="%", secondary_y=True, title_font_color="#3498db")
    fig.update_layout(**_layout("الحرارة والرطوبة  |  Temp & Humidity"))
    return fig

def build_pm_chart(daily):
    """PM2.5 / PM10 chart — shown only for CSV files."""
    fig = go.Figure()
    if "PM25_mean" in daily.columns:
        fig.add_trace(go.Scatter(x=daily["DayLabel"], y=daily["PM25_mean"],
            mode="lines+markers", name="PM 2.5 (μg/m³)",
            line=dict(color="#e74c3c",width=2), marker=dict(size=6),
            hovertemplate="PM2.5: %{y:.1f}<extra></extra>"))
    if "PM10_mean" in daily.columns:
        fig.add_trace(go.Scatter(x=daily["DayLabel"], y=daily["PM10_mean"],
            mode="lines+markers", name="PM 10 (μg/m³)",
            line=dict(color="#8F3F97",width=2,dash="dash"), marker=dict(size=6),
            hovertemplate="PM10: %{y:.1f}<extra></extra>"))
    # WHO limits
    fig.add_hline(y=15, line_dash="dot", line_color="#e74c3c", line_width=1,
                  annotation_text="WHO PM2.5 limit (15)", annotation_font_size=9)
    fig.add_hline(y=45, line_dash="dot", line_color="#8F3F97", line_width=1,
                  annotation_text="WHO PM10 limit (45)", annotation_font_size=9)
    fig.update_layout(**_layout("جسيمات الغبار  |  Particulate Matter (PM)"),
                      yaxis_title="μg/m³")
    return fig

def build_co_smoke(daily):
    fig = go.Figure()
    if "CO_mean" in daily.columns and daily["CO_mean"].notna().any():
        fig.add_trace(go.Bar(x=daily["DayLabel"], y=daily["CO_mean"],
            name="CO (ppm)", marker_color="#8F3F97",
            hovertemplate="CO: %{y:.0f} ppm<extra></extra>"))
    if "SMOKE_mean" in daily.columns and daily["SMOKE_mean"].notna().any():
        fig.add_trace(go.Bar(x=daily["DayLabel"], y=daily["SMOKE_mean"],
            name="SMOKE / PM2.5", marker_color="#c0392b",
            hovertemplate="%{y:.0f}<extra></extra>"))
    fig.update_layout(**_layout("الملوثات  |  Pollutants"), barmode="group")
    return fig

def build_risk_donut(df):
    counts = df["RiskLevel"].value_counts()
    order  = ["آمن","متوسط","تحذير","حرج","طوارئ"]
    vals   = [counts.get(k,0) for k in order]
    colors = [RISK_COLORS[k] for k in order]
    fig = go.Figure(go.Pie(
        labels=order, values=vals, hole=0.58,
        marker=dict(colors=colors, line=dict(color="#fff",width=2)),
        textfont_size=12,
        hovertemplate="<b>%{label}</b><br>%{value:,} قراءة (%{percent})<extra></extra>",
    ))
    fig.add_annotation(text=f"<b>{sum(vals):,}</b><br>قراءة",
                       x=0.5,y=0.5,showarrow=False,font=dict(size=14,color="#333"))
    fig.update_layout(**_layout("توزيع مستويات الخطر  |  Risk Distribution"))
    return fig

def build_hourly(df):
    hourly = df.dropna(subset=["Hour"]).groupby("Hour")["AQI"].mean().reindex(range(24)).reset_index()
    hourly.columns=["Hour","AQI"]
    colors=[aqi_color(v) if pd.notna(v) else "#ccc" for v in hourly["AQI"]]
    fig=go.Figure()
    fig.add_trace(go.Bar(x=hourly["Hour"],y=hourly["AQI"].fillna(0),
        marker_color=colors,name="AQI ساعي",
        hovertemplate="الساعة %{x}:00<br>AQI: %{y:.0f}<extra></extra>"))
    fig.add_hline(y=100,line_dash="dot",line_color="#FF7E00",line_width=1.2,
                  annotation_text="100",annotation_font_color="#FF7E00",annotation_font_size=9)
    fig.update_layout(**_layout("متوسط AQI لكل ساعة  |  Hourly AQI"),
        xaxis=dict(tickmode="array",tickvals=[0,3,6,9,12,15,18,21],
                   ticktext=["12AM","3AM","6AM","9AM","12PM","3PM","6PM","9PM"]))
    return fig

def build_trend_forecast(daily):
    aqi=daily["AQI_mean"].values; x=np.arange(len(aqi))
    slope,intercept,r_val,p_val,se=stats.linregress(x,aqi)
    trend=slope*x+intercept
    x_fut=np.arange(len(aqi),len(aqi)+7); y_fut=slope*x_fut+intercept
    ci=1.96*se*np.sqrt(1+1/len(x)+(x_fut-x.mean())**2/((x-x.mean())**2).sum())
    smooth=gaussian_filter1d(aqi,sigma=1.0)

    fig=go.Figure()
    for lo,hi,col in [(0,50,"#00E400"),(50,100,"#FFFF00"),(100,150,"#FF7E00"),
                       (150,200,"#FF0000"),(200,300,"#8F3F97"),(300,500,"#7E0023")]:
        fig.add_hrect(y0=lo,y1=hi,fillcolor=col,opacity=0.07,line_width=0)
    fig.add_trace(go.Scatter(x=list(daily["DayLabel"]),y=smooth,
        mode="lines",name="AQI (منعّم)",line=dict(color="#3498db",width=2.5)))
    fig.add_trace(go.Scatter(x=list(daily["DayLabel"]),y=aqi,
        mode="markers",name="AQI فعلي",
        marker=dict(color=[aqi_color(v) for v in aqi],size=11,
                    line=dict(color="white",width=1.5))))
    all_x=list(daily["DayLabel"])+[f"يوم+{j+1}" for j in range(7)]
    fig.add_trace(go.Scatter(x=all_x,y=np.append(trend,y_fut),
        mode="lines",name=f"اتجاه ({slope:+.1f}/يوم)",
        line=dict(color="#FF7E00",width=2,dash="dash")))
    fig.add_trace(go.Scatter(x=[f"يوم+{j+1}" for j in range(7)],y=np.clip(y_fut,0,500),
        mode="markers+text",
        marker=dict(color=[aqi_color(max(0,v)) for v in y_fut],size=13,
                    symbol="diamond",line=dict(color="white",width=1.5)),
        text=[f"{v:.0f}" for v in y_fut],textposition="top center",
        textfont=dict(size=9),name="تنبؤ 7 أيام"))
    fig.add_trace(go.Scatter(
        x=[f"يوم+{j+1}" for j in range(7)]+[f"يوم+{j+1}" for j in range(6,-1,-1)],
        y=list(np.clip(y_fut+ci,0,600))+list(np.clip(y_fut-ci,0,600)),
        fill="toself",fillcolor="rgba(143,63,151,0.15)",
        line=dict(color="rgba(0,0,0,0)"),name="فترة ثقة 95%"))
    fig.add_vline(x=len(aqi)-0.5,line_dash="dot",line_color="#888",line_width=1)
    fig.update_layout(**_layout(f"تحليل الاتجاه والتنبؤ  |  R²={r_val**2:.3f}  p={p_val:.3f}"),
                      xaxis_tickangle=-30)
    return fig

# ─────────────────────────────────────────────
#  UI COMPONENTS
# ─────────────────────────────────────────────
def aqi_scale_card():
    rows=[]
    for i,(rng,ar,en,desc,bg,tc) in enumerate(SCALE_DATA):
        rows.append(dbc.Row([
            dbc.Col(html.Div(FACE_EMOJIS[i],style={
                "backgroundColor":bg,"borderRadius":"10px","fontSize":"2.1rem",
                "textAlign":"center","padding":"6px 4px","lineHeight":"1.1","minWidth":"60px"}),
                width="auto"),
            dbc.Col([
                html.Div([html.Span(rng,style={"fontWeight":"700","fontSize":"0.95rem","marginLeft":"8px"}),
                          html.Span(ar, style={"fontWeight":"700","fontSize":"0.95rem","color":"#222"})]),
                html.Div(desc,style={"fontSize":"0.76rem","color":"#555","marginTop":"2px"}),
            ],style={"paddingRight":"4px"}),
        ],align="center",style={"borderBottom":"1px solid #eee","padding":"8px 10px","margin":"0"}))
    return dbc.Card([
        dbc.CardHeader(html.H6("دليل مؤشر جودة الهواء  |  AQI Reference Guide",
            className="mb-0 text-center fw-bold"),
            style={"backgroundColor":"#f0f4ff","borderBottom":"2px solid #2c5f9e"}),
        dbc.CardBody(rows,style={"padding":"0"}),
    ],style={"borderRadius":"12px","overflow":"hidden","border":"1px solid #dde"})

def stat_cards(df, ftype):
    aqi_mean=df["AQI"].mean(); lv=get_aqi_level(aqi_mean)
    tc="#000" if lv["color"] in ("#00E400","#FFFF00","#FF7E00") else "#fff"
    cards=[
        ("📊", f"{len(df):,}",            "إجمالي القراءات",       "#2c5f9e"),
        ("🌡️", f"{df['Temperature'].mean():.1f}°م" if "Temperature" in df and df["Temperature"].notna().any() else "—",
                                           "متوسط الحرارة",         "#e67e22"),
        ("💨", f"{aqi_mean:.1f}",          f"متوسط AQI",            lv["color"] if lv["color"]!="#FFFF00" else "#b8a000"),
        ("⚠️", f"{df['AQI'].max():.0f}",   "أعلى AQI مُسجّل",      "#c0392b"),
        ("📅", f"{df['DateStr'].nunique()}","أيام الرصد",            "#27ae60"),
        ("📁", "CSV" if ftype=="csv" else "Excel", "نوع الملف",     "#555"),
    ]
    return dbc.Row([
        dbc.Col(dbc.Card([dbc.CardBody([
            html.Div(icon,style={"fontSize":"1.5rem","textAlign":"center"}),
            html.H5(val, className="text-center fw-bold mb-0",style={"color":color,"fontSize":"1.2rem"}),
            html.P(label,className="text-center text-muted mb-0",style={"fontSize":"0.7rem"}),
        ],style={"padding":"10px 5px"})],
        style={"borderRadius":"10px","border":f"1px solid {color}33","backgroundColor":"#fff","height":"100%"}),
        width=2) for icon,val,label,color in cards
    ],className="g-2 mb-3")

def file_info_badge(filename, ftype, df):
    color = "success" if ftype=="csv" else "primary"
    icon  = "📄" if ftype=="csv" else "📊"
    return dbc.Alert([
        html.B(f"{icon} {filename}  "),
        f"— {len(df):,} قراءة  |  {df['DateStr'].nunique()} يوم  |  "
        f"نوع: {'CSV (AirLink)' if ftype=='csv' else 'Excel'}",
    ], color=color, className="mt-2 mb-0 py-2")

# ─────────────────────────────────────────────
#  APP LAYOUT
# ─────────────────────────────────────────────
app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.BOOTSTRAP,
        "https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap"],
    title="EnviroAI — الجامعة المستنصرية",
    suppress_callback_exceptions=True,
)
server = app.server  # gunicorn entry point

header = dbc.Navbar(dbc.Container([
    html.Div([
        html.H5("🌍 EnviroAI",className="mb-0 fw-bold",style={"color":"white","fontSize":"1.2rem"}),
        html.Small("Air Quality & Weather Analysis System",
                   style={"color":"#a8c8ff","display":"block","fontSize":"0.7rem"}),
    ]),
    html.Div([
        html.Div("كلية العلوم / الجامعة المستنصرية",
                 style={"color":"white","fontWeight":"600","textAlign":"right","fontSize":"0.82rem"}),
        html.Div("بالتعاون مع وزارة البيئة / مديرية البيئة الحضرية — قسم مراقبة نوعية الهواء والضوضاء",
                 style={"color":"#a8c8ff","textAlign":"right","fontSize":"0.68rem"}),
    ]),
],fluid=True,style={"display":"flex","justifyContent":"space-between","alignItems":"center"}),
color="#1a3a6b",dark=True,style={"padding":"10px 0","direction":"rtl"})

upload_section = dbc.Container([
    dbc.Row(dbc.Col(
        dcc.Upload(id="upload-data",
            children=html.Div([
                html.Div("📁",style={"fontSize":"2.4rem"}),
                html.Div("اسحب وأفلت ملف البيانات هنا",style={"fontWeight":"600","fontSize":"1rem"}),
                html.Div("يدعم Excel (.xlsx) و CSV (تصدير AirLink)",
                         style={"color":"#555","fontSize":"0.85rem","marginTop":"4px"}),
                html.Div("Date/Time, AQI, Temperature, Humidity, PM2.5, PM10, CO, SMOKE …",
                         style={"color":"#aaa","fontSize":"0.72rem","marginTop":"3px"}),
            ],style={"textAlign":"center","padding":"18px"}),
            style={"border":"2px dashed #2c5f9e","borderRadius":"14px",
                   "backgroundColor":"#f0f5ff","cursor":"pointer"},
            multiple=False),
        width={"size":8,"offset":2}),className="my-3"),
    dbc.Row(dbc.Col(html.Div(id="upload-status"),width={"size":8,"offset":2})),
],fluid=True)

app.layout = html.Div([
    header,
    upload_section,
    html.Div(id="dashboard-content"),
    dcc.Store(id="stored-data"),
    dcc.Store(id="stored-ftype"),
],style={"fontFamily":"Cairo, Arial, sans-serif","direction":"rtl","backgroundColor":"#f5f7fb"})

# ─────────────────────────────────────────────
#  CALLBACKS
# ─────────────────────────────────────────────
@app.callback(
    Output("stored-data",   "data"),
    Output("stored-ftype",  "data"),
    Output("upload-status", "children"),
    Input("upload-data",    "contents"),
    State("upload-data",    "filename"),
    prevent_initial_call=True,
)
def store_upload(contents, filename):
    if not contents: return None, None, ""
    df, err, ftype = parse_file(contents, filename)
    if err:
        return None, None, dbc.Alert(f"❌ {err}", color="danger", className="mt-2")
    return (df.to_json(date_format="iso", orient="split"),
            ftype,
            file_info_badge(filename, ftype, df))

@app.callback(
    Output("dashboard-content","children"),
    Input("stored-data",  "data"),
    Input("stored-ftype", "data"),
    prevent_initial_call=True,
)
def render_dashboard(json_data, ftype):
    if not json_data: return ""
    df = _load_df(json_data)
    return _build_layout(df, ftype or "excel")

def _load_df(json_data):
    df = pd.read_json(json_data, orient="split")
    df["Date"]     = pd.to_datetime(df["Date"], errors="coerce")
    df["DateStr"]  = df["Date"].dt.strftime("%Y-%m-%d")
    df["DayLabel"] = df["Date"].dt.strftime("%-d %b")
    df["Hour"]     = pd.to_datetime(df["Time"].astype(str), errors="coerce").dt.hour \
                     if "Time" in df.columns else df["Date"].dt.hour
    return df

def _build_layout(df, ftype):
    daily   = daily_stats(df)
    has_pm  = "PM25_mean" in daily.columns or "PM10_mean" in daily.columns
    tabs_list = [
        dbc.Tab(label="📊 نظرة عامة",   tab_id="overview"),
        dbc.Tab(label="⏰ نمط ساعي",     tab_id="hourly"),
        dbc.Tab(label="⚠️ مستوى الخطر", tab_id="risk"),
        dbc.Tab(label="📈 تنبؤ واتجاه", tab_id="forecast"),
        dbc.Tab(label="📋 دليل AQI",    tab_id="guide"),
    ]
    if has_pm:
        tabs_list.insert(1, dbc.Tab(label="🌫️ PM2.5/PM10", tab_id="pm"))

    return dbc.Container([
        html.Hr(style={"margin":"8px 0"}),
        dbc.Tabs(tabs_list, id="main-tabs", active_tab="overview",
                 style={"marginBottom":"14px"}),
        html.Div(id="tab-content-area"),
        dcc.Store(id="df-store-local",  data=df.to_json(date_format="iso",orient="split")),
        dcc.Store(id="ftype-store-local",data=ftype),
    ],fluid=True,style={"paddingBottom":"40px"})

@app.callback(
    Output("tab-content-area","children"),
    Input("main-tabs",        "active_tab"),
    State("df-store-local",   "data"),
    State("ftype-store-local","data"),
    prevent_initial_call=True,
)
def switch_tab(tab, json_data, ftype):
    if not json_data: return ""
    df    = _load_df(json_data)
    daily = daily_stats(df)

    if tab == "overview":
        return dbc.Container([
            stat_cards(df, ftype or "excel"),
            dbc.Row([
                dbc.Col(dcc.Graph(figure=build_aqi_bar(daily),    config={"displayModeBar":False}),width=12,className="mb-3"),
                dbc.Col(dcc.Graph(figure=build_temp_humidity(daily),config={"displayModeBar":False}),width=6),
                dbc.Col(dcc.Graph(figure=build_co_smoke(daily),   config={"displayModeBar":False}),width=6),
            ])],fluid=True)

    if tab == "pm":
        return dbc.Container(dbc.Row(
            dbc.Col(dcc.Graph(figure=build_pm_chart(daily),config={"displayModeBar":False}),width=12)
        ),fluid=True)

    if tab == "hourly":
        return dbc.Container(dbc.Row(
            dbc.Col(dcc.Graph(figure=build_hourly(df),config={"displayModeBar":False}),width=12)
        ),fluid=True)

    if tab == "risk":
        rc = df["RiskLevel"].value_counts().reset_index()
        rc.columns=["المستوى","العدد"]
        rc["النسبة"]=(rc["العدد"]/len(df)*100).round(1).astype(str)+"%"
        return dbc.Container([
            dbc.Row([
                dbc.Col(dcc.Graph(figure=build_risk_donut(df),config={"displayModeBar":False}),width=5),
                dbc.Col([
                    dbc.Table.from_dataframe(rc[["المستوى","العدد","النسبة"]],
                        striped=True,bordered=True,hover=True,
                        style={"fontSize":"0.88rem","textAlign":"center"},className="mt-3"),
                    html.P(f"القراءات الخطرة (حرج + طوارئ): {df[df['RiskLevel'].isin(['حرج','طوارئ'])].shape[0]:,}",
                           className="text-muted mt-2",style={"fontSize":"0.8rem"}),
                ],width=7),
            ])],fluid=True)

    if tab == "forecast":
        return dbc.Container([
            dbc.Row(dbc.Col(dcc.Graph(figure=build_trend_forecast(daily),config={"displayModeBar":False}),width=12)),
            dbc.Row(dbc.Col(dbc.Alert([html.B("ملاحظة: "),
                "التنبؤ مبني على الانحدار الخطي — للتوجيه فقط."],color="info"),
                width={"size":10,"offset":1})),
        ],fluid=True)

    if tab == "guide":
        return dbc.Container(
            dbc.Row(dbc.Col(aqi_scale_card(),width={"size":8,"offset":2}),className="my-3"),
        fluid=True)
    return ""

# ─────────────────────────────────────────────
#  RUN
# ─────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
