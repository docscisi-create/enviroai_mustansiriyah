"""
EnviroAI — Air Quality & Weather Analysis System
كلية العلوم / الجامعة المستنصرية
"""
import base64, io, os, json, uuid
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
import warnings
warnings.filterwarnings("ignore")

# ── In-memory session cache (server-side) ────────────────────────────
_CACHE = {}   # session_id -> {"df": df, "ftype": ftype, "meta": {...}}

# ══════════════════════════════════════════════
#  AQI SCALE
# ══════════════════════════════════════════════
AQI_LEVELS = [
    {"range":(0,50),   "color":"#00E400","ar":"جيد",              "tc":"#000"},
    {"range":(51,100), "color":"#FFFF00","ar":"معتدل",            "tc":"#000"},
    {"range":(101,150),"color":"#FF7E00","ar":"غير صحي للحساسين","tc":"#000"},
    {"range":(151,200),"color":"#FF0000","ar":"غير صحي",          "tc":"#fff"},
    {"range":(201,300),"color":"#8F3F97","ar":"غير صحي جداً",     "tc":"#fff"},
    {"range":(301,500),"color":"#7E0023","ar":"خطير",             "tc":"#fff"},
]
RISK_COLORS = {"آمن":"#00E400","متوسط":"#FFFF00","تحذير":"#FF7E00","حرج":"#FF0000","طوارئ":"#8F3F97"}
FACES = ["😊","🙂","😐","😷","😷","😷"]
SCALE_DESCS = [
    "جودة الهواء مرضية وتشكل خطرًا قليلاً.",
    "يجب على الأفراد الحساسين تجنب الأنشطة الخارجية المكثفة.",
    "الجمهور والحساسون معرضون لخطر مشاكل تنفسية.",
    "يزداد احتمال حدوث آثار ضارة على الجمهور.",
    "سيتأثر عموم الجمهور. يُنصح بالبقاء في المنازل.",
    "خطير جداً. يجب على الجميع البقاء في الداخل.",
]

def get_lv(v):
    v = max(0, float(v))
    for lv in AQI_LEVELS:
        if lv["range"][0] <= v <= lv["range"][1]: return lv
    return AQI_LEVELS[-1]

def aqi_color(v):
    try: return get_lv(v)["color"]
    except: return "#ccc"

# ══════════════════════════════════════════════
#  FILE PARSING
# ══════════════════════════════════════════════
REQUIRED_XL = {"Date","Time","Temperature","Humidity","AQI","CO","SMOKE","RiskIndex","RiskLevel"}

def _add_derived(df):
    df["Date"]     = pd.to_datetime(df["Date"], errors="coerce")
    df             = df.dropna(subset=["Date"])
    df["DateStr"]  = df["Date"].dt.strftime("%Y-%m-%d")
    df["DayLabel"] = df["Date"].dt.strftime("%-d %b")
    if "Time" in df.columns:
        df["Hour"] = pd.to_datetime(df["Time"].astype(str), errors="coerce").dt.hour
    else:
        df["Hour"] = df["Date"].dt.hour
    if "RiskLevel" not in df.columns:
        def _rl(v):
            if v<=50:  return "آمن"
            if v<=100: return "متوسط"
            if v<=150: return "تحذير"
            if v<=200: return "حرج"
            return "طوارئ"
        df["RiskLevel"] = df["AQI"].apply(lambda v: _rl(float(v)) if pd.notna(v) else "متوسط")
    return df

def parse_file(contents, filename):
    _, b64 = contents.split(",", 1)
    raw    = base64.b64decode(b64)
    fname  = filename.lower()

    if fname.endswith((".xlsx",".xls",".xlsm")):
        try:
            sheets = pd.read_excel(io.BytesIO(raw), sheet_name=None)
        except Exception as e:
            return None, f"تعذّر قراءة Excel: {e}", "excel"
        for _, sdf in sheets.items():
            if REQUIRED_XL.issubset(set(sdf.columns)):
                return _add_derived(sdf.copy()), "", "excel"
        return None, "لم يُعثر على الأعمدة المطلوبة في ملف Excel.\nالأعمدة المطلوبة: Date, Time, Temperature, Humidity, AQI, CO, SMOKE, RiskIndex, RiskLevel", "excel"

    try:
        enc  = chardet.detect(raw)["encoding"] or "latin-1"
        text = raw.decode(enc, errors="replace")
        tmp  = pd.read_csv(io.StringIO(text), header=None, on_bad_lines="skip")
        hrow = 0
        for i, row in tmp.iterrows():
            vals = [str(v).lower() for v in row if pd.notna(v)]
            if any("date" in v or "time" in v for v in vals):
                hrow = i; break
        df = pd.read_csv(io.StringIO(text), header=None, skiprows=hrow+1, on_bad_lines="skip")
        cols = [str(c).strip() if pd.notna(c) else f"c{i}"
                for i,c in enumerate(tmp.iloc[hrow].tolist()[:len(df.columns)])]
        df.columns = cols
        rename = {}
        for c in df.columns:
            cl = c.lower()
            if ("date" in cl or "time" in cl) and "Date" not in rename.values(): rename[c]="Date"
            elif "aqi" in cl and "AQI" not in rename.values():                   rename[c]="AQI"
            elif ("2.5" in cl or "pm2.5" in cl or "pm 2.5" in cl):              rename[c]="PM25"
            elif ("pm 10" in cl or "pm10" in cl):                               rename[c]="PM10"
            elif ("pm 1" in cl or "pm1" in cl) and "PM1" not in rename.values():rename[c]="PM1"
            elif "temp" in cl and "Temperature" not in rename.values():         rename[c]="Temperature"
            elif "hum" in cl and "Humidity" not in rename.values():             rename[c]="Humidity"
            elif "dew" in cl:                                                   rename[c]="DewPoint"
        df = df.rename(columns=rename)
        if "AQI" not in df.columns or "Date" not in df.columns:
            return None, "ملف CSV لا يحتوي أعمدة AQI أو Date المطلوبة.", "csv"
        for c in ["AQI","PM25","PM10","PM1","Temperature","Humidity"]:
            if c in df.columns: df[c] = pd.to_numeric(df[c], errors="coerce")
        if "CO"    not in df.columns: df["CO"]    = np.nan
        if "SMOKE" not in df.columns: df["SMOKE"] = df.get("PM25", np.nan)
        if "RiskIndex" not in df.columns:
            df["RiskIndex"] = df["AQI"].apply(lambda v: float(v)*0.9 if pd.notna(v) else np.nan)
        df = df[df["AQI"].notna() & (df["AQI"]>0)]
        return _add_derived(df), "", "csv"
    except Exception as e:
        return None, f"تعذّر قراءة CSV: {e}", "csv"

def daily_stats(df):
    agg = dict(
        DayLabel  =("DayLabel","first"),
        AQI_mean  =("AQI","mean"), AQI_max=("AQI","max"),
    )
    for col,key in [("Temperature","Temp_mean"),("Humidity","Hum_mean"),
                    ("CO","CO_mean"),("SMOKE","SMOKE_mean"),
                    ("PM25","PM25_mean"),("PM10","PM10_mean")]:
        if col in df.columns: agg[key]=(col,"mean")
    return df.groupby("DateStr").agg(**agg).round(2).reset_index()

# ══════════════════════════════════════════════
#  CHARTS
# ══════════════════════════════════════════════
def _base(title):
    return dict(
        title=dict(text=title,font=dict(size=13,color="#1a1a2e"),x=0.5,xanchor="center"),
        paper_bgcolor="white", plot_bgcolor="#f8f9fa",
        font=dict(family="Cairo,Arial",color="#333"),
        margin=dict(t=55,b=45,l=55,r=35),
        legend=dict(bgcolor="rgba(255,255,255,.9)",bordercolor="#ddd",borderwidth=1),
        hovermode="x unified",
        xaxis=dict(gridcolor="#ebebeb",tickangle=-30),
        yaxis=dict(gridcolor="#ebebeb"),
    )

def ch_aqi_bar(daily):
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=daily["DayLabel"], y=daily["AQI_mean"],
        marker_color=[aqi_color(v) for v in daily["AQI_mean"]],
        marker_line_color="rgba(0,0,0,.1)", marker_line_width=1,
        text=daily["AQI_mean"].round(0).astype(int), textposition="outside",
        hovertemplate="<b>%{x}</b><br>AQI: %{y:.0f}<extra></extra>",
    ))
    for y,lbl,col in [(50,"جيد 50","#00E400"),(100,"معتدل 100","#b8a000"),
                      (150,"غير صحي 150","#FF7E00"),(200,"خطر 200","#FF0000"),(300,"طوارئ 300","#8F3F97")]:
        fig.add_hline(y=y,line_dash="dot",line_color=col,line_width=1.3,opacity=.8,
                      annotation_text=lbl,annotation_position="right",
                      annotation_font_size=9,annotation_font_color=col)
    fig.update_layout(**_base("متوسط AQI اليومي  |  Daily Average AQI"),
                      yaxis_title="AQI",
                      yaxis_range=[0,max(daily["AQI_mean"].max()*1.25,230)])
    return fig

def ch_temp_hum(daily):
    fig = make_subplots(specs=[[{"secondary_y":True}]])
    if "Temp_mean" in daily.columns:
        fig.add_trace(go.Scatter(x=daily["DayLabel"],y=daily["Temp_mean"],
            mode="lines+markers",name="حرارة (°م)",
            line=dict(color="#e67e22",width=2.5),marker=dict(size=7)),secondary_y=False)
    if "Hum_mean" in daily.columns:
        fig.add_trace(go.Bar(x=daily["DayLabel"],y=daily["Hum_mean"],
            name="رطوبة (%)",opacity=.28,marker_color="#3498db"),secondary_y=True)
    fig.update_yaxes(title_text="°م",secondary_y=False,title_font_color="#e67e22",gridcolor="#ebebeb")
    fig.update_yaxes(title_text="%", secondary_y=True, title_font_color="#3498db",showgrid=False)
    b = _base("درجة الحرارة والرطوبة  |  Temperature & Humidity")
    fig.update_layout(**b)
    return fig

def ch_pm(daily):
    fig = go.Figure()
    if "PM25_mean" in daily.columns:
        fig.add_trace(go.Scatter(x=daily["DayLabel"],y=daily["PM25_mean"],
            mode="lines+markers",name="PM 2.5 (μg/m³)",
            line=dict(color="#e74c3c",width=2.5),marker=dict(size=7)))
    if "PM10_mean" in daily.columns:
        fig.add_trace(go.Scatter(x=daily["DayLabel"],y=daily["PM10_mean"],
            mode="lines+markers",name="PM 10 (μg/m³)",
            line=dict(color="#8F3F97",width=2,dash="dash"),marker=dict(size=6)))
    fig.add_hline(y=15,line_dash="dot",line_color="#e74c3c",line_width=1.2,
                  annotation_text="WHO PM2.5 ≤15",annotation_font_size=9,annotation_font_color="#e74c3c")
    fig.add_hline(y=45,line_dash="dot",line_color="#8F3F97",line_width=1.2,
                  annotation_text="WHO PM10 ≤45",annotation_font_size=9,annotation_font_color="#8F3F97")
    fig.update_layout(**_base("جسيمات الغبار  |  Particulate Matter (PM)"),yaxis_title="μg/m³")
    return fig

def ch_pollutants(daily):
    fig = go.Figure()
    if "CO_mean" in daily.columns and daily["CO_mean"].notna().any():
        fig.add_trace(go.Bar(x=daily["DayLabel"],y=daily["CO_mean"],name="CO (ppm)",marker_color="#8F3F97"))
    smoke_col = "SMOKE_mean" if "SMOKE_mean" in daily.columns else "PM25_mean"
    if smoke_col in daily.columns and daily[smoke_col].notna().any():
        fig.add_trace(go.Bar(x=daily["DayLabel"],y=daily[smoke_col],name="SMOKE / PM2.5",marker_color="#c0392b"))
    fig.update_layout(**_base("الملوثات  |  Pollutants"),barmode="group")
    return fig

def ch_donut(df):
    counts=df["RiskLevel"].value_counts()
    order=["آمن","متوسط","تحذير","حرج","طوارئ"]
    vals=[counts.get(k,0) for k in order]
    fig=go.Figure(go.Pie(labels=order,values=vals,hole=.58,
        marker=dict(colors=[RISK_COLORS[k] for k in order],line=dict(color="#fff",width=2)),
        hovertemplate="<b>%{label}</b><br>%{value:,} قراءة  %{percent}<extra></extra>"))
    fig.add_annotation(text=f"<b>{sum(vals):,}</b><br>قراءة",x=.5,y=.5,showarrow=False,
                       font=dict(size=14,color="#333"))
    fig.update_layout(**_base("توزيع مستويات الخطر  |  Risk Distribution"))
    return fig

def ch_hourly(df):
    h=df.dropna(subset=["Hour"]).groupby("Hour")["AQI"].mean().reindex(range(24)).reset_index()
    h.columns=["Hour","AQI"]
    fig=go.Figure(go.Bar(x=h["Hour"],y=h["AQI"].fillna(0),
        marker_color=[aqi_color(v) if pd.notna(v) else "#ddd" for v in h["AQI"]],
        hovertemplate="الساعة %{x}:00<br>AQI: %{y:.0f}<extra></extra>"))
    fig.add_hline(y=100,line_dash="dot",line_color="#FF7E00",line_width=1.5,
                  annotation_text="100",annotation_font_color="#FF7E00",annotation_font_size=9)
    lyt = _base("متوسط AQI لكل ساعة  |  Hourly AQI Pattern")
    lyt["xaxis"] = dict(tickmode="array",tickvals=[0,3,6,9,12,15,18,21],
                   ticktext=["12AM","3AM","6AM","9AM","12PM","3PM","6PM","9PM"],
                   tickangle=0, gridcolor="#ebebeb")
    fig.update_layout(**lyt)
    return fig

def ch_trend(daily):
    aqi=daily["AQI_mean"].values; x=np.arange(len(aqi))
    slope,intercept,r,p,se=stats.linregress(x,aqi)
    trend=slope*x+intercept
    xf=np.arange(len(aqi),len(aqi)+7); yf=np.clip(slope*xf+intercept,0,500)
    ci=1.96*se*np.sqrt(1+1/len(x)+(xf-x.mean())**2/((x-x.mean())**2).sum())
    smooth=gaussian_filter1d(aqi,sigma=1.0)
    fig=go.Figure()
    for lo,hi,col in [(0,50,"#00E400"),(50,100,"#FFFF00"),(100,150,"#FF7E00"),
                      (150,200,"#FF0000"),(200,300,"#8F3F97"),(300,500,"#7E0023")]:
        fig.add_hrect(y0=lo,y1=hi,fillcolor=col,opacity=.07,line_width=0)
    fig.add_trace(go.Scatter(x=list(daily["DayLabel"]),y=smooth,mode="lines",
        name="AQI منعّم",line=dict(color="#3498db",width=2.5)))
    fig.add_trace(go.Scatter(x=list(daily["DayLabel"]),y=aqi,mode="markers",name="AQI فعلي",
        marker=dict(color=[aqi_color(v) for v in aqi],size=11,line=dict(color="white",width=1.5))))
    all_x=list(daily["DayLabel"])+[f"+{j+1}" for j in range(7)]
    fig.add_trace(go.Scatter(x=all_x,y=np.append(trend,yf),mode="lines",
        name=f"اتجاه ({slope:+.1f}/يوم)",line=dict(color="#FF7E00",width=2,dash="dash")))
    fig.add_trace(go.Scatter(x=[f"+{j+1}" for j in range(7)],y=yf,mode="markers+text",
        marker=dict(color=[aqi_color(v) for v in yf],size=13,symbol="diamond",
                    line=dict(color="white",width=1.5)),
        text=[f"{v:.0f}" for v in yf],textposition="top center",textfont=dict(size=9),
        name="تنبؤ 7 أيام"))
    fig.add_trace(go.Scatter(
        x=[f"+{j+1}" for j in range(7)]+[f"+{j+1}" for j in range(6,-1,-1)],
        y=list(np.clip(yf+ci,0,600))+list(np.clip(yf-ci,0,600)),
        fill="toself",fillcolor="rgba(143,63,151,.15)",
        line=dict(color="rgba(0,0,0,0)"),name="فترة ثقة 95%"))
    fig.add_vline(x=len(aqi)-.5,line_dash="dot",line_color="#aaa",line_width=1)
    fig.update_layout(**_base(f"تحليل الاتجاه والتنبؤ  |  R²={r**2:.3f}  p={p:.3f}"),
                      xaxis_tickangle=-30)
    return fig

# ══════════════════════════════════════════════
#  UI HELPERS
# ══════════════════════════════════════════════
def kpi_row(df, ftype):
    aqi_m=df["AQI"].mean(); lv=get_lv(aqi_m)
    items=[
        ("📊","إجمالي القراءات",  f"{len(df):,}",          "#2c5f9e"),
        ("📅","أيام الرصد",       str(df["DateStr"].nunique()),"#27ae60"),
        ("🌡️","متوسط الحرارة",   f"{df['Temperature'].mean():.1f}°م" if "Temperature" in df.columns and df["Temperature"].notna().any() else "—", "#e67e22"),
        ("💨","متوسط AQI",        f"{aqi_m:.1f}",           lv["color"] if lv["color"]!="#FFFF00" else "#b8a000"),
        ("⚠️","أعلى AQI مُسجّل", f"{df['AQI'].max():.0f}", "#c0392b"),
        ("📁","نوع الملف",        "CSV (AirLink)" if ftype=="csv" else "Excel", "#555"),
    ]
    cols=[]
    for icon,label,val,color in items:
        cols.append(dbc.Col(dbc.Card(dbc.CardBody([
            html.Div(icon,style={"fontSize":"1.6rem","textAlign":"center","lineHeight":"1"}),
            html.Div(val, style={"fontWeight":"700","fontSize":"1.15rem","color":color,
                                 "textAlign":"center","margin":"4px 0 2px"}),
            html.Div(label,style={"fontSize":"0.68rem","color":"#777","textAlign":"center"}),
        ],style={"padding":"10px 6px"}),
        style={"borderRadius":"10px","border":f"2px solid {color}22",
               "backgroundColor":"#fff","height":"100%"}),width=2))
    return dbc.Row(cols, className="g-2 mb-3")

def aqi_guide_card():
    rows=[]
    for i,lv in enumerate(AQI_LEVELS):
        lo,hi=lv["range"]
        rows.append(dbc.Row([
            dbc.Col(html.Div(FACES[i],style={
                "background":lv["color"],"borderRadius":"10px","fontSize":"2rem",
                "textAlign":"center","padding":"5px 0","minWidth":"56px","lineHeight":"1.2"}),
                width="auto"),
            dbc.Col([
                html.Span(f"{lo}–{hi}  ",style={"fontWeight":"700","fontSize":".9rem"}),
                html.Span(lv["ar"],      style={"fontWeight":"700","fontSize":".9rem","color":"#222"}),
                html.Br(),
                html.Small(SCALE_DESCS[i],style={"color":"#666"}),
            ]),
        ],align="center",className="py-2",
          style={"borderBottom":"1px solid #eee","margin":"0","paddingRight":"8px"}))
    return dbc.Card([
        dbc.CardHeader(html.B("دليل مؤشر جودة الهواء  |  AQI Reference Guide",
            className="d-block text-center"),
            style={"background":"#eef2ff","borderBottom":"2px solid #2c5f9e"}),
        dbc.CardBody(rows,style={"padding":"0 8px"}),
    ],style={"borderRadius":"12px","overflow":"hidden"})

def welcome_screen():
    return dbc.Container([
        dbc.Row(dbc.Col([
            html.Div([
                html.Div("🌍", style={"fontSize":"4rem","textAlign":"center"}),
                html.H3("نظام EnviroAI لتحليل جودة الهواء",
                        className="text-center fw-bold mt-2",style={"color":"#1a3a6b"}),
                html.P("Air Quality & Weather Analysis System",
                       className="text-center text-muted mb-4"),
            ]),
            dbc.Card([
                dbc.CardHeader(html.B("📋 كيفية الاستخدام  |  How to Use"),
                               style={"background":"#eef2ff","borderBottom":"2px solid #2c5f9e"}),
                dbc.CardBody([
                    dbc.Row([
                        dbc.Col([
                            html.H6("📤 الخطوة 1 — رفع الملف", style={"color":"#1a3a6b"}),
                            html.P("اسحب وأفلت ملف البيانات في المربع أعلاه، أو انقر لاختياره من جهازك."),
                            html.Hr(),
                            html.H6("📊 الخطوة 2 — استعراض التحليل", style={"color":"#1a3a6b"}),
                            html.P("تظهر لوحة التحكم تلقائياً بعد رفع الملف مع 6 تبويبات تفاعلية."),
                            html.Hr(),
                            html.H6("🗂️ أنواع الملفات المدعومة", style={"color":"#1a3a6b","marginBottom":"8px"}),
                            dbc.Badge("✅ Excel (.xlsx)  — بيانات المحطة الرئيسية", color="primary", className="me-2 mb-2 d-block"),
                            dbc.Badge("✅ CSV — تصدير AirLink للمحطة المتنقلة", color="success", className="mb-1 d-block"),
                        ], width=5),
                        dbc.Col([
                            html.H6("📈 ماذا يعرض النظام؟", style={"color":"#1a3a6b"}),
                            html.Ul([
                                html.Li("📊 نظرة عامة — مؤشرات AQI اليومي + درجة الحرارة + الملوثات"),
                                html.Li("🌫️ PM2.5 / PM10 — جسيمات الغبار مع حدود منظمة الصحة العالمية"),
                                html.Li("⏰  كيف يتغير AQI خلال 24 ساعة"),
                                html.Li("⚠️ مستوى الخطر — توزيع الفئات (آمن / تحذير / حرج ...)"),
                                html.Li("📈 تنبؤ — تحليل الاتجاه + توقعات 7 أيام"),
                                html.Li("📋 دليل AQI — شرح المقياس الرسمي بالألوان"),
                            ], style={"fontSize":".88rem","paddingRight":"1.2rem"}),
                        ], width=7),
                    ]),
                ]),
            ], className="mb-4", style={"borderRadius":"12px"}),
            html.H6("🎨 مقياس AQI الرسمي:", className="mb-3", style={"color":"#1a3a6b"}),
            dbc.Row([
                dbc.Col(html.Div([
                    html.Div(FACES[i], style={
                        "background":lv["color"],"color":lv["tc"],
                        "borderRadius":"10px","textAlign":"center",
                        "padding":"10px 4px","fontSize":"1.4rem","lineHeight":"1",
                    }),
                    html.Div(f"{lv['range'][0]}–{lv['range'][1]}",
                             style={"textAlign":"center","fontSize":".72rem","fontWeight":"600",
                                    "color":lv["color"] if lv["color"]!="#FFFF00" else "#7a6a00","marginTop":"4px"}),
                    html.Div(lv["ar"],
                             style={"textAlign":"center","fontSize":".7rem","color":"#555"}),
                ]), width=2)
                for i,lv in enumerate(AQI_LEVELS)
            ], className="g-2"),
        ], width={"size":10,"offset":1}), className="my-4"),
    ], fluid=True)

def build_dashboard(session_id):
    if session_id not in _CACHE:
        return welcome_screen()
    data = _CACHE[session_id]
    df    = data["df"]
    ftype = data["ftype"]
    daily = daily_stats(df)
    has_pm = "PM25_mean" in daily.columns or "PM10_mean" in daily.columns

    G = lambda fig: dcc.Graph(figure=fig, config={"displayModeBar":False},
                              style={"height":"420px"})

    # ── Tab contents ────────────────────────────────────────────────
    overview = dbc.Container([
        kpi_row(df, ftype),
        dbc.Row([
            dbc.Col(G(ch_aqi_bar(daily)),    width=12, className="mb-3"),
            dbc.Col(G(ch_temp_hum(daily)),   width=6),
            dbc.Col(G(ch_pollutants(daily)), width=6),
        ]),
    ], fluid=True)

    rc = df["RiskLevel"].value_counts().reset_index()
    rc.columns=["المستوى","العدد"]
    rc["النسبة%"]=(rc["العدد"]/len(df)*100).round(1).astype(str)+"%"
    danger = df[df["RiskLevel"].isin(["حرج","طوارئ"])].shape[0]
    risk_tab = dbc.Container([
        dbc.Row([
            dbc.Col(G(ch_donut(df)),width=5),
            dbc.Col([
                dbc.Table.from_dataframe(rc,striped=True,bordered=True,hover=True,
                    style={"fontSize":".88rem","textAlign":"center"},className="mt-3"),
                dbc.Alert(f"⚠️ القراءات الخطرة (حرج+طوارئ): {danger:,}",
                          color="warning",className="mt-2 py-2"),
            ],width=7),
        ]),
    ],fluid=True)

    tabs_list = [
        dbc.Tab(overview,  label="📊 نظرة عامة",   tab_id="overview"),
    ]
    if has_pm:
        tabs_list.append(dbc.Tab(
            dbc.Container(dbc.Row(dbc.Col(G(ch_pm(daily)),width=12)),fluid=True),
            label="🌫️ PM2.5/PM10", tab_id="pm"))
    tabs_list += [
        dbc.Tab(dbc.Container(dbc.Row(dbc.Col(G(ch_hourly(df)),width=12)),fluid=True),
                label="⏰ كل ساعة", tab_id="hourly"),
        dbc.Tab(risk_tab, label="⚠️ مستوى الخطر", tab_id="risk"),
        dbc.Tab(dbc.Container([
            dbc.Row(dbc.Col(G(ch_trend(daily)),width=12)),
            dbc.Row(dbc.Col(dbc.Alert([html.B("ملاحظة: "),
                "التنبؤ مبني على الانحدار الخطي — للتوجيه فقط وليس تنبؤاً جوياً رسمياً."],
                color="info",className="mt-2"),width={"size":10,"offset":1})),
        ],fluid=True), label="📈 تنبؤ واتجاه", tab_id="forecast"),
        dbc.Tab(dbc.Container(
            dbc.Row(dbc.Col(aqi_guide_card(),width={"size":8,"offset":2}),className="my-3"),
        fluid=True), label="📋 دليل AQI", tab_id="guide"),
    ]

    return dbc.Container([
        html.Hr(style={"margin":"8px 0 14px"}),
        dbc.Tabs(tabs_list, active_tab="overview", style={"fontWeight":"600"}),
    ], fluid=True, style={"paddingBottom":"50px"})

# ══════════════════════════════════════════════
#  APP
# ══════════════════════════════════════════════
app = dash.Dash(__name__,
    external_stylesheets=[
        dbc.themes.BOOTSTRAP,
        "https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap",
    ],
    title="EnviroAI — الجامعة المستنصرية",
    suppress_callback_exceptions=True,
)
server = app.server

HEADER = dbc.Navbar(
    dbc.Container(dbc.Row([
        dbc.Col(html.Div([
            html.Span("🌍 EnviroAI", style={"color":"white","fontWeight":"700","fontSize":"1.2rem"}),
            html.Br(),
            html.Span("Air Quality & Weather Analysis System",
                      style={"color":"#a8c8ff","fontSize":".68rem"}),
        ]),width="auto"),
        dbc.Col(html.Div([
            html.Div("كلية العلوم / الجامعة المستنصرية",
                     style={"color":"white","fontWeight":"600","textAlign":"right","fontSize":".85rem"}),
            html.Div("بالتعاون مع وزارة البيئة / مديرية البيئة الحضرية — قسم مراقبة نوعية الهواء والضوضاء",
                     style={"color":"#a8c8ff","textAlign":"right","fontSize":".68rem"}),
        ]),),
    ],align="center",justify="between"),fluid=True),
    color="#1a3a6b", dark=True,
    style={"padding":"8px 0","direction":"rtl","borderBottom":"3px solid #2c5f9e"},
)

UPLOAD = dbc.Container([
    dbc.Row(dbc.Col(
        dcc.Upload(id="upload",
            children=html.Div([
                html.Div("📁",style={"fontSize":"2.5rem","lineHeight":"1"}),
                html.Div("اسحب وأفلت ملف البيانات هنا",
                         style={"fontWeight":"700","fontSize":"1rem","marginTop":"6px"}),
                html.Div("يدعم Excel (.xlsx) و CSV (تصدير AirLink)",
                         style={"color":"#555","fontSize":".82rem","marginTop":"3px"}),
                html.Div("Date/Time, AQI, Temperature, Humidity, PM2.5, PM10, CO, SMOKE ...",
                         style={"color":"#999","fontSize":".72rem","marginTop":"2px"}),
            ],style={"textAlign":"center","padding":"20px 10px"}),
            style={
                "border":"2px dashed #2c5f9e","borderRadius":"14px",
                "background":"#f0f5ff","cursor":"pointer",
            },
            multiple=False,
        ),
        width={"size":8,"offset":2}
    ),className="mt-3 mb-1"),
    dbc.Row(dbc.Col(html.Div(id="upload-msg"),width={"size":8,"offset":2})),
],fluid=True)

app.layout = html.Div([
    HEADER,
    UPLOAD,
    dcc.Store(id="session-id"),   # stores only a short UUID string
    html.Div(id="page-body", children=welcome_screen()),
], style={"fontFamily":"Cairo,Arial,sans-serif","direction":"rtl","background":"#f5f7fb","minHeight":"100vh"})

# ── Callback 1: generate session ID on page load ──────────────────────
@app.callback(Output("session-id","data"), Input("session-id","data"))
def init_session(sid):
    return sid or str(uuid.uuid4())

# ── Callback 2: upload → parse → cache → show status ─────────────────
@app.callback(
    Output("upload-msg","children"),
    Output("page-body","children"),
    Input("upload","contents"),
    State("upload","filename"),
    State("session-id","data"),
    prevent_initial_call=True,
)
def on_upload(contents, filename, session_id):
    if not contents:
        return no_update, no_update

    df, err, ftype = parse_file(contents, filename)
    if err:
        msg = dbc.Alert([html.B("❌  "), err], color="danger", className="mt-2 py-2")
        return msg, no_update

    # Store in server-side cache
    _CACHE[session_id or "default"] = {"df": df, "ftype": ftype}

    rows, days = len(df), df["DateStr"].nunique()
    icon = "📄" if ftype=="csv" else "📊"
    badge_color = "success" if ftype=="csv" else "primary"
    msg = dbc.Alert([
        html.B(f"{icon} {filename}  "),
        f"— {rows:,} قراءة  |  {days} يوم  |  نوع: {'CSV (AirLink)' if ftype=='csv' else 'Excel'}",
    ], color=badge_color, className="mt-2 py-2")

    dashboard = build_dashboard(session_id or "default")
    return msg, dashboard

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
