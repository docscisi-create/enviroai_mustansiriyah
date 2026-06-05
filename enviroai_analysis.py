"""
EnviroAI — محطة علوم الجو التعليمية
كلية العلوم  / الجامعة المستنصرية
بالتعاون مع وزارة البيئة / مديرية البيئة الحضرية
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
from matplotlib.patches import FancyBboxPatch
import matplotlib.patheffects as pe
import seaborn as sns
from scipy import stats
from scipy.ndimage import gaussian_filter1d
import warnings
import os

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
#  AQI COLOR PALETTE  (رموز ألوان AQI الرسمية)
# ─────────────────────────────────────────────
AQI_LEVELS = [
    {"label": "Good\nجيد",              "range": (0,   50),  "color": "#00E400", "face_color": "#00E400", "text_color": "#000000"},
    {"label": "Moderate\nمعتدل",         "range": (51,  100), "color": "#FFFF00", "face_color": "#FFFF00", "text_color": "#000000"},
    {"label": "Unhealthy (Sensitive)\nغير صحي للحساسين", "range": (101, 150), "color": "#FF7E00", "face_color": "#FF7E00", "text_color": "#000000"},
    {"label": "Unhealthy\nغير صحي",      "range": (151, 200), "color": "#FF0000", "face_color": "#FF0000", "text_color": "#FFFFFF"},
    {"label": "Very Unhealthy\nغير صحي جداً", "range": (201, 300), "color": "#8F3F97", "face_color": "#8F3F97", "text_color": "#FFFFFF"},
    {"label": "Hazardous\nخطير",         "range": (301, 500), "color": "#7E0023", "face_color": "#7E0023", "text_color": "#FFFFFF"},
]

def get_aqi_level(aqi_value):
    """Return AQI level dict for a given AQI value."""
    for level in AQI_LEVELS:
        lo, hi = level["range"]
        if lo <= aqi_value <= hi:
            return level
    return AQI_LEVELS[-1]

def aqi_color_array(aqi_values):
    """Return list of colors corresponding to each AQI value."""
    return [get_aqi_level(v)["color"] for v in aqi_values]


# ─────────────────────────────────────────────
#  LOAD DATA
# ─────────────────────────────────────────────
REQUIRED_COLUMNS = {"Date", "Time", "Temperature", "Humidity", "AQI",
                    "CO", "SMOKE", "RiskIndex", "RiskLevel"}

def load_data(filepath: str) -> dict:
    """
    Load an EnviroAI Excel file flexibly.

    Supports:
    - Any filename (not just EnviroAI_2026-06-05.xlsx)
    - Arabic sheet names (البيانات / الملخص / الخطرة) or fallback to first sheet
    - Auto-detects the data sheet by checking for required columns
    - Generates a 'dangerous' subset automatically if the sheet is missing
    """
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"الملف غير موجود: {filepath}")

    ext = os.path.splitext(filepath)[-1].lower()
    if ext not in (".xlsx", ".xls", ".xlsm"):
        raise ValueError(f"صيغة الملف غير مدعومة: {ext}. استخدم .xlsx أو .xls")

    sheets = pd.read_excel(filepath, sheet_name=None)
    sheet_names = list(sheets.keys())
    print(f"  ℹ  Sheets found: {sheet_names}")

    # ── Locate the main data sheet ──
    data_sheet = None
    for candidate in ["البيانات", "Data", "data", "Sheet1", "Sheet 1", sheet_names[0]]:
        if candidate in sheets:
            cols = set(sheets[candidate].columns)
            if REQUIRED_COLUMNS.issubset(cols):
                data_sheet = candidate
                break

    # fallback: pick first sheet that has all required columns
    if data_sheet is None:
        for name, sdf in sheets.items():
            if REQUIRED_COLUMNS.issubset(set(sdf.columns)):
                data_sheet = name
                break

    if data_sheet is None:
        found_cols = {name: list(sdf.columns) for name, sdf in sheets.items()}
        raise ValueError(
            f"لم يتم العثور على ورقة تحتوي الأعمدة المطلوبة:\n{REQUIRED_COLUMNS}\n"
            f"الأعمدة الموجودة في الملف:\n{found_cols}"
        )

    print(f"  ✔  Using data sheet: '{data_sheet}'")
    df = sheets[data_sheet].copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df["DateStr"]  = df["Date"].dt.strftime("%Y-%m-%d")
    df["DayLabel"] = df["Date"].dt.strftime("%-d %b")
    df["Hour"]     = pd.to_datetime(df["Time"].astype(str), errors="coerce").dt.hour

    # ── Summary sheet (optional) ──
    summary = None
    for name in ["الملخص", "Summary", "summary"]:
        if name in sheets:
            summary = sheets[name]
            break
    if summary is None:
        # auto-generate a basic summary
        summary = pd.DataFrame({
            "المقياس": ["إجمالي", "متوسط AQI", "أعلى AQI", "متوسط CO", "أعلى SMOKE"],
            "القيمة":  [len(df), round(df["AQI"].mean(), 1),
                        df["AQI"].max(), round(df["CO"].mean(), 1), df["SMOKE"].max()]
        })

    # ── Dangerous readings sheet (optional) ──
    dangerous = None
    for name in ["الخطرة", "Dangerous", "dangerous"]:
        if name in sheets:
            dangerous = sheets[name]
            break
    if dangerous is None:
        # auto-generate: حرج + طوارئ
        dangerous = df[df["RiskLevel"].isin(["حرج", "طوارئ", "Critical", "Emergency"])].copy()
        print(f"  ℹ  Auto-generated dangerous subset: {len(dangerous)} rows")

    return {"raw": df, "summary": summary, "dangerous": dangerous, "filename": os.path.basename(filepath)}


# ─────────────────────────────────────────────
#  DAILY AGGREGATION
# ─────────────────────────────────────────────
def daily_stats(df: pd.DataFrame) -> pd.DataFrame:
    agg = df.groupby("DateStr").agg(
        DayLabel=("DayLabel", "first"),
        Temp_mean=("Temperature", "mean"),
        Temp_max=("Temperature", "max"),
        Humidity_mean=("Humidity", "mean"),
        AQI_mean=("AQI", "mean"),
        AQI_max=("AQI", "max"),
        CO_mean=("CO", "mean"),
        CO_max=("CO", "max"),
        SMOKE_mean=("SMOKE", "mean"),
        SMOKE_max=("SMOKE", "max"),
        RiskIndex_mean=("RiskIndex", "mean"),
        Count=("AQI", "count"),
    ).round(2).reset_index()
    return agg


# ─────────────────────────────────────────────
#  FIGURE 1 — AQI OVERVIEW DASHBOARD
# ─────────────────────────────────────────────
def plot_aqi_dashboard(data: dict, save_path: str = "output"):
    os.makedirs(save_path, exist_ok=True)
    df = data["raw"]
    daily = daily_stats(df)

    fig = plt.figure(figsize=(18, 12), facecolor="#0d1117")
    fig.patch.set_facecolor("#0d1117")

    gs = gridspec.GridSpec(3, 3, figure=fig, hspace=0.45, wspace=0.35,
                           top=0.88, bottom=0.06, left=0.06, right=0.97)

    title_text = (
        "تقرير جودة الهواء — محطة علوم الجو التعليمية\n"
        "كلية العلوم ب / الجامعة المستنصرية    |    أبريل – مايو 2026"
    )
    fig.text(0.5, 0.94, title_text, ha="center", va="top", fontsize=14,
             color="white", fontweight="bold",
             fontproperties=_arabic_font())

    # ── Subplot 1: Daily AQI bar (coloured by level) ──
    ax1 = fig.add_subplot(gs[0, :])
    ax1.set_facecolor("#161b22")
    colors = aqi_color_array(daily["AQI_mean"].values)
    bars = ax1.bar(range(len(daily)), daily["AQI_mean"], color=colors,
                   edgecolor="#30363d", linewidth=0.5, width=0.65, zorder=3)
    # value labels on bars
    for i, (bar, val) in enumerate(zip(bars, daily["AQI_mean"])):
        lv = get_aqi_level(val)
        ax1.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 3,
                 f"{val:.0f}", ha="center", va="bottom", fontsize=9,
                 color=lv["color"], fontweight="bold")
    # threshold lines
    for thresh, lbl, col in [(100, "Moderate|معتدل", "#FFFF00"),
                              (150, "Unhealthy|غير صحي", "#FF7E00"),
                              (200, "Unhealthy|غير صحي", "#FF0000"),
                              (300, "Very Unhealthy|غير صحي جداً", "#8F3F97")]:
        ax1.axhline(thresh, color=col, linewidth=0.8, linestyle="--", alpha=0.6, zorder=2)
        ax1.text(len(daily) - 0.3, thresh + 2, lbl, ha="right", va="bottom",
                 fontsize=7, color=col, alpha=0.85)
    ax1.set_xticks(range(len(daily)))
    ax1.set_xticklabels(daily["DayLabel"], rotation=30, ha="right", fontsize=9, color="#c9d1d9")
    ax1.set_ylabel("AQI", color="#c9d1d9", fontsize=10)
    ax1.set_title("متوسط مؤشر جودة الهواء اليومي   |   Daily AQI Average",
                  color="white", fontsize=11, pad=8)
    ax1.tick_params(colors="#c9d1d9")
    ax1.spines[:].set_color("#30363d")
    ax1.set_ylim(0, max(daily["AQI_mean"].max() * 1.18, 320))
    ax1.yaxis.grid(True, color="#30363d", linewidth=0.5, zorder=0)
    _add_aqi_legend_bar(ax1)

    # ── Subplot 2: Temperature line ──
    ax2 = fig.add_subplot(gs[1, 0])
    _style_ax(ax2, "#161b22")
    ax2.plot(range(len(daily)), daily["Temp_mean"], color="#FF7E00",
             marker="o", markersize=5, linewidth=2, zorder=3, label="متوسط")
    ax2.fill_between(range(len(daily)), daily["Temp_mean"], alpha=0.15, color="#FF7E00")
    ax2.plot(range(len(daily)), daily["Temp_max"], color="#FF0000",
             marker="^", markersize=4, linewidth=1.2, linestyle="--", alpha=0.7, label="أعلى")
    ax2.set_xticks(range(len(daily)))
    ax2.set_xticklabels(daily["DayLabel"], rotation=45, ha="right", fontsize=7, color="#c9d1d9")
    ax2.set_title("درجة الحرارة (°م)\nTemperature (°C)", color="white", fontsize=9)
    ax2.set_ylabel("°C", color="#c9d1d9", fontsize=8)
    ax2.legend(fontsize=7, facecolor="#161b22", edgecolor="#30363d",
               labelcolor="#c9d1d9", loc="upper left")
    ax2.yaxis.grid(True, color="#30363d", linewidth=0.4)

    # ── Subplot 3: CO bar ──
    ax3 = fig.add_subplot(gs[1, 1])
    _style_ax(ax3, "#161b22")
    co_colors = ["#FF0000" if v > 200 else "#FF7E00" if v > 100 else "#8F3F97"
                 for v in daily["CO_mean"]]
    ax3.bar(range(len(daily)), daily["CO_mean"], color=co_colors,
            edgecolor="#30363d", linewidth=0.4, width=0.6)
    ax3.set_xticks(range(len(daily)))
    ax3.set_xticklabels(daily["DayLabel"], rotation=45, ha="right", fontsize=7, color="#c9d1d9")
    ax3.set_title("أول أكسيد الكربون (CO)\nCarbon Monoxide", color="white", fontsize=9)
    ax3.set_ylabel("ppm", color="#c9d1d9", fontsize=8)
    ax3.yaxis.grid(True, color="#30363d", linewidth=0.4)

    # ── Subplot 4: SMOKE line ──
    ax4 = fig.add_subplot(gs[1, 2])
    _style_ax(ax4, "#161b22")
    ax4.plot(range(len(daily)), daily["SMOKE_mean"], color="#c9d1d9",
             marker="s", markersize=4, linewidth=1.8, zorder=3)
    ax4.fill_between(range(len(daily)), daily["SMOKE_mean"], alpha=0.2, color="#7E0023")
    ax4.set_xticks(range(len(daily)))
    ax4.set_xticklabels(daily["DayLabel"], rotation=45, ha="right", fontsize=7, color="#c9d1d9")
    ax4.set_title("مؤشر الدخان (SMOKE)\nSmoke Index", color="white", fontsize=9)
    ax4.set_ylabel("ppm", color="#c9d1d9", fontsize=8)
    ax4.yaxis.grid(True, color="#30363d", linewidth=0.4)

    # ── Subplot 5: Risk distribution donut ──
    ax5 = fig.add_subplot(gs[2, 0])
    _style_ax(ax5, "#161b22")
    risk_counts = data["raw"]["RiskLevel"].value_counts()
    ordered_keys = ["آمن", "متوسط", "تحذير", "حرج", "طوارئ"]
    risk_colors_map = {"آمن": "#00E400", "متوسط": "#FFFF00", "تحذير": "#FF7E00",
                       "حرج": "#FF0000", "طوارئ": "#8F3F97"}
    sizes = [risk_counts.get(k, 0) for k in ordered_keys]
    colors_pie = [risk_colors_map[k] for k in ordered_keys]
    wedges, texts, autotexts = ax5.pie(
        sizes, labels=None, colors=colors_pie, autopct="%1.1f%%",
        startangle=90, pctdistance=0.75,
        wedgeprops={"linewidth": 1.5, "edgecolor": "#0d1117"},
    )
    for at in autotexts:
        at.set_fontsize(7)
        at.set_color("white")
    circle = plt.Circle((0, 0), 0.5, color="#161b22")
    ax5.add_patch(circle)
    ax5.text(0, 0, f"4,910\nقراءة", ha="center", va="center",
             fontsize=8, color="white", fontweight="bold")
    ax5.set_title("توزيع مستويات الخطر\nRisk Level Distribution",
                  color="white", fontsize=9)
    legend_patches = [mpatches.Patch(color=risk_colors_map[k], label=k) for k in ordered_keys]
    ax5.legend(handles=legend_patches, fontsize=7, facecolor="#161b22",
               edgecolor="#30363d", labelcolor="#c9d1d9",
               loc="lower center", ncol=2, bbox_to_anchor=(0.5, -0.15))

    # ── Subplot 6: Hourly AQI heatmap pattern ──
    ax6 = fig.add_subplot(gs[2, 1])
    _style_ax(ax6, "#161b22")
    df_h = data["raw"].dropna(subset=["Hour"])
    hourly_avg = df_h.groupby("Hour")["AQI"].mean().reindex(range(24), fill_value=np.nan)
    hours = hourly_avg.index.values
    vals  = hourly_avg.values
    valid = ~np.isnan(vals)
    h_colors = [get_aqi_level(v)["color"] if not np.isnan(v) else "#30363d" for v in vals]
    ax6.bar(hours, np.where(valid, vals, 0), color=h_colors,
            edgecolor="#0d1117", linewidth=0.4, width=0.85)
    ax6.set_xlim(-0.5, 23.5)
    ax6.set_xticks([0, 3, 6, 9, 12, 15, 18, 21])
    ax6.set_xticklabels(["12AM","3AM","6AM","9AM","12PM","3PM","6PM","9PM"],
                        fontsize=7, color="#c9d1d9")
    ax6.set_title("متوسط AQI لكل ساعة\nHourly AQI Pattern", color="white", fontsize=9)
    ax6.set_ylabel("AQI", color="#c9d1d9", fontsize=8)
    ax6.yaxis.grid(True, color="#30363d", linewidth=0.4)

    # ── Subplot 7: Trend + linear regression ──
    ax7 = fig.add_subplot(gs[2, 2])
    _style_ax(ax7, "#161b22")
    x = np.arange(len(daily))
    y = daily["AQI_mean"].values
    slope, intercept, r_val, p_val, _ = stats.linregress(x, y)
    trend = slope * x + intercept
    aqi_smooth = gaussian_filter1d(y, sigma=1)
    ax7.scatter(x, y, c=aqi_color_array(y), s=50, zorder=4, edgecolors="#30363d", linewidths=0.5)
    ax7.plot(x, aqi_smooth, color="#58a6ff", linewidth=1.5, zorder=3, label="AQI (smooth)")
    ax7.plot(x, trend, color="#FF7E00", linewidth=2, linestyle="--", zorder=5,
             label=f"Trend (slope={slope:.1f}/day)")
    ax7.fill_between(x, trend, alpha=0.1, color="#FF7E00")
    ax7.set_xticks(x)
    ax7.set_xticklabels(daily["DayLabel"], rotation=45, ha="right", fontsize=7, color="#c9d1d9")
    ax7.set_title(f"اتجاه AQI  |  R²={r_val**2:.2f}  p={p_val:.3f}\nLinear Trend Analysis",
                  color="white", fontsize=9)
    ax7.legend(fontsize=7, facecolor="#161b22", edgecolor="#30363d",
               labelcolor="#c9d1d9", loc="upper left")
    ax7.yaxis.grid(True, color="#30363d", linewidth=0.4)

    plt.savefig(f"{save_path}/01_aqi_dashboard.png", dpi=150,
                bbox_inches="tight", facecolor="#0d1117")
    plt.close()
    print(f"  ✔  Saved: {save_path}/01_aqi_dashboard.png")


# ─────────────────────────────────────────────
#  FIGURE 2 — AQI OFFICIAL SCALE LEGEND
# ─────────────────────────────────────────────
def plot_aqi_scale_reference(save_path: str = "output"):
    """Reproduce the official AQI reference card (دليل مؤشر جودة الهواء)."""
    os.makedirs(save_path, exist_ok=True)

    fig, ax = plt.subplots(figsize=(10, 7), facecolor="#1c1c1c")
    ax.set_facecolor("#1c1c1c")
    ax.set_xlim(0, 10)
    ax.set_ylim(0, len(AQI_LEVELS) + 1.2)
    ax.axis("off")

    fig.text(0.5, 0.96, "دليل مؤشر جودة الهواء  |  AQI Reference Guide",
             ha="center", va="top", fontsize=14, color="white",
             fontweight="bold", fontproperties=_arabic_font())

    descriptions = [
        "جودة الهواء مرضية وتشكل خطرًا قليلاً أو منعدماً. يوصى بتهوية منزلك.\n"
        "Air quality is satisfactory and poses little or no risk.",

        "يجب على الأفراد الحساسين تجنب الأنشطة الخارجية المكثفة.\n"
        "Sensitive groups should limit prolonged outdoor exertion.",

        "الجمهور والأفراد الحساسون معرضون لخطر الإصابة بمشاكل تنفسية.\n"
        "Members of sensitive groups may experience health effects.",

        "يزداد احتمال حدوث آثار ضارة. يجب تجنب الأنشطة الخارجية المكثفة.\n"
        "Everyone may begin to experience health effects.",

        "سيتأثر عموم الجمهور بشكل ملحوظ. ينبغي البقاء في المنازل.\n"
        "Health alert: everyone may experience more serious health effects.",

        "قد تؤدي إلى أمراض خطيرة. يجب على الجميع تجنب الرياضة والبقاء في الداخل.\n"
        "Health warnings of emergency conditions. Everyone is at risk.",
    ]

    for i, level in enumerate(AQI_LEVELS):
        row = len(AQI_LEVELS) - i - 1
        y = row + 0.15

        # colored box
        box = FancyBboxPatch((0.1, y), 9.8, 0.78,
                             boxstyle="round,pad=0.05",
                             facecolor=level["color"],
                             edgecolor="white", linewidth=0.8, zorder=2)
        ax.add_patch(box)

        # range + label
        lo, hi = level["range"]
        ax.text(0.35, y + 0.52, f"{lo} – {hi}",
                ha="left", va="center", fontsize=11, fontweight="bold",
                color=level["text_color"])
        short_label = level["label"].split("\n")[-1]  # Arabic part
        ax.text(0.35, y + 0.24, short_label,
                ha="left", va="center", fontsize=9,
                color=level["text_color"], style="italic")

        # description text
        desc = descriptions[i]
        ax.text(5.0, y + 0.39, desc,
                ha="center", va="center", fontsize=7,
                color=level["text_color"], wrap=True,
                multialignment="center",
                fontproperties=_arabic_font())

    plt.savefig(f"{save_path}/02_aqi_reference_card.png", dpi=150,
                bbox_inches="tight", facecolor="#1c1c1c")
    plt.close()
    print(f"  ✔  Saved: {save_path}/02_aqi_reference_card.png")


# ─────────────────────────────────────────────
#  FIGURE 3 — DETAILED POLLUTANT ANALYSIS
# ─────────────────────────────────────────────
def plot_pollutant_analysis(data: dict, save_path: str = "output"):
    os.makedirs(save_path, exist_ok=True)
    df = data["raw"]
    daily = daily_stats(df)

    fig, axes = plt.subplots(2, 2, figsize=(16, 10), facecolor="#0d1117")
    fig.patch.set_facecolor("#0d1117")
    fig.suptitle(
        "تحليل الملوثات — محطة الجامعة المستنصرية  |  Pollutant Analysis",
        fontsize=13, color="white", fontweight="bold", y=0.98
    )

    x = np.arange(len(daily))
    labels = daily["DayLabel"].values

    # ── Panel A: AQI + Risk Index overlay ──
    ax = axes[0, 0]
    _style_ax(ax, "#161b22")
    colors_aqi = aqi_color_array(daily["AQI_mean"].values)
    ax.bar(x - 0.2, daily["AQI_mean"], width=0.38, color=colors_aqi,
           edgecolor="#30363d", linewidth=0.4, label="AQI", zorder=3)
    ax2r = ax.twinx()
    ax2r.plot(x, daily["RiskIndex_mean"], color="#FF7E00", linewidth=2,
              marker="D", markersize=5, zorder=4, label="Risk Index")
    ax2r.set_ylabel("Risk Index", color="#FF7E00", fontsize=8)
    ax2r.tick_params(colors="#FF7E00")
    ax2r.spines[:].set_color("#30363d")
    ax.set_xticks(x); ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=7, color="#c9d1d9")
    ax.set_title("AQI vs مؤشر الخطر", color="white", fontsize=10)
    ax.set_ylabel("AQI", color="#c9d1d9", fontsize=8)
    ax.yaxis.grid(True, color="#30363d", linewidth=0.4)
    _add_aqi_legend_bar(ax, compact=True)

    # ── Panel B: CO & SMOKE dual bar ──
    ax = axes[0, 1]
    _style_ax(ax, "#161b22")
    ax.bar(x - 0.2, daily["CO_mean"], width=0.38, color="#8F3F97",
           edgecolor="#30363d", linewidth=0.4, label="CO (ppm)", zorder=3)
    ax.bar(x + 0.2, daily["SMOKE_mean"], width=0.38, color="#c9d1d9",
           edgecolor="#30363d", linewidth=0.4, label="SMOKE", zorder=3)
    ax.set_xticks(x); ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=7, color="#c9d1d9")
    ax.set_title("CO و الدخان اليومي  |  CO & Smoke Daily", color="white", fontsize=10)
    ax.set_ylabel("ppm", color="#c9d1d9", fontsize=8)
    ax.yaxis.grid(True, color="#30363d", linewidth=0.4)
    ax.legend(fontsize=8, facecolor="#161b22", edgecolor="#30363d",
              labelcolor="#c9d1d9", loc="upper left")

    # ── Panel C: Humidity vs Temperature scatter coloured by AQI ──
    ax = axes[1, 0]
    _style_ax(ax, "#161b22")
    sc = ax.scatter(df["Temperature"], df["Humidity"],
                    c=df["AQI"], cmap="RdYlGn_r",
                    s=8, alpha=0.5, edgecolors="none",
                    vmin=0, vmax=300)
    cbar = plt.colorbar(sc, ax=ax, pad=0.02)
    cbar.set_label("AQI", color="#c9d1d9", fontsize=8)
    cbar.ax.yaxis.set_tick_params(color="#c9d1d9")
    plt.setp(cbar.ax.yaxis.get_ticklabels(), color="#c9d1d9", fontsize=7)
    ax.set_xlabel("درجة الحرارة (°م)  |  Temperature (°C)",
                  color="#c9d1d9", fontsize=8)
    ax.set_ylabel("الرطوبة (%)  |  Humidity (%)", color="#c9d1d9", fontsize=8)
    ax.set_title("الحرارة / الرطوبة ملوّنة بـ AQI\nTemp vs Humidity (coloured by AQI)",
                 color="white", fontsize=9)
    ax.yaxis.grid(True, color="#30363d", linewidth=0.4)
    ax.xaxis.grid(True, color="#30363d", linewidth=0.4)

    # ── Panel D: Dangerous readings timeline ──
    ax = axes[1, 1]
    _style_ax(ax, "#161b22")
    danger_df = data["dangerous"].copy()
    danger_df["Date"] = pd.to_datetime(danger_df["Date"])
    danger_df["DateStr"] = danger_df["Date"].dt.strftime("%Y-%m-%d")
    danger_daily = danger_df.groupby("DateStr").agg(
        AQI_max=("AQI", "max"),
        Count=("AQI", "count")
    ).reset_index()
    dd_x = range(len(danger_daily))
    dd_colors = aqi_color_array(danger_daily["AQI_max"].values)
    bars = ax.bar(dd_x, danger_daily["Count"], color=dd_colors,
                  edgecolor="#30363d", linewidth=0.5, width=0.6, zorder=3)
    for bar, count in zip(bars, danger_daily["Count"]):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1,
                str(count), ha="center", va="bottom", fontsize=8, color="#c9d1d9")
    ax.set_xticks(dd_x)
    ax.set_xticklabels(danger_daily["DateStr"], rotation=45, ha="right",
                       fontsize=7, color="#c9d1d9")
    ax.set_title("عدد القراءات الخطرة يومياً (حرج + طوارئ)\nDangerous Readings per Day",
                 color="white", fontsize=9)
    ax.set_ylabel("عدد القراءات  |  Count", color="#c9d1d9", fontsize=8)
    ax.yaxis.grid(True, color="#30363d", linewidth=0.4)

    plt.tight_layout(rect=[0, 0, 1, 0.97])
    plt.savefig(f"{save_path}/03_pollutant_analysis.png", dpi=150,
                bbox_inches="tight", facecolor="#0d1117")
    plt.close()
    print(f"  ✔  Saved: {save_path}/03_pollutant_analysis.png")


# ─────────────────────────────────────────────
#  FIGURE 4 — FORECAST & STATISTICS REPORT
# ─────────────────────────────────────────────
def plot_forecast_report(data: dict, save_path: str = "output"):
    os.makedirs(save_path, exist_ok=True)
    df = data["raw"]
    daily = daily_stats(df)
    aqi = daily["AQI_mean"].values
    x = np.arange(len(aqi))

    # linear regression
    slope, intercept, r_val, p_val, se = stats.linregress(x, aqi)
    trend = slope * x + intercept

    # forecast 7 days ahead
    x_future = np.arange(len(aqi), len(aqi) + 7)
    y_future = slope * x_future + intercept
    ci = 1.96 * se * np.sqrt(1 + 1 / len(x) + (x_future - x.mean()) ** 2 / ((x - x.mean()) ** 2).sum())

    fig, axes = plt.subplots(1, 2, figsize=(16, 7), facecolor="#0d1117")
    fig.patch.set_facecolor("#0d1117")
    fig.suptitle(
        "تحليل الاتجاه والتنبؤ — AQI  |  Trend Analysis & Forecast",
        fontsize=13, color="white", fontweight="bold", y=0.99
    )

    # ── Left: full trend view ──
    ax = axes[0]
    _style_ax(ax, "#161b22")
    scatter_colors = aqi_color_array(aqi)
    ax.scatter(x, aqi, c=scatter_colors, s=80, zorder=5,
               edgecolors="white", linewidths=0.4)
    smooth = gaussian_filter1d(aqi, sigma=1.2)
    ax.plot(x, smooth, color="#58a6ff", linewidth=2, label="AQI (smoothed)", zorder=4)
    ax.plot(x, trend, color="#FF7E00", linewidth=2, linestyle="--", label=f"Trend +{slope:.1f}/day", zorder=6)
    # AQI threshold bands
    band_defs = [(0, 50, "#00E400", 0.06), (50, 100, "#FFFF00", 0.06),
                 (100, 150, "#FF7E00", 0.07), (150, 200, "#FF0000", 0.07),
                 (200, 300, "#8F3F97", 0.07), (300, 500, "#7E0023", 0.07)]
    for lo, hi, col, alpha in band_defs:
        ax.axhspan(lo, hi, color=col, alpha=alpha, zorder=0)
    ax.set_xticks(x)
    ax.set_xticklabels(daily["DayLabel"], rotation=45, ha="right", fontsize=8, color="#c9d1d9")
    ax.set_ylabel("AQI", color="#c9d1d9", fontsize=9)
    ax.set_title(f"منحنى AQI مع خط الاتجاه  |  R²={r_val**2:.3f}", color="white", fontsize=10)
    ax.legend(fontsize=8, facecolor="#161b22", edgecolor="#30363d", labelcolor="#c9d1d9")
    ax.yaxis.grid(True, color="#30363d", linewidth=0.5)
    _add_stats_box(ax, aqi, slope, r_val, p_val)

    # ── Right: 7-day forecast ──
    ax2 = axes[1]
    _style_ax(ax2, "#161b22")
    ax2.scatter(x, aqi, c=scatter_colors, s=60, zorder=4, edgecolors="white", linewidths=0.4)
    ax2.plot(x, trend, color="#FF7E00", linewidth=1.5, linestyle="--", zorder=3)
    forecast_colors = aqi_color_array(np.clip(y_future, 0, 500))
    ax2.scatter(x_future, y_future, c=forecast_colors, s=90, marker="D",
                edgecolors="white", linewidths=0.8, zorder=5)
    ax2.plot(x_future, y_future, color="#8F3F97", linewidth=2,
             linestyle="-", label="7-day forecast", zorder=4)
    ax2.fill_between(x_future, y_future - ci, y_future + ci,
                     color="#8F3F97", alpha=0.2, label="95% CI")
    for j, (xi, yi) in enumerate(zip(x_future, y_future)):
        lv = get_aqi_level(max(0, yi))
        ax2.annotate(f"Day+{j+1}\n{yi:.0f}", (xi, yi),
                     textcoords="offset points", xytext=(0, 10),
                     ha="center", fontsize=7, color=lv["color"])
    ax2.axvline(len(aqi) - 0.5, color="#c9d1d9", linewidth=1, linestyle=":", alpha=0.6)
    ax2.text(len(aqi) - 0.3, ax2.get_ylim()[0] + 10, "⟵ تاريخي",
             fontsize=8, color="#c9d1d9", alpha=0.7)
    ax2.text(len(aqi) + 0.1, ax2.get_ylim()[0] + 10, "تنبؤ ⟶",
             fontsize=8, color="#8F3F97")
    for lo, hi, col, alpha in band_defs:
        ax2.axhspan(lo, hi, color=col, alpha=alpha, zorder=0)
    ax2.set_title("التنبؤ بـ AQI — 7 أيام قادمة\n7-Day AQI Forecast",
                  color="white", fontsize=10)
    ax2.set_ylabel("AQI", color="#c9d1d9", fontsize=9)
    ax2.legend(fontsize=8, facecolor="#161b22", edgecolor="#30363d", labelcolor="#c9d1d9")
    ax2.yaxis.grid(True, color="#30363d", linewidth=0.5)

    plt.tight_layout()
    plt.savefig(f"{save_path}/04_forecast_report.png", dpi=150,
                bbox_inches="tight", facecolor="#0d1117")
    plt.close()
    print(f"  ✔  Saved: {save_path}/04_forecast_report.png")


# ─────────────────────────────────────────────
#  FIGURE 5 — OFFICIAL REPORT STYLE (mimics Image 2)
# ─────────────────────────────────────────────
def plot_official_report(data: dict, save_path: str = "output"):
    os.makedirs(save_path, exist_ok=True)
    df = data["raw"]
    daily = daily_stats(df)

    fig = plt.figure(figsize=(18, 10), facecolor="white")
    fig.patch.set_facecolor("white")

    # Header bar
    header_ax = fig.add_axes([0, 0.88, 1, 0.12])
    header_ax.set_facecolor("#003366")
    header_ax.axis("off")
    header_ax.text(
        0.5, 0.6,
        "تقرير جودة الهواء من محطة المستنصرية — قسم علوم الجو / كلية العلوم ب",
        ha="center", va="center", fontsize=13, color="white",
        fontweight="bold", fontproperties=_arabic_font()
    )
    header_ax.text(
        0.5, 0.2,
        "بالتعاون مع وزارة البيئة / مديرية البيئة الحضرية  |  قسم مراقبة نوعية الهواء والضوضاء",
        ha="center", va="center", fontsize=9, color="#aad4ff",
        fontproperties=_arabic_font()
    )

    gs = gridspec.GridSpec(2, 2, figure=fig, top=0.86, bottom=0.05,
                           left=0.05, right=0.97, hspace=0.4, wspace=0.3)

    # AQI bar coloured
    ax1 = fig.add_subplot(gs[0, :])
    x = np.arange(len(daily))
    colors = aqi_color_array(daily["AQI_mean"].values)
    bars = ax1.bar(x, daily["AQI_mean"], color=colors,
                   edgecolor="#cccccc", linewidth=0.6, width=0.65)
    for bar, val in zip(bars, daily["AQI_mean"]):
        ax1.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1.5,
                 f"{val:.0f}", ha="center", va="bottom", fontsize=9, fontweight="bold")
    ax1.set_xticks(x)
    ax1.set_xticklabels(daily["DayLabel"], rotation=30, ha="right", fontsize=9)
    ax1.set_ylabel("AQI", fontsize=10)
    ax1.set_title("مؤشر جودة الهواء اليومي  |  Daily AQI Index", fontsize=11,
                  fontweight="bold")
    ax1.yaxis.grid(True, color="#eeeeee", linewidth=0.5)
    ax1.set_facecolor("#f8f9fa")
    ax1.spines[:].set_color("#cccccc")
    _add_aqi_scale_bar(ax1)

    # Temperature & Humidity
    ax2 = fig.add_subplot(gs[1, 0])
    ax2.set_facecolor("#f8f9fa")
    ax2.spines[:].set_color("#cccccc")
    ax2.plot(x, daily["Temp_mean"], color="#e74c3c", marker="o", markersize=5,
             linewidth=2, label="درجة الحرارة (°م)")
    ax2b = ax2.twinx()
    ax2b.bar(x, daily["Humidity_mean"], alpha=0.3, color="#3498db",
             width=0.5, label="الرطوبة (%)")
    ax2b.set_ylabel("الرطوبة (%)", fontsize=8, color="#3498db")
    ax2b.tick_params(colors="#3498db")
    ax2b.spines[:].set_color("#cccccc")
    ax2.set_xticks(x)
    ax2.set_xticklabels(daily["DayLabel"], rotation=45, ha="right", fontsize=7)
    ax2.set_title("الحرارة والرطوبة  |  Temp & Humidity", fontsize=9, fontweight="bold")
    ax2.set_ylabel("°C", fontsize=8, color="#e74c3c")
    ax2.yaxis.grid(True, color="#eeeeee", linewidth=0.4)

    # Risk level bar
    ax3 = fig.add_subplot(gs[1, 1])
    ax3.set_facecolor("#f8f9fa")
    ax3.spines[:].set_color("#cccccc")
    risk_counts = df["RiskLevel"].value_counts()
    keys = ["آمن", "متوسط", "تحذير", "حرج", "طوارئ"]
    risk_colors_map = {"آمن": "#00E400", "متوسط": "#FFFF00",
                       "تحذير": "#FF7E00", "حرج": "#FF0000", "طوارئ": "#8F3F97"}
    vals = [risk_counts.get(k, 0) for k in keys]
    bars3 = ax3.barh(keys, vals, color=[risk_colors_map[k] for k in keys],
                     edgecolor="#cccccc", linewidth=0.5)
    for bar, val in zip(bars3, vals):
        ax3.text(bar.get_width() + 20, bar.get_y() + bar.get_height() / 2,
                 str(val), va="center", fontsize=9)
    ax3.set_title("توزيع مستويات الخطر  |  Risk Levels", fontsize=9, fontweight="bold")
    ax3.set_xlabel("عدد القراءات", fontsize=8)
    ax3.xaxis.grid(True, color="#eeeeee", linewidth=0.4)

    plt.savefig(f"{save_path}/05_official_report.png", dpi=150,
                bbox_inches="tight", facecolor="white")
    plt.close()
    print(f"  ✔  Saved: {save_path}/05_official_report.png")


# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────
def _arabic_font():
    """Return font properties that support Arabic (fallback gracefully)."""
    try:
        from matplotlib.font_manager import FontProperties
        return FontProperties(family="DejaVu Sans")
    except Exception:
        return None

def _style_ax(ax, bg_color):
    ax.set_facecolor(bg_color)
    ax.tick_params(colors="#c9d1d9", labelsize=8)
    ax.spines[:].set_color("#30363d")

def _add_aqi_legend_bar(ax, compact=False):
    """Add a small AQI colour legend below the axes."""
    from mpl_toolkits.axes_grid1.inset_locator import inset_axes
    legend_patches = []
    for level in AQI_LEVELS:
        lo, hi = level["range"]
        lbl = f"{lo}-{hi}" if compact else f"{lo}-{hi}\n{level['label'].split(chr(10))[-1]}"
        legend_patches.append(mpatches.Patch(color=level["color"], label=lbl))
    ax.legend(handles=legend_patches, loc="upper left",
              fontsize=6 if compact else 7,
              facecolor="#0d1117" if not compact else "#161b22",
              edgecolor="#30363d", labelcolor="#c9d1d9",
              ncol=len(AQI_LEVELS), bbox_to_anchor=(0, -0.22 if not compact else -0.28))

def _add_aqi_scale_bar(ax):
    """Add AQI colour scale bar below axis for official-style report."""
    scale_data = [(f"{l['range'][0]}-{l['range'][1]}", l["color"]) for l in AQI_LEVELS]
    for i, (lbl, col) in enumerate(scale_data):
        ax.annotate("", xy=(0, 0), xytext=(0, 0))  # dummy
    patches = [mpatches.Patch(color=col, label=lbl) for lbl, col in scale_data]
    ax.legend(handles=patches, loc="upper right", fontsize=7,
              ncol=len(scale_data), framealpha=0.9,
              bbox_to_anchor=(1, -0.18))

def _add_stats_box(ax, aqi, slope, r_val, p_val):
    stats_text = (
        f"n = {len(aqi)}  |  μ = {np.mean(aqi):.1f}\n"
        f"σ = {np.std(aqi):.1f}  |  max = {np.max(aqi):.0f}\n"
        f"Slope = +{slope:.1f}/day\n"
        f"R² = {r_val**2:.3f}  |  p = {p_val:.3f}"
    )
    ax.text(0.02, 0.97, stats_text, transform=ax.transAxes,
            fontsize=8, va="top", ha="left", color="#c9d1d9",
            bbox=dict(boxstyle="round,pad=0.4", facecolor="#161b22",
                      edgecolor="#30363d", alpha=0.9))


# ─────────────────────────────────────────────
#  GENERATE SUMMARY CSV
# ─────────────────────────────────────────────
def export_summary_csv(data: dict, save_path: str = "output"):
    os.makedirs(save_path, exist_ok=True)
    daily = daily_stats(data["raw"])

    # add AQI level column
    daily["AQI_Level_EN"] = daily["AQI_mean"].apply(
        lambda v: get_aqi_level(v)["label"].split("\n")[0])
    daily["AQI_Level_AR"] = daily["AQI_mean"].apply(
        lambda v: get_aqi_level(v)["label"].split("\n")[-1])

    # linear regression for trend
    x = np.arange(len(daily))
    slope, intercept, r_val, *_ = stats.linregress(x, daily["AQI_mean"].values)
    daily["AQI_Trend"] = (slope * x + intercept).round(2)

    daily.to_csv(f"{save_path}/daily_summary.csv", index=False, encoding="utf-8-sig")
    print(f"  ✔  Saved: {save_path}/daily_summary.csv")
    return daily


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────
# ─────────────────────────────────────────────
#  BATCH MODE — analyse multiple files at once
# ─────────────────────────────────────────────
def analyse_file(excel_path: str, output_dir: str) -> None:
    """Run full analysis pipeline on a single Excel file."""
    fname = os.path.basename(excel_path)
    stem  = os.path.splitext(fname)[0]
    file_out = os.path.join(output_dir, stem)
    os.makedirs(file_out, exist_ok=True)

    print(f"\n{'─'*60}")
    print(f"  ► File : {fname}")
    print(f"  ► Output: {file_out}/")
    print(f"{'─'*60}")

    data  = load_data(excel_path)
    total = len(data["raw"])
    days  = data["raw"]["DateStr"].nunique()
    print(f"  ✔  Loaded {total:,} readings across {days} days")

    print("  ► Generating figures...")
    plot_aqi_scale_reference(file_out)
    plot_aqi_dashboard(data, file_out)
    plot_pollutant_analysis(data, file_out)
    plot_forecast_report(data, file_out)
    plot_official_report(data, file_out)

    print("  ► Exporting summary CSV...")
    export_summary_csv(data, file_out)

    print(f"\n  ► Quick stats for {fname}:")
    print(f"     Readings   : {total:,}")
    print(f"     AQI mean   : {data['raw']['AQI'].mean():.1f}")
    print(f"     AQI max    : {data['raw']['AQI'].max():.0f}")
    print(f"     Temp mean  : {data['raw']['Temperature'].mean():.1f} °C")
    dangerous_days = pd.to_datetime(data["dangerous"]["Date"], errors="coerce").dt.date.nunique()
    print(f"     Danger days: {dangerous_days}")
    print(f"  ✔  Done → {file_out}/")


def main(inputs: list, output_dir: str = "output") -> None:
    """
    Main entry point — supports one or many Excel files.

    Parameters
    ----------
    inputs      : list of file paths (glob patterns already expanded by argparse/glob)
    output_dir  : root output folder; each file gets its own sub-folder inside
    """
    banner = [
        "╔══════════════════════════════════════════════════════════╗",
        "║   EnviroAI — محطة علوم الجو — الجامعة المستنصرية        ║",
        "║   Air Quality & Weather Analysis System  |  v2.0  |  2026║",
        "║   كلية العلوم — قسم مراقبة نوعية الهواء والضوضاء       ║",
        "╚══════════════════════════════════════════════════════════╝",
    ]
    print("\n" + "\n".join(banner) + "\n")

    if not inputs:
        print("  ✖  No input files provided. Use --input to specify one or more Excel files.")
        print("  Example: python enviroai_analysis.py --input data/*.xlsx\n")
        return

    valid   = [f for f in inputs if os.path.isfile(f)]
    invalid = [f for f in inputs if not os.path.isfile(f)]

    if invalid:
        print(f"  ⚠  Skipping missing files: {invalid}")
    if not valid:
        print("  ✖  No valid files to process.")
        return

    print(f"  ℹ  Files to process ({len(valid)}):")
    for f in valid:
        size_kb = os.path.getsize(f) // 1024
        print(f"     • {os.path.basename(f)}  ({size_kb} KB)")

    errors = []
    for i, fpath in enumerate(valid, 1):
        print(f"\n[{i}/{len(valid)}] Processing: {os.path.basename(fpath)}")
        try:
            analyse_file(fpath, output_dir)
        except Exception as exc:
            print(f"  ✖  ERROR in {os.path.basename(fpath)}: {exc}")
            errors.append((fpath, str(exc)))

    print("\n" + "═" * 60)
    print(f"  ✔  Finished: {len(valid) - len(errors)}/{len(valid)} files succeeded")
    if errors:
        print(f"  ✖  Failed files:")
        for fpath, msg in errors:
            print(f"     • {os.path.basename(fpath)}: {msg}")
    print(f"  ℹ  All output saved under: {os.path.abspath(output_dir)}/")
    print("═" * 60 + "\n")


# ─────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────
if __name__ == "__main__":
    import argparse
    import glob

    parser = argparse.ArgumentParser(
        description="EnviroAI — Air Quality Analysis — Mustansiriyah University",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Single file:
  python enviroai_analysis.py --input EnviroAI_2026-06-05.xlsx

  # Multiple specific files:
  python enviroai_analysis.py --input file1.xlsx file2.xlsx file3.xlsx

  # All xlsx files in a folder (glob):
  python enviroai_analysis.py --input data/*.xlsx

  # Interactive prompt (no --input):
  python enviroai_analysis.py
        """
    )
    parser.add_argument(
        "--input", "-i",
        nargs="*",
        default=None,
        help="One or more Excel file paths (.xlsx). Supports glob patterns like data/*.xlsx"
    )
    parser.add_argument(
        "--output", "-o",
        default="output",
        help="Root output directory (default: output/). Each file gets a sub-folder."
    )
    args = parser.parse_args()

    # ── Resolve input files ──
    if args.input:
        # expand any glob patterns (Windows doesn't auto-expand *)
        resolved = []
        for pattern in args.input:
            expanded = glob.glob(pattern)
            if expanded:
                resolved.extend(expanded)
            else:
                resolved.append(pattern)   # pass as-is; error handled in main()
        input_files = sorted(set(resolved))
    else:
        # interactive mode: prompt user to enter path(s)
        print("\n  EnviroAI — لم يتم تحديد ملفات عبر --input")
        print("  أدخل مسار ملف Excel أو أكثر (اضغط Enter بعد كل مسار، أدخل سطراً فارغاً للبدء):\n")
        input_files = []
        while True:
            try:
                path = input("  > ملف Excel: ").strip().strip('"').strip("'")
            except (EOFError, KeyboardInterrupt):
                break
            if not path:
                break
            # expand glob if entered
            expanded = glob.glob(path)
            if expanded:
                input_files.extend(expanded)
                print(f"    ✔ تم إضافة {len(expanded)} ملف/ملفات")
            else:
                input_files.append(path)
                print(f"    ✔ تم إضافة: {path}")
        if not input_files:
            print("  ✖  لم يتم إدخال أي ملفات. خروج.\n")
            raise SystemExit(0)

    main(inputs=input_files, output_dir=args.output)
