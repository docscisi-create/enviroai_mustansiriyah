# 🌍 EnviroAI — محطة علوم الجو التعليمية
### كلية العلوم ب / الجامعة المستنصرية

> **Air Quality & Weather Analysis System**  
> بالتعاون مع وزارة البيئة / مديرية البيئة الحضرية — قسم مراقبة نوعية الهواء والضوضاء

---

## 📌 نظرة عامة | Overview

نظام Python كامل لتحليل وتصوير بيانات جودة الهواء والطقس المُجمَّعة من محطة الرصد التعليمية في الجامعة المستنصرية — بغداد / العراق.

يعتمد النظام على **رموز AQI الرسمية** (المقياس الدولي لجودة الهواء) ويُصدر:
- 5 لوحات تحليلية احترافية (PNG)
- ملف CSV يومي مُعالَج مع مستوى AQI لكل يوم
- تقرير بأسلوب التقارير الرسمية

---

## 🎨 رموز AQI المعتمدة | Official AQI Color Scale

| النطاق | المستوى | اللون |
|--------|---------|-------|
| 0 – 50   | جيد / Good                        | 🟢 `#00E400` |
| 51 – 100 | معتدل / Moderate                  | 🟡 `#FFFF00` |
| 101 – 150| غير صحي للحساسين / Unhealthy (Sensitive) | 🟠 `#FF7E00` |
| 151 – 200| غير صحي / Unhealthy               | 🔴 `#FF0000` |
| 201 – 300| غير صحي جداً / Very Unhealthy     | 🟣 `#8F3F97` |
| 301 – 500| خطير / Hazardous                  | 🔴 `#7E0023` |

---

## 📊 المخرجات | Output Figures

| الملف | الوصف |
|-------|-------|
| `01_aqi_dashboard.png`      | لوحة تحكم شاملة: AQI يومي + حرارة + CO + دخان + توزيع الخطر + نمط ساعي + اتجاه |
| `02_aqi_reference_card.png` | بطاقة مرجعية رسمية لمقياس AQI بالعربية والإنجليزية |
| `03_pollutant_analysis.png` | تحليل الملوثات: AQI vs Risk Index، CO & Smoke، Temp/Humidity scatter، أيام الخطر |
| `04_forecast_report.png`    | تحليل الاتجاه الخطي + التنبؤ بـ AQI لـ 7 أيام مع فترة ثقة 95% |
| `05_official_report.png`    | تقرير بأسلوب التقارير الرسمية للجامعة |
| `daily_summary.csv`         | ملخص يومي مُعالَج مع مستوى AQI ومنحنى الاتجاه |

---

## 🚀 التثبيت والتشغيل | Installation & Usage

### 1. استنساخ المستودع
```bash
git clone https://github.com/<YOUR_USERNAME>/enviroai-mustansiriyah.git
cd enviroai-mustansiriyah
```

### 2. تثبيت المتطلبات
```bash
pip install -r requirements.txt
```

### 3. تشغيل التحليل

**ملف واحد:**
```bash
python enviroai_analysis.py --input EnviroAI_2026-06-05.xlsx
```

**ملفات متعددة (أكثر من تاريخ رصد):**
```bash
python enviroai_analysis.py --input april.xlsx may.xlsx june.xlsx
```

**كل ملفات xlsx في مجلد (glob):**
```bash
python enviroai_analysis.py --input data/*.xlsx --output output
```

**الوضع التفاعلي (بدون --input):**
```bash
python enviroai_analysis.py
# سيطلب منك إدخال مسار كل ملف يدوياً
```

### الخيارات
| الخيار | الاختصار | الوصف | القيمة الافتراضية |
|--------|----------|-------|-------------------|
| `--input`  | `-i` | مسار ملف واحد أو أكثر (.xlsx) | ← يطلب تفاعلياً |
| `--output` | `-o` | مجلد الحفظ (كل ملف يحصل على مجلد فرعي) | `output/` |

### هيكل المخرجات عند معالجة ملفات متعددة
```
output/
├── EnviroAI_April/
│   ├── 01_aqi_dashboard.png
│   ├── 02_aqi_reference_card.png
│   └── daily_summary.csv
├── EnviroAI_May/
│   ├── 01_aqi_dashboard.png
│   └── ...
└── EnviroAI_June/
    └── ...
```

---

## 📁 هيكل المشروع | Project Structure

```
enviroai-mustansiriyah/
├── enviroai_analysis.py        # ← الكود الرئيسي
├── requirements.txt            # ← المتطلبات
├── README.md                   # ← هذا الملف
├── EnviroAI_2026-06-05.xlsx   # ← ملف البيانات (غير مرفوع)
└── output/
    ├── 01_aqi_dashboard.png
    ├── 02_aqi_reference_card.png
    ├── 03_pollutant_analysis.png
    ├── 04_forecast_report.png
    ├── 05_official_report.png
    └── daily_summary.csv
```

---

## 📋 المتطلبات | Requirements

```
pandas>=2.0.0
openpyxl>=3.1.0
matplotlib>=3.7.0
seaborn>=0.12.0
numpy>=1.24.0
scipy>=1.10.0
Pillow>=10.0.0
```

---

## 📈 إحصاءات البيانات | Data Summary (April–May 2026)

| المؤشر | القيمة |
|--------|--------|
| إجمالي القراءات  | 4,910  |
| متوسط AQI        | 94.6   |
| أعلى AQI         | 462 (16 أبريل — طوارئ) |
| متوسط الحرارة    | 26.8 °C |
| أيام خطرة        | 8 / 13 |
| اتجاه AQI        | تصاعدي (+18/يوم) |

---

## 👥 لجنة إعداد التقرير

- **أ.م.د. حسام طارق محمد**
- **م. نعم داري ابراهيم**
- بإشراف رئيس قسم علوم الجو

---

## 📜 الترخيص | License

للأغراض الأكاديمية والبحثية — كلية العلوم ب / الجامعة المستنصرية  
For academic and research purposes — College of Science B / Al-Mustansiriyah University
