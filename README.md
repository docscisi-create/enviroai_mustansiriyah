<div align="center">

<h1>🌍 EnviroAI</h1>
<h3>Air Quality & Weather Analysis System</h3>

<p>
  <b>كلية العلوم / الجامعة المستنصرية — بغداد / العراق</b><br/>
  بالتعاون مع <b>وزارة البيئة / مديرية البيئة الحضرية</b><br/>
  قسم مراقبة نوعية الهواء والضوضاء
</p>

<p>
  <img src="https://img.shields.io/badge/Python-3.9%2B-blue?logo=python"/>
  <img src="https://img.shields.io/badge/Dash-2.14%2B-informational?logo=plotly"/>
  <img src="https://img.shields.io/badge/Deploy-Railway-blueviolet?logo=railway"/>
  <img src="https://img.shields.io/badge/AQI-Official%20Scale-orange"/>
</p>

</div>

---

## 🚀 نشر التطبيق على Railway

### 1. رفع على GitHub
```bash
git init
git add .
git commit -m "initial: EnviroAI Dash web app"
git remote add origin https://github.com/USERNAME/enviroai.git
git push -u origin main
```

### 2. ربط Railway
1. افتح [railway.app](https://railway.app) → **New Project** → **Deploy from GitHub**
2. اختر المستودع
3. Railway يكتشف `Procfile` تلقائياً ويشغّل:
   ```
   web: gunicorn app:server
   ```
4. **بعد النشر يظهر رابط تلقائياً** مثل: `https://enviroai-xxxx.up.railway.app`

---

## 📁 هيكل الملفات

```
enviroai/
├── app.py            ← التطبيق الرئيسي (Dash + gunicorn server)
├── requirements.txt  ← المتطلبات
├── Procfile          ← أمر تشغيل Railway
├── .gitignore
└── README.md
```

---

## ⚙️ متطلبات Railway

| الملف | المحتوى |
|-------|---------|
| `Procfile` | `web: gunicorn app:server` |
| `requirements.txt` | dash, plotly, pandas, gunicorn … |
| `app.py` | يحتوي `server = app.server` (مطلوب لـ gunicorn) |

---

## 📊 مميزات التطبيق

- **رفع أي ملف Excel** مباشرة من المتصفح
- **5 تبويبات تفاعلية**: نظرة عامة — جودة الهواء — مستوى الخطر — تنبؤ — دليل AQI
- **رموز AQI الرسمية** بالألوان والوجوه التحذيرية
- **تنبؤ 7 أيام** مع فترة ثقة 95%
- يعمل على **أي متصفح** بدون تثبيت

---

## 📋 متطلبات ملف البيانات

يجب أن يحتوي ملف Excel على الأعمدة:
`Date, Time, Temperature, Humidity, AQI, CO, SMOKE, RiskIndex, RiskLevel`

---

## 🎨 رموز AQI المعتمدة

| النطاق | المستوى | اللون |
|:------:|---------|:-----:|
| 0–50 | جيد / Good | 🟢 |
| 51–100 | معتدل / Moderate | 🟡 |
| 101–150 | غير صحي للحساسين | 🟠 |
| 151–200 | غير صحي / Unhealthy | 🔴 |
| 201–300 | غير صحي جداً / Very Unhealthy | 🟣 |
| 301–500 | خطير / Hazardous | 🔴 |
