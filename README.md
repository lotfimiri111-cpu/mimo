# مذكرتي Pro — v18 Commercial

نظام توليد وبيع عروض PowerPoint لمذكرات التخرج (Licence / Master / Doctorat)

---

## 🏗️ بنية المشروع

```
v18-commercial/
├── app.py                    # Flask API الرئيسي
├── data.db                   # قاعدة بيانات SQLite (تُنشأ تلقائياً)
├── requirements.txt
├── render.yaml               # إعداد Render.com
│
├── core/
│   ├── models.py             # PresentationRequest
│   ├── themes.py             # 12 ثيم احترافي
│   ├── database.py           # طبقة البيانات (Orders, Codes, Settings)
│   └── preview.py            # PPTX→PDF→JPG engine مع Watermark
│
├── engine/
│   └── pipeline.py           # محركات Canva / Premium / Classic
│
├── public/
│   ├── index.html            # صفحة الإدخال (الطالب يملأ معلوماته)
│   ├── preview.html          # المعاينة الاحترافية المحمية
│   └── download.html         # صفحة إدخال كود التحميل
│
├── admin/
│   └── index.html            # لوحة الإدارة الكاملة
│
└── storage/                  # (لا يُخدَّم مباشرة — محمي)
    ├── pptx/                 # ملفات PPTX الأصلية (خارج public)
    ├── receipts/             # وصولات الدفع
    └── preview_cache/        # صور الشرائح المؤقتة (PDF + JPG)
```

---

## 🚀 رحلة الطالب

```
1. يفتح الموقع → يملأ معلومات المذكرة
2. يضغط "توليد العرض" (يستغرق 10-30 ثانية)
3. يُحوَّل تلقائياً إلى صفحة المعاينة
4. يشاهد جميع الشرائح بجودة كاملة (مع Watermark خفيف)
5. إذا أعجبه → يضغط "تحميل العرض"
6. تظهر معلومات الدفع (CCP / BaridiMob)
7. يدفع ويرفع وصل التسديد
8. يستقبل كود التحميل (بعد موافقة Admin)
9. يدخل الكود في صفحة /download
10. يحمّل ملف PPTX الأصلي بجودة كاملة
```

---

## 🔑 نظام الأكواد

- كل كود مرتبط بطلب واحد فقط
- صالح للاستخدام مرة واحدة (قابل للتعديل)
- صلاحية 7 أيام (قابلة للتغيير من الإعدادات)
- تسجيل IP + وقت كل تحميل
- يمكن تعطيله من لوحة الإدارة

---

## 🛡️ الحماية

- ملفات PPTX خارج /public (لا رابط مباشر)
- Watermark على جميع صور المعاينة
- Right-click معطّل على صفحة المعاينة
- Drag & Drop معطّل على الصور
- Headers: no-store, no-cache على صور الشرائح
- كل صورة تُطلب عبر token مؤقت (preview_token)

---

## ⚙️ التثبيت المحلي

```bash
pip install -r requirements.txt
python app.py
# يفتح على http://localhost:5000
# لوحة الإدارة: http://localhost:5000/admin
# كلمة المرور الافتراضية: admin1234
```

### المتطلبات الإضافية (نظام)
```bash
# Ubuntu/Debian
sudo apt install libreoffice poppler-utils

# أو عبر Docker — انظر Dockerfile
```

---

## 🌐 النشر على Render.com

1. ارفع المشروع على GitHub
2. أنشئ Web Service على Render
3. Build Command:
   ```
   pip install -r requirements.txt
   ```
4. Start Command:
   ```
   gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120
   ```
5. أضف في Render → Environment:
   - `SECRET_KEY` (أي قيمة عشوائية)

> **ملاحظة**: ليبر أوفيس وpoppler يجب أن يكونا متاحَين في بيئة Render
> اضف في Build Command:
> ```
> apt-get install -y libreoffice poppler-utils && pip install -r requirements.txt
> ```

---

## 💰 الإعدادات الافتراضية

| الإعداد | القيمة |
|---------|--------|
| السعر | 800 دج |
| صلاحية الكود | 7 أيام |
| كلمة مرور Admin | admin1234 |

**غيّر كلمة المرور فوراً من لوحة الإدارة → الإعدادات**

---

## 🔮 التطوير المستقبلي

- [ ] دفع إلكتروني تلقائي (BaridiPay / CIB)
- [ ] تسجيل حسابات الطلاب
- [ ] إشعارات SMS/Email تلقائية
- [ ] AI لاقتراح محتوى الشرائح
- [ ] اشتراكات شهرية
- [ ] تطبيق جوال

---

**مذكرتي Pro v18 — بُني بـ Flask + python-pptx + LibreOffice + pdf2image**
