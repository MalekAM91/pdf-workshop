# Private PDF Workshop — ورشة PDF الخاصة

## 👉 [malekam91.github.io/pdf-workshop](https://malekam91.github.io/pdf-workshop/)

Open the link on any device (Windows, Mac, iPhone, Android) and start working. There is nothing to install.

افتح الرابط من أي جهاز (ويندوز، ماك، آيفون، أندرويد) وابدأ على طول. ما فيه شي تثبّته.

---

### What it does / وش يسوي

| English | العربي |
|---|---|
| PDF → Word (editable, Arabic and English) | PDF ← Word (نص قابل للتعديل، عربي وإنجليزي) |
| OCR for scanned PDFs (Arabic and English) | قراءة ضوئية OCR للملفات الممسوحة (عربي وإنجليزي) |
| PDF → images (PNG) · PDF → text | PDF ← صور · PDF ← نص |
| Images → PDF | صور ← PDF |
| Merge files, reorder, rotate, delete pages | دمج ملفات، ترتيب، تدوير، حذف صفحات |
| Extract pages (e.g. `1-3, 5`) | استخراج صفحات (مثلاً `1-3, 5`) |
| Write on pages, white-out text | الكتابة على الصفحات وتغطية النص |
| Watermark | علامة مائية |

### Privacy / الخصوصية

- **Your files never leave your device.** Everything runs inside your browser, and no server receives them.
- **Nothing is saved.** Close the page and the files are gone.
- **No outside code.** All libraries are bundled in this repository.

- **ملفاتك ما تطلع من جهازك.** كل شي يشتغل داخل المتصفح، وما فيه سيرفر يستقبلها.
- **ما يُحفظ شي.** تقفل الصفحة وتنمسح الملفات.
- **ما فيه أكواد من برا.** كل المكتبات موجودة داخل هذا المستودع.

### Limits / حدود

OCR is very good on clear scans but can misread a letter on blurry or skewed pages, so check important figures.

القراءة الضوئية دقيقة في المسح الواضح، لكن ممكن تغلط في حرف إذا الصفحة مائلة أو غير واضحة، فراجع الأرقام المهمة.

---

<details>
<summary>Desktop version (optional, advanced) / نسخة الكمبيوتر (اختيارية، للمتقدمين)</summary>

The `desktop/` folder has a Python version with extra tools: password protect and unlock, compression, metadata removal and page numbers.

مجلد `desktop/` فيه نسخة بايثون فيها أدوات زيادة: حماية بكلمة مرور وفكها، ضغط الحجم، إزالة البيانات الوصفية، وترقيم الصفحات.

**How to run it / طريقة التشغيل:**

1. Install Python 3.9 or newer. / ثبّت Python 3.9 أو أحدث.
2. Run `desktop/run_windows.bat` on Windows, or `desktop/run_mac.command` on Mac. / شغّل `desktop/run_windows.bat` على ويندوز، أو `desktop/run_mac.command` على ماك.

**Libraries / المكتبات:**

- pypdf (BSD)
- PyMuPDF and pdf2docx (AGPL-3.0)
- python-docx (MIT)

</details>

**Web libraries / مكتبات الموقع:**

- pdf.js (Apache-2.0)
- tesseract.js and its Arabic and English data (Apache-2.0)
- pdf-lib (MIT)
- JSZip (MIT)
