# PDF Workshop — ورشة PDF المحلية

A small PDF toolkit that runs **on your own computer**. It opens a page in your browser at `127.0.0.1`, works in **Arabic and English**, needs no internet after the first install, and keeps no copies of your files.

برنامج خفيف لملفات PDF يشتغل **على جهازك أنت**. يفتح صفحة في المتصفح على `127.0.0.1`، ويدعم **العربي والإنجليزي**، وما يحتاج إنترنت بعد أول تثبيت، وما يحتفظ بأي نسخة من ملفاتك.

[العربية ↓](#بالعربي)

---

## English

### Download

On this page click **Code → Download ZIP**, then unzip it anywhere.

### Run

**Windows**
1. Install Python from <https://www.python.org/downloads/>. During setup, tick **Add python.exe to PATH**.
2. Double-click `run_windows.bat`.
   - The first run installs the libraries into a `.venv` folder next to the script. This needs internet and takes a minute or two.
   - After that, the page opens in your browser.

**macOS**
Double-click `run_mac.command`, or run `./run_mac.command` in Terminal.

**Manual**
```bash
pip install -r requirements.txt
python pdf_workshop.py
```

Use the **English / عربي** button at the top of the page to switch language.

### Tools

| Group | Tool | Library |
|---|---|---|
| Convert | PDF to Word: English keeps its layout and tables; Arabic becomes clean editable right-to-left text | pdf2docx / pdf.js + python-docx |
| Convert | PDF to text (TXT) | pdf.js |
| Convert | PDF to images (PNG/JPG, 100–300 DPI) | PyMuPDF |
| Convert | Images to PDF (A4 or original size) | PyMuPDF |
| Organize | Merge files in the order you set | pypdf |
| Organize | Split: chosen pages, one file per page, or every N pages | pypdf |
| Organize | Reorder, delete and rotate pages | pypdf |
| Protect | Password protection (AES-256) with print and copy controls | pypdf |
| Protect | Remove a password you know | pypdf |
| Protect | Remove metadata (author, software, dates) before sending a file outside | pypdf |
| Edit | Diagonal watermark, in Arabic or English | PyMuPDF |
| Edit | Page numbers ("Page 1 of 10" or «صفحة 1 من 10») | PyMuPDF |
| Edit | Compress, with three levels | PyMuPDF |

### Security

- **Local only.** The server listens on `127.0.0.1`, so nobody else on the network can reach it.
- **Secret token for each session.** The link carries a random token that changes every time you start the program, and any request without it is refused. This stops other websites open in your browser from sending requests to the program, and blocks DNS-rebinding attacks.
- **Nothing is stored.** Each job runs in a temporary folder that is deleted as soon as the result is sent, and file names are never logged.
- **No internet needed.** pdf.js is bundled in `vendor/`, and the page's Content-Security-Policy blocks all outside connections.
- **Stops by itself.** The program stops two minutes after you close the page. You can also click **Stop the program**.

### Known limits

- **Scanned PDFs** are images with no text. PDF to Word marks those pages as having no text, because reading them needs OCR (for example Tesseract).
- **Arabic tables** come out as lines of text. To keep the table itself, choose **Keep layout and tables (pdf2docx)**; Arabic word order is weaker in that mode.

---

## بالعربي

### التنزيل

من صفحة المستودع اضغط **Code ← Download ZIP**، وفك الضغط في أي مكان.

### التشغيل

**ويندوز**
1. ثبّت Python من <https://www.python.org/downloads/>، وأثناء التثبيت فعّل خيار **Add python.exe to PATH**.
2. اضغط مرتين على `run_windows.bat`.
   - أول مرة يثبّت المكتبات في مجلد `.venv` جنب الملف. يحتاج إنترنت ويأخذ دقيقة أو دقيقتين.
   - بعدها تنفتح الصفحة في المتصفح.

**ماك**
اضغط مرتين على `run_mac.command`، أو شغّل `./run_mac.command` من Terminal.

تقدر تغيّر اللغة من زر **English / عربي** أعلى الصفحة.

### الأدوات

- **تحويل:**
  - PDF إلى Word: العربي يطلع نص قابل للتعديل باتجاه صحيح، والإنجليزي يحافظ على التنسيق والجداول.
  - PDF إلى نص، PDF إلى صور، صور إلى PDF.
- **تنظيم:** دمج، تقسيم، استخراج صفحات، ترتيب جديد، حذف، تدوير.
- **حماية:**
  - حماية بكلمة مرور (AES-256) مع التحكم بالطباعة والنسخ.
  - فك الحماية لملف تعرف كلمة مروره.
  - إزالة البيانات الوصفية (اسم الكاتب، البرنامج، التواريخ) قبل إرسال الملف لجهة خارجية.
- **تعديل:** علامة مائية بالعربي أو الإنجليزي، ترقيم الصفحات، ضغط الحجم.

### الأمان

- **السيرفر محلي فقط.** يسمع على `127.0.0.1`، فما أحد في الشبكة يقدر يوصل له.
- **رمز سري لكل جلسة.** الرابط فيه رمز عشوائي يتغير كل مرة تشغّل البرنامج، وأي طلب بدون الرمز ينرفض. هذا يمنع أي موقع ثاني مفتوح في المتصفح من إرسال طلبات للبرنامج.
- **ما يحفظ شي.** كل عملية تشتغل في مجلد مؤقت ينحذف أول ما تنتهي، والبرنامج ما يسجّل أسماء الملفات.
- **ما يحتاج إنترنت.** مكتبة pdf.js موجودة في مجلد `vendor/`، والصفحة تمنع أي اتصال خارجي.
- **يقفل نفسه.** يتوقف البرنامج بعد دقيقتين من إغلاق الصفحة، أو تضغط «إيقاف البرنامج».

### حدود معروفة

- **الملفات الممسوحة بالسكانر** ما فيها نص، فتحتاج OCR، ومو موجود في البرنامج حالياً.
- **الجداول العربية** تطلع في Word كسطور نص. إذا تبي الجدول نفسه، اختر «الحفاظ على التنسيق والجداول»، بس ترتيب الكلمات العربية فيه أضعف.

---

### Libraries and licences / المكتبات والتراخيص

| Library | Licence |
|---|---|
| pypdf | BSD |
| PyMuPDF | AGPL-3.0 |
| pdf2docx | AGPL-3.0 |
| python-docx | MIT |
| pdf.js 3.11.174, Mozilla, bundled in `vendor/` | Apache-2.0 |

PyMuPDF and pdf2docx are AGPL-licensed. Using this tool inside your organisation is fine. If you redistribute a modified version, share its source under the same licence.
