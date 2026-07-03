/**
 * ============================================================
 *  محرّك الكتابة في Google Sheet — للوحة تحكم موقع سمية عسيري
 * ============================================================
 * طريقة التركيب (مرة واحدة فقط):
 * 1) افتحي جدول الـ Google Sheet المرتبط بالموقع.
 * 2) من القائمة: الإضافات (Extensions) ← Apps Script.
 * 3) احذفي أي كود موجود في الملف، والصقي هذا الكود كاملاً مكانه.
 * 4) احفظي (Ctrl+S)، ثم من زر "نشر" (Deploy) أعلى يمين الصفحة اختاري
 *    "نشر جديد" (New deployment).
 * 5) اختاري النوع (Select type) ← تطبيق ويب (Web app).
 * 6) في "تنفيذ باسم" (Execute as) اختاري: أنا (Me).
 * 7) في "من له صلاحية الوصول" (Who has access) اختاري: أي مستخدم (Anyone).
 * 8) اضغطي "نشر" (Deploy) ثم وافقي على صلاحيات الوصول لحسابك.
 * 9) انسخي رابط "تطبيق الويب" (Web app URL) — هذا هو الرابط الذي
 *    تضعينه في أعلى صفحة لوحة التحكم (admin) في خانة "رابط الخدمة".
 *
 * ملاحظة: أي تعديل لاحق على هذا الكود يتطلب عمل "نشر جديد" مرة أخرى
 * (أو تعديل نفس الإصدار من خلال "إدارة عمليات النشر").
 * ============================================================
 */

// اسم رأس الأعمدة الثابت المستخدم في الموقع — لا تغيّري ترتيبه
var COLUMNS = ['النوع', 'الترتيب', 'العنوان', 'المجال', 'الوصف', 'الرابط', 'الصورة', 'الأيقونة'];

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: 'الخدمة تعمل بنجاح ✅' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    // تأكيد وجود صف رأس الأعمدة الصحيح — يُنشئه تلقائياً إن كان الجدول فارغاً
    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
      lastRow = 1;
    }

    var values = lastRow > 0 ? sheet.getRange(1, 1, lastRow, COLUMNS.length).getValues() : [COLUMNS];

    var items = body.rows ? body.rows : [body];
    var results = [];

    items.forEach(function (r) {
      var type = (r.type || '').toString().trim();
      var title = (r.title || '').toString().trim();
      var newRow = [
        type,
        r.order || '',
        title,
        r.category || '',
        r.desc || '',
        r.link || '',
        r.image || '',
        r.icon || ''
      ];

      // البحث عن صف موجود بنفس (النوع + العنوان) لتحديثه بدل تكراره
      var foundRowIndex = -1;
      for (var i = 1; i < values.length; i++) {
        var rowType = (values[i][0] || '').toString().trim();
        var rowTitle = (values[i][2] || '').toString().trim();
        if (rowType === type && rowTitle === title && title !== '') {
          foundRowIndex = i + 1; // فهرس الصف الفعلي في الشيت (1-based)
          break;
        }
      }

      if (foundRowIndex > 0) {
        sheet.getRange(foundRowIndex, 1, 1, COLUMNS.length).setValues([newRow]);
        results.push({ type: type, title: title, action: 'updated' });
      } else {
        sheet.appendRow(newRow);
        values.push(newRow);
        results.push({ type: type, title: title, action: 'added' });
      }
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, results: results }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
