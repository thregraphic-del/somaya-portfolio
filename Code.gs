/**
 * ============================================================
 *  محرّك الكتابة في Google Sheet — للوحة تحكم موقع سمية عسيري
 *  (النسخة المحدَّثة — تدعم 5 تبويبات منفصلة)
 * ============================================================
 * طريقة التركيب (مرة واحدة فقط):
 * 1) افتحي جدول الـ Google Sheet الجديد (ذو التبويبات الخمسة).
 * 2) من القائمة: الإضافات (Extensions) ← Apps Script.
 * 3) احذفي أي كود موجود في الملف، والصقي هذا الكود كاملاً مكانه.
 * 4) احفظي (Ctrl+S)، ثم من زر "نشر" (Deploy) أعلى يمين الصفحة اختاري
 *    "نشر جديد" (New deployment).
 * 5) اختاري النوع (Select type) ← تطبيق ويب (Web app).
 * 6) في "تنفيذ باسم" (Execute as) اختاري: أنا (Me).
 * 7) في "من له صلاحية الوصول" (Who has access) اختاري: أي مستخدم (Anyone).
 * 8) اضغطي "نشر" (Deploy) ثم وافقي على صلاحيات الوصول لحسابك.
 * 9) انسخي رابط "تطبيق الويب" (Web app URL) — هذا هو الرابط الذي
 *    تضعينه في أعلى صفحة لوحة التحكم (admin) في خانة "رابط خدمة الحفظ".
 *
 * ملاحظة: إذا كان هذا نفس المشروع القديم وتحدّثين الكود فقط، استخدمي
 * "إدارة عمليات النشر" (Manage deployments) ← تعديل ← إصدار جديد ← نشر،
 * حتى يبقى نفس الرابط القديم يعمل دون الحاجة لتحديثه في لوحة التحكم.
 * ============================================================
 */

// أسماء التبويبات — يجب أن تطابق أسماء الأوراق (Sheets) في الملف بالضبط
var TABS = {
  profile:   'الملف_الشخصي',
  skill:     'المهارات',
  client:    'الشعارات',
  dashboard: 'الداشبورد',
  blog:      'المدونة'
};

// أعمدة موحّدة لكل التبويبات
var COLUMNS = ['الترتيب', 'العنوان', 'المجال', 'الوصف', 'الرابط', 'الصورة', 'الأيقونة'];

function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
  }
  return sheet;
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: 'الخدمة تعمل بنجاح ✅' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var items = body.rows ? body.rows : [body];
    var results = [];

    items.forEach(function (r) {
      var type = (r.type || '').toString().trim().toLowerCase();
      var tabName = TABS[type];
      if (!tabName) {
        results.push({ type: type, action: 'skipped-unknown-type' });
        return;
      }

      var sheet = getOrCreateSheet_(tabName);
      var lastRow = sheet.getLastRow();
      var values = lastRow > 0 ? sheet.getRange(1, 1, lastRow, COLUMNS.length).getValues() : [COLUMNS];

      var title = (r.title || '').toString().trim();
      var newRow = [
        r.order || '',
        title,
        r.category || '',
        r.desc || '',
        r.link || '',
        r.image || '',
        r.icon || ''
      ];

      // البحث عن صف موجود بنفس العنوان داخل هذا التبويب لتحديثه بدل تكراره
      var foundRowIndex = -1;
      for (var i = 1; i < values.length; i++) {
        var rowTitle = (values[i][1] || '').toString().trim();
        if (rowTitle === title && title !== '') {
          foundRowIndex = i + 1; // فهرس الصف الفعلي في الشيت (1-based)
          break;
        }
      }

      if (foundRowIndex > 0) {
        sheet.getRange(foundRowIndex, 1, 1, COLUMNS.length).setValues([newRow]);
        results.push({ type: type, title: title, tab: tabName, action: 'updated' });
      } else {
        sheet.appendRow(newRow);
        results.push({ type: type, title: title, tab: tabName, action: 'added' });
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
