/**
 * ===================================================================
 *  نظام تسجيل غياب الطلاب - Backend API (Google Apps Script)
 *  Student Attendance System - Backend
 * ===================================================================
 *
 *  التبويبات المطلوبة في جدول البيانات:
 *    1) "بيانات_الطلاب"  ->  A: رقم الطالب | B: اسم الطالب | C: الصف | D: الفصل
 *    2) "سجل_الغياب"     ->  A: التاريخ | B: رقم الطالب | C: اسم الطالب | D: الصف | E: الفصل | F: الحالة
 *
 *  طريقة النشر:
 *    Deploy > New deployment > Type: Web app
 *    Execute as: Me
 *    Who has access: Anyone
 *    انسخ الرابط (Web app URL) والصقه في الواجهة الأمامية.
 * ===================================================================
 */

// أسماء التبويبات - يجب أن تطابق جدول البيانات تماماً
var STUDENTS_SHEET   = 'بيانات_الطلاب';
var ATTENDANCE_SHEET = 'سجل_الغياب';

/**
 * بناء استجابة JSON موحّدة.
 * ملاحظة: تطبيقات Apps Script لا تدعم رؤوس CORS المخصصة،
 * لذلك تعتمد الواجهة على إرسال نوع المحتوى text/plain لتجنب طلب preflight.
 */
function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * doGet - قراءة البيانات
 *  action=getStudents  ->  يعيد مصفوفة الطلاب [{id, name, grade, classRoom}]
 */
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';

    if (action === 'getStudents') {
      return jsonOutput(getStudents());
    }

    /* ===== توجيهات الموديول الجديد (إضافة فقط — لا تحذف القديم) ===== */
    if (action === 'getAllAttendance') {
      return jsonOutput(getAllAttendance());
    }
    if (action === 'logViolation') {
      return jsonOutput(logViolation(e.parameter));
    }
    if (action === 'getViolations') {
      return jsonOutput(getViolations());
    }
    /* ===== توجيهات وحدة إدارة الطلاب (CRUD) ===== */
    if (action === 'addStudent') {
      return jsonOutput(addStudent(e.parameter));
    }
    if (action === 'updateStudent') {
      return jsonOutput(updateStudent(e.parameter));
    }
    if (action === 'deleteStudent') {
      return jsonOutput(deleteStudent(e.parameter));
    }
    /* ===== توجيهات حزمة الميزات المتقدمة (تسوية الإنذارات + عداد الأجهزة) ===== */
    if (action === 'settleAbsence') {
      return jsonOutput(settleAbsence(e.parameter));
    }
    if (action === 'devicePing') {
      return jsonOutput(devicePing(e.parameter));
    }
    /* ===== توجيهات سجل الإجراءات المتخذة مع الطالب ===== */
    if (action === 'logAction') {
      return jsonOutput(logAction(e.parameter));
    }
    if (action === 'getActions') {
      return jsonOutput(getActions(e.parameter));
    }
    /* ===== توجيهات رقم تليفون الطالب (عمود جديد في شيت الطلاب) ===== */
    if (action === 'setStudentPhone') {
      return jsonOutput(setStudentPhone(e.parameter));
    }
    if (action === 'getStudentPhone') {
      return jsonOutput(getStudentPhone(e.parameter));
    }
    /* ============================================================== */

    // استجابة افتراضية للتحقق من أن الخدمة تعمل
    return jsonOutput({
      ok: true,
      message: 'نظام تسجيل الغياب يعمل بنجاح. استخدم action=getStudents.'
    });

  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

/**
 * قراءة تبويب "بيانات_الطلاب" وتجاهل صف العناوين.
 * يعيد: [{ id, name, grade, classRoom }]
 */
function getStudents() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(STUDENTS_SHEET);
  if (!sheet) {
    throw new Error('لم يتم العثور على التبويب: ' + STUDENTS_SHEET);
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // لا يوجد سوى صف العناوين أو فارغ

  // قراءة الأعمدة A:D من الصف الثاني حتى آخر صف
  var values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

  var students = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var id  = String(row[0]).trim();
    if (id === '') continue; // تجاهل الصفوف الفارغة

    students.push({
      id:        id,
      name:      String(row[1]).trim(),
      grade:     String(row[2]).trim(),
      classRoom: String(row[3]).trim()
    });
  }
  return students;
}

/* ===================================================================
 *  إعدادات تقسيم البيانات (Data Sharding)
 * =================================================================== */

// مفتاح التقسيم:
//   'classRoom' = شيت مستقل لكل فصل (مثل "2/أ")  ← الافتراضي
//   'grade'     = شيت مستقل لكل صف  (مثل "الثاني")
var SHARD_BY = 'classRoom';

// صف العناوين الموحّد لكل شيت غياب (نفس ترتيب أعمدة النظام الأصلي)
var SHARD_HEADERS = ['التاريخ', 'رقم الطالب', 'اسم الطالب', 'الصف', 'الفصل', 'الحالة'];

/**
 * doPost - حفظ / تحديث الغياب مع تقسيم البيانات (Data Sharding).
 *
 * الحمولة القادمة من الواجهة كما هي دون أي تغيير:
 *   { date: "YYYY-MM-DD", records: [ {id, name, grade, classRoom, status}, ... ] }
 *
 * المنطق:
 *   1) تُجمَّع السجلات حسب اسم الفصل (مفتاح التقسيم SHARD_BY).
 *   2) لكل فصل: إن لم يوجد شيت باسمه يُنشأ تلقائياً مع صف العناوين.
 *   3) إن وُجد سجل بنفس (التاريخ + رقم الطالب) داخل شيت الفصل -> يُحدّث عمود الحالة.
 *      وإلا -> يُضاف صف جديد.
 */
function doPost(e) {
  // قفل لمنع الكتابة المتزامنة وتلف البيانات
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOutput({ ok: false, error: 'لا توجد بيانات في الطلب.' });
    }

    var payload = JSON.parse(e.postData.contents);
    var date    = payload.date;
    var records = payload.records;

    if (!date || !records || !records.length) {
      return jsonOutput({ ok: false, error: 'الحمولة غير صحيحة: التاريخ أو السجلات مفقودة.' });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // (1) تجميع السجلات حسب اسم الفصل (مفتاح التقسيم)
    var groups = {}; // { sheetName: [record, ...] }
    for (var i = 0; i < records.length; i++) {
      var rec = records[i];
      var shardName = String(rec[SHARD_BY] || '').trim();
      if (shardName === '') shardName = 'غير_محدد'; // حماية: سجل بلا فصل
      if (!groups[shardName]) groups[shardName] = [];
      groups[shardName].push(rec);
    }

    var totalInserted = 0;
    var totalUpdated  = 0;
    var sheetsTouched = [];

    // (2) معالجة كل فصل في شيته الخاص
    for (var sheetName in groups) {
      if (!groups.hasOwnProperty(sheetName)) continue;

      var sheet        = getOrCreateShardSheet(ss, sheetName);
      var classRecords = groups[sheetName];

      // فهرس السجلات الموجودة داخل هذا الشيت: "التاريخ|رقم الطالب" -> رقم الصف
      var existing = buildShardIndex(sheet);

      var newRows = [];
      for (var j = 0; j < classRecords.length; j++) {
        var r      = classRecords[j];
        var id     = String(r.id).trim();
        var status = String(r.status).trim();
        var key    = date + '|' + id;

        if (existing.hasOwnProperty(key)) {
          // (3أ) تحديث عمود الحالة (العمود السادس F) للصف الموجود
          sheet.getRange(existing[key], 6).setValue(status);
          totalUpdated++;
        } else {
          // (3ب) تجهيز صف جديد
          newRows.push([
            date,
            id,
            String(r.name).trim(),
            String(r.grade).trim(),
            String(r.classRoom).trim(),
            status
          ]);
          existing[key] = -1; // منع التكرار داخل نفس الطلب
          totalInserted++;
        }
      }

      // إضافة الصفوف الجديدة دفعة واحدة (أسرع وأكثر أماناً)
      if (newRows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, SHARD_HEADERS.length).setValues(newRows);
      }
      sheetsTouched.push(sheetName);
    }

    return jsonOutput({
      ok: true,
      message: 'تم حفظ الغياب بنجاح ✓  (جديد: ' + totalInserted +
               '، محدّث: ' + totalUpdated + ') في: ' + sheetsTouched.join('، '),
      inserted: totalInserted,
      updated: totalUpdated,
      sheets: sheetsTouched
    });

  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * إيجاد شيت الفصل بالاسم، أو إنشاؤه تلقائياً مع صف العناوين عند عدم وجوده.
 */
function getOrCreateShardSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(SHARD_HEADERS);
    // تنسيق صف العناوين: خط عريض + تجميد الصف الأول
    sheet.getRange(1, 1, 1, SHARD_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * بناء فهرس لسجلات شيت فصل معيّن: "التاريخ|رقم الطالب" -> رقم الصف (1-based).
 * يعيد استخدام normalizeDate الموجودة مسبقاً لتوحيد صيغة التاريخ.
 */
function buildShardIndex(sheet) {
  var index   = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return index; // لا يوجد سوى صف العناوين

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // العمودان A و B فقط
  for (var i = 0; i < data.length; i++) {
    var id = String(data[i][1]).trim();
    if (id === '') continue;
    var key = normalizeDate(data[i][0]) + '|' + id;
    index[key] = i + 2; // +2 لأن الصفوف 1-based ونبدأ بعد العناوين
  }
  return index;
}

/**
 * بناء فهرس للسجلات الحالية في "سجل_الغياب".
 * المفتاح: "التاريخ|رقم الطالب"  ->  رقم الصف (1-based)
 * يتم تطبيع التاريخ ليطابق صيغة YYYY-MM-DD سواء كان نصاً أو كائن Date.
 */
function buildIndex(sheet) {
  var index   = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return index;

  // قراءة العمودين A (التاريخ) و B (رقم الطالب) فقط
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  for (var i = 0; i < data.length; i++) {
    var dateVal = data[i][0];
    var id      = String(data[i][1]).trim();
    if (id === '') continue;

    var dateKey = normalizeDate(dateVal);
    var key     = dateKey + '|' + id;
    index[key]  = i + 2; // +2 لأن الصفوف 1-based ونبدأ من الصف الثاني
  }
  return index;
}

/**
 * تطبيع التاريخ إلى صيغة YYYY-MM-DD.
 * يتعامل مع التواريخ المخزّنة كنص أو ككائن Date.
 */
function normalizeDate(value) {
  if (value instanceof Date) {
    var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }
  return String(value).trim();
}


/* ===================================================================
 *  ▼▼▼  موديول إضافي: الإحصائيات + المخالفات  ▼▼▼
 *  أُلحق بنهاية الملف — لا يعدّل أي كود أصلي بالأعلى.
 *  التوجيه يتم عبر الفروع المضافة داخل doGet.
 * =================================================================== */

// اسم شيت المخالفات (يُنشأ تلقائياً إن لم يكن موجوداً)
var VIOLATIONS_SHEET = 'شيت المخالفات';

/**
 * قراءة كل سجلات الغياب من جميع شيتات الفصول (بعد تطبيق التقسيم Sharding).
 * تمرّ على كل تبويبات الملف وتأخذ فقط الشيتات التي تطابق صف عناوين الغياب،
 * فتتجاهل تلقائياً شيت "بيانات_الطلاب" و "شيت المخالفات" وأي شيت آخر.
 * (تشمل أيضاً شيت "سجل_الغياب" القديم إن وُجد، فلا تُفقد البيانات السابقة).
 */
function readAttendanceRows() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var rows   = [];

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];

    // تجاهل الشيتات النظامية صراحةً
    if (sheet.getName() === STUDENTS_SHEET || sheet.getName() === VIOLATIONS_SHEET) continue;

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;

    // التحقق أن هذا شيت غياب فعلاً (أول خلية في العناوين = "التاريخ")
    var firstHeader = String(sheet.getRange(1, 1).getValue()).trim();
    if (firstHeader !== SHARD_HEADERS[0]) continue;

    var values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for (var i = 0; i < values.length; i++) {
      var r = values[i];
      if (String(r[1]).trim() === '') continue; // تجاهل الصفوف الفارغة
      rows.push({
        date:      normalizeDate(r[0]),
        id:        String(r[1]).trim(),
        name:      String(r[2]).trim(),
        grade:     String(r[3]).trim(),
        classRoom: String(r[4]).trim(),
        status:    String(r[5]).trim()
      });
    }
  }
  return rows;
}

/**
 * نقطة وصول getAllAttendance:
 * تُعيد كل سجلات الغياب ليحسب الـ Frontend النسب لحظياً (فلترة فورية بلا تكرار طلبات).
 */
function getAllAttendance() {
  return readAttendanceRows();
}

/**
 * نقطة وصول logViolation:
 * تسجّل مخالفة في "شيت المخالفات" بالأعمدة: اسم الطالب | التاريخ | المخالفة | الصف.
 * تستقبل المعطيات عبر باراميترات الرابط (e.parameter): name, type, grade.
 */
function logViolation(params) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var name  = String((params && params.name)  || '').trim();
    var type  = String((params && params.type)  || '').trim();
    var grade = String((params && params.grade) || '').trim();

    if (!name || !type) {
      return { ok: false, error: 'اسم الطالب أو نوع المخالفة مفقود.' };
    }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(VIOLATIONS_SHEET);

    // إنشاء الشيت تلقائياً مع صف العناوين عند أول استخدام
    if (!sheet) {
      sheet = ss.insertSheet(VIOLATIONS_SHEET);
      sheet.appendRow(['اسم الطالب', 'التاريخ', 'المخالفة', 'الصف']);
    }

    var tz    = ss.getSpreadsheetTimeZone();
    var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    sheet.appendRow([name, today, type, grade]);

    return { ok: true, message: 'تم تسجيل المخالفة بنجاح ✓' };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}


/* ===================================================================
 *  ▼▼▼  موديول إضافي: قراءة المخالفات (للتقارير ولوحات الصدارة)  ▼▼▼
 * =================================================================== */

/**
 * قراءة كل المخالفات من "شيت المخالفات".
 * الأعمدة: A: اسم الطالب | B: التاريخ | C: المخالفة | D: الصف
 * تُعيد: [{ name, date, type, grade }]
 */
function getViolations() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(VIOLATIONS_SHEET);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (String(r[0]).trim() === '') continue; // تجاهل الصفوف الفارغة
    rows.push({
      name:  String(r[0]).trim(),
      date:  normalizeDate(r[1]),
      type:  String(r[2]).trim(),
      grade: String(r[3]).trim()
    });
  }
  return rows;
}


/* ===================================================================
 *  ▼▼▼  موديول إضافي: إدارة الطلاب (CRUD)  ▼▼▼
 *  إضافة / تعديل / حذف طلاب في تبويب "بيانات_الطلاب"
 *  أُلحق بنهاية الملف — التوجيه عبر doGet (addStudent / updateStudent / deleteStudent)
 * =================================================================== */

/**
 * إيجاد رقم الصف (1-based) لطالب في "بيانات_الطلاب" حسب رقم الطالب.
 * يعيد -1 إن لم يوجد.
 */
function findStudentRow(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var target = String(id).trim();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === target) return i + 2; // +2 لأن البيانات تبدأ من الصف الثاني
  }
  return -1;
}

/**
 * addStudent - إضافة طالب جديد.
 * المعطيات: id, name, grade, classRoom
 * يرفض الإضافة لو كان رقم الطالب مستخدماً مسبقاً.
 */
function addStudent(params) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var id        = String((params && params.id)        || '').trim();
    var name      = String((params && params.name)      || '').trim();
    var grade     = String((params && params.grade)     || '').trim();
    var classRoom = String((params && params.classRoom) || '').trim();

    if (!id || !name || !grade || !classRoom) {
      return { ok: false, error: 'كل الحقول مطلوبة: رقم الطالب، الاسم، الصف، الفصل.' };
    }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(STUDENTS_SHEET);
    if (!sheet) return { ok: false, error: 'لم يتم العثور على التبويب: ' + STUDENTS_SHEET };

    // منع تكرار رقم الطالب
    if (findStudentRow(sheet, id) !== -1) {
      return { ok: false, error: 'رقم الطالب (' + id + ') مستخدم بالفعل لطالب آخر.' };
    }

    sheet.appendRow([id, name, grade, classRoom]);
    return { ok: true, message: 'تمت إضافة الطالب «' + name + '» بنجاح ✓' };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

/**
 * updateStudent - تعديل بيانات طالب.
 * المعطيات: originalId (الرقم الأصلي للبحث) + id, name, grade, classRoom (القيم الجديدة)
 */
function updateStudent(params) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var originalId = String((params && params.originalId) || '').trim();
    var id         = String((params && params.id)         || '').trim();
    var name       = String((params && params.name)       || '').trim();
    var grade      = String((params && params.grade)      || '').trim();
    var classRoom  = String((params && params.classRoom)  || '').trim();

    if (!originalId || !id || !name || !grade || !classRoom) {
      return { ok: false, error: 'كل الحقول مطلوبة للتعديل.' };
    }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(STUDENTS_SHEET);
    if (!sheet) return { ok: false, error: 'لم يتم العثور على التبويب: ' + STUDENTS_SHEET };

    var row = findStudentRow(sheet, originalId);
    if (row === -1) {
      return { ok: false, error: 'لم يتم العثور على طالب برقم (' + originalId + ').' };
    }

    // لو غُيّر رقم الطالب، نتأكد أن الرقم الجديد غير مستخدم لطالب آخر
    if (id !== originalId && findStudentRow(sheet, id) !== -1) {
      return { ok: false, error: 'الرقم الجديد (' + id + ') مستخدم بالفعل لطالب آخر.' };
    }

    // تحديث الأعمدة الأربعة دفعة واحدة A:D
    sheet.getRange(row, 1, 1, 4).setValues([[id, name, grade, classRoom]]);
    return { ok: true, message: 'تم تعديل بيانات الطالب «' + name + '» بنجاح ✓' };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

/**
 * deleteStudent - حذف طالب نهائياً من "بيانات_الطلاب" حسب رقمه.
 * المعطيات: id
 * ملاحظة: لا يحذف سجلات الغياب التاريخية للطالب من شيتات الفصول (تبقى للأرشيف).
 */
function deleteStudent(params) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var id = String((params && params.id) || '').trim();
    if (!id) return { ok: false, error: 'رقم الطالب مطلوب للحذف.' };

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(STUDENTS_SHEET);
    if (!sheet) return { ok: false, error: 'لم يتم العثور على التبويب: ' + STUDENTS_SHEET };

    var row = findStudentRow(sheet, id);
    if (row === -1) {
      return { ok: false, error: 'لم يتم العثور على طالب برقم (' + id + ').' };
    }

    sheet.deleteRow(row);
    return { ok: true, message: 'تم حذف الطالب بنجاح ✓' };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}



/* ===================================================================
 *  ▼▼▼  حزمة الميزات المتقدمة: تسوية الإنذارات + عداد الأجهزة  ▼▼▼
 *  أُلحقت بنهاية الملف — لا تعدّل أي كود أصلي بالأعلى.
 * =================================================================== */

/**
 * settleAbsence — تسوية الإنذار:
 * تحويل أيام غياب محددة لطالب من «غياب» إلى «غياب بعذر» في كل شيتات الفصول.
 * المعاملات: id=رقم الطالب ، dates=قائمة تواريخ مفصولة بفواصل (yyyy-MM-dd,...)
 */
function settleAbsence(params) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var id = String((params && params.id) || '').trim();
    var datesArr = String((params && params.dates) || '')
      .split(',').map(function (s) { return s.trim(); }).filter(String);
    if (!id || !datesArr.length) {
      return { ok: false, error: 'بيانات غير مكتملة: مطلوب رقم الطالب وتواريخ التسوية' };
    }
    var dateSet = {};
    for (var i = 0; i < datesArr.length; i++) dateSet[datesArr[i]] = true;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var updated = 0;

    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      if (sheet.getName() === STUDENTS_SHEET || sheet.getName() === VIOLATIONS_SHEET) continue;
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) continue;
      // شيتات الغياب فقط (أول عنوان = "التاريخ")
      if (String(sheet.getRange(1, 1).getValue()).trim() !== SHARD_HEADERS[0]) continue;

      var values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (var r = 0; r < values.length; r++) {
        if (String(values[r][1]).trim() !== id) continue;
        var dk = normalizeDate(values[r][0]);
        if (!dateSet[dk]) continue;
        // نحوّل «غياب» فقط (بدون عذر) — «غياب بعذر» يبقى كما هو
        if (String(values[r][5]).trim() === 'غياب') {
          sheet.getRange(r + 2, 6).setValue('غياب بعذر');
          updated++;
        }
      }
    }
    return { ok: true, updated: updated, message: 'تم تحويل ' + updated + ' يوم إلى «غياب بعذر» ✓' };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

/**
 * devicePing — عداد الأجهزة المتصلة:
 * كل جهاز يرسل نبضة كل دقيقة بمعرّفه؛ يُحتسب «متصلاً» إن كانت آخر نبضة خلال 5 دقائق.
 * التخزين في CacheService (مؤقت وخفيف — لا يلمس الشيتات إطلاقاً).
 */
function devicePing(params) {
  try {
    var cache = CacheService.getScriptCache();
    var raw = cache.get('active_devices');
    var map = {};
    try { map = raw ? JSON.parse(raw) : {}; } catch (e) { map = {}; }

    var now = Date.now();
    var device = String((params && params.device) || '').trim();
    if (device) map[device] = now;

    // تنظيف الأجهزة الخاملة (أقدم من 5 دقائق) وعدّ النشِط
    var fresh = {};
    var count = 0;
    for (var k in map) {
      if (now - map[k] < 5 * 60 * 1000) { fresh[k] = map[k]; count++; }
    }
    cache.put('active_devices', JSON.stringify(fresh), 600); // صلاحية 10 دقائق
    return { ok: true, count: count };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}


/* ===================================================================
 *  ▼▼▼  سجل الإجراءات المتخذة مع الطالب (رسالة / استدعاء ولي أمر…)  ▼▼▼
 *  أُلحق بنهاية الملف — لا يعدّل أي كود أصلي بالأعلى.
 * =================================================================== */

var ACTIONS_SHEET = 'سجل_الاجراءات';
var ACTIONS_HEADERS = ['التاريخ', 'رقم الطالب', 'اسم الطالب', 'نوع الإجراء', 'ملاحظات'];

// إيجاد/إنشاء شيت الإجراءات مع صف العناوين
function getActionsSheet(ss) {
  var sheet = ss.getSheetByName(ACTIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ACTIONS_SHEET);
    sheet.getRange(1, 1, 1, ACTIONS_HEADERS.length).setValues([ACTIONS_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * logAction — تسجيل إجراء جديد متخذ مع طالب.
 * المعاملات: id ، name ، type ، note (اختياري)
 */
function logAction(params) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var id   = String((params && params.id)   || '').trim();
    var name = String((params && params.name) || '').trim();
    var type = String((params && params.type) || '').trim();
    var note = String((params && params.note) || '').trim();
    if (!id || !type) {
      return { ok: false, error: 'بيانات غير مكتملة: مطلوب رقم الطالب ونوع الإجراء' };
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getActionsSheet(ss);
    var today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    sheet.appendRow([today, id, name, type, note]);
    return { ok: true, message: 'تم تسجيل الإجراء بنجاح ✓' };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

/**
 * getActions — إرجاع الإجراءات؛ إن مُرِّر id تُرجَع إجراءات هذا الطالب فقط.
 * يعيد: [{ date, id, name, type, note }]
 */
function getActions(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ACTIONS_SHEET);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var filterId = String((params && params.id) || '').trim();
  var values = sheet.getRange(2, 1, lastRow - 1, ACTIONS_HEADERS.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    var id = String(r[1]).trim();
    if (id === '') continue;
    if (filterId && id !== filterId) continue;
    out.push({
      date: normalizeDate(r[0]),
      id:   id,
      name: String(r[2]).trim(),
      type: String(r[3]).trim(),
      note: String(r[4]).trim()
    });
  }
  return out;
}


/* ===================================================================
 *  ▼▼▼  رقم تليفون الطالب — عمود E في شيت «بيانات_الطلاب»  ▼▼▼
 *  أُلحق بنهاية الملف — لا يعدّل أي كود أصلي بالأعلى.
 *  هذا الرقم هو نفسه رقم واتساب الذي يُرسَل عليه تلقائياً.
 * =================================================================== */

var PHONE_HEADER = 'رقم التليفون';
var PHONE_COL    = 5; // العمود E (بعد: A رقم، B اسم، C صف، D فصل)

// التأكد من وجود عنوان عمود الهاتف
function ensurePhoneHeader(sheet) {
  var h = String(sheet.getRange(1, PHONE_COL).getValue()).trim();
  if (h === '') sheet.getRange(1, PHONE_COL).setValue(PHONE_HEADER);
}

/**
 * setStudentPhone — حفظ/تحديث رقم تليفون الطالب في العمود E.
 * المعطيات: id ، phone
 */
function setStudentPhone(params) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var id    = String((params && params.id)    || '').trim();
    var phone = String((params && params.phone) || '').trim();
    if (!id) return { ok: false, error: 'رقم الطالب مطلوب' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(STUDENTS_SHEET);
    if (!sheet) return { ok: false, error: 'تبويب الطلاب غير موجود' };

    ensurePhoneHeader(sheet);
    var row = findStudentRow(sheet, id);
    if (row === -1) return { ok: false, error: 'لم يُعثر على الطالب رقم ' + id };

    // تنسيق الخلية كنص للحفاظ على الصفر البادئ (01xxxxxxxxx)
    sheet.getRange(row, PHONE_COL).setNumberFormat('@').setValue(phone);
    return { ok: true, message: 'تم حفظ رقم التليفون ✓', phone: phone };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

/**
 * getStudentPhone — قراءة رقم تليفون الطالب من العمود E.
 * المعطيات: id
 */
function getStudentPhone(params) {
  try {
    var id = String((params && params.id) || '').trim();
    if (!id) return { ok: false, error: 'رقم الطالب مطلوب' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(STUDENTS_SHEET);
    if (!sheet) return { ok: true, phone: '' };

    var row = findStudentRow(sheet, id);
    if (row === -1) return { ok: true, phone: '' };

    var v = String(sheet.getRange(row, PHONE_COL).getValue()).trim();
    return { ok: true, phone: v };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}
