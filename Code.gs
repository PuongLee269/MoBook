var API_VERSION = 4;
var TIME_ZONE = "Asia/Bangkok";
var RESPONSE_SHEET = "Responses";
var CLASS_SHEET = "Classes";

function doGet(e) {
  try {
    var action = e && e.parameter ? String(e.parameter.action || "") : "";

    if (action === "config") {
      return jsonOutput_({
        success: true,
        version: API_VERSION,
        classes: getClassConfigs_()
      });
    }

    if (action === "eligibility") {
      var name = String(e.parameter.name || "").trim();
      var classCode = normalizeClassCode_(e.parameter.classCode);
      return jsonOutput_(checkEligibility_(name, classCode));
    }

    var records = getResponseObjects_();

    if (action === "results") {
      var resultClass = normalizeClassCode_(e.parameter.classCode);
      var config = findClassConfig_(resultClass);
      if (!config) {
        return jsonOutput_({success: false, version: API_VERSION, message: "Mã lớp không tồn tại."});
      }
      var period = getPeriod_(config.frequency, new Date());
      return jsonOutput_({
        success: true,
        version: API_VERSION,
        classCode: resultClass,
        periodKey: period.key,
        periodLabel: period.label,
        topSlots: buildTopSlots_(records, resultClass, period.key)
      });
    }

    return jsonOutput_(records);
  } catch (error) {
    return jsonOutput_({success: false, version: API_VERSION, message: error.message});
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents || "{}");
    var action = String(data.action || "submit");

    if (action === "saveConfig") {
      verifyAdminKey_(data.adminKey);
      var classes = saveClassConfigs_(data.classes || []);
      return jsonOutput_({success: true, version: API_VERSION, classes: classes});
    }

    if (action === "resetResponse") {
      return resetResponse_(data);
    }

    if (action === "adminDeleteResponse") {
      verifyAdminKey_(data.adminKey);
      return adminDeleteResponse_(data);
    }

    if (action === "resetClassVotes") {
      verifyAdminKey_(data.adminKey);
      return resetClassVotes_(data);
    }

    if (action !== "submit") {
      return jsonOutput_({success: false, version: API_VERSION, message: "Hành động không hợp lệ."});
    }

    return submitResponse_(data);
  } catch (error) {
    return jsonOutput_({success: false, version: API_VERSION, message: error.message});
  }
}

function submitResponse_(data) {
  var name = String(data.name || "").trim();
  var classCode = normalizeClassCode_(data.classCode);

  if (!name) {
    return jsonOutput_({success: false, version: API_VERSION, message: "Vui lòng nhập tên."});
  }

  var config = findClassConfig_(classCode);
  if (!config || config.active === false) {
    return jsonOutput_({success: false, version: API_VERSION, message: "Mã lớp không hợp lệ hoặc đã ngừng hoạt động."});
  }

  var availability = Array.isArray(data.availability) ? data.availability : [];
  if (!availability.length) {
    return jsonOutput_({success: false, version: API_VERSION, message: "Vui lòng chọn ít nhất một buổi rảnh."});
  }

  var period = getPeriod_(config.frequency, new Date());
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    var records = getResponseObjects_();
    var normalizedName = normalizeName_(name);
    var duplicate = records.some(function (item) {
      return normalizeName_(item.Name) === normalizedName
        && normalizeClassCode_(item.ClassCode) === classCode
        && String(item.PeriodKey || "") === period.key;
    });

    if (duplicate) {
      return jsonOutput_({
        success: false,
        version: API_VERSION,
        code: "DUPLICATE",
        message: config.frequency === "monthly"
          ? "Bạn đã điền lịch cho lớp này trong tháng hiện tại."
          : "Bạn đã điền lịch cho lớp này trong tuần hiện tại."
      });
    }

    var responseId = Utilities.getUuid();
    var resetToken = Utilities.getUuid();

    appendResponse_({
      Timestamp: new Date(),
      ResponseId: responseId,
      ResetToken: resetToken,
      Name: name,
      Morning: numberInRange_(data.morning, 1, 5),
      Afternoon: numberInRange_(data.afternoon, 1, 5),
      Evening: numberInRange_(data.evening, 1, 5),
      Availability: JSON.stringify(availability),
      ClassCode: classCode,
      Frequency: config.frequency,
      PeriodKey: period.key
    });

    records = getResponseObjects_();

    return jsonOutput_({
      success: true,
      version: API_VERSION,
      responseId: responseId,
      resetToken: resetToken,
      classCode: classCode,
      periodKey: period.key,
      periodLabel: period.label,
      topSlots: buildTopSlots_(records, classCode, period.key)
    });
  } finally {
    lock.releaseLock();
  }
}

function resetResponse_(data) {
  var responseId = String(data.responseId || "");
  var resetToken = String(data.resetToken || "");
  var name = String(data.name || "").trim();
  var classCode = normalizeClassCode_(data.classCode);
  var periodKey = String(data.periodKey || "");

  if (!periodKey) {
    var config = findClassConfig_(classCode);
    if (config) periodKey = getPeriod_(config.frequency, new Date()).key;
  }

  if ((!responseId || !resetToken) && (!name || !classCode || !periodKey)) {
    return jsonOutput_({success: false, version: API_VERSION, message: "Thiếu thông tin để xóa lượt vừa gửi."});
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    var requiredHeaders = ["Timestamp", "Name", "Morning", "Afternoon", "Evening", "Availability", "ClassCode", "Frequency", "PeriodKey", "ResponseId", "ResetToken"];
    var sheet = getOrCreateSheet_(RESPONSE_SHEET, requiredHeaders);
    var headers = ensureHeaders_(sheet, requiredHeaders);
    var values = sheet.getDataRange().getValues();
    var responseIdIndex = headers.indexOf("ResponseId");
    var resetTokenIndex = headers.indexOf("ResetToken");
    var nameIndex = headers.indexOf("Name");
    var classIndex = headers.indexOf("ClassCode");
    var periodIndex = headers.indexOf("PeriodKey");

    for (var i = values.length - 1; i >= 1; i--) {
      var exactMatch = responseId && resetToken
        && String(values[i][responseIdIndex] || "") === responseId
        && String(values[i][resetTokenIndex] || "") === resetToken;

      var legacyMatch = !responseId && !resetToken
        && normalizeName_(values[i][nameIndex]) === normalizeName_(name)
        && normalizeClassCode_(values[i][classIndex]) === classCode
        && String(values[i][periodIndex] || "") === periodKey;

      if (exactMatch || legacyMatch) {
        sheet.deleteRow(i + 1);
        return jsonOutput_({success: true, version: API_VERSION, deleted: true});
      }
    }

    return jsonOutput_({success: false, version: API_VERSION, message: "Không tìm thấy lượt vừa gửi để đặt lại."});
  } finally {
    lock.releaseLock();
  }
}

function adminDeleteResponse_(data) {
  var rowNumber = Math.floor(Number(data.rowNumber));
  if (!rowNumber || rowNumber < 2) {
    return jsonOutput_({success: false, version: API_VERSION, message: "Dòng dữ liệu không hợp lệ."});
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getOrCreateSheet_(RESPONSE_SHEET, ["Timestamp", "Name", "ClassCode"]);
    if (rowNumber > sheet.getLastRow()) {
      return jsonOutput_({success: false, version: API_VERSION, message: "Lượt vote không còn tồn tại. Hãy làm mới trang."});
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    var row = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
    var nameIndex = headers.indexOf("Name");
    var classIndex = headers.indexOf("ClassCode");

    if (data.name && normalizeName_(row[nameIndex]) !== normalizeName_(data.name)) {
      return jsonOutput_({success: false, version: API_VERSION, message: "Dữ liệu đã thay đổi. Hãy làm mới trang."});
    }
    if (data.classCode && normalizeClassCode_(row[classIndex]) !== normalizeClassCode_(data.classCode)) {
      return jsonOutput_({success: false, version: API_VERSION, message: "Dữ liệu đã thay đổi. Hãy làm mới trang."});
    }

    sheet.deleteRow(rowNumber);
    return jsonOutput_({success: true, version: API_VERSION, deleted: 1});
  } finally {
    lock.releaseLock();
  }
}

function resetClassVotes_(data) {
  var classCode = normalizeClassCode_(data.classCode);
  if (!classCode) {
    return jsonOutput_({success: false, version: API_VERSION, message: "Mã lớp không hợp lệ."});
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getOrCreateSheet_(RESPONSE_SHEET, ["Timestamp", "Name", "ClassCode"]);
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      return jsonOutput_({success: true, version: API_VERSION, deleted: 0});
    }

    var headers = values[0].map(String);
    var classIndex = headers.indexOf("ClassCode");
    if (classIndex < 0) {
      return jsonOutput_({success: true, version: API_VERSION, deleted: 0});
    }

    var deleted = 0;
    for (var i = values.length - 1; i >= 1; i--) {
      if (normalizeClassCode_(values[i][classIndex]) === classCode) {
        sheet.deleteRow(i + 1);
        deleted++;
      }
    }

    return jsonOutput_({success: true, version: API_VERSION, deleted: deleted, classCode: classCode});
  } finally {
    lock.releaseLock();
  }
}

function checkEligibility_(name, classCode) {
  if (!name) {
    return {success: false, version: API_VERSION, allowed: false, message: "Vui lòng nhập tên."};
  }

  var config = findClassConfig_(classCode);
  if (!config || config.active === false) {
    return {success: false, version: API_VERSION, allowed: false, message: "Mã lớp không hợp lệ hoặc đã ngừng hoạt động."};
  }

  var period = getPeriod_(config.frequency, new Date());
  var normalizedName = normalizeName_(name);
  var duplicate = getResponseObjects_().some(function (item) {
    return normalizeName_(item.Name) === normalizedName
      && normalizeClassCode_(item.ClassCode) === classCode
      && String(item.PeriodKey || "") === period.key;
  });

  return {
    success: true,
    version: API_VERSION,
    allowed: !duplicate,
    classCode: classCode,
    periodKey: period.key,
    periodLabel: period.label,
    message: duplicate
      ? (config.frequency === "monthly"
        ? "Bạn đã điền lịch cho lớp này trong tháng hiện tại."
        : "Bạn đã điền lịch cho lớp này trong tuần hiện tại.")
      : ""
  };
}

function getClassConfigs_() {
  var sheet = getOrCreateSheet_(CLASS_SHEET, ["Code", "Frequency", "Active", "UpdatedAt"]);
  var values = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < values.length; i++) {
    var code = normalizeClassCode_(values[i][0]);
    if (!code) continue;
    result.push({
      code: code,
      frequency: values[i][1] === "monthly" ? "monthly" : "weekly",
      active: values[i][2] === "" ? true : parseBoolean_(values[i][2]),
      updatedAt: values[i][3] || ""
    });
  }

  return result;
}

function saveClassConfigs_(classes) {
  if (!Array.isArray(classes)) throw new Error("Danh sách mã lớp không hợp lệ.");

  var seen = {};
  var cleaned = classes.map(function (item) {
    var code = normalizeClassCode_(item.code);
    if (!/^[A-Z0-9_-]{1,20}$/.test(code)) {
      throw new Error("Mã lớp không hợp lệ: " + code);
    }
    if (seen[code]) throw new Error("Mã lớp bị trùng: " + code);
    seen[code] = true;
    return {
      code: code,
      frequency: item.frequency === "monthly" ? "monthly" : "weekly",
      active: item.active !== false
    };
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getOrCreateSheet_(CLASS_SHEET, ["Code", "Frequency", "Active", "UpdatedAt"]);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 4).setValues([["Code", "Frequency", "Active", "UpdatedAt"]]);
    if (cleaned.length) {
      var now = new Date();
      var rows = cleaned.map(function (item) {
        return [item.code, item.frequency, item.active, now];
      });
      sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    }
  } finally {
    lock.releaseLock();
  }

  return getClassConfigs_();
}

function findClassConfig_(classCode) {
  var code = normalizeClassCode_(classCode);
  var classes = getClassConfigs_();
  for (var i = 0; i < classes.length; i++) {
    if (classes[i].code === code) return classes[i];
  }
  return null;
}

function getResponseObjects_() {
  var headers = ["Timestamp", "Name", "Morning", "Afternoon", "Evening", "Availability", "ClassCode", "Frequency", "PeriodKey", "ResponseId", "ResetToken"];
  var sheet = getOrCreateSheet_(RESPONSE_SHEET, headers);
  ensureHeaders_(sheet, headers);

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  var actualHeaders = values[0].map(function (header) { return String(header); });
  var result = [];

  for (var i = 1; i < values.length; i++) {
    if (values[i].join("") === "") continue;
    var obj = {};
    for (var j = 0; j < actualHeaders.length; j++) {
      obj[actualHeaders[j]] = values[i][j];
    }
    obj._row = i + 1;
    result.push(obj);
  }

  return result;
}

function appendResponse_(record) {
  var requiredHeaders = ["Timestamp", "Name", "Morning", "Afternoon", "Evening", "Availability", "ClassCode", "Frequency", "PeriodKey", "ResponseId", "ResetToken"];
  var sheet = getOrCreateSheet_(RESPONSE_SHEET, requiredHeaders);
  var headers = ensureHeaders_(sheet, requiredHeaders);
  var row = headers.map(function (header) {
    return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : "";
  });
  sheet.appendRow(row);
}

function getOrCreateSheet_(name, headers) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function ensureHeaders_(sheet, requiredHeaders) {
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function (header) {
    return String(header);
  });

  requiredHeaders.forEach(function (header) {
    if (headers.indexOf(header) === -1) {
      headers.push(header);
      sheet.getRange(1, headers.length).setValue(header);
    }
  });

  return headers;
}

function buildTopSlots_(records, classCode, periodKey) {
  var counter = {};

  records.forEach(function (item) {
    if (normalizeClassCode_(item.ClassCode) !== classCode) return;
    if (String(item.PeriodKey || "") !== periodKey) return;

    var slots = item.Availability || [];
    if (typeof slots === "string") {
      try {
        slots = JSON.parse(slots);
      } catch (error) {
        slots = [];
      }
    }

    if (!Array.isArray(slots)) return;
    slots.forEach(function (slot) {
      var key = String(slot.day || "") + " " + String(slot.period || "");
      key = key.trim();
      if (key) counter[key] = (counter[key] || 0) + 1;
    });
  });

  return Object.keys(counter).map(function (key) {
    return {slot: key, votes: counter[key]};
  }).sort(function (a, b) {
    if (b.votes !== a.votes) return b.votes - a.votes;
    return a.slot.localeCompare(b.slot);
  }).slice(0, 5);
}

function getPeriod_(frequency, date) {
  var dateText = Utilities.formatDate(date, TIME_ZONE, "yyyy-MM-dd");
  var parts = dateText.split("-").map(Number);

  if (frequency === "monthly") {
    return {
      key: parts[0] + "-M" + pad2_(parts[1]),
      label: "tháng " + pad2_(parts[1]) + "/" + parts[0]
    };
  }

  var current = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  var day = current.getUTCDay() || 7;
  current.setUTCDate(current.getUTCDate() + 4 - day);
  var weekYear = current.getUTCFullYear();
  var yearStart = new Date(Date.UTC(weekYear, 0, 1));
  var week = Math.ceil((((current - yearStart) / 86400000) + 1) / 7);

  return {
    key: weekYear + "-W" + pad2_(week),
    label: "tuần " + pad2_(week) + "/" + weekYear
  };
}

function verifyAdminKey_(providedKey) {
  var expectedKey = PropertiesService.getScriptProperties().getProperty("ADMIN_KEY");
  if (expectedKey && String(providedKey || "") !== expectedKey) {
    throw new Error("Khóa quản trị không đúng.");
  }
}

function normalizeName_(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function normalizeClassCode_(value) {
  return String(value || "").trim().toUpperCase();
}

function parseBoolean_(value) {
  if (typeof value === "boolean") return value;
  var text = String(value).toLowerCase();
  return text !== "false" && text !== "0" && text !== "no";
}

function numberInRange_(value, min, max) {
  var number = Number(value);
  if (!isFinite(number)) return min;
  return Math.max(min, Math.min(max, number));
}

function pad2_(value) {
  return String(value).padStart(2, "0");
}

function jsonOutput_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
