function doPost(e) {
  try {
    // รับข้อมูล JSON
    var data = JSON.parse(e.postData.contents);
    var links = data.links || [];
    var note = data.note || "";

    // เปิดหรือสร้าง Spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ถ้ายังไม่มี Spreadsheet (รันเป็นครั้งแรก)
    if (!ss) {
      // สร้าง Spreadsheet ใหม่
      ss = SpreadsheetApp.create("GET LINK - ระบบบันทึกลิงก์");
      var sheet = ss.getActiveSheet();
      sheet.setName("Links Data");

      // สร้างหัวตาราง
      setupSheet(sheet);

      // บันทึก URL ของ Spreadsheet ไว้ใน Properties
      var props = PropertiesService.getScriptProperties();
      props.setProperty("SPREADSHEET_ID", ss.getId());

      // แสดง URL ให้ผู้ใช้รู้
      Logger.log("Spreadsheet URL: " + ss.getUrl());
    }

    var sheet = ss.getSheetByName("Links Data");
    if (!sheet) {
      sheet = ss.insertSheet("Links Data");
      setupSheet(sheet);
    }

    // เวลาปัจจุบัน (ไทย)
    var now = new Date();
    var thaiTime = Utilities.formatDate(now, "Asia/Bangkok", "HH:mm:ss");
    var thaiDate = Utilities.formatDate(now, "Asia/Bangkok", "dd/MM/yyyy");

    // เพิ่มข้อมูลแต่ละลิงก์ (แยกแถว)
    for (var i = 0; i < links.length; i++) {
      var rowData = [
        links[i],        // คอลัมน์ A: ลิงก์
        thaiDate,        // คอลัมน์ B: วันที่
        thaiTime,        // คอลัมน์ C: เวลา
        note             // คอลัมน์ D: โน้ต
      ];

      sheet.appendRow(rowData);
    }

    // จัดรูปแบบแถวล่าสุด
    formatLastRows(sheet, links.length);

    return ContentService.createTextOutput(JSON.stringify({
      "status": "success",
      "message": "บันทึกสำเร็จ " + links.length + " ลิงก์",
      "spreadsheetUrl": ss.getUrl()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error",
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ฟังก์ชันสร้างโครงสร้างตารางเริ่มต้น
function setupSheet(sheet) {
  // ล้างข้อมูลเดิม
  sheet.clear();

  // หัวตาราง
  var headers = ["ลิงก์ (Link)", "วันที่ (Date)", "เวลา (Time)", "โน้ต (Note)"];
  sheet.appendRow(headers);

  // จัดรูปแบบหัวตาราง
  var headerRange = sheet.getRange(1, 1, 1, 4);
  headerRange.setBackground("#2563eb");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(12);
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");

  // ตั้งค่าความกว้างคอลัมน์
  sheet.setColumnWidth(1, 400);  // ลิงก์
  sheet.setColumnWidth(2, 120);  // วันที่
  sheet.setColumnWidth(3, 100);  // เวลา
  sheet.setColumnWidth(4, 250);  // โน้ต

  // ตั้งค่าความสูงแถวหัวตาราง
  sheet.setRowHeight(1, 35);

  // ตรึงแถวหัวตาราง
  sheet.setFrozenRows(1);

  // เปิดใช้งานการห่อข้อความอัตโนมัติ
  headerRange.setWrap(true);

  // ตั้งค่าการกรอง
  var filterRange = sheet.getRange(1, 1, 1, 4);
  filterRange.createFilter();

  // ตั้งชื่อ Spreadsheet
  SpreadsheetApp.getActiveSpreadsheet().rename("GET LINK - ระบบบันทึกลิงก์");
}

// ฟังก์ชันจัดรูปแบบแถวล่าสุด
function formatLastRows(sheet, count) {
  var lastRow = sheet.getLastRow();
  var startRow = lastRow - count + 1;

  if (startRow < 2) startRow = 2;

  var range = sheet.getRange(startRow, 1, count, 4);

  // สลับสีพื้นหลัง (zebra striping)
  for (var i = 0; i < count; i++) {
    var row = sheet.getRange(startRow + i, 1, 1, 4);
    if ((startRow + i) % 2 === 0) {
      row.setBackground("#f1f5f9");
    } else {
      row.setBackground("#ffffff");
    }

    // จัดรูปแบบข้อความ
    row.setFontSize(11);
    row.setVerticalAlignment("middle");

    // ลิงก์เป็นสีน้ำเงินและมี underline
    var linkCell = sheet.getRange(startRow + i, 1);
    linkCell.setFontColor("#2563eb");
    linkCell.setFontLine("underline");

    // จัดกึ่งกลางวันที่และเวลา
    sheet.getRange(startRow + i, 2).setHorizontalAlignment("center");
    sheet.getRange(startRow + i, 3).setHorizontalAlignment("center");
  }

  // ปรับความสูงแถวอัตโนมัติ
  sheet.autoResizeRows(startRow, count);
}

// ฟังก์ชันเริ่มต้น (รันครั้งแรกเพื่อสร้างตาราง)
function initialize() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.create("GET LINK - ระบบบันทึกลิงก์");
  }

  var sheet = ss.getSheetByName("Links Data");
  if (!sheet) {
    sheet = ss.insertSheet("Links Data");
  }

  setupSheet(sheet);

  // บันทึก ID ไว้ใช้อ้างอิง
  var props = PropertiesService.getScriptProperties();
  props.setProperty("SPREADSHEET_ID", ss.getId());

  Logger.log("✅ สร้างตารางสำเร็จ!");
  Logger.log("📎 Spreadsheet URL: " + ss.getUrl());

  // แสดงข้อความแจ้งเตือน
  SpreadsheetApp.getUi().alert(
    "สร้างตารางสำเร็จ!",
    "ระบบ GET LINK พร้อมใช้งานแล้ว\n\nSpreadsheet URL:\n" + ss.getUrl(),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ฟังก์ชันทดสอบ (สำหรับทดสอบการทำงาน)
function testSubmit() {
  var testData = {
    links: [
      "https://s.shopee.co.th/7fWSvaHhRy",
      "https://s.shopee.co.th/7fWSvaHhRy",
      "https://s.shopee.co.th/7fWSvaHhRy"
    ],
    note: "ทดสอบระบบ",
    timestamp: new Date().toISOString()
  };

  var e = {
    postData: {
      contents: JSON.stringify(testData)
    }
  };

  var result = doPost(e);
  Logger.log(result.getContent());
}

// ฟังก์ชัน doGet สำหรับตรวจสอบว่า Web App ทำงานได้
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    "status": "ok",
    "message": "GET LINK API is running",
    "version": "1.0"
  })).setMimeType(ContentService.MimeType.JSON);
}