/**
 * 1. ฟังก์ชันตั้งค่า (รันตัวนี้เพื่อสร้างตารางในไฟล์ปัจจุบัน)
 * กดเลือก "setupInActiveSheet" แล้วกด Run
 */
function setupInActiveSheet() {
  try {
    // ดึง Spreadsheet ที่เรากำลังเปิดรันสคริปต์อยู่
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Links Data");
    
    // ถ้ายังไม่มี Sheet ชื่อ Links Data ให้สร้างใหม่
    if (!sheet) {
      sheet = ss.insertSheet("Links Data");
    } else {
      sheet.clear(); // ล้างข้อมูลเก่าออกถ้าจะ Setup ใหม่
    }
    
    // 1. สร้างหัวตาราง
    var headers = ["ลิงก์สินค้า", "ลิงก์ร้านค้า", "ลิงก์หมวดหมู่", "วันที่", "เวลา", "โน้ต", "จำนวนรวม"];
    sheet.getRange(1, 1, 1, 7).setValues([headers]);
    
    // 2. จัดรูปแบบหัวตาราง
    sheet.getRange(1, 1, 1, 7)
         .setBackground("#2563eb")
         .setFontColor("#ffffff")
         .setFontWeight("bold")
         .setHorizontalAlignment("center");
    
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 400);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 250);

    // 3. ใส่ข้อมูลตัวอย่าง (Dummy Data)
    var dummy = [
      [
        "https://sample.link/product-1\nhttps://sample.link/product-2", 
        "https://sample.link/shop", 
        "https://sample.link/category", 
        Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy"), 
        Utilities.formatDate(new Date(), "Asia/Bangkok", "HH:mm:ss"), 
        "ข้อมูลตัวอย่าง: ระบบเชื่อมต่อสำเร็จ", 
        "2"
      ]
    ];
    sheet.getRange(2, 1, 1, 7).setValues(dummy);
    
    // จัดขอบและตำแหน่ง
    sheet.getRange(2, 1, 1, 7).setVerticalAlignment("middle").setWrap(true);

    SpreadsheetApp.getUi().alert("✅ สำเร็จ! สร้างตารางและข้อมูลตัวอย่างในไฟล์นี้เรียบร้อยแล้ว");
    
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ ข้อผิดพลาด: " + e.toString());
  }
}

/**
 * 2. ฟังก์ชันรับข้อมูลจาก HTML (GitHub Pages)
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Links Data") || ss.getSheets()[0];
    
    var now = new Date();
    var grouped = groupLinks(data.links || []);
    
    var rowData = [
      grouped.p.join("\n"), 
      grouped.s, 
      grouped.c, 
      Utilities.formatDate(now, "Asia/Bangkok", "dd/MM/yyyy"),
      Utilities.formatDate(now, "Asia/Bangkok", "HH:mm:ss"),
      data.note || "",
      (data.links || []).length
    ];

    sheet.appendRow(rowData);
    
    // จัดรูปแบบแถวใหม่ที่เพิ่มเข้ามา
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, 7).setVerticalAlignment("middle").setWrap(true);
    
    return ContentService.createTextOutput(JSON.stringify({"status": "success"}))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": err.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ฟังก์ชันช่วยแยกลิงก์
 */
function groupLinks(links) {
  var p = [], s = "", c = "";
  links.forEach(function(l, i) {
    if (i === 0) p.push(l);
    else if (i === 1) s = l;
    else if (i === 2) c = l;
    else p.push(l);
  });
  return {p: p, s: s, c: c};
}

function doGet(e) {
  return ContentService.createTextOutput("Status: Online");
}
