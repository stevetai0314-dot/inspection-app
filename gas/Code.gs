var SHEET_NAME = "異常回報";
var TZ = "Asia/Ho_Chi_Minh";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        "日期","時間戳記","巡檢人員","機台號",
        "帶寬左","帶寬右","勾高左","勾高右",
        "電熱左","電熱右","外觀左","外觀右"
      ]);
    }

    var today = Utilities.formatDate(new Date(), TZ, "yyyy/MM/dd");

    sheet.appendRow([
      today,
      data.timestamp       || "",
      data.inspector       || "",
      String(data.so_may   || "").trim(),
      data.do_rong_trai    || "",
      data.do_rong_phai    || "",
      data.cao_trai        || "",
      data.cao_phai        || "",
      data.nhiet_trai      || "",
      data.nhiet_phai      || "",
      data.ngoai_quan_trai || "",
      data.ngoai_quan_phai || ""
    ]);

    return jsonOut({ status: "ok" });

  } catch (err) {
    return jsonOut({ status: "error", msg: err.message });
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet || sheet.getLastRow() <= 1) return jsonOut({});

    var yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    var yStr = Utilities.formatDate(yesterday, TZ, "yyyy/MM/dd");

    var rows = sheet.getDataRange().getValues();
    var result = {};

    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      var rowDate = (r[0] && typeof r[0].getTime === 'function')
          ? Utilities.formatDate(r[0], TZ, "yyyy/MM/dd")
          : String(r[0]).trim();
      if (rowDate !== yStr) continue;

      var mNo = String(r[3]).trim();
      if (!mNo) continue;
      if (!result[mNo]) result[mNo] = [];

      if ((r[4] !== "" || r[5] !== "") && result[mNo].indexOf("帶寬") < 0) result[mNo].push("帶寬");
      if ((r[6] !== "" || r[7] !== "") && result[mNo].indexOf("勾高") < 0) result[mNo].push("勾高");
      if ((r[8] !== "" || r[9] !== "") && result[mNo].indexOf("電熱") < 0) result[mNo].push("電熱");
      if ((r[10]!== "" || r[11]!== "") && result[mNo].indexOf("外觀") < 0) result[mNo].push("外觀");
    }

    return jsonOut(result);

  } catch (err) {
    return jsonOut({});
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
