function doGet(e) {
  var action = e.parameter.action;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventaris");
  
  if (action === "getItems") {
    var rows = sheet.getDataRange().getValues();
    var data = [];
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0]) {
        data.push({
          id: rows[i][0],
          kode: rows[i][1],
          nama: rows[i][2],
          ruangan: rows[i][3],
          kondisi: rows[i][4],
          penanggung: rows[i][5],
          tahun: rows[i][6],
          keterangan: rows[i][7],
          coordX: rows[i][8] || 50,
          coordY: rows[i][9] || 50
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success", data: data}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var json = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventaris");
    var action = json.action;
    
    if (action === "saveItem") {
      var item = json.data;
      var rows = sheet.getDataRange().getValues();
      var rowIndex = -1;
      
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] === item.id) {
          rowIndex = i + 1;
          break;
        }
      }
      
      var rowData = [item.id, item.kode, item.nama, item.ruangan, item.kondisi, item.penanggung, item.tahun, item.keterangan, item.coordX, item.coordY];
      
      if (rowIndex > 0) {
        sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
      }
      
      return ContentService.createTextOutput(JSON.stringify({status: "success"}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "deleteItem") {
      var id = json.id;
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0] === id) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({status: "success"}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
