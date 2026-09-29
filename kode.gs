function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Inventaris");
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var rows = data.slice(1);
    
    var result = [];
    var totalBarang = 0;
    var kondisiBaik = 0;
    
    rows.forEach(function(row) {
      var obj = {};
      headers.forEach(function(header, index) {
        obj[header.toString().trim()] = row[index];
      });
      result.push(obj);
      
      // Hitung ringkasan sederhana
      totalBarang += Number(row[6]) || 0; // Kolom JML
      if(row[10] && row[10].toString().toLowerCase() === "baik") {
        kondisiBaik += Number(row[6]) || 0;
      }
    });
    
    var responseData = {
      status: "success",
      summary: {
        totalItems: totalBarang,
        goodCondition: kondisiBaik,
        totalRows: rows.length
      },
      data: result
    };
    
    return ContentService.createTextOutput(JSON.stringify(responseData))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
