function exportSheetDataToJson() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (i == 0) {
      const fileName = 'meta.json'
      const data = {
        "projectCode": sheets[0].getRange(1,2).getValue(),
        "projectName": sheets[0].getRange(2,2).getValue(),
        "projectDate": sheets[0].getRange(3,2).getValue(),
        "aud": sheets[0].getRange(4,2).getValue(),
        "col": sheets[0].getRange(5, 2, 1, sheets[0].getLastColumn() - 1).getValues()
      }
      const jsonFile = DriveApp.createFile(`plproto/${fileName}`, JSON.stringify(data));
      Logger.log("PLProto:Metaデータが作成されました。ファイル名: " + jsonFile.getName());
    } else {
      const lastRow = sheets[i].getLastRow();
      const lastColumn = sheets[i].getLastColumn();
      const dataRange = sheets[i].getRange(1, 1, lastRow, lastColumn);
      const dataValues = dataRange.getValues();
      const fileName = sheets[i].getName();
      const iss = sheets[i].getRange(1,2).getValue();
      const isc = sheets[i].getRange(1,3).getValue();
      let body = []
      for (let k = 2; k < lastRow; k++) {
        const rowData = dataValues[k];
        let section = [];
        // 各セルの値をオブジェクトに追加
        for (let j = 2; j < lastColumn; j++) {
          section.push(
            {
              "object": rowData[1],
              "color": sheets[i].getRange(k+1,j+1).getBackground(),
              "pointer": rowData[0],
              "content": rowData[j]
            }
          )
        }
        body.push(section);
      }
      const data = {
        "projectCode": sheets[0].getRange(1,2).getValue(),
	      "layerName": fileName.split('.')[1],
	      "layerIndex": i-1,
	      "layerNumber": `${fileName.split('.')[0]}.`,
	      "stt": sheets[i].getRange(1,4).getValue(),
	      "end": sheets[i].getRange(1,5).getValue(),
	      "aud": sheets[0].getRange(4,2).getValue(),
	      "iss": iss,
	      "isc": isc,
	      "body": body
      }
      const jsonFile = DriveApp.createFile(`plproto/layout/${fileName}.json`, JSON.stringify(data));
      Logger.log("PLProto:Layerファイルが作成されました。ファイル名: " + jsonFile.getName());
    }
  }
}
