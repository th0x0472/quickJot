function doPost(e) {

  const scriptProperties = PropertiesService.getScriptProperties();
  const expectedKey = scriptProperties.getProperty('SECRET_KEY');

  try {
    const jsonString = e.postData.contents;
    const data = JSON.parse(jsonString);

    const receivedKey = data.KEY;
    const receivedMsg = data.MSG;

    if (receivedKey !== expectedKey){
      throw new Error('Authentication error.');
    }

    // 4. 後続処理（必要に応じてここに処理を記述します）
    main(receivedMsg);

    // 成功レスポンスの返却
    const successResponse = {
      status: 'success',
      message: receivedMsg + 'を記録しました'
    };

    return ContentService.createTextOutput(JSON.stringify(successResponse))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {

    let errorResponse;

    if (error.message === 'Authentication error.'){
      errorResponse = {
        status: 401,
        error: 'Authentication failed.',
        message: 'KEYが一致しないか、指定されていません'
      };
    }else{
      errorResponse = {
        status: 400,
        error: 'Bad request.',
        message: 'リクエストの解析に失敗しました'
      }
    }

    return ContentService.createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheet(spreadsheetApp){

  const today = new Date();
  const timeZone = Session.getScriptTimeZone();
  const sheetName = Utilities.formatDate(today, timeZone, 'yyyyMM');

  let sheet = spreadsheetApp.getSheetByName(sheetName);

  if (!sheet){
    let template = spreadsheetApp.getSheetByName('template');
    sheet = template.copyTo(spreadsheetApp);
    sheet.setName(sheetName);
    spreadsheetApp.setActiveSheet(sheet);
    spreadsheetApp.moveActiveSheet(1);
  }

  return sheet;
}

function main(msg){

  const scriptProperties = PropertiesService.getScriptProperties();
  const spreadsheetUrl = scriptProperties.getProperty('SPREADSHEET_URL');

  const spreadsheetApp = SpreadsheetApp.openByUrl(spreadsheetUrl);
  const sheet = getSheet(spreadsheetApp);

  const today = new Date();
  const timeZone = Session.getScriptTimeZone();
  const date = Utilities.formatDate(today, timeZone, 'yyyy/MM/dd');
  const currentTime = Utilities.formatDate(today, timeZone, 'HH:mm');

  const targetRow = sheet.getLastRow() + 1;
  const targetRange = sheet.getRange('A' + targetRow + ':C' + targetRow);

  const newValues = [[date, currentTime, msg]];
  targetRange.setValues(newValues);

}

function test(){

  main('Hello world.')

}