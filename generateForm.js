// GoogleスプレッドシートからGoogleフォームを作成する別のスクリプト
function generateForm(spreadsheetUrl, sheetName) {

  // スプレッドシートを開く
  let spreadsheet
  if (spreadsheetUrl) {
    try {
      spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    } catch (e) {
      var msg = 'スプレッドシートURLとシート名を正しく入力してください。'
      Logger.log(msg);
      SpreadsheetApp.getUi().alert(msg)
      return
    }
  } else {
    // URLが空の場合、現在開いているスプレッドシートを使用
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }
// 指定したシートを取得（存在しない場合は新規作成）
let sheet = spreadsheet.getSheetByName(sheetName);
if (!sheet) {
  var msg = 'スプレッドシートURLとシート名を正しく入力してください。'
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg)
  return
}
  // シートのデータを取得
  const data = sheet.getDataRange().getValues();

  // Googleフォームを新規作成
  const form = FormApp.create('新規生成されたフォーム');

  // シートのデータを読み込んでフォームに質問を追加
  for (let i = 2; i < data.length; i++) { // 3行目からスタート
    let title       = data[i][COLUMN_MAPPING.QUESTION_TITLE]
    let description = data[i][COLUMN_MAPPING.DESCRIPTION]
    let isRequired  = data[i][COLUMN_MAPPING.REQUIRED]
    let method      = data[i][COLUMN_MAPPING.INPUT_METHOD]
    let choices = [];
    let columnInfo = [];
    let rowInfo = [];

    if (method == '') continue;
    switch (method) {
      case FormItems.TEXT.name:
        // テキスト
        item = form.addTextItem();
        item.setTitle(title);
        item.setHelpText(description);
        break;
      case FormItems.PARAGRAPH_TEXT.name:
        // 段落テキスト
        item = form.addParagraphTextItem();
        item.setTitle(title);
        item.setHelpText(description);
        break;
      case FormItems.MULTIPLE_CHOICE.name:
        // ラジオボタン
        item = form.addMultipleChoiceItem();
        item.setTitle(title);
        item.setHelpText(description);
        for (i++; data[i][COLUMN_MAPPING.INPUT_METHOD] == ''; i++) {
          var choice = data[i][COLUMN_MAPPING.RANGE];
          if (choice.startsWith(PREFIX_TAG)){
            if(choice.includes(tags.OTHER_OPTION)) {
              item.showOtherOption(true)
            }
          } else {
            choices.push(choice)
          }
        }
        console.log(choices)
        item.setChoices(choices.map(option => item.createChoice(option)));
        break;
      case FormItems.CHECKBOX.name:
        // チェックボックス
        item = form.addCheckboxItem();
        item.setTitle(title);
        item.setHelpText(description);
        for (i++; data[i][COLUMN_MAPPING.INPUT_METHOD] == ''; i++) {
          console.log(i)
          var choice = data[i][COLUMN_MAPPING.RANGE];
          if (choice.startsWith(PREFIX_TAG)){
            if(choice.includes(tags.OTHER_OPTION)) {
              item.showOtherOption(true)
            }
          } else {
            choices.push(choice)
          }
        }
        console.log(choices)
        item.setChoices(choices.map(option => item.createChoice(option)));
        break;
      case FormItems.LIST.name:
        // プルダウンメニュー
        item = form.addListItem();
        item.setTitle(title);
        item.setHelpText(description);
        for (i++; data[i][COLUMN_MAPPING.INPUT_METHOD] == ''; i++) {
          var choice = data[i][COLUMN_MAPPING.RANGE];
          choices.push(choice);
        }
        item.setChoices(choices.map(option => item.createChoice(option)));
        break;

      case FormItems.FILE_UPLOAD.name:
        // ファイルアップロード
        item = form.addPageBreakItem();
        item.setTitle(title);
        item.setHelpText(description);
        break;
      case FormItems.SCALE.name:
        // 均等目盛
        item = form.addScaleItem();
        item.setTitle(title);
        item.setHelpText(description);
        break;
      case FormItems.GRID.name:
        // 選択式グリッド
        item = form.addGridItem();
        item.setTitle(title);
        item.setHelpText(description);
        var fGrid = true
        for (i++; data[i][COLUMN_MAPPING.INPUT_METHOD] == ''; i++) {
          var choice = data[i][COLUMN_MAPPING.RANGE];
          if(choice.includes(tags.GRID_ROW)) {
            fGrid = true
            continue;
          }
          if(choice.includes(tags.GRID_COLUMN)) {
            fGrid = false
            continue
          } 
          if(fGrid) {
            rowInfo.push(choice)
          } else {
            columnInfo.push(choice)
          }
        }
        item.setRows(rowInfo);
        item.setColumns(columnInfo);
        break;
      case FormItems.CHECKBOX_GRID.name:
        // チェックボックスグリッド
        item = form.addCheckboxItem();
        item.setTitle(title);
        item.setHelpText(description);
        var fGrid = true
        for (i++; data[i][COLUMN_MAPPING.INPUT_METHOD] == ''; i++) {
          var choice = data[i][COLUMN_MAPPING.RANGE];
          if(choice.includes(tags.GRID_ROW)) {
            fGrid = true
            continue;
          }
          if(choice.includes(tags.GRID_COLUMN)) {
            fGrid = false
            continue
          } 
          if(fGrid) {
            rowInfo.push(choice)
          } else {
            columnInfo.push(choice)
          }
        }
        item.setRows(rowInfo);
        item.setColumns(columnInfo);
        break;

      case FormItems.DATE.name:
        // 日付
        item = form.addDateItem();
        item.setTitle(title);
        item.setHelpText(description);
        break;
      case FormItems.TIME.name:
          // 時刻
        item = form.addTimeItem();
        item.setTitle(title);
        item.setHelpText(description);
        break;
      case FormItems.DATE_TIME.name:
        item = form.addDateTimeItem();
        item.setTitle(title);
        item.setHelpText(description);
        break;
      
      case FormItems.PAGE_BREAK.name:
        //セクション
        item = form.addPageBreakItem();
        item.setTitle(title);
        item.setHelpText(description);
        break;
      case FormItems.SECTION_HEADER.name:
        // タイトルと説明'
        item = form.addSectionHeaderItem();
        item.setTitle(title);
        item.setHelpText(description);
        break;
      case FormItems.IMAGE.name:
          // 画像
          item = form.addImageItem();
          item.setTitle(title);
          break;
      case FormItems.VIDEO.name:
        // 動画
        item = form.addVideoItem();
        item.setTitle(title);
        break;
      case FormItems.FORM_TITLE.name:
        // フォーム説明
        form.setTitle(title);
        break;
      case FormItems.FORM_DESCRIPTION.name:
        // フォーム説明
        form.setDescription(title);
        break;
      default:
        Logger.log(`未対応の入力方法: ${method}`);
        break;
    }
    if (isRequired == 'Y') {
      item.setRequired(true);
    }
  }
  msg = 'フォームが作成されました: ' + form.getEditUrl()
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg)
}

function mainGenerateForm() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(panelSheetName);
  const spreadsheetUrl = sheet.getRange('C3').getValue();
  const sheetName = sheet.getRange('C5').getValue();
  generateForm(spreadsheetUrl, sheetName);
}
