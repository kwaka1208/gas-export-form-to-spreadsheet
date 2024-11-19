function extractForm(formUrl, spreadsheetUrl, sheetName) {
  const form = FormApp.openByUrl(formUrl);
  const items = form.getItems();

  // スプレッドシートを開く
  let spreadsheet;
  if (spreadsheetUrl) {
    try {
      spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
    } catch (e) {
      // スプレッドシートが見つからない場合、新規作成
      spreadsheet = SpreadsheetApp.create('Form Questions Output');
      Logger.log('新しいスプレッドシートが作成されました: ' + spreadsheet.getUrl());
    }
  } else {
    // URLが空の場合、現在開いているスプレッドシートを使用
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }

  // 指定したシートを取得（存在しない場合は新規作成）
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    // 既存のシートが存在する場合、データを消去（書式などは残す）
    sheet.clearContents();
  }

  // ヘッダー行を設定
  sheet.appendRow(COLUMN_NAME);
  sheet.appendRow(COLUMN_INDEX);

  // フォームのタイトルと説明を取得
  const formTitle = form.getTitle();
  const formDescription = form.getDescription();
  // フォームのタイトルと説明をスプレッドシートに追加
  row = Array(COLUMN_INDEX.length)
  row[COLUMN_MAPPING.QUESTION_TITLE] = formTitle;
  row[COLUMN_MAPPING.INPUT_METHOD] = FormItems.FORM_TITLE.name;
  sheet.appendRow(row);

  row = Array(COLUMN_INDEX.length);
  row[COLUMN_MAPPING.QUESTION_TITLE] = formDescription;
  row[COLUMN_MAPPING.INPUT_METHOD] = FormItems.FORM_DESCRIPTION.name;
  sheet.appendRow(row);

  let currentSectionTitle = '';
  items.forEach((item) => {
    let type = item.getType();
    const title = item.getTitle();
    const helpText = item.getHelpText();
    var isRequired = isRequiredForItem(item)
    let choices = [];
    let range = '';
    let inputMethod = '';
    let columnInfo = [];
    let rowInfo = [];

    /*
      Formのタイプから名前を取得
    */
    inputMethod = GetNameByType(type)

    switch (type) {
      case FormApp.ItemType.MULTIPLE_CHOICE:
        choices = item.asMultipleChoiceItem().getChoices().map(choice => {
          return choice.getValue();
        });
        if (item.asMultipleChoiceItem().hasOtherOption()) {
          choices.push('tag:' + tags.OTHER_OPTION)
        }
        break;
      case FormApp.ItemType.CHECKBOX:
        choices = item.asCheckboxItem().getChoices().map(choice => {
          return choice.getValue();
        });
        if (item.asCheckboxItem().hasOtherOption()) {
          choices.push('tag:' + tags.OTHER_OPTION)
        }
        break;
      case FormApp.ItemType.LIST:
        choices = item.asListItem().getChoices().map(choice => choice.getValue());
        break;
      case FormApp.ItemType.SCALE:
        const lowerBound = item.asScaleItem().getLowerBound();
        const upperBound = item.asScaleItem().getUpperBound();
        const leftLabel = item.asScaleItem().getLeftLabel(); 
        const rightLabel = item.asScaleItem().getRightLabel();
        range = `${leftLabel ? leftLabel : '最小値のラベルなし'}：${lowerBound}, ${rightLabel ? rightLabel : '最大値のラベルなし'}：${upperBound}`;
        break;
      case FormApp.ItemType.GRID:
        columnInfo = item.asGridItem().getColumns();
        rowInfo = item.asGridItem().getRows();
        break;
      case FormApp.ItemType.CHECKBOX_GRID:
        columnInfo = item.asCheckboxGridItem().getColumns();
        rowInfo = item.asCheckboxGridItem().getRows();
        break;
      /*
        以下は名前ののみなので処理の必要なし
      */
      // case FormApp.ItemType.TEXT:
      // case FormApp.ItemType.PARAGRAPH_TEXT:
      // case FormApp.ItemType.FILE_UPLOAD:
      // case FormApp.ItemType.DATE:
      // case FormApp.ItemType.TIME:
      // case FormApp.ItemType.SECTION_HEADER:
      // case FormApp.ItemType.PAGE_BREAK:
      // case FormApp.ItemType.IMAGE:
      // case FormApp.ItemType.VIDEO:
      // case FormApp.ItemType.DATE_TIME:
      // case FormApp.ItemType.DURATION:
      default:
        break;
    }

    // 設問、説明、必須情報、入力方法、追加情報、分岐情報、検証情報を出力
    row = Array(COLUMN_INDEX.length);
    row[COLUMN_MAPPING.QUESTION_TITLE]    = title;
    row[COLUMN_MAPPING.DESCRIPTION]       = helpText;
    row[COLUMN_MAPPING.REQUIRED]          = isRequired;
    row[COLUMN_MAPPING.INPUT_METHOD]      = inputMethod;
    sheet.appendRow(row);

    row = Array(COLUMN_INDEX.length);
    if (range != '') {
        row[COLUMN_MAPPING.RANGE]             = range;
        sheet.appendRow(row);
    }
    // グリッド形式の列情報と行情報を縦方向に追加
    if (rowInfo.length > 0) {
      row = Array(COLUMN_INDEX.length);
      row[COLUMN_MAPPING.RANGE]             = 'tag:' + tags.GRID_ROW;
      sheet.appendRow(row);
      rowInfo.forEach((row) => {
        let gridRow = [];
        gridRow[COLUMN_MAPPING.RANGE] = row;
        sheet.appendRow(gridRow);
      });
    }
    if (columnInfo.length > 0) {
      row = Array(COLUMN_INDEX.length);
      row[COLUMN_MAPPING.RANGE]             = 'tag:' + tags.GRID_COLUMN;
      sheet.appendRow(row);
      columnInfo.forEach((column) => {
        let gridColumn = [];
        gridColumn[COLUMN_MAPPING.RANGE] = column;
        sheet.appendRow(gridColumn);
      });
    }

    // 選択肢を縦方向に追加（グリッド形式以外）
    if (choices.length > 0) {
      choices.forEach((choice) => {
        let choiceRow = [];
        choiceRow[COLUMN_MAPPING.RANGE] = choice;
        sheet.appendRow(choiceRow);
      });
    }
  });

  var msgComplete = 'フォームの情報がスプレッドシートに出力されました';
  SpreadsheetApp.getUi().alert(msgComplete);
  Logger.log(msgComplete + ': ' + spreadsheet.getUrl());
}

function isRequiredForItem(item) {
  var result;
  switch (item.getType()) {
    case FormApp.ItemType.TEXT:
      // テキスト
      result = item.asTextItem().isRequired();
      break;
    case FormApp.ItemType.PARAGRAPH_TEXT:
      // パラグラフテキスト
      result = item.asParagraphTextItem().isRequired();
      break;
    case FormApp.ItemType.MULTIPLE_CHOICE:
      // ラジオボタン
      result = item.asMultipleChoiceItem().isRequired();
      break;
    case FormApp.ItemType.CHECKBOX:
      // チェックボックス
      result = item.asCheckboxItem().isRequired();
      break;
    case FormApp.ItemType.LIST:
      // プルダウン
      result = item.asListItem().isRequired();
      break;
    case FormApp.ItemType.SCALE:
      result = item.asScaleItem().isRequired();
      break;
    case FormApp.ItemType.GRID:
      result = item.asGridItem().isRequired();
      break;
    case FormApp.ItemType.CHECKBOX_GRID:
      result = item.asCheckboxGridItem().isRequired();
      break;
    case FormApp.ItemType.DATE:
      result = item.asDateItem().isRequired();
      break;
    case FormApp.ItemType.TIME:
      result = item.asTimeItem().isRequired();
      break;
    case FormApp.ItemType.DATE_TIME:
      result = item.asDateTimeItem().isRequired();
      break;
    case FormApp.ItemType.SECTION_HEADER:
      result = false;
      break;
    case FormApp.ItemType.DURATION:
      result = item.asDurationItem().isRequired();
      break;
    // 必要に応じて他のタイプも追加可能
    default:
      result = item.isRequired ? item.isRequired() : false;
  }
  return result ? 'Y' : '';
}

function mainExtractForm() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(panelSheetName);
  const formUrl = sheet.getRange('C2').getValue();
  const spreadsheetUrl = sheet.getRange('C3').getValue();
  const sheetName = sheet.getRange('C4').getValue();

  if (sheetName == panelSheetName) {
    SpreadsheetApp.getUi().alert('出力先には' + panelSheetName + '以外の名前を使ってください');
    return;
  }

  if (formUrl && sheetName) {
    extractForm(formUrl, spreadsheetUrl, sheetName);
  } else {
    SpreadsheetApp.getUi().alert('フォームURLとシート名を正しく入力してください。');
  }
}
