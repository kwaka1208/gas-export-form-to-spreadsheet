const panelSheetName = 'パネル'

function exportGoogleFormToSheet(formUrl, spreadsheetUrl, sheetName) {
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
  sheet.appendRow(['質問内容', '必須', '入力方法', '選択肢など']);

  let currentSectionTitle = '';
  items.forEach((item) => {
    let type = item.getType();
    const title = item.getTitle();
    var isRequired = isRequiredForItem(item)
    let choices = [];
    let inputMethod = '';
    let columnInfo = [];
    let rowInfo = [];
    let additionalInfo = '';

    switch (type) {
      case FormApp.ItemType.TEXT:
        inputMethod = 'テキストボックス';
        break;
      case FormApp.ItemType.PARAGRAPH_TEXT:
        inputMethod = 'パラグラフテキスト';
        break;
      case FormApp.ItemType.MULTIPLE_CHOICE:
        inputMethod = 'ラジオボタン';
        choices = item.asMultipleChoiceItem().getChoices().map(choice => choice.getValue());
        break;
      case FormApp.ItemType.CHECKBOX:
        inputMethod = 'チェックボックス';
        choices = item.asCheckboxItem().getChoices().map(choice => choice.getValue());
        break;
      case FormApp.ItemType.LIST:
        inputMethod = 'プルダウン';
        choices = item.asListItem().getChoices().map(choice => choice.getValue());
        break;
      case FormApp.ItemType.FILE_UPLOAD:
        inputMethod = 'ファイルアップロード';
        break;
      case FormApp.ItemType.SCALE:
        inputMethod = '均等目盛';
        const lowerBound = item.asScaleItem().getLowerBound();
        const upperBound = item.asScaleItem().getUpperBound();
        const leftLabel = item.asScaleItem().getLeftLabel(); 
        const rightLabel = item.asScaleItem().getRightLabel();
        additionalInfo = `${leftLabel ? leftLabel : '最小値のラベルなし'}：${lowerBound}, ${rightLabel ? rightLabel : '最大値のラベルなし'}：${upperBound}`;
        break;
      case FormApp.ItemType.GRID:
        inputMethod = '選択式グリッド形式';
        columnInfo = item.asGridItem().getColumns();
        rowInfo = item.asGridItem().getRows();
        break;
      case FormApp.ItemType.CHECKBOX_GRID:
        inputMethod = 'チェックボックスグリッド形式';
        columnInfo = item.asCheckboxGridItem().getColumns();
        rowInfo = item.asCheckboxGridItem().getRows();
        break;
      case FormApp.ItemType.DATE:
        inputMethod = '日付';
        break;
      case FormApp.ItemType.TIME:
        inputMethod = '時刻';
        break;
      case FormApp.ItemType.SECTION_HEADER:
        inputMethod = 'セクションヘッダー';
        break;
      default:
        inputMethod = 'その他';
    }

    // 設問、必須情報、入力方法、追加情報を出力
    const row = [title, isRequired, inputMethod, additionalInfo];
    sheet.appendRow(row);

    // グリッド形式の列情報と行情報を縦方向に追加
    if (rowInfo.length > 0) {
      sheet.appendRow(['', '', '', '[項目]']);
      rowInfo.forEach((row) => {
        sheet.appendRow(['', '', '', row]);
      });
    }
    if (columnInfo.length > 0) {
      sheet.appendRow(['', '', '', '[選択肢]']);
      columnInfo.forEach((column) => {
        sheet.appendRow(['', '', '', column]);
      });
    }

    // 選択肢を縦方向に追加（グリッド形式以外）
    if (choices.length > 0) {
      choices.forEach((choice) => {
        sheet.appendRow(['', '', '', choice]);
      });
    }
  });

  var msgComplete = 'フォームの情報がスプレッドシートに出力されました'
  SpreadsheetApp.getUi().alert(msgComplete);
  Logger.log(msgComplete + ': ' + spreadsheet.getUrl());
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('フォームエクスポート')
    .addItem('フォームからスプレッドシートへ出力', 'exportFromSheetData')
    .addToUi();
}

function exportFromSheetData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(panelSheetName);
  const formUrl = sheet.getRange('C2').getValue();
  const spreadsheetUrl = sheet.getRange('C3').getValue();
  const sheetName = sheet.getRange('C4').getValue();

  if (sheetName == panelSheetName) {
    SpreadsheetApp.getUi().alert('出力先には' + panelSheetName + '以外の名前を使ってください');
    return
  }

  if (formUrl && sheetName) {
    exportGoogleFormToSheet(formUrl, spreadsheetUrl, sheetName);
  } else {
    SpreadsheetApp.getUi().alert('フォームURLとシート名を正しく入力してください。');
  }
}

function isRequiredForItem(item) {
  var result
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
    // case FormApp.ItemType.FILE_UPLOAD:
    //   // ファイルアップロード
    //   result = item.asListItem().isRequired();
    //   break;
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
      result = false
      break;
    // 必要に応じて他のタイプも追加可能
    default:
      result = item.isRequired ? item.isRequired() : false;
  }
  return result ? 'Y' : ''
}
