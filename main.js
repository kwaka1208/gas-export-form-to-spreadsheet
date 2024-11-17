function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('フォームエクスポート')
      .addItem('フォームからスプレッドシートへ出力', 'exportFromSheetData')
      .addToUi();
  }
  