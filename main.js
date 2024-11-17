const panelSheetName = 'パネル'

const COLUMNS = {
  'GNO':          'No.',
  'GNAME':        'グループ名',
  'QNO':          'No.',
  'METHOD':       '入力方法',
  'REQ':          '必須',
  'QTITLE':       '設問内容',
  'DESCRIPTION':  '説明',
  'INO':          'No.',
  'RANGE':        '選択肢・入力範囲',
  'MEMO':         'メモ',
  'DISPLAY':      '表示形式',
}

const COLUMN_INDEX = Object.keys(COLUMNS);
const COLUMN_NAME = Object.values(COLUMNS)
const MAPPING_KEYS = {
  QUESTION_TITLE:   'QTITLE',
  DESCRIPTION:      'DESCRIPTION',
  REQUIRED:         'REQ',
  INPUT_METHOD:     'METHOD',
  RANGE:            'RANGE',
};

// 動的にマッピングを生成
const COLUMN_MAPPING = Object.fromEntries(
  Object.entries(MAPPING_KEYS).map(([key, column]) => [key, COLUMN_INDEX.indexOf(column)])
);


/**
 *  シートオープン時のメニュー追加
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('フォームエクスポート')
      .addItem('フォームからスプレッドシートへ出力', 'mainExtractForm')
      .addToUi();
  }
  