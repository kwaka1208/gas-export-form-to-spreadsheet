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

  const FormItems = {
    "TEXT": {
      "type": FormApp.ItemType.TEXT, 
      "name": "テキスト", 
    },
    "PARAGRAPH_TEXT": {
      "type": FormApp.ItemType.PARAGRAPH_TEXT, 
      "name": "段落テキスト", 
    },
    "MULTIPLE_CHOICE": {
      "type": FormApp.ItemType.MULTIPLE_CHOICE, 
      "name": "ラジオボタン", 
    },
    "CHECKBOX": {
      "type": FormApp.ItemType.CHECKBOX, 
      "name": "チェックボックス", 
    },
    "LIST": {
      "type": FormApp.ItemType.LIST, 
      "name": "プルダウン", 
    },
    "FILE_UPLOAD": {
      "type": FormApp.ItemType.FILE_UPLOAD, 
      "name": "ファイルアップロード", 
    },
    "SCALE": {
      "type": FormApp.ItemType.SCALE, 
      "name": "均等目盛", 
    },
    "GRID": {
      "type": FormApp.ItemType.GRID, 
      "name": "選択式グリッド", 
    },
    "CHECKBOX_GRID": {
      "type": FormApp.ItemType.CHECKBOX_GRID, 
      "name": "チェックボックスグリッド", 
    },
    "DATE": {
      "type": FormApp.ItemType.DATE, 
      "name": "日付", 
    },
    "TIME": {
      "type": FormApp.ItemType.TIME, 
      "name": "時刻", 
    },
    "DATE_TIME": {
      "type": FormApp.ItemType.DATE_TIME, 
      "name": "日付時刻", 
    },
    "PAGE_BREAK": {
      "type": FormApp.ItemType.PAGE_BREAK, 
      "name": "セクション", 
    },
    "SECTION_HEADER": {
      "type": FormApp.ItemType.SECTION_HEADER, 
      "name": "タイトルと説明", 
    },
    "IMAGE": {
      "type": FormApp.ItemType.IMAGE, 
      "name": "画像", 
    },
    "VIDEO": {
      "type": FormApp.ItemType.VIDEO, 
      "name": "動画", 
    },
    "FORM_TITLE": {
      "type": 0, 
      "name": "フォームタイトル", 
    },
    "FORM_DESCRIPTION": {
      "type": 0, 
      "name": "フォーム説明", 
    },
  }

const PREFIX_TAG = 'tag:'
const tags = {
  "OTHER_OPTION": "+OTHER",
  "GRID_ROW": "GRID_ROW",
  "GRID_COLUMN": "GRID_COLUMN",
}

function GetNameByType(_type) {
  for (const [key, value] of Object.entries(FormItems)) {
    if (value.type === _type) {
      return value.name; // 対応するnameを返す
    }
  }
  return null; // 該当なしの場合
}

/**
 *  シートオープン時のメニュー追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('フォームを展開', 'mainExtractForm')
    .addItem('フォームを作成', 'mainGenerateForm')
    .addToUi();
}
