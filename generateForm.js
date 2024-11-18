// GoogleスプレッドシートからGoogleフォームを作成する別のスクリプト
function generateForm() {
    // 前回作ったスプレッドシートのURLを指定（必要に応じて変更してください）
    const sheetUrl = 'YOUR_SPREADSHEET_URL';
    const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
    const sheetName = 'YOUR_SHEET_NAME';
    const sheet = spreadsheet.getSheetByName(sheetName);
  
    if (!sheet) {
      Logger.log(`指定されたシート名「${sheetName}」が見つかりません。`);
      return;
    }
  
    // シートのデータを取得
    const data = sheet.getDataRange().getValues();
  
    // Googleフォームを新規作成
    const form = FormApp.create('新規生成されたフォーム');
  
    // シートのデータを読み込んでフォームに質問を追加
    for (let i = 2; i < data.length; i++) { // 3行目からスタート
      const questionText = data[i][3]; // D列に質問内容
      const inputType = data[i][4];    // E列に入力方法
      const isRequired = data[i][5] === 'Y'; // F列に必須か任意か（Yの場合は必須）
      const options = data[i][7] ? data[i][7].split('\n') : []; // H列に選択肢（行区切り）
      const validationInfo = {          // I, J, K, L 列に回答の検証に関する情報
        pattern: data[i][8],            // I列に正規表現パターン
        helpText: data[i][9],           // J列にヘルプテキスト
        minLength: data[i][10],         // K列に最小文字数
        maxLength: data[i][11]          // L列に最大文字数
      };
  
      let item;
      switch (inputType.toLowerCase()) {
        case 'テキスト':
          item = form.addTextItem();
          item.setTitle(questionText);
          if (validationInfo.pattern) {
            const textValidation = FormApp.createTextValidation()
              .requireTextMatchesPattern(validationInfo.pattern)
              .setHelpText(validationInfo.helpText || '入力が無効です')
              .build();
            item.setValidation(textValidation);
          }
          if (validationInfo.minLength || validationInfo.maxLength) {
            const lengthValidation = FormApp.createTextValidation();
            if (validationInfo.minLength) {
              lengthValidation.requireTextLengthGreaterThanOrEqualTo(parseInt(validationInfo.minLength));
            }
            if (validationInfo.maxLength) {
              lengthValidation.requireTextLengthLessThanOrEqualTo(parseInt(validationInfo.maxLength));
            }
            item.setValidation(lengthValidation.build());
          }
          break;
        case '段落テキスト':
          item = form.addParagraphTextItem();
          item.setTitle(questionText);
          break;
        case 'チェックボックス':
          item = form.addCheckboxItem();
          item.setTitle(questionText);
          item.setChoices(options.map(option => item.createChoice(option)));
          break;
        case 'ラジオボタン':
          item = form.addMultipleChoiceItem();
          item.setTitle(questionText);
          item.setChoices(options.map(option => item.createChoice(option)));
          break;
        case 'ドロップダウン':
          item = form.addListItem();
          item.setTitle(questionText);
          item.setChoices(options.map(option => item.createChoice(option)));
          break;
        case '日付':
          item = form.addDateItem();
          item.setTitle(questionText);
          break;
        case '時刻':
          item = form.addTimeItem();
          item.setTitle(questionText);
          break;
        case '日付時刻':
          item = form.addDateTimeItem();
          item.setTitle(questionText);
          break;
        case 'セクション':
          item = form.addSectionHeaderItem();
          item.setTitle(questionText);
          break;
        default:
          Logger.log(`未対応の入力方法: ${inputType}`);
      }
  
      if (item && isRequired && inputType.toLowerCase() !== 'セクションヘッダー') {
        item.setRequired(true);
      }
    }
  
    Logger.log('フォームが作成されました: ' + form.getEditUrl());
  }
  
  // 注意: スプレッドシートのURLとシート名を適切に設定してください。
  