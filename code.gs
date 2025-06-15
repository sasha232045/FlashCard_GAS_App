/**
 * 単語学習アプリ - Google Apps Script サーバーサイドコード
 * スプレッドシートを使用した高度な学習管理システム
 */

// スプレッドシートID（実際の使用時には自分のスプレッドシートIDに変更してください）
const SPREADSHEET_ID = 'your_spreadsheet_id_here';
const SHEET_NAME = '単語リスト';

/**
 * スプレッドシートIDが正しく設定されているかチェック
 */
function checkSpreadsheetIdConfiguration() {
  console.log('=== スプレッドシートID設定チェック ===');
  console.log('現在設定されているSPREADSHEET_ID: ' + SPREADSHEET_ID);
  
  if (SPREADSHEET_ID === 'your_spreadsheet_id_here') {
    console.log('⚠️ 警告: スプレッドシートIDが未設定です！');
    console.log('⚠️ これが原因で毎回新しいスプレッドシートが作成されている可能性があります');
    console.log('⚠️ 対策: 実際のスプレッドシートIDをcode.gsのSPREADSHEET_IDに設定してください');
    return false;
  } else {
    console.log('✓ スプレッドシートIDが設定されています');
    return true;
  }
}

/**
 * Webアプリのエントリーポイント
 */
function doGet() {
  console.log('アプリ開始: doGet関数が呼び出されました');
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('究極の単語学習アプリ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * スプレッドシートのインスタンスを取得
 */
function getSpreadsheet() {
  try {
    console.log('=== getSpreadsheet関数開始 ===');
    console.log('設定されたスプレッドシートID: ' + SPREADSHEET_ID);
    
    // スプレッドシートID設定チェック
    const isIdConfigured = checkSpreadsheetIdConfiguration();
    
    // IDが設定されていない場合は、新しいスプレッドシートを作成
    if (SPREADSHEET_ID === 'your_spreadsheet_id_here') {
      console.log('スプレッドシートIDが未設定のため、新しいシートを作成します');
      const timestamp = new Date().getTime();
      const newSs = SpreadsheetApp.create('単語学習アプリ_データベース_' + timestamp);
      console.log('新しいスプレッドシート作成成功');
      console.log('新しいスプレッドシートID: ' + newSs.getId());
      console.log('新しいスプレッドシートURL: ' + newSs.getUrl());
      console.log('※このIDをcode.gsの SPREADSHEET_ID に設定してください※');
      return newSs;
    }
    
    console.log('既存のスプレッドシートを開こうとしています...');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('既存スプレッドシート取得成功: ' + ss.getName());
    return ss;
    
  } catch (error) {
    console.error('既存スプレッドシート取得エラー: ' + error.toString());
    console.log('エラー回復のため新しいスプレッドシートを作成します');
    
    try {
      const timestamp = new Date().getTime();
      const newSs = SpreadsheetApp.create('単語学習アプリ_データベース_エラー回復_' + timestamp);
      console.log('エラー回復用スプレッドシート作成成功');
      console.log('回復用スプレッドシートID: ' + newSs.getId());
      console.log('回復用スプレッドシートURL: ' + newSs.getUrl());
      console.log('※このIDをcode.gsの SPREADSHEET_ID に設定してください※');
      return newSs;
    } catch (createError) {
      console.error('新しいスプレッドシート作成も失敗: ' + createError.toString());
      throw new Error('スプレッドシートの作成・取得ができません: ' + createError.toString());
    }
  }
}

/**
 * ワードシートを取得または作成
 */
function getWordSheet() {
  try {
    console.log('=== getWordSheet関数開始 ===');
    const ss = getSpreadsheet();
    console.log('スプレッドシート取得完了: ' + ss.getName());
    
    let sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      console.log('シートが存在しないため新規作成します: ' + SHEET_NAME);
      
      // まず新しいシートを作成
      sheet = ss.insertSheet(SHEET_NAME);
      console.log('新しいシート作成完了: ' + SHEET_NAME);
      
      // デフォルトシートがある場合は削除（ただし、最後のシートは削除しない）
      const sheets = ss.getSheets();
      console.log('現在のシート数: ' + sheets.length);
      
      // シートが2枚以上ある場合のみ削除を試行
      if (sheets.length > 1) {
        for (let i = 0; i < sheets.length; i++) {
          const currentSheet = sheets[i];
          const sheetName = currentSheet.getName();
          console.log('シート ' + i + ': ' + sheetName + ' (ID: ' + currentSheet.getSheetId() + ')');
          
          // 新しく作成したシートとは異なる「シート1」を探して削除
          if ((sheetName === 'シート1' || sheetName.indexOf('Sheet') === 0) && 
              currentSheet.getSheetId() !== sheet.getSheetId()) {
            console.log('不要なデフォルトシートを削除します: ' + sheetName);
            try {
              // 削除前に再度シート数を確認
              const currentSheets = ss.getSheets();
              if (currentSheets.length > 1) {
                ss.deleteSheet(currentSheet);
                console.log('デフォルトシート削除完了: ' + sheetName);
              } else {
                console.log('シートが1つしかないため削除をスキップ: ' + sheetName);
              }
            } catch (deleteError) {
              console.log('デフォルトシート削除をスキップ: ' + deleteError.toString());
            }
            break;
          }
        }
      } else {
        console.log('シートが1つしかないため、デフォルトシートの削除をスキップします');
      }
      
      // ヘッダー行を設定
      const headers = [
        'ID', '表面_テキスト', '裏面_テキスト', '備考_例文', 'セクション', 'タグ',
        '読み上げパターン', '読み上げ間隔', '次回レビュー日', '間隔', 'EF',
        'メモ1', 'メモ2', 'メモ3', 'メモ4', 'メモ5',
        'ステータス_Listening', 'ステータス_Reading', 'ステータス_Speaking', 'ステータス_Writing',
        '登録日'
      ];
      
      console.log('ヘッダー行を設定中...');
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setBackground('#4CAF50')
        .setFontColor('white').setFontWeight('bold');
      
      // 列幅を調整
      sheet.setColumnWidth(1, 80);   // ID
      sheet.setColumnWidth(2, 150);  // 表面
      sheet.setColumnWidth(3, 150);  // 裏面
      sheet.setColumnWidth(4, 200);  // 備考
      sheet.setColumnWidth(5, 100);  // セクション
      sheet.setColumnWidth(6, 100);  // タグ
      sheet.setColumnWidth(7, 120);  // 読み上げパターン
      sheet.setColumnWidth(8, 80);   // 読み上げ間隔
      
      console.log('ヘッダー行設定完了');
    } else {
      console.log('既存シートを使用: ' + SHEET_NAME);
    }
    
    console.log('=== getWordSheet関数完了 ===');
    return sheet;
    
  } catch (error) {
    console.error('=== getWordSheet関数でエラー発生 ===');
    console.error('エラー詳細: ' + error.toString());
    console.error('エラースタック: ' + (error.stack || '取得できません'));
    throw error;
  }
}

/**
 * 初期テストデータを追加
 */
function addInitialTestData(sheet) {
  console.log('初期テストデータを追加します');
  
  // 既存のデータがないことを確認
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    console.log('既にデータが存在するため、テストデータ追加をスキップします');
    return;
  }
  
  const today = new Date();
  const todayStr = today.toLocaleDateString('ja-JP');
  
  const testData = [
    [
      Date.now(), // ユニークなタイムスタンプID
      'Hello', // 表面
      'こんにちは', // 裏面
      'Hello, how are you? - 元気ですか？', // 備考・例文
      '基本会話', // セクション
      '挨拶,日常会話', // タグ
      'E2,J1,E1', // 読み上げパターン
      1.5, // 読み上げ間隔
      '', // 次回レビュー日（空白）
      1, // 間隔
      2.5, // EF
      'ヘロー', // メモ1（カタカナ読み）
      '/həˈloʊ/', // メモ2（発音記号）
      'ハの音を意識して', // メモ3（発音のコツ）
      '', // メモ4
      '', // メモ5
      '未着手', // ステータス_Listening
      '未着手', // ステータス_Reading
      '未着手', // ステータス_Speaking
      '未着手', // ステータス_Writing
      todayStr // 登録日
    ],
    [
      Date.now() + 1,
      'Thank you',
      'ありがとう',
      'Thank you very much. - どうもありがとうございます。',
      '基本会話',
      '挨拶,感謝',
      'E1,J1',
      1.0,
      '',
      1,
      2.5,
      'サンキュー',
      '/θæŋk juː/',
      'thの音は舌を軽く噛んで',
      '',
      '',
      '未着手',
      '未着手',
      '未着手',
      '未着手',
      todayStr
    ],
    [
      Date.now() + 2,
      'Good morning',
      'おはよう',
      'Good morning! Have a nice day. - おはよう！良い一日を。',
      '基本会話',
      '挨拶,朝',
      'E2,J1',
      1.2,
      '',
      1,
      2.5,
      'グッドモーニング',
      '/ɡʊd ˈmɔːrnɪŋ/',
      'モーニングの最後は軽く',
      '',
      '',
      '未着手',
      '未着手',
      '未着手',
      '未着手',
      todayStr
    ],
    [
      Date.now() + 3,
      'Beautiful',
      '美しい',
      'What a beautiful day! - なんて美しい日でしょう！',
      '形容詞',
      '形容詞,感情',
      'E3,J1',
      1.0,
      '',
      1,
      2.5,
      'ビューティフル',
      '/ˈbjuːtɪfəl/',
      'ビューの部分を強調',
      '',
      '',
      '未着手',
      '未着手',
      '未着手',
      '未着手',
      todayStr
    ],
    [
      Date.now() + 4,
      'Important',
      '重要な',
      'This is very important. - これはとても重要です。',
      '形容詞',
      '形容詞,ビジネス',
      'E2,J1,E1',
      1.3,
      '',
      1,
      2.5,
      'インポータント',
      '/ɪmˈpɔːrtənt/',
      'ポーの部分にアクセント',
      '',
      '',
      '未着手',
      '未着手',
      '未着手',
      '未着手',
      todayStr
    ]
  ];
  
  try {
    // データを一括挿入
    const range = sheet.getRange(2, 1, testData.length, testData[0].length);
    range.setValues(testData);
    
    console.log(`${testData.length}件のテストデータを追加しました`);
    
    // 追加したデータの詳細をログ出力
    testData.forEach((row, index) => {
      console.log(`テストデータ${index + 1}: ${row[1]} → ${row[2]} (ID: ${row[0]})`);
    });
    
    // セルの書式設定
    sheet.getRange(2, 1, testData.length, 1).setNumberFormat('0'); // ID列を数値フォーマット
    sheet.getRange(2, 8, testData.length, 1).setNumberFormat('0.0'); // 読み上げ間隔列
    sheet.getRange(2, 10, testData.length, 1).setNumberFormat('0'); // 間隔列
    sheet.getRange(2, 11, testData.length, 1).setNumberFormat('0.0'); // EF列
    
  } catch (error) {
    console.error('テストデータ追加エラー: ' + error.toString());
    throw new Error('テストデータの追加に失敗しました: ' + error.toString());
  }
}

/**
 * 全ての単語データを取得
 */
function getAllWords() {
  try {
    console.log('=== getAllWords関数開始 ===');
    console.log('スプレッドシートID設定値: ' + SPREADSHEET_ID);
    
    const sheet = getWordSheet();
    console.log('シート取得成功: ' + sheet.getName());
    console.log('シートのID: ' + sheet.getSheetId());
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    console.log('シートの最終行: ' + lastRow);
    console.log('シートの最終列: ' + lastColumn);
    
    // シートの基本情報を詳細確認
    const maxRows = sheet.getMaxRows();
    const maxColumns = sheet.getMaxColumns();
    console.log('シートの最大行数: ' + maxRows);
    console.log('シートの最大列数: ' + maxColumns);
    
    // ヘッダー行を確認
    if (lastRow >= 1) {
      const headerRange = sheet.getRange(1, 1, 1, Math.min(lastColumn, 21));
      const headers = headerRange.getValues()[0];
      console.log('ヘッダー行の内容:');
      headers.forEach((header, index) => {
        console.log(`  列${index + 1}: "${header}"`);
      });
    }
    
    // データ行があるかチェック
    if (lastRow >= 2) {
      console.log('データ行のサンプル確認（最初の3行まで）:');
      const sampleRows = Math.min(lastRow - 1, 3);
      for (let i = 2; i <= 1 + sampleRows; i++) {
        const rowRange = sheet.getRange(i, 1, 1, Math.min(lastColumn, 21));
        const rowData = rowRange.getValues()[0];
        console.log(`行${i}の内容:`);
        rowData.forEach((cell, index) => {
          console.log(`  列${index + 1}: "${cell}" (型: ${typeof cell})`);
        });
      }
    }
    
    if (lastRow <= 1) {
      console.log('データが存在しないため、テストデータを追加します');
      
      // テストデータを追加
      try {
        addInitialTestData(sheet);
        console.log('テストデータ追加完了、再帰呼び出しします');
        
        // テストデータ追加後に再度取得（ただし、無限ループを防ぐため1回のみ）
        const newLastRow = sheet.getLastRow();
        console.log('テストデータ追加後の最終行: ' + newLastRow);
        
        if (newLastRow <= 1) {
          console.error('テストデータ追加後もデータが存在しません');
          return [];
        }
        
      } catch (addError) {
        console.error('テストデータ追加エラー: ' + addError.toString());
        return [];
      }
    }
    
    // データを再取得
    const finalLastRow = sheet.getLastRow();
    const finalLastColumn = sheet.getLastColumn();
    console.log('最終的なデータ行数: ' + (finalLastRow - 1));
    console.log('最終的な列数: ' + finalLastColumn);
    
    if (finalLastRow <= 1) {
      console.log('最終確認でもデータが存在しません');
      return [];
    }
    
    // データ範囲を明確に指定して取得
    const dataStartRow = 2;
    const dataRowCount = finalLastRow - 1;
    const dataColumnCount = 21; // 固定で21列を取得
    
    console.log(`データ取得範囲: 行${dataStartRow}～${finalLastRow}、列1～${dataColumnCount}`);
    console.log(`取得するセル範囲: ${sheet.getRange(dataStartRow, 1, dataRowCount, dataColumnCount).getA1Notation()}`);
    
    const data = sheet.getRange(dataStartRow, 1, dataRowCount, dataColumnCount).getValues();
    console.log('取得したデータ配列サイズ: ' + data.length + ' x ' + (data.length > 0 ? data[0].length : 0));
    
    // 取得したrawデータの詳細確認
    console.log('=== 取得したrawデータの詳細 ===');
    data.forEach((row, index) => {
      console.log(`データ行${index + 1} (シート行${index + 2}):`);
      row.forEach((cell, colIndex) => {
        console.log(`  列${colIndex + 1}: "${cell}" (型: ${typeof cell})`);
      });
      console.log('  ---');
    });
    
    if (data.length === 0) {
      console.log('データ配列が空です');
      return [];
    }
    
    const words = data.map((row, index) => {
      console.log(`データ行${index + 1}処理中: ID="${row[0]}", 表面="${row[1]}", 裏面="${row[2]}"`);
      
      const today = new Date();
      const todayStr = today.toLocaleDateString('ja-JP');
      
      const word = {
        id: row[0] || (Date.now() + index),
        front: row[1] || '',
        back: row[2] || '',
        example: row[3] || '',
        section: row[4] || 'デフォルト',
        tags: row[5] || '',
        readPattern: row[6] || 'E1,J1',
        readInterval: row[7] || 1.0,
        nextReviewDate: row[8] || '',
        interval: row[9] || 1,
        ef: row[10] || 2.5,
        memo1: row[11] || '',
        memo2: row[12] || '',
        memo3: row[13] || '',
        memo4: row[14] || '',
        memo5: row[15] || '',
        statusListening: row[16] || '未着手',
        statusReading: row[17] || '未着手',
        statusSpeaking: row[18] || '未着手',
        statusWriting: row[19] || '未着手',
        registeredDate: row[20] || todayStr
      };
      
      console.log(`変換後オブジェクト: ID=${word.id}, front="${word.front}", back="${word.back}"`);
      return word;
    });
    
    console.log('=== 単語処理完了 ===');
    words.forEach((word, index) => {
      console.log(`単語${index + 1}: ID=${word.id}, ${word.front} → ${word.back}`);
      console.log(`  セクション: ${word.section}, タグ: ${word.tags}`);
      console.log(`  読み上げ: ${word.readPattern}, 間隔: ${word.readInterval}秒`);
      console.log(`  ステータス - L:${word.statusListening}, R:${word.statusReading}, S:${word.statusSpeaking}, W:${word.statusWriting}`);
      if (word.memo1 || word.memo2 || word.memo3) {
        console.log(`  メモ - 1:${word.memo1}, 2:${word.memo2}, 3:${word.memo3}`);
      }
      console.log('  ---');
    });
    
    console.log(`=== getAllWords関数完了: ${words.length}件返却 ===`);
    console.log('返却前の最終確認 - wordsの型: ' + typeof words);
    console.log('返却前の最終確認 - words.length: ' + words.length);
    console.log('返却前の最終確認 - JSON化テスト: ' + JSON.stringify(words).substring(0, 200) + '...');
    
    return words;
    
  } catch (error) {
    console.error('=== getAllWords関数でエラー発生 ===');
    console.error('エラー詳細: ' + error.toString());
    console.error('エラースタック: ' + (error.stack || '取得できません'));
    
    // エラー時は空配列を返す
    return [];
  }
}

/**
 * 新しい単語を追加
 */
function addWord(wordData) {
  console.log('新規単語追加開始: ' + wordData.front);
  const sheet = getWordSheet();
  const lastRow = sheet.getLastRow();
  
  // 新しいIDを生成（タイムスタンプベース）
  let newId;
  if (lastRow > 1) {
    const existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    newId = Math.max(...existingIds) + 1;
  } else {
    newId = Date.now(); // タイムスタンプをIDとして使用
  }
  
  const now = new Date();
  const todayStr = now.toLocaleDateString('ja-JP');
  
  const newRow = [
    newId,
    wordData.front || '',
    wordData.back || '',
    wordData.example || '',
    wordData.section || 'デフォルト',
    wordData.tags || '',
    wordData.readPattern || 'E1,J1',
    wordData.readInterval || 1.0,
    wordData.nextReviewDate || '', // 空白（今日が復習日）
    wordData.interval || 1,
    wordData.ef || 2.5,
    wordData.memo1 || '',
    wordData.memo2 || '',
    wordData.memo3 || '',
    wordData.memo4 || '',
    wordData.memo5 || '',
    wordData.statusListening || '未着手',
    wordData.statusReading || '未着手',
    wordData.statusSpeaking || '未着手',
    wordData.statusWriting || '未着手',
    todayStr
  ];
  
  sheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
  console.log('新規単語追加完了: ID=' + newId);
  return newId;
}

/**
 * 単語データを更新
 */
function updateWord(wordData) {
  console.log('単語更新開始: ID=' + wordData.id);
  const sheet = getWordSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == wordData.id) {
      const updateRow = [
        wordData.id,
        wordData.front || '',
        wordData.back || '',
        wordData.example || '',
        wordData.section || '',
        wordData.tags || '',
        wordData.readPattern || '',
        wordData.readInterval || 1,
        wordData.nextReviewDate || '',
        wordData.interval || 1,
        wordData.ef || 2.5,
        wordData.memo1 || '',
        wordData.memo2 || '',
        wordData.memo3 || '',
        wordData.memo4 || '',
        wordData.memo5 || '',
        wordData.statusListening || '未着手',
        wordData.statusReading || '未着手',
        wordData.statusSpeaking || '未着手',
        wordData.statusWriting || '未着手',
        data[i][20] // 登録日は変更しない
      ];
      
      sheet.getRange(i + 1, 1, 1, updateRow.length).setValues([updateRow]);
      console.log('単語更新完了: ID=' + wordData.id);
      return true;
    }
  }
  
  console.error('更新対象の単語が見つかりません: ID=' + wordData.id);
  return false;
}

/**
 * 単語を削除
 */
function deleteWord(wordId) {
  console.log('単語削除開始: ID=' + wordId);
  const sheet = getWordSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == wordId) {
      sheet.deleteRow(i + 1);
      console.log('単語削除完了: ID=' + wordId);
      return true;
    }
  }
  
  console.error('削除対象の単語が見つかりません: ID=' + wordId);
  return false;
}

/**
 * 学習ステータスを初回学習中に更新
 */
function updateStatusToLearning(wordId) {
  console.log('ステータス更新（学習中）: ID=' + wordId);
  const sheet = getWordSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == wordId) {
      // 未着手のステータスのみを学習中に更新
      const updates = {};
      if (data[i][16] === '未着手') updates.listening = '学習中';
      if (data[i][17] === '未着手') updates.reading = '学習中';
      if (data[i][18] === '未着手') updates.speaking = '学習中';
      if (data[i][19] === '未着手') updates.writing = '学習中';
      
      if (Object.keys(updates).length > 0) {
        sheet.getRange(i + 1, 17, 1, 4).setValues([[
          updates.listening || data[i][16],
          updates.reading || data[i][17],
          updates.speaking || data[i][18],
          updates.writing || data[i][19]
        ]]);
        console.log('ステータス更新完了: ' + JSON.stringify(updates));
      }
      return true;
    }
  }
  
  return false;
}

/**
 * SRSアルゴリズムによる次回復習日の計算
 */
function calculateNextReview(wordId, quality) {
  console.log(`SRS計算開始: ID=${wordId}, 評価=${quality}`);
  const sheet = getWordSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == wordId) {
      let interval = data[i][9] || 1;
      let ef = data[i][10] || 2.5;
      
      // SM-2アルゴリズム
      if (quality >= 3) {
        if (interval === 1) {
          interval = 6;
        } else {
          interval = Math.round(interval * ef);
        }
      } else {
        interval = 1;
      }
      
      ef = ef + (0.1 - (5 - quality) * (0.08 + (5 - quality) * 0.02));
      ef = Math.max(1.3, ef);
      
      const nextDate = new Date();
      nextDate.setDate(nextDate.getDate() + interval);
      
      sheet.getRange(i + 1, 9, 1, 3).setValues([[
        nextDate.toLocaleDateString('ja-JP'),
        interval,
        ef
      ]]);
      
      console.log(`SRS計算完了: 次回=${nextDate.toLocaleDateString('ja-JP')}, 間隔=${interval}日, EF=${ef}`);
      return true;
    }
  }
  
  return false;
}

/**
 * 今日復習すべき単語を取得
 */
function getTodayReviewWords() {
  console.log('今日の復習単語取得開始');
  const words = getAllWords();
  const today = new Date().toLocaleDateString('ja-JP');
  
  const reviewWords = words.filter(word => {
    if (!word.nextReviewDate) return true; // 初回学習
    return word.nextReviewDate <= today;
  });
  
  console.log(`今日の復習対象: ${reviewWords.length}件`);
  return reviewWords;
}

/**
 * セクション一覧を取得
 */
function getSections() {
  const words = getAllWords();
  const sections = [...new Set(words.map(word => word.section).filter(s => s))];
  console.log('セクション一覧取得: ' + sections.join(', '));
  return sections;
}

/**
 * タグ一覧を取得
 */
function getTags() {
  const words = getAllWords();
  const allTags = words.map(word => word.tags).join(',').split(',');
  const tags = [...new Set(allTags.map(tag => tag.trim()).filter(t => t))];
  console.log('タグ一覧取得: ' + tags.join(', '));
  return tags;
}

/**
 * スプレッドシート情報を取得（デバッグ用）
 */
function getSpreadsheetInfo() {
  try {
    const ss = getSpreadsheet();
    const info = {
      id: ss.getId(),
      name: ss.getName(),
      url: ss.getUrl(),
      sheets: ss.getSheets().map(sheet => ({
        name: sheet.getName(),
        rows: sheet.getLastRow(),
        columns: sheet.getLastColumn()
      }))
    };
    console.log('スプレッドシート情報: ' + JSON.stringify(info, null, 2));
    return info;
  } catch (error) {
    console.error('スプレッドシート情報取得エラー: ' + error.toString());
    return { error: error.toString() };
  }
}

/**
 * スプレッドシートの詳細情報をデバッグ出力
 */
function debugSpreadsheetInfo() {
  try {
    console.log('=== スプレッドシート詳細情報デバッグ開始 ===');
    console.log('設定されたSPREADSHEET_ID: ' + SPREADSHEET_ID);
    console.log('設定されたSHEET_NAME: ' + SHEET_NAME);
    
    // スプレッドシート基本情報
    const ss = getSpreadsheet();
    console.log('スプレッドシート名: ' + ss.getName());
    console.log('スプレッドシートID: ' + ss.getId());
    console.log('スプレッドシートURL: ' + ss.getUrl());
    
    // 全シート情報
    const sheets = ss.getSheets();
    console.log('シート総数: ' + sheets.length);
    sheets.forEach((sheet, index) => {
      console.log(`シート${index + 1}: 名前="${sheet.getName()}", ID=${sheet.getSheetId()}`);
      console.log(`  最大行数: ${sheet.getMaxRows()}, 最大列数: ${sheet.getMaxColumns()}`);
      console.log(`  最終行: ${sheet.getLastRow()}, 最終列: ${sheet.getLastColumn()}`);
    });
    
    // ターゲットシートの詳細情報
    const targetSheet = ss.getSheetByName(SHEET_NAME);
    if (targetSheet) {
      console.log('=== ターゲットシート詳細 ===');
      console.log('シート名: ' + targetSheet.getName());
      console.log('シートID: ' + targetSheet.getSheetId());
      console.log('最終行: ' + targetSheet.getLastRow());
      console.log('最終列: ' + targetSheet.getLastColumn());
      
      // 実際のデータ範囲をチェック
      const dataRange = targetSheet.getDataRange();
      console.log('データ範囲: ' + dataRange.getA1Notation());
      console.log('データ範囲の行数: ' + dataRange.getNumRows());
      console.log('データ範囲の列数: ' + dataRange.getNumColumns());
      
      // 少しデータを実際に読んでみる
      if (targetSheet.getLastRow() >= 1) {
        console.log('=== 実際のデータサンプル ===');
        const sampleRange = targetSheet.getRange(1, 1, Math.min(3, targetSheet.getLastRow()), Math.min(5, targetSheet.getLastColumn()));
        const sampleData = sampleRange.getValues();
        sampleData.forEach((row, rowIndex) => {
          console.log(`行${rowIndex + 1}: [${row.map(cell => `"${cell}"`).join(', ')}]`);
        });
      }
    } else {
      console.log('ターゲットシートが見つかりません: ' + SHEET_NAME);
    }
    
    console.log('=== スプレッドシート詳細情報デバッグ完了 ===');
    return {
      spreadsheetName: ss.getName(),
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      sheetCount: sheets.length,
      targetSheetExists: !!targetSheet,
      targetSheetName: targetSheet ? targetSheet.getName() : null,
      lastRow: targetSheet ? targetSheet.getLastRow() : 0,
      lastColumn: targetSheet ? targetSheet.getLastColumn() : 0
    };
    
  } catch (error) {
    console.error('=== デバッグ情報取得エラー ===');
    console.error('エラー詳細: ' + error.toString());
    console.error('エラースタック: ' + (error.stack || '取得できません'));
    throw error;
  }
}