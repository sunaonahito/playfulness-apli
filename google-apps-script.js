/**
 * プレイフル診断アプリ - Google Apps Script (修正版)
 * スプレッドシートへのデータ蓄積機能
 */

// スプレッドシートIDを設定
const SPREADSHEET_ID = '1O31BlTX6tfNv8Eo6vxbalQZGtsyM7E4fOYPl2-st8cY';
const SHEET_NAME = 'プレイフル診断データ';

/**
 * POSTリクエストを処理する関数
 */
function doPost(e) {
  try {
    // リクエストデータの取得
    const data = JSON.parse(e.parameter.data);
    console.log('受信データ:', data);

    // スプレッドシートにデータを保存
    const result = saveToSpreadsheet(data);
    
    const response = {
      success: result.success,
      message: result.success ? 'データが正常に保存されました' : result.error,
      timestamp: new Date().toISOString()
    };

    return ContentService
      .createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('エラー:', error);
    const errorResponse = {
      success: false,
      message: 'サーバーエラーが発生しました: ' + error.toString(),
      timestamp: new Date().toISOString()
    };
    
    return ContentService
      .createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GETリクエストを処理する関数（テスト用）
 */
function doGet(e) {
  const response = {
    success: true,
    message: 'プレイフル診断アプリ API は正常に動作しています',
    timestamp: new Date().toISOString()
  };
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * スプレッドシートにデータを保存
 */
function saveToSpreadsheet(data) {
  try {
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
      createHeaders(sheet);
    }

    // ヘッダー行が存在しない場合は作成
    if (sheet.getLastRow() === 0) {
      createHeaders(sheet);
    }

    // データ行を作成
    const rowData = [
      new Date(data.timestamp),      // A: 診断日時
      data.name || '',               // B: 名前
      data.age || '',                // C: 年齢
      data.email || '',              // D: メールアドレス
      parseFloat(data.totalScore) || 0,  // E: 総合スコア
      parseFloat(data.factor1) || 0,     // F: 日常の楽しさ発見
      parseFloat(data.factor2) || 0,     // G: 自由感・解放感
      parseFloat(data.factor3) || 0,     // H: 創造的・自発的活動
      parseFloat(data.factor4) || 0,     // I: 子どもとのプレイフル交流
      parseFloat(data.factor5) || 0,     // J: 社会的つながりでの楽しさ
      data.userAgent || '',          // K: ユーザーエージェント
    ];

    // 各質問の回答を追加（Q1-Q25）
    if (data.answers) {
      for (let i = 1; i <= 25; i++) {
        rowData.push(data.answers[i] || '');
      }
    } else {
      // answersがない場合は空文字で埋める
      for (let i = 1; i <= 25; i++) {
        rowData.push('');
      }
    }

    // データを追加
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);

    // 日付列の書式設定
    sheet.getRange(newRow, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');

    console.log(`データを行 ${newRow} に保存しました`);
    
    return {
      success: true,
      rowNumber: newRow
    };

  } catch (error) {
    console.error('スプレッドシート保存エラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ヘッダー行を作成
 */
function createHeaders(sheet) {
  const headers = [
    '診断日時', '名前', '年齢', 'メールアドレス',
    '総合スコア', '日常の楽しさ発見', '自由感・解放感', '創造的・自発的活動',
    '子どもとのプレイフル交流', '社会的つながりでの楽しさ', 'ユーザーエージェント'
  ];

  // Q1-Q25の列を追加
  for (let i = 1; i <= 25; i++) {
    headers.push(`Q${i}`);
  }

  // ヘッダーを設定
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーのスタイル設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#FF8A9B');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  
  console.log('ヘッダー行を作成しました');
}

/**
 * テスト関数 - 手動実行でデータ保存をテスト
 */
function testSaveData() {
  const testData = {
    timestamp: new Date().toISOString(),
    name: 'テスト太郎',
    age: 30,
    email: 'test@example.com',
    totalScore: 3.5,
    factor1: 3.2,
    factor2: 3.8,
    factor3: 3.1,
    factor4: 3.9,
    factor5: 3.4,
    userAgent: 'Test Browser',
    answers: {
      1: 4, 2: 3, 3: 5, 4: 2, 5: 4,
      6: 3, 7: 4, 8: 3, 9: 5, 10: 3,
      11: 4, 12: 3, 13: 4, 14: 5, 15: 3,
      16: 4, 17: 5, 18: 4, 19: 3, 20: 4,
      21: 3, 22: 4, 23: 3, 24: 5, 25: 4
    }
  };
  
  const result = saveToSpreadsheet(testData);
  console.log('テスト結果:', result);
  return result;
}