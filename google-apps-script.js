/**
 * プレイフル診断アプリ - Google Apps Script
 * スプレッドシートへのデータ蓄積機能
 */

// スプレッドシートIDを設定（実際のスプレッドシートIDに変更してください）
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
const SHEET_NAME = 'プレイフル診断データ';

/**
 * POSTリクエストを処理する関数
 * @param {Event} e - POSTイベントオブジェクト
 * @return {ContentService} - JSON レスポンス
 */
function doPost(e) {
  try {
    // CORSヘッダーを設定
    const response = {
      success: false,
      message: '',
      timestamp: new Date().toISOString()
    };

    // リクエストデータの取得
    const data = JSON.parse(e.postData.contents);
    console.log('受信データ:', data);

    // データ検証
    if (!validateData(data)) {
      response.message = 'データの形式が正しくありません';
      return createResponse(response);
    }

    // スプレッドシートにデータを保存
    const result = saveToSpreadsheet(data);
    
    if (result.success) {
      response.success = true;
      response.message = 'データが正常に保存されました';
      response.rowNumber = result.rowNumber;
    } else {
      response.message = result.error;
    }

    return createResponse(response);

  } catch (error) {
    console.error('エラー:', error);
    return createResponse({
      success: false,
      message: 'サーバーエラーが発生しました: ' + error.toString(),
      timestamp: new Date().toISOString()
    });
  }
}

/**
 * GETリクエストを処理する関数（テスト用）
 */
function doGet(e) {
  return createResponse({
    success: true,
    message: 'プレイフル診断アプリ API は正常に動作しています',
    timestamp: new Date().toISOString()
  });
}

/**
 * データ検証関数
 * @param {Object} data - 検証するデータ
 * @return {boolean} - 検証結果
 */
function validateData(data) {
  const requiredFields = [
    'timestamp', 'name', 'age', 'gender', 'email',
    'totalScore', 'factor1', 'factor2', 'factor3', 'factor4', 'factor5',
    'answers'
  ];

  for (const field of requiredFields) {
    if (data[field] === undefined || data[field] === null) {
      console.error(`必須フィールドが不足: ${field}`);
      return false;
    }
  }

  // 年齢の範囲チェック
  if (data.age < 18 || data.age > 99) {
    console.error('年齢が範囲外:', data.age);
    return false;
  }

  // スコアの範囲チェック
  const scores = [data.totalScore, data.factor1, data.factor2, data.factor3, data.factor4, data.factor5];
  for (const score of scores) {
    if (score < 1 || score > 5) {
      console.error('スコアが範囲外:', score);
      return false;
    }
  }

  return true;
}

/**
 * スプレッドシートにデータを保存
 * @param {Object} data - 保存するデータ
 * @return {Object} - 保存結果
 */
function saveToSpreadsheet(data) {
  try {
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = createNewSheet(spreadsheet);
    }

    // ヘッダー行が存在しない場合は作成
    if (sheet.getLastRow() === 0) {
      createHeaders(sheet);
    }

    // データ行を作成
    const rowData = [
      new Date(data.timestamp),  // A: 診断日時
      data.name,                 // B: 名前
      data.age,                  // C: 年齢
      data.gender,               // D: 性別
      data.email,                // E: メールアドレス
      parseFloat(data.totalScore), // F: 総合スコア
      parseFloat(data.factor1),    // G: 日常の楽しさ発見
      parseFloat(data.factor2),    // H: 自由感・解放感
      parseFloat(data.factor3),    // I: 創造的・自発的活動
      parseFloat(data.factor4),    // J: 子どもとのプレイフル交流
      parseFloat(data.factor5),    // K: 社会的つながりでの楽しさ
      data.userAgent,            // L: ユーザーエージェント
    ];

    // 各質問の回答を追加（Q1-Q25）
    for (let i = 1; i <= 25; i++) {
      rowData.push(data.answers[i] || '');
    }

    // データを追加
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);

    // セルの書式設定
    formatNewRow(sheet, newRow);

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
 * 新しいシートを作成
 * @param {Spreadsheet} spreadsheet - スプレッドシートオブジェクト
 * @return {Sheet} - 作成されたシート
 */
function createNewSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet(SHEET_NAME);
  console.log('新しいシートを作成しました:', SHEET_NAME);
  return sheet;
}

/**
 * ヘッダー行を作成
 * @param {Sheet} sheet - シートオブジェクト
 */
function createHeaders(sheet) {
  const headers = [
    '診断日時', '名前', '年齢', '性別', 'メールアドレス',
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
  
  // 列幅を自動調整
  sheet.autoResizeColumns(1, headers.length);
  
  console.log('ヘッダー行を作成しました');
}

/**
 * 新しい行の書式設定
 * @param {Sheet} sheet - シートオブジェクト
 * @param {number} rowNumber - 行番号
 */
function formatNewRow(sheet, rowNumber) {
  // 交互の行色設定
  if (rowNumber % 2 === 0) {
    const range = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn());
    range.setBackground('#f8f9fa');
  }
  
  // 日付列の書式設定
  sheet.getRange(rowNumber, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  
  // スコア列の書式設定（小数点2桁）
  for (let col = 6; col <= 11; col++) {
    sheet.getRange(rowNumber, col).setNumberFormat('0.00');
  }
}

/**
 * CORSヘッダー付きレスポンスを作成
 * @param {Object} data - レスポンスデータ
 * @return {ContentService} - HTTPレスポンス
 */
function createResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
}

/**
 * OPTIONSリクエスト処理（CORS対応）
 */
function doOptions(e) {
  return createResponse({ message: 'CORS preflight OK' });
}

/**
 * データ統計取得関数（管理用）
 */
function getStatistics() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { totalResponses: 0 };
    }
    
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    const values = dataRange.getValues();
    
    const stats = {
      totalResponses: values.length,
      averageAge: 0,
      averageTotalScore: 0,
      genderDistribution: {},
      lastResponseDate: null
    };
    
    let ageSum = 0;
    let scoreSum = 0;
    
    values.forEach(row => {
      // 年齢
      ageSum += row[2];
      
      // 総合スコア
      scoreSum += row[5];
      
      // 性別分布
      const gender = row[3];
      stats.genderDistribution[gender] = (stats.genderDistribution[gender] || 0) + 1;
      
      // 最新の回答日時
      if (!stats.lastResponseDate || row[0] > stats.lastResponseDate) {
        stats.lastResponseDate = row[0];
      }
    });
    
    stats.averageAge = ageSum / values.length;
    stats.averageTotalScore = scoreSum / values.length;
    
    return stats;
    
  } catch (error) {
    console.error('統計取得エラー:', error);
    return { error: error.toString() };
  }
}