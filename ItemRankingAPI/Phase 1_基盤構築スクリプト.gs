/**
 * 楽天ランキングAPI × GAS 連携システム
 * * 【Phase 1】基盤構築（スクリプトプロパティ設定解説・設定シート作成・API疎通テスト）
 * * 開発者: 開発担当AI
 * * ==========================================================================
 * 【重要・事前設定：スクリプトプロパティの登録手順】
 * * 本スクリプトを実行する前に、Google Apps Script (GAS) のエディタ上で
 * 以下の「プロパティ」を手動で追加してください。これにより安全にAPIを利用できます。
 * * 1. GASエディタの画面左側にあるメニューから「プロジェクトの設定」（歯車アイコン ⚙️）をクリックします。
 * 2. 画面を一番下までスクロールし、「スクリプトプロパティを編集」ボタンをクリックします。
 * 3. 「スクリプトプロパティを追加」をクリックし、以下の【プロパティ名（キー）】と【値】のペアを登録します。
 * * ① 【プロパティ名】: RAKUTEN_APPLICATION_ID
 * 【値】: (ご自身の楽天アプリケーションIDを入力)
 * * ② 【プロパティ名】: RAKUTEN_ACCESS_KEY
 * 【値】: (ご自身の楽天アクセスキーを入力)
 * * ③ 【プロパティ名】: TARGET_SPREADSHEET_ID
 * 【値】: (書き込み対象とするGoogleスプレッドシートのIDを入力)
 * ※スプレッドシートのURL内の 「.../d/【ここの部分】/edit...」 です。
 * * 4. 入力後、「スクリプトプロパティを保存」ボタンを必ず押してください。
 * * ==========================================================================
 * 【重要：楽天デベロッパー側の「アプリURL（リファラ）」設定とGASの制約】
 * * 楽天デベロッパーの必須入力項目である「アプリURL」に登録したドメインは、
 * 楽天API側で自動的にアクセス許可リスト（リファラ制限）の対象になります。
 * * GAS（Google）の共有ドメインは楽天側で弾かれる場合があるため、
 * [対策] 楽天アプリ管理画面（ https://webservice.rakuten.co.jp/app/list ）にアクセスし、
 * 対象アプリの「アプリURL」欄を【 https://example.com 】のような独自ドメインに設定し保存してください。
 * (※ご自身のホームページURLやブログURLがある場合は、そちらでも構いません)
 * ==========================================================================
 */

// ==========================================
// 定数・グローバル定義
// ==========================================
const RAKUTEN_API_ENDPOINT = 'https://openapi.rakuten.co.jp/ichibaranking/api/IchibaItem/Ranking/20220601';

/**
 * スクリプトプロパティから必要な値を安全に取得するヘルパー関数
 * * @param {string} key プロパティ名
 * @returns {string} プロパティ値
 * @throws {Error} プロパティが存在しない場合はエラーを投げます
 */
function getRequiredProperty_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error(`プロパティエラー: 「${key}」がスクリプトプロパティに設定されていません。GASの設定メニュー（歯車アイコン）から設定してください。`);
  }
  return value;
}

/**
 * 接続先スプレッドシートオブジェクトを取得するヘルパー関数
 * * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} Spreadsheetオブジェクト
 */
function getTargetSpreadsheet_() {
  const spreadsheetId = getRequiredProperty_('TARGET_SPREADSHEET_ID');
  try {
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    throw new Error(`スプレッドシート接続エラー: ID「${spreadsheetId}」を開くことができませんでした。IDが正しいか、権限があるかを確認してください。詳細: ${e.message}`);
  }
}

/**
 * 【タスク 1-2】設定シート作成・初期化関数
 * 対象スプレッドシートに指定の名前（デフォルト:「設定」）のシートが存在しない場合、新規に自動生成します。
 * 既に存在する場合は、上書きせずにその旨をログに記録します。
 * 
 * @param {string} [sheetName='設定'] 作成する設定シートの名前
 */
function initializeSettingSheet(sheetName = '設定') {
  try {
    const ss = getTargetSpreadsheet_();
    let settingSheet = ss.getSheetByName(sheetName);

    if (!settingSheet) {
      // 設定シートがない場合は新規作成
      settingSheet = ss.insertSheet(sheetName);
      
      // ヘッダーの設定
      const headers = ['ジャンル名', 'ジャンルID', '取得件数', '有効フラグ'];
      settingSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // デザイン調整（ヘッダー背景色、太字、中央揃え）
      const headerRange = settingSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#F3F4F6') // 薄いグレー
                 .setFontWeight('bold')
                 .setHorizontalAlignment('center');
      
      // サンプルデータの書き込み（スイーツ: 100283）
      settingSheet.getRange(2, 1, 1, 4).setValues([['スイーツ', 100283, 10, true]]);
      
      // 列幅の自動調整
      settingSheet.autoResizeColumns(1, headers.length);
      
      console.log(`✅ 「${sheetName}」シートを新規に作成し、ヘッダーとサンプルデータを初期化しました。`);
    } else {
      console.log(`ℹ️ 「${sheetName}」シートは既に存在するため、初期化処理をスキップしました。`);
    }
  } catch (e) {
    console.error(`エラーが発生しました (initializeSettingSheet): ${e.message}`);
  }
}

/**
 * 【テスト】initializeSettingSheet関数の動作検証用テスト関数
 * 一時的なテスト用シート名（例: 「設定_TEST」）を用いて設定シートを作成し、
 * 正しくヘッダーや初期値、フォーマットが設定されるかを検証します。
 * テスト終了後は、作成した一時シートを自動で削除し、元の環境に影響を与えないようにします。
 */
function testInitializeSettingSheet() {
  console.log('--- testInitializeSettingSheet 開始 ---');
  const ss = getTargetSpreadsheet_();
  const testSheetName = '設定_TEST';
  
  // 以前のテストシートが残っていれば削除しておく
  const oldTestSheet = ss.getSheetByName(testSheetName);
  if (oldTestSheet) {
    ss.deleteSheet(oldTestSheet);
  }
  
  try {
    // 1. 関数を実行
    initializeSettingSheet(testSheetName);
    
    // 2. 作成されたシートの取得
    const sheet = ss.getSheetByName(testSheetName);
    if (!sheet) {
      throw new Error(`テスト失敗: シート「${testSheetName}」が作成されませんでした。`);
    }
    
    // 3. ヘッダー行の検証
    const headers = sheet.getRange(1, 1, 1, 4).getValues()[0];
    const expectedHeaders = ['ジャンル名', 'ジャンルID', '取得件数', '有効フラグ'];
    for (let i = 0; i < expectedHeaders.length; i++) {
      if (headers[i] !== expectedHeaders[i]) {
        throw new Error(`テスト失敗: ヘッダー ${i + 1} 列目が「${headers[i]}」になっています（期待値: 「${expectedHeaders[i]}」）。`);
      }
    }
    
    // 4. サンプルデータの検証
    const sampleRow = sheet.getRange(2, 1, 1, 4).getValues()[0];
    const expectedSample = ['スイーツ', 100283, 10, true];
    if (sampleRow[0] !== expectedSample[0] ||
        Number(sampleRow[1]) !== expectedSample[1] ||
        Number(sampleRow[2]) !== expectedSample[2] ||
        sampleRow[3] !== expectedSample[3]) {
      throw new Error(`テスト失敗: サンプルデータが正しくありません。「${sampleRow}」になっています。`);
    }
    
    console.log('✅ testInitializeSettingSheet: すべての検証項目をクリアしました！');
    
  } catch (error) {
    console.error(`❌ testInitializeSettingSheet 失敗: ${error.message}`);
  } finally {
    // 一時テストシートのクリーンアップ
    const sheetToDelete = ss.getSheetByName(testSheetName);
    if (sheetToDelete) {
      ss.deleteSheet(sheetToDelete);
      console.log(`🧹 テスト用一時シート「${testSheetName}」を削除しました。`);
    }
    console.log('--- testInitializeSettingSheet 終了 ---');
  }
}

/**
 * 楽天商品ランキングAPIを呼び出してデータを取得する内部関数
 * 
 * @param {number} genreId ジャンルID
 * @param {number} page ページ番号（1〜34）
 * @returns {Object[]|null} 商品データの配列。エラー発生時は null を返します。
 * @private
 */
function fetchRakutenRanking_(genreId, page) {
  let applicationId;
  let accessKey;
  try {
    applicationId = getRequiredProperty_('RAKUTEN_APPLICATION_ID');
    accessKey = getRequiredProperty_('RAKUTEN_ACCESS_KEY');
  } catch (e) {
    console.error(`認証情報取得エラー: ${e.message}`);
    return null;
  }

  // スクリプトプロパティからリファラ用のURLを取得（未設定時はデフォルトのGASドメインを使用）
  const registeredAppUrl = PropertiesService.getScriptProperties().getProperty('RAKUTEN_APP_URL') || 'https://script.google.com/';

  const url = `${RAKUTEN_API_ENDPOINT}?applicationId=${applicationId}&accessKey=${accessKey}&genreId=${genreId}&page=${page}&formatVersion=2`;

  const options = {
    method: 'get',
    headers: {
      // GASのUrlFetchAppにおけるリファラ制限による403エラーを回避するため、
      // 楽天管理画面に登録されているリファラURLをRefererおよびOriginに設定します。
      'Referer': registeredAppUrl,
      'Origin': registeredAppUrl
    },
    muteHttpExceptions: true // エラー時のレスポンスコードを取得するため有効化
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      
      let rawItems = [];
      // レスポンスの主要構造の抽出
      if (data && data.Items && Array.isArray(data.Items)) {
        rawItems = data.Items;
      } else if (data && data.Products && Array.isArray(data.Products)) {
        rawItems = data.Products;
      } else if (Array.isArray(data)) {
        rawItems = data;
      } else {
        console.warn('APIの応答構造が想定と異なります。返却データ:', JSON.stringify(data).substring(0, 200));
        return null;
      }

      // レスポンスの正規化処理（ネスト構造 Items[i].Item.itemName もフラット構造にマッピング）
      return rawItems.map(item => {
        const target = (item && item.Item) ? item.Item : item;
        
        return {
          rank: target.rank || null,
          itemName: target.itemName || null,
          catchcopy: target.catchcopy || null,
          itemCode: target.itemCode || null,
          itemPrice: target.itemPrice ? Number(target.itemPrice) : null,
          itemUrl: target.itemUrl || null,
          reviewCount: target.reviewCount ? Number(target.reviewCount) : null,
          reviewAverage: target.reviewAverage ? Number(target.reviewAverage) : null,
          shopName: target.shopName || null,
          shopUrl: target.shopUrl || null,
          availability: target.availability !== undefined ? Number(target.availability) : null,
          genreId: target.genreId || null
        };
      });

    } else {
      console.error(`❌ 楽天APIリクエスト失敗。ステータスコード: ${responseCode}`);
      console.error(`エラー詳細: ${response.getContentText()}`);
      return null;
    }
  } catch (e) {
    console.error(`ネットワーク・または通信エラーが発生しました: ${e.message}`);
    return null;
  }
}

/**
 * 【タスク 1-3】楽天ランキングAPI 疎通テスト関数
 * スクリプトプロパティの認証情報とダミーリファラ（設定済みの場合は RAKUTEN_APP_URL）を使用して、
 * 特定ジャンル（スイーツ: 100283）のランキング情報を1件お試しで取得し、
 * APIとの疎通が正常に行えるか検証します。
 */
function testRakutenApiConnection() {
  console.log('--- testRakutenApiConnection 開始 ---');
  const testGenreId = 100283;
  const testPage = 1;

  try {
    console.log(`疎通テスト実行中...（ジャンルID: ${testGenreId}, ページ: ${testPage}）`);
    const items = fetchRakutenRanking_(testGenreId, testPage);

    if (items && items.length > 0) {
      console.log('✅ API疎通テストに成功しました！');
      console.log(`取得商品数: ${items.length} 件`);
      
      // 1位の商品の詳細をログ出力して検証
      const topItem = items[0];
      console.log(`【1位の商品サンプル】`);
      console.log(`順位 (rank): ${topItem.rank || '未取得'}`);
      console.log(`商品名 (itemName): ${topItem.itemName ? topItem.itemName.substring(0, 40) + '...' : '未取得'}`);
      console.log(`価格 (itemPrice): ${topItem.itemPrice ? topItem.itemPrice + ' 円' : '未取得'}`);
      console.log(`ショップ名 (shopName): ${topItem.shopName || '未取得'}`);
    } else {
      throw new Error('APIから有効な商品リストを取得できませんでした。レスポンスが空、またはエラーが発生しています。');
    }
  } catch (error) {
    console.error(`❌ testRakutenApiConnection 失敗: ${error.message}`);
  }
  console.log('--- testRakutenApiConnection 終了 ---');
}