/**
 * 楽天ランキングAPI × GAS 連携システム
 * * 【Phase 1】基盤構築（スクリプトプロパティ設定解説・設定シート作成・API疎通テスト）
 * * 開発者: 開発担当AI
 * * ==========================================================================
 * 【手動設定：スクリプトプロパティの登録手順】
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
 * * 対象スプレッドシートに「設定」シートが存在しない場合、新規に自動生成します。
 * 既に存在する場合は、上書きせずにその旨をログに記録します。
 */
function initializeSettingSheet() {
  try {
    const ss = getTargetSpreadsheet_();
    let settingSheet = ss.getSheetByName('設定');

    if (!settingSheet) {
      // 設定シートがない場合は新規作成
      settingSheet = ss.insertSheet('設定');
      
      // ヘッダーの設定
      const headers = ['ジャンル名', 'ジャンルID', '取得件数', '有効フラグ'];
      settingSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // デザイン調整（ヘッダー背景色、太字、枠線）
      const headerRange = settingSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#F3F4F6') // 薄いグレー
                 .setFontWeight('bold')
                 .setHorizontalAlignment('center');
      
      // サンプルデータの書き込み（スイーツ: 100283）
      settingSheet.getRange(2, 1, 1, 4).setValues([['スイーツ', 100283, 10, true]]);
      
      // 列幅の自動調整
      settingSheet.autoResizeColumns(1, headers.length);
      
      console.log('✅ 「設定」シートを新規に作成し、ヘッダーとサンプルデータを初期化しました。');
    } else {
      console.log('ℹ️ 「設定」シートは既に存在するため、初期化処理をスキップしました。');
    }
  } catch (e) {
    console.error(`エラーが発生しました (initializeSettingSheet): ${e.message}`);
  }
}

/**
 * 【タスク 1-3】楽天ランキングAPI 疎通テスト関数
 * * スクリプトプロパティに保存された楽天の認証情報を使用して、
 * 特定ジャンル（デフォルト: スイーツ 100283）のランキング情報を1件お試しで呼び出し、
 * 正常に応答（API疎通）があるかをログ出力で検証します。
 */
function testRakutenApiConnection() {
  console.log('楽天ランキングAPIの疎通テストを開始します...');

  let applicationId;
  let accessKey;
  try {
    applicationId = getRequiredProperty_('RAKUTEN_APPLICATION_ID');
    accessKey = getRequiredProperty_('RAKUTEN_ACCESS_KEY');
  } catch (e) {
    console.error(e.message);
    return;
  }

  // 疎通確認用のパラメータ設定（スイーツ: 100283, 1ページ目）
  const testGenreId = 100283;
  const url = `${RAKUTEN_API_ENDPOINT}?applicationId=${applicationId}&accessKey=${accessKey}&genreId=${testGenreId}&page=1&formatVersion=2`;

  const options = {
    method: 'get',
    muteHttpExceptions: true // エラー時のレスポンスコードを取得するため有効化
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      console.log('✅ API疎通テストに成功しました！');
      
      // 取得データの一部を簡易ログ出力
      if (data && data.Products && data.Products.length > 0) {
        console.log(`取得商品数: ${data.Products.length} 件`);
        const topItem = data.Products[0];
        console.log(`【1位の商品サンプル】`);
        console.log(`商品名: ${topItem.itemName}`);
        console.log(`価格: ${topItem.itemPrice} 円`);
        console.log(`ショップ名: ${topItem.shopName}`);
      } else if (data && data.length > 0) {
        // formatVersion=2 の場合、直接配列で取得できるケースもあります
        console.log(`取得商品数: ${data.length} 件`);
        const topItem = data[0];
        console.log(`【1位の商品サンプル】`);
        console.log(`商品名: ${topItem.itemName}`);
        console.log(`価格: ${topItem.itemPrice} 円`);
        console.log(`ショップ名: ${topItem.shopName}`);
      } else {
        console.warn('APIは200を返しましたが、ランキング商品リストが空でした。');
      }
    } else {
      console.error(`❌ API疎通テストに失敗しました。ステータスコード: ${responseCode}`);
      console.error(`エラー内容: ${response.getContentText()}`);
    }
  } catch (e) {
    console.error(`ネットワーク・またはその他のエラーが発生しました: ${e.message}`);
  }
}