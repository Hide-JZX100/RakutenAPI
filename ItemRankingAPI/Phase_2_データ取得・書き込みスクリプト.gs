/**
 * 楽天ランキングAPI × GAS 連携システム
 * * 【Phase 2】データ取得・書き込み処理（設定シート読み込み、ランキング取得、スプレッドシート上書き保存）
 * * 開発者: 開発担当AI
 * * ==========================================================================
 * 本スクリプトは、同一フォルダ（プロジェクト）内の `Phase 1_基盤構築スクリプト.gs` に定義された
 * グローバル関数（getTargetSpreadsheet_ など）を呼び出します。
 * ==========================================================================
 */

/**
 * 「設定」シートから有効なジャンル設定一覧をオブジェクト配列として取得します。
 * 有効フラグが TRUE の行のみを対象とします。
 * 
 * @param {string} [sheetName='設定'] 読み込み対象の設定シート名
 * @returns {Object[]} 設定情報の配列。各要素は { genreName, genreId, limitCount } の形式。
 * @private
 */
function getActiveGenreSettings_(sheetName = '設定') {
  try {
    const ss = getTargetSpreadsheet_();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`設定シート「${sheetName}」が見つかりません。`);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      console.warn(`「${sheetName}」シートにデータが存在しません。`);
      return [];
    }

    // A2からDの最終行までデータ範囲を取得（ヘッダー行は除く）
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 4);
    const values = dataRange.getValues();

    const activeSettings = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const genreName = row[0];
      const genreId = row[1];
      const limitCount = row[2];
      const isActive = row[3];

      // 有効フラグが明確に TRUE の場合のみ追加
      if (isActive === true) {
        activeSettings.push({
          genreName: String(genreName),
          genreId: Number(genreId),
          limitCount: Number(limitCount)
        });
      }
    }

    console.log(`ℹ️ 「${sheetName}」シートから有効な設定を ${activeSettings.length} 件取得しました。`);
    return activeSettings;
  } catch (e) {
    console.error(`エラーが発生しました (getActiveGenreSettings_): ${e.message}`);
    throw e;
  }
}

/**
 * 【テスト】getActiveGenreSettings_関数の動作検証用テスト関数
 * 現在の「設定」シートから有効な設定データを正常に読み込み、
 * ログに出力して構造を検証します。
 */
function testReadSettings() {
  console.log('--- testReadSettings 開始 ---');
  try {
    const settings = getActiveGenreSettings_();
    
    if (settings && settings.length > 0) {
      console.log(`✅ 読み込みテスト成功: 設定件数 ${settings.length} 件`);
      settings.forEach((setting, index) => {
        console.log(`[設定 ${index + 1}] ジャンル名: ${setting.genreName}, ジャンルID: ${setting.genreId}, 取得件数: ${setting.limitCount}`);
      });
    } else {
      console.warn('⚠️ 有効なジャンル設定が見つかりませんでした。「設定」シートの有効フラグがTRUEになっているデータがあるか確認してください。');
    }
  } catch (error) {
    console.error(`❌ testReadSettings 失敗: ${error.message}`);
  }
  console.log('--- testReadSettings 終了 ---');
}
