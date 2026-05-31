/**
 * 楽天ランキングAPI × GAS 連携システム
 * * 【Phase 2】データ取得・書き込み処理（設定シート読み込み、ランキング取得、スプレッドシート上書き保存）
 * * 開発者: 開発担当AI
 * * ==========================================================================
 * 本スクリプトは、同一フォルダ（プロジェクト）内の `Phase 1_基盤構築スクリプト.gs` に定義された
 * グローバル関数（getTargetSpreadsheet_、fetchRakutenRanking_ など）を呼び出します。
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
 * 取得した商品ランキングデータをスプレッドシート書き込み用の二次元配列形式に変換します。
 * 指定された取得件数（limitCount）に応じてデータをスライスします。
 * 
 * @param {Object[]} items 正規化された商品データの配列
 * @param {number} genreId ジャンルID
 * @param {number} limitCount 取得件数制限
 * @returns {Array[]} 書き込み用二次元配列（ヘッダー行は含まない）
 * @private
 */
function transformRankingData_(items, genreId, limitCount) {
  if (!items || items.length === 0) {
    return [];
  }

  // 順位（rank）の昇順でソート（1位から30位へ並び替え）
  // 元のitems配列を破壊しないようスプレッド構文でコピーしてからソートします
  const sortedItems = [...items].sort((a, b) => (a.rank || 0) - (b.rank || 0));

  // 設定された取得件数にスライス
  const targetItems = sortedItems.slice(0, limitCount);
  const timestamp = new Date();

  // 各行を二次元配列にマッピング
  return targetItems.map(item => {
    return [
      timestamp,                    // A: 取得日時
      item.rank,                    // B: 順位
      item.itemName,                // C: 商品名
      item.catchcopy,               // D: キャッチコピー
      item.itemPrice,               // E: 価格
      item.itemUrl,                 // F: 商品URL
      item.reviewCount,             // G: レビュー数
      item.reviewAverage,           // H: 評点
      item.shopName,                // I: ショップ名
      item.availability,            // J: 在庫
      genreId                       // K: ジャンルID
    ];
  });
}

/**
 * 指定された二次元配列のランキングデータを、ジャンル名のシートへ書き込みます。
 * シートが存在しない場合は新規作成し、ヘッダーを設定します。
 * 既に存在する場合は、古いデータをクリアした上で上書き保存します。
 * 
 * @param {Array[]} values 書き込み用二次元配列（データ行のみ）
 * @param {string} genreName ジャンル名（シート名になります）
 * @private
 */
function writeRankingToGenreSheet_(values, genreName) {
  try {
    const ss = getTargetSpreadsheet_();
    let sheet = ss.getSheetByName(genreName);
    
    const headers = [
      '取得日時', '順位', '商品名', 'キャッチコピー', '価格', 
      '商品URL', 'レビュー数', '評点', 'ショップ名', '在庫', 'ジャンルID'
    ];

    if (!sheet) {
      // シートが存在しない場合は新規追加
      sheet = ss.insertSheet(genreName);
      
      // ヘッダー行を設定
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // ヘッダーのデザイン調整（太字、背景色：薄い緑、中央揃え）
      sheet.getRange(1, 1, 1, headers.length)
           .setBackground('#E5F3E5') // 薄い緑
           .setFontWeight('bold')
           .setHorizontalAlignment('center');
      
      console.log(`✅ 新規にランキングシート「${genreName}」を作成しました。`);
    } else {
      // 既存シートがある場合は、2行目以降の既存データ行をすべてクリア
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        // A2から最終列・最終行の範囲をクリア
        sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
      }
      console.log(`🧹 既存の「${genreName}」シートのデータをクリアしました。`);
    }

    if (values && values.length > 0) {
      // 2行目から二次元配列データを一括書き込み
      sheet.getRange(2, 1, values.length, headers.length).setValues(values);
      
      // 列幅の自動調整
      sheet.autoResizeColumns(1, headers.length);
      console.log(`💾 「${genreName}」シートに ${values.length} 件のランキングデータを書き込みました。`);
    } else {
      console.warn(`⚠️ 「${genreName}」シートに書き込むデータがありませんでした。`);
    }
  } catch (e) {
    console.error(`エラーが発生しました (writeRankingToGenreSheet_): ${e.message}`);
    throw e;
  }
}

/**
 * 楽天ランキングデータ取得・書き込みのメイン実行関数
 * 設定シートから有効な全ジャンルの情報を取得し、APIを巡回して各ランキングシートを更新します。
 * 実行結果は「実行ログ」シートに自動記録されます。
 */
function exportRakutenRankings() {
  console.log('🚀 楽天ランキング取得・更新バッチ処理を開始します...');
  let totalCount = 0;
  let status = 'SUCCESS';
  let executionMessage = '正常終了';

  try {
    // 1. 有効なジャンル設定を取得
    const settings = getActiveGenreSettings_();
    if (!settings || settings.length === 0) {
      const warnMsg = '有効なジャンル設定が「設定」シートに登録されていません。';
      console.warn(`⚠️ ${warnMsg}`);
      writeExecutionLog_('SUCCESS', 0, warnMsg);
      return;
    }

    // 2. ジャンルごとに巡回
    for (let i = 0; i < settings.length; i++) {
      const setting = settings[i];
      console.log(`\n--- [${i + 1}/${settings.length}] ジャンル: ${setting.genreName} (ID: ${setting.genreId}) の処理を開始 ---`);
      
      // APIからデータを取得 (テストのため1ページ目＝30位まで)
      console.log('楽天APIからデータを取得中...');
      const items = fetchRakutenRanking_(setting.genreId, 1);
      
      if (items && items.length > 0) {
        // 取得したデータを二次元配列に変換＆取得件数にスライス
        console.log('データを整形中...');
        const values = transformRankingData_(items, setting.genreId, setting.limitCount);
        
        // スプレッドシートに書き込み
        writeRankingToGenreSheet_(values, setting.genreName);
        totalCount += values.length;
      } else {
        // API取得失敗時はエラーを投げ、ログにエラーを残すようにします
        throw new Error(`ジャンル「${setting.genreName}」のデータ取得に失敗しました（データが空、または通信エラー）。`);
      }

      // 楽天APIの負荷低減（レート制限対策）のため、ループの間に1秒スリープを挟む
      if (i < settings.length - 1) {
        console.log('次のジャンル取得まで1秒間待機します...');
        Utilities.sleep(1000);
      }
    }
    
    console.log('\n✨ すべてのジャンルのランキングデータ更新処理が正常に終了しました。');
    writeExecutionLog_(status, totalCount, executionMessage);

  } catch (e) {
    status = 'ERROR';
    executionMessage = e.message;
    console.error(`❌ バッチ処理中に重大なエラーが発生しました: ${e.message}`);
    writeExecutionLog_(status, totalCount, executionMessage);
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

/**
 * 【テスト】transformRankingData_関数およびAPI連携の動作検証用テスト関数
 * 有効な最初のジャンル設定を用いて実際にAPIから取得し、二次元配列へ正しく変換
 * されるかを検証します。
 */
function testFetchAndTransform() {
  console.log('--- testFetchAndTransform 開始 ---');
  try {
    const settings = getActiveGenreSettings_();
    if (!settings || settings.length === 0) {
      console.warn('⚠️ 有効なジャンル設定が見つからないためテストをスキップします。');
      return;
    }

    const testSetting = settings[0];
    console.log(`テスト実行ジャンル: ${testSetting.genreName} (ID: ${testSetting.genreId}, 件数: ${testSetting.limitCount})`);

    // 1. APIから取得 (テストのため1ページ目のみ)
    const items = fetchRakutenRanking_(testSetting.genreId, 1);
    if (!items || items.length === 0) {
      throw new Error('APIからデータを取得できませんでした。');
    }

    // 2. 二次元配列に変換 (スライス処理も含む)
    const values = transformRankingData_(items, testSetting.genreId, testSetting.limitCount);

    // 3. データの簡易検証 (アサーション)
    console.log(`✅ 変換後の行数: ${values.length} 行 (期待値: 最大 ${testSetting.limitCount} 行)`);
    if (values.length > 0) {
      const firstRow = values[0];
      console.log('【変換後の1行目サンプル】');
      console.log(`取得日時: ${firstRow[0]}`);
      console.log(`順位: ${firstRow[1]}`);
      console.log(`商品名: ${firstRow[2] ? firstRow[2].substring(0, 30) + '...' : 'null'}`);
      console.log(`価格: ${firstRow[4]} 円`);
      console.log(`ジャンルID: ${firstRow[10]}`);

      // 件数の一致確認
      const expectedLength = Math.min(items.length, testSetting.limitCount);
      if (values.length !== expectedLength) {
        throw new Error(`テスト失敗: 変換後の行数(${values.length})` + 
                         `が、期待される件数(${expectedLength})と一致しません。`);
      }
      console.log('✅ testFetchAndTransform: 二次元配列への変換および件数制限スライスが正常に動作しました！');
    } else {
      throw new Error('テスト失敗: 変換後の二次元配列が空です。');
    }

  } catch (error) {
    console.error(`❌ testFetchAndTransform 失敗: ${error.message}`);
  }
  console.log('--- testFetchAndTransform 終了 ---');
}

/**
 * 【テスト】exportRakutenRankings関数の結合動作検証用テスト関数
 * メインのバッチ処理を実行し、実際にシートが自動作成・上書きされ、
 * ランキングデータが綺麗に格納されるかを確認します。
 */
function testExportRankings() {
  console.log('--- testExportRankings（結合テスト） 開始 ---');
  exportRakutenRankings();
  console.log('--- testExportRankings（結合テスト） 終了 ---');
}
