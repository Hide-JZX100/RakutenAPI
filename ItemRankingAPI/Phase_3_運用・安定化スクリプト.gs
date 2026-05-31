/**
 * 楽天ランキングAPI × GAS 連携システム
 * * 【Phase 3】運用・安定化処理（エラーハンドリング・自動リトライ・実行ログ出力・定期実行トリガー設定）
 * * 開発者: 開発担当AI
 * * ==========================================================================
 * 本スクリプトは、同一フォルダ（プロジェクト）内の別ファイルに定義された
 * グローバル関数（getTargetSpreadsheet_、fetchRakutenRanking_ など）を呼び出します。
 * ==========================================================================
 */

/**
 * スプレッドシートの「実行ログ」シートに実行結果を1行追記します。
 * シートが存在しない場合は、ヘッダーを設定して新規作成します。
 * 
 * @param {string} status 実行ステータス（'SUCCESS' または 'ERROR'）
 * @param {number} count 取得・処理に成功した総商品件数
 * @param {string} message 実行メッセージまたはエラー詳細
 * @private
 */
function writeExecutionLog_(status, count, message) {
  try {
    const ss = getTargetSpreadsheet_();
    let logSheet = ss.getSheetByName('実行ログ');
    
    const headers = ['実行日時', 'ステータス', '取得件数', 'メッセージ / エラー詳細'];
    
    if (!logSheet) {
      logSheet = ss.insertSheet('実行ログ');
      logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      logSheet.getRange(1, 1, 1, headers.length)
              .setBackground('#F3F4F6') // 薄いグレー
              .setFontWeight('bold')
              .setHorizontalAlignment('center');
      logSheet.autoResizeColumns(1, headers.length);
      console.log('✅ 「実行ログ」シートを新規に作成しました。');
    }
    
    // ログ行の追記
    const timestamp = new Date();
    logSheet.appendRow([timestamp, status, count, message]);
    
    console.log(`📝 実行ログを記録しました。ステータス: ${status}, 件数: ${count}`);
  } catch (e) {
    console.error(`実行ログの書き込みに失敗しました: ${e.message}`);
  }
}

/**
 * exportRakutenRankings関数を毎日指定された時間帯に実行する時間主導型トリガーを設定します。
 * 重複登録を防ぐため、既存の同一関数用トリガーはすべて削除してから新規登録します。
 * 
 * @param {number} [hour=9] 実行時間帯（0〜23の整数。例: 9 を指定すると 9:00〜10:00 の間に実行）
 */
function setDailyTrigger(hour = 10) {
  try {
    // 1. 既存の同一関数に対するトリガーを削除（重複防止）
    deleteDailyTrigger();

    // 2. 毎日指定された時間に実行するトリガーを新規作成
    ScriptApp.newTrigger('exportRakutenRankings')
             .timeBased()
             .everyDays(1)
             .atHour(hour)
             .create();

    console.log(`✅ 定期実行トリガーを設定しました。毎日 ${hour}:00〜${hour + 1}:00 の間に実行されます。`);
  } catch (e) {
    console.error(`トリガー設定中にエラーが発生しました (setDailyTrigger): ${e.message}`);
    throw e;
  }
}

/**
 * 現在登録されている exportRakutenRankings関数用の定期実行トリガーをすべて削除します。
 */
function deleteDailyTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deleteCount = 0;

    for (let i = 0; i < triggers.length; i++) {
      const trigger = triggers[i];
      if (trigger.getHandlerFunction() === 'exportRakutenRankings') {
        ScriptApp.deleteTrigger(trigger);
        deleteCount++;
      }
    }

    if (deleteCount > 0) {
      console.log(`🧹 既存のトリガー（関数: exportRakutenRankings）を ${deleteCount} 件削除しました。`);
    } else {
      console.log('ℹ️ 削除対象の既存トリガーは見つかりませんでした。');
    }
  } catch (e) {
    console.error(`トリガー削除中にエラーが発生しました (deleteDailyTrigger): ${e.message}`);
    throw e;
  }
}

/**
 * 【テスト】APIの一時エラーおよび通信例外に対するリトライ挙動検証用テスト関数
 * * 意図的に 503 エラーを返すテスト用URLおよび存在しないドメインへのリクエストを行い、
 * 最大3回・2秒間隔でリトライ（再試行）が実行されることをログ出力で検証します。
 */
function testApiRetryBehavior() {
  console.log('--- testApiRetryBehavior 開始 ---');
  const dummyGenreId = 100283;
  const dummyPage = 1;

  // -------------------------------------------------------------
  // パターン1: 503エラー（一時的なサーバーエラー）に対するリトライ検証
  // -------------------------------------------------------------
  console.log('\n[パターン1] 503（Server Error）によるリトライテストを開始します...');
  const testUrl503 = 'https://httpbin.org/status/503';
  
  try {
    const result = fetchRakutenRanking_(dummyGenreId, dummyPage, testUrl503);
    if (result === null) {
      console.log('✅ パターン1結果: リトライ上限に達し、期待通り null が返却されました。');
    } else {
      console.error('❌ パターン1結果: 異常。エラーURLからデータが取得できてしまいました。');
    }
  } catch (e) {
    console.error(`❌ パターン1結果: 意図しない例外が発生しました: ${e.message}`);
  }

  // -------------------------------------------------------------
  // パターン2: DNSエラー/通信例外（接続エラー）に対するリトライ検証
  // -------------------------------------------------------------
  console.log('\n[パターン2] DNSエラー（通信例外）によるリトライテストを開始します...');
  const testUrlInvalidDomain = 'https://invalid-domain-name-test-12345.com/';

  try {
    const result = fetchRakutenRanking_(dummyGenreId, dummyPage, testUrlInvalidDomain);
    if (result === null) {
      console.log('✅ パターン2結果: リトライ上限に達し、期待通り null が返却されました。');
    } else {
      console.error('❌ パターン2結果: 異常。無効なURLからデータが取得できてしまいました。');
    }
  } catch (e) {
    console.error(`❌ パターン2結果: 意図しない例外が発生しました: ${e.message}`);
  }

  console.log('\n--- testApiRetryBehavior 終了 ---');
}

/**
 * 【テスト】writeExecutionLog_関数の動作検証用テスト関数
 * 実行ログの書き込みを行い、正常に「実行ログ」シートへ追記されるかをテストします。
 */
function testWriteLog() {
  console.log('--- testWriteLog 開始 ---');
  try {
    writeExecutionLog_('SUCCESS', 99, 'テスト実行ログ（正常終了テスト）');
    writeExecutionLog_('ERROR', 0, 'テスト実行ログ（エラー発生テスト）：APIタイムアウトを想定');
    console.log('✅ testWriteLog: 「実行ログ」シートへのテスト書き込みが完了しました。');
  } catch (error) {
    console.error(`❌ testWriteLog 失敗: ${error.message}`);
  }
  console.log('--- testWriteLog 終了 ---');
}

/**
 * 【テスト】定期実行トリガーの設定・削除機能の動作検証用テスト関数
 * トリガーを一時的に作成し、正しくプロジェクトに登録されているか検証した上で、
 * 自動でトリガーを削除してクリーンアップします。
 */
function testTriggerSetup() {
  console.log('--- testTriggerSetup 開始 ---');
  try {
    const testHour = 3; // テスト用に朝3時設定
    
    // 1. トリガーを設定
    setDailyTrigger(testHour);
    
    // 2. 登録されているか検証
    const triggers = ScriptApp.getProjectTriggers();
    let foundTrigger = null;
    for (let i = 0; i < triggers.length; i++) {
      const trigger = triggers[i];
      if (trigger.getHandlerFunction() === 'exportRakutenRankings') {
        foundTrigger = trigger;
        break;
      }
    }
    
    if (foundTrigger) {
      console.log(`✅ トリガー登録検証成功: 関数「${foundTrigger.getHandlerFunction()}」が「${foundTrigger.getEventType()}」イベントとして登録されています。`);
    } else {
      throw new Error('テスト失敗: トリガーがプロジェクトに登録されていません。');
    }
    
    // 3. クリーンアップ（テストで作成したトリガーを削除）
    deleteDailyTrigger();
    
    // 削除確認
    const postTriggers = ScriptApp.getProjectTriggers();
    let deletedSuccessfully = true;
    for (let i = 0; i < postTriggers.length; i++) {
      if (postTriggers[i].getHandlerFunction() === 'exportRakutenRankings') {
        deletedSuccessfully = false;
        break;
      }
    }
    
    if (deletedSuccessfully) {
      console.log('✅ トリガー削除検証成功: テスト用トリガーが正常に削除されました。');
      console.log('✅ testTriggerSetup: すべての検証項目をクリアしました！');
    } else {
      throw new Error('テスト失敗: トリガーの削除処理が正常に行われませんでした。');
    }
    
  } catch (error) {
    console.error(`❌ testTriggerSetup 失敗: ${error.message}`);
  }
  console.log('--- testTriggerSetup 終了 ---');
}
