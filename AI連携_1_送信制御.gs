/**
 * ============================================================================
 * AI連携_1_送信制御.gs（設定・トリガー・メール送信・暗号化）
 * ※定数名の衝突を完全に防ぐため getAiBridgeConfig_() 関数で設定を取得します
 * ============================================================================
 */

function getAiBridgeConfig_() {
  return {
    RECIPIENT_EMAIL: 'capsdoctor365@caps365.jp',
    SUBJECT_PREFIX: '【AI連携暗号化JSON】医師不在・シフト状況パック',
    EXT_SHIFT_SS_ID: '1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co',
    DEFAULT_SECRET_KEY: 'CAPS_MNYS_AI_AGENT_SECRET_2026_KEY_#9876',

    PENDING_FLAG_KEY: 'AI_BRIDGE_PENDING_SEND',
    LAST_SENT_HASH_KEY: 'AI_BRIDGE_LAST_SENT_HASH',
    LAST_SENT_TIME_KEY: 'AI_BRIDGE_LAST_SENT_TIME_MS',

    MIN_INTERVAL_MS: 30 * 60 * 1000,

    EXCLUDED_CLINIC_NOS: new Set(['1000', '1002', '1004', '1005', '1017', '1018']),
    EXCLUDED_CLINIC_NAMES: [
      '出張インフルエンザワクチン',
      '【埼玉】ワクチンバックアップシフト',
      '職域新型コロナワクチン接種',
      '【千葉】ワクチンバックアップシフト',
      '【東京】ワクチンバックアップシフト',
      '院外勤務（内科）',
      '院外勤務（小児科）',
      '【関東】バックアップシフト',
      '欠勤',
      '有給',
      '嘱託医業務',
      '医師会業務',
      '休館'
    ]
  };
}

function triggerAiBridgeExport() {
  const cfg = getAiBridgeConfig_();
  const props = PropertiesService.getScriptProperties();
  try {
    props.deleteProperty(cfg.PENDING_FLAG_KEY);
    sendEncryptedShiftContextToCaps(false);
  } catch (e) {
    props.setProperty(cfg.PENDING_FLAG_KEY, String(Date.now()));
    Logger.log('直接送信できなかったため、1回限りの代理送信予約をセットしました: ' + e.message);
  }
}

function installAiBridgeFallbackTrigger() {
  const funcName = 'processPendingAiBridgeExport';
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === funcName) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(funcName).timeBased().everyMinutes(10).create();
  Logger.log('✅ AI連携のバックグラウンド代理送信トリガー（10分間隔）を設定しました。');
}

function processPendingAiBridgeExport() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  try {
    const cfg = getAiBridgeConfig_();
    const props = PropertiesService.getScriptProperties();
    const pendingTimestamp = props.getProperty(cfg.PENDING_FLAG_KEY);
    if (!pendingTimestamp) return;

    props.deleteProperty(cfg.PENDING_FLAG_KEY);

    if (Date.now() - Number(pendingTimestamp) > 2 * 60 * 60 * 1000) {
      Logger.log('古い代理送信フラグのため破棄しました。');
      return;
    }

    sendEncryptedShiftContextToCaps(false);
  } catch (e) {
    Logger.log('代理送信時にエラーが発生しました: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

function testSendEncryptedShiftContextManual() {
  sendEncryptedShiftContextToCaps(true);
}

function sendEncryptedShiftContextToCaps(isForceManual) {
  const cfg = getAiBridgeConfig_();
  const props = PropertiesService.getScriptProperties();
  const nowMs = Date.now();

  const lastSentTimeMs = Number(props.getProperty(cfg.LAST_SENT_TIME_KEY) || 0);
  if (!isForceManual && lastSentTimeMs > 0 && (nowMs - lastSentTimeMs) < cfg.MIN_INTERVAL_MS) {
    Logger.log('🛑 暴走防止ストッパー作動：直近30分以内に送信済みのためスキップしました。');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const payloadObj = buildAiShiftContextPayload_(ss);

  if (!payloadObj) {
    Logger.log('AI連携送信スキップ：ペイロードが生成できませんでした。');
    return;
  }

  const contentForHash = JSON.stringify({
    u: payloadObj.unfilledList,
    a: payloadObj.activeShiftsByDate,
    o: payloadObj.ouboShiftsByDate,
    s: payloadObj.streakWarnings,
    k: payloadObj.spotCandidateRows,
    m: payloadObj.dailyMetricsByDate
  });
  const currentHash = computeSha256Hex_(contentForHash);
  const lastSentHash = props.getProperty(cfg.LAST_SENT_HASH_KEY) || '';

  if (!isForceManual && currentHash === lastSentHash) {
    Logger.log('🛑 2重送信防止ストッパー作動：前回と内容が同一のためスキップしました。');
    return;
  }

  props.setProperty(cfg.LAST_SENT_TIME_KEY, String(nowMs));
  props.setProperty(cfg.LAST_SENT_HASH_KEY, currentHash);

  const secretKey = props.getProperty('AI_BRIDGE_SECRET_KEY') || cfg.DEFAULT_SECRET_KEY;
  const jsonString = JSON.stringify(payloadObj);
  const encryptedPackage = encryptAndSignPayload_(jsonString, secretKey);

  const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  const subject = `${cfg.SUBJECT_PREFIX} (${nowStr})`;

  const body = [
    '※このメールは医療法人側システム（mijyusoku）からAIエージェント会議用に自動送信された暗号化データです。',
    `生成日時: ${nowStr}`,
    `不在枠件数: ${payloadObj.unfilledList.length}件`,
    `対象日数: ${Object.keys(payloadObj.activeShiftsByDate).length}日分`,
    `日別充足指標日数: ${Object.keys(payloadObj.dailyMetricsByDate || {}).length}日分`,
    `連勤注意候補件数: ${(payloadObj.streakWarnings || []).length}件`,
    `個別打診候補件数: ${Math.max(0, (payloadObj.spotCandidateRows || []).length - 1)}件`,
    `複数日程おまとめ医師数: ${Math.max(0, (payloadObj.bundledCandidateRows || []).length - 1)}名`,
    '',
    '---BEGIN_ENCRYPTED_AI_PAYLOAD---',
    JSON.stringify(encryptedPackage),
    '---END_ENCRYPTED_AI_PAYLOAD---'
  ].join('\n');

  MailApp.sendEmail({
    to: cfg.RECIPIENT_EMAIL,
    subject: subject,
    body: body,
    noReply: true
  });

  Logger.log(`✅ 送信完了（不在枠: ${payloadObj.unfilledList.length}件 / 日別指標: ${Object.keys(payloadObj.dailyMetricsByDate || {}).length}日分）`);
}

function computeSha256Hex_(str) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return digest.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function encryptAndSignPayload_(plainText, secretKey) {
  const textBlob = Utilities.newBlob(plainText, 'application/json', 'payload.json');
  const gzipBytes = Utilities.gzip(textBlob).getBytes();
  const ivStr = Utilities.getUuid().replace(/-/g, '').slice(0, 16);

  const cipherBytes = xorStreamCipher_(gzipBytes, secretKey, ivStr);
  const cipherBase64 = Utilities.base64Encode(cipherBytes);

  const sigBytes = Utilities.computeHmacSha256Signature(ivStr + ':' + cipherBase64, secretKey);
  const sigBase64 = Utilities.base64Encode(sigBytes);

  return { v: 1, iv: ivStr, data: cipherBase64, sig: sigBase64 };
}

function xorStreamCipher_(inputBytes, secretKey, ivStr) {
  const output = new Array(inputBytes.length);
  let keyStream = [];
  let blockIndex = 0;

  for (let i = 0; i < inputBytes.length; i++) {
    if (i % 32 === 0) {
      keyStream = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        `${secretKey}|${ivStr}|${blockIndex}`,
        Utilities.Charset.UTF_8
      );
      blockIndex++;
    }
    output[i] = (inputBytes[i] ^ keyStream[i % 32]) & 0xff;
    if (output[i] > 127) output[i] -= 256;
  }
  return output;
}