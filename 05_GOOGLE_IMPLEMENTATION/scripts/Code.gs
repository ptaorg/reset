/**
 * PTA適正運営スターターキット v0.3.1
 * Googleフォーム・スプレッドシート実装パッケージ
 *
 * 設計原則：
 * - 入会した会員だけを管理する
 * - 入会しない者には何も提出させない
 * - 非加入届、加入・非加入二択フォーム、未提出者督促、学校名簿照合、非会員児童一覧は作らない
 * - PTA書類は学校提出物ではなく、PTAが直接受領する
 */

const PTAKIT_VERSION = 'v0.3.1';

const SHEET = Object.freeze({
  CONFIG: 'コード設定',
  LINKS: 'フォームURL',
  MEMBERS: '会員台帳',
  FEES: '会費管理',
  WITHDRAWALS: '退会管理',
  UPDATES: '更新受付',
  OFFICERS: '役員承諾',
  PERMISSIONS: '権限管理',
  DELETION: '削除予定',
  LOG: 'ログ',
  SELF_CHECK: '自己点検'
});

const PREFIX_KEY = Object.freeze({
  JOIN: 'JOIN_PREFIX',
  LEAVE: 'LEAVE_PREFIX',
  UPDATE: 'UPDATE_PREFIX',
  OFFICER: 'OFFICER_PREFIX'
});

const FORM_KEY = Object.freeze({
  JOIN: 'JOIN_FORM_ID',
  LEAVE: 'LEAVE_FORM_ID',
  UPDATE: 'UPDATE_FORM_ID',
  OFFICER: 'OFFICER_FORM_ID'
});

const HEADERS = Object.freeze({
  [SHEET.CONFIG]: ['キー','値','説明'],
  [SHEET.LINKS]: ['フォーム種別','フォームID','編集URL','回答URL','作成日時','備考'],
  [SHEET.MEMBERS]: ['受付番号','受付日時','ステータス','会員氏名','ふりがな','メール','電話番号','関係','児童生徒氏名','学年','組','会費年度','会費対象','個人情報同意','フォーム回答ID','最終更新日時','退会日','備考'],
  [SHEET.FEES]: ['会費ID','作成日時','会員受付番号','会費年度','会費額','納入状況','納入日','方法','備考'],
  [SHEET.WITHDRAWALS]: ['退会受付番号','受付日時','会員受付番号','会員氏名','メール','退会希望日','処理状況','会員台帳処理日','会費処理日','削除予定日','フォーム回答ID','備考'],
  [SHEET.UPDATES]: ['更新受付番号','受付日時','会員受付番号','会員氏名','新メール','新電話番号','変更内容','処理状況','フォーム回答ID','備考'],
  [SHEET.OFFICERS]: ['役員承諾番号','受付日時','会員受付番号','氏名','希望役職','承諾内容','個人情報規程確認','守秘義務確認','処理状況','フォーム回答ID','備考'],
  [SHEET.PERMISSIONS]: ['対象者','役割','閲覧可能シート','編集可能シート','付与日','削除予定日','削除日','備考'],
  [SHEET.DELETION]: ['対象ID','対象区分','氏名','発生日','削除予定日','状態','処理日','備考'],
  [SHEET.LOG]: ['日時','操作者','種別','対象番号','処理','結果','備考'],
  [SHEET.SELF_CHECK]: ['確認日時','区分','項目','結果','詳細','対応']
});

const DEFAULT_CONFIG = [
  ['PTA_NAME','〇〇学校PTA','フォーム名・通知文に使用するPTA名'],
  ['ADMIN_EMAIL','','管理者通知先メールアドレス。空欄なら通知しない'],
  ['PRIVACY_POLICY_URL','','プライバシーポリシーURL。未公開の場合は空欄可'],
  ['FEE_AMOUNT','0','年会費額。実運用前に変更'],
  ['FEE_YEAR', String(new Date().getFullYear()), '会費年度'],
  ['JOIN_PREFIX','J-','入会受付番号の接頭辞'],
  ['LEAVE_PREFIX','L-','退会受付番号の接頭辞'],
  ['UPDATE_PREFIX','U-','更新受付番号の接頭辞'],
  ['OFFICER_PREFIX','O-','役員承諾番号の接頭辞'],
  ['DELETION_AFTER_DAYS','365','退会後削除予定日の既定日数'],
  [FORM_KEY.JOIN,'','自動設定。手入力しない'],
  [FORM_KEY.LEAVE,'','自動設定。手入力しない'],
  [FORM_KEY.UPDATE,'','自動設定。手入力しない'],
  [FORM_KEY.OFFICER,'','自動設定。手入力しない']
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PTA適正運営 v0.3.1')
    .addItem('1 初期シート作成・修復', 'ptaInstallOrRepair')
    .addItem('2 フォーム作成・トリガー設定', 'ptaCreateForms')
    .addSeparator()
    .addItem('自己点検を実行', 'ptaRunSelfCheck')
    .addItem('フォームURLを一覧へ出力', 'ptaWriteFormUrls')
    .addItem('削除予定通知トリガーを設定', 'ptaInstallDeletionReminder')
    .addSeparator()
    .addItem('トリガー一覧をログ出力', 'ptaListTriggers')
    .addItem('本キットのトリガーを削除', 'ptaUninstallKitTriggers')
    .addToUi();
}

function ptaInstallOrRepair() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(function(name) {
    const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    ensureHeader_(sheet, HEADERS[name]);
  });
  seedConfig_();
  styleWorkbook_();
  appendLog_('system','初期設定','', '必要シート・ヘッダー作成/修復', '成功', PTAKIT_VERSION);
  ptaRunSelfCheck();
}

function ptaCreateForms() {
  ptaInstallOrRepair();
  const config = getConfig_();
  const forms = [
    {kind:'入会申込', key:FORM_KEY.JOIN, handler:'handleJoinSubmit', create:function(){ return buildJoinForm_(config); }},
    {kind:'退会通知', key:FORM_KEY.LEAVE, handler:'handleWithdrawalSubmit', create:function(){ return buildWithdrawalForm_(config); }},
    {kind:'会員情報更新', key:FORM_KEY.UPDATE, handler:'handleUpdateSubmit', create:function(){ return buildUpdateForm_(config); }},
    {kind:'役員承諾', key:FORM_KEY.OFFICER, handler:'handleOfficerSubmit', create:function(){ return buildOfficerForm_(config); }}
  ];
  forms.forEach(function(def) {
    const form = openOrCreateForm_(def, config);
    removeTriggersForHandler_(def.handler);
    ScriptApp.newTrigger(def.handler).forForm(form).onFormSubmit().create();
    setConfigValue_(def.key, form.getId());
  });
  ptaWriteFormUrls();
  appendLog_('system','フォーム作成','', 'フォーム作成・トリガー設定', '成功', '既存IDが有効な場合は再利用');
  notifyAdmin_('PTAフォーム設定が完了しました', 'v0.3.1のフォーム作成・トリガー設定が完了しました。フォームURLシートを確認してください。');
}

function ptaWriteFormUrls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.LINKS);
  const config = getConfig_();
  clearDataRows_(sheet);
  const defs = [
    ['入会申込', FORM_KEY.JOIN],
    ['退会通知', FORM_KEY.LEAVE],
    ['会員情報更新', FORM_KEY.UPDATE],
    ['役員承諾', FORM_KEY.OFFICER]
  ];
  defs.forEach(function(row) {
    const id = config[row[1]];
    if (!id) return;
    try {
      const f = FormApp.openById(id);
      appendObjectRow_(SHEET.LINKS, {
        'フォーム種別': row[0],
        'フォームID': id,
        '編集URL': f.getEditUrl(),
        '回答URL': f.getPublishedUrl(),
        '作成日時': new Date(),
        '備考': '学校には加入・非加入・未提出を共有しない'
      });
    } catch (err) {
      appendObjectRow_(SHEET.LINKS, {
        'フォーム種別': row[0],
        'フォームID': id,
        '編集URL': '',
        '回答URL': '',
        '作成日時': new Date(),
        '備考': 'フォームを開けません: ' + err.message
      });
    }
  });
}

function ptaInstallDeletionReminder() {
  removeTriggersForHandler_('notifyDeletionDue');
  ScriptApp.newTrigger('notifyDeletionDue').timeBased().everyDays(1).atHour(9).create();
  appendLog_('system','トリガー','', '削除予定通知トリガーを設定', '成功', '毎日9時頃');
}

function ptaListTriggers() {
  const triggers = ScriptApp.getProjectTriggers().map(function(t){
    return [new Date(), 'system', 'トリガー一覧', '', t.getHandlerFunction(), '確認', t.getEventType()];
  });
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.LOG);
  if (triggers.length) sh.getRange(sh.getLastRow()+1,1,triggers.length,triggers[0].length).setValues(triggers);
}

function ptaUninstallKitTriggers() {
  const handlers = ['handleJoinSubmit','handleWithdrawalSubmit','handleUpdateSubmit','handleOfficerSubmit','notifyDeletionDue'];
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (handlers.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  appendLog_('system','トリガー','', '本キット関連トリガー削除', '成功', handlers.join(','));
}

function ptaRunSelfCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = [];
  const now = new Date();
  Object.keys(HEADERS).forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) {
      out.push([now,'シート','必須シート: '+name,'NG','存在しない','ptaInstallOrRepairを実行']);
      return;
    }
    const current = sh.getRange(1,1,1,HEADERS[name].length).getValues()[0];
    const ok = HEADERS[name].every(function(h,i){ return current[i] === h; });
    out.push([now,'シート','ヘッダー: '+name, ok?'OK':'NG', ok?'一致':'不一致', ok?'':'ヘッダー修復が必要']);
  });
  const forbiddenSheets = ['非加入届','非会員児童一覧','未提出者一覧','未加入者一覧','学校名簿','学校名簿照合','役員免除申請','家庭事情'];
  forbiddenSheets.forEach(function(name){
    out.push([now,'禁止機能','禁止シート: '+name, ss.getSheetByName(name)?'NG':'OK', ss.getSheetByName(name)?'存在する':'存在しない', ss.getSheetByName(name)?'削除又は利用停止':'']);
  });
  const config = getConfig_();
  out.push([now,'設定','管理者メール', config.ADMIN_EMAIL ? 'OK':'注意', config.ADMIN_EMAIL ? config.ADMIN_EMAIL : '未設定', '通知不要なら空欄可']);
  ['JOIN_FORM_ID','LEAVE_FORM_ID','UPDATE_FORM_ID','OFFICER_FORM_ID'].forEach(function(k){
    out.push([now,'設定','フォームID '+k, config[k] ? 'OK':'注意', config[k] ? '設定済み' : '未設定', 'ptaCreateFormsで作成']);
  });
  const sh = ss.getSheetByName(SHEET.SELF_CHECK);
  clearDataRows_(sh);
  if (out.length) sh.getRange(2,1,out.length,out[0].length).setValues(out);
  appendLog_('system','自己点検','', '自己点検を実行', '成功', 'NG件数=' + out.filter(function(r){return r[3]==='NG';}).length);
}

function handleJoinSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const config = getConfig_();
    const data = responseToObject_(e);
    const no = nextSequence_(PREFIX_KEY.JOIN);
    appendObjectRow_(SHEET.MEMBERS, {
      '受付番号': no,
      '受付日時': new Date(),
      'ステータス': '入会中',
      '会員氏名': data['会員氏名'] || '',
      'ふりがな': data['ふりがな'] || '',
      'メール': data['メール'] || '',
      '電話番号': data['電話番号'] || '',
      '関係': data['児童生徒との関係'] || '',
      '児童生徒氏名': data['児童生徒氏名（必要最小限）'] || '',
      '学年': data['学年（必要最小限）'] || '',
      '組': data['組（必要最小限）'] || '',
      '会費年度': config.FEE_YEAR || String(new Date().getFullYear()),
      '会費対象': '対象',
      '個人情報同意': '同意済み',
      'フォーム回答ID': responseId_(e),
      '最終更新日時': new Date(),
      '備考': '学校名簿照合禁止'
    });
    appendObjectRow_(SHEET.FEES, {
      '会費ID': 'F-' + no,
      '作成日時': new Date(),
      '会員受付番号': no,
      '会費年度': config.FEE_YEAR || String(new Date().getFullYear()),
      '会費額': Number(config.FEE_AMOUNT || 0),
      '納入状況': '未確認',
      '備考': 'PTAが会員本人に直接案内。学校徴収金と混合しない'
    });
    appendLog_('form','入会受付',no,'入会申込を会員台帳へ登録','成功','');
    notifyAdmin_('PTA入会申込を受け付けました（' + no + '）', '受付番号：' + no + '\n会員氏名：' + (data['会員氏名'] || '') + '\nメール：' + (data['メール'] || '') + '\n\n未提出者・非加入者・学校名簿との照合は行わないでください。');
  } finally {
    lock.releaseLock();
  }
}

function handleWithdrawalSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const config = getConfig_();
    const data = responseToObject_(e);
    const no = nextSequence_(PREFIX_KEY.LEAVE);
    const memberNo = data['会員受付番号'] || '';
    const desired = data['退会希望日'] || '';
    const deleteDate = addDays_(new Date(), Number(config.DELETION_AFTER_DAYS || 365));
    appendObjectRow_(SHEET.WITHDRAWALS, {
      '退会受付番号': no,
      '受付日時': new Date(),
      '会員受付番号': memberNo,
      '会員氏名': data['会員氏名'] || '',
      'メール': data['メール'] || '',
      '退会希望日': desired,
      '処理状況': '未処理',
      '削除予定日': deleteDate,
      'フォーム回答ID': responseId_(e),
      '備考': '退会は許可制ではない。退会理由を求めない'
    });
    appendObjectRow_(SHEET.DELETION, {
      '対象ID': memberNo,
      '対象区分': '退会者',
      '氏名': data['会員氏名'] || '',
      '発生日': new Date(),
      '削除予定日': deleteDate,
      '状態': '未処理',
      '備考': '会計・監査上必要な保存期間後に削除'
    });
    appendLog_('form','退会通知',memberNo,'退会通知を登録','成功',no);
    notifyAdmin_('PTA退会通知を受け付けました（' + no + '）', '退会受付番号：' + no + '\n会員受付番号：' + memberNo + '\n\n退会は許可制ではありません。退会理由や家庭事情を求めないでください。');
  } finally {
    lock.releaseLock();
  }
}

function handleUpdateSubmit(e) {
  const data = responseToObject_(e);
  const no = nextSequence_(PREFIX_KEY.UPDATE);
  appendObjectRow_(SHEET.UPDATES, {
    '更新受付番号': no,
    '受付日時': new Date(),
    '会員受付番号': data['会員受付番号'] || '',
    '会員氏名': data['会員氏名'] || '',
    '新メール': data['新しいメール'] || '',
    '新電話番号': data['新しい電話番号'] || '',
    '変更内容': data['変更内容メモ'] || '',
    '処理状況': '管理者確認待ち',
    'フォーム回答ID': responseId_(e),
    '備考': '本人確認後に会員台帳を更新'
  });
  appendLog_('form','会員情報更新',data['会員受付番号'] || '', '更新受付を登録', '成功', no);
  notifyAdmin_('PTA会員情報更新を受け付けました（' + no + '）', '管理者が本人確認のうえ、会員台帳を更新してください。');
}

function handleOfficerSubmit(e) {
  const data = responseToObject_(e);
  const no = nextSequence_(PREFIX_KEY.OFFICER);
  appendObjectRow_(SHEET.OFFICERS, {
    '役員承諾番号': no,
    '受付日時': new Date(),
    '会員受付番号': data['会員受付番号'] || '',
    '氏名': data['氏名'] || '',
    '希望役職': data['希望役職'] || '',
    '承諾内容': data['承諾内容'] || '本人承諾',
    '個人情報規程確認': '確認済み',
    '守秘義務確認': '確認済み',
    '処理状況': '受付',
    'フォーム回答ID': responseId_(e),
    '備考': '免除申請・家庭事情管理には使用しない'
  });
  appendLog_('form','役員承諾',no,'役員承諾を登録','成功','');
  notifyAdmin_('PTA役員立候補・承諾を受け付けました（' + no + '）', '氏名：' + (data['氏名'] || '') + '\n希望役職：' + (data['希望役職'] || '') + '\n\nこれは本人承諾記録です。免除申請や家庭事情管理には使用しないでください。');
}

function notifyDeletionDue() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.DELETION);
  if (!sh || sh.getLastRow() < 2) return;
  const values = sh.getRange(2,1,sh.getLastRow()-1,HEADERS[SHEET.DELETION].length).getValues();
  const today = new Date();
  const due = values.filter(function(r){
    const deadline = r[4];
    const status = r[5];
    if (!(deadline instanceof Date)) return false;
    const days = Math.ceil((deadline.getTime() - today.getTime()) / 86400000);
    return status !== '削除済み' && days <= 30;
  });
  if (due.length) notifyAdmin_('PTA個人情報の保存期限確認が必要です', '削除予定又は期限経過の情報があります。\n対象件数：' + due.length + '\n\n削除予定シートを確認してください。');
}

function buildJoinForm_(config) {
  const form = FormApp.create((config.PTA_NAME || 'PTA') + ' 入会申込フォーム');
  form.setDescription('PTAは学校ではありません。PTAは任意加入の団体です。入会を希望する方のみ、このフォームを送信してください。入会しない場合は何も提出する必要はありません。このフォームは学校提出物ではありません。学校は加入・非加入・未提出を把握しません。');
  addRequiredConfirm_(form, 'PTAは任意加入の団体であることを確認しました');
  addRequiredConfirm_(form, '入会しない場合は何も提出する必要がないことを確認しました');
  addRequiredConfirm_(form, 'このフォームは学校提出物ではないことを確認しました');
  addRequiredConfirm_(form, '学校は加入・非加入・未提出を把握しないことを確認しました');
  form.addTextItem().setTitle('会員氏名').setRequired(true);
  form.addTextItem().setTitle('ふりがな').setRequired(false);
  addEmailItem_(form, 'メール', true);
  form.addTextItem().setTitle('電話番号').setRequired(false);
  form.addMultipleChoiceItem().setTitle('児童生徒との関係').setChoiceValues(['保護者','教職員','その他']).setRequired(true);
  form.addTextItem().setTitle('児童生徒氏名（必要最小限）').setRequired(false);
  form.addListItem().setTitle('学年（必要最小限）').setChoiceValues(['未就学','1','2','3','4','5','6','中1','中2','中3','高1','高2','高3','その他']).setRequired(false);
  form.addTextItem().setTitle('組（必要最小限）').setRequired(false);
  const pp = config.PRIVACY_POLICY_URL ? '\nプライバシーポリシー: ' + config.PRIVACY_POLICY_URL : '';
  addRequiredConfirm_(form, '個人情報の利用目的・保存期間・学校名簿照合禁止を確認し、同意します' + pp);
  return form;
}

function buildWithdrawalForm_(config) {
  const form = FormApp.create((config.PTA_NAME || 'PTA') + ' 退会通知フォーム');
  form.setDescription('退会は許可制ではありません。このフォームは退会意思を記録し、会員台帳・会費処理・削除予定を整理するためのものです。退会理由や家庭事情は求めません。');
  form.addTextItem().setTitle('会員受付番号').setRequired(true);
  form.addTextItem().setTitle('会員氏名').setRequired(true);
  addEmailItem_(form, 'メール', true);
  form.addDateItem().setTitle('退会希望日').setRequired(false);
  addRequiredConfirm_(form, '退会は許可制ではないことを確認しました');
  addRequiredConfirm_(form, '退会理由の記載は不要であることを確認しました');
  return form;
}

function buildUpdateForm_(config) {
  const form = FormApp.create((config.PTA_NAME || 'PTA') + ' 会員情報更新フォーム');
  form.setDescription('会員本人の連絡先等を更新するためのフォームです。未提出者確認や学校名簿照合には使用しません。');
  form.addTextItem().setTitle('会員受付番号').setRequired(true);
  form.addTextItem().setTitle('会員氏名').setRequired(true);
  addEmailItem_(form, '新しいメール', false);
  form.addTextItem().setTitle('新しい電話番号').setRequired(false);
  form.addParagraphTextItem().setTitle('変更内容メモ').setRequired(false);
  return form;
}

function buildOfficerForm_(config) {
  const form = FormApp.create((config.PTA_NAME || 'PTA') + ' 役員立候補・承諾フォーム');
  form.setDescription('役員を引き受ける本人が送信するフォームです。くじ引き結果、免除申請、家庭事情の管理には使用しません。');
  form.addTextItem().setTitle('会員受付番号').setRequired(true);
  form.addTextItem().setTitle('氏名').setRequired(true);
  form.addTextItem().setTitle('希望役職').setRequired(true);
  form.addParagraphTextItem().setTitle('承諾内容').setRequired(false);
  addRequiredConfirm_(form, '本人の意思で役員就任を承諾します');
  addRequiredConfirm_(form, '個人情報取扱規程を確認します');
  addRequiredConfirm_(form, '守秘義務を確認します');
  return form;
}

function addRequiredConfirm_(form, text) {
  form.addCheckboxItem().setTitle('確認: ' + text).setChoiceValues(['確認しました']).setRequired(true);
}

function addEmailItem_(form, title, required) {
  const validation = FormApp.createTextValidation().requireTextIsEmail().setHelpText('メールアドレス形式で入力してください').build();
  form.addTextItem().setTitle(title).setValidation(validation).setRequired(!!required);
}

function openOrCreateForm_(def, config) {
  const existingId = config[def.key];
  if (existingId) {
    try { return FormApp.openById(existingId); } catch (err) { appendLog_('system','フォーム再作成','', def.kind, '注意', err.message); }
  }
  return def.create();
}

function ensureHeader_(sheet, headers) {
  const existing = sheet.getRange(1,1,1,headers.length).getValues()[0];
  const same = headers.every(function(h,i){ return existing[i] === h; });
  if (!same) sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

function seedConfig_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.CONFIG);
  const existing = sheet.getDataRange().getValues().slice(1).map(function(r){ return r[0]; });
  const rows = DEFAULT_CONFIG.filter(function(r){ return existing.indexOf(r[0]) === -1; });
  if (rows.length) sheet.getRange(sheet.getLastRow()+1,1,rows.length,3).setValues(rows);
}

function styleWorkbook_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const cols = HEADERS[name].length;
    sh.getRange(1,1,1,cols).setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');
    sh.autoResizeColumns(1, Math.min(cols, 12));
  });
}

function clearDataRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return;
  sheet.getRange(2,1,sheet.getLastRow()-1,Math.max(1,sheet.getLastColumn())).clearContent();
}

function getConfig_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.CONFIG);
  const config = {};
  if (!sheet || sheet.getLastRow() < 2) return config;
  sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues().forEach(function(r){ if (r[0]) config[String(r[0])] = r[1]; });
  return config;
}

function setConfigValue_(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.CONFIG);
  const values = sheet.getDataRange().getValues();
  for (let i=1; i<values.length; i++) {
    if (values[i][0] === key) { sheet.getRange(i+1,2).setValue(value); return; }
  }
  sheet.appendRow([key, value, '自動追加']);
}

function responseToObject_(e) {
  const obj = {};
  if (!e || !e.response) return obj;
  e.response.getItemResponses().forEach(function(ir) {
    let v = ir.getResponse();
    if (Array.isArray(v)) v = v.join(' / ');
    obj[ir.getItem().getTitle()] = v;
  });
  return obj;
}

function responseId_(e) {
  return e && e.response ? e.response.getId() : '';
}

function nextSequence_(prefixKey) {
  const config = getConfig_();
  const prefix = config[prefixKey] || 'NO-';
  const props = PropertiesService.getScriptProperties();
  const key = 'SEQ_' + prefixKey;
  const current = Number(props.getProperty(key) || '0') + 1;
  props.setProperty(key, String(current));
  return prefix + Utilities.formatString('%05d', current);
}

function appendObjectRow_(sheetName, obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('シートが見つかりません: ' + sheetName);
  const headers = HEADERS[sheetName] || sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const row = headers.map(function(h){ return obj[h] !== undefined ? obj[h] : ''; });
  sh.appendRow(row);
}

function appendLog_(operator, type, targetNo, action, result, note) {
  try {
    appendObjectRow_(SHEET.LOG, {'日時':new Date(),'操作者':operator,'種別':type,'対象番号':targetNo,'処理':action,'結果':result,'備考':note || ''});
  } catch (err) {
    console.error(err);
  }
}

function notifyAdmin_(subject, body) {
  const config = getConfig_();
  if (!config.ADMIN_EMAIL) return;
  MailApp.sendEmail(String(config.ADMIN_EMAIL), subject, body);
}

function removeTriggersForHandler_(handler) {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === handler) ScriptApp.deleteTrigger(t);
  });
}

function addDays_(date, days) {
  const d = new Date(date.getTime());
  d.setDate(d.getDate() + Number(days || 0));
  return d;
}
