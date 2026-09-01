/**
 * nv-devices v2: 貸与物管理システム GASバックエンド（人ベース構造）
 *
 * シート構成:
 * - 社員: 社員マスタ（在籍状況を含む）
 * - 通信機器: PC/Wi-Fi/Mac付属品（SSマーケットレンタル）
 * - 社用携帯: au携帯
 * - ヘッドフォン: 人に紐づく初回貸与日で2年交換判定
 * - 追加備品: モニター等、人によって異なる備品
 * - 種別マスタ: 追加備品の種別リスト（自由に追加可能）
 * - 設定: メール署名など
 *
 * セットアップ:
 * 1. 新規スプレッドシート作成 → Apps Script にこのコードを貼り付け
 * 2. setupInitialData() を1回実行
 * 3. デプロイ > ウェブアプリ（実行: 自分 / アクセス: 全員）
 * 4. URLを index.html の GAS_URL に設定
 */

// ============================================================
// Web API
// ============================================================

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'getData';
  let result;
  try {
    if (action === 'getData') {
      result = getData();
    } else {
      result = { error: 'unknown action: ' + action };
    }
  } catch (err) {
    result = { error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case 'updateStatus':
        result = updateStatus(body.sheet, body.key, body.status);
        break;
      case 'updateEmployeeStatus':
        result = updateEmployeeStatus(body.name, body.status);
        break;
      case 'transferDevice':
        result = transferDevice(body.sheet, body.key, body.newOwner, body.oldOwner);
        break;
      case 'addExtraItem':
        result = addExtraItem(body.owner, body.itemType, body.note);
        break;
      case 'removeExtraItem':
        result = removeExtraItem(body.rowId);
        break;
      case 'addEmployee':
        result = addEmployee(body.name, body.note, body.hfDate);
        break;
      case 'addDevice':
        result = addDevice(body.sheet, body.row);
        break;
      case 'updateRow':
        result = updateRow(body.sheet, body.keyCol, body.key, body.values);
        break;
      case 'deleteRow':
        result = deleteRowByKey(body.sheet, body.keyCol, body.key);
        break;
      case 'ocrDocument':          // 貸与書類の読み取り（ocr.gs）
        result = ocrDocument_(body);
        break;
      default:
        result = { error: 'unknown action: ' + body.action };
    }
  } catch (err) {
    result = { error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// データ取得
// ============================================================

function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    employees: sheetToObjects(ss.getSheetByName('社員')),
    devices: sheetToObjects(ss.getSheetByName('通信機器')),
    mobiles: sheetToObjects(ss.getSheetByName('社用携帯')),
    headphones: sheetToObjects(ss.getSheetByName('ヘッドフォン')),
    extras: sheetToObjects(ss.getSheetByName('追加備品'), true),
    itemTypes: sheetToObjects(ss.getSheetByName('種別マスタ')),
    config: getConfig(ss)
  };
}

function sheetToObjects(sheet, withRowId) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = {};
    let hasData = false;
    for (let j = 0; j < headers.length; j++) {
      let v = values[i][j];
      if (v instanceof Date) {
        v = Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd');
      } else {
        v = String(v);
      }
      if (v !== '') hasData = true;
      row[headers[j]] = v;
    }
    if (hasData) {
      if (withRowId) row['_row'] = i + 1;
      rows.push(row);
    }
  }
  return rows;
}

function getConfig(ss) {
  const sheet = ss.getSheetByName('設定');
  if (!sheet) return {};
  const values = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) config[String(values[i][0])] = String(values[i][1]);
  }
  return config;
}

// ============================================================
// データ更新
// ============================================================

/** 機器のステータス更新（返却チェック / ストック化） */
function updateStatus(sheetName, key, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { error: 'sheet not found: ' + sheetName };

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const keyCol = keyColumnFor(sheetName, headers);
  const statusCol = headers.indexOf('ステータス');
  if (keyCol < 0 || statusCol < 0) return { error: 'column not found' };

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][keyCol]) === String(key)) {
      sheet.getRange(i + 1, statusCol + 1).setValue(status);
      return { ok: true };
    }
  }
  return { error: 'key not found: ' + key };
}

/** 社員の在籍ステータス更新（在籍中/退職予定/退職済み） */
function updateEmployeeStatus(name, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('社員');
  if (!sheet) return { error: 'sheet not found: 社員' };
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const nameCol = headers.indexOf('名前');
  const statusCol = headers.indexOf('在籍状況');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][nameCol]) === String(name)) {
      sheet.getRange(i + 1, statusCol + 1).setValue(status);
      return { ok: true };
    }
  }
  return { error: 'employee not found: ' + name };
}

/** 機器のスライド（保有者変更＋履歴を備考に自動追記） */
function transferDevice(sheetName, key, newOwner, oldOwner) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { error: 'sheet not found: ' + sheetName };

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const keyCol = keyColumnFor(sheetName, headers);
  const ownerCol = headers.indexOf(sheetName === 'ヘッドフォン' ? '名前' : '保有者');
  const noteCol = headers.indexOf('備考');
  const statusCol = headers.indexOf('ステータス');
  if (keyCol < 0 || ownerCol < 0) return { error: 'column not found' };

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][keyCol]) === String(key)) {
      sheet.getRange(i + 1, ownerCol + 1).setValue(newOwner);
      if (noteCol >= 0 && oldOwner) {
        const oldNote = String(values[i][noteCol] || '');
        const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM');
        const newNote = (oldNote ? oldNote + ' / ' : '') + '旧' + oldOwner + '（' + stamp + 'スライド）';
        sheet.getRange(i + 1, noteCol + 1).setValue(newNote);
      }
      if (statusCol >= 0) {
        sheet.getRange(i + 1, statusCol + 1).setValue('使用中');
      }
      return { ok: true };
    }
  }
  return { error: 'key not found: ' + key };
}

/** 追加備品の登録 */
function addExtraItem(owner, itemType, note) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('追加備品');
  if (!sheet) return { error: 'sheet not found: 追加備品' };
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  sheet.appendRow([owner, itemType, today, note || '', '使用中']);
  return { ok: true };
}

/** 追加備品の削除（行番号指定） */
function removeExtraItem(rowId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('追加備品');
  if (!sheet) return { error: 'sheet not found: 追加備品' };
  const row = Number(rowId);
  if (!row || row < 2) return { error: 'invalid row' };
  sheet.deleteRow(row);
  return { ok: true };
}

function keyColumnFor(sheetName, headers) {
  if (sheetName === '通信機器') return headers.indexOf('シリアル');
  if (sheetName === '社用携帯') return headers.indexOf('製番');
  if (sheetName === 'ヘッドフォン') return headers.indexOf('名前');
  return -1;
}

/** 社員追加（社員シート＋ヘッドフォンシートに行を追加） */
function addEmployee(name, note, hfDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const empSheet = ss.getSheetByName('社員');
  if (!empSheet) return { error: 'sheet not found: 社員' };

  // 重複チェック
  const values = empSheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(name)) {
      return { error: '同名の社員がすでに存在します: ' + name };
    }
  }

  empSheet.appendRow([name, '在籍中', note || '']);

  // ヘッドフォンシートにも行追加（貸与日は任意）
  const hfSheet = ss.getSheetByName('ヘッドフォン');
  if (hfSheet) {
    hfSheet.appendRow([name, hfDate || '', '', '', '', '', '', '', '使用中']);
  }
  return { ok: true };
}

/** 機器の新規登録（列名→値のオブジェクトで受け取り） */
function addDevice(sheetName, rowObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { error: 'sheet not found: ' + sheetName };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);

  // キー重複チェック
  const keyCol = keyColumnFor(sheetName, headers);
  if (keyCol >= 0) {
    const keyName = headers[keyCol];
    const keyVal = rowObj[keyName];
    if (keyVal) {
      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][keyCol]) === String(keyVal)) {
          return { error: keyName + ' が重複しています: ' + keyVal };
        }
      }
    }
  }

  const newRow = headers.map(h => rowObj[h] !== undefined ? rowObj[h] : '');
  const rowIndex = sheet.getLastRow() + 1;
  sheet.getRange(rowIndex, 1, 1, headers.length).setNumberFormat('@').setValues([newRow]);
  return { ok: true };
}

/** 行の更新（キー列で行を特定し、指定された列だけ書き換え） */
function updateRow(sheetName, keyColName, key, valuesObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { error: 'sheet not found: ' + sheetName };
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const keyCol = headers.indexOf(keyColName);
  if (keyCol < 0) return { error: 'key column not found: ' + keyColName };

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][keyCol]) === String(key)) {
      for (const colName in valuesObj) {
        const col = headers.indexOf(colName);
        if (col >= 0) {
          sheet.getRange(i + 1, col + 1).setNumberFormat('@').setValue(valuesObj[colName]);
        }
      }
      return { ok: true };
    }
  }
  return { error: 'key not found: ' + key };
}

/** 行の削除（キー列で特定） */
function deleteRowByKey(sheetName, keyColName, key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { error: 'sheet not found: ' + sheetName };
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const keyCol = headers.indexOf(keyColName);
  if (keyCol < 0) return { error: 'key column not found: ' + keyColName };

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][keyCol]) === String(key)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { error: 'key not found: ' + key };
}

// ============================================================
// 初期セットアップ（1回だけ実行）
// ============================================================

function setupInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---- 社員シート ----
  let sheet = getOrCreateSheet(ss, '社員');
  sheet.clear();
  const empHeaders = ['名前', '在籍状況', '備考'];
  const employees = [
    ['今村 邦之', '在籍中', ''],
    ['田開 友規', '在籍中', ''],
    ['髙山 博樹', '在籍中', ''],
    ['三宮 洋太', '在籍中', ''],
    ['冨谷 四季', '在籍中', ''],
    ['齊藤 麻莉', '在籍中', ''],
    ['小川 さつき', '在籍中', ''],
    ['正嵜 あずさ', '在籍中', ''],
    ['進藤 真世', '在籍中', ''],
    ['畠山 直人', '在籍中', ''],
    ['長澤 咲', '在籍中', ''],
    ['市嶋 拓也', '在籍中', ''],
    ['大河原 拓実', '在籍中', ''],
    ['壽美 晃二郎', '在籍中', ''],
    ['岩田 早織', '在籍中', ''],
    ['原田 将寛', '在籍中', ''],
    ['小池 拓哉', '在籍中', ''],
    ['十亀 奈津美', '在籍中', ''],
    ['藤澤', '在籍中', ''],
    ['渡邉 隼人', '在籍中', ''],
    ['村野 佳世', '在籍中', ''],
    ['芳賀 航', '在籍中', ''],
    ['関口 純', '在籍中', ''],
    ['林 佑樹', '在籍中', ''],
    ['粕谷 優希', '在籍中', ''],
    ['岸本 峻伍', '在籍中', ''],
    ['石原 朱理', '在籍中', ''],
    ['折原 千尋', '在籍中', ''],
    ['上野 隼弥', '在籍中', ''],
    ['貝瀬 英行', '在籍中', ''],
    ['牧川 倫子', '在籍中', ''],
    ['勝間田 雄大', '在籍中', ''],
    ['小西 輝', '在籍中', ''],
    ['並木 奨平', '在籍中', ''],
    ['六鹿 比斗志', '在籍中', ''],
    ['板井 和弥', '在籍中', ''],
    ['橘 百惠', '在籍中', ''],
    ['半田 樹梨', '在籍中', ''],
    ['清原 以於理', '在籍中', ''],
    ['金田 明子', '在籍中', ''],
    ['小島 一輝', '退職済み', ''],
    ['幡野 菜月', '退職済み', ''],
    ['鈴木 萌子', '退職済み', ''],
    ['小田切 咲', '退職済み', ''],
    ['派遣', '在籍中', '派遣社員'],
    ['貸出用', '在籍中', 'ストック管理用']
  ];
  sheet.getRange(1, 1, 1, empHeaders.length).setValues([empHeaders]).setFontWeight('bold');
  sheet.getRange(2, 1, employees.length, empHeaders.length).setNumberFormat('@').setValues(employees);
  sheet.setFrozenRows(1);

  // ---- 通信機器シート ----
  sheet = getOrCreateSheet(ss, '通信機器');
  sheet.clear();
  const deviceHeaders = ['保有者', '種別', '商品名', 'シリアル', '契約番号', '契約開始', '契約終了', '備考', 'ステータス'];
  const devices = [
    ['三宮 洋太', 'PC', 'DELL Pro 13Plus', '60249753', 'B25021241', '2025/09/01', '2027/09/30', '', '使用中'],
    ['今村 邦之', 'PC', 'HP EliteBook 630 G9', '30716759', 'B24024556', '2024/09/01', '2027/06/30', '', '使用中'],
    ['髙山 博樹', 'PC', 'HP EliteBook 630 G9', '31166409', 'B24024556', '2024/09/01', '2027/06/30', '', '使用中'],
    ['田開 友規', 'PC', 'HP EliteBook 630 G9', '30863552', 'B25016055', '2025/07/01', '2027/06/30', 'マウスのみ付属なし', '使用中'],
    ['長澤 咲', 'PC', 'Latitude 5320', '60024442', 'B24016491', '2024/07/01', '2027/06/30', '', '使用中'],
    ['壽美 晃二郎', 'PC', 'Latitude 5320', '60232250', 'B24016491', '2024/07/01', '2027/06/30', '', '使用中'],
    ['芳賀 航', 'PC', 'Latitude 5320', '60022400', 'B25007064', '2025/04/01', '2028/03/31', '', '使用中'],
    ['橘 百惠', 'PC', 'Latitude 5320', '60193834', 'B25008592', '2025/05/01', '2028/04/30', '旧ポイさんPC', '使用中'],
    ['関口 純', 'PC', 'Latitude 5320', '60193933', 'B25008592', '2025/05/01', '2028/04/30', '', '使用中'],
    ['正嵜 あずさ', 'PC', 'Latitude 5320', '60228949', 'B25018561', '2025/07/01', '2028/06/30', '', '使用中'],
    ['半田 樹梨', 'PC', 'Latitude 5320', '60047755', 'B25025555', '2025/10/01', '2028/09/30', '旧幡野さんPC', '使用中'],
    ['派遣', 'PC', 'Latitude 5320', '60063748', 'B25026265', '2025/10/01', '2028/09/30', 'セキュリティのみ設定', '使用中'],
    ['林 佑樹', 'PC', 'Latitude 5320', '60086334', 'B25029341', '2025/12/01', '2028/12/31', '', '使用中'],
    ['岸本 峻伍', 'PC', 'Latitude 5320', '60065704', 'B25035361', '2026/01/01', '2028/12/31', '', '使用中'],
    ['粕谷 優希', 'PC', 'Latitude 5320', '60068040', 'B25035361', '2026/01/01', '2028/12/31', '', '使用中'],
    ['折原 千尋', 'PC', 'Latitude 5320', '31089371', 'B25035363', '2026/02/01', '2029/01/31', '', '使用中'],
    ['石原 朱理', 'PC', 'Latitude 5320', '60085504', 'B25035363', '2026/02/01', '2029/01/31', '', '使用中'],
    ['上野 隼弥', 'PC', 'Latitude 5320', '60203601', 'B25035903', '2026/01/01', '2028/12/31', '2/12 鈴木さんPCをスライド', '使用中'],
    ['牧川 倫子', 'PC', 'Latitude 5320', '60230072', 'B26011408', '2026/06/01', '2029/05/31', '', '使用中'],
    ['勝間田 雄大', 'PC', 'Latitude 5320', '60047229', 'B26014496', '2026/06/01', '2029/05/31', '', '使用中'],
    ['小西 輝', 'PC', 'Latitude 5320', '60082183', 'B26014496', '2026/06/01', '2029/05/31', '', '使用中'],
    ['並木 奨平', 'PC', 'Latitude 5320', '60196477', 'B26014496', '2026/06/01', '2029/05/31', '', '使用中'],
    ['六鹿 比斗志', 'PC', 'Latitude 5320', '60196576', 'B26014496', '2026/06/01', '2029/05/31', '7/1～', '使用中'],
    ['板井 和弥', 'PC', 'Latitude 5320', '60059307', 'B26014496', '2026/06/01', '2029/05/31', '7/1～', '使用中'],
    ['清原 以於理', 'PC', 'Latitude 5320', '60067685', 'B26014496', '2026/09/01', '', '9/1～', '使用中'],
    ['貸出用', 'PC', 'MacBook Pro 2022 M2', '31150217', 'B24016287', '2024/07/15', '2027/07/14', '', '使用中'],
    ['冨谷 四季', 'PC', 'MacBook Pro 2023 M2Pro', '31168151', 'B24024556', '2024/09/01', '2027/06/30', 'キーボード&マウスも貸与', '使用中'],
    ['村野 佳世', 'PC', 'MacBook Pro 14インチ M4', '31761925', 'B25003842', '2025/04/01', '2028/03/31', 'マウスは別途購入予定', '使用中'],
    ['冨谷 四季', 'Mac付属品', 'Apple Magic Keyboard', '31578615', 'B24024556', '2024/09/01', '2027/06/30', '', '使用中'],
    ['冨谷 四季', 'Mac付属品', 'Apple Magic Mouse 2', '31578103', 'B24024556', '2024/09/01', '2027/06/30', '', '使用中'],
    ['田開 友規', 'Wi-Fi', 'Pocket WiFi SoftBank 601HW', '31352178', 'B24024556', '2024/09/01', '2027/06/30', '', '使用中'],
    ['貸出用', 'Wi-Fi', 'Pocket WiFi SoftBank 602HW', '10047484', 'B24024556', '2024/09/01', '2027/06/30', '', '使用中'],
    ['貸出用', 'Wi-Fi', 'Pocket WiFi SoftBank 602HW', '90006807', 'B24024556', '2024/09/01', '2027/06/30', '', '使用中'],
    ['齊藤 麻莉', 'PC', 'ProBook 635 Aero G8', '30704862', 'B24016491', '2024/07/01', '2027/06/30', '', '使用中'],
    ['渡邉 隼人', 'PC', 'ProBook 635 Aero G8', '30704886', 'B24016491', '2024/07/01', '2027/06/30', '', '使用中'],
    ['畠山 直人', 'PC', 'ProBook 635 Aero G8（交換用）', '31166843', 'B24016491', '2024/07/01', '2027/06/30', '', '使用中'],
    ['原田 将寛', 'PC', 'ProBook 635 Aero G8', '31160179', 'B24016491', '2024/07/01', '2027/06/30', '', '使用中'],
    ['市嶋 拓也', 'PC', 'ProBook 635 Aero G8', '31160186', 'B24016491', '2024/07/01', '2027/06/30', '', '使用中'],
    ['十亀 奈津美', 'PC', 'ProBook 635 Aero G8', '31160346', 'B24016491', '2024/07/01', '2027/06/30', '', '使用中'],
    ['小池 拓哉', 'PC', 'ProBook 635 Aero G8', '31160995', 'B24016491', '2024/07/01', '2027/06/30', '', '使用中'],
    ['岩田 早織', 'PC', 'ProBook 635 Aero G8', '31165549', 'B24016491', '2024/07/01', '2027/06/30', '', '使用中'],
    ['大河原 拓実', 'PC', 'ProBook 635 Aero G8', '31165761', 'B24016491', '2024/07/01', '2027/06/30', '', '使用中'],
    ['藤澤', 'PC', 'ProBook 635 Aero G8', '31166867', 'B24016491', '2024/07/01', '2027/06/30', '', '使用中'],
    ['進藤 真世', 'PC', 'ProBook 635 Aero G8', '31160100', 'B24024556', '2024/09/01', '2027/06/30', '', '使用中']
  ];
  sheet.getRange(1, 1, 1, deviceHeaders.length).setValues([deviceHeaders]).setFontWeight('bold');
  sheet.getRange(2, 1, devices.length, deviceHeaders.length).setNumberFormat('@').setValues(devices);
  sheet.setFrozenRows(1);

  // ---- 社用携帯シート ----
  sheet = getOrCreateSheet(ss, '社用携帯');
  sheet.clear();
  const mobileHeaders = ['保有者', '商品名', '製番', '電話番号', 'ICCID', '申込番号', 'ステータス'];
  const mobiles = [
    ['並木 奨平', 'au SCG33', '351781751732249', '080-6864-3159', '8981300121318512765', 'EA8023573-00', '使用中'],
    ['勝間田 雄大', 'au SCG33', '351781751745506', '080-6867-7164', '8981300121318512773', 'EAA024679-00', '使用中'],
    ['牧川 倫子', 'au SCG33', '351781751745498', '080-6864-3158', '8981300121318512781', 'EAA224344-00', '使用中'],
    ['小西 輝', 'au SCG33', '351781750587180', '080-6784-0431', '8981300171337636129', 'EA8675150-00', '使用中'],
    ['芳賀 航', 'au SCG33', '351781750591976', '080-6784-0433', '8981300171337636178', 'EA7674750-00', '使用中'],
    ['長澤 咲', 'au SCG33', '351781750604399', '080-6782-9436', '8981300171337636103', 'EA9775001-00', '使用中'],
    ['十亀 奈津美', 'au SCG33', '351781750725442', '080-6782-9441', '8981300171337636202', 'EAE583583-00', '使用中'],
    ['小川 さつき', 'au SCG33', '351781750604449', '080-6783-9024', '8981300171337636111', 'EA9575974-00', '使用中'],
    ['小池 拓哉', 'au SCG33', '351781750558306', '080-6783-9025', '8981300171337636400', 'EAD685415-00', '使用中'],
    ['林 佑樹', 'au SCG33', '351781750587347', '080-6783-9026', '8981300171337636137', 'EA8675149-00', '使用中'],
    ['進藤 真世', 'au SCG33', '351781750668998', '080-6783-9027', '8981300171337636434', 'EAC173738-00', '使用中'],
    ['市嶋 拓也', 'au SCG33', '351781750575458', '080-6782-6645', '8981300171337636228', 'EAD884687-00', '使用中'],
    ['上野 隼弥', 'au SCG33', '351781750604407', '080-6782-9437', '8981300171337636145', 'EA8474576-00', '使用中'],
    ['壽美 晃二郎', 'au SCG33', '351781750669137', '080-6784-0436', '8981300171337636483', 'EAA275781-00', '使用中'],
    ['原田 将寛', 'au SCG33', '351781750669129', '080-6783-9021', '8981300171337636426', 'EAC375639-00', '使用中'],
    ['粕谷 優希', 'au SCG33', '351781750575417', '080-6782-6644', '8981300171337636491', 'EA9875124-00', '使用中'],
    ['岸本 峻伍', 'au SCG33', '351781750603623', '080-6783-9022', '8981300171337636194', 'EA7375048-00', '使用中'],
    ['六鹿 比斗志', 'au SCG33', '351781750558199', '080-6783-9023', '8981300171337636475', 'EAA574599-00', '使用中'],
    ['岩田 早織', 'au SCG33', '351781750575482', '080-6784-0434', '8981300171337636442', 'EAB375594-00', '使用中'],
    ['大河原 拓実', 'au SCG33', '351781750669087', '080-6782-6643', '8981300171337636210', 'EAE485355-00', '使用中'],
    ['渡邉 隼人', 'au SCG33', '351781750602245', '080-6782-9442', '8981300171337636418', 'EAD484351-00', '使用中'],
    ['畠山 直人', 'au SCG33', '351781750726465', '080-6782-9440', '8981300171337636467', 'EAA674396-00', '使用中'],
    ['折原 千尋', 'au SCG33', '351781750726267', '080-6782-9438', '8981300171337636459', 'EAA776886-00', '使用中'],
    ['板井 和弥', 'au SCG33', '351781750590101', '080-6782-9439', '8981300171337636152', 'EAA7975095-00', '使用中'],
    ['石原 朱理', 'au SCG33', '351781750587248', '080-6784-0435', '8981300171337636160', 'EA7775861-00', '使用中'],
    ['関口 純', 'au SCG33', '351781750590002', '080-6784-0432', '8981300171337636186', 'EA7375048-00', '使用中'],
    ['清原 以於理', 'au SCG33', '351781752418863', '080-9034-3121', '8981300171344229231', 'EAD652558', '使用中'],
    ['橘 百惠', 'au SCG33', '351781752432070', '080-9034-3145', '8981300171344229207', 'EAB244310', '使用中'],
    ['半田 樹梨', 'au SCG33', '351781752467118', '080-9034-3117', '8981300171344229249', 'EA7743245', '使用中']
  ];
  sheet.getRange(1, 1, 1, mobileHeaders.length).setValues([mobileHeaders]).setFontWeight('bold');
  sheet.getRange(2, 1, mobiles.length, mobileHeaders.length).setNumberFormat('@').setValues(mobiles);
  sheet.setFrozenRows(1);

  // ---- ヘッドフォンシート ----
  sheet = getOrCreateSheet(ss, 'ヘッドフォン');
  sheet.clear();
  const hfHeaders = ['名前', '初回貸与日', '1回目交換', '2回目交換', '3回目交換', '4回目交換', '5回目交換', '備考', 'ステータス'];
  const headphones = [
    ['今村 邦之', '', '', '', '', '', '', '', '使用中'],
    ['田開 友規', '', '2026/01/16', '', '', '', '', '', '使用中'],
    ['髙山 博樹', '', '', '', '', '', '', '', '使用中'],
    ['十亀 奈津美', '', '2026/04/22', '', '', '', '', '', '使用中'],
    ['進藤 真世', '', '2025/07/02', '', '', '', '', '', '使用中'],
    ['畠山 直人', '', '2025/09/29', '', '', '', '', '', '使用中'],
    ['長澤 咲', '', '2026/05/28', '', '', '', '', '', '使用中'],
    ['市嶋 拓也', '', '', '', '', '', '', '', '使用中'],
    ['大河原 拓実', '', '', '', '', '', '', '', '使用中'],
    ['三宮 洋太', '', '', '', '', '', '', '', '使用中'],
    ['壽美 晃二郎', '', '', '', '', '', '', '', '使用中'],
    ['岩田 早織', '', '2026/01/05', '', '', '', '', '', '使用中'],
    ['原田 将寛', '', '', '', '', '', '', '', '使用中'],
    ['小池 拓哉', '', '', '', '', '', '', '', '使用中'],
    ['冨谷 四季', '', '', '', '', '', '', '', '使用中'],
    ['齊藤 麻莉', '', '', '', '', '', '', '2024/11返却済', '返却済'],
    ['村野 佳世', '2025/04/01', '', '', '', '', '', '', '使用中'],
    ['芳賀 航', '2025/04/01', '', '', '', '', '', '', '使用中'],
    ['関口 純', '2025/05/01', '', '', '', '', '', '', '使用中'],
    ['正嵜 あずさ', '2025/09/11', '', '', '', '', '', '', '使用中'],
    ['幡野 菜月', '2025/10/01', '', '', '', '', '', '退職', '返却済'],
    ['林 佑樹', '2025/12/01', '', '', '', '', '', '', '使用中'],
    ['粕谷 優希', '2026/01/01', '', '', '', '', '', '', '使用中'],
    ['岸本 峻伍', '2026/01/01', '', '', '', '', '', '', '使用中'],
    ['鈴木 萌子', '2026/01/01', '', '', '', '', '', '退職', '返却済'],
    ['石原 朱理', '2026/02/01', '', '', '', '', '', '', '使用中'],
    ['折原 千尋', '2026/02/01', '', '', '', '', '', '', '使用中'],
    ['上野 隼弥', '2026/04/01', '', '', '', '', '', '', '使用中'],
    ['渡邉 隼人', '2026/05/01', '', '', '', '', '', '', '使用中'],
    ['貝瀬 英行', '2026/05/01', '', '', '', '', '', '', '使用中'],
    ['牧川 倫子', '2026/06/01', '', '', '', '', '', '', '使用中'],
    ['勝間田 雄大', '2026/06/01', '', '', '', '', '', '', '使用中'],
    ['小西 輝', '2026/06/01', '', '', '', '', '', '', '使用中'],
    ['並木 奨平', '2026/06/01', '', '', '', '', '', '', '使用中'],
    ['六鹿 比斗志', '2026/07/01', '', '', '', '', '', '', '使用中'],
    ['板井 和弥', '2026/07/01', '', '', '', '', '', '', '使用中'],
    ['橘 百惠', '2026/08/01', '', '', '', '', '', '', '使用中'],
    ['半田 樹梨', '2026/08/01', '', '', '', '', '', '', '使用中'],
    ['清原 以於理', '2026/09/01', '', '', '', '', '', '', '使用中'],
    ['金田 明子', '', '', '', '', '', '', '', '使用中']
  ];
  sheet.getRange(1, 1, 1, hfHeaders.length).setValues([hfHeaders]).setFontWeight('bold');
  sheet.getRange(2, 1, headphones.length, hfHeaders.length).setNumberFormat('@').setValues(headphones);
  sheet.setFrozenRows(1);

  // ---- 追加備品シート ----
  sheet = getOrCreateSheet(ss, '追加備品');
  sheet.clear();
  const extraHeaders = ['保有者', '種別', '貸与日', '備考', 'ステータス'];
  sheet.getRange(1, 1, 1, extraHeaders.length).setValues([extraHeaders]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // ---- 種別マスタシート ----
  sheet = getOrCreateSheet(ss, '種別マスタ');
  sheet.clear();
  const typeData = [
    ['種別名'],
    ['モニター'],
    ['キーボード'],
    ['マウス'],
    ['ドッキングステーション'],
    ['外付けSSD'],
    ['Webカメラ']
  ];
  sheet.getRange(1, 1, typeData.length, 1).setValues(typeData);
  sheet.getRange(1, 1, 1, 1).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // ---- 設定シート ----
  sheet = getOrCreateSheet(ss, '設定');
  sheet.clear();
  const configData = [
    ['キー', '値'],
    ['署名者名', '齊藤 麻莉'],
    ['会社名', 'ナウビレッジ株式会社'],
    ['部署名', '管理部'],
    ['発注先会社', '株式会社SSマーケット'],
    ['発注先担当者', '田中様']
  ];
  sheet.getRange(1, 1, configData.length, 2).setValues(configData);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  sheet.setFrozenRows(1);

  SpreadsheetApp.flush();
  Logger.log('セットアップ完了');
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}
