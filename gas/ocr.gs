/**
 * 貸与書類の読み取り（Gemini API）
 * ------------------------------------------------------------
 * 紙で届く貸与書類をスマホで撮影し、シリアル・契約番号・契約日などを
 * 自動抽出するための処理。抽出結果はそのまま保存せず、必ずフロント側の
 * 確認フォームを経由して登録する。
 *
 * 導入手順は gas/README.md を参照。
 */

// 使用モデル。精度が足りない場合は 'gemini-2.5-pro' に変更する。
const OCR_MODEL = 'gemini-2.5-flash';
const GEMINI_ENDPOINT = 'https://generativelanguage.googleapis.com/v1beta/models/' + OCR_MODEL + ':generateContent';

const OCR_PROMPT = [
  'あなたは社内の貸与物管理担当者です。添付された貸与書類（納品書・貸与明細など）の画像から、',
  '貸与された機器の情報を抽出してください。',
  '',
  '【ルール】',
  '- 書類に記載されている機器を1台ずつ配列の要素にすること。同じ書類に複数台あれば全て挙げる。',
  '- 記載が無い項目は空文字にすること。推測で埋めてはいけない。',
  '- シリアル番号・製番は書類の記載を1文字も変えずに転記すること。',
  '- 文字が潰れていて判読できない箇所がある場合は、その機器の confidence を "low" にし、',
  '  note に「シリアル4文字目が不鮮明」のように具体的に書くこと。',
  '- 日付は必ず yyyy/mm/dd 形式に変換すること（令和などの和暦も西暦に直す）。',
  '- kind は次のいずれかを選ぶ: PC / 携帯 / Wi-Fi / Mac付属品',
  '  ノートパソコン・デスクトップは PC、携帯電話・スマートフォンは 携帯、',
  '  モバイルルーターは Wi-Fi、電源アダプタや変換アダプタなどの付属品は Mac付属品。',
  '- tel と iccid は 携帯 の場合のみ記入する。'
].join('\n');

const OCR_SCHEMA = {
  type: 'OBJECT',
  properties: {
    devices: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          kind:       {type: 'STRING', description: 'PC / 携帯 / Wi-Fi / Mac付属品'},
          serial:     {type: 'STRING', description: 'シリアル番号（携帯の場合は製番）'},
          contractNo: {type: 'STRING', description: '契約番号'},
          startDate:  {type: 'STRING', description: '契約開始日 yyyy/mm/dd'},
          endDate:    {type: 'STRING', description: '契約終了日 yyyy/mm/dd'},
          tel:        {type: 'STRING', description: '電話番号（携帯のみ）'},
          iccid:      {type: 'STRING', description: 'ICCID（携帯のみ）'},
          note:       {type: 'STRING', description: '備考・判読できなかった箇所'},
          confidence: {type: 'STRING', description: 'high / low'}
        },
        required: ['kind', 'serial']
      }
    }
  },
  required: ['devices']
};

/**
 * doPost から呼ぶ本体。
 * body = {action:'ocrDocument', images:[{mimeType:'image/jpeg', data:'<base64>'}, ...]}
 * 戻り値 = {devices:[...]}
 */
function ocrDocument_(body) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY がスクリプトプロパティに設定されていません');

  const images = body.images || [];
  if (!images.length) throw new Error('画像が送信されていません');
  if (images.length > 5) throw new Error('画像は一度に5枚までにしてください');

  const parts = [{text: OCR_PROMPT}];
  images.forEach(function (img) {
    if (!img || !img.data) throw new Error('画像データが不正です');
    parts.push({inline_data: {mime_type: img.mimeType || 'image/jpeg', data: img.data}});
  });

  const payload = {
    contents: [{role: 'user', parts: parts}],
    generationConfig: {
      temperature: 0,                      // 転記作業なので揺らぎを排除する
      responseMimeType: 'application/json',
      responseSchema: OCR_SCHEMA
    }
  };

  const res = UrlFetchApp.fetch(GEMINI_ENDPOINT + '?key=' + encodeURIComponent(key), {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code !== 200) throw new Error('Gemini APIエラー(' + code + '): ' + text.slice(0, 300));

  const json = JSON.parse(text);
  const cand = json.candidates && json.candidates[0];
  const out = cand && cand.content && cand.content.parts && cand.content.parts[0] && cand.content.parts[0].text;
  if (!out) throw new Error('書類から情報を読み取れませんでした。撮り直してください。');

  const parsed = JSON.parse(out);
  return {devices: (parsed.devices || []).map(normalizeOcrDevice_)};
}

/** 抽出結果の表記ゆれをシートの形式に揃える */
function normalizeOcrDevice_(d) {
  return {
    kind:       KIND_ALIAS_[String(d.kind || '').trim()] || String(d.kind || '').trim() || 'PC',
    serial:     String(d.serial || '').trim().toUpperCase().replace(/\s+/g, ''),
    contractNo: String(d.contractNo || '').trim(),
    startDate:  normalizeOcrDate_(d.startDate),
    endDate:    normalizeOcrDate_(d.endDate),
    tel:        String(d.tel || '').trim(),
    iccid:      String(d.iccid || '').trim().replace(/\s+/g, ''),
    note:       String(d.note || '').trim(),
    confidence: d.confidence === 'low' ? 'low' : 'high'
  };
}

const KIND_ALIAS_ = {
  'パソコン': 'PC', 'ノートパソコン': 'PC', 'ノートPC': 'PC', 'PC': 'PC',
  '携帯': '携帯', 'スマートフォン': '携帯', '携帯電話': '携帯', 'スマホ': '携帯',
  'WiFi': 'Wi-Fi', 'Wi-Fi': 'Wi-Fi', 'モバイルルーター': 'Wi-Fi', 'ポケットWiFi': 'Wi-Fi',
  '付属品': 'Mac付属品', 'Mac付属品': 'Mac付属品'
};

/** 2025-04-01 / 2025.4.1 / 令和7年4月1日 などを 2025/04/01 に寄せる */
function normalizeOcrDate_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  const wareki = s.match(/令和\s*(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日/);
  if (wareki) {
    return pad4_(2018 + Number(wareki[1])) + '/' + pad2_(wareki[2]) + '/' + pad2_(wareki[3]);
  }
  const m = s.match(/(\d{4})\D{1,3}(\d{1,2})\D{1,3}(\d{1,2})/);
  if (!m) return s;   // 解釈できないものは原文のまま返し、画面で目視修正させる
  return m[1] + '/' + pad2_(m[2]) + '/' + pad2_(m[3]);
}
function pad2_(n) { return String(Number(n)).padStart(2, '0'); }
function pad4_(n) { return String(Number(n)); }
