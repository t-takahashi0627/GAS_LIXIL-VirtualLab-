// Webアプリとしてのエントリポイント
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Upload");
}

// CSVを処理してDriveにスプレッドシートを保存する
// CSVを処理してDriveにスプレッドシートを保存する
function processCSV(csvText, fileName) {
  const rows = Utilities.parseCsv(csvText);
  const processedData = processData(rows);

  // ファイル名が未指定ならデフォルト名を使用
  const sheetName = fileName || "CSV_" + new Date().toISOString().slice(0, 10);
  const spreadsheet = SpreadsheetApp.create(sheetName);
  const sheet = spreadsheet.getSheets()[0];

  sheet.getRange(1, 1, processedData.length, processedData[0].length).setValues(processedData);

  return spreadsheet.getUrl(); // URLを返す
}

// CSVデータを整形処理
function processData(data) {
  const headers = [
    "お客さま氏名", "お客さま氏名（フリガナ）", "お客さまTEL", "お客さま郵便番号", "お客さま都道府県",
    "お客さま市区町村町域", "ご依頼内容", "製品のブランド", "製品品番", "取付年月（年）", "取付年月（月）",
    "ご希望の訪問日がある場合（任意）なし", "ご希望の訪問日がある場合（任意）あり", "ご希望の訪問日",
    "LTS得意先コード（ご依頼元）", "会社名、部署", "ご依頼元担当者", "ご依頼元住所", "ご依頼元TEL", "ご依頼元FAX",
    "LTS得意先コード（ご請求先）", "LTS得意先コード（ご請求先）", "ご請求先", "ご請求先（フリガナ）", "ご請求先担当者", "ご請求先TEL"
  ];

  const result = [headers];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const newRow = [];

    newRow[0]  = row[5]  || "";
    newRow[1]  = row[0]  || "";
    newRow[2]  = row[6]  || "";

    // row[7] には「郵便番号＋住所」全体が入っている前提
    let postalAddress = row[7] || "";
    let postalCode = "";
    let prefecture = "";
    let cityAddress = "";

    // 郵便番号を抽出（例：123-4567）
    const postalMatch = postalAddress.match(/\d{3}-\d{4}/);
    postalCode = postalMatch ? postalMatch[0] : "478-0000";

    // 47都道府県一覧
    const prefectures = [
      "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
      "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
      "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県",
      "岐阜県", "静岡県", "愛知県", "三重県",
      "滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県",
      "鳥取県", "島根県", "岡山県", "広島県", "山口県",
      "徳島県", "香川県", "愛媛県", "高知県",
      "福岡県", "佐賀県", "長崎県", "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"
    ];

    // 都道府県を検索
    let matchedPref = "";
    for (const pref of prefectures) {
      if (postalAddress.includes(pref)) {
        matchedPref = pref;
        break;
      }
    }
    prefecture = matchedPref || "愛知県";

    // 市区町村以下の住所
    if (matchedPref) {
      const index = postalAddress.indexOf(matchedPref);
      cityAddress = postalAddress.slice(index + matchedPref.length).trim();
      if (!cityAddress) cityAddress = "不明";
    } else {
      cityAddress = "不明";
    }

    newRow[3] = postalCode;
    newRow[4] = prefecture;
    newRow[5] = cityAddress;

    newRow[6]  = row[18] || "";

    // 対応するブランド名（インデックス 10, 11, 12）
    const brandFlags = [row[10], row[11], row[12]];
    const brandLabels = ["LIXIL", "TOSTEM", "INAX"];

    // "1" が付いているブランド名だけを抽出
    const selectedBrands = brandFlags.map((flag, i) => flag === "1" ? brandLabels[i] : null).filter(Boolean);

    // 結果をセット（なければ "INAX"）
    newRow[7] = selectedBrands.length > 0 ? selectedBrands.join(" ") : "INAX";

    newRow[8]  = row[13] || "";
    newRow[9]  = row[14] || "";
    newRow[10] = "";
    newRow[11] = row[19] || "";
    newRow[12] = "";
    newRow[13] = "";
    newRow[14] = row[23] || "";
    newRow[15] = row[21] || "";
    newRow[16] = row[22] || "";
    newRow[17] = "";
    newRow[18] = "";
    newRow[19] = "";
    newRow[20] = "";
    newRow[21] = "";
    newRow[22] = "";
    newRow[23] = row[31] || "";
    newRow[24] = row[32] || "";
    newRow[25] = row[29] || "";

    result.push(newRow);
  }

  return result;
}