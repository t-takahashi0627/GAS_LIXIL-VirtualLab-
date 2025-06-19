// 郵便番号マスタ設定（IMPORTRANGEで使う）
const ZIP_MASTER_ID = '1FdicA2XbJCvcDUi5L2iNwTYsUwJy7VZClWibXwtdudk';
const ZIP_MASTER_SHEET = 'utf_ken_all_SJIS'; // シート名

/**
 * Webアプリとしてのエントリポイント
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Upload");
}

/**
 * CSVを受け取り、スプレッドシートに整形保存
 */
function processCSV(csvText, fileName, zipList, prefList, addrList) {
 const rows = Utilities.parseCsv(csvText);
  const processedData = processData(rows, zipList, prefList, addrList);

  // ファイル名が未指定ならデフォルト名を使用
  const sheetName = fileName || "CSV_" + new Date().toISOString().slice(0, 10);
  const spreadsheet = SpreadsheetApp.create(sheetName);
  const sheet = spreadsheet.getSheets()[0];

  sheet.getRange(1, 1, processedData.length, processedData[0].length).setValues(processedData);

  return spreadsheet.getUrl(); // URLを返す
}
 
/**
 * 郵便番号から住所を検索（マスタ参照）
 */
function lookupAddress(zipCode) {
  const sheet = SpreadsheetApp.openById(ZIP_MASTER_ID).getSheetByName(ZIP_MASTER_SHEET);
  const data = sheet.getRange('A:D').getValues();
  const key = zipCode.replace(/-/g, '');

  for (let i = 0; i < data.length; i++) {
    const rowZip = (data[i][0] + '').replace(/-/g, '');
    if (rowZip === key) {
      return {
        pref: data[i][1],
        addr: (data[i][2] || '') + (data[i][3] || '')
      };
    }
  }
  return null;
}

/**
 * CSVデータの整形処理
 */
function processData(data, zipList, prefList, addrList) {
  const headers = [
    "お客さま氏名", "お客さま氏名（フリガナ）", "お客さまTEL", "お客さま郵便番号", "お客さま都道府県",
    "お客さま市区町村町域", "ご依頼内容", "製品のブランド", "製品品番", "取付年月（年）", "取付年月（月）",
    "ご希望の訪問日がある場合（任意）なし", "ご希望の訪問日がある場合（任意）あり", "ご希望の訪問日",
    "LTS得意先コード（ご依頼元）", "会社名、部署", "ご依頼元担当者", "ご依頼元住所", "ご依頼元TEL", "ご依頼元FAX",
    "LTS得意先コード（ご請求先）", "LTS得意先コード（ご請求先）", "ご請求先", "ご請求先（フリガナ）", "ご請求先担当者", "ご請求先TEL"
  ];

  const result = [headers];

  // 入力値の整形
  //const zip = uiZip?.match(/^\d{7}$/) ? uiZip.replace(/^(\d{3})(\d{4})$/, '$1-$2') : (uiZip || '478-0000');
  //const pref = uiPref || "愛知県";
  //const addr = uiAddr || "不明";

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const newRow = [];

    // 郵便番号・都道府県・市区町村（それぞれの行に対応する値を使用）
    const zip = zipList?.[i - 1] || "478-0000";
    const pref = prefList?.[i - 1] || "愛知県";
    const addr = addrList?.[i - 1] || "不明";


    const val5 = (row[5] || "").trim();
    newRow[0] = val5 !== "" ? val5 : "不明";

    const val0 = (row[0] || "").trim();
    newRow[1] = val0 !== "" ? val0 : "フメイ";

    const tel = (row[6] || "").replace(/-/g, "").trim();
    newRow[2] = tel !== "" ? tel : "999999999";

    newRow[3] = zip;      // 郵便番号（HTML入力値）
    newRow[4] = pref;   // 都道府県（HTML入力値）
    newRow[5] = addr;   // 市区町村以下（HTML入力値）

    const val18 = (row[18] || "").trim();
    newRow[6] = val18 !== "" ? val18 : "修理依頼";

    const brandFlags = [row[10], row[11], row[12]];
    const brandLabels = ["LIXIL", "TOSTEM", "INAX"];
    const selectedIndexes = brandFlags.map((flag, i) => flag === "1" ? i : -1).filter(i => i !== -1);
    newRow[7] = selectedIndexes.length === 1 ? brandLabels[selectedIndexes[0]] : "INAX";

    const val13 = (row[13] || "").trim();
    newRow[8] = val13 !== "" ? val13 : "ﾌﾒｲ";

    const dateValue = (row[14] || "").trim();
    const dateMatch = dateValue.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
    newRow[9] = dateMatch ? dateMatch[1] : "";
    newRow[10] = dateMatch ? dateMatch[2].padStart(2, '0') : "";

    const val19 = (row[19] || "").trim();
    if (val19 !== "") {
      newRow[11] = 0;
      newRow[12] = 1;
      newRow[13] = val19;
    } else {
      newRow[11] = 1;
      newRow[12] = 0;
      newRow[13] = "";
    }

    //------------newRow14の処理---------------------
    const val23 = (row[23] || "").trim();
    newRow[14] = val23 !== "" ? val23 : "99999";

    //------------newRow15の処理---------------------
    const val21 = (row[21] || "").trim();
    newRow[15] = val21 !== "" ? val21 : "不明";

    //------------newRow16の処理---------------------
    const val22 = (row[22] || "").trim();
    newRow[16] = val22 !== "" ? val22 : "不明";

    //------------newRow17の処理---------------------
    newRow[17] = "不明";

    //------------newRow18の処理---------------------
    newRow[18] = "999999999";

    //------------newRow19の処理---------------------
    newRow[19] = "999999999";

    //------------newRow20の処理---------------------
    newRow[20] = "9999";

    //------------newRow21の処理---------------------
    newRow[21] = "9999";

    //------------newRow22の処理---------------------
    newRow[22] = "";


    //------------newRow23の処理---------------------
    const val31 = (row[31] || "").trim();
    newRow[23] = val31 !== "" ? val31 : "ﾌﾒｲ";

    //------------newRow24の処理---------------------
    const val32 = (row[32] || "").trim();
    newRow[24] = val32 !== "" ? val32 : "不明";

    //------------newRow25の処理---------------------
    const val29 = (row[29] || "").trim();
    newRow[25] = val29 !== "" ? val29 : "999999999";

    result.push(newRow);
  }

  return result;
}
