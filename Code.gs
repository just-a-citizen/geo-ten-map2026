/**
 * フォームが送信されたときに実行される関数
 */
function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  
  // フォームの回答内容を取得
  const pref = e.values[1];           // 都道府県
  const city = e.values[2] || "";     // 市区町村（空の可能性がある）
  const spotName = e.values[3];        
  const spotAddr = e.values[4] || "";  
  const reason = e.values[5];          

  // --- 出身地の座標取得ロジックを強化 ---
  let originSearchKey;
  if (city === "") {
    // 市区町村が空なら「茨城県 県庁」で検索して、確実に県庁所在地の座標を取る
    originSearchKey = pref + " 県庁";
  } else {
    originSearchKey = pref + city;
  }

  const originLoc = getCoordinates(originSearchKey);
  if (originLoc) {
    sheet.getRange(row, 7).setValue(originLoc.lat); 
    sheet.getRange(row, 8).setValue(originLoc.lng); 
  }

  // スポットの座標（ここは今まで通り）
  const spotSearchKey = spotAddr || (pref + " " + spotName);
  const spotLoc = getCoordinates(spotSearchKey);
  if (spotLoc) {
    sheet.getRange(row, 9).setValue(spotLoc.lat); 
    sheet.getRange(row, 10).setValue(spotLoc.lng); 
  }
}

/**
 * 住所から緯度経度を返す補助関数（日本国内に限定）
 */
function getCoordinates(address) {
  if (!address || address === "") return null;
  
  try {
    // 常に「日本」を頭につけることで、海外への誤爆を防ぐ
    const response = Maps.newGeocoder().geocode("日本 " + address);
    if (response.status === 'OK' && response.results.length > 0) {
      const result = response.results[0].geometry.location;
      return { lat: result.lat, lng: result.lng };
    }
  } catch (e) {
    console.log("エラー: " + e.message);
  }
  return null;
}

/**
 * すでにシートにあるデータに対して一括で緯度経度を補完する関数
 */
function fillMissingCoordinates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues(); // 全データを取得
  
  // 2行目から順番にチェック（1行目は見出しなので i=1 から開始）
  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const pref = data[i][1];      // B列
    const city = data[i][2] || ""; // C列
    const spotName = data[i][3];   // D列
    const spotAddr = data[i][4] || ""; // E列

    // 出身地の座標が空（G列かH列が空）なら取得
    if (!data[i][6] || !data[i][7]) {
      const originLoc = getCoordinates(pref + city);
      if (originLoc) {
        sheet.getRange(row, 7).setValue(originLoc.lat);
        sheet.getRange(row, 8).setValue(originLoc.lng);
      }
    }

    // スポットの座標が空（I列かJ列が空）なら取得
    if (!data[i][8] || !data[i][9]) {
      const spotSearchKey = spotAddr || (pref + " " + spotName);
      const spotLoc = getCoordinates(spotSearchKey);
      if (spotLoc) {
        sheet.getRange(row, 9).setValue(spotLoc.lat);
        sheet.getRange(row, 10).setValue(spotLoc.lng);
      }
    }
  }
}

// Webアプリとして表示するための設定（修正版）
function doGet(e) {
  // URLの末尾に ?page=relation と付けると新しいマップを表示
  const page = e.parameter.page || 'index'; 
  
  return HtmlService.createHtmlOutputFromFile(page)
      .setTitle('おすすめ穴場スポットマップ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// getMapData() は今のものをそのまま使えます（変更なし）

// スプレッドシートのデータをHTML側に送る関数
function getMapData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0]; 
  const data = sheet.getDataRange().getValues();

  // Dateオブジェクトなどが含まれるとエラーになるため、
  // 全てのデータを「文字列」や「数値」として扱えるように安全に処理する
  const safeData = data.map(row => {
    return row.map(cell => {
      if (cell instanceof Date) {
        return Utilities.formatDate(cell, "JST", "yyyy/MM/dd HH:mm:ss");
      }
      return cell;
    });
  });

  return safeData;
}