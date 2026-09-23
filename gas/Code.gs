// ==== 事前設定（GASエディタの「プロジェクトの設定」→「スクリプト プロパティ」に設定） ====
//   GITHUB_TOKEN     : GitHub Personal Access Token（対象リポジトリへの Contents 書き込み権限）
//   GITHUB_OWNER     : リポジトリのオーナー名（例: fe1000mg）
//   GITHUB_REPO      : リポジトリ名（例: -）
//   GITHUB_ICON_PATH : アイコン画像を置くパス（例: 生活/在庫管理アプリ/icons/）
//   GITHUB_BRANCH    : コミット先ブランチ（未設定なら main）
//
// スプレッドシートには「CategoryColors」というシートを追加してください（無ければ自動作成されます）。
// 1行目のヘッダー: categorySub, color

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'list';
  if (action === 'colors') {
    return getColors_();
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  const items = rows
    .filter(r => r[0])
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });
  return ContentService.createTextOutput(JSON.stringify(items))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  if (data.action === 'uploadIcon') {
    return uploadIconToGithub_(data);
  }
  if (data.action === 'setColor') {
    return setColor_(data);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');

  if (data.action === 'upsert') {
    const item = data.item;
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][idCol] === item.id) { rowIndex = i + 1; break; }
    }
    const row = headers.map(h => item[h] !== undefined ? item[h] : '');
    if (rowIndex === -1) {
      sheet.appendRow(row);
    } else {
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    }
  } else if (data.action === 'delete') {
    for (let i = 1; i < values.length; i++) {
      if (values[i][idCol] === data.id) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- 小分類ごとのアイコン色 ----
function getColors_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CategoryColors');
  const result = {};
  if (sheet) {
    const rows = sheet.getDataRange().getValues();
    rows.shift();
    rows.forEach(r => { if (r[0]) result[r[0]] = r[1]; });
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function setColor_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('CategoryColors');
  if (!sheet) {
    sheet = ss.insertSheet('CategoryColors');
    sheet.appendRow(['categorySub', 'color']);
  }
  const values = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.categorySub) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1) {
    sheet.appendRow([data.categorySub, data.color]);
  } else {
    sheet.getRange(rowIndex, 2).setValue(data.color);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- アイコン画像をGitHubリポジトリへ直接コミット ----
// GitHubのトークンはこの関数（サーバーサイド）にのみ保持され、ブラウザ側には一切渡らない。
function uploadIconToGithub_(data) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('GITHUB_TOKEN');
  const owner = props.getProperty('GITHUB_OWNER');
  const repo = props.getProperty('GITHUB_REPO');
  const basePath = props.getProperty('GITHUB_ICON_PATH') || 'icons/';
  const branch = props.getProperty('GITHUB_BRANCH') || 'main';

  if (!token || !owner || !repo) {
    return jsonOutput_({ ok: false, error: 'GitHub連携が未設定です（スクリプトプロパティを確認してください）' });
  }
  if (!data.name || !data.base64) {
    return jsonOutput_({ ok: false, error: '品名または画像データがありません' });
  }

  const path = basePath.replace(/\/?$/, '/') + data.name + '.png';
  const encodedPath = path.split('/').map(encodeURIComponent).join('/');
  const apiUrl = 'https://api.github.com/repos/' + owner + '/' + repo + '/contents/' + encodedPath;
  const headers = {
    Authorization: 'token ' + token,
    'User-Agent': 'inventory-app-gas'
  };

  let sha = null;
  const getRes = UrlFetchApp.fetch(apiUrl + '?ref=' + encodeURIComponent(branch), {
    headers: headers,
    muteHttpExceptions: true
  });
  if (getRes.getResponseCode() === 200) {
    sha = JSON.parse(getRes.getContentText()).sha;
  }

  const payload = {
    message: 'Update icon: ' + data.name,
    content: data.base64,
    branch: branch
  };
  if (sha) payload.sha = sha;

  const putRes = UrlFetchApp.fetch(apiUrl, {
    method: 'put',
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const ok = putRes.getResponseCode() === 200 || putRes.getResponseCode() === 201;
  return jsonOutput_({ ok: ok, error: ok ? '' : putRes.getContentText() });
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
