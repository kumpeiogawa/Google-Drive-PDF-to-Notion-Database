// 以下は必要に応じて取得の上入力してください
// NotionのAPIキー
const NOTION_API_KEY = "ntn_xxxxxxxxxxxxxxxxxxxx"
// Notion内のSurvey DatabaseのID
const NOTION_DB_ID = "xxxxxxxxxxxxxxxxx"
// GeminiのAPIキー
const GEMINI_API_KEY = "AIxxxxxxxxxxxxxxxxxxxxxx"
// 本スクリプトが走査する（論文PDFを探す）Google Drive FolderのID
const ROOT_FOLDER_ID = "xxxxxxxxxxxxxxxxxxxxxxxx"

// Geminiに投げるプロンプト
const prompt = `

Tell me the authors, authors' affiliations, title, journal/proceeding name, published year, summary (in Japanese), and DOI (if found) of this paper. 

Your summary must follow the following format, and each of them should be about 100-150 characters long in Japanese. 
"1. どんなもの？\n(your answer here)\n2. 先行研究と比べてどこがすごい？\n(your answer here)\n3. 技術や手法のキモはどこ？\n(your answer here)\n4. どうやって有効だと検証した？\n(your answer here)\n5. 議論はある？\n(your answer here)"

Remove all the citations from your answers.

Format your answer in json with keys of "title", "authors", "affiliation", "proc", "year", "summary", and "doi". Each value must be in one single string (not list). 
`

var successFolder = null;
var errorFolder = null;
var overwriteFolder = null;
var status = "";

// この関数を定期実行に設定してください
// 実行間隔は最小10分（それより短くする場合は下のコードを読んで最大実行時間を調整してください）
function refreshAllFiles() {

  const startTime = new Date();
  var num_files = 0;
  var num_processed_files = 0;

  try {
    var running = CacheService.getScriptCache().get('running');
    if (running != null) {
      throw new Error(`Script already running. Try again after a few minutes.`);
    }
    else {
      CacheService.getScriptCache().put('running', true);
    }

    updateStatus("Start running script");

    const pageList = getNotionPageList();

    // 各種フォルダの取得
    const rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
    const folders = rootFolder.getFolders();

    while (folders.hasNext()) {
      const folder = folders.next();
      if (folder.getName() == "success")
        successFolder = folder;
      if (folder.getName() == "error")
        errorFolder = folder;
      if (folder.getName() == "overwrite")
        overwriteFolder = folder;
    }

    const updateFiles = rootFolder.getFilesByType(MimeType.PDF);
    const overwriteFiles = overwriteFolder.getFilesByType(MimeType.PDF);

    var timeUp = false;
    while (updateFiles.hasNext() || overwriteFiles.hasNext()) {

      num_files++;
      var overwrite = false;
      var file = null;

      if (updateFiles.hasNext()) {
        file = updateFiles.next()
        overwrite = false;
      } else if (overwriteFiles.hasNext()) {
        file = overwriteFiles.next();
        overwrite = true;
      }

      // 実行時間が5分を超えたらAIに投げるのは中断して残りファイル数だけ数える（GASの実行時間制限のため）
      if(timeUp) continue;
      if ((new Date()) - startTime > 5 * 60 * 1000){
        timeUp = true;
        updateStatus(`Exceeding maximum execution time. Run the script again if you have more papers to update.`)
        continue;
      }

      try {
        num_processed_files++;
        processFile(file, pageList, overwrite);
      } catch (e) {
        updateStatus(e.message)
        file.setDescription(e);
        file.moveTo(errorFolder);
        updateStatus(`<span style="color:red;">Failed to update ${file.getId()}: ${file.getName()}, ${e.message}</span>`);
      }

      CacheService.getScriptCache().remove('running');
    }
  } catch (e) {
    updateStatus(`<span style="color:red;">${e.message}</span>`);
  }

  updateStatus(`Finished running script. Total time = ${(new Date() - startTime) / 1000} sec. ${num_files} files found, ${num_processed_files} files processed, ${num_files - num_processed_files} files waiting for update.`);
  return;

}

// HTMLページを表示するための関数
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

// ステータスの更新用
function getStatus() {
  var s = CacheService.getScriptCache().get('status');
  console.log(s)
  return s == null ? "" : s;
}

// 対象DB内のページ（論文）リストを，論文タイトルをキー，NotionのページIDとGoogle Driveの論文ファイルのIDを値とする連想配列で返す
// Google Driveの論文ファイルのIDは使ってない．上書き時に古いファイルを消すようにしようかと思ったが別にしなくていいかなと思い放置
function getNotionPageList() {

  updateStatus("Obtaining paper list from Notion database");
  var dict = {}

  var has_more = true;
  var next_cursor = undefined;
  while (has_more) {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': `Bearer ${NOTION_API_KEY}`,
        'Notion-Version': '2022-06-28'
      },
      payload: JSON.stringify({
        'start_cursor': next_cursor
      }),
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch('https://api.notion.com/v1/databases/' + NOTION_DB_ID + '/query', options);
    const result = JSON.parse(response.getContentText());
    const pages = result['results'];

    for (var i = 0; i < pages.length; i++) {
      var title = pages[i]['properties']['Title']['title'][0]['plain_text'];

      var fileURL = pages[i]['properties']['PDF Link']['url'];
      var fileID = fileURL.slice(32, -18)
      dict[title] = JSON.parse(`{"page_id": "${pages[i]['id']}", "file_id": "${fileID}"}`);
    }

    has_more = result['has_more']
    if (has_more) {
      next_cursor = result['next_cursor']
      Utilities.sleep(500);
    }
  }

  updateStatus(`${Object.keys(dict).length} entries found in the database`);
  return dict;

}

// ステータスメッセージをアップデート
function updateStatus(message) {
  status += (new Date().toLocaleTimeString('ja-JP')) + ": " + message + "<br>";
  CacheService.getScriptCache().put('status', status);
  console.log(message);
}

// 論文ファイル1つ1つに対する処理
function processFile(file, pageList, overwrite = false) {
  updateStatus("Processing file id " + file.getId() + ": " + file.getName())

  // Gemini APIにファイルをアップロード
  var fileBlob = file.getBlob();
  const uploadedFile = uploadFileToGemini(fileBlob);
  updateStatus(`Uploading file to Gemini API: ${uploadedFile.name}`);

  // ファイルの処理が完了するまで待機 (ポーリング)
  const activeFile = pollFileState(uploadedFile.name);
  updateStatus(`PDF file ready : ${activeFile.uri}`);

  const payload = {
    contents: [
      {
        parts: [
          { text: prompt },
          {
            file_data: {
              mime_type: activeFile.mime_type,
              file_uri: activeFile.uri
            }
          }
        ],
      },
    ],
  };

  updateStatus("Waiting for Gemini response")
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response);
  const content = data['candidates'][0]['content']['parts'][0]['text'];
  updateStatus("Response obtained from Gemini")

  // Gemini APIの利用制限（1分あたり10リクエスト）に引っかからないようにするため数秒待機
  // Geminiの応答だけで6秒以上かかるので別に要らないかも
  updateStatus("Waiting 6 sec (Gemini API quota restriction)");
  Utilities.sleep(6000);

  updateStatus("Converting Gemini response to json")
  var text = content.substring(content.indexOf("{"), content.lastIndexOf("}") + 1);
  var json = JSON.parse(text)

  // JSONから値を取得（たまにGeminiのレスポンス次第で取れてない値あり）
  var _title = json.title ? json["title"] : "";
  var _authors = json.authors ? json["authors"] : "";
  var _affiliation = json.affiliation ? json["affiliation"] : "";
  var _proc = json.proc ? json["proc"] : "";
  var _year = json.year ? json["year"] : "";
  var _summary = json.summary ? json["summary"] : "";
  var _doi = json.doi ? json["doi"] : "";
  var val = [_title, file.getUrl(), _authors, _affiliation, _proc, _year, _summary, _doi]

  if (_title != "") {
    if (_title in pageList && !overwrite) {
      throw new Error(`Error: The paper seems to be already in the papers list. To overwrite, move the PDF to the "overwrite" folder and try again.`);
    }
    else {
      file.setName(_title + ".pdf");

      updateStatus("Sending paper info to Notion database");
      if (overwrite) updateStatus("Warning: Overwriting existing data");
      createNotionEntry(file.getId(), val, pageList, overwrite);
      file.moveTo(successFolder);
      updateStatus(`<span style="color:blue;">Successfully updated "${_title}"</span>`)
    }
  }
  else {
    throw new Error(`Error: The title cannot be obtained. Check the PDF file and try again.`);
  }
}

// Gemini APIに論文ファイルをアップする
function uploadFileToGemini(fileBlob) {
  const url = `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${GEMINI_API_KEY}`;

  const options = {
    method: 'post',
    payload: fileBlob,
    headers: {
    },
    contentType: fileBlob.getContentType(),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode !== 200) {
    throw new Error(`ファイルアップロード失敗 (HTTP ${responseCode}): ${responseBody}`);
  }

  const result = JSON.parse(responseBody);
  return result.file;
}

// Gemini APIへのファイルアップロード状況を5秒おきに取得し，ファイルが準備できるまで待つ
function pollFileState(fileName) {
  const url = `https://generativelanguage.googleapis.com/v1beta/${fileName}?key=${GEMINI_API_KEY}`;
  const options = {
    method: 'get',
    muteHttpExceptions: true
  };

  while (true) {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode !== 200) {
      throw new Error(`ファイル状態の取得に失敗 (HTTP ${responseCode}): ${responseBody}`);
    }

    const file = JSON.parse(responseBody);

    if (file.state === 'ACTIVE') {
      return file;
    } else if (file.state === 'FAILED') {
      const errorMessage = file.error ? file.error.message : '理由不明';
      throw new Error(`ファイルの処理に失敗しました。理由: ${errorMessage}`);
    }

    // 5秒待機
    Utilities.sleep(5000);
  }
}

// 論文の1ページ目の画像を取得する
function savePdfThumbnailAsPng(pdfId) {
  // 1. Drive APIを使ってファイルのメタデータを取得する
  //    fieldsパラメータで、必ず'thumbnailLink'を指定します。
  const fileMetadata = Drive.Files.get(pdfId, { fields: "id, name, thumbnailLink" });

  // 2. サムネイルリンクの存在を確認
  if (!fileMetadata.thumbnailLink) {
    console.error("このファイルにはサムネイルがありません。別の方法を試してください。");
    return;
  }

  // サムネイルはサイズが小さい場合があります。URLの末尾の '=s220' を変更すると
  // より大きな画像を取得できることがあります。（例: '=s800'）
  const imageUrl = fileMetadata.thumbnailLink.replace('=s220', '=c');

  // 3. サムネイルのURLから画像データを取得する
  //    API経由で取得したURLにアクセスするには、認証トークンが必要です。
  const response = UrlFetchApp.fetch(imageUrl, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });

  const imageBlob = response.getBlob();
  return imageBlob;
}

// 情報をNotionに送信する
function createNotionEntry(pdfId, val, pageList, overwrite) {

  var fileBlob = savePdfThumbnailAsPng(pdfId);

  const getUploadUrlOptions = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${NOTION_API_KEY}`,
      'Notion-Version': '2022-06-28'
    },
    payload: JSON.stringify({
    }),
    muteHttpExceptions: true // エラー時に例外をスローせず、レスポンスを返す
  };
  const getUploadUrlResponse = UrlFetchApp.fetch('https://api.notion.com/v1/file_uploads', getUploadUrlOptions);
  const getUploadUrlResult = JSON.parse(getUploadUrlResponse.getContentText());

  if (getUploadUrlResponse.getResponseCode() !== 200) {
    throw new Error(`URL取得エラー: ${getUploadUrlResult.message}`);
  }

  const uploadUrl = getUploadUrlResult["upload_url"];
  const fileUploadID = getUploadUrlResult["id"];

  // 1. 境界文字列を定義
  const boundary = '----GasBoundary' + new Date().getTime();

  // 2. ペイロードを組み立て
  let data = '';

  // ファイルフィールド部分
  data += '--' + boundary + '\r\n';
  data += 'Content-Disposition: form-data; name="' + 'file' + '"; filename="' + fileBlob.getName() + '"\r\n';
  data += 'Content-Type: ' + fileBlob.getContentType() + '\r\n\r\n';
  data += fileBlob.getDataAsString()
  data += '\r\n--' + boundary + '--\r\n';

  const uploadFileOptions = {
    method: "POST", headers: {
      'Authorization': `Bearer ${NOTION_API_KEY}`,
      'Notion-Version': '2022-06-28'
    }, payload: { meta: JSON.stringify({ id: 5, subject: "foo", date: "2018-10-05" }), file: fileBlob }
  }

  const uploadFileResponse = UrlFetchApp.fetch(uploadUrl, uploadFileOptions);

  if (uploadFileResponse.getResponseCode() !== 200) {
    throw new Error(`ファイルアップロードエラー: ${uploadFileResponse.getContentText()}`);
  }

  const method = overwrite ? 'patch' : 'post';
  const updatePageOptions = {
    method: method,
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${NOTION_API_KEY}`,
      'Notion-Version': '2022-06-28'
    },
    payload: JSON.stringify({
      "parent": {
        "type": "database_id",
        "database_id": NOTION_DB_ID
      },
      "cover": {
        "type": "file_upload",
        "file_upload": {
          "id": fileUploadID
        }
      },
      "properties": {
        "Title": {
          "title": [
            {
              "text": {
                "content": val[0]
              }
            }
          ]
        },
        "PDF Link": {
          "url": val[1]
        },
        "Authors": {
          "rich_text": [
            {
              "text": {
                "content": val[2]
              }
            }
          ]
        },
        "Author Affiliation": {
          "rich_text": [
            {
              "text": {
                "content": val[3]
              }
            }
          ]
        },
        "Journal/Proceeding": {
          "rich_text": [
            {
              "text": {
                "content": val[4]
              }
            }
          ]
        },
        "Year": {
          "number": parseInt(val[5])
        },
        "Summary": {
          "rich_text": [
            {
              "text": {
                "content": val[6]
              }
            }
          ]
        },
        "DOI": {
          "rich_text": [
            {
              "text": {
                "content": val[7]
              }
            }
          ]
        }

      }
    }),
    muteHttpExceptions: true
  };

  const notionPageId = overwrite ? ("/" + pageList[val[0]]['page_id']) : "";
  const updatePageResponse = UrlFetchApp.fetch(`https://api.notion.com/v1/pages${notionPageId}`, updatePageOptions);
  const updatePageResult = JSON.parse(updatePageResponse.getContentText());

  if (updatePageResponse.getResponseCode() !== 200) {
    throw new Error(`ページ更新エラー: ${updatePageResult.message}`);
  }
}


