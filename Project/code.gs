function onOpen() {
  buildMenu();
}

function buildMenu() {
  const ui = DocumentApp.getUi();
  ui.createMenu('ファイルを編集')
    .addItem('読み込み', 'loadTextFile')
    .addItem('保存', 'saveAsNewTextFile')
    .addItem('ファイルの詳細', 'showFileDetails');

  // HTMLファイルならhtmlで表示を追加
  const props = PropertiesService.getDocumentProperties().getProperties();
  if (props.mimeType === 'text/html') {
    ui.createMenu('ファイルを編集')
      .addItem('読み込み', 'loadTextFile')
      .addItem('保存', 'saveAsNewTextFile')
      .addItem('ファイルの詳細', 'showFileDetails')
      .addItem('HTMLで表示', 'showHtmlContent')
      .addToUi();
  } else {
    ui.createMenu('ファイルを編集')
      .addItem('読み込み', 'loadTextFile')
      .addItem('保存', 'saveAsNewTextFile')
      .addItem('ファイルの詳細', 'showFileDetails')
      .addToUi();
  }
}

function loadTextFile() {
  const ui = DocumentApp.getUi();
  const response = ui.prompt('読み込みたいファイルのIDまたはGoogleドライブの共有リンクを入力してください（HTMLやJSなどのテキストファイル）');

  if (response.getSelectedButton() === ui.Button.OK) {
    let input = response.getResponseText().trim();

    // URLからIDを抽出
    const fileIdFromUrl = extractFileIdFromUrl(input);
    const fileId = fileIdFromUrl || input;

    if (!fileId.match(/^[a-zA-Z0-9_-]{10,}$/)) {
      ui.alert('有効なファイルIDまたは共有リンクを入力してください。');
      return;
    }

    try {
      const file = DriveApp.getFileById(fileId);
      const content = file.getBlob().getDataAsString();

      // ドキュメントに書き込む
      const doc = DocumentApp.getActiveDocument();
      const body = doc.getBody();
      body.clear();
      body.setText(content);

      // ファイル情報を保存
      PropertiesService.getDocumentProperties().setProperties({
        fileId: fileId,
        fileName: file.getName(),
        mimeType: file.getMimeType(),
        parentFolderId: file.getParents().hasNext()
          ? file.getParents().next().getId()
          : ''
      });

      ui.alert('ファイルの内容を読み込みました。');

      // メニュー更新（htmlで表示 の有無切替のため）
      buildMenu();

    } catch (e) {
      ui.alert('ファイルの読み込みに失敗しました: ' + e.message);
    }
  }
}

function saveAsNewTextFile() {
  const props = PropertiesService.getDocumentProperties().getProperties();
  const fileId = props.fileId;
  const originalName = props.fileName;
  const mimeType = props.mimeType || 'text/plain';
  const folderId = props.parentFolderId;

  if (!fileId || !originalName) {
    DocumentApp.getUi().alert('先に「読み込み」を行ってください。');
    return;
  }

  try {
    const extension = getExtensionFromMimeType(mimeType) || getExtensionFromFilename(originalName);
    const baseName = originalName.replace(/\.[^/.]+$/, '');

    // タイムスタンプの生成（;と^を使う形式）
    const now = new Date();
    const pad = n => n.toString().padStart(2, '0');
    const timestamp = `${now.getFullYear()};${pad(now.getMonth() + 1)};${pad(now.getDate())}^${pad(now.getHours())};${pad(now.getMinutes())};${pad(now.getSeconds())}`;
    const newName = `${baseName}-edit;${timestamp}.${extension}`;

    const content = DocumentApp.getActiveDocument().getBody().getText();
    const blob = Utilities.newBlob(content, mimeType, newName);
    const newFile = DriveApp.createFile(blob);

    // 元フォルダに移動
    if (folderId) {
      const folder = DriveApp.getFolderById(folderId);
      folder.addFile(newFile);
      DriveApp.getRootFolder().removeFile(newFile);
    }

    DocumentApp.getUi().alert(`ファイルを保存しました: ${newName}`);

  } catch (e) {
    DocumentApp.getUi().alert('保存に失敗しました: ' + e.message);
  }
}

function showFileDetails() {
  const props = PropertiesService.getDocumentProperties().getProperties();
  const ui = DocumentApp.getUi();

  if (!props.fileId || !props.fileName) {
    ui.alert('まだファイルを読み込んでいません。');
    return;
  }

  const folderId = props.parentFolderId;
  let folderUrl = 'なし';
  if (folderId) {
    folderUrl = `https://drive.google.com/drive/folders/${folderId}`;
  }

  const message = 
    `ファイル名: ${props.fileName}\n` +
    `ファイルID: ${props.fileId}\n` +
    `保存先フォルダ: ${folderUrl}`;

  ui.alert('現在編集中のファイル詳細', message, ui.ButtonSet.OK);
}

// htmlで表示用
function showHtmlContent() {
  const props = PropertiesService.getDocumentProperties().getProperties();
  if (!props.fileId || !props.fileName) {
    DocumentApp.getUi().alert('まだファイルを読み込んでいません。');
    return;
  }
  const htmlContent = DocumentApp.getActiveDocument().getBody().getText();

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(2000)
    .setHeight(1500)
    .setTitle('HTMLファイルの内容');

  DocumentApp.getUi().showModalDialog(htmlOutput, 'HTMLで表示');
}

// URLからGoogleドライブファイルIDを抽出するヘルパー関数
function extractFileIdFromUrl(url) {
  const regex = /\/d\/([a-zA-Z0-9_-]{10,})/;
  const match = url.match(regex);
  return match ? match[1] : null;
}

// MIMEタイプから拡張子を取得
function getExtensionFromMimeType(mimeType) {
  const map = {
    'text/html': 'html',
    'application/javascript': 'js',
    'text/css': 'css',
    'application/json': 'json',
    'text/x-python': 'py',
    'text/plain': 'txt'
  };
  return map[mimeType] || null;
}

// ファイル名から拡張子を取得
function getExtensionFromFilename(filename) {
  const match = filename.match(/\.([a-z0-9]+)$/i);
  return match ? match[1] : 'txt';
}
