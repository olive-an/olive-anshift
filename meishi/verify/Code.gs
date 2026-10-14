// 検証用スクリプト。名刺データは扱わない。
function doGet() {
  return HtmlService.createHtmlOutputFromFile('verify')
    .setTitle('Olive 名刺 検証')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 検証A: ログイン中の利用者のメールアドレスが取れるか
function whoAmI() {
  return {
    active: Session.getActiveUser().getEmail(),
    effective: Session.getEffectiveUser().getEmail()
  };
}

// 検証D: ドライブの使用量が取れるか
function driveInfo() {
  return { used: DriveApp.getStorageUsed(), limit: DriveApp.getStorageLimit() };
}

// 検証E: 画像を非公開でドライブに保存し、読み戻せるか
function saveAndReadBack(base64) {
  var it = DriveApp.getFoldersByName('名刺検証');
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder('名刺検証');
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg', 'test.jpg');
  var file = folder.createFile(blob);
  var back = Utilities.base64Encode(file.getBlob().getBytes());
  return {
    id: file.getId(),
    sharing: String(file.getSharingAccess()),
    bytes: blob.getBytes().length,
    readBack: 'data:image/jpeg;base64,' + back
  };
}
