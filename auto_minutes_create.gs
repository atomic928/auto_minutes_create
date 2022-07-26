/* 一回デバッグで実行して権限許可しないとフォームから実行できないっぽい */
/* フォームに投稿があったときにこの関数を実行する */
function onFormSubmit(e){
  console.log(e.namedValues);
  var date = e.namedValues["日付を入力してください"][0];
  var startHour = e.namedValues["何時から始めますか"][0];
  var startMinute = e.namedValues["何分から始めますか"][0];
  createMinutesTemplate(date, startHour, startMinute);
}

/* 議事録を作成する関数 */
function createMinutesTemplate(date, startHour, startMinute) {
  /* MTGの日時の設定 */
  var year = date.slice(0,4);
  var month = date.slice(5,7);
  var day = date.slice(-2);
  var dayOfWeek = judgeDayOfWeek(year, month, day);
  var startHour = startHour;
  /* startMinuteが設定されていない場合00とする */
  if (startMinute == '') {
    startMinute = '00';
  }
  var startMinute = startMinute;

  /* フォルダに入ってるファイルの数 */
  var numberOfFiles = countFolders()+1;

  /* 各種担当 */
  var facilitator = selectFacilitator(numberOfFiles);
  var minutesCharge = selectMinutesCharge(numberOfFiles);

  /* 議事録のタイトル */
  var title = year + '_' + month + '_' + day + '_AssembRe_第' + numberOfFiles + '回議事録';

  /* GoogleDocumentを任意の名前で作成する */
  var folder = DriveApp.getFolderById('19dfHi8qZ2Gh6iXa9NSQJK01BCPFbLaoQ')
  var minutes_template = DocumentApp.create(title);
  var fileId = minutes_template.getId();
  var body = minutes_template.getBody();
  /* 本体にヘッダーを挿入する */
  body.appendParagraph('----------------------------------------------------------------------------------------------------').setHeading(DocumentApp.ParagraphHeading.NORMAL);
  /* 本体に今日の日付を挿入する */
  body.appendParagraph('日時：' + year + '年 ' + month + '月 ' + day + '日 ' + dayOfWeek + '曜日 ' + startHour + '：' + startMinute + ' ~ ').setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.NORMAL);
  /* 本体に会議の参加者入力欄を挿入する */
  body.appendParagraph('出席者：').setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.NORMAL);
  body.appendParagraph('進行：' + facilitator).setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.NORMAL);
  body.appendParagraph('議事録：' + minutesCharge).setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.NORMAL);
  body.appendParagraph('----------------------------------------------------------------------------------------------------').setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.NORMAL);
  /* 本体に見出しをつける */
  body.appendParagraph('アジェンダ').setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.NORMAL);
  body.appendParagraph('').setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.NORMAL);
  body.appendParagraph('議題').setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.NORMAL);
  body.appendParagraph('').setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.NORMAL);
  body.appendParagraph('宿題を置く場所').setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.NORMAL);
  body.appendParagraph('').setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.NORMAL);
  body.appendParagraph('次回やること').setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.NORMAL);
  body.appendParagraph('').setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.NORMAL);
  minutes_template.saveAndClose();

  var file = DriveApp.getFileById(fileId);
  /* 共有権限の設定 */
  setAccessPermission(file);
  folder.addFile(file);

  /* slackへ投稿 */
  var url = file.getUrl();
  var message = year + '年' + month + '月' + day + '日' + dayOfWeek + '曜日 ' + startHour + '：' + startMinute + '~ 行われるMTGの議事録です。'
  send(url, message);
}

/* フォルダに入っているファイルの数を数える関数 */
function countFolders() {
  var folder = DriveApp.getFolderById("19dfHi8qZ2Gh6iXa9NSQJK01BCPFbLaoQ");
  var contents = folder.getFiles();
  var i = 0;
  while (contents.hasNext()) {
    file = contents.next();
    i++
  }
  Logger.log("このフォルダ内のファイル数は" + i + "です。");
  return i;
}

/* 編集権限の設定をする関数 */
function setAccessPermission(document) {
  var access;
  var permission;
  var status = 1;

  //①リンク共有有効(編集権限あり)
  if (status == 1) {
    access = DriveApp.Access.ANYONE_WITH_LINK;
    permission = DriveApp.Permission.EDIT;
    document.setSharing(access, permission);
  }

  //②リンク共有有効(閲覧権限のみ)
  if (status == 2) {
    access = DriveApp.Access.ANYONE_WITH_LINK;
    permission = DriveApp.Permission.VIEW;
    document.setSharing(access, permission);
  }

  //③リンク共有無効
  if (status == 3) {
    access = DriveApp.Access.PRIVATE;
    permission = DriveApp.Permission.EDIT;
    document.setSharing(access, permission);
  }
}

/* 進行役を決定する関数 */
function selectFacilitator(numberOfFiles) {
  var facilitatorList = ["出口", "関", "仲山"]
  return facilitatorList[(numberOfFiles+2)%3];
}

/* 議事録担当を決定する関数 */
function selectMinutesCharge(numberOfFiles) {
  var minutesChargeList = ["野村", "小池", "柴田", "成瀨", "上條", "児玉", "古川", "福田"]
  return minutesChargeList[(numberOfFiles+2)%8];
}

/* 日時から曜日を取得する関数 */
function judgeDayOfWeek(year, month, day) {
  var dateList = ["日", "月", "火", "水", "木", "金", "土"];
  var date = new Date(year, month-1, day)
  var dayNum = date.getDay();
  return dateList[dayNum];
}

/* 議事録が出来たときに投稿するslackのチャンネル */
const SLACK_CHANNEL_URL = PropertiesService
    .getScriptProperties()
    .getProperty('SLACK_CHANNEL_URL')

/* slackに投稿するための関数 */
function send(url, message) {
  // 投稿するメッセージ
  const text = message + '\n' + url;

  // 投稿者名とアイコンを設定する
  const data = {
     "username" : 'Slack Panda',
     "icon_emoji": ':panda_face:',
     text,
  };

  const params = {
    "method" : "POST",
    "contentType" : "application/json",
    "payload" : JSON.stringify(data),
  };

  UrlFetchApp.fetch(SLACK_CHANNEL_URL, params)
}
