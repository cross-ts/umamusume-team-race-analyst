// Sheet Names
const TEAM_RACE_SUMMARY = 'チームレース戦歴';
const SHEET_FOR_ADD = '各種追加用シート';
const HALL_OF_FAME_UMAMUSUME = '殿堂入りウマ娘';
const MASTER_DATA = 'マスターデータ';

// Drive FolderNames
const TEAM_RACE_FOLDER = 'ウマ娘チームレース'

// 関数共通のシート取得
const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const ADD_SHEET = SPREADSHEET.getSheetByName(SHEET_FOR_ADD);


// Validateエラー用のアラート表示
function alert(title, message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(title, message, ui.ButtonSet.OK);
}

DOCUMENT_RESOURCE = {
  'title': 'tmp',
  'mimeType': 'image/png'
}
UMAMUSUME_NAMES_RANGE = 'B3:B17'
POINT_RANGE = 'C3:C17'
// ウマ娘チームレースディレクトリに上げたスクショからスコアの自動入力を行う
function inputDataFromScreenShot() {
  // 画像をGoogle Documentに変換しocrを使いスコアをtextとして取得する
  var folder = DriveApp.getFoldersByName(TEAM_RACE_FOLDER);
  if (!folder.hasNext()) {
    alert('フォルダが存在しません', `Google Driveに「${TEAM_RACE_FOLDER}」という名前のフォルダを作成してください。`);
  }
  var files = folder.next().getFiles();
  var raw_texts = [];
  while (files.hasNext()) {
    var file = files.next();
    var blobData = file.getBlob();
    var fileID = Drive.Files.insert(DOCUMENT_RESOURCE, blobData, {ocr: true}).id;
    var document = DocumentApp.openById(fileID);
    var text = document.getBody().getText();
    raw_texts = raw_texts.concat(text.split('\n'));
    Drive.Files.remove(fileID);
    Drive.Files.remove(file.getId());
  }

  // 取得したスコアとウマ娘名で突合を行いスコアを入力する
  var umamusumeNames = ADD_SHEET.getRange(UMAMUSUME_NAMES_RANGE).getValues();
  var updateValues = [];
  umamusumeNames.forEach(e => {
    var name = e[0];
    var index = 1;
    raw_texts.some(str => {
      if (~str.indexOf(name)) {
        return true;
      }
      index += 1;
    })
    var point = parseInt(raw_texts[index].replace(/,/g, ''));
    updateValues.push([point]);
  })
  ADD_SHEET.getRange(POINT_RANGE).setValues(updateValues);
}

const ADD_POINT_RANGE = 'A3:C17'
const HALL_OF_FAME_UMAMUSUME_UPDATE_RANGE = 'D%d:M%d'
const RESET_POINT_RANGE = 'C3:C17'
const RESET_INPUT = [
  [""],[""],[""],
  [""],[""],[""],
  [""],[""],[""],
  [""],[""],[""],
  [""],[""],[""],
]
function saveTeamRacePoints() {
  var inputData = ADD_SHEET.getRange(ADD_POINT_RANGE).getValues();
  var fameUmamusumeSheet = SPREADSHEET.getSheetByName(HALL_OF_FAME_UMAMUSUME);

  // NOTE: この辺りRangeListにして1回のみ更新にした方が早そう
  inputData.some(e => {
    var umamusumeID = e[0];
    var umamusumeName = e[1];
    var teamRacePoint = e[2];
    if (teamRacePoint == "") {
      alert(`${umamusumeName}のポイント欄が空です`, 'レースポイントを入力してください');
      return true;
    }
    // HACK: 殿堂入りウマ娘IDはROWS() - 1で採番していることを逆手に取る
    var targetUmamusumeRange = HALL_OF_FAME_UMAMUSUME_UPDATE_RANGE.replace(/%d/g, parseInt(umamusumeID + 1));
    var updateRange = fameUmamusumeSheet.getRange(targetUmamusumeRange);

    // 今回のスコアを先頭にして10戦以上前のスコアを落として更新する
    var updateValues = updateRange.getValues();
    updateValues[0].unshift(teamRacePoint);
    updateValues[0].pop();
    updateRange.setValues(updateValues);
  });

  // 入力を削除
  ADD_SHEET.getRange(RESET_POINT_RANGE).setValues(RESET_INPUT)
}

// 新規殿堂入りウマ娘の追加
const HALL_OF_FAME_UMAMUSUME_INPUT_RANGE = 'E3:F3';
const HALL_OF_FAME_UMAMUSUME_NAME_RANGE = 'B2:B';
const WRITABLE_RANGE = 'B%d:C%d';
function saveNewHallOfFameUmamusume() {
  // 入力の確認
  var inputRange = ADD_SHEET.getRange(HALL_OF_FAME_UMAMUSUME_INPUT_RANGE)
  var newUmamusume = inputRange.getValues();
  var umamusumeName = newUmamusume[0][0];
  var trainedDate = newUmamusume[0][1];
  if (umamusumeName == "") {
    alert("名前が空です", "新規登録するウマ娘名を選択してください");
    return;
  }
  if (trainedDate == "") {
    alert("育成日が空です", "新規登録するウマ娘の育成日を選択してください");
    return;
  }

  // 書き込み
  var fameUmamusumeSheet = SPREADSHEET.getSheetByName(HALL_OF_FAME_UMAMUSUME);
  var columnUmamusumeName = fameUmamusumeSheet.getRange(HALL_OF_FAME_UMAMUSUME_NAME_RANGE);
  var umamusumeNames = columnUmamusumeName.getValues();
  counter = 0;
  while ( umamusumeNames[counter] && umamusumeNames[counter][0] != "" ) {
    counter++;
  }
  var blankRange = fameUmamusumeSheet.getRange(WRITABLE_RANGE.replace(/%d/g, counter + 2));
  blankRange.setValues(newUmamusume);

  // 入力のリセット
  inputRange.setValues([["",""]]);
}
