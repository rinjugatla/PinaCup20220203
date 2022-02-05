// 1行目はヘッダ
const OffsetHeader = 2;
// 予選シート
const EntoryId = 0;
const QualiId = 308016459;
const MainId = 1528242347;

function onOpen() {
  // スプレッドシートを開いたときに呼びだされる
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('大会用コマンド', [
      { name: '予選組作成', functionName: 'updateQualifyingList' },
    ]);
}

function updateQualifyingList(){
  new QualifyingTournament();
}

class LockController {
  constructor() {
    this.quali = Utility.GetRows(QualiId);
    this.main = Utility.GetRows(MainId);
  }

  IsQualiLocked() {
    // 予選大会がロックされているか
    Logger.log(this.quali);
    return this.quali.length > 1;
  }

  IsMainLocked() {
    // 本線大会がロックされているか
    return this.main.length > 1;
  }
}

class Utility {
  // ヘッダ取得, list
  static GetHeaders(sheetId) {
    let rows = this.GetRows(sheetId);
    return rows[0];
  }

  // すべての行を取得
  static GetRows(sheetId){
    const sh = this.GetSheetById(sheetId);
    return sh.getDataRange().getValues();
  }

  // GIDからシートを取得
  static GetSheetById(gid) {
    for (const sheet of SpreadsheetApp.getActive().getSheets()) {
      if (sheet.getSheetId() == gid) {
        return sheet;
      }
    }
    return null;
  }

  // https://www.nxworld.net/js-array-shuffle.html
  // JavaScript：配列内の要素をシャッフル（ランダムソート）する方法
  static Shuffle([...array]) {
    for (let i = array.length - 1; i >= 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [array[i], array[j]] = [array[j], array[i]];
    }
    return array;
  }

  static DeletePrevData(sheetId) {
    // 前回実行時の結果を削除
    // ヘッダーのみ残す
    const sh = Utility.GetSheetById(sheetId);
    const rows = sh.getDataRange().getValues();
    // ヘッダーしかない場合は処理不要
    if (rows.length < OffsetHeader) { return; }

    const headers = rows[0];
    let emptyData = [];
    for (let rowi = 0; rowi < rows.length - 1; rowi++) {
      emptyData.push(new Array(headers.length));
    }
    sh.getRange(OffsetHeader, 1, rows.length - 1, headers.length).setValues(emptyData);
  }
}

class QualifyingTournament {
  constructor() {
    // 予選大会リストの初期化 -> 作成
    let lock = new LockController();
    if(lock.IsQualiLocked()){
      Browser.msgBox('予選組が作成済みです。\n再作成する場合は予選組データを削除してください。');
      return;
    }
    Utility.DeletePrevData(QualiId);
    this.CreateQualifyingList();
  }

  CreateQualifyingList() {
    // 予選大会参加者のランダムリストを作成
    // 実行するたびにシャッフルされるので注意
    let entory = new EntoryController(EntoryId, QualiId);
    let shuffledUsers = entory.ShuffledUsers;
    // 予選組シートに結果を反映
    const sh = Utility.GetSheetById(QualiId);
    sh.getRange(OffsetHeader, 1, shuffledUsers.length, shuffledUsers[0].length).setValues(shuffledUsers);
  }
}

class MainTournament {

}

class EntoryController {
  // entorySheetId: 
  // targetSheetId:
  constructor(entorySheetId, targetSheetId) {
    this.EntorySheetId = entorySheetId;
    this.TargetSheetId = targetSheetId;
    this.Users = this.GetUsers();
    this.AttendUsers = this.ExtructAttendUserIds();
    this.ShuffledUsers = this.CreateShuffledUserList();
  }

  // 参加者辞書を取得
  GetUsers() {
    const sh = Utility.GetSheetById(this.EntorySheetId);
    const rows = sh.getDataRange().getValues();

    let headers = rows[0];
    // key: nomber
    let users = {};
    for (let rowi = 1; rowi < rows.length; rowi++) {
      const row = rows[rowi];
      let id = 0;
      let rowData = {};
      for (let col = 0; col < row.length; col++) {
        let colName = headers[col];
        if (colName == 'No') { id = Number(row[col]); }
        rowData[colName] = row[col];
      }
      users[id] = rowData;
    }

    return users;
  }

  // 応募者dictから参加者idを抽出
  ExtructAttendUserIds() {
    let ids = [];
    for (const id in this.Users) {
      if (this.Users[id]['参加'] == '〇') { ids.push(id); }
    }
    return ids;
  }

  // 参加者をシャッフル
  // targetHeaders: 出力先のヘッダー
  CreateShuffledUserList() {
    let targetHeaders = Utility.GetHeaders(QualiId);
    let shuffledIds = Utility.Shuffle(this.AttendUsers);
    let shuffledUsers = [];
    for (let i = 0; i < shuffledIds.length; i++) {
      let values = this.CreateValueList(targetHeaders, this.Users[shuffledIds[i]]);
      shuffledUsers.push(values);
    }

    // 組情報を追加
    let addedGroupKey = this.AddGroupKey(targetHeaders, shuffledUsers);
    return addedGroupKey;
  }

  // ヘッダー順にdictからlistを作成
  CreateValueList(headers, valueDict) {
    let result = [];
    for (const header of headers) {
      if (header in valueDict) {
        result.push(valueDict[header]);
      } else {
        // 存在しない値はnullを代入
        result.push(null);
      }
    }
    return result;
  }

  // 参加者listに組情報を追加
  AddGroupKey(headers, users) {
    const aCode = 65;
    let groupIndex = headers.indexOf('組');
    let result = [];
    for (let i = 0; i < users.length; i++) {
      let values = users[i].slice();
      // 3名ずつ同じ組に入れる A A A B B B C...
      values[groupIndex] = String.fromCharCode(65 + Math.trunc(i / 3));
      result.push(values);
    }
    return result;
  }
}

class Test {
  shuffle_test() {
    const array = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    const result = shuffle(array);
    Logger.log(result);
  }
}