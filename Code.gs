/**
 * Webアプリとして画面を表示する
 *
 * どんな関数か: GASのWebアプリが決めている特別な入り口の関数（名前固定、GASが自動で呼ぶ）
 * 使ってるメソッド: HtmlService.createTemplateFromFile() / .evaluate() / .setTitle()
 * できあがるもの: ブラウザに表示できる完成済みHTML（HtmlOutputというオブジェクト）
 */
// ユーザーがURLを開いた瞬間最初に呼ばれる関数
// return はブラウザに渡される
function doGet() {
  // index.htmlをテンプレートとして読み、setTitleでタイトルを設定
  //addMetaTag(名前, 内容)
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("シフト提出")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * membersシートから
 * 校舎一覧とメンバー一覧を取得する
 *
 * A列：name
 * B列：school_name
 *
 * どんな関数か: スプレッドシートを読み取って、フロント用に整形して返すデータ取得系の関数
 * 使ってるメソッド: SpreadsheetApp.getActiveSpreadsheet() / getSheetByName() / getRange().getValues() / map・filter・Set
 * できあがるもの: { schools: ["新宿校", "渋谷校"], members: [{name, school_name}, ...] } という形のオブジェクト
 */
function getMembers() {
  // このスクリプトが紐づいているスプレッドシートを取ってくる
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // getSheetByName()でmembersとあるシートをとってくる
  const sheet = spreadsheet.getSheetByName("members");

  // sheetでmembersシートが見つからなければエラーを吐く
  if (!sheet) {
    throw new Error("membersシートが見つかりません。");
  }
  // そのシートでデータが入ってる一番下の行を取得
  const lastRow = sheet.getLastRow();
  // lastRowが1だったらヘッダーしかなくデータがないということなので、空の配列を返して終わる
  if (lastRow < 2) {
    return {
      schools: [],
      members: [],
    };
  }
  // getRange(開始行, 開始列, 行数, 列数)
  // スプシmembersの1列2行目から最終列までを取得。
  // 列数2なのでa列とb列を取得
  // values は[["田中", "渋谷校"], ["鈴木", "新宿校"]]みたいな配列
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const members = values
    // 配列をmapで[{name: "田中", school_name: "渋谷校"}, ...]というオブジェクトの形へ
    .map(function (row) {
      return {
        // string型にしnameには.trimで空白除いてnameを。
        name: String(row[0] || "").trim(),
        school_name: String(row[1] || "").trim(),
      };
    })
    // 名前か校舎が空白であればその行は消す
    .filter(function (member) {
      return member.name !== "" && member.school_name !== "";
    });

  // membersからshool_nameだけの配列を作成。
  const schools = [
    ...new Set(
      members.map(function (member) {
        return member.school_name;
      }),
    ),
  ].sort();
  // 最後に学校名とメンバーをオブジェクト形式で返す
  return {
    schools: schools,
    members: members,
  };
}

/**
 * Settingsシートから時間の選択肢一覧を取得する
 *
 * A列：time
 *
 * どんな関数か: スプレッドシートを読み取って、フロント用に整形して返すデータ取得系の関数
 * 使ってるメソッド: SpreadsheetApp.getActiveSpreadsheet() / getSheetByName() / getRange().getValues() / map・filter
 * できあがるもの: { times: ["09:00", "09:30", ...] } という形のオブジェクト
 * 呼ばれるとこ: index.html側のloadTimeOptionsから
 * 影響先: 画面は触らない。戻り値をフロント側がプルダウンに反映する
 */
function getTimeOptions() {
  const settingSheet = SpreadsheetApp.getActiveSpreadsheet();
  const settings = settingSheet.getSheetByName("settings");

  if (!settings) {
    throw new Error("settingsシートが見つかりません");
  }
  const settingLastRow = settings.getLastRow();
  if (settingLastRow < 2) {
    return {
      times: [],
    };
  }
  const timeValues = settings.getRange(2, 1, settingLastRow - 1, 1).getValues();
  const times = timeValues
    .map(function (row) {
      const value = row[0];
      if (value instanceof Date) {
        return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
      }
      return String(value || "").trim();
    })
    .filter(function (time) {
      return time !== "";
    });

  return {
    times: times,
  };
}
/**
 * シフトと授業時間を保存する
 * 呼ばれるとこ: index.html側のsendShiftsから
 * 影響先: shifts_DB, lessons_DBシートに行を追加
 *
 * どんな関数か: フロントから届いたデータをチェックしてから、スプレッドシートに書き込む保存系の関数
 * 使ってるメソッド: Utilities.getUuid() / getRange().setValues() / forEach
 * できあがるもの: shifts_DBに選んだ日数分の行、lessons_DBに(日数×授業数)分の行が追加される。
 *              戻り値は { success, shift_count, lesson_count, ... } という登録件数の報告オブジェクト
 */
function saveShifts(request) {
  // リクエスト自体が無ければ即エラー
  if (!request) {
    throw new Error("送信データがありません。");
  }

  // フロントから届いたデータを取り出す。空でも安全なようにString+trim
  const schoolName = String(request.school_name || "").trim();
  const userName = String(request.user_name || "").trim();
  const staffStart = String(request.staff_start || "").trim();
  const staffEnd = String(request.staff_end || "").trim();
  // 配列はStringに変換すると壊れるので、そのまま受け取る（中身のチェックはvalidateRequest_で）
  const selectedDates = request.selected_dates;
  const lessons = request.lessons;

  // ここで中身をチェック。ダメならthrowされてこの先は実行されない
  validateRequest_(schoolName, userName, staffStart, staffEnd, selectedDates, lessons);

  // 書き込み先のシートを取ってくる
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const shiftsSheet = spreadsheet.getSheetByName("shifts_DB");
  const lessonsSheet = spreadsheet.getSheetByName("lessons_DB");

  if (!shiftsSheet) {
    throw new Error("shifts_DBシートが見つかりません。");
  }

  if (!lessonsSheet) {
    throw new Error("lessons_DBシートが見つかりません。");
  }

  // Utilities.getUuid()はGAS標準の「重複しないランダムなID文字列」を作る機能
  // batchIdは今回の送信1回分（複数日まとめて）をひとまとめに識別するID
  const batchId = Utilities.getUuid();
  const timestamp = new Date();
  // シートに書き込む行を溜めておく配列
  const shiftRows = [];
  const lessonRows = [];

  // 選んだ日付ごとに、シフト1行分を作る
  selectedDates.forEach(function (date) {
    // shiftIdはその1日分だけを識別するID（batchIdは全体、shiftIdは1日ごと、という粒度の違い）
    const shiftId = Utilities.getUuid();

    // shifts_DBシートの列順（id, batchId, 校舎, 名前, 日付, 開始, 終了, ステータス, 登録日時）に合わせて並べる
    shiftRows.push([
      shiftId,
      batchId,
      schoolName,
      userName,
      date,
      staffStart,
      staffEnd,
      "PENDING", // 承認前の初期ステータス。あとで別処理が承認/却下に変える想定
      timestamp,
    ]);

    // その日にある授業時間ごとに、授業1行分を作る（shiftIdで上のシフト行と紐づく＝どの日の授業か分かるように）
    lessons.forEach(function (lesson, index) {
      // lessons_DBシートの列順（id, shiftId, 何コマ目か, 開始, 終了, 登録日時）に合わせて並べる
      lessonRows.push([
        Utilities.getUuid(),
        shiftId,
        index + 1,
        lesson.start,
        lesson.end,
        timestamp,
      ]);
    });
  });

  // 溜めた行を、シートの一番下（getLastRow()+1行目）からまとめて書き込む
  // shiftRows[0].lengthは1行分の列数（配列の要素数）
  shiftsSheet
    .getRange(shiftsSheet.getLastRow() + 1, 1, shiftRows.length, shiftRows[0].length)
    .setValues(shiftRows);

  lessonsSheet
    .getRange(lessonsSheet.getLastRow() + 1, 1, lessonRows.length, lessonRows[0].length)
    .setValues(lessonRows);

  // 何件登録できたかをブラウザ側へ返す。sendShiftsの成功メッセージで使われる
  return {
    success: true,
    shift_count: shiftRows.length,
    lesson_count: lessonRows.length,
    selected_date_count: selectedDates.length,
    lesson_type_count: lessons.length,
  };
}

/**
 * 送信データを検証する
 * 呼ばれるとこ: saveShiftsの一番最初
 * 影響先: 画面は触らない。ダメならthrowして呼び出し元(saveShifts)ごと止める
 * 末尾の_はGASの慣習で「外部から直接呼ばせたくない内部用関数」って印
 *
 * どんな関数か: 保存前の最終チェックだけをやるバリデーション系の関数（データは何も作らない）
 * 使ってるメソッド: Array.isArray() / throw new Error() / forEach
 * できあがるもの: 何も返さない（戻り値なし）。問題があればErrorを投げて処理を止めるだけ
 */
function validateRequest_(schoolName, userName, staffStart, staffEnd, selectedDates, lessons) {
  // 1個ずつ順番にチェックして、ダメだったらその場でthrowして終わり（下のチェックは実行されない）
  // 校舎名が空文字ならNG
  if (!schoolName) {
    throw new Error("校舎名を選択してください。");
  }

  // 名前が空文字ならNG
  if (!userName) {
    throw new Error("名前を選択してください。");
  }

  // 配列かどうか＋1件以上あるかを両方チェック
  if (!Array.isArray(selectedDates) || selectedDates.length === 0) {
    throw new Error("勤務日を1日以上選択してください。");
  }

  // 開始・終了どちらかが空文字ならNG
  if (!staffStart || !staffEnd) {
    throw new Error("スタッフ勤務時間を入力してください。");
  }

  // 文字列同士でも"09:00" >= "17:00"みたいに比較できる
  if (staffStart >= staffEnd) {
    throw new Error("スタッフ終了時間は開始時間より後にしてください。");
  }

  // そもそも配列で届いてなければ形式がおかしいのでNG
  if (!Array.isArray(lessons)) {
    throw new Error("授業時間の形式が正しくありません");
  }

  // フロント側でも7件に制限してるが、サーバー側でも念のため二重チェック
  if (lessons.length > 7) {
    throw new Error("授業時間は最大7件までです。");
  }

  // 授業時間を1件ずつチェック（何件目かエラー文に入れたいので、indexが使えるforEach）
  lessons.forEach(function (lesson, index) {
    // 1件分の開始・終了を、空でも安全なようにString+trim
    const lessonStart = String(lesson.start || "").trim();
    const lessonEnd = String(lesson.end || "").trim();

    // どちらかが空文字ならNG
    if (!lessonStart || !lessonEnd) {
      throw new Error("授業" + (index + 1) + "の時間を入力してください。");
    }

    // 開始が終了以降になってたらNG
    if (lessonStart >= lessonEnd) {
      throw new Error("授業" + (index + 1) + "の終了時間は開始時間より後にしてください。");
    }

    // スタッフ勤務時間の範囲内に収まってるかもチェック
    if (lessonStart < staffStart || lessonEnd > staffEnd) {
      throw new Error("授業" + (index + 1) + "はスタッフ勤務時間内に設定してください。");
    }
  });
}
