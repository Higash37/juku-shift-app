function doGet() {
  return (
    HtmlService.createTemplateFromFile("index")
      // HtmlTemplateをdoGetが返せるアウトプットに変換する
      // テンプレートに埋め込まれた<?= ... ?>や<?!= ... ?>を実際に実行して、結果をHTMLに埋め込む
      // サーバーの値をindex.htmlに渡したくなったら使う
      .evaluate()
      .setTitle("シフト提出")
      // 画面幅に合わせて正しい大きさで表示するためのタグ
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
  );
}

// スプシからスタッフメンバーを持ってくる
function getMembers() {
  // スプシを読み込む
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // スタッフメンバーが載っているシートを呼ぶ
  const sheet = spreadsheet.getSheetByName("members");

  // シートがなければエラーを吐く
  if (!sheet) {
    throw new Error("membersシートが見つかりません。");
  }
  // ヘッダー以外にデータがあるか確認する
  // なければ空配列を返す
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      schools: [],
      members: [],
    };
  }

  // 2行目1列目から最終行の2列目までの値を回収
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  // valuesは2次元配列を想定
  // [
  //   ["田中", "渋谷校"],
  //   ["", "新宿校"],      // 名前が空の行(ゴミデータ)
  //   ["鈴木", "渋谷校"],
  // ]
  // 2次元配列からオブジェクト配列に変更
  // 毎回row[0]などと書かないといけないため
  const members = values
    // オブジェクト配列をmapで作成
    .map(function (row) {
      return {
        // 要素をオブジェクト形式で保存
        // trim()で後空白は除去
        name: String(row[0] || "").trim(),
        school_name: String(row[1] || "").trim(),
      };
    })
    // 名前か校舎名がないrowははじく
    // 基本的に全員どこかしらの校舎には所属しているため
    .filter(function (member) {
      return member.name !== "" && member.school_name !== "";
    });

  // membersから校舎名だけの配列を作成
  // ソートにて並び替え、Set()にて重複を避ける
  const schools = [
    ...new Set(
      members.map(function (member) {
        return member.school_name;
      }),
    ),
  ].sort();
  return {
    schools: schools,
    members: members,
  };
}

function getTimeOptions() {
  // スプレッドシートファイルを取得
  const settingSheet = SpreadsheetApp.getActiveSpreadsheet();
  // 設定用のファイルを取得
  const settings = settingSheet.getSheetByName("settings");

  // なければエラーを吐く
  if (!settings) {
    throw new Error("settingsシートが見つかりません");
  }
  // 設定シートの最後の行番号を取得
  const settingLastRow = settings.getLastRow();
  // ヘッダーしかなければ空配列を返す
  if (settingLastRow < 2) {
    return {
      times: [],
    };
  }
  // 設定シートの2行目から最終行までの1列を取得
  const timeValues = settings.getRange(2, 1, settingLastRow - 1, 1).getValues();
  //  選択肢になる時間配列[9:00, 9:30, ...]を作る
  const times = timeValues
    .map(function (row) {
      // value に時間を入れる
      const value = row[0];
      // もし時間が日付型だと日付・時刻・曜日・タイムゾーンを含んだ長い形になるのでinstanceofでvalue型にする
      if (value instanceof Date) {
        // Sessionを使うことでGASのタイムゾーンを取得できる
        // valueを"HH:mm"の形にする
        return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
      }
      // もしDate型でなければString型として返す
      return String(value || "").trim();
    })
    // 時間がない行が最終行までの途中にあればフィルタリングする
    .filter(function (time) {
      return time !== "";
    });

  // 時間配列を返す
  return {
    times: times,
  };
}

// HTMLから送られてきたリクエストを受け取り、スプレッドシートに保存する
function saveShifts(request) {
  // requestがなければエラーを吐く
  // requestの未定義状態で処理を進めてしまうクラッシュ対応
  if (!request) {
    throw new Error("送信データがありません。");
  }

  // それぞれ要素をストリング型で保存
  const schoolName = String(request.school_name || "").trim();
  const userName = String(request.user_name || "").trim();
  const staffStart = String(request.staff_start || "").trim();
  const staffEnd = String(request.staff_end || "").trim();

  // 下記は日付文字型の配列とオブジェクト配列を想定。そのためチェックは別の場所に
  const selectedDates = request.selected_dates;
  const lessons = request.lessons;

  //
  validateRequest_(schoolName, userName, staffStart, staffEnd, selectedDates, lessons);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // シフトシートを取得
  const shiftsSheet = spreadsheet.getSheetByName("shifts_DB");
  // 授業シートを取得
  const lessonsSheet = spreadsheet.getSheetByName("lessons_DB");

  // シフトシート・授業シートがなければエラーを吐く
  if (!shiftsSheet) {
    throw new Error("shifts_DBシートが見つかりません。");
  }

  if (!lessonsSheet) {
    throw new Error("lessons_DBシートが見つかりません。");
  }

  // シフト特定用のユニークなidを生成
  const batchId = Utilities.getUuid();
  // 作成時を生成
  const timestamp = new Date();
  const shiftRows = [];
  const lessonRows = [];

  // 日付の数だけシフトと授業をshift用と授業用で
  selectedDates.forEach(function (date) {
    const shiftId = Utilities.getUuid();

    // shiftRowsに1日のシフトをプッシュする
    shiftRows.push([
      shiftId,
      batchId,
      schoolName,
      userName,
      date,
      staffStart,
      staffEnd,
      "PENDING",
      timestamp,
    ]);

    // 授業日はシフト1日につき複数送信可能なため、forEachの中で再度授業用forEachをする
    lessons.forEach(function (lesson, index) {
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

  // シフトシートの最終行の次の行にシフト配列を入れる
  shiftsSheet
    .getRange(shiftsSheet.getLastRow() + 1, 1, shiftRows.length, shiftRows[0].length)
    .setValues(shiftRows);

  // 授業用シートの最終列に授業を入れる
  // ただ、授業が空で送信された際にlessonRowsがundefinedとなりエラー表示が出てしまうのでifで囲む
  if (lessonRows.length > 0) {
    lessonsSheet
      .getRange(lessonsSheet.getLastRow() + 1, 1, lessonRows.length, lessonRows[0].length)
      .setValues(lessonRows);
  }

  // 送信ボタンの下に送信内容表示のためカウントを返す
  return {
    success: true, // 成功フラグ
    shift_count: shiftRows.length, //シフト数
    lesson_count: lessonRows.length, // 1日につきの授業総数
    selected_date_count: selectedDates.length, // 選択された日付数
    lesson_type_count: lessons.length, // 実際に書き込まれる授業数
  };
}

// 入力情報が正しいのかをバリデーション
function validateRequest_(schoolName, userName, staffStart, staffEnd, selectedDates, lessons) {
  if (!schoolName) {
    throw new Error("校舎名を選択してください。");
  }

  if (!userName) {
    throw new Error("名前を選択してください。");
  }

  if (!Array.isArray(selectedDates) || selectedDates.length === 0) {
    throw new Error("勤務日を1日以上選択してください。");
  }

  if (!staffStart || !staffEnd) {
    throw new Error("スタッフ勤務時間を入力してください。");
  }

  if (staffStart >= staffEnd) {
    throw new Error("スタッフ終了時間は開始時間より後にしてください。");
  }

  if (!Array.isArray(lessons)) {
    throw new Error("授業時間の形式が正しくありません");
  }

  if (lessons.length > 7) {
    throw new Error("授業時間は最大7件までです。");
  }

  lessons.forEach(function (lesson, index) {
    const lessonStart = String(lesson.start || "").trim();
    const lessonEnd = String(lesson.end || "").trim();

    if (!lessonStart || !lessonEnd) {
      throw new Error("授業" + (index + 1) + "の時間を入力してください。");
    }

    if (lessonStart >= lessonEnd) {
      throw new Error("授業" + (index + 1) + "の終了時間は開始時間より後にしてください。");
    }

    if (lessonStart < staffStart || lessonEnd > staffEnd) {
      // 講師から「授業とスタッフ勤務時間の前後関係がわかりにくい」という声があったが、
      // 将来のシフト差し引き機能のためにこの制約自体は維持し、
      // フォーム側の説明文で分かりやすくする方針にした。
      throw new Error("授業" + (index + 1) + "はスタッフ勤務時間内に設定してください。");
    }
  });
}
