/**
 * Webアプリとして画面を表示する
 */
function doGet() {
  return HtmlService.createTemplateFromFile("Index").evaluate().setTitle("シフト提出");
}

/**
 * membersシートから
 * 校舎一覧とメンバー一覧を取得する
 *
 * A列：name
 * B列：school_name
 */
function getMembers() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = spreadsheet.getSheetByName("members");

  if (!sheet) {
    throw new Error("membersシートが見つかりません。");
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      schools: [],
      members: [],
    };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  const members = values
    .map(function (row) {
      return {
        name: String(row[0] || "").trim(),

        school_name: String(row[1] || "").trim(),
      };
    })
    .filter(function (member) {
      return member.name !== "" && member.school_name !== "";
    });

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

/**
 * シフトと授業時間を保存する
 */
function saveShifts(request) {
  if (!request) {
    throw new Error("送信データがありません。");
  }

  const schoolName = String(request.school_name || "").trim();

  const userName = String(request.user_name || "").trim();

  const staffStart = String(request.staff_start || "").trim();

  const staffEnd = String(request.staff_end || "").trim();

  const selectedDates = request.selected_dates;

  const lessons = request.lessons;

  validateRequest_(schoolName, userName, staffStart, staffEnd, selectedDates, lessons);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const shiftsSheet = spreadsheet.getSheetByName("shifts_DB");

  const lessonsSheet = spreadsheet.getSheetByName("lessons_DB");

  if (!shiftsSheet) {
    throw new Error("shifts_DBシートが見つかりません。");
  }

  if (!lessonsSheet) {
    throw new Error("lessons_DBシートが見つかりません。");
  }

  const batchId = Utilities.getUuid();

  const timestamp = new Date();

  const shiftRows = [];
  const lessonRows = [];

  selectedDates.forEach(function (date) {
    const shiftId = Utilities.getUuid();

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

  shiftsSheet
    .getRange(shiftsSheet.getLastRow() + 1, 1, shiftRows.length, shiftRows[0].length)
    .setValues(shiftRows);

  lessonsSheet
    .getRange(lessonsSheet.getLastRow() + 1, 1, lessonRows.length, lessonRows[0].length)
    .setValues(lessonRows);

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
 */
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
      throw new Error("授業" + (index + 1) + "はスタッフ勤務時間内に設定してください。");
    }
  });
}
