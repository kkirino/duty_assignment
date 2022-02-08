function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("サイドバーを開く")
    .addItem("開く", "openSidebar")
    .addToUi();
}

function openSidebar() {
  const htmlOutput = HtmlService.createTemplateFromFile("sidebar").evaluate();
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

const ss = SpreadsheetApp.getActiveSpreadsheet();

function getSettings() {
  const ws = ss.getSheetByName("settings");
  const titles = ws
    .getRange(2, 1, ws.getLastRow() - 1, 1)
    .getValues()
    .map(function (row) {
      return row[0];
    });
  const values = ws
    .getRange(2, 2, ws.getLastRow() - 1, 1)
    .getValues()
    .map(function (row) {
      return row[0];
    });
  const settings = {};
  settings.calendarIdInside =
    values[titles.indexOf("大学病院当番のカレンダーID")];
  settings.calendarIdOutside =
    values[titles.indexOf("ネーベン・外勤のカレンダーID")];
  settings.manualPageUrl = values[titles.indexOf("マニュアルページのURL")];
  holidayStringsArray = values[titles.indexOf("追加の祝日")].split(",");
  settings.holidays = holidayStringsArray.map(function (value) {
    return {
      month: value.split("/")[0] - 1,
      date: value.split("/")[1],
    };
  });
  return settings;
}

function getDefinition() {
  const ws = ss.getSheetByName("definition");
  const sheetInfo = ws
    .getRange(2, 1, ws.getLastRow() - 1, ws.getLastColumn())
    .getValues();
  const definition = [];
  const dayOfWeekJa = ["日", "月", "火", "水", "木", "金", "土"];
  sheetInfo.forEach(function (row) {
    for (let i = 0; i < 1 + 4 * Number(row[2] === 99); i++) {
      for (let j = 0; j < 1 + Number(row[5] === "通常・祝日共通"); j++) {
        for (let k = 0; k < 1 + Number(row[9] === "通常・祝日共通"); k++) {
          definition.push({
            hospital: row[0],
            work: row[1],
            weekInMonthStart:
              (i + 1) * Number(row[2] === 99) + row[2] * Number(row[2] !== 99),
            dayOfWeekStart: dayOfWeekJa.indexOf(row[3]),
            isSaturdayCount: Number(row[4] === "はい"),
            isHolidayStart:
              j * Number(row[5] === "通常・祝日共通") +
              Number(row[5] === "祝日"),
            hoursStart: row[6],
            minutesStart: row[7],
            isNextDayEnd: Number(row[8] === "翌日"),
            isHolidayEnd:
              k * Number(row[9] === "通常・祝日共通") +
              Number(row[9] === "祝日"),
            hoursEnd: row[10],
            minutesEnd: row[11],
            doctorInDefault: row[12],
          });
        }
      }
    }
  });
  return definition;
}

function isHoliday(date) {
  const holidaysInCalendar = CalendarApp.getCalendarById(
    "ja.japanese#holiday@group.v.calendar.google.com"
  );
  const holidaysInSettings = getSettings().holidays;
  function reduceSum(previousValue, currentValue) {
    return previousValue + currentValue;
  }
  return Number(
    holidaysInCalendar.getEventsForDay(date).length > 0 ||
      holidaysInSettings
        .map(function (holiday) {
          holiday.month === date.getMonth() && holiday.date === date.getDate();
        })
        .reduce(reduceSum) > 0
  );
}

function getDateInfo(userInfo) {
  const dateFrom = new Date(userInfo.dateFrom);
  const dateTo = new Date(userInfo.dateTo);
  dateFrom.setHours(0);
  dateTo.setHours(0);
  const duration =
    (dateTo.valueOf() - dateFrom.valueOf()) / (1000 * 60 * 60 * 24) + 1;
  const dateInfo = Array.from(new Array(duration)).map(function (v, i) {
    let dateStart = new Date(dateFrom.valueOf() + 1000 * 60 * 60 * 24 * i);
    let datePrevious = new Date(dateStart.valueOf() - 1000 * 60 * 60 * 24);
    let dateNext = new Date(dateStart.valueOf() + 1000 * 60 * 60 * 24);
    return {
      dateStart: dateStart,
      weekInMonthStart: [
        Math.floor((dateStart.getDate() + 6) / 7),
        Math.floor((datePrevious.getDate() + 6) / 7),
      ],
      dayOfWeekStart: dateStart.getDay(),
      isHolidayStart: isHoliday(dateStart),
      dateEnd: [dateStart, dateNext],
      isHolidayEnd: [isHoliday(dateStart), isHoliday(dateNext)],
    };
  });
  return dateInfo
}

function createSchedules(userInfo) {
  if (userInfo.dateFrom == "" || userInfo.dateTo == "") {
    return "未入力です";
  } else {
    try {
      const dateInfo = getDateInfo(userInfo);
      const definition = getDefinition();
      const inside = ss.insertSheet();
      inside.setName("【作業中】大学当番");
      const outside = ss.insertSheet();
      outside.setName("【作業中】外勤");
      dateInfo.forEach(function (date) {
        if (
          date.dayOfWeekStart === 0 ||
          date.dayOfWeekStart === 6 ||
          date.isHolidayStart === 1
        ) {
          ["当直", "Super", "オンコール"].forEach(function (value) {
            inside.appendRow([date.dateStart.toLocaleDateString(), value]);
          });
        } else if (
          date.dayOfWeekStart === 1 ||
          date.dayOfWeekStart === 3 ||
          date.dayOfWeekStart === 5
        ) {
          ["外来係", "当直", "Super", "オンコール", "急患センター"].forEach(
            function (value) {
              inside.appendRow([date.dateStart.toLocaleDateString(), value]);
            }
          );
        } else {
          ["当直", "Super", "オンコール", "急患センター"].forEach(function (
            value
          ) {
            inside.appendRow([date.dateStart.toLocaleDateString(), value]);
          });
        }
        definition.forEach(function (def) {
          if (
            def.weekInMonthStart ===
              date.weekInMonthStart[def.isSaturdayCount] &&
            def.dayOfWeekStart === date.dayOfWeekStart &&
            def.isHolidayStart === date.isHolidayStart &&
            def.isHolidayEnd === date.isHolidayEnd[def.isNextDayEnd]
          ) {
            outside.appendRow([
              def.hospital,
              def.work,
              def.doctorInDefault,
              date.dateStart.toLocaleDateString(),
              def.hoursStart,
              def.minutesStart,
              date.dateEnd[def.isNextDayEnd].toLocaleDateString(),
              def.hoursEnd,
              def.minutesEnd,
            ]);
          }
        });
      });
      return "予定を作成しました";
    } catch {
      return "予定を作成できませんでした";
    }
  }
}

function writeCalendarEvents(isCheckedInside, isCheckedOutside) {
  if (!isCheckedInside && !isCheckedOutside) {
    return "チェックボックスにチェックを入れてください";
  } else {
    try {
      if (isCheckedOutside) {
        const calendarId = getSettings().calendarIdOutside;
        const calendar = CalendarApp.getCalendarById(calendarId);
        const ws = ss.getSheetByName("【作業中】外勤");
        const sheetInfo = ws
          .getRange(1, 1, ws.getLastRow(), ws.getLastColumn())
          .getValues();
        sheetInfo.forEach(function (row) {
          let title = row[2] + " " + row[0] + row[1];
          let startTime = new Date(row[3]);
          startTime.setHours(row[4]);
          startTime.setMinutes(row[5]);
          let endTime = new Date(row[6]);
          endTime.setHours(row[7]);
          endTime.setMinutes(row[8]);
          calendar.createEvent(title, startTime, endTime);
        });
        const today = new Date();
        ws.setName(
          today.toLocaleDateString() + " " + today.toLocaleTimeString()
        );
      }
      if (isCheckedInside) {
        const calendarId = getSettings().calendarIdInside;
        const calendar = CalendarApp.getCalendarById(calendarId);
        const ws = ss.getSheetByName("【作業中】大学当番");
        const sheetInfo = ws
          .getRange(1, 1, ws.getLastRow(), ws.getLastColumn())
          .getValues();
          Logger.log(sheetInfo);
        sheetInfo.forEach(function (row) {
          let title = "【" + row[1] + "】" + row[2];
          let date = new Date(row[0]);
          calendar.createAllDayEvent(title, date);
        });
        const today = new Date();
        ws.setName(
          today.toLocaleDateString() + " " + today.toLocaleTimeString()
        );
      }
      return "予定をカレンダーに登録しました";
    } catch {
      return "予定をカレンダーに登録できませんでした";
    }
  }
}
