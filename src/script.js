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
const dayOfWeekJa = ["日", "月", "火", "水", "木", "金", "土"];

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
  const holidayStrings = values[titles.indexOf("追加の祝日")].split(",");
  settings.holidays = holidayStrings.map(function (value) {
    return {
      month: value.split("/")[0] - 1,
      date: value.split("/")[1],
    };
  });
  const outpatientDayStrings =
    values[titles.indexOf("大学当番外来係の曜日")].split(",");
  const firstAidCenterDayStrings =
    values[titles.indexOf("大学当番急患センターの曜日")].split(",");
  settings.outpatientDays = [];
  settings.firstAidCenterDays = [];
  outpatientDayStrings.forEach(function (value) {
    if (dayOfWeekJa.indexOf(value) > -1) {
      settings.outpatientDays.push(dayOfWeekJa.indexOf(value));
    }
  });
  firstAidCenterDayStrings.forEach(function (value) {
    if (dayOfWeekJa.indexOf(value) > -1) {
      settings.firstAidCenterDays.push(dayOfWeekJa.indexOf(value));
    }
  });
  return settings;
}

function getDefinition() {
  const ws = ss.getSheetByName("definition");
  const sheetInfo = ws
    .getRange(2, 1, ws.getLastRow() - 1, ws.getLastColumn())
    .getValues();
  const definition = [];
  sheetInfo.forEach(function (row) {
    for (let i = 0; i < 1 + 4 * Number(row[2] === 99); i++) {
      for (let j = 0; j < 1 + Number(row[5] === "通常・祝日共通"); j++) {
        for (let k = 0; k < 1 + Number(row[6] === "通常・祝日共通"); k++) {
          for (let l = 0; l < 1 + Number(row[10] === "通常・祝日共通"); l++) {
            definition.push({
              hospital: row[0],
              work: row[1],
              weekInMonthStart:
                (i + 1) * Number(row[2] === 99) +
                row[2] * Number(row[2] !== 99),
              dayOfWeekStart: dayOfWeekJa.indexOf(row[3]),
              isSaturdayCount: Number(row[4] === "はい"),
              isHolidayPrevious:
                j * Number(row[5] === "通常・祝日共通") +
                Number(row[5] === "祝日"),
              isHolidayStart:
                k * Number(row[6] === "通常・祝日共通") +
                Number(row[6] === "祝日"),
              hoursStart: row[7],
              minutesStart: row[8],
              isNextDayEnd: Number(row[9] === "翌日"),
              isHolidayEnd:
                l * Number(row[10] === "通常・祝日共通") +
                Number(row[10] === "祝日"),
              hoursEnd: row[11],
              minutesEnd: row[12],
              doctorInDefault: row[13],
            });
          }
        }
      }
    }
  });
  return definition;
}

function reduceSum(previousValue, currentValue) {
  return previousValue + currentValue;
}

function isHoliday(date) {
  const holidaysInCalendar = CalendarApp.getCalendarById(
    "ja.japanese#holiday@group.v.calendar.google.com"
  );
  const holidaysInSettings = getSettings().holidays;
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
    const dateStart = new Date(dateFrom.valueOf() + 1000 * 60 * 60 * 24 * i);
    const datePrevious = new Date(dateStart.valueOf() - 1000 * 60 * 60 * 24);
    const dateNext = new Date(dateStart.valueOf() + 1000 * 60 * 60 * 24);
    return {
      dateStart: dateStart,
      weekInMonthStart: [
        Math.floor((dateStart.getDate() + 6) / 7),
        Math.floor((datePrevious.getDate() + 6) / 7),
      ],
      dayOfWeekStart: dateStart.getDay(),
      isHolidayPrevious: isHoliday(datePrevious),
      isHolidayStart: isHoliday(dateStart),
      dateEnd: [dateStart, dateNext],
      isHolidayEnd: [isHoliday(dateStart), isHoliday(dateNext)],
    };
  });
  return dateInfo;
}

function createSchedules(userInfo) {
  if (userInfo.dateFrom == "" || userInfo.dateTo == "") {
    return "未入力です";
  } else {
    try {
      const dateInfo = getDateInfo(userInfo);
      const definition = getDefinition();
      const settings = getSettings();
      const tmpForInside = ss.getSheetByName("template_大学当番");
      const tmpForOutside = ss.getSheetByName("template_外勤");
      const sheetNames = ss.getSheets().map(function (sheet) {
        return sheet.getName();
      });
      const sheetNamesWithOldDate = sheetNames
        .filter(function (string) {
          const regex = new RegExp(
            "^20[0-9]{2}/[0-9]{1,}/[0-9]{1,} [0-9]{1,}:[0-9]{1,}:[0-9]{1,}$"
          );
          return string.search(regex) > -1;
        })
        .filter(function (string) {
          return (
            (new Date().valueOf() - new Date(string).valueOf()) /
              (1000 * 60 * 60 * 24) >
            90
          );
        });
      sheetNamesWithOldDate.forEach(function (name) {
        const sheetForDelete = ss.getSheetByName(name);
        ss.deleteSheet(sheetForDelete);
      });
      const sheetNameForInside = "【作業中】大学当番";
      const sheetNameForOutside = "【作業中】外勤";
      if (
        sheetNames.indexOf(sheetNameForInside) > -1 ||
        sheetNames.indexOf(sheetNameForOutside) > -1
      ) {
        return "作業中シートが既に存在します。新たなシートは作成しませんでした。";
      } else {
        const inside = tmpForInside.copyTo(ss);
        inside.setName(sheetNameForInside);
        inside
          .getRange(2, 1, inside.getLastRow() - 1, inside.getLastColumn())
          .clearContent();
        const outside = tmpForOutside.copyTo(ss);
        outside.setName(sheetNameForOutside);
        outside
          .getRange(2, 1, outside.getLastRow() - 1, outside.getLastColumn())
          .clearContent();
        const insideInfo = [];
        const outsideInfo = [];
        dateInfo.forEach(function (date) {
          const needForOutpatient =
            settings.outpatientDays
              .map(function (i) {
                return i === date.dayOfWeekStart;
              })
              .reduce(reduceSum) > 0;
          const needForFirstAidCenter =
            settings.firstAidCenterDays
              .map(function (i) {
                return i === date.dayOfWeekStart;
              })
              .reduce(reduceSum) > 0;
          ["当直", "Super", "オンコール"].forEach(function (value) {
            insideInfo.push([date.dateStart.toLocaleDateString(), value]);
          });
          if (needForFirstAidCenter && !isHoliday(date.dateStart)) {
            insideInfo.push([
              date.dateStart.toLocaleDateString(),
              "急患センター",
            ]);
          }
          if (needForOutpatient && !isHoliday(date.dateStart)) {
            // inside.appendRow([date.dateStart.toLocaleDateString(), "外来係"]);
            insideInfo.push([date.dateStart.toLocaleDateString(), "外来係"]);
          }
          definition
            .filter(function (def) {
              return (
                def.weekInMonthStart ===
                  date.weekInMonthStart[def.isSaturdayCount] &&
                def.dayOfWeekStart === date.dayOfWeekStart &&
                def.isHolidayPrevious === date.isHolidayPrevious &&
                def.isHolidayStart === date.isHolidayStart &&
                def.isHolidayEnd === date.isHolidayEnd[def.isNextDayEnd]
              );
            })
            .forEach(function (def) {
              outsideInfo.push([
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
            });
        });
        inside.getRange(2, 1, insideInfo.length, 2).setValues(insideInfo);
        outside.getRange(2, 1, outsideInfo.length, 9).setValues(outsideInfo);
        return "予定を作成しました";
      }
    } catch (e) {
      Logger.log(e);
      return "予定を作成できませんでした";
    }
  }
}

function writeCalendarEvents(isCheckedInside, isCheckedOutside) {
  if (!isCheckedInside && !isCheckedOutside) {
    return "どちらかの予定を選んでください";
  } else {
    if (isCheckedInside) {
      try {
        const calendarId = getSettings().calendarIdInside;
        const calendar = CalendarApp.getCalendarById(calendarId);
        const ws = ss.getSheetByName("【作業中】大学当番");
        const sheetInfo = ws
          .getRange(2, 1, ws.getLastRow() - 1, ws.getLastColumn())
          .getValues();
        const schedules = [];
        let error = 0;
        sheetInfo.forEach(function (row, index) {
          const title = "【" + row[1] + "】" + row[2];
          const date = new Date(row[0]);
          const blankExists =
            row
              .map(function (value) {
                return value === "";
              })
              .reduce(reduceSum) > 0;
          if (blankExists) {
            ws.getRange(2 + index, 4).setValue(
              "この行に空欄があります(あとで消してください)"
            );
            error += 1;
          }
          schedules.push({
            title: title,
            date: date,
          });
        });
        if (error > 0) {
          return "予定の記載に空欄があります";
        } else {
          schedules.forEach(function (schedule) {
            calendar.createAllDayEvent(schedule.title, schedule.date);
            Utilities.sleep(500);
          });
          const today = new Date();
          ws.setName(
            today.toLocaleDateString() + " " + today.toLocaleTimeString()
          );
          const protection = ws.protect();
          protection.setWarningOnly(true);
          return "予定をカレンダーに登録しました";
        }
      } catch (e) {
        Logger.log(e);
        return "予定をカレンダーに登録できませんでした";
      }
    } else {
      try {
        const calendarId = getSettings().calendarIdOutside;
        const calendar = CalendarApp.getCalendarById(calendarId);
        const ws = ss.getSheetByName("【作業中】外勤");
        const sheetInfo = ws
          .getRange(2, 1, ws.getLastRow() - 1, ws.getLastColumn())
          .getValues();
        const schedules = [];
        let error = 0;
        sheetInfo.forEach(function (row, index) {
          const title = row[2] + " " + row[0] + row[1];
          const startTime = new Date(row[3]);
          startTime.setHours(row[4]);
          startTime.setMinutes(row[5]);
          const endTime = new Date(row[6]);
          endTime.setHours(row[7]);
          endTime.setMinutes(row[8]);
          const timeErrorExists = startTime.valueOf() >= endTime.valueOf();
          const blankExists =
            row
              .map(function (value) {
                return value === "";
              })
              .reduce(reduceSum) > 0;
          if (timeErrorExists || blankExists) {
            ws.getRange(2 + index, 10).setValue(
              "この行の記載に空欄または誤りがあります(あとで消してください)"
            );
            error += 1;
          }
          schedules.push({
            title: title,
            startTime: startTime,
            endTime: endTime,
          });
        });
        if (error > 0) {
          return "予定の記載に空欄または誤りがあります";
        } else {
          schedules.forEach(function (schedule) {
            calendar.createEvent(
              schedule.title,
              schedule.startTime,
              schedule.endTime
            );
            Utilities.sleep(500);
          });
          const today = new Date();
          ws.setName(
            today.toLocaleDateString() + " " + today.toLocaleTimeString()
          );
          const protection = ws.protect();
          protection.setWarningOnly(true);
          return "予定をカレンダーに登録しました";
        }
      } catch (e) {
        Logger.log(e);
        return "予定をカレンダーに登録できませんでした";
      }
    }
  }
}
