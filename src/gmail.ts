// NOTE: 自前で定義しないとテストできないので
// https://developers.google.com/apps-script/reference/calendar/event-color?hl=ja
// enum EventColor {
//   PALE_BLUE = '1',
//   PALE_GREEN = '2',
//   MAUVE = '3',
//   PALE_RED = '4',
//   YELLOW = '5',
//   ORANGE = '6',
//   CYAN = '7',
//   GRAY = '8',
//   BLUE = '9',
//   GREEN = '10',
//   RED = '11',
// }

interface Room {
  name: string;
  color: GoogleAppsScript.Calendar.EventColor;
}

interface sheetHeaderSetting {
  reservationID: string;
  icalUID: string;
}

interface sheetSetting {
  name: string;
  header: sheetHeaderSetting;
}

interface Setting {
  homeName: string;
  rooms: Room[];
  sheet: sheetSetting;
}

const rooms: Room[] = [
  { name: '101', color: GoogleAppsScript.Calendar.EventColor.PALE_RED },
  { name: '102', color: GoogleAppsScript.Calendar.EventColor.ORANGE },
  { name: '201', color: GoogleAppsScript.Calendar.EventColor.PALE_BLUE },
  { name: '202', color: GoogleAppsScript.Calendar.EventColor.CYAN },
  { name: '203', color: GoogleAppsScript.Calendar.EventColor.BLUE },
  { name: '204', color: GoogleAppsScript.Calendar.EventColor.PALE_GREEN },
];

const defaultSetting: Setting = {
  homeName: '氷見C邸',
  rooms: rooms,
  sheet: {
    name: 'main',
    header: {
      reservationID: 'reservation_id',
      icalUID: 'ical_uid',
    },
  },
};

const extractPlainBodies = (
  threads: GoogleAppsScript.Gmail.GmailThread[]
): string[] => {
  return threads
    .flatMap(thread => thread.getMessages())
    .map(message => message.getPlainBody());
};

const createSchedulePerDay = (
  gmailApp: GoogleAppsScript.Gmail.GmailApp = GmailApp,
  calendarApp: GoogleAppsScript.Calendar.CalendarApp = CalendarApp,
  spreadsheetApp: GoogleAppsScript.Spreadsheet.SpreadsheetApp = SpreadsheetApp,
  setting: Setting = defaultSetting
): void => {
  const newer_than = '1d';
  // Gmailから過去一日分の予約リクエストメールを取得
  const threads = gmailApp.search(
    `subject:【ADDress】${setting.homeName}：予約リクエスト自動承認のお知らせ newer_than:${newer_than}`
  );

  // メール本文を抽出
  const messagePlainBodies = extractPlainBodies(threads);

  // 本文から必要情報を切り出してデータ格納
  interface ReservationDetail {
    id: string | null;
    room: string;
    reserver: string;
    arriveAt: string | null;
    leaveAt: string | null;
    parking: string;
    raw: string;
  }

  const reservationDetails: ReservationDetail[] = messagePlainBodies.map(
    message => {
      const extractReservation = (
        message: string,
        prefixRegexp: RegExp
      ): string | null => {
        const matched = message.match(new RegExp(prefixRegexp + `(.*)\r\n`));
        return matched ? matched[1] : null;
      };
      const arriveAtStr = extractReservation(message, /・到着日時\s*：/);
      const arriveAt =
        arriveAtStr
          ?.replace(/年|月/g, '/')
          ?.replace(/日 \((\d{2}:\d{2})ごろ\)/, ' $1') || null;
      const leaveAtStr = extractReservation(message, /・出発日時\s*：/);
      const leaveAt =
        leaveAtStr
          ?.replace(/年|月/g, '/')
          ?.replace(/日 \((\d{2}:\d{2})まで\)/, ' $1') || null;
      const parkingStr = extractReservation(message, /・駐車場利用\s*：/);
      const parking = parkingStr === '予約なし' ? '' : '🚗';

      return {
        id: extractReservation(message, /予約ID\s*：/),
        room: extractReservation(message, /部屋番号\s*：/) || '',
        reserver: extractReservation(message, /予約者名\s*：/) || '',
        arriveAt: arriveAt,
        leaveAt: leaveAt,
        parking: parking,
        raw: message,
      };
    }
  );

  // Gカレンダーにスケジュールを作成
  const calendar = calendarApp.getDefaultCalendar();
  return reservationDetails
    .filter(r => {
      console.log(r);
      return r.id && r.arriveAt && r.leaveAt;
    })
    .forEach(reservation => {
      const startAt = new Date(reservation.arriveAt as string);
      const endAt = new Date(reservation.leaveAt as string);
      const title = `${reservation.room} ${reservation.reserver} ${reservation.parking}`;
      const option = { description: reservation.raw };
      const event = calendar.createEvent(title, startAt, endAt, option);
      const eventColor = setting.rooms.find(
        r => r.name === reservation.room
      )?.color;
      if (eventColor) {
        event.setColor(eventColor.toString());
      }
      // スプレッドシートの１行目の、reservationIDの列の最後にreservation.idを記入し、icalUIDの最後にevent.getId()を記入
      const sheet = spreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName(setting.sheet.name);
      if (!sheet) {
        console.error('Sheet not found: ' + setting.sheet.name);
        return;
      }
      const firstRowRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
      const firstRowValues = firstRowRange.getValues()[0];
      const reservationIDColumnIndex =
        firstRowValues.findIndex(
          v => v === setting.sheet.header.reservationID
        ) + 1;
      const icalUIDColumnIndex =
        firstRowValues.findIndex(v => v === setting.sheet.header.icalUID) + 1;
      if (reservationIDColumnIndex === 0) {
        console.error(
          'Column not found: ' + setting.sheet.header.reservationID
        );
        return;
      }
      if (icalUIDColumnIndex === 0) {
        console.error('Column not found: ' + setting.sheet.header.icalUID);
        return;
      }
      const nextRow = sheet.getLastRow() + 1;
      sheet
        .getRange(nextRow, reservationIDColumnIndex)
        .setValue(reservation.id);
      sheet.getRange(nextRow, icalUIDColumnIndex).setValue(event.getId());
    });
};

// const deleteSchedulePerDay = () => {
//   // 設定値の取得
//   // const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
//   // const homeName = sheet.getRange(1, 2).getValue();
//   const homeName = setting.homeName;

//   // Gmailから過去一日分の予約キャンセルメールを取得
//   const threads = GmailApp.search(
//     'subject:【ADDress】' +
//       homeName +
//       '：予約キャンセル受信のお知らせ newer_than:1d'
//   );

//   // メール本文を抽出
//   const texts = [];
//   threads.forEach(thread => {
//     const messages = thread.getMessages();

//     messages.forEach(message => {
//       texts.push(message.getPlainBody());
//     });
//   });

//   // 本文から必要情報を切り出してデータ格納
//   const datas = [];
//   texts.forEach(text => {
//     const obj = {};

//     let start = text.indexOf('■予約ID：');
//     let end = text.indexOf('\r\n', start);
//     let str = text.slice(start + 6, end);
//     obj['id'] = str;

//     start = text.indexOf('・到着日時　　：');
//     end = text.indexOf('日 (', start);
//     str = text.slice(start + 8, end);
//     str = str.replace('年', '/');
//     str = str.replace('月', '/');
//     str += ' ';
//     str += text.substr(end + 3, 5);
//     str += ':00';
//     obj['arriveAt'] = str;

//     start = text.indexOf('・出発日時　　：');
//     end = text.indexOf('日 (', start);
//     str = text.slice(start + 8, end);
//     str = str.replace('年', '/');
//     str = str.replace('月', '/');
//     str += ' ';
//     str += text.substr(end + 3, 5);
//     str += ':00';
//     obj['leaveAt'] = str;

//     datas.push(obj);
//   });

//   // Gカレンダーのスケジュールを削除
//   const calendar = CalendarApp.getDefaultCalendar();
//   datas.forEach(data => {
//     const startAt = new Date(data['arriveAt']);
//     const endAt = new Date(data['leaveAt']);
//     const query = data['id'];
//     const event = calendar.getEvents(startAt, endAt, { search: query });
//     if (event.length) {
//       event[0].deleteEvent();
//     }
//   });
// };
// const updateSchedulePerDay = () => {};

export {
  createSchedulePerDay,
  // deleteSchedulePerDay,
  // updateSchedulePerDay
};

export type { Room, Setting };
