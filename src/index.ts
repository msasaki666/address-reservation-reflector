/**
 * Copyright 2023 Motoaki Sasaki
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

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
  newer_than: string;
}

const rooms: Room[] = [
  { name: '101', color: CalendarApp.EventColor.PALE_RED },
  { name: '102', color: CalendarApp.EventColor.ORANGE },
  { name: '201', color: CalendarApp.EventColor.PALE_BLUE },
  { name: '202', color: CalendarApp.EventColor.CYAN },
  { name: '203', color: CalendarApp.EventColor.BLUE },
  { name: '204', color: CalendarApp.EventColor.PALE_GREEN },
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
  newer_than: '1d',
};

const extractPlainBodies = (
  threads: GoogleAppsScript.Gmail.GmailThread[]
): string[] => {
  return threads
    .flatMap(thread => thread.getMessages())
    .map(message => message.getPlainBody());
};

const extractSpecificReservationDetail = (
  plainBody: string,
  prefixRegexp: RegExp
): string | null => {
  const matched = plainBody.match(
    new RegExp(prefixRegexp.source + `(.*)(?:\r\n|\n)`)
  );
  return matched ? matched[1] : null;
};

// eslint-disable-next-line @typescript-eslint/no-unused-vars
const createSchedulePerDay = (
  gmailApp: GoogleAppsScript.Gmail.GmailApp = GmailApp,
  calendarApp: GoogleAppsScript.Calendar.CalendarApp = CalendarApp,
  spreadsheetApp: GoogleAppsScript.Spreadsheet.SpreadsheetApp = SpreadsheetApp,
  setting: Setting = defaultSetting
): void => {
  const threads = gmailApp.search(
    `subject:【ADDress】${setting.homeName}：予約リクエスト自動承認のお知らせ newer_than:${setting.newer_than}`
  );
  const messagePlainBodies = extractPlainBodies(threads);
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
      const arriveAtStr = extractSpecificReservationDetail(
        message,
        /到着日時\s*：/
      );
      const arriveAt =
        arriveAtStr
          ?.replace(/年|月/g, '/')
          ?.replace(/日 \((\d{2}:\d{2})ごろ\)/, ' $1') || null;
      const leaveAtStr = extractSpecificReservationDetail(
        message,
        /出発日時\s*：/
      );
      const leaveAt =
        leaveAtStr
          ?.replace(/年|月/g, '/')
          ?.replace(/日 \((\d{2}:\d{2})まで\)/, ' $1') || null;
      const parkingStr = extractSpecificReservationDetail(
        message,
        /駐車場利用\s*：/
      );
      const parking = parkingStr === '予約なし' ? '' : '🚗';

      return {
        id: extractSpecificReservationDetail(message, /予約ID\s*：/),
        room: extractSpecificReservationDetail(message, /部屋番号\s*：/) || '',
        reserver:
          extractSpecificReservationDetail(message, /予約者名\s*：/) || '',
        arriveAt: arriveAt,
        leaveAt: leaveAt,
        parking: parking,
        raw: message,
      };
    }
  );

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

// eslint-disable-next-line @typescript-eslint/no-unused-vars
const deleteSchedulePerDay = (
  gmailApp: GoogleAppsScript.Gmail.GmailApp = GmailApp,
  calendarApp: GoogleAppsScript.Calendar.CalendarApp = CalendarApp,
  spreadsheetApp: GoogleAppsScript.Spreadsheet.SpreadsheetApp = SpreadsheetApp,
  setting: Setting = defaultSetting
): void => {
  const threads = gmailApp.search(
    `subject:【ADDress】${setting.homeName}：予約キャンセル受信のお知らせ newer_than:${setting.newer_than}`
  );
  const messagePlainBodies = extractPlainBodies(threads);

  const extractedReservationIDs = messagePlainBodies
    .map(message => {
      return extractSpecificReservationDetail(message, /予約ID\s*：/);
    })
    .filter(id => id);

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
    firstRowValues.findIndex(v => v === setting.sheet.header.reservationID) + 1;
  const icalUIDColumnIndex =
    firstRowValues.findIndex(v => v === setting.sheet.header.icalUID) + 1;
  if (reservationIDColumnIndex === 0) {
    console.error('Column not found: ' + setting.sheet.header.reservationID);
    return;
  }
  if (icalUIDColumnIndex === 0) {
    console.error('Column not found: ' + setting.sheet.header.icalUID);
    return;
  }
  const offset = 2;
  const reservationIDInputtedRange = sheet.getRange(
    offset,
    reservationIDColumnIndex,
    sheet.getLastRow(),
    reservationIDColumnIndex
  );

  const reservationIDs = reservationIDInputtedRange
    .getValues()
    .map((v): string => {
      return v[0];
    });
  extractedReservationIDs.forEach(id => {
    const index = reservationIDs.indexOf(id as string);
    if (index === -1) return;

    const deletionTargetRowIndex = index + offset;
    const icalUID = sheet
      .getRange(deletionTargetRowIndex, icalUIDColumnIndex)
      .getValue();
    const event = calendarApp.getEventById(icalUID);
    if (event) {
      event.deleteEvent();
      sheet.deleteRow(deletionTargetRowIndex);
    }
  });
};
// const updateSchedulePerDay = () => {};
