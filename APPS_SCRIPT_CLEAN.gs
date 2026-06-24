const TZ = 'Asia/Seoul';

const SHEET_NAMES = {
  VISIT: '방문일지',
  PRINTER: '프린터',
  CAREER: '커리어존',
  CONNECT: '커넥트룸',
  MEMBERS: 'APP_MEMBERS',
  LOUNGE_GUESTS: 'APP_LOUNGE_GUESTS',
  SETTINGS: 'APP_SETTINGS',
  VISIT_LOGS: 'APP_VISIT_LOGS',
  PRINTER_LOGS: 'APP_PRINTER_LOGS',
  CAREER_LOGS: 'APP_CAREER_LOGS',
  CONNECT_LOGS: 'APP_CONNECT_LOGS',
  NAVER_SYNC: '네이버예약_동기화'
};

const HEADERS = {
  visit: ['연번', '날짜', '성별', '나이'],
  printer: ['연번', '날짜', '성별', '나이', '사용 매수'],
  career: [
    '연번', '날짜', '사용목적', '인원수', '비고', '시작시간', '종료시간',
    '20~29세(남)', '20~29세(여)', '30~39세(남)', '30~39세(여)', '~19세(남)', '~19세(여)'
  ],
  connect: [
    '연번', '날짜', '사용목적', '인원수', '비고', '시작시간', '종료시간',
    '20~29세(남)', '20~29세(여)', '30~39세(남)', '30~39세(여)', '~19세(남)', '~19세(여)'
  ],
  members: ['memberId', 'name', 'gender', 'age', 'phone', 'phoneLastDigits', 'status', 'role', 'createdAt', 'updatedAt'],
  loungeGuests: ['guestId', 'name', 'gender', 'birthDate', 'age', 'role', 'validDate', 'createdAt'],
  settings: ['key', 'value', 'updatedAt'],
  visitLogs: ['logId', 'userId', 'memberType', 'name', 'gender', 'age', 'phone', 'phoneLastDigits', 'date', 'source', 'createdBy', 'createdAt'],
  printerLogs: ['logId', 'userId', 'name', 'gender', 'age', 'phone', 'phoneLastDigits', 'date', 'count', 'createdBy', 'createdAt'],
  roomLogs: ['logId', 'userId', 'spaceType', 'name', 'memberNameKey', 'gender', 'age', 'phone', 'phoneLastDigits', 'date', 'startTime', 'endTime', 'purpose', 'companionsJson', 'companionCount', 'createdBy', 'createdAt', 'linkedReservationId', 'dedupeStatus'],
  naverSync: [
    'reservationId', 'status', 'bookerName', 'maskedPhone', 'usageDate', 'startTime', 'endTime',
    'placeName', 'normalizedPlace', 'headcount', 'purpose', 'requestedAt', 'confirmedAt',
    'completedAt', 'cancelledAt', 'cancelReason', 'calendarEventId', 'calendarStatus',
    'onsiteMemberName', 'onsiteMemberNameKey', 'onsiteCheckedInAt', 'onsiteCheckinSource',
    'lastSyncedAt', 'sourceFileName', 'sourceModifiedAt', 'rawUsageText'
  ]
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('청년이봄홈페이지 관리')
    .addItem('공개 시트 헤더 새로고침', 'refreshPublicSheetHeaders')
    .addItem('월간 운영 달력 생성', 'generateMonthlyOperationCalendar')
    .addItem('특정 월 운영 달력 생성', 'refreshSelectedOperationCalendar')
    .addItem('월별 시트 전환 실행', 'archiveAndResetMonthlySheets')
    .addItem('네이버 예약 CSV 동기화', 'syncNaverReservationCsv')
    .addToUi();
}

function doGet(e) {
  try {
    const action = str_(e && e.parameter && e.parameter.action);
    if (!action) {
      return ContentService.createTextOutput('Space management Apps Script is running.')
        .setMimeType(ContentService.MimeType.TEXT);
    }

    const payload = JSON.parse(str_(e.parameter.payload, '{}'));
    const result = dispatchWebAction_(action, payload);
    const callback = str_(e.parameter.callback);
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return json_(result);
  } catch (error) {
    return json_({ ok: false, error: error.message || String(error) });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    if (data.action) {
      return json_(dispatchWebAction_(data.action, data.payload || {}));
    }

    const sheetType = str_(data.sheetType, 'visit');
    if (sheetType === 'visit') {
      appendPublicVisitLog_({ gender: data.gender, age: data.age }, str_(data.date, today_()));
    } else if (sheetType === 'printer') {
      appendPublicPrinterLog_({ gender: data.gender, age: data.age }, { date: str_(data.date, today_()), count: Number(data.count || 0) });
    } else if (sheetType === 'career') {
      appendPublicCareerLog_(
        { gender: data.gender, age: data.age },
        {
          date: str_(data.date, today_()),
          purpose: str_(data.purpose),
          companions: [],
          startTime: str_(data.startTime),
          endTime: str_(data.endTime)
        }
      );
    } else if (sheetType === 'connect') {
      appendPublicConnectLog_(
        { gender: data.gender, age: data.age },
        {
          date: str_(data.date, today_()),
          purpose: str_(data.purpose),
          companions: [],
          startTime: str_(data.startTime),
          endTime: str_(data.endTime)
        }
      );
    }

    return json_({ ok: true });
  } catch (error) {
    return json_({ ok: false, error: error.message || String(error) });
  }
}

function dispatchWebAction_(action, payload) {
  switch (action) {
    case 'bootstrap':
      return { ok: true, adminSheetUrl: getAppSetting_('ADMIN_SHEET_URL') };
    case 'resolveFacilityEntry':
      return handleResolveFacilityEntry_(payload);
    case 'getFacilityMemberById':
      return { ok: true, member: getMemberById_(str_(payload.memberId)) };
    case 'registerFacilityMember':
      return handleRegisterFacilityMember_(payload);
    case 'getLoungeGuestById':
      return { ok: true, guest: getLoungeGuestById_(str_(payload.guestId)) };
    case 'registerLoungeGuest':
      return handleRegisterLoungeGuest_(payload);
    case 'submitVisitLog':
      return handleSubmitVisitLog_(payload);
    case 'submitUsageRecord':
      return handleSubmitUsageRecord_(payload);
    case 'searchMembers':
      return handleSearchMembers_(payload);
    case 'getStats':
      return handleGetStats_(payload);
    default:
      throw new Error('지원하지 않는 요청입니다: ' + action);
  }
}

function handleResolveFacilityEntry_(payload) {
  const rawInput = str_(payload.inputValue);
  const adminPassword = getAppSetting_('ADMIN_PASSWORD');
  if (adminPassword && rawInput === adminPassword) {
    return { ok: true, mode: 'admin', adminSheetUrl: getAppSetting_('ADMIN_SHEET_URL') };
  }

  const digits = digits_(rawInput);
  if (!digits) throw new Error('숫자 또는 관리자 비밀번호를 입력해 주세요.');

  const exact = findMembersByExactPhone_(digits);
  if (exact.length > 0) return { ok: true, mode: 'member', member: exact[0] };

  const tail = findMembersByLastDigits_(digits);
  if (tail.length === 0) return { ok: true, mode: 'signup', suggestedPhone: digits };
  if (tail.length === 1) return { ok: true, mode: 'member', member: tail[0] };

  return {
    ok: true,
    mode: 'ambiguous',
    message: '같은 끝자리 4자리를 사용하는 회원이 있습니다. 전체 전화번호를 다시 입력해 주세요.'
  };
}

function handleRegisterFacilityMember_(payload) {
  const member = normalizeFacilityMember_(payload.member || {});
  if (!member.name || !member.gender || !member.age || !member.phone) {
    throw new Error('성함, 성별, 나이, 연락처를 모두 입력해 주세요.');
  }

  const existing = findMembersByExactPhone_(member.phone);
  if (existing.length > 0) {
    return { ok: true, existing: true, member: existing[0] };
  }

  const row = {
    memberId: Utilities.getUuid(),
    name: member.name,
    gender: member.gender,
    age: member.age,
    phone: member.phone,
    phoneLastDigits: member.phoneLastDigits,
    status: 'approved',
    role: 'user',
    createdAt: nowText_(),
    updatedAt: nowText_()
  };
  appendObjectRow_(SHEET_NAMES.MEMBERS, HEADERS.members, row);
  return { ok: true, existing: false, member: memberRowToObject_(row) };
}

function handleRegisterLoungeGuest_(payload) {
  const guest = payload.guest || {};
  const birthDate = str_(guest.birthDate);
  const gender = str_(guest.gender);
  const age = Number(guest.age || calcAge_(birthDate));
  if (!birthDate || !gender || !age) {
    throw new Error('성별과 생년월일을 확인해 주세요.');
  }

  const row = {
    guestId: Utilities.getUuid(),
    name: '라운지 이용자',
    gender: gender,
    birthDate: birthDate,
    age: age,
    role: 'lounge_guest',
    validDate: today_(),
    createdAt: nowText_()
  };
  appendObjectRow_(SHEET_NAMES.LOUNGE_GUESTS, HEADERS.loungeGuests, row);
  return { ok: true, guest: loungeRowToObject_(row) };
}

function handleSubmitVisitLog_(payload) {
  runMonthlySheetRolloverIfNeeded();
  const member = normalizeSubject_(payload.member || {});
  const date = str_(payload.date, today_());
  const source = str_(payload.source, 'system');
  const createdBy = str_(payload.createdBy, 'self');
  const allowDuplicate = Boolean(payload.allowDuplicate);

  if (!member.gender || !member.age) throw new Error('방문일지 대상 정보를 찾을 수 없습니다.');
  if (!allowDuplicate && member.id && hasVisitLog_(member.id, date)) {
    return { ok: true, created: false };
  }

  appendObjectRow_(SHEET_NAMES.VISIT_LOGS, HEADERS.visitLogs, {
    logId: Utilities.getUuid(),
    userId: member.id,
    memberType: member.role,
    name: member.name,
    gender: member.gender,
    age: member.age,
    phone: member.phone,
    phoneLastDigits: member.phoneLastDigits,
    date: date,
    source: source,
    createdBy: createdBy,
    createdAt: nowText_()
  });
  appendPublicVisitLog_(member, date);
  return { ok: true, created: true };
}

function handleSubmitUsageRecord_(payload) {
  runMonthlySheetRolloverIfNeeded();
  const type = str_(payload.type);
  const member = normalizeSubject_(payload.member || {});
  const data = payload.data || {};
  const createdBy = str_(payload.createdBy, 'self');
  const date = str_(data.date, today_());

  if (type === 'printer') {
    const count = Number(data.count || 0);
    if (!count) throw new Error('프린터 사용 매수를 입력해 주세요.');
    if (member.id && getPrinterCountForDate_(member.id, date) + count > 10) {
      throw new Error('프린터는 하루 최대 10장까지 기록할 수 있습니다.');
    }

    appendObjectRow_(SHEET_NAMES.PRINTER_LOGS, HEADERS.printerLogs, {
      logId: Utilities.getUuid(),
      userId: member.id,
      name: member.name,
      gender: member.gender,
      age: member.age,
      phone: member.phone,
      phoneLastDigits: member.phoneLastDigits,
      date: date,
      count: count,
      createdBy: createdBy,
      createdAt: nowText_()
    });
    appendPublicPrinterLog_(member, { date: date, count: count });
    return { ok: true };
  }

  const roomData = {
    date: date,
    startTime: normalizeTime_(data.startTime),
    endTime: normalizeTime_(data.endTime),
    purpose: str_(data.purpose),
    companions: Array.isArray(data.companions) ? data.companions : [],
    createdBy: createdBy
  };
  if (!roomData.startTime || !roomData.endTime || !roomData.purpose) {
    throw new Error('시작 시간, 종료 시간, 이용 목적을 모두 입력해 주세요.');
  }

  const roomRow = {
    logId: Utilities.getUuid(),
    userId: member.id,
    spaceType: type,
    name: member.name,
    memberNameKey: normalizeNameKey_(member.name),
    gender: member.gender,
    age: member.age,
    phone: member.phone,
    phoneLastDigits: member.phoneLastDigits,
    date: roomData.date,
    startTime: roomData.startTime,
    endTime: roomData.endTime,
    purpose: roomData.purpose,
    companionsJson: JSON.stringify(roomData.companions),
    companionCount: roomData.companions.length,
    createdBy: roomData.createdBy,
    createdAt: nowText_(),
    linkedReservationId: '',
    dedupeStatus: 'standalone'
  };

  const matchedReservation = tryMatchNaverReservationForRoom_(type, member, roomData);
  if (matchedReservation) {
    roomRow.linkedReservationId = matchedReservation.reservationId;
    roomRow.dedupeStatus = 'matched_naver';
  }

  if (type === 'careerZone') {
    appendObjectRow_(SHEET_NAMES.CAREER_LOGS, HEADERS.roomLogs, roomRow);
    if (!matchedReservation) appendPublicCareerLog_(member, roomData);
    renderMonthlyOperationCalendar_(date);
    return { ok: true };
  }

  if (type === 'connectRoom') {
    appendObjectRow_(SHEET_NAMES.CONNECT_LOGS, HEADERS.roomLogs, roomRow);
    if (!matchedReservation) appendPublicConnectLog_(member, roomData);
    renderMonthlyOperationCalendar_(date);
    return { ok: true };
  }

  throw new Error('지원하지 않는 일지 종류입니다.');
}

function handleSearchMembers_(payload) {
  const term = str_(payload.term);
  if (!term) return { ok: true, members: [] };
  const digits = digits_(term);
  const rows = getRowsAsObjects_(SHEET_NAMES.MEMBERS, HEADERS.members).map(memberRowToObject_);
  const matches = rows.filter(function(member) {
    return member.name.indexOf(term) !== -1 ||
      (digits && (member.phone.indexOf(digits) !== -1 || member.phoneLastDigits.indexOf(digits) !== -1));
  });
  return { ok: true, members: matches.slice(0, 50) };
}

function handleGetStats_(payload) {
  const range = getPeriodRange_(str_(payload.period, 'daily'));
  const careerRows = getRowsInRange_(SHEET_NAMES.CAREER_LOGS, HEADERS.roomLogs, range);
  const connectRows = getRowsInRange_(SHEET_NAMES.CONNECT_LOGS, HEADERS.roomLogs, range);
  return {
    ok: true,
    range: range,
    totals: {
      visit: countRowsInRange_(SHEET_NAMES.VISIT_LOGS, HEADERS.visitLogs, range),
      printer: sumPrinterInRange_(range),
      careerZone: careerRows.length,
      connectRoom: connectRows.length
    },
    breakdowns: {
      careerZone: buildParticipantBreakdown_(careerRows),
      connectRoom: buildParticipantBreakdown_(connectRows)
    }
  };
}

function appendPublicVisitLog_(member, date) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.VISIT, HEADERS.visit);
  sheet.appendRow([nextSeq_(sheet), date, member.gender, member.age]);
}

function appendPublicPrinterLog_(member, data) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.PRINTER, HEADERS.printer);
  sheet.appendRow([nextSeq_(sheet), data.date, member.gender, member.age, Number(data.count || 0)]);
}

function appendPublicCareerLog_(member, data) {
  const members = [member].concat(data.companions || []);
  const stats = buildParticipantStats_(members);
  const payload = {
    date: data.date,
    purpose: data.purpose,
    headcount: members.length,
    participants: buildParticipantDetail_(members),
    startTime: data.startTime,
    endTime: data.endTime
  };
  const sheet = getOrCreateSheet_(SHEET_NAMES.CAREER, HEADERS.career);
  sheet.appendRow([
    nextSeq_(sheet),
    payload.date,
    payload.purpose,
    payload.headcount,
    payload.participants,
    payload.startTime,
    payload.endTime,
    stats.byAgeGender['20대_남성'] || 0,
    stats.byAgeGender['20대_여성'] || 0,
    stats.byAgeGender['30대_남성'] || 0,
    stats.byAgeGender['30대_여성'] || 0,
    stats.byAgeGender['19세 이하_남성'] || 0,
    stats.byAgeGender['19세 이하_여성'] || 0
  ]);
  upsertRoomCalendarEvent_('커리어존', payload);
}

function appendPublicConnectLog_(member, data) {
  const members = [member].concat(data.companions || []);
  const stats = buildParticipantStats_(members);
  const payload = {
    date: data.date,
    purpose: data.purpose,
    headcount: members.length,
    participants: buildParticipantDetail_(members),
    startTime: data.startTime,
    endTime: data.endTime
  };
  const sheet = getOrCreateSheet_(SHEET_NAMES.CONNECT, HEADERS.connect);
  sheet.appendRow([
    nextSeq_(sheet),
    payload.date,
    payload.purpose,
    payload.headcount,
    payload.participants,
    payload.startTime,
    payload.endTime,
    stats.byAgeGender['20대_남성'] || 0,
    stats.byAgeGender['20대_여성'] || 0,
    stats.byAgeGender['30대_남성'] || 0,
    stats.byAgeGender['30대_여성'] || 0,
    stats.byAgeGender['19세 이하_남성'] || 0,
    stats.byAgeGender['19세 이하_여성'] || 0
  ]);
  upsertRoomCalendarEvent_('커넥트룸', payload);
}

function tryMatchNaverReservationForRoom_(spaceType, member, roomData) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.NAVER_SYNC, HEADERS.naverSync);
  if (sheet.getLastRow() < 2) return null;

  const rows = getRowsAsObjects_(SHEET_NAMES.NAVER_SYNC, HEADERS.naverSync);
  const targetSpace = normalizeRoomSpaceType_(spaceType);
  const memberNameKey = normalizeNameKey_(member.name);

  const candidates = rows.filter(function(row) {
    if (str_(row.calendarStatus) === 'cancelled' || isCancelledNaverStatus_(row.status)) return false;
    if (normalizeRoomSpaceTypeFromNaver_(row.normalizedPlace || row.placeName) !== targetSpace) return false;
    if (normalizeSheetDate_(row.usageDate) !== roomData.date) return false;
    if (normalizeSheetTime_(row.startTime) !== roomData.startTime) return false;
    if (normalizeSheetTime_(row.endTime) !== roomData.endTime) return false;
    return true;
  });

  if (candidates.length === 0) return null;

  let matched = candidates.find(function(row) {
    return normalizeNameKey_(row.bookerName) === memberNameKey;
  });

  if (!matched && candidates.length === 1) matched = candidates[0];
  if (!matched) return null;

  updateNaverReservationCheckin_(sheet, matched.reservationId, {
    onsiteMemberName: member.name,
    onsiteMemberNameKey: memberNameKey,
    onsiteCheckedInAt: nowText_(),
    onsiteCheckinSource: spaceType
  });

  return matched;
}

function upsertRoomCalendarEvent_(spaceName, payload) {
  const calendarId = getAppSetting_('SPACE_CALENDAR_ID');
  const calendar = calendarId ? CalendarApp.getCalendarById(calendarId) : CalendarApp.getDefaultCalendar();
  if (!calendar || !payload.date || !payload.startTime || !payload.endTime) return;

  const start = buildDateTime_(payload.date, payload.startTime);
  const end = buildDateTime_(payload.date, payload.endTime);
  const title = '[' + spaceName + '] ' + payload.purpose;
  const description =
    '공간: ' + spaceName + '\n' +
    '인원수: ' + payload.headcount + '\n' +
    '참여자: ' + payload.participants + '\n' +
    '시간: ' + payload.startTime + ' ~ ' + payload.endTime;

  calendar.createEvent(title, start, end, { description: description });
}

function syncNaverReservationCsv() {
  const folderName = getAppSetting_('NAVER_CSV_FOLDER_NAME') || '네이버예약CSV';
  const calendarId = getAppSetting_('SPACE_CALENDAR_ID');
  const calendar = calendarId ? CalendarApp.getCalendarById(calendarId) : CalendarApp.getDefaultCalendar();
  const folder = getRequiredFolder_(folderName);
  const sheet = getOrCreateSheet_(SHEET_NAMES.NAVER_SYNC, HEADERS.naverSync);
  const existingMap = buildExistingReservationMap_(sheet);
  const files = folder.getFilesByType(MimeType.CSV);
  const touchedMonthKeys = {};

  let created = 0;
  let updated = 0;
  let cancelled = 0;

  while (files.hasNext()) {
    const file = files.next();
    const rows = Utilities.parseCsv(file.getBlob().getDataAsString('UTF-8').replace(/^\uFEFF/, ''));
    const headerIndex = rows.findIndex(function(row) {
      return row.join(',').indexOf('예약번호') !== -1;
    });
    if (headerIndex < 0) continue;

    const headers = rows[headerIndex];
    const dataRows = rows.slice(headerIndex + 1).filter(function(row) {
      return String(row[0] || '').trim();
    });

    dataRows.forEach(function(row) {
      const record = mapNaverRow_(headers, row, file);
      if (!record) return;

      const result = upsertNaverReservation_(sheet, calendar, existingMap[record.reservationId], record);
      if (record.usageDate) touchedMonthKeys[record.usageDate.slice(0, 7)] = true;
      if (result === 'created') created += 1;
      else if (result === 'updated') updated += 1;
      else if (result === 'cancelled') cancelled += 1;
    });

    file.setTrashed(true);
  }

  Object.keys(touchedMonthKeys).forEach(function(monthKey) {
    renderMonthlyOperationCalendarByMonthKey_(monthKey);
  });

  SpreadsheetApp.getUi().alert(
    '네이버 예약 동기화 완료\n\n' +
    '신규: ' + created + '\n' +
    '수정: ' + updated + '\n' +
    '취소/노쇼: ' + cancelled
  );
}

function archiveAndResetMonthlySheets() {
  const currentMonthKey = Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  const propertyKey = 'MONTHLY_SHEET_ROLLOVER_LAST_RUN';
  if (PropertiesService.getScriptProperties().getProperty(propertyKey) === currentMonthKey) {
    SpreadsheetApp.getUi().alert('이번 달 시트 전환은 이미 완료되었습니다.');
    return;
  }
  runMonthlySheetRolloverIfNeeded();
  SpreadsheetApp.getUi().alert('월별 시트 전환이 완료되었습니다.');
}

function runMonthlySheetRolloverIfNeeded() {
  const propertyKey = 'MONTHLY_SHEET_ROLLOVER_LAST_RUN';
  const currentMonthKey = Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(propertyKey) === currentMonthKey) return;

  const ss = getSpreadsheet_();
  const previousMonthLabel = Utilities.formatDate(new Date(new Date().getFullYear(), new Date().getMonth() - 1, 1), TZ, 'M월');
  [
    { name: SHEET_NAMES.VISIT, headers: HEADERS.visit },
    { name: SHEET_NAMES.PRINTER, headers: HEADERS.printer },
    { name: SHEET_NAMES.CAREER, headers: HEADERS.career },
    { name: SHEET_NAMES.CONNECT, headers: HEADERS.connect }
  ].forEach(function(item) {
    const current = ss.getSheetByName(item.name);
    if (current) {
      const archiveName = item.name + '_' + previousMonthLabel;
      if (!ss.getSheetByName(archiveName)) {
        current.setName(archiveName);
        current.hideSheet();
      }
    }
    const fresh = ss.getSheetByName(item.name) || ss.insertSheet(item.name);
    fresh.clear();
    ensureSheetHeaders_(fresh, item.headers);
  });

  props.setProperty(propertyKey, currentMonthKey);
}

function refreshPublicSheetHeaders() {
  ensureSheetHeaders_(getOrCreateSheet_(SHEET_NAMES.VISIT, HEADERS.visit), HEADERS.visit);
  ensureSheetHeaders_(getOrCreateSheet_(SHEET_NAMES.PRINTER, HEADERS.printer), HEADERS.printer);
  ensureSheetHeaders_(getOrCreateSheet_(SHEET_NAMES.CAREER, HEADERS.career), HEADERS.career);
  ensureSheetHeaders_(getOrCreateSheet_(SHEET_NAMES.CONNECT, HEADERS.connect), HEADERS.connect);
  SpreadsheetApp.getUi().alert('공개 시트 헤더를 최신 형식으로 다시 맞췄습니다.');
}

function generateMonthlyOperationCalendar() {
  const today = new Date();
  renderMonthlyOperationCalendarByMonthKey_(
    Utilities.formatDate(today, TZ, 'yyyy-MM')
  );
  SpreadsheetApp.getUi().alert('이번 달 운영 달력을 생성했습니다.');
}

function renderMonthlyOperationCalendar_(dateStr) {
  const monthKey = str_(dateStr).slice(0, 7);
  if (!monthKey) return;
  renderMonthlyOperationCalendarByMonthKey_(monthKey);
}

function renderMonthlyOperationCalendarByMonthKey_(monthKey) {
  const parts = monthKey.split('-').map(Number);
  if (parts.length < 2 || !parts[0] || !parts[1]) return;

  const year = parts[0];
  const month = parts[1] - 1;
  const sheetName = monthKey;
  const ss = getSpreadsheet_();
  const sheet = createOperationCalendarSheet_(ss, sheetName);
  const calendarId = getAppSetting_('SPACE_CALENDAR_ID');
  const calendar = calendarId ? CalendarApp.getCalendarById(calendarId) : CalendarApp.getDefaultCalendar();
  createOperationCalendarView_(sheet, calendar, year, month);
}

function refreshSelectedOperationCalendar() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('달력을 표시할 월을 입력하세요.\n형식: YYYY-MM\n예: 2026-06');
  if (!response || response.getSelectedButton() === ui.Button.CANCEL) return;

  const input = str_(response.getResponseText());
  if (!/^\d{4}-\d{2}$/.test(input)) {
    ui.alert('형식이 잘못되었습니다. YYYY-MM 형식으로 입력해 주세요.');
    return;
  }

  renderMonthlyOperationCalendarByMonthKey_(input);
  ui.alert(input + ' 운영 달력을 생성했습니다.');
}

function createOperationCalendarSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  return sheet;
}

function createOperationCalendarView_(sheet, calendar, year, month) {
  if (!sheet || !calendar || !year || month === undefined) return;

  const titleRange = sheet.getRange(1, 1, 1, 7);
  titleRange.merge();
  titleRange.setValue(year + '년 ' + (month + 1) + '월')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF');

  const days = ['일', '월', '화', '수', '목', '금', '토'];
  const headerRow = 3;

  sheet.setHiddenGridlines(true);

  for (var i = 0; i < 7; i += 1) {
    const cell = sheet.getRange(headerRow, i + 1);
    cell.setValue(days[i])
      .setFontWeight('bold')
      .setBackground('#E8F0FE')
      .setHorizontalAlignment('center');

    if (i === 0) cell.setFontColor('#FF0000');
    else if (i === 6) cell.setFontColor('#0000FF');
  }

  const startOfMonth = new Date(year, month, 1);
  const endOfMonth = new Date(year, month + 1, 0, 23, 59, 59);
  let events = [];

  try {
    events = calendar.getEvents(startOfMonth, endOfMonth);
  } catch (error) {
    Logger.log('운영 달력 캘린더 조회 오류: ' + (error.message || error));
  }

  const eventMap = {};
  events.forEach(function(event) {
    if (!isOperationCalendarEvent_(event)) return;

    const start = event.getStartTime();
    const key = start.getFullYear() + '-' + (start.getMonth() + 1) + '-' + start.getDate();
    if (!eventMap[key]) eventMap[key] = [];

    const startStr = pad2_(start.getHours()) + ':' + pad2_(start.getMinutes());
    const endTime = event.getEndTime();
    const endStr = pad2_(endTime.getHours()) + ':' + pad2_(endTime.getMinutes());
    const title = event.getTitle() || '';
    const isCancelled = title.indexOf('[취소]') !== -1 || title.indexOf('[노쇼]') !== -1;

    eventMap[key].push({
      time: startStr + '~' + endStr,
      title: title,
      description: event.getDescription() || '',
      isCancelled: isCancelled
    });
  });

  let currentRow = 4;
  const firstDay = new Date(year, month, 1).getDay();
  const lastDay = new Date(year, month + 1, 0).getDate();
  let currentDay = 1 - firstDay;

  for (var week = 0; week < 6; week += 1) {
    if (currentDay > lastDay) break;

    const weekEventCounts = [];
    for (var dayIndex = 0; dayIndex < 7; dayIndex += 1) {
      const tempDay = currentDay + dayIndex;
      if (tempDay > 0 && tempDay <= lastDay) {
        const tempKey = year + '-' + (month + 1) + '-' + tempDay;
        weekEventCounts.push((eventMap[tempKey] || []).length);
      } else {
        weekEventCounts.push(0);
      }
    }

    const maxEventsInWeek = Math.max(4, weekEventCounts[0], weekEventCounts[1], weekEventCounts[2], weekEventCounts[3], weekEventCounts[4], weekEventCounts[5], weekEventCounts[6]);

    for (var dayOfWeek = 0; dayOfWeek < 7; dayOfWeek += 1) {
      const col = dayOfWeek + 1;
      const virtualDay = currentDay + dayOfWeek;

      if (virtualDay <= 0 || virtualDay > lastDay) {
        sheet.getRange(currentRow, col, 1 + maxEventsInWeek, 1)
          .setBackground('#F9F9F9')
          .setBorder(true, true, true, true, null, null);
      } else {
        const dateCell = sheet.getRange(currentRow, col);
        dateCell.setValue(virtualDay)
          .setFontSize(12)
          .setFontWeight('bold')
          .setHorizontalAlignment('left')
          .setBackground('#FFFFFF')
          .setBorder(true, true, true, true, null, null);

        if (dayOfWeek === 0) dateCell.setFontColor('#FF0000');
        else if (dayOfWeek === 6) dateCell.setFontColor('#0000FF');

        const dayKey = year + '-' + (month + 1) + '-' + virtualDay;
        const dayEvents = eventMap[dayKey] || [];

        dayEvents.forEach(function(item, idx) {
          const eventRow = currentRow + 1 + idx;
          const eventCell = sheet.getRange(eventRow, col);
          eventCell.setValue(item.time + ' ' + item.title)
            .setBackground(item.isCancelled ? '#F0F0F0' : '#FFFACD')
            .setFontSize(9)
            .setWrap(true)
            .setVerticalAlignment('middle')
            .setBorder(true, true, true, true, null, null);

          if (item.isCancelled) {
            eventCell.setFontColor('#999999').setFontLine('line-through');
          }

          if (item.description) {
            eventCell.setNote(item.description);
          }
        });

        if (dayEvents.length < maxEventsInWeek) {
          sheet.getRange(currentRow + 1 + dayEvents.length, col, maxEventsInWeek - dayEvents.length, 1)
            .setBackground('#FFFFFF')
            .setBorder(true, true, true, true, null, null);
        }
      }
    }

    sheet.setRowHeight(currentRow, 25);
    for (var rowOffset = 1; rowOffset <= maxEventsInWeek; rowOffset += 1) {
      sheet.setRowHeight(currentRow + rowOffset, 45);
    }

    currentRow += 1 + maxEventsInWeek;
    currentDay += 7;
  }

  for (var colIndex = 1; colIndex <= 7; colIndex += 1) {
    sheet.setColumnWidth(colIndex, 150);
  }
}

function isOperationCalendarEvent_(event) {
  const title = str_(event.getTitle());
  return title.indexOf('커리어존') !== -1 || title.indexOf('커넥트룸') !== -1 || title.indexOf('AI 커리어존') !== -1 || title.indexOf('청년 커넥트룸') !== -1;
}

function getAppSetting_(key) {
  const prop = PropertiesService.getScriptProperties().getProperty(key);
  if (prop) return String(prop).trim();

  const rows = getRowsAsObjects_(SHEET_NAMES.SETTINGS, HEADERS.settings);
  const found = rows.find(function(row) { return str_(row.key) === key; });
  return found ? str_(found.value) : '';
}

function getMemberById_(memberId) {
  const row = getRowsAsObjects_(SHEET_NAMES.MEMBERS, HEADERS.members).find(function(item) {
    return str_(item.memberId) === memberId;
  });
  return row ? memberRowToObject_(row) : null;
}

function getLoungeGuestById_(guestId) {
  const row = getRowsAsObjects_(SHEET_NAMES.LOUNGE_GUESTS, HEADERS.loungeGuests).find(function(item) {
    return str_(item.guestId) === guestId;
  });
  return row ? loungeRowToObject_(row) : null;
}

function findMembersByExactPhone_(phone) {
  return getRowsAsObjects_(SHEET_NAMES.MEMBERS, HEADERS.members)
    .map(memberRowToObject_)
    .filter(function(row) { return row.phone === phone; });
}

function findMembersByLastDigits_(digits) {
  return getRowsAsObjects_(SHEET_NAMES.MEMBERS, HEADERS.members)
    .map(memberRowToObject_)
    .filter(function(row) { return row.phoneLastDigits === digits; });
}

function hasVisitLog_(userId, date) {
  return getRowsAsObjects_(SHEET_NAMES.VISIT_LOGS, HEADERS.visitLogs).some(function(row) {
    return str_(row.userId) === userId && str_(row.date) === date;
  });
}

function getPrinterCountForDate_(userId, date) {
  return getRowsAsObjects_(SHEET_NAMES.PRINTER_LOGS, HEADERS.printerLogs).reduce(function(sum, row) {
    if (str_(row.userId) !== userId || str_(row.date) !== date) return sum;
    return sum + Number(row.count || 0);
  }, 0);
}

function countRowsInRange_(sheetName, headers, range) {
  return getRowsAsObjects_(sheetName, headers).filter(function(row) {
    const date = normalizeSheetDate_(row.date);
    return date >= range.start && date <= range.end;
  }).length;
}

function getRowsInRange_(sheetName, headers, range) {
  return getRowsAsObjects_(sheetName, headers).filter(function(row) {
    const date = normalizeSheetDate_(row.date);
    return date >= range.start && date <= range.end;
  });
}

function sumPrinterInRange_(range) {
  return getRowsAsObjects_(SHEET_NAMES.PRINTER_LOGS, HEADERS.printerLogs).reduce(function(sum, row) {
    const date = normalizeSheetDate_(row.date);
    if (date < range.start || date > range.end) return sum;
    return sum + Number(row.count || 0);
  }, 0);
}

function buildParticipantBreakdown_(rows) {
  const breakdown = {
    totalParticipants: 0,
    byGender: {},
    byAgeGroup: {},
    byAgeGender: {}
  };

  (rows || []).forEach(function(row) {
    const participants = extractParticipantsFromRoomRow_(row);
    participants.forEach(function(person) {
      const gender = str_(person.gender, '미기재');
      const age = Number(person.age || 0);
      const ageGroup = getAgeGroupLabel_(age);
      const ageGenderKey = ageGroup + '_' + gender;

      breakdown.totalParticipants += 1;
      breakdown.byGender[gender] = (breakdown.byGender[gender] || 0) + 1;
      breakdown.byAgeGroup[ageGroup] = (breakdown.byAgeGroup[ageGroup] || 0) + 1;
      breakdown.byAgeGender[ageGenderKey] = (breakdown.byAgeGender[ageGenderKey] || 0) + 1;
    });
  });

  return breakdown;
}

function buildParticipantStats_(members) {
  const stats = {
    byGender: {},
    byAgeGroup: {},
    byAgeGender: {}
  };

  (members || []).forEach(function(person) {
    const gender = str_(person.gender, '미기재');
    const age = Number(person.age || 0);
    const ageGroup = getAgeGroupLabel_(age);
    const ageGenderKey = ageGroup + '_' + gender;

    stats.byGender[gender] = (stats.byGender[gender] || 0) + 1;
    stats.byAgeGroup[ageGroup] = (stats.byAgeGroup[ageGroup] || 0) + 1;
    stats.byAgeGender[ageGenderKey] = (stats.byAgeGender[ageGenderKey] || 0) + 1;
  });

  return stats;
}

function extractParticipantsFromRoomRow_(row) {
  const participants = [{
    name: str_(row.name),
    gender: str_(row.gender),
    age: Number(row.age || 0)
  }];

  try {
    const companions = JSON.parse(str_(row.companionsJson, '[]'));
    if (Array.isArray(companions)) {
      companions.forEach(function(companion) {
        participants.push({
          name: '',
          gender: str_(companion.gender),
          age: Number(companion.age || 0)
        });
      });
    }
  } catch (error) {}

  return participants.filter(function(person) {
    return person.gender || person.age;
  });
}

function getAgeGroupLabel_(age) {
  if (!age) return '미기재';
  if (age <= 19) return '19세 이하';
  if (age <= 29) return '20대';
  if (age <= 39) return '30대';
  if (age <= 49) return '40대';
  return '50세 이상';
}

function getPeriodRange_(period) {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  if (period === 'weekly') {
    const day = start.getDay();
    start.setDate(start.getDate() - (day === 0 ? 6 : day - 1));
  } else if (period === 'monthly') {
    start.setDate(1);
  }
  return {
    start: Utilities.formatDate(start, TZ, 'yyyy-MM-dd'),
    end: Utilities.formatDate(now, TZ, 'yyyy-MM-dd')
  };
}

function getOrCreateSheet_(name, headers) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  ensureSheetHeaders_(sheet, headers);
  return sheet;
}

function getSpreadsheet_() {
  const spreadsheetId = getAppSetting_('SPREADSHEET_ID');
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('SPREADSHEET_ID 스크립트 속성을 설정해 주세요.');
  }
  return active;
}

function ensureSheetHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const same = headers.every(function(header, index) { return String(current[index] || '') === header; });
    if (!same) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }

  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
}

function getRowsAsObjects_(sheetName, headers) {
  const sheet = getOrCreateSheet_(sheetName, headers);
  if (sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  return values.map(function(row) {
    const obj = {};
    headers.forEach(function(header, index) { obj[header] = row[index]; });
    return obj;
  });
}

function appendObjectRow_(sheetName, headers, obj) {
  const sheet = getOrCreateSheet_(sheetName, headers);
  sheet.appendRow(headers.map(function(header) { return obj[header] !== undefined ? obj[header] : ''; }));
}

function nextSeq_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  return Number(sheet.getRange(lastRow, 1).getValue() || 0) + 1;
}

function normalizeFacilityMember_(member) {
  const phone = digits_(member.phone);
  return {
    name: str_(member.name),
    gender: str_(member.gender),
    age: Number(member.age || 0),
    phone: phone,
    phoneLastDigits: phone.slice(-4)
  };
}

function normalizeSubject_(member) {
  const phone = digits_(member.phone);
  return {
    id: str_(member.id || member.memberId || member.guestId),
    name: str_(member.name),
    gender: str_(member.gender),
    age: Number(member.age || 0),
    phone: phone,
    phoneLastDigits: str_(member.phoneLastDigits || phone.slice(-4)),
    role: str_(member.role, 'user')
  };
}

function memberRowToObject_(row) {
  return {
    id: str_(row.memberId),
    name: str_(row.name),
    gender: str_(row.gender),
    age: Number(row.age || 0),
    phone: digits_(row.phone),
    phoneLastDigits: str_(row.phoneLastDigits),
    status: str_(row.status, 'approved'),
    role: str_(row.role, 'user')
  };
}

function loungeRowToObject_(row) {
  return {
    id: str_(row.guestId),
    name: str_(row.name, '라운지 이용자'),
    gender: str_(row.gender),
    birthDate: str_(row.birthDate),
    age: Number(row.age || 0),
    role: str_(row.role, 'lounge_guest'),
    validDate: str_(row.validDate)
  };
}

function buildParticipantDetail_(members) {
  return (members || []).map(function(item) {
    return str_(item.gender) + '/' + String(item.age || '');
  }).join(', ');
}

function mapNaverRow_(headers, row, file) {
  const map = {};
  headers.forEach(function(header, index) { map[str_(header)] = str_(row[index]); });
  const usage = parseNaverUsageDateTime_(pickFirstValue_(map, ['이용일시']));
  const status = pickFirstValue_(map, ['상태']);
  const placeName = pickFirstValue_(map, ['상품']);
  const normalizedPlace = normalizeNaverPlaceName_(placeName);
  const normalizedSpaceType = normalizeRoomSpaceTypeFromNaver_(normalizedPlace);

  if (!isSupportedNaverReservationSpace_(normalizedSpaceType)) return null;

  return {
    reservationId: pickFirstValue_(map, ['예약번호']),
    status: status,
    bookerName: pickFirstValue_(map, ['예약자', '예약자명']),
    maskedPhone: pickFirstValue_(map, ['전화번호', '휴대폰번호']),
    usageDate: usage.date,
    startTime: usage.startTime,
    endTime: usage.endTime,
    placeName: placeName,
    normalizedPlace: normalizedPlace,
    headcount: pickFirstValue_(map, ['예약자입력정보1-이용 인원(명)', '예약자입력정보1_이용 인원(명)', '이용 인원(명)']),
    purpose: pickFirstValue_(map, ['예약자입력정보2-이용 목적', '예약자입력정보2_이용 목적', '이용 목적']),
    requestedAt: pickFirstValue_(map, ['예약신청일시']),
    confirmedAt: pickFirstValue_(map, ['예약확정일시']),
    completedAt: pickFirstValue_(map, ['이용완료일시']),
    cancelledAt: pickFirstValue_(map, ['예약취소일시']),
    cancelReason: pickFirstValue_(map, ['취소사유']),
    calendarEventId: '',
    calendarStatus: isCancelledNaverStatus_(status) ? 'cancelled' : 'active',
    lastSyncedAt: nowText_(),
    sourceFileName: file.getName(),
    sourceModifiedAt: Utilities.formatDate(file.getLastUpdated(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    rawUsageText: pickFirstValue_(map, ['이용일시'])
  };
}

function buildExistingReservationMap_(sheet) {
  const headers = HEADERS.naverSync;
  if (sheet.getLastRow() < 2) return {};
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  const map = {};
  values.forEach(function(row, index) {
    const obj = {};
    headers.forEach(function(header, col) { obj[header] = row[col]; });
    map[str_(obj.reservationId)] = { rowNumber: index + 2, values: obj };
  });
  return map;
}

function updateNaverReservationCheckin_(sheet, reservationId, values) {
  if (sheet.getLastRow() < 2) return;

  const headers = HEADERS.naverSync;
  const allRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  const reservationIdKey = str_(reservationId);
  const rowIndex = allRows.findIndex(function(row) {
    return str_(row[0]) === reservationIdKey;
  });
  if (rowIndex < 0) return;

  const currentRow = allRows[rowIndex];
  const objectRow = {};
  headers.forEach(function(header, index) {
    objectRow[header] = currentRow[index];
  });

  objectRow.onsiteMemberName = values.onsiteMemberName || '';
  objectRow.onsiteMemberNameKey = values.onsiteMemberNameKey || '';
  objectRow.onsiteCheckedInAt = values.onsiteCheckedInAt || '';
  objectRow.onsiteCheckinSource = values.onsiteCheckinSource || '';
  objectRow.lastSyncedAt = nowText_();

  const updatedRow = headers.map(function(header) {
    return objectRow[header] !== undefined ? objectRow[header] : '';
  });

  sheet.getRange(rowIndex + 2, 1, 1, updatedRow.length).setValues([updatedRow]);
  applyNaverReservationRowStyle_(sheet, rowIndex + 2, objectRow);
}

function upsertNaverReservation_(sheet, calendar, existing, record) {
  if (existing && existing.values && existing.values.calendarEventId) {
    record.calendarEventId = existing.values.calendarEventId;
  }
  if (existing && existing.values) {
    record.onsiteMemberName = str_(existing.values.onsiteMemberName);
    record.onsiteMemberNameKey = str_(existing.values.onsiteMemberNameKey);
    record.onsiteCheckedInAt = str_(existing.values.onsiteCheckedInAt);
    record.onsiteCheckinSource = str_(existing.values.onsiteCheckinSource);
  }

  record.calendarEventId = upsertNaverCalendarEvent_(calendar, record);
  const row = HEADERS.naverSync.map(function(header) { return record[header] || ''; });

  if (!existing) {
    sheet.appendRow(row);
    applyNaverReservationRowStyle_(sheet, sheet.getLastRow(), record);
    return record.calendarStatus === 'cancelled' ? 'cancelled' : 'created';
  }

  sheet.getRange(existing.rowNumber, 1, 1, row.length).setValues([row]);
  applyNaverReservationRowStyle_(sheet, existing.rowNumber, record);
  return record.calendarStatus === 'cancelled' ? 'cancelled' : 'updated';
}

function upsertNaverCalendarEvent_(calendar, record) {
  if (!record.usageDate || !record.startTime || !record.endTime || !record.normalizedPlace) {
    return record.calendarEventId || '';
  }

  const start = buildDateTime_(record.usageDate, record.startTime);
  const end = buildDateTime_(record.usageDate, record.endTime);
  const title = (record.calendarStatus === 'cancelled' ? '[취소] ' : '[예약] ') + record.normalizedPlace + ' / ' + record.bookerName;
  const description =
    '예약번호: ' + record.reservationId + '\n' +
    '상태: ' + record.status + '\n' +
    '예약자: ' + record.bookerName + '\n' +
    '전화번호: ' + record.maskedPhone + '\n' +
    '이용일시: ' + record.rawUsageText + '\n' +
    '인원수: ' + record.headcount + '\n' +
    '이용목적: ' + record.purpose + '\n' +
    '취소사유: ' + record.cancelReason;

  let event = null;
  if (record.calendarEventId) {
    try { event = calendar.getEventById(record.calendarEventId); } catch (e) {}
  }
  if (!event) {
    event = calendar.createEvent(title, start, end, { description: description });
  } else {
    event.setTitle(title);
    event.setTime(start, end);
    event.setDescription(description);
  }
  return event.getId();
}

function parseNaverUsageDateTime_(text) {
  const raw = str_(text);
  const dateMatch = raw.match(/(\d{2,4})\.\s*(\d{1,2})\.\s*(\d{1,2})/);
  const timeMatch = raw.match(/(오전|오후)\s*(\d{1,2}):(\d{2})~(\d{1,2}):(\d{2})/);
  if (!dateMatch || !timeMatch) return { date: '', startTime: '', endTime: '' };

  let year = Number(dateMatch[1]);
  if (year < 100) year += 2000;
  const month = pad2_(dateMatch[2]);
  const day = pad2_(dateMatch[3]);
  const startHour = convertHour_(timeMatch[1], Number(timeMatch[2]));
  let endHour = convertHour_(timeMatch[1], Number(timeMatch[4]));
  const startMinute = Number(timeMatch[3]);
  const endMinute = Number(timeMatch[5]);
  if (endHour < startHour || (endHour === startHour && endMinute <= startMinute)) endHour += 12;

  return {
    date: year + '-' + month + '-' + day,
    startTime: pad2_(startHour) + ':' + pad2_(startMinute),
    endTime: pad2_(endHour) + ':' + pad2_(endMinute)
  };
}

function normalizeNaverPlaceName_(placeName) {
  return str_(placeName)
    .replace('[동아리실] ', '')
    .replace('[Club room] ', '')
    .replace('3층 ', '')
    .replace('3rd floor ', '')
    .trim();
}

function normalizeRoomSpaceType_(spaceType) {
  if (spaceType === 'careerZone') return 'career';
  if (spaceType === 'connectRoom') return 'connect';
  return str_(spaceType);
}

function normalizeRoomSpaceTypeFromNaver_(placeName) {
  const normalized = normalizeNameKey_(normalizeNaverPlaceName_(placeName));
  if (normalized.indexOf('커리어') !== -1 || normalized.indexOf('career') !== -1) return 'career';
  if (normalized.indexOf('커넥트') !== -1 || normalized.indexOf('connect') !== -1) return 'connect';
  return normalized;
}

function isSupportedNaverReservationSpace_(spaceType) {
  return spaceType === 'career' || spaceType === 'connect';
}

function isCancelledNaverStatus_(status) {
  const normalized = normalizeNameKey_(status);
  return normalized.indexOf('취소') !== -1 || normalized.indexOf('노쇼') !== -1 || normalized.indexOf('noshow') !== -1;
}

function applyNaverReservationRowStyle_(sheet, rowNumber, record) {
  if (!sheet || !rowNumber) return;
  const isCancelled = isCancelledNaverStatus_(record.status) || str_(record.calendarStatus) === 'cancelled';
  const range = sheet.getRange(rowNumber, 1, 1, HEADERS.naverSync.length);
  range.setFontLine(isCancelled ? 'line-through' : 'none');
  range.setFontColor(isCancelled ? '#777777' : '#000000');
}

function getRequiredFolder_(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) throw new Error('드라이브 폴더를 찾을 수 없습니다: ' + folderName);
  return folders.next();
}

function buildDateTime_(dateStr, timeStr) {
  const d = dateStr.split('-').map(Number);
  const t = timeStr.split(':').map(Number);
  return new Date(d[0], d[1] - 1, d[2], t[0], t[1], 0);
}

function normalizeSheetDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  }
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, TZ, 'yyyy-MM-dd');
  }
  return text;
}

function normalizeSheetTime_(value) {
  if (!value && value !== 0) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'HH:mm');
  }
  const text = String(value).trim();
  if (/^\d{2}:\d{2}$/.test(text)) return text;

  const parsed = new Date('1970-01-01T' + text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, TZ, 'HH:mm');
  }
  return normalizeTime_(text);
}

function normalizeTime_(value) {
  const d = digits_(value);
  if (!d) return '';
  const padded = d.padStart(4, '0').slice(0, 4);
  return pad2_(Math.min(Number(padded.substring(0, 2)), 23)) + ':' + pad2_(Math.min(Number(padded.substring(2, 4)), 59));
}

function calcAge_(birthDate) {
  const birth = new Date(birthDate);
  if (isNaN(birth.getTime())) return 0;
  const now = new Date();
  let age = now.getFullYear() - birth.getFullYear();
  const monthDiff = now.getMonth() - birth.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && now.getDate() < birth.getDate())) age -= 1;
  return age;
}

function pickFirstValue_(map, keys) {
  for (let i = 0; i < keys.length; i += 1) {
    const value = str_(map[keys[i]]);
    if (value) return value;
  }
  return '';
}

function json_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function str_(value, fallback) {
  const result = value === undefined || value === null ? '' : String(value).trim();
  return result || (fallback || '');
}

function digits_(value) {
  return str_(value).replace(/\D/g, '');
}

function normalizeNameKey_(value) {
  return str_(value).replace(/\s+/g, '').toLowerCase();
}

function pad2_(value) {
  return String(value).padStart(2, '0');
}

function convertHour_(period, hour) {
  if (period === '오후' && hour !== 12) return hour + 12;
  if (period === '오전' && hour === 12) return 0;
  return hour;
}

function today_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

function nowText_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}
