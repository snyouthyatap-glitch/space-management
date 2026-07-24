const TZ = 'Asia/Seoul';
const QR_ENTRY_TOKEN = 'ybom-qr-2026';
const MEMBER_RETENTION_YEARS = 3;

const SHEET_NAMES = {
  VISIT: '방문일지',
  PRINTER: '프린터',
  CAREER: '커리어존',
  CONNECT: '커넥트룸',
  MEMBERS: 'APP_MEMBERS',
  VISIT_LOGS: 'APP_VISIT_LOGS',
  PRINTER_LOGS: 'APP_PRINTER_LOGS',
  CAREER_LOGS: 'APP_CAREER_LOGS',
  CONNECT_LOGS: 'APP_CONNECT_LOGS',
  RETENTION_LOGS: 'APP_RETENTION_LOGS',
  NAVER_SYNC: '네이버예약_동기화'
};

const HEADERS = {
  visit: ['연번', '날짜', '성별', '나이'],
  printer: ['연번', '날짜', '성별', '나이', '사용 매수'],
  career: [
    '연번', '날짜', '사용목적', '인원수', '비고', '시작시간', '종료시간',
    '20~29세(남)', '20~29세(여)', '30~39세(남)', '30~39세(여)', '~19세(남)', '~19세(여)',
    '40세 이상(남)', '40세 이상(여)'
  ],
  connect: [
    '연번', '날짜', '사용목적', '인원수', '비고', '시작시간', '종료시간',
    '20~29세(남)', '20~29세(여)', '30~39세(남)', '30~39세(여)', '~19세(남)', '~19세(여)',
    '40세 이상(남)', '40세 이상(여)'
  ],
  members: [
    'memberId', 'name', 'gender', 'birthDate', 'age', 'phone', 'phoneLastDigits', 'isSeongnamResident',
    'privacyConsent', 'optionalConsent', 'consentAt', 'status', 'role', 'createdAt', 'updatedAt',
    'lastUsedAt', 'retentionHoldUntil', 'retentionHoldReason'
  ],
  visitLogs: ['logId', 'userId', 'memberType', 'name', 'gender', 'age', 'phone', 'phoneLastDigits', 'date', 'source', 'createdBy', 'createdAt'],
  printerLogs: ['logId', 'userId', 'name', 'gender', 'age', 'phone', 'phoneLastDigits', 'date', 'count', 'createdBy', 'createdAt', 'requestId'],
  roomLogs: [
    'logId', 'userId', 'spaceType', 'name', 'memberNameKey', 'gender', 'age', 'phone', 'phoneLastDigits',
    'date', 'startTime', 'endTime', 'purpose', 'companionsJson', 'companionCount', 'createdBy', 'createdAt',
    'linkedReservationId', 'dedupeStatus', 'headcount', 'participants', 'male20s', 'female20s', 'male30s',
    'female30s', 'maleUnder19', 'femaleUnder19', 'male40Plus', 'female40Plus'
  ],
  retentionLogs: [
    'runId', 'action', 'processedAt', 'memberCount', 'visitLogCount', 'printerLogCount',
    'careerLogCount', 'connectLogCount', 'reservationLinkCount'
  ],
  naverSync: [
    'reservationId', 'status', 'bookerName', 'maskedPhone', 'usageDate', 'startTime', 'endTime',
    'placeName', 'spaceType', 'headcount', 'purpose', 'requestedAt', 'confirmedAt', 'completedAt',
    'cancelledAt', 'cancelReason', 'onsiteMemberName', 'onsiteMemberNameKey', 'onsiteCheckedInAt',
    'onsiteCheckinSource', 'usageLogId',
    'submittedAt', 'submittedByMemberId', 'lastSyncedAt', 'sourceFileName', 'sourceModifiedAt',
    'rawUsageText'
  ]
};

const NAVER_CSV_FIELDS = {
  reservationId: ['예약번호'],
  status: ['상태', '예약상태'],
  bookerName: ['예약자', '예약자명', '방문자명'],
  phone: ['전화번호', '휴대폰번호', '방문자전화번호'],
  usageDateTime: ['이용일시', '이용 일시'],
  placeName: ['상품', '상품명', '예약상품'],
  headcount: [
    '예약자입력정보2-이용 인원(명)',
    '예약자입력정보2_이용 인원(명)',
    '예약자입력정보1-이용 인원(명)',
    '예약자입력정보1_이용 인원(명)',
    '이용 인원(명)',
    '수량'
  ],
  purpose: [
    '예약자입력정보5-이용 목적',
    '예약자입력정보5_이용 목적',
    '예약자입력정보2-이용 목적',
    '예약자입력정보2_이용 목적',
    '이용 목적',
    '요청사항'
  ],
  requestedAt: ['예약신청일시'],
  confirmedAt: ['예약확정일시'],
  completedAt: ['이용완료일시'],
  cancelledAt: ['예약취소일시'],
  cancelReason: ['취소사유']
};

const NAVER_CSV_REQUIRED_HEADERS = ['예약번호', '상태', '예약자', '전화번호', '이용일시', '상품'];

function onOpen() {
  ensureDataSheetSchemas_();
  runMonthlySheetRolloverIfNeeded();
  ensurePublicFormulaSheetSchemas_();
  SpreadsheetApp.getUi()
    .createMenu('청년이봄홈페이지 관리')
    .addItem('월별 시트 전환 실행', 'archiveAndResetMonthlySheets')
    .addItem('간소화 시트 수식 재설정', 'setupPublicFormulaSheets')
    .addItem('네이버 예약 CSV 동기화', 'syncNaverReservationCsv')
    .addSeparator()
    .addItem('마지막 이용일 다시 계산', 'backfillMemberLastUsedAt')
    .addItem('3년 초과 회원 미리보기', 'previewExpiredMembers')
    .addItem('3년 초과 개인정보 파기', 'purgeExpiredMemberData')
    .addItem('월간 개인정보 파기 트리거 설정', 'setupRetentionCleanupTrigger')
    .addToUi();
}

function doGet(e) {
  const callback = str_(e && e.parameter && e.parameter.callback);
  try {
    const action = str_(e && e.parameter && e.parameter.action);
    if (!action) {
      return ContentService.createTextOutput('Space management Apps Script is running.')
        .setMimeType(ContentService.MimeType.TEXT);
    }

    const payload = JSON.parse(str_(e.parameter.payload, '{}'));
    const result = dispatchWebAction_(action, payload);
    return webResponse_(result, callback);
  } catch (error) {
    return webResponse_({ ok: false, error: error.message || String(error) }, callback);
  }
}

function dispatchWebAction_(action, payload) {
  requireEntryToken_(payload);
  switch (action) {
    case 'resolveFacilityEntry':
      return handleResolveFacilityEntry_(payload);
    case 'updateMemberBirthDate':
      return handleUpdateMemberBirthDate_(payload);
    case 'registerFacilityMember':
      return handleRegisterFacilityMember_(payload);
    case 'submitVisitLog':
      return handleSubmitVisitLog_(payload);
    case 'submitUsageRecord':
      return handleSubmitUsageRecord_(payload);
    case 'submitReservationUsage':
      return handleSubmitReservationUsage_(payload);
    default:
      throw new Error('지원하지 않는 요청입니다: ' + action);
  }
}

function handleResolveFacilityEntry_(payload) {
  const rawInput = str_(payload.inputValue);
  const digits = digits_(rawInput);
  if (digits.length !== 4 && !isValidPhone_(digits)) {
    throw new Error('전화번호 뒷 4자리 또는 전체 휴대폰 번호를 입력해 주세요.');
  }

  if (digits.length === 11) {
    const exact = findMembersByField_('phone', digits);
    if (exact.length === 0) return { ok: true, mode: 'signup', suggestedPhone: digits };
    const activeExact = exact.filter(isActiveMember_);
    if (activeExact.length === 0) throw new Error('현재 이용할 수 없는 회원입니다.');
    return { ok: true, mode: 'member', member: publicMember_(activeExact[0]) };
  }

  const tail = findMembersByField_('phoneLastDigits', digits).filter(isActiveMember_);
  if (tail.length === 0) return { ok: true, mode: 'signup', suggestedPhone: digits };
  if (tail.length === 1) return { ok: true, mode: 'member', member: publicMember_(tail[0]) };

  return {
    ok: true,
    mode: 'ambiguous',
    message: '같은 끝자리 4자리를 사용하는 회원이 있습니다. 전체 전화번호를 다시 입력해 주세요.'
  };
}

function handleRegisterFacilityMember_(payload) {
  const member = normalizeFacilityMember_(payload.member || {});
  if (!member.name || !isValidGender_(member.gender) || !isValidBirthDate_(member.birthDate) || !isValidPhone_(member.phone) || !member.privacyConsent) {
    throw new Error('필수 정보와 개인정보 수집·이용 동의를 모두 확인해 주세요.');
  }

  return withScriptLock_(function() {
    const existing = findMembersByField_('phone', member.phone);
    if (existing.length > 0) {
      const activeExisting = existing.filter(isActiveMember_);
      if (activeExisting.length === 0) throw new Error('현재 이용할 수 없는 회원입니다.');
      return { ok: true, existing: true, member: publicMember_(activeExisting[0]) };
    }

    const row = {
      memberId: Utilities.getUuid(),
      name: member.name,
      gender: member.gender,
      age: '',
      phone: member.phone,
      phoneLastDigits: member.phoneLastDigits,
      status: 'approved',
      role: 'user',
      createdAt: nowText_(),
      updatedAt: nowText_(),
      birthDate: member.birthDate,
      isSeongnamResident: member.isSeongnamResident,
      privacyConsent: true,
      optionalConsent: false,
      consentAt: nowText_(),
      lastUsedAt: today_(),
      retentionHoldUntil: '',
      retentionHoldReason: ''
    };
    appendObjectRow_(SHEET_NAMES.MEMBERS, HEADERS.members, row);
    return { ok: true, existing: false, member: publicMember_(memberRowToObject_(row)) };
  });
}

function requireEntryToken_(payload) {
  if (str_(payload && payload.entryToken) !== QR_ENTRY_TOKEN) {
    throw new Error('현장 QR 코드로 다시 접속해 주세요.');
  }
}

function handleUpdateMemberBirthDate_(payload) {
  const memberId = str_(payload.memberId);
  const birthDate = normalizeBirthDate_(payload.birthDate);
  const source = normalizeVisitSource_(payload.source || 'facility-birthdate-update');
  if (!memberId || !isValidBirthDate_(birthDate)) {
    throw new Error('회원 정보와 생년월일을 확인해 주세요.');
  }

  return withScriptLock_(function() {
    const sheet = getOrCreateSheet_(SHEET_NAMES.MEMBERS, HEADERS.members);
    ensureSheetHeaders_(sheet, HEADERS.members);
    const rowNumber = findFirstRowByColumnValue_(sheet, memberColumn_(sheet, 'memberId'), memberId);
    if (!rowNumber) throw new Error('회원 정보를 찾을 수 없습니다.');

    const member = memberAtRow_(sheet, rowNumber);
    if (!isActiveMember_(member)) throw new Error('현재 이용할 수 없는 회원입니다.');
    if (member.birthDate && member.birthDate !== birthDate) {
      throw new Error('이미 다른 생년월일이 등록된 회원입니다. 직원에게 문의해 주세요.');
    }

    const date = today_();
    const reservations = prepareVisitWrite_(member, date);
    if (!member.birthDate) {
      sheet.getRange(rowNumber, memberColumn_(sheet, 'birthDate')).setValue(birthDate);
      sheet.getRange(rowNumber, memberColumn_(sheet, 'updatedAt')).setValue(nowText_());
    }
    const updatedMember = memberAtRow_(sheet, rowNumber);
    const visit = ensureVisitLogLocked_(updatedMember, date, source, reservations);
    return {
      ok: true,
      member: publicMember_(updatedMember),
      visit: visit
    };
  });
}

function handleSubmitVisitLog_(payload) {
  const member = requireActiveMember_(payload.memberId);
  const date = today_();
  const source = normalizeVisitSource_(payload.source);

  if (!member.birthDate) throw new Error('생년월일을 먼저 입력해 주세요.');

  return withScriptLock_(function() {
    return ensureVisitLogLocked_(member, date, source);
  });
}

function prepareVisitWrite_(member, date) {
  const memberSheet = getOrCreateSheet_(SHEET_NAMES.MEMBERS, HEADERS.members);
  const visitSheet = getOrCreateSheet_(SHEET_NAMES.VISIT_LOGS, HEADERS.visitLogs);
  ensureSheetHeaders_(memberSheet, HEADERS.members);
  ensureSheetHeaders_(visitSheet, HEADERS.visitLogs);
  return getTodayReservationsForMember_(member, date);
}

function ensureVisitLogLocked_(member, date, source, preparedReservations) {
  const reservations = Array.isArray(preparedReservations)
    ? preparedReservations
    : prepareVisitWrite_(member, date);
  if (hasVisitLog_(member.id, date)) {
    updateMemberLastUsedAt_(member.id, date);
    return {
      ok: true,
      created: false,
      date: date,
      reservations: reservations
    };
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
    createdBy: 'self',
    createdAt: nowText_()
  });
  updateMemberLastUsedAt_(member.id, date);
  return {
    ok: true,
    created: true,
    date: date,
    reservations: reservations
  };
}

function handleSubmitUsageRecord_(payload) {
  const requestedType = str_(payload.type);
  if (requestedType !== 'printer' && requestedType !== 'careerZone' && requestedType !== 'connectRoom') {
    throw new Error('지원하지 않는 일지 종류입니다.');
  }

  const member = requireActiveMember_(payload.memberId);
  const data = payload.data || {};
  const type = requestedType === 'careerZone'
    ? normalizeCareerRoomType_(data.spaceType)
    : requestedType;
  if (requestedType === 'careerZone' && !type) {
    throw new Error('이용할 커리어존 Room1 또는 Room2를 선택해 주세요.');
  }
  const date = today_();
  if (!member.birthDate) throw new Error('생년월일을 먼저 입력해 주세요.');
  if (!hasVisitLog_(member.id, date)) throw new Error('오늘 방문 확인을 먼저 진행해 주세요.');

  if (type === 'printer') {
    const count = Number(data.count || 0);
    const requestedId = str_(payload.requestId);
    if (requestedId && !/^[A-Za-z0-9_-]{8,100}$/.test(requestedId)) {
      throw new Error('프린터 요청 정보를 확인해 주세요.');
    }
    const requestId = requestedId || Utilities.getUuid();
    if (!Number.isInteger(count) || count < 1 || count > 10) {
      throw new Error('프린터 사용 매수는 1장부터 10장까지 입력해 주세요.');
    }

    return withScriptLock_(function() {
      const sheet = getOrCreateSheet_(SHEET_NAMES.PRINTER_LOGS, HEADERS.printerLogs);
      ensureSheetHeaders_(sheet, HEADERS.printerLogs);
      const requestRows = findRowsByColumnValue_(
        sheet,
        HEADERS.printerLogs.indexOf('requestId') + 1,
        requestId
      );
      if (requestRows.length > 0) {
        const existing = objectAtRow_(sheet, requestRows[0], HEADERS.printerLogs);
        if (str_(existing.userId) !== str_(member.id)) {
          throw new Error('이미 사용된 프린터 요청번호입니다. 다시 제출해 주세요.');
        }
        return { ok: true, created: false, requestId: requestId };
      }
      if (getPrinterCountForDate_(member.id, date) + count > 10) {
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
        createdBy: 'self',
        createdAt: nowText_(),
        requestId: requestId
      });
      return { ok: true, created: true, requestId: requestId };
    });
  }

  const roomData = {
    date: date,
    startTime: normalizeTime_(data.startTime),
    endTime: normalizeTime_(data.endTime),
    purpose: str_(data.purpose),
    companions: normalizeCompanions_(data.companions)
  };
  if (!roomData.startTime || !roomData.endTime || !roomData.purpose) {
    throw new Error('시작 시간, 종료 시간, 이용 목적을 모두 입력해 주세요.');
  }
  if (roomData.endTime <= roomData.startTime) {
    throw new Error('종료 시간은 시작 시간보다 늦어야 합니다.');
  }
  if (roomData.purpose.length > 200) throw new Error('이용 목적은 200자 이내로 입력해 주세요.');

  return withScriptLock_(function() {
    const roomLogSheet = getRoomLogSheetName_(type);
    const reservationSheet = ensureNaverReservationSchema_();
    const overlappingReservations = findOverlappingNaverReservationEntries_(
      reservationSheet,
      type,
      roomData.date,
      roomData.startTime,
      roomData.endTime
    );
    const otherReservations = overlappingReservations.filter(function(entry) {
      return !reservationBelongsToMember_(entry.values, member);
    });
    if (otherReservations.length > 0) {
      throw new Error('해당 시간대에는 다른 이용자의 네이버 예약이 있습니다. 예약 시간을 확인해 주세요.');
    }

    let ownReservations = overlappingReservations.filter(function(entry) {
      return reservationBelongsToMember_(entry.values, member);
    });
    if (ownReservations.length > 1) {
      const exactMatches = ownReservations.filter(function(entry) {
        return normalizeSheetTime_(entry.values.startTime) === roomData.startTime &&
          normalizeSheetTime_(entry.values.endTime) === roomData.endTime;
      });
      if (exactMatches.length !== 1) {
        throw new Error('겹치는 본인 예약이 여러 건입니다. 화면 위의 예약 일정을 선택해 제출해 주세요.');
      }
      ownReservations = exactMatches;
    }
    if (ownReservations.length === 1) {
      return submitReservationUsageLocked_(
        member,
        reservationSheet,
        ownReservations[0],
        roomData.companions,
        { duplicateAsError: true, source: 'manual-linked-reservation' }
      );
    }

    assertNoRoomLogConflict_(roomLogSheet, member, type, roomData);
    const roomRow = buildRoomLogRow_(member, type, roomData, {
      createdBy: 'self',
      linkedReservationId: '',
      dedupeStatus: 'manual'
    });
    appendObjectRow_(roomLogSheet, HEADERS.roomLogs, roomRow);
    return { ok: true, linkedReservation: false };
  });
}

function handleSubmitReservationUsage_(payload) {
  const member = requireActiveMember_(payload.memberId);
  const reservationId = str_(payload.reservationId);
  const date = today_();
  if (!reservationId) throw new Error('예약 정보를 다시 선택해 주세요.');
  if (!member.birthDate) throw new Error('생년월일을 먼저 입력해 주세요.');
  if (!hasVisitLog_(member.id, date)) throw new Error('오늘 방문 확인을 먼저 진행해 주세요.');

  return withScriptLock_(function() {
    const sheet = ensureNaverReservationSchema_();
    const entry = findNaverReservationEntry_(sheet, reservationId);
    if (!entry) throw new Error('예약 정보를 찾을 수 없습니다. 시트에서 예약 동기화를 다시 실행해 주세요.');
    return submitReservationUsageLocked_(member, sheet, entry, payload.companions, {
      duplicateAsError: true,
      repairedAsSuccess: true,
      source: 'reservation-card'
    });
  });
}

function submitReservationUsageLocked_(member, sheet, entry, companionsValue, options) {
  const reservation = entry.values;
  const reservationId = str_(reservation.reservationId);
  const date = today_();
  const settings = options || {};
  if (normalizeSheetDate_(reservation.usageDate) !== date) {
    throw new Error('오늘 이용하는 예약만 제출할 수 있습니다.');
  }
  if (isCancelledNaverStatus_(reservation.status)) {
    throw new Error('취소 또는 노쇼 처리된 예약입니다.');
  }
  if (!reservationBelongsToMember_(reservation, member)) {
    throw new Error('로그인한 회원의 예약과 일치하지 않습니다.');
  }
  if (str_(reservation.usageLogId)) {
    if (settings.duplicateAsError) throw new Error('이미 동일한 시간대로 제출하였습니다.');
    return { ok: true, created: false, reservationId: reservationId };
  }

  const linkedLog = findLinkedReservationLog_(reservationId);
  if (linkedLog) {
    if (str_(linkedLog.row.userId) !== str_(member.id)) {
      throw new Error('이미 다른 회원의 이용일지와 연결된 예약입니다. 관리자에게 확인해 주세요.');
    }
    updateNaverReservationSubmission_(sheet, entry.rowNumber, member, reservation, linkedLog.row.logId);
    if (settings.duplicateAsError && !settings.repairedAsSuccess) {
      throw new Error('이미 동일한 시간대로 제출하였습니다.');
    }
    return { ok: true, created: false, repaired: true, reservationId: reservationId };
  }

  const spaceType = str_(reservation.spaceType);
  const roomLogSheet = getRoomLogSheetName_(spaceType);
  const expectedCompanionCount = Math.max(0, Number(reservation.headcount || 1) - 1);
  const roomData = {
    date: date,
    startTime: normalizeTime_(reservation.startTime),
    endTime: normalizeTime_(reservation.endTime),
    purpose: str_(reservation.purpose, '네이버 예약 이용'),
    companions: normalizeCompanions_(companionsValue, expectedCompanionCount)
  };
  if (roomData.companions.length !== expectedCompanionCount) {
    throw new Error('예약 인원에 맞게 동반 이용자의 성별과 나이를 모두 입력해 주세요.');
  }
  if (!roomData.startTime || !roomData.endTime || roomData.endTime <= roomData.startTime) {
    throw new Error('예약 시간이 올바르지 않습니다. CSV 내용을 확인해 주세요.');
  }
  assertNoRoomLogConflict_(roomLogSheet, member, spaceType, roomData);

  const roomRow = buildRoomLogRow_(member, spaceType, roomData, {
    createdBy: 'naver-reservation',
    linkedReservationId: reservationId,
    dedupeStatus: settings.source || 'matched_naver'
  });
  appendObjectRow_(roomLogSheet, HEADERS.roomLogs, roomRow);
  updateNaverReservationSubmission_(sheet, entry.rowNumber, member, reservation, roomRow.logId);
  return { ok: true, created: true, reservationId: reservationId, linkedReservation: true };
}

function buildRoomLogRow_(member, spaceType, roomData, metadata) {
  const details = metadata || {};
  const row = {
    logId: Utilities.getUuid(),
    userId: member.id,
    spaceType: spaceType,
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
    companionsJson: JSON.stringify(roomData.companions || []),
    companionCount: (roomData.companions || []).length,
    createdBy: str_(details.createdBy, 'self'),
    createdAt: nowText_(),
    linkedReservationId: str_(details.linkedReservationId),
    dedupeStatus: str_(details.dedupeStatus)
  };
  const summary = buildRoomSummary_(member, roomData);
  Object.keys(summary).forEach(function(key) { row[key] = summary[key]; });
  return row;
}

function updateNaverReservationSubmission_(sheet, rowNumber, member, reservation, logId) {
  updateNaverReservationRow_(sheet, rowNumber, {
    onsiteMemberName: member.name,
    onsiteMemberNameKey: normalizeNameKey_(member.name),
    onsiteCheckedInAt: nowText_(),
    onsiteCheckinSource: str_(reservation.spaceType),
    usageLogId: logId,
    submittedAt: nowText_(),
    submittedByMemberId: member.id
  });
}

function getRoomLogSheetName_(spaceType) {
  if (isCareerRoomType_(spaceType)) return SHEET_NAMES.CAREER_LOGS;
  if (spaceType === 'connectRoom') return SHEET_NAMES.CONNECT_LOGS;
  throw new Error('지원하지 않는 예약 장소입니다.');
}

function assertNoRoomLogConflict_(sheetName, member, spaceType, roomData) {
  const sheet = getOrCreateSheet_(sheetName, HEADERS.roomLogs);
  const conflicts = findObjectsByColumnValue_(sheet, HEADERS.roomLogs, 'date', roomData.date).filter(function(row) {
    return roomResourcesConflict_(row.spaceType, spaceType) &&
      timeRangesOverlap_(row.startTime, row.endTime, roomData.startTime, roomData.endTime);
  });
  const exactSameMember = conflicts.some(function(row) {
    return str_(row.userId) === str_(member.id) &&
      normalizeSheetTime_(row.startTime) === roomData.startTime &&
      normalizeSheetTime_(row.endTime) === roomData.endTime;
  });
  if (exactSameMember) throw new Error('이미 동일한 시간대로 제출하였습니다.');
  if (conflicts.length > 0) {
    throw new Error('해당 시간대에는 이미 제출된 이용일지가 있습니다. 다른 시간을 확인해 주세요.');
  }
}

function findOverlappingNaverReservationEntries_(sheet, spaceType, date, startTime, endTime) {
  return findEntriesByColumnValue_(sheet, HEADERS.naverSync, 'usageDate', date).filter(function(entry) {
    const reservation = entry.values;
    return roomResourcesConflict_(reservation.spaceType, spaceType) &&
      !isCancelledNaverStatus_(reservation.status) &&
      timeRangesOverlap_(reservation.startTime, reservation.endTime, startTime, endTime);
  });
}

function normalizeCareerRoomType_(value) {
  const type = str_(value);
  if (type === 'careerRoom1' || type === 'careerRoom2') return type;
  return '';
}

function isCareerRoomType_(value) {
  const type = str_(value);
  return type === 'careerZone' || type === 'careerRoom1' || type === 'careerRoom2';
}

function roomResourcesConflict_(left, right) {
  const leftType = str_(left);
  const rightType = str_(right);
  if (leftType === rightType) return true;
  if (!leftType || !rightType) return true;
  if (isCareerRoomType_(leftType) && isCareerRoomType_(rightType)) {
    return leftType === 'careerZone' || rightType === 'careerZone';
  }
  return false;
}

function timeRangesOverlap_(startA, endA, startB, endB) {
  const normalizedStartA = normalizeSheetTime_(startA);
  const normalizedEndA = normalizeSheetTime_(endA);
  const normalizedStartB = normalizeSheetTime_(startB);
  const normalizedEndB = normalizeSheetTime_(endB);
  if (!normalizedStartA || !normalizedEndA || !normalizedStartB || !normalizedEndB) return false;
  return normalizedStartA < normalizedEndB && normalizedStartB < normalizedEndA;
}

function buildRoomSummary_(member, data) {
  const members = [member].concat(data.companions || []);
  const stats = buildParticipantStats_(members);
  return {
    headcount: members.length,
    participants: buildParticipantDetail_(members),
    male20s: stats['20대_남성'] || 0,
    female20s: stats['20대_여성'] || 0,
    male30s: stats['30대_남성'] || 0,
    female30s: stats['30대_여성'] || 0,
    maleUnder19: stats['19세 이하_남성'] || 0,
    femaleUnder19: stats['19세 이하_여성'] || 0,
    male40Plus: stats['40세 이상_남성'] || 0,
    female40Plus: stats['40세 이상_여성'] || 0
  };
}

function backfillMemberLastUsedAt() {
  ensureDataSheetSchemas_();
  const result = withScriptLock_(backfillMemberLastUsedAt_);
  SpreadsheetApp.getUi().alert(
    '마지막 이용일 계산을 완료했습니다. 전체 ' + result.total + '명 중 ' + result.updated + '명을 갱신했습니다.'
  );
  return result;
}

function backfillMemberLastUsedAt_() {
  const latestByMember = {};
  getRowsAsObjects_(SHEET_NAMES.VISIT_LOGS, HEADERS.visitLogs).forEach(function(log) {
    const memberId = str_(log.userId);
    const visitDate = dateOnly_(log.date);
    if (memberId && visitDate && (!latestByMember[memberId] || visitDate > latestByMember[memberId])) {
      latestByMember[memberId] = visitDate;
    }
  });

  const sheet = getOrCreateSheet_(SHEET_NAMES.MEMBERS, HEADERS.members);
  if (sheet.getLastRow() < 2) return { total: 0, updated: 0 };

  const rowCount = sheet.getLastRow() - 1;
  const memberIdColumn = HEADERS.members.indexOf('memberId');
  const createdAtColumn = HEADERS.members.indexOf('createdAt');
  const lastUsedColumn = HEADERS.members.indexOf('lastUsedAt');
  const values = sheet.getRange(2, 1, rowCount, HEADERS.members.length).getValues();
  let updated = 0;
  const lastUsedValues = values.map(function(row) {
    const memberId = str_(row[memberIdColumn]);
    const current = dateOnly_(row[lastUsedColumn]);
    const createdAt = dateOnly_(row[createdAtColumn]);
    const calculated = [current, latestByMember[memberId], createdAt]
      .filter(Boolean)
      .sort()
      .pop() || '';
    if (calculated !== current) updated += 1;
    return [calculated];
  });
  sheet.getRange(2, lastUsedColumn + 1, rowCount, 1).setValues(lastUsedValues);
  return { total: rowCount, updated: updated };
}

function previewExpiredMembers() {
  ensureDataSheetSchemas_();
  const result = withScriptLock_(function() {
    backfillMemberLastUsedAt_();
    return getRetentionCandidates_();
  });
  if (result.length === 0) {
    SpreadsheetApp.getUi().alert('마지막 이용일로부터 3년이 지난 회원이 없습니다.');
    return [];
  }

  const lines = result.slice(0, 20).map(function(item) {
    return maskName_(item.member.name) + ' / 끝자리 ' + item.member.phoneLastDigits + ' / 마지막 이용일 ' + item.lastUsedAt;
  });
  if (result.length > lines.length) lines.push('외 ' + (result.length - lines.length) + '명');
  SpreadsheetApp.getUi().alert('3년 초과 회원 ' + result.length + '명\n\n' + lines.join('\n'));
  return result.map(function(item) {
    return { memberId: item.member.id, lastUsedAt: item.lastUsedAt };
  });
}

function purgeExpiredMemberData() {
  ensureDataSheetSchemas_();
  const candidates = withScriptLock_(function() {
    backfillMemberLastUsedAt_();
    return getRetentionCandidates_();
  });
  if (candidates.length === 0) {
    SpreadsheetApp.getUi().alert('마지막 이용일로부터 3년이 지난 회원이 없습니다.');
    return emptyRetentionResult_();
  }

  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    '개인정보 파기 확인',
    '회원 ' + candidates.length + '명의 회원정보를 삭제하고 상세 일지의 이름·전화번호·회원번호를 제거합니다. 계속하시겠습니까?',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return emptyRetentionResult_();

  const result = runRetentionCleanup();
  ui.alert(
    '개인정보 파기를 완료했습니다.\n회원 ' + result.memberCount + '명\n상세 일지 ' + result.totalLogCount + '건'
  );
  return result;
}

function runRetentionCleanup() {
  ensureDataSheetSchemas_();
  return withScriptLock_(function() {
    backfillMemberLastUsedAt_();
    const candidates = getRetentionCandidates_();
    if (candidates.length === 0) {
      return emptyRetentionResult_();
    }

    const memberIds = candidates.map(function(item) { return item.member.id; });
    const logCounts = anonymizeMemberLogs_(memberIds);
    const memberSheet = getOrCreateSheet_(SHEET_NAMES.MEMBERS, HEADERS.members);
    candidates
      .map(function(item) { return item.rowNumber; })
      .sort(function(a, b) { return b - a; })
      .forEach(function(rowNumber) { memberSheet.deleteRow(rowNumber); });
    appendRetentionAudit_('scheduled_retention', candidates.length, logCounts);
    return retentionResult_(candidates.length, logCounts);
  });
}

function setupRetentionCleanupTrigger() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    '월간 자동 파기 설정',
    '먼저 3년 초과 회원 미리보기 결과를 확인했나요? 설정 후 매월 1일 자동으로 개인정보 파기가 실행됩니다.',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runRetentionCleanup') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('runRetentionCleanup')
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .create();
  ui.alert('매월 1일 오전 3시에 개인정보 보관기간을 확인하도록 설정했습니다.');
}

function getRetentionCandidates_() {
  const sheet = getOrCreateSheet_(SHEET_NAMES.MEMBERS, HEADERS.members);
  if (sheet.getLastRow() < 2) return [];

  const cutoff = retentionCutoff_();
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.members.length).getValues();
  return values.map(function(row, index) {
    const obj = {};
    HEADERS.members.forEach(function(header, column) { obj[header] = row[column]; });
    const member = memberRowToObject_(obj);
    return {
      rowNumber: index + 2,
      member: member,
      lastUsedAt: member.lastUsedAt || dateOnly_(member.createdAt)
    };
  }).filter(function(item) {
    return item.member.id &&
      item.lastUsedAt &&
      item.lastUsedAt < cutoff &&
      !hasActiveRetentionHold_(item.member);
  });
}

function hasActiveRetentionHold_(member) {
  return Boolean(member && member.retentionHoldUntil && member.retentionHoldUntil >= today_());
}

function anonymizeMemberLogs_(memberIds) {
  const ids = {};
  (memberIds || []).forEach(function(memberId) {
    if (memberId) ids[memberId] = true;
  });
  return {
    visitLogCount: anonymizeLogSheet_(
      SHEET_NAMES.VISIT_LOGS,
      HEADERS.visitLogs,
      ids,
      ['userId', 'name', 'phone', 'phoneLastDigits']
    ),
    printerLogCount: anonymizeLogSheet_(
      SHEET_NAMES.PRINTER_LOGS,
      HEADERS.printerLogs,
      ids,
      ['userId', 'name', 'phone', 'phoneLastDigits']
    ),
    careerLogCount: anonymizeLogSheet_(
      SHEET_NAMES.CAREER_LOGS,
      HEADERS.roomLogs,
      ids,
      ['userId', 'name', 'memberNameKey', 'phone', 'phoneLastDigits']
    ),
    connectLogCount: anonymizeLogSheet_(
      SHEET_NAMES.CONNECT_LOGS,
      HEADERS.roomLogs,
      ids,
      ['userId', 'name', 'memberNameKey', 'phone', 'phoneLastDigits']
    ),
    reservationLinkCount: anonymizeNaverMemberLinks_(ids)
  };
}

function anonymizeLogSheet_(sheetName, headers, memberIds, personalHeaders) {
  const sheet = getOrCreateSheet_(sheetName, headers);
  if (sheet.getLastRow() < 2) return 0;

  const rowCount = sheet.getLastRow() - 1;
  const values = sheet.getRange(2, 1, rowCount, headers.length).getValues();
  const memberIdColumn = headers.indexOf('userId');
  const personalColumns = personalHeaders.map(function(header) { return headers.indexOf(header); });
  let changed = 0;
  values.forEach(function(row) {
    if (!memberIds[str_(row[memberIdColumn])]) return;
    personalColumns.forEach(function(column) {
      if (column >= 0) row[column] = '';
    });
    changed += 1;
  });
  if (changed > 0) {
    sheet.getRange(2, 1, rowCount, headers.length).setValues(values);
  }
  return changed;
}

function anonymizeNaverMemberLinks_(memberIds) {
  const sheet = ensureNaverReservationSchema_();
  if (sheet.getLastRow() < 2) return 0;

  const rowCount = sheet.getLastRow() - 1;
  const values = sheet.getRange(2, 1, rowCount, HEADERS.naverSync.length).getValues();
  const memberIdColumn = HEADERS.naverSync.indexOf('submittedByMemberId');
  const personalColumns = ['onsiteMemberName', 'onsiteMemberNameKey', 'submittedByMemberId']
    .map(function(header) { return HEADERS.naverSync.indexOf(header); });
  let changed = 0;
  values.forEach(function(row) {
    if (!memberIds[str_(row[memberIdColumn])]) return;
    personalColumns.forEach(function(column) {
      if (column >= 0) row[column] = '';
    });
    changed += 1;
  });
  if (changed > 0) {
    sheet.getRange(2, 1, rowCount, HEADERS.naverSync.length).setValues(values);
  }
  return changed;
}

function appendRetentionAudit_(action, memberCount, logCounts) {
  appendObjectRow_(SHEET_NAMES.RETENTION_LOGS, HEADERS.retentionLogs, {
    runId: Utilities.getUuid(),
    action: action,
    processedAt: nowText_(),
    memberCount: memberCount,
    visitLogCount: logCounts.visitLogCount || 0,
    printerLogCount: logCounts.printerLogCount || 0,
    careerLogCount: logCounts.careerLogCount || 0,
    connectLogCount: logCounts.connectLogCount || 0,
    reservationLinkCount: logCounts.reservationLinkCount || 0
  });
}

function retentionResult_(memberCount, logCounts) {
  const result = {
    memberCount: memberCount,
    visitLogCount: logCounts.visitLogCount || 0,
    printerLogCount: logCounts.printerLogCount || 0,
    careerLogCount: logCounts.careerLogCount || 0,
    connectLogCount: logCounts.connectLogCount || 0,
    reservationLinkCount: logCounts.reservationLinkCount || 0
  };
  result.totalLogCount = result.visitLogCount + result.printerLogCount + result.careerLogCount +
    result.connectLogCount + result.reservationLinkCount;
  return result;
}

function emptyRetentionResult_() {
  return retentionResult_(0, {});
}

function retentionCutoff_() {
  const parts = today_().split('-').map(Number);
  const cutoff = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  cutoff.setUTCFullYear(cutoff.getUTCFullYear() - MEMBER_RETENTION_YEARS);
  return Utilities.formatDate(cutoff, 'UTC', 'yyyy-MM-dd');
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
  const installedMonthKey = props.getProperty(propertyKey);
  if (installedMonthKey === currentMonthKey) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [
    { name: SHEET_NAMES.VISIT, headers: HEADERS.visit },
    { name: SHEET_NAMES.PRINTER, headers: HEADERS.printer },
    { name: SHEET_NAMES.CAREER, headers: HEADERS.career },
    { name: SHEET_NAMES.CONNECT, headers: HEADERS.connect }
  ].forEach(function(item) {
    const current = ss.getSheetByName(item.name);
    if (current && installedMonthKey) {
      freezeSheetFormulasAsValues_(current);
      current.setName(makeUniqueSheetName_(item.name + '_' + installedMonthKey));
      current.hideSheet();
    }
    const fresh = ss.getSheetByName(item.name) || ss.insertSheet(item.name);
    installPublicSheetFormula_(fresh, item.name, item.headers, currentMonthKey, !installedMonthKey);
  });

  props.setProperty(propertyKey, currentMonthKey);
}

function setupPublicFormulaSheets() {
  ensureDataSheetSchemas_();
  backfillRoomSummaryColumns_(SHEET_NAMES.CAREER_LOGS);
  backfillRoomSummaryColumns_(SHEET_NAMES.CONNECT_LOGS);

  const monthKey = Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [
    { name: SHEET_NAMES.VISIT, headers: HEADERS.visit },
    { name: SHEET_NAMES.PRINTER, headers: HEADERS.printer },
    { name: SHEET_NAMES.CAREER, headers: HEADERS.career },
    { name: SHEET_NAMES.CONNECT, headers: HEADERS.connect }
  ].forEach(function(item) {
    const sheet = ss.getSheetByName(item.name) || ss.insertSheet(item.name);
    installPublicSheetFormula_(sheet, item.name, item.headers, monthKey, true);
  });

  PropertiesService.getScriptProperties().setProperty('MONTHLY_SHEET_ROLLOVER_LAST_RUN', monthKey);

  SpreadsheetApp.getUi().alert('간소화 시트를 상세 일지 기반 수식으로 전환했습니다.');
}

function ensurePublicFormulaSheetSchemas_() {
  const monthKey = Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [
    { name: SHEET_NAMES.VISIT, headers: HEADERS.visit },
    { name: SHEET_NAMES.PRINTER, headers: HEADERS.printer },
    { name: SHEET_NAMES.CAREER, headers: HEADERS.career },
    { name: SHEET_NAMES.CONNECT, headers: HEADERS.connect }
  ].forEach(function(item) {
    const sheet = ss.getSheetByName(item.name) || ss.insertSheet(item.name);
    const currentHeaders = sheet.getRange(1, 1, 1, item.headers.length).getValues()[0];
    const headersMatch = item.headers.every(function(header, index) {
      return str_(currentHeaders[index]) === header;
    });
    const formula = str_(sheet.getRange('A2').getFormula());
    const roomFormulaCurrent = item.name !== SHEET_NAMES.CAREER && item.name !== SHEET_NAMES.CONNECT
      ? true
      : formula.indexOf('!AC2:INDEX(') !== -1;
    if (!headersMatch || !formula || !roomFormulaCurrent) {
      if (item.name === SHEET_NAMES.CAREER) backfillRoomSummaryColumns_(SHEET_NAMES.CAREER_LOGS);
      if (item.name === SHEET_NAMES.CONNECT) backfillRoomSummaryColumns_(SHEET_NAMES.CONNECT_LOGS);
      installPublicSheetFormula_(sheet, item.name, item.headers, monthKey, true);
    }
  });
}

function installPublicSheetFormula_(sheet, sheetName, headers, monthKey, preserveExisting) {
  if (preserveExisting) backupPublicSheetIfNeeded_(sheet);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange('A2').setFormula(buildPublicSheetFormula_(sheetName, monthKey));
}

function buildPublicSheetFormula_(sheetName, monthKey) {
  if (sheetName === SHEET_NAMES.VISIT) {
    return '=IFERROR(LET(last,MAX(2,COUNTA(APP_VISIT_LOGS!A:A)),dates,APP_VISIT_LOGS!I2:INDEX(APP_VISIT_LOGS!I:I,last),' +
      'data,FILTER(HSTACK(dates,APP_VISIT_LOGS!E2:INDEX(APP_VISIT_LOGS!E:E,last),APP_VISIT_LOGS!F2:INDEX(APP_VISIT_LOGS!F:F,last)),' +
      'IFERROR(TEXT(dates,"yyyy-mm"),LEFT(dates&"",7))="' + monthKey + '"),HSTACK(SEQUENCE(ROWS(data)),data)),"")';
  }
  if (sheetName === SHEET_NAMES.PRINTER) {
    return '=IFERROR(LET(last,MAX(2,COUNTA(APP_PRINTER_LOGS!A:A)),dates,APP_PRINTER_LOGS!H2:INDEX(APP_PRINTER_LOGS!H:H,last),' +
      'data,FILTER(HSTACK(dates,APP_PRINTER_LOGS!D2:INDEX(APP_PRINTER_LOGS!D:D,last),APP_PRINTER_LOGS!E2:INDEX(APP_PRINTER_LOGS!E:E,last),' +
      'APP_PRINTER_LOGS!I2:INDEX(APP_PRINTER_LOGS!I:I,last)),IFERROR(TEXT(dates,"yyyy-mm"),LEFT(dates&"",7))="' + monthKey + '"),' +
      'HSTACK(SEQUENCE(ROWS(data)),data)),"")';
  }

  const source = sheetName === SHEET_NAMES.CAREER ? 'APP_CAREER_LOGS' : 'APP_CONNECT_LOGS';
  return '=IFERROR(LET(last,MAX(2,COUNTA(' + source + '!A:A)),dates,' + source + '!J2:INDEX(' + source + '!J:J,last),' +
    'data,FILTER(HSTACK(dates,' + source + '!M2:INDEX(' + source + '!M:M,last),' + source + '!T2:INDEX(' + source + '!T:T,last),' +
    source + '!U2:INDEX(' + source + '!U:U,last),' + source + '!K2:INDEX(' + source + '!K:K,last),' + source + '!L2:INDEX(' + source + '!L:L,last),' +
    source + '!V2:INDEX(' + source + '!V:V,last),' + source + '!W2:INDEX(' + source + '!W:W,last),' + source + '!X2:INDEX(' + source + '!X:X,last),' +
    source + '!Y2:INDEX(' + source + '!Y:Y,last),' + source + '!Z2:INDEX(' + source + '!Z:Z,last),' + source + '!AA2:INDEX(' + source + '!AA:AA,last),' +
    source + '!AB2:INDEX(' + source + '!AB:AB,last),' + source + '!AC2:INDEX(' + source + '!AC:AC,last)),' +
    'IFERROR(TEXT(dates,"yyyy-mm"),LEFT(dates&"",7))="' + monthKey + '"),' +
    'HSTACK(SEQUENCE(ROWS(data)),data)),"")';
}

function backupPublicSheetIfNeeded_(sheet) {
  if (sheet.getLastRow() < 2 || str_(sheet.getRange('A2').getFormula())) return;
  const timestamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss');
  const backup = sheet.copyTo(SpreadsheetApp.getActiveSpreadsheet());
  backup.setName(makeUniqueSheetName_(sheet.getName() + '_수식전백업_' + timestamp));
  backup.hideSheet();
}

function makeUniqueSheetName_(baseName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let name = baseName.slice(0, 90);
  let suffix = 2;
  while (ss.getSheetByName(name)) {
    name = (baseName.slice(0, 86) + '_' + suffix).slice(0, 90);
    suffix += 1;
  }
  return name;
}

function freezeSheetFormulasAsValues_(sheet) {
  const range = sheet.getDataRange();
  if (range.getNumRows() < 2 || !str_(sheet.getRange('A2').getFormula())) return;
  const values = range.getValues();
  range.clearContent();
  range.setValues(values);
}

function backfillRoomSummaryColumns_(sheetName) {
  const sheet = getOrCreateSheet_(sheetName, HEADERS.roomLogs);
  if (sheet.getLastRow() < 2) return;

  const rows = getRowsAsObjects_(sheetName, HEADERS.roomLogs);
  const summaryStart = HEADERS.roomLogs.indexOf('headcount');
  const summaryHeaders = HEADERS.roomLogs.slice(summaryStart);
  const values = rows.map(function(row) {
    let companions = [];
    try {
      companions = JSON.parse(str_(row.companionsJson, '[]'));
      if (!Array.isArray(companions)) companions = [];
    } catch (error) {}
    const summary = buildRoomSummary_(
      { gender: str_(row.gender), age: Number(row.age || 0) },
      { date: str_(row.date), purpose: str_(row.purpose), startTime: str_(row.startTime), endTime: str_(row.endTime), companions: companions }
    );
    return summaryHeaders.map(function(header) {
      return row[header] !== '' && row[header] !== undefined ? row[header] : summary[header];
    });
  });
  sheet.getRange(2, summaryStart + 1, values.length, summaryHeaders.length).setValues(values);
}


function getMemberById_(memberId) {
  if (!memberId) return null;
  const sheet = getOrCreateSheet_(SHEET_NAMES.MEMBERS, HEADERS.members);
  const rowNumber = findFirstRowByColumnValue_(sheet, memberColumn_(sheet, 'memberId'), memberId);
  return rowNumber ? memberAtRow_(sheet, rowNumber) : null;
}

function updateMemberLastUsedAt_(memberId, date) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.MEMBERS, HEADERS.members);
  ensureSheetHeaders_(sheet, HEADERS.members);
  const rowNumber = findFirstRowByColumnValue_(sheet, memberColumn_(sheet, 'memberId'), memberId);
  if (!rowNumber) throw new Error('회원 정보를 찾을 수 없습니다.');
  sheet.getRange(rowNumber, memberColumn_(sheet, 'lastUsedAt')).setValue(dateOnly_(date));
  sheet.getRange(rowNumber, memberColumn_(sheet, 'updatedAt')).setValue(nowText_());
}

function findMembersByField_(field, value) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.MEMBERS, HEADERS.members);
  const target = digits_(value);
  const column = memberColumn_(sheet, field);
  const phoneColumn = memberColumn_(sheet, 'phone');
  const phoneTailColumn = memberColumn_(sheet, 'phoneLastDigits');
  const isPhoneSearch = field === 'phone';
  const isPhoneTailSearch = field === 'phoneLastDigits';
  if (!isPhoneSearch && !isPhoneTailSearch) {
    return findRowsByColumnValue_(sheet, column, value).map(function(rowNumber) {
      return memberAtRow_(sheet, rowNumber);
    });
  }

  const rowNumbers = {};
  findRowsByColumnValue_(sheet, column, value).forEach(function(rowNumber) {
    rowNumbers[rowNumber] = true;
  });

  if (isPhoneSearch) {
    findRowsByColumnRegex_(sheet, phoneColumn, buildPhoneRegex_(target)).forEach(function(rowNumber) {
      rowNumbers[rowNumber] = true;
    });
  } else {
    findRowsByColumnRegex_(sheet, phoneTailColumn, buildPhoneTailRegex_(target)).forEach(function(rowNumber) {
      rowNumbers[rowNumber] = true;
    });
    findRowsByColumnRegex_(sheet, phoneColumn, buildPhoneSuffixRegex_(target)).forEach(function(rowNumber) {
      rowNumbers[rowNumber] = true;
    });
  }

  return Object.keys(rowNumbers)
    .map(Number)
    .map(function(rowNumber) { return memberAtRow_(sheet, rowNumber); })
    .filter(function(member) {
      if (isPhoneSearch) return member.phone === target;
      return member.phoneLastDigits === target || member.phone.slice(-4) === target;
    });
}

function memberColumn_(sheet, header) {
  const lastColumn = Math.max(sheet.getLastColumn(), HEADERS.members.length);
  const currentHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const target = str_(header).trim().toLowerCase();
  const index = currentHeaders.findIndex(function(value) {
    return str_(value).trim().toLowerCase() === target;
  });
  if (index < 0) throw new Error('APP_MEMBERS 시트에서 ' + header + ' 열을 찾을 수 없습니다.');
  return index + 1;
}

function memberAtRow_(sheet, rowNumber) {
  const width = Math.max(sheet.getLastColumn(), HEADERS.members.length);
  const currentHeaders = sheet.getRange(1, 1, 1, width).getValues()[0];
  const values = sheet.getRange(rowNumber, 1, 1, width).getValues()[0];
  const columnByHeader = {};
  currentHeaders.forEach(function(header, index) {
    const key = str_(header).trim().toLowerCase();
    if (key && columnByHeader[key] === undefined) columnByHeader[key] = index;
  });
  const row = {};
  HEADERS.members.forEach(function(header) {
    const index = columnByHeader[header.toLowerCase()];
    row[header] = index === undefined ? '' : values[index];
  });
  return memberRowToObject_(row);
}

function requireActiveMember_(memberId) {
  const member = getMemberById_(str_(memberId));
  if (!member) throw new Error('회원 정보를 찾을 수 없습니다. 처음부터 다시 진행해 주세요.');
  if (!isActiveMember_(member)) throw new Error('현재 이용할 수 없는 회원입니다.');
  return member;
}

function publicMember_(member) {
  return {
    id: member.id,
    name: maskName_(member.name),
    gender: member.gender,
    age: member.age,
    hasBirthDate: Boolean(member.birthDate),
    phoneLastDigits: member.phoneLastDigits
  };
}

function maskName_(value) {
  const name = str_(value, '회원');
  if (name.length <= 1) return name;
  if (name.length === 2) return name.charAt(0) + '○';
  return name.charAt(0) + new Array(name.length - 1).join('○') + name.charAt(name.length - 1);
}

function isActiveMember_(member) {
  return str_(member && member.status, 'approved') === 'approved';
}

function hasVisitLog_(userId, date) {
  if (!userId) return false;
  const sheet = getOrCreateSheet_(SHEET_NAMES.VISIT_LOGS, HEADERS.visitLogs);
  return findRowsByColumnValue_(sheet, 2, userId).some(function(rowNumber) {
    return normalizeSheetDate_(sheet.getRange(rowNumber, 9).getValue()) === date;
  });
}

function getPrinterCountForDate_(userId, date) {
  if (!userId) return 0;
  const sheet = getOrCreateSheet_(SHEET_NAMES.PRINTER_LOGS, HEADERS.printerLogs);
  return findRowsByColumnValue_(sheet, 2, userId).reduce(function(sum, rowNumber) {
    const values = sheet.getRange(rowNumber, 8, 1, 2).getValues()[0];
    if (normalizeSheetDate_(values[0]) !== date) return sum;
    return sum + Number(values[1] || 0);
  }, 0);
}

function findRowsByColumnValue_(sheet, column, value) {
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, column, sheet.getLastRow() - 1, 1)
    .createTextFinder(str_(value))
    .matchEntireCell(true)
    .findAll()
    .map(function(range) { return range.getRow(); });
}

function findRowsByColumnRegex_(sheet, column, pattern) {
  if (!pattern || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, column, sheet.getLastRow() - 1, 1)
    .createTextFinder(pattern)
    .useRegularExpression(true)
    .matchEntireCell(true)
    .findAll()
    .map(function(range) { return range.getRow(); });
}

function buildPhoneRegex_(phone) {
  if (!isValidPhone_(phone)) return '';
  return '^0?' + phone.substring(1).split('').join('[^0-9]*') + '[^0-9]*$';
}

function buildPhoneSuffixRegex_(phoneTail) {
  if (!/^\d{4}$/.test(phoneTail)) return '';
  return '.*' + phoneTail.split('').join('[^0-9]*') + '[^0-9]*$';
}

function buildPhoneTailRegex_(phoneTail) {
  if (!/^\d{4}$/.test(phoneTail)) return '';
  const withoutLeadingZero = phoneTail.replace(/^0+/, '');
  if (!withoutLeadingZero) return '^0+$';
  return '^0*' + withoutLeadingZero.split('').join('[^0-9]*') + '[^0-9]*$';
}

function findFirstRowByColumnValue_(sheet, column, value) {
  const rows = findRowsByColumnValue_(sheet, column, value);
  return rows.length > 0 ? rows[0] : 0;
}

function normalizeSheetDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  }
  return str_(value);
}





function buildParticipantStats_(members) {
  const stats = {};

  (members || []).forEach(function(person) {
    const gender = str_(person.gender, '미기재');
    const age = Number(person.age || 0);
    const ageGroup = getAgeGroupLabel_(age);
    const ageGenderKey = ageGroup + '_' + gender;
    stats[ageGenderKey] = (stats[ageGenderKey] || 0) + 1;
  });

  return stats;
}




function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = findExistingSheet_(ss, name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    ensureSheetHeaders_(sheet, headers);
  }
  return sheet;
}

function findExistingSheet_(spreadsheet, name) {
  const expectedName = str_(name).trim().toLowerCase();
  const exact = spreadsheet.getSheetByName(name);
  const matching = spreadsheet.getSheets().filter(function(sheet) {
    return str_(sheet.getName()).trim().toLowerCase() === expectedName;
  });
  if (exact && exact.getLastRow() > 1) return exact;

  const populated = matching.find(function(sheet) {
    return sheet.getLastRow() > 1;
  });
  return populated || exact || matching[0] || null;
}

function getAgeGroupLabel_(age) {
  if (age <= 19) return '19세 이하';
  if (age <= 29) return '20대';
  if (age <= 39) return '30대';
  return '40세 이상';
}

function ensureDataSheetSchemas_() {
  [
    { name: SHEET_NAMES.MEMBERS, headers: HEADERS.members },
    { name: SHEET_NAMES.VISIT_LOGS, headers: HEADERS.visitLogs },
    { name: SHEET_NAMES.PRINTER_LOGS, headers: HEADERS.printerLogs },
    { name: SHEET_NAMES.CAREER_LOGS, headers: HEADERS.roomLogs },
    { name: SHEET_NAMES.CONNECT_LOGS, headers: HEADERS.roomLogs },
    { name: SHEET_NAMES.RETENTION_LOGS, headers: HEADERS.retentionLogs }
  ].forEach(function(item) {
    ensureSheetHeaders_(getOrCreateSheet_(item.name, item.headers), item.headers);
  });
  ensureNaverReservationSchema_();
}

function ensureSheetHeaders_(sheet, headers) {
  let shouldStyle = false;
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    shouldStyle = true;
  } else {
    const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const conflict = headers.some(function(header, index) {
      return current[index] && String(current[index]) !== header;
    });
    if (conflict) throw new Error(sheet.getName() + ' 시트의 열 순서를 확인해 주세요.');
    const missing = headers.some(function(header, index) { return !current[index]; });
    if (missing) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      shouldStyle = true;
    }
  }

  if (shouldStyle) {
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4285F4')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
  }
}

function ensureNaverReservationSchema_() {
  const sheet = getOrCreateSheet_(SHEET_NAMES.NAVER_SYNC, HEADERS.naverSync);
  if (sheet.getLastRow() === 0) {
    ensureSheetHeaders_(sheet, HEADERS.naverSync);
    return sheet;
  }

  const lastColumn = Math.max(1, sheet.getLastColumn());
  const currentHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(value) {
    return str_(value);
  });
  const exact = HEADERS.naverSync.every(function(header, index) {
    return currentHeaders[index] === header;
  });
  if (exact) return sheet;

  const oldRows = sheet.getLastRow() < 2
    ? []
    : sheet.getRange(2, 1, sheet.getLastRow() - 1, lastColumn).getValues();
  const migratedRows = oldRows.map(function(row) {
    const old = {};
    currentHeaders.forEach(function(header, index) {
      if (header) old[header] = row[index];
    });
    if (!old.spaceType) {
      old.spaceType = normalizeNaverSpaceType_(old.placeName || old.normalizedPlace);
    }
    return HEADERS.naverSync.map(function(header) {
      return old[header] !== undefined ? old[header] : '';
    });
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, HEADERS.naverSync.length).setValues([HEADERS.naverSync]);
  if (migratedRows.length > 0) {
    sheet.getRange(2, 1, migratedRows.length, HEADERS.naverSync.length).setValues(migratedRows);
  }
  sheet.getRange(1, 1, 1, HEADERS.naverSync.length)
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  return sheet;
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

function objectAtRow_(sheet, rowNumber, headers) {
  const values = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const obj = {};
  headers.forEach(function(header, index) { obj[header] = values[index]; });
  return obj;
}

function findEntriesByColumnValue_(sheet, headers, header, value) {
  const column = headers.indexOf(header) + 1;
  if (column < 1) throw new Error('검색할 열을 찾을 수 없습니다: ' + header);
  return findRowsByColumnValue_(sheet, column, value).map(function(rowNumber) {
    return {
      rowNumber: rowNumber,
      values: objectAtRow_(sheet, rowNumber, headers)
    };
  });
}

function findObjectsByColumnValue_(sheet, headers, header, value) {
  return findEntriesByColumnValue_(sheet, headers, header, value).map(function(entry) {
    return entry.values;
  });
}

function appendObjectRow_(sheetName, headers, obj) {
  const sheet = getOrCreateSheet_(sheetName, headers);
  ensureSheetHeaders_(sheet, headers);
  const rowNumber = sheet.getLastRow() + 1;
  if (sheetName === SHEET_NAMES.MEMBERS) {
    const phoneColumn = memberColumn_(sheet, 'phone');
    const phoneTailColumn = memberColumn_(sheet, 'phoneLastDigits');
    sheet.getRange(rowNumber, phoneColumn).setNumberFormat('@');
    sheet.getRange(rowNumber, phoneTailColumn).setNumberFormat('@');
  }
  sheet.getRange(rowNumber, 1, 1, headers.length)
    .setValues([headers.map(function(header) { return obj[header] !== undefined ? obj[header] : ''; })]);
}


function normalizeFacilityMember_(member) {
  const phone = digits_(member.phone);
  const birthDate = normalizeBirthDate_(member.birthDate);
  return {
    name: str_(member.name),
    gender: str_(member.gender),
    birthDate: birthDate,
    phone: phone,
    phoneLastDigits: phone.slice(-4),
    isSeongnamResident: str_(member.isSeongnamResident) === '관내' ? '관내' : '관외',
    privacyConsent: member.privacyConsent === true || str_(member.privacyConsent).toLowerCase() === 'true'
  };
}


function memberRowToObject_(row) {
  const birthDate = normalizeBirthDate_(row.birthDate);
  const phone = normalizeStoredMemberPhone_(row.phone);
  const storedPhoneTail = digits_(row.phoneLastDigits);
  return {
    id: str_(row.memberId),
    name: str_(row.name),
    gender: str_(row.gender),
    age: birthDate ? calcAge_(birthDate) : Number(row.age || 0),
    birthDate: birthDate,
    phone: phone,
    phoneLastDigits: storedPhoneTail
      ? storedPhoneTail.slice(-4).padStart(4, '0')
      : phone.slice(-4),
    status: str_(row.status, 'approved'),
    role: str_(row.role, 'user'),
    isSeongnamResident: str_(row.isSeongnamResident),
    createdAt: str_(row.createdAt),
    lastUsedAt: dateOnly_(row.lastUsedAt),
    retentionHoldUntil: dateOnly_(row.retentionHoldUntil),
    retentionHoldReason: str_(row.retentionHoldReason)
  };
}

function normalizeStoredMemberPhone_(value) {
  const phone = digits_(value);
  if (phone.length === 10 && phone.indexOf('10') === 0) return '0' + phone;
  return phone;
}

function buildParticipantDetail_(members) {
  return (members || []).map(function(item) {
    return str_(item.gender) + '/' + String(item.age || '');
  }).join(', ');
}

function syncNaverReservationCsv() {
  ensureDataSheetSchemas_();
  const folder = getNaverCsvFolder_();
  const fileIterator = folder.getFiles();
  const files = [];
  while (fileIterator.hasNext()) {
    const file = fileIterator.next();
    if (/\.csv$/i.test(file.getName())) files.push(file);
  }
  files.sort(function(a, b) { return a.getLastUpdated().getTime() - b.getLastUpdated().getTime(); });

  const latestRecords = {};
  let csvRowCount = 0;
  let ignoredRowCount = 0;
  let parseErrorCount = 0;
  const parsedFiles = [];

  files.forEach(function(file) {
    try {
      const rows = Utilities.parseCsv(readNaverCsvText_(file));
      const headerIndex = findNaverCsvHeaderIndex_(rows);
      if (headerIndex < 0) throw new Error('네이버 예약 CSV 헤더를 찾을 수 없습니다.');

      const headers = normalizeNaverCsvHeaders_(rows[headerIndex]);
      const fileRecords = {};
      validateNaverCsvHeaders_(headers);
      rows.slice(headerIndex + 1).forEach(function(row) {
        if (!row.some(function(cell) { return str_(cell); })) return;
        csvRowCount += 1;
        const record = mapNaverReservationRow_(headers, row, file);
        if (!record) {
          ignoredRowCount += 1;
          return;
        }
        fileRecords[record.reservationId] = record;
      });
      Object.keys(fileRecords).forEach(function(reservationId) {
        latestRecords[reservationId] = fileRecords[reservationId];
      });
      parsedFiles.push(file);
    } catch (error) {
      parseErrorCount += 1;
      Logger.log('네이버 예약 CSV 읽기 오류 (' + file.getName() + '): ' + (error.message || error));
    }
  });

  const sheet = ensureNaverReservationSchema_();
  const existingMap = buildExistingNaverReservationMap_(sheet);
  const summary = {
    created: 0,
    updated: 0,
    unchanged: 0,
    cancelled: 0,
    noShow: 0
  };

  Object.keys(latestRecords).sort().forEach(function(reservationId) {
    const record = latestRecords[reservationId];
    const result = upsertNaverReservation_(sheet, existingMap[reservationId], record);
    summary[result.changeType] += 1;
    if (isNoShowNaverStatus_(record.status)) summary.noShow += 1;
    else if (isCancelledNaverStatus_(record.status)) summary.cancelled += 1;
  });

  let trashedFileCount = 0;
  parsedFiles.forEach(function(file) {
    file.setTrashed(true);
    trashedFileCount += 1;
  });

  SpreadsheetApp.getUi().alert(
    '네이버 예약 동기화 완료\n\n' +
    'CSV 파일: ' + files.length + '개\n' +
    '확인 행: ' + csvRowCount + '건\n' +
    '신규: ' + summary.created + '건\n' +
    '수정: ' + summary.updated + '건\n' +
    '변경 없음: ' + summary.unchanged + '건\n' +
    '취소: ' + summary.cancelled + '건 / 노쇼: ' + summary.noShow + '건\n' +
    '대상 외 장소: ' + ignoredRowCount + '건\n' +
    'CSV 오류: ' + parseErrorCount + '개\n' +
    'Drive 휴지통 이동: ' + trashedFileCount + '개'
  );
  return summary;
}

function getNaverCsvFolder_() {
  const folderId = str_(PropertiesService.getScriptProperties().getProperty('NAVER_CSV_FOLDER_ID'));
  if (folderId) return DriveApp.getFolderById(folderId);

  const folderName = str_(
    PropertiesService.getScriptProperties().getProperty('NAVER_CSV_FOLDER_NAME'),
    '네이버예약CSV'
  );
  const folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) {
    throw new Error('구글 드라이브에서 ' + folderName + ' 폴더를 찾을 수 없습니다.');
  }
  return folders.next();
}

function readNaverCsvText_(file) {
  const blob = file.getBlob();
  let text = blob.getDataAsString('UTF-8').replace(/^\uFEFF/, '');
  if (text.indexOf('\uFFFD') !== -1) {
    text = blob.getDataAsString('EUC-KR').replace(/^\uFEFF/, '');
  }
  return text;
}

function findNaverCsvHeaderIndex_(rows) {
  return rows.findIndex(function(row) {
    const headers = normalizeNaverCsvHeaders_(row);
    return NAVER_CSV_REQUIRED_HEADERS.every(function(header) {
      return headers.indexOf(header) !== -1;
    });
  });
}

function normalizeNaverCsvHeaders_(headers) {
  return (headers || []).map(function(header) {
    return str_(header).replace(/^\uFEFF/, '');
  });
}

function validateNaverCsvHeaders_(headers) {
  const missing = NAVER_CSV_REQUIRED_HEADERS.filter(function(header) {
    return headers.indexOf(header) === -1;
  });
  if (missing.length > 0) {
    throw new Error('필수 열이 없습니다: ' + missing.join(', '));
  }
}

function mapNaverReservationRow_(headers, row, file) {
  const values = {};
  headers.forEach(function(header, index) { values[str_(header)] = str_(row[index]); });

  const reservationId = pickFirstValue_(values, NAVER_CSV_FIELDS.reservationId);
  const placeName = pickFirstValue_(values, NAVER_CSV_FIELDS.placeName);
  const spaceType = normalizeNaverSpaceType_(placeName);
  if (!reservationId || !spaceType) return null;

  const status = pickFirstValue_(values, NAVER_CSV_FIELDS.status);
  const bookerName = pickFirstValue_(values, NAVER_CSV_FIELDS.bookerName);
  const maskedPhone = pickFirstValue_(values, NAVER_CSV_FIELDS.phone);
  const rawUsageText = pickFirstValue_(values, NAVER_CSV_FIELDS.usageDateTime);
  const usage = parseNaverUsageDateTime_(rawUsageText);
  const headcount = normalizeReservationHeadcount_(pickFirstValue_(values, NAVER_CSV_FIELDS.headcount));
  if (!status || !bookerName || !maskedPhone || !usage.date || !usage.startTime || !usage.endTime) {
    throw new Error('예약번호 ' + reservationId + '의 필수 예약 정보를 해석할 수 없습니다.');
  }
  if (headcount > 11) {
    throw new Error('예약번호 ' + reservationId + '의 이용 인원이 11명을 초과합니다.');
  }

  return {
    reservationId: reservationId,
    status: status,
    bookerName: bookerName,
    maskedPhone: maskedPhone,
    usageDate: usage.date,
    startTime: usage.startTime,
    endTime: usage.endTime,
    placeName: placeName,
    spaceType: spaceType,
    headcount: headcount,
    purpose: pickFirstValue_(values, NAVER_CSV_FIELDS.purpose),
    requestedAt: pickFirstValue_(values, NAVER_CSV_FIELDS.requestedAt),
    confirmedAt: pickFirstValue_(values, NAVER_CSV_FIELDS.confirmedAt),
    completedAt: pickFirstValue_(values, NAVER_CSV_FIELDS.completedAt),
    cancelledAt: pickFirstValue_(values, NAVER_CSV_FIELDS.cancelledAt),
    cancelReason: pickFirstValue_(values, NAVER_CSV_FIELDS.cancelReason),
    lastSyncedAt: nowText_(),
    sourceFileName: file.getName(),
    sourceModifiedAt: Utilities.formatDate(file.getLastUpdated(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    rawUsageText: rawUsageText
  };
}

function pickFirstValue_(map, keys) {
  for (let i = 0; i < keys.length; i += 1) {
    if (str_(map[keys[i]])) return str_(map[keys[i]]);
  }
  return '';
}

function normalizeReservationHeadcount_(value) {
  const match = str_(value).match(/\d+/);
  return match ? Math.max(1, Number(match[0])) : 1;
}

function parseNaverUsageDateTime_(value) {
  const raw = str_(value);
  const dateMatch = raw.match(/(\d{2,4})[.\-/]\s*(\d{1,2})[.\-/]\s*(\d{1,2})/);
  const timeMatch = raw.match(/(오전|오후|AM|PM)?\s*(\d{1,2}):(\d{2})\s*~\s*(오전|오후|AM|PM)?\s*(\d{1,2}):(\d{2})/i);
  if (!dateMatch || !timeMatch) return { date: '', startTime: '', endTime: '' };

  let year = Number(dateMatch[1]);
  if (year < 100) year += 2000;
  const startPeriod = str_(timeMatch[1]).toUpperCase();
  const endPeriod = str_(timeMatch[4]).toUpperCase() || startPeriod;
  const startHour = convertReservationHour_(startPeriod, Number(timeMatch[2]));
  let endHour = convertReservationHour_(endPeriod, Number(timeMatch[5]));
  const startMinute = Number(timeMatch[3]);
  const endMinute = Number(timeMatch[6]);
  if (!timeMatch[4] && (endHour < startHour || (endHour === startHour && endMinute <= startMinute))) {
    endHour += 12;
  }
  if (endHour > 23) return { date: '', startTime: '', endTime: '' };

  return {
    date: year + '-' + pad2_(dateMatch[2]) + '-' + pad2_(dateMatch[3]),
    startTime: pad2_(startHour) + ':' + pad2_(startMinute),
    endTime: pad2_(endHour) + ':' + pad2_(endMinute)
  };
}

function convertReservationHour_(period, hour) {
  const normalized = str_(period).toUpperCase();
  if (normalized === '오전' || normalized === 'AM') return hour === 12 ? 0 : hour;
  if (normalized === '오후' || normalized === 'PM') return hour === 12 ? 12 : hour + 12;
  return hour;
}

function normalizeNaverSpaceType_(placeName) {
  const normalized = normalizeNameKey_(placeName);
  if (normalized.indexOf('커리어존') !== -1) {
    if (normalized.indexOf('room1') !== -1 || normalized.indexOf('룸1') !== -1) return 'careerRoom1';
    if (normalized.indexOf('room2') !== -1 || normalized.indexOf('룸2') !== -1) return 'careerRoom2';
    return 'careerZone';
  }
  if (normalized.indexOf('커넥트룸') !== -1) return 'connectRoom';
  return '';
}

function buildExistingNaverReservationMap_(sheet) {
  if (sheet.getLastRow() < 2) return {};
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.naverSync.length).getValues();
  const map = {};
  rows.forEach(function(row, index) {
    const values = {};
    HEADERS.naverSync.forEach(function(header, column) { values[header] = row[column]; });
    if (str_(values.reservationId)) {
      map[str_(values.reservationId)] = { rowNumber: index + 2, values: values };
    }
  });
  return map;
}

function upsertNaverReservation_(sheet, existing, record) {
  preserveNaverOperationalFields_(record, existing && existing.values);
  const changed = !existing || hasNaverReservationChanged_(existing.values, record);

  if (!existing) {
    sheet.appendRow(HEADERS.naverSync.map(function(header) { return record[header] || ''; }));
    applyNaverReservationRowStyle_(sheet, sheet.getLastRow(), record);
    return { changeType: 'created' };
  }
  if (!changed) {
    return { changeType: 'unchanged' };
  }

  sheet.getRange(existing.rowNumber, 1, 1, HEADERS.naverSync.length)
    .setValues([HEADERS.naverSync.map(function(header) { return record[header] || ''; })]);
  applyNaverReservationRowStyle_(sheet, existing.rowNumber, record);
  return { changeType: 'updated' };
}

function preserveNaverOperationalFields_(record, existing) {
  if (!existing) return;
  [
    'onsiteMemberName', 'onsiteMemberNameKey', 'onsiteCheckedInAt',
    'onsiteCheckinSource', 'usageLogId', 'submittedAt', 'submittedByMemberId'
  ].forEach(function(key) {
    record[key] = existing[key] || '';
  });
}

function hasNaverReservationChanged_(before, after) {
  return [
    'status', 'bookerName', 'maskedPhone', 'usageDate', 'startTime', 'endTime', 'placeName',
    'spaceType', 'headcount', 'purpose', 'requestedAt', 'confirmedAt', 'completedAt',
    'cancelledAt', 'cancelReason'
  ].some(function(key) {
    return str_(before[key]) !== str_(after[key]);
  });
}

function isNoShowNaverStatus_(status) {
  const normalized = normalizeNameKey_(status);
  return normalized.indexOf('노쇼') !== -1 || normalized.indexOf('noshow') !== -1;
}

function isCancelledNaverStatus_(status) {
  const normalized = normalizeNameKey_(status);
  return normalized.indexOf('취소') !== -1 || isNoShowNaverStatus_(status);
}

function getReservationSpaceLabel_(spaceType) {
  if (spaceType === 'careerRoom1') return '커리어존 Room1';
  if (spaceType === 'careerRoom2') return '커리어존 Room2';
  if (spaceType === 'careerZone') return '커리어존';
  return '커넥트룸';
}

function getTodayReservationsForMember_(member, date) {
  const sheet = ensureNaverReservationSchema_();
  return findObjectsByColumnValue_(sheet, HEADERS.naverSync, 'usageDate', date)
    .filter(function(reservation) {
      return !isCancelledNaverStatus_(reservation.status) &&
        reservationBelongsToMember_(reservation, member);
    })
    .sort(function(a, b) {
      return str_(a.startTime).localeCompare(str_(b.startTime));
    })
    .map(publicNaverReservation_);
}

function reservationBelongsToMember_(reservation, member) {
  if (normalizeNameKey_(reservation.bookerName) !== normalizeNameKey_(member.name)) return false;
  const reservationPhoneTail = digits_(reservation.maskedPhone).slice(-4);
  return !reservationPhoneTail || reservationPhoneTail === str_(member.phoneLastDigits);
}

function publicNaverReservation_(reservation) {
  return {
    id: str_(reservation.reservationId),
    spaceType: str_(reservation.spaceType),
    placeName: getReservationSpaceLabel_(reservation.spaceType),
    date: normalizeSheetDate_(reservation.usageDate),
    startTime: normalizeSheetTime_(reservation.startTime),
    endTime: normalizeSheetTime_(reservation.endTime),
    headcount: Math.max(1, Number(reservation.headcount || 1)),
    purpose: str_(reservation.purpose, '네이버 예약 이용'),
    submitted: Boolean(str_(reservation.usageLogId))
  };
}

function findNaverReservationEntry_(sheet, reservationId) {
  const entries = findEntriesByColumnValue_(sheet, HEADERS.naverSync, 'reservationId', reservationId);
  return entries.length > 0 ? entries[0] : null;
}

function updateNaverReservationRow_(sheet, rowNumber, changes) {
  const values = sheet.getRange(rowNumber, 1, 1, HEADERS.naverSync.length).getValues()[0];
  Object.keys(changes).forEach(function(header) {
    const column = HEADERS.naverSync.indexOf(header);
    if (column >= 0) values[column] = changes[header] || '';
  });
  sheet.getRange(rowNumber, 1, 1, HEADERS.naverSync.length).setValues([values]);
}

function findLinkedReservationLog_(reservationId) {
  const sheetNames = [SHEET_NAMES.CAREER_LOGS, SHEET_NAMES.CONNECT_LOGS];
  for (let i = 0; i < sheetNames.length; i += 1) {
    const sheet = getOrCreateSheet_(sheetNames[i], HEADERS.roomLogs);
    const entries = findEntriesByColumnValue_(
      sheet,
      HEADERS.roomLogs,
      'linkedReservationId',
      reservationId
    );
    if (entries.length > 0) {
      return { sheetName: sheetNames[i], rowNumber: entries[0].rowNumber, row: entries[0].values };
    }
  }
  return null;
}

function applyNaverReservationRowStyle_(sheet, rowNumber, record) {
  const cancelled = isCancelledNaverStatus_(record.status);
  const range = sheet.getRange(rowNumber, 1, 1, HEADERS.naverSync.length);
  range.setFontLine(cancelled ? 'line-through' : 'none');
  range.setFontColor(cancelled ? '#777777' : '#000000');
}

function normalizeNameKey_(value) {
  return str_(value).toLowerCase().replace(/\s+/g, '');
}

function normalizeSheetTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'HH:mm');
  }
  const raw = str_(value);
  if (/^\d{1,2}:\d{2}$/.test(raw)) {
    const parts = raw.split(':').map(Number);
    if (parts[0] <= 23 && parts[1] <= 59) return pad2_(parts[0]) + ':' + pad2_(parts[1]);
  }
  return normalizeTime_(raw);
}















function normalizeTime_(value) {
  const d = digits_(value);
  if (d.length < 3 || d.length > 4) return '';
  const padded = d.padStart(4, '0');
  const hour = Number(padded.substring(0, 2));
  const minute = Number(padded.substring(2, 4));
  if (hour > 23 || minute > 59) return '';
  return pad2_(hour) + ':' + pad2_(minute);
}

function normalizeCompanions_(value, maxCount) {
  const companions = Array.isArray(value) ? value : [];
  const limit = Number.isInteger(maxCount) && maxCount >= 0 ? maxCount : 10;
  if (companions.length > limit) throw new Error('추가 인원수를 확인해 주세요.');
  return companions.map(function(person) {
    const gender = str_(person && person.gender);
    const age = Number(person && person.age);
    if (!isValidGender_(gender) || !Number.isInteger(age) || age < 1 || age > 100) {
      throw new Error('추가 인원의 성별과 나이를 확인해 주세요.');
    }
    return { gender: gender, age: age };
  });
}

function normalizeVisitSource_(value) {
  const source = str_(value);
  const allowed = ['facility-qr', 'facility-phone-recheck', 'facility-signup', 'facility-signup-existing', 'facility-birthdate-update'];
  return allowed.indexOf(source) >= 0 ? source : 'facility-qr';
}

function isValidGender_(value) {
  return value === '남성' || value === '여성';
}

function isValidPhone_(value) {
  return /^010\d{8}$/.test(digits_(value));
}

function calcAge_(birthDate) {
  const birth = str_(birthDate).split('-').map(Number);
  const today = today_().split('-').map(Number);
  if (birth.length !== 3 || birth.some(function(value) { return !Number.isFinite(value); })) return 0;
  let age = today[0] - birth[0];
  if (today[1] < birth[1] || (today[1] === birth[1] && today[2] < birth[2])) age -= 1;
  return age;
}

function normalizeBirthDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  }
  const raw = str_(value);
  const digits = digits_(raw);
  if (digits.length === 8) {
    return digits.substring(0, 4) + '-' + digits.substring(4, 6) + '-' + digits.substring(6, 8);
  }
  return raw;
}

function isValidBirthDate_(value) {
  const normalized = normalizeBirthDate_(value);
  const match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return false;

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const birth = new Date(year, month - 1, day);
  return year >= 1900 &&
    birth.getFullYear() === year &&
    birth.getMonth() === month - 1 &&
    birth.getDate() === day &&
    normalized <= today_();
}

function json_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function webResponse_(payload, callback) {
  if (callback && /^[A-Za-z_$][\w$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(payload);
}

function withScriptLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function str_(value, fallback) {
  const result = value === undefined || value === null ? '' : String(value).trim();
  return result || (fallback || '');
}

function digits_(value) {
  return str_(value).replace(/\D/g, '');
}

function pad2_(value) {
  return String(value).padStart(2, '0');
}

function dateOnly_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  }
  const text = str_(value);
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return '';
  return match[1] + '-' + match[2] + '-' + match[3];
}


function today_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

function nowText_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}
