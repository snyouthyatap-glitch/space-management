const TZ = 'Asia/Seoul';
const QR_ENTRY_TOKEN_PROPERTY = 'QR_ENTRY_TOKEN';
const QR_ENTRY_PAGE_URL = 'https://snyouthyatap-glitch.github.io/space-management/';
const PHONE_TAIL_LOOKUP_LIMIT = 60;
const PHONE_TAIL_LOOKUP_WINDOW_SECONDS = 300;
const MEMBER_RETENTION_YEARS = 3;
const FACILITY_MIN_AGE = 19;
const FACILITY_MAX_AGE = 39;
const ADMIN_ACCESS_HASH_PROPERTY = 'ADMIN_ACCESS_HASH';
const ADMIN_SESSION_SECONDS = 1800;
const PENDING_CALENDAR_MONTHS_PROPERTY = 'PENDING_OPERATION_CALENDAR_MONTHS';

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
  RESERVATION_REQUESTS: 'APP_RESERVATION_REQUESTS',
  RETENTION_LOGS: 'APP_RETENTION_LOGS',
  ACCESS_ALERTS: 'APP_ACCESS_ALERTS',
  NAVER_SYNC: '네이버예약_동기화'
};

const HEADERS = {
  visit: ['연번', '날짜', '성별', '나이'],
  printer: ['연번', '날짜', '성별', '나이', '사용 매수'],
  career: [
    '연번', '날짜', '사용공간', '사용목적', '인원수', '비고', '시작시간', '종료시간',
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
  reservationRequests: [
    'requestId', 'status', 'memberId', 'name', 'memberNameKey', 'gender', 'age', 'phone', 'phoneLastDigits',
    'spaceType', 'placeId', 'placeName', 'date', 'startTime', 'endTime', 'purpose', 'headcount',
    'companionsJson', 'companionCount', 'requestedAt', 'decidedAt', 'decidedBy', 'decisionNote',
    'usageLogId', 'submittedAt', 'submittedByMemberId', 'updatedAt'
  ],
  retentionLogs: [
    'runId', 'action', 'processedAt', 'memberCount', 'visitLogCount', 'printerLogCount',
    'careerLogCount', 'connectLogCount', 'reservationLinkCount', 'accessAlertCount'
  ],
  accessAlerts: [
    'alertId', 'userId', 'name', 'phoneLastDigits', 'age', 'reason', 'date', 'attemptedAt', 'inputType'
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
  // 대리예약이면 실제 이용자인 방문자 정보를 먼저 쓴다.
  bookerName: ['방문자명', '예약자', '예약자명'],
  phone: ['방문자전화번호', '전화번호', '휴대폰번호'],
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
    .addItem('이번 달 운영 달력 갱신', 'generateMonthlyOperationCalendar')
    .addItem('선택 월 운영 달력 갱신', 'refreshSelectedOperationCalendar')
    .addSeparator()
    .addItem('마지막 이용일 다시 계산', 'backfillMemberLastUsedAt')
    .addItem('3년 초과 회원 미리보기', 'previewExpiredMembers')
    .addItem('3년 초과 개인정보 파기', 'purgeExpiredMemberData')
    .addItem('월간 개인정보 파기 트리거 설정', 'setupRetentionCleanupTrigger')
    .addSeparator()
    .addItem('관리자 비밀번호 설정', 'setupAdminAccessKey')
    .addItem('현장 QR 토큰 새로 발급', 'setupQrEntryToken')
    .addToUi();
  refreshPendingOperationCalendars_();
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
  if (action === 'adminChallenge') return handleAdminChallenge_();
  if (action === 'adminLogin') return handleAdminLogin_(payload);
  if (action.indexOf('admin') === 0) {
    requireAdminSession_(payload, action !== 'adminDashboardSnapshot');
  } else {
    requireEntryToken_(payload);
  }

  switch (action) {
    case 'warmup':
      return { ok: true };
    case 'adminLookupMember':
      return handleAdminLookupMember_(payload);
    case 'adminCompleteReservationUsage':
      return handleAdminCompleteReservationUsage_(payload);
    case 'adminDashboardSnapshot':
      return handleAdminDashboardSnapshot_(payload);
    case 'adminUpdateReservationStatus':
      return handleAdminUpdateReservationStatus_(payload);
    case 'adminEnsureNaverReservation':
      return handleAdminEnsureNaverReservation_(payload);
    case 'adminUpdateReservationRequestStatus':
      return handleAdminUpdateReservationRequestStatus_(payload);
    case 'adminMonthlyPerformance':
      return handleAdminMonthlyPerformance_(payload);
    case 'resolveFacilityEntry':
      return handleResolveFacilityEntry_(payload);
    case 'updateMemberBirthDate':
      return handleUpdateMemberBirthDate_(payload);
    case 'registerFacilityMember':
      return handleRegisterFacilityMember_(payload);
    case 'submitVisitLog':
      return handleSubmitVisitLog_(payload);
    case 'getTodayReservations':
      return handleGetTodayReservations_(payload);
    case 'getSpaceAvailability':
      return handleGetSpaceAvailability_(payload);
    case 'submitSpaceReservationRequest':
      return handleSubmitSpaceReservationRequest_(payload);
    case 'submitUsageRecord':
      return handleSubmitUsageRecord_(payload);
    case 'submitReservationUsage':
      return handleSubmitReservationUsage_(payload);
    default:
      throw new Error('지원하지 않는 요청입니다: ' + action);
  }
}

function setupAdminAccessKey() {
  const ui = SpreadsheetApp.getUi();
  const first = ui.prompt(
    '직원 프로그램 관리자 비밀번호 설정',
    '8자 이상의 관리자 비밀번호를 입력해 주세요. 직원 프로그램 연결 설정에 같은 값을 넣습니다.',
    ui.ButtonSet.OK_CANCEL
  );
  if (first.getSelectedButton() !== ui.Button.OK) return;

  const accessKey = str_(first.getResponseText());
  if (accessKey.length < 8) {
    ui.alert('관리자 비밀번호는 8자 이상이어야 합니다.');
    return;
  }

  const second = ui.prompt(
    '직원 프로그램 관리자 비밀번호 확인',
    '같은 비밀번호를 한 번 더 입력해 주세요.',
    ui.ButtonSet.OK_CANCEL
  );
  if (second.getSelectedButton() !== ui.Button.OK) return;
  if (str_(second.getResponseText()) !== accessKey) {
    ui.alert('입력한 비밀번호가 서로 다릅니다.');
    return;
  }

  PropertiesService.getScriptProperties()
    .setProperty(ADMIN_ACCESS_HASH_PROPERTY, sha256Hex_(accessKey));
  ui.alert('관리자 비밀번호를 설정했습니다.');
}

function handleAdminChallenge_() {
  const challenge = Utilities.getUuid();
  CacheService.getScriptCache().put('admin-challenge:' + challenge, '1', 120);
  return { ok: true, challenge: challenge };
}

function handleAdminLogin_(payload) {
  const challenge = str_(payload.challenge);
  const proof = str_(payload.proof).toLowerCase();
  const cache = CacheService.getScriptCache();
  const challengeKey = 'admin-challenge:' + challenge;
  const challengeExists = challenge && cache.get(challengeKey);
  cache.remove(challengeKey);

  const storedHash = str_(
    PropertiesService.getScriptProperties().getProperty(ADMIN_ACCESS_HASH_PROPERTY)
  ).toLowerCase();
  if (!storedHash) {
    throw new Error('구글 시트 관리 메뉴에서 관리자 비밀번호를 먼저 설정해 주세요.');
  }
  if (!challengeExists || !proof || proof !== sha256Hex_(challenge + ':' + storedHash)) {
    throw new Error('관리자 비밀번호가 올바르지 않습니다.');
  }

  const sessionToken = Utilities.getUuid();
  cache.put('admin-session:' + sessionToken, '1', ADMIN_SESSION_SECONDS);
  return {
    ok: true,
    sessionToken: sessionToken,
    expiresInSeconds: ADMIN_SESSION_SECONDS
  };
}

function requireAdminSession_(payload, refreshSession) {
  const sessionToken = str_(payload && payload.adminSession);
  const cache = CacheService.getScriptCache();
  const key = 'admin-session:' + sessionToken;
  if (!sessionToken || !cache.get(key)) {
    throw new Error('관리자 인증이 만료되었습니다. 다시 로그인해 주세요.');
  }
  if (refreshSession !== false) cache.put(key, '1', ADMIN_SESSION_SECONDS);
}

function handleAdminDashboardSnapshot_(payload) {
  const recentLimit = Math.min(500, Math.max(20, Number(payload && payload.limit) || 100));
  const calendarEntries = getStaffCalendarEntries_(payload && payload.calendarMonth);
  return {
    ok: true,
    serverTime: nowText_(),
    events: getStaffDashboardEvents_(recentLimit),
    todayReservations: getStaffTodayReservations_(calendarEntries),
    calendarEntries: calendarEntries,
    checkIns: getStaffMonthCheckIns_(payload && payload.calendarMonth),
    reservationRequests: getStaffReservationRequests_(calendarEntries),
    summary: getStaffDashboardSummary_()
  };
}

// 직원 프로그램이 네이버 화면에서 읽은 오늘 예약을 '네이버 예약' 시트에 맞춘다.
// - 확정·승인·이용완료: 시트에 없으면 한 줄 추가, 있으면 바뀐 내용(날짜·시간·공간·인원 등)만 고침
// - 취소: 시트에 있으면 '취소'로 바꿈
// - 이미 일지가 제출됐거나 직원이 처리(이용완료·취소·노쇼)한 예약은 건드리지 않음
// CSV 동기화를 깜빡해도 이용자 웹 '오늘 예약 일정'과 운영 캘린더에 오늘 예약이 보이게 하려는 것이다.
function handleAdminEnsureNaverReservation_(payload) {
  const reservationId = str_(payload && payload.reservationId).replace(/^naverSmartPlace:/, '');
  const status = str_(payload && payload.status, '확정');
  if (!reservationId) throw new Error('맞출 네이버 예약번호가 없습니다.');
  const cancelled = status === '취소';
  if (!cancelled && ['확정', '승인', '이용완료'].indexOf(status) === -1) {
    throw new Error('확정 또는 취소된 네이버 예약만 시트에 맞출 수 있습니다.');
  }

  return withScriptLock_(function() {
    const sheet = ensureNaverReservationSchema_();
    const existing = findNaverReservationEntry_(sheet, reservationId);
    const now = nowText_();

    if (existing) {
      const current = str_(existing.values.status);
      const alreadyHandled = isCancelledNaverStatus_(current) ||
        current === '이용완료' ||
        str_(existing.values.submittedAt) ||
        str_(existing.values.usageLogId);
      if (alreadyHandled) {
        return { ok: true, created: false, updated: false, reservationId: reservationId };
      }
      if (!cancelled) {
        const detailChanges = diffNaverReservationDetails_(existing.values, payload, status);
        const changedFields = Object.keys(detailChanges);
        if (changedFields.length === 0) {
          return { ok: true, created: false, updated: false, reservationId: reservationId };
        }
        detailChanges.lastSyncedAt = now;
        updateNaverReservationRow_(sheet, existing.rowNumber, detailChanges);
        applyNaverReservationRowStyle_(sheet, existing.rowNumber, Object.assign({}, existing.values, detailChanges));
        clearNaverReservationMonthCache_(normalizeSheetDate_(existing.values.usageDate));
        if (detailChanges.usageDate) clearNaverReservationMonthCache_(detailChanges.usageDate);
        return {
          ok: true,
          created: false,
          updated: true,
          changedFields: changedFields,
          reservationId: reservationId
        };
      }
      const changes = {
        status: '취소',
        completedAt: '',
        cancelledAt: now,
        cancelReason: '네이버에서 취소됨(직원 프로그램 확인)',
        lastSyncedAt: now
      };
      updateNaverReservationRow_(sheet, existing.rowNumber, changes);
      applyNaverReservationRowStyle_(sheet, existing.rowNumber, Object.assign({}, existing.values, changes));
      clearNaverReservationMonthCache_(normalizeSheetDate_(existing.values.usageDate));
      return { ok: true, created: false, updated: true, reservationId: reservationId };
    }

    // 시트에 없는 취소 건은 남길 필요가 없다.
    if (cancelled) return { ok: true, created: false, updated: false, reservationId: reservationId };

    const record = {
      reservationId: reservationId,
      status: status,
      bookerName: cleanNaverBookerName_(payload.bookerName) || '예약자',
      maskedPhone: maskNaverPhone_(payload.maskedPhone),
      usageDate: normalizeSheetDate_(payload.usageDate),
      startTime: normalizeSheetTime_(payload.startTime),
      endTime: normalizeSheetTime_(payload.endTime),
      placeName: str_(payload.placeName),
      spaceType: str_(payload.spaceType),
      headcount: safePositiveInteger_(payload.headcount, 1),
      purpose: str_(payload.purpose, '네이버 예약 이용'),
      confirmedAt: now,
      lastSyncedAt: now,
      sourceFileName: '직원 프로그램 네이버 확인',
      rawUsageText: str_(payload.rawUsageText)
    };
    if (!record.usageDate || !record.startTime || !record.endTime || !record.spaceType) {
      throw new Error('네이버 예약 ' + reservationId + '의 날짜·시간·공간 정보를 확인할 수 없습니다.');
    }

    sheet.appendRow(HEADERS.naverSync.map(function(header) { return record[header] || ''; }));
    applyNaverReservationRowStyle_(sheet, sheet.getLastRow(), record);
    clearNaverReservationMonthCache_(record.usageDate);
    return { ok: true, created: true, updated: false, reservationId: reservationId };
  });
}

// 네이버에서 바뀐(목록에 '변경' 표시) 예약의 새 값 중 시트와 다른 것만 돌려준다.
// 빈 값은 '모름'으로 보고 덮어쓰지 않는다. 이름은 띄어쓰기, 연락처는 뒷 4자리만 비교한다.
function diffNaverReservationDetails_(current, payload, status) {
  const incoming = {
    usageDate: normalizeSheetDate_(payload.usageDate),
    startTime: normalizeSheetTime_(payload.startTime),
    endTime: normalizeSheetTime_(payload.endTime),
    placeName: str_(payload.placeName),
    spaceType: str_(payload.spaceType),
    headcount: payload.headcount ? String(safePositiveInteger_(payload.headcount, 1)) : '',
    purpose: str_(payload.purpose),
    bookerName: cleanNaverBookerName_(payload.bookerName),
    maskedPhone: maskNaverPhone_(payload.maskedPhone),
    rawUsageText: str_(payload.rawUsageText)
  };
  const changes = {};
  Object.keys(incoming).forEach(function(field) {
    const next = incoming[field];
    if (!next) return;
    let before = current[field];
    if (field === 'usageDate') before = normalizeSheetDate_(before);
    else if (field === 'startTime' || field === 'endTime') before = normalizeSheetTime_(before);
    else before = str_(before);
    if (field === 'bookerName' && normalizeNameKey_(before) === normalizeNameKey_(next)) return;
    if (field === 'maskedPhone' && digits_(before).slice(-4) === digits_(next).slice(-4)) return;
    if (field === 'headcount' && Number(before) === Number(next)) return;
    // 상품명은 CSV('AI 커리어존 (청취Room1)')와 화면('커리어존 Room1') 표기가 달라서, 공간이 바뀐 경우에만 고친다.
    if (field === 'placeName' && str_(current.spaceType) === incoming.spaceType) return;
    if (before !== next) changes[field] = next;
  });
  if (str_(current.status) === '신청' && status !== '신청') {
    changes.status = status;
    changes.confirmedAt = str_(current.confirmedAt, nowText_());
  }
  // rawUsageText만 다른 건 표시 차이일 뿐이라 변경으로 보지 않는다.
  if (Object.keys(changes).length === 1 && changes.rawUsageText) return {};
  return changes;
}

function clearNaverReservationMonthCache_(usageDate) {
  if (!usageDate) return;
  markOperationCalendarMonthPending_(usageDate);
  const monthKey = usageDate.slice(0, 7);
  if (isValidMonthKey_(monthKey)) {
    CacheService.getScriptCache().remove('staff-dashboard-calendar-v1:' + monthKey);
  }
}

// CSV와 같은 형태로 가운데 자리를 가린다. 회원 확인은 뒷 4자리만 쓴다.
function maskNaverPhone_(value) {
  const digits = digits_(value);
  if (digits.length < 8) return str_(value);
  return digits.slice(0, 3) + '-****-' + digits.slice(-4);
}

function handleAdminUpdateReservationStatus_(payload) {
  const reservationId = str_(payload && payload.reservationId);
  const status = str_(payload && payload.status);
  const allowedStatuses = ['확정', '이용완료', '취소', '노쇼'];
  if (!reservationId) throw new Error('변경할 예약번호가 없습니다.');
  if (allowedStatuses.indexOf(status) === -1) {
    throw new Error('지원하지 않는 예약 상태입니다.');
  }

  return withScriptLock_(function() {
    const sheet = ensureNaverReservationSchema_();
    const entry = findNaverReservationEntry_(sheet, reservationId);
    if (!entry) throw new Error('예약번호 ' + reservationId + '을(를) 찾을 수 없습니다.');

    const now = nowText_();
    const changes = { status: status };
    if (status === '취소' || status === '노쇼') {
      changes.completedAt = '';
      changes.cancelledAt = now;
      changes.cancelReason = status === '노쇼'
        ? '직원 프로그램에서 노쇼 처리'
        : '직원 프로그램에서 취소 처리';
    } else {
      changes.cancelledAt = '';
      changes.cancelReason = '';
      if (status === '이용완료') {
        changes.completedAt = str_(entry.values.completedAt, now);
      } else {
        changes.completedAt = '';
        changes.confirmedAt = str_(entry.values.confirmedAt, now);
      }
    }

    updateNaverReservationRow_(sheet, entry.rowNumber, changes);
    applyNaverReservationRowStyle_(
      sheet,
      entry.rowNumber,
      Object.assign({}, entry.values, changes)
    );

    const usageDate = normalizeSheetDate_(entry.values.usageDate);
    markOperationCalendarMonthPending_(usageDate);
    const monthKey = usageDate.slice(0, 7);
    if (isValidMonthKey_(monthKey)) {
      CacheService.getScriptCache().remove('staff-dashboard-calendar-v1:' + monthKey);
    }
    return {
      ok: true,
      reservationId: reservationId,
      status: status,
      cancelled: isCancelledNaverStatus_(status)
    };
  });
}

function handleAdminUpdateReservationRequestStatus_(payload) {
  const requestId = str_(
    payload && (payload.reservationRequestId || payload.requestId || payload.reservationId)
  ).replace(/^(web|reservationRequest|webReservation):/, '');
  const status = str_(payload && payload.status);
  const note = str_(payload && payload.note);
  const allowedStatuses = ['신청', '승인', '반려', '취소'];
  if (!requestId) throw new Error('변경할 예약 신청번호가 없습니다.');
  if (allowedStatuses.indexOf(status) === -1) {
    throw new Error('지원하지 않는 예약 신청 상태입니다.');
  }

  return withScriptLock_(function() {
    const sheet = getOrCreateSheet_(SHEET_NAMES.RESERVATION_REQUESTS, HEADERS.reservationRequests);
    ensureSheetHeaders_(sheet, HEADERS.reservationRequests);
    const entry = findReservationRequestEntry_(sheet, requestId);
    if (!entry) throw new Error('예약 신청 ' + requestId + '을(를) 찾을 수 없습니다.');
    if (str_(entry.values.usageLogId)) {
      throw new Error('이미 이용일지가 제출된 예약은 상태를 변경할 수 없습니다.');
    }

    if (status === '승인' || status === '신청') {
      assertNoReservationRequestConflict_(
        str_(entry.values.spaceType),
        normalizeSheetDate_(entry.values.date),
        normalizeSheetTime_(entry.values.startTime),
        normalizeSheetTime_(entry.values.endTime),
        requestId
      );
    }

    const now = nowText_();
    updateReservationRequestRow_(sheet, entry.rowNumber, {
      status: status,
      decidedAt: status === '신청' ? '' : now,
      decidedBy: 'admin',
      decisionNote: note,
      updatedAt: now
    });
    const date = normalizeSheetDate_(entry.values.date);
    clearReservationCaches_(date);
    markOperationCalendarMonthPending_(date);
    return {
      ok: true,
      requestId: 'web:' + requestId,
      reservationId: 'web:' + requestId,
      status: status,
      cancelled: status === '취소' || status === '반려'
    };
  });
}

function handleAdminMonthlyPerformance_(payload) {
  const monthKey = str_(payload && payload.month);
  if (!isValidMonthKey_(monthKey)) {
    throw new Error('실적을 조회할 월을 YYYY-MM 형식으로 입력해 주세요.');
  }

  const categories = {
    visit: createPerformanceCategory_('visit', '방문자', false),
    connect: createPerformanceCategory_('connect', '커넥트룸', true),
    career: createPerformanceCategory_('career', '커리어존', false),
    printer: createPerformanceCategory_('printer', '프린터', false)
  };
  const totalUniqueKeys = {};
  const memberBirthDates = getPerformanceMemberBirthDates_();
  const visitRows = getRowsAsObjectsForMonth_(
    SHEET_NAMES.VISIT_LOGS,
    HEADERS.visitLogs,
    'date',
    monthKey
  );
  const printerRows = getRowsAsObjectsForMonth_(
    SHEET_NAMES.PRINTER_LOGS,
    HEADERS.printerLogs,
    'date',
    monthKey
  );
  const connectRows = getRowsAsObjectsForMonth_(
    SHEET_NAMES.CONNECT_LOGS,
    HEADERS.roomLogs,
    'date',
    monthKey
  );
  const careerRows = getRowsAsObjectsForMonth_(
    SHEET_NAMES.CAREER_LOGS,
    HEADERS.roomLogs,
    'date',
    monthKey
  );

  visitRows.forEach(function(row) {
    addPerformancePerson_(
      categories.visit,
      normalizeSheetDate_(row.date),
      performancePersonKey_(row.userId, 'visit:' + str_(row.logId)),
      row.gender,
      performanceAgeForDate_(row.userId, row.age, row.date, memberBirthDates),
      totalUniqueKeys
    );
  });

  printerRows.forEach(function(row) {
    addPerformancePerson_(
      categories.printer,
      normalizeSheetDate_(row.date),
      performancePersonKey_(row.userId, 'printer:' + str_(row.logId)),
      row.gender,
      performanceAgeForDate_(row.userId, row.age, row.date, memberBirthDates),
      totalUniqueKeys
    );
  });

  [
    { rows: connectRows, category: categories.connect },
    { rows: careerRows, category: categories.career }
  ].forEach(function(source) {
    source.rows.forEach(function(row) {
      addPerformanceRoomLog_(source.category, row, totalUniqueKeys, memberBirthDates);
    });
  });

  const order = ['visit', 'connect', 'career', 'printer'];
  const summary = order.map(function(key) {
    const category = categories[key];
    return {
      key: category.key,
      label: category.label,
      uniqueUsers: Object.keys(category.uniqueKeys).length,
      totalUsers: category.totalUsers
    };
  });
  const totalUsers = summary.reduce(function(sum, item) {
    return sum + item.totalUsers;
  }, 0);

  const daily = {};
  order.forEach(function(key) {
    daily[key] = buildPerformanceDailyRows_(categories[key], monthKey);
  });
  return {
    ok: true,
    month: monthKey,
    generatedAt: nowText_(),
    summary: summary,
    totalUsers: totalUsers,
    totalUniqueUsers: Object.keys(totalUniqueKeys).length,
    daily: daily,
    simplified: {
      visit: buildPerformanceVisitSimplifiedRows_(visitRows, memberBirthDates),
      printer: buildPerformancePrinterSimplifiedRows_(printerRows, memberBirthDates),
      career: buildPerformanceRoomSimplifiedRows_(careerRows, true),
      connect: buildPerformanceRoomSimplifiedRows_(connectRows, false)
    },
    uniqueBasis: '회원 ID 월간 중복 제거, 식별 불가 비회원 및 동반자는 기록별 인원'
  };
}

function buildPerformanceVisitSimplifiedRows_(rows, memberBirthDates) {
  return sortPerformanceSimplifiedRows_(rows).map(function(row) {
    return {
      date: normalizeSheetDate_(row.date),
      gender: str_(row.gender),
      age: performanceAgeForDate_(row.userId, row.age, row.date, memberBirthDates)
    };
  });
}

function buildPerformancePrinterSimplifiedRows_(rows, memberBirthDates) {
  return sortPerformanceSimplifiedRows_(rows).map(function(row) {
    return {
      date: normalizeSheetDate_(row.date),
      gender: str_(row.gender),
      age: performanceAgeForDate_(row.userId, row.age, row.date, memberBirthDates),
      count: performanceCount_(row.count)
    };
  });
}

function buildPerformanceRoomSimplifiedRows_(rows, includeSpace) {
  return sortPerformanceSimplifiedRows_(rows).map(function(row) {
    return {
      date: normalizeSheetDate_(row.date),
      space: includeSpace ? performanceCareerSpaceLabel_(row.spaceType) : '',
      purpose: str_(row.purpose),
      headcount: Math.max(1, performanceCount_(row.headcount)),
      note: str_(row.participants),
      startTime: normalizeSheetTime_(row.startTime),
      endTime: normalizeSheetTime_(row.endTime),
      male20s: performanceCount_(row.male20s),
      female20s: performanceCount_(row.female20s),
      male30s: performanceCount_(row.male30s),
      female30s: performanceCount_(row.female30s),
      maleUnder19: performanceCount_(row.maleUnder19),
      femaleUnder19: performanceCount_(row.femaleUnder19),
      male40Plus: performanceCount_(row.male40Plus),
      female40Plus: performanceCount_(row.female40Plus)
    };
  });
}

function performanceCareerSpaceLabel_(spaceType) {
  if (spaceType === 'careerRoom1') return 'Room1';
  if (spaceType === 'careerRoom2') return 'Room2';
  return '미지정';
}

function sortPerformanceSimplifiedRows_(rows) {
  return rows.slice().sort(function(left, right) {
    return normalizeSheetDate_(left.date).localeCompare(normalizeSheetDate_(right.date)) ||
      str_(left.createdAt).localeCompare(str_(right.createdAt)) ||
      Number(left.__rowNumber || 0) - Number(right.__rowNumber || 0);
  });
}

function createPerformanceCategory_(key, label, includeAllAges) {
  return {
    key: key,
    label: label,
    includeAllAges: includeAllAges === true,
    totalUsers: 0,
    uniqueKeys: {},
    daily: {}
  };
}

function createPerformanceCounts_() {
  return {
    maleUnder19: 0,
    femaleUnder19: 0,
    male20To24: 0,
    female20To24: 0,
    male25To29: 0,
    female25To29: 0,
    male30To34: 0,
    female30To34: 0,
    male35To39: 0,
    female35To39: 0,
    male40Plus: 0,
    female40Plus: 0,
    unknown: 0,
    total: 0
  };
}

function getPerformanceDay_(category, date) {
  if (!category.daily[date]) category.daily[date] = createPerformanceCounts_();
  return category.daily[date];
}

function addPerformancePerson_(category, date, personKey, gender, age, totalUniqueKeys) {
  if (!isPerformanceAgeIncluded_(category, age)) return;
  const day = getPerformanceDay_(category, date);
  const bucket = performanceBucket_(gender, age);
  day[bucket] += 1;
  day.total += 1;
  category.totalUsers += 1;
  category.uniqueKeys[personKey] = true;
  totalUniqueKeys[personKey] = true;
}

function addPerformanceRoomLog_(category, row, totalUniqueKeys, memberBirthDates) {
  const date = normalizeSheetDate_(row.date);
  const people = performanceRoomPeople_(row);
  if (people.length > 0) {
    people[0].age = performanceAgeForDate_(row.userId, people[0].age, date, memberBirthDates);
  }
  const headcount = Math.max(
    1,
    people.length,
    performanceCount_(row.headcount),
    performanceCount_(row.companionCount) + 1
  );
  const logId = str_(row.logId, category.key + '-row-' + str_(row.__rowNumber));
  people.slice(0, headcount).forEach(function(person, index) {
    if (!isPerformanceAgeIncluded_(category, person.age)) return;
    const personKey = index === 0
      ? performancePersonKey_(row.userId, category.key + ':' + logId + ':lead')
      : 'companion:' + logId + ':' + index;
    addPerformancePerson_(category, date, personKey, person.gender, person.age, totalUniqueKeys);
  });

  if (category.includeAllAges && people.length < headcount) {
    const day = getPerformanceDay_(category, date);
    const missingCount = headcount - people.length;
    day.unknown += missingCount;
    day.total += missingCount;
    category.totalUsers += missingCount;
    for (let index = people.length; index < headcount; index += 1) {
      const personKey = index === 0
        ? performancePersonKey_(row.userId, category.key + ':' + logId + ':lead')
        : 'companion:' + logId + ':' + index;
      category.uniqueKeys[personKey] = true;
      totalUniqueKeys[personKey] = true;
    }
  }
}

function isPerformanceAgeIncluded_(category, age) {
  if (category.includeAllAges) return true;
  const numericAge = Number(age);
  return Number.isFinite(numericAge) && numericAge >= FACILITY_MIN_AGE && numericAge <= FACILITY_MAX_AGE;
}

function getPerformanceMemberBirthDates_() {
  const sheet = getOrCreateSheet_(SHEET_NAMES.MEMBERS, HEADERS.members);
  if (sheet.getLastRow() < 2) return {};
  const headerContext = memberHeaderContext_(sheet);
  const idColumn = memberColumn_(sheet, 'memberId', headerContext);
  const birthDateColumn = memberColumn_(sheet, 'birthDate', headerContext);
  const rowCount = sheet.getLastRow() - 1;
  const ids = sheet.getRange(2, idColumn, rowCount, 1).getValues();
  const birthDates = sheet.getRange(2, birthDateColumn, rowCount, 1).getValues();
  const result = {};
  ids.forEach(function(row, index) {
    const memberId = str_(row[0]);
    const birthDate = normalizeBirthDate_(birthDates[index][0]);
    if (memberId && isValidBirthDate_(birthDate)) result[memberId] = birthDate;
  });
  return result;
}

function performanceAgeForDate_(userId, storedAge, usageDate, memberBirthDates) {
  const birthDate = (memberBirthDates || {})[str_(userId)];
  return birthDate ? calcAgeOnDate_(birthDate, normalizeSheetDate_(usageDate)) : Number(storedAge || 0);
}

function performanceRoomPeople_(row) {
  const participantText = str_(row.participants);
  if (participantText) {
    const parsed = participantText.split(',').map(function(item) {
      const match = str_(item).match(/^(남성|여성)\s*\/\s*(\d{1,3})$/);
      return match ? { gender: match[1], age: Number(match[2]) } : null;
    }).filter(function(person) { return person; });
    if (parsed.length > 0) return parsed;
  }

  const people = [{ gender: row.gender, age: row.age }];
  let companions = [];
  try {
    companions = JSON.parse(str_(row.companionsJson, '[]'));
    if (!Array.isArray(companions)) companions = [];
  } catch (error) {}
  return people.concat(companions);
}

function performanceCount_(value) {
  const number = Number(value || 0);
  return isFinite(number) ? Math.max(0, Math.floor(number)) : 0;
}

function performancePersonKey_(userId, fallback) {
  const id = str_(userId);
  return id ? 'user:' + id : 'record:' + fallback;
}

function performanceBucket_(gender, age) {
  const normalizedGender = str_(gender);
  const numericAge = Number(age || 0);
  if ((normalizedGender !== '남성' && normalizedGender !== '여성') || numericAge < 1) {
    return 'unknown';
  }
  const prefix = normalizedGender === '남성' ? 'male' : 'female';
  if (numericAge <= 19) return prefix + 'Under19';
  if (numericAge <= 24) return prefix + '20To24';
  if (numericAge <= 29) return prefix + '25To29';
  if (numericAge <= 34) return prefix + '30To34';
  if (numericAge <= 39) return prefix + '35To39';
  return prefix + '40Plus';
}

function buildPerformanceDailyRows_(category, monthKey) {
  const parts = monthKey.split('-').map(Number);
  const lastDay = new Date(parts[0], parts[1], 0).getDate();
  const rows = [];
  for (let day = 1; day <= lastDay; day += 1) {
    const date = monthKey + '-' + pad2_(day);
    const counts = category.daily[date] || createPerformanceCounts_();
    if (counts.total > 0) rows.push(Object.assign({ date: date }, counts));
  }
  return rows;
}

function getRowsAsObjectsForMonth_(sheetName, headers, dateHeader, monthKey) {
  const sheet = getOrCreateSheet_(sheetName, headers);
  if (sheet.getLastRow() < 2) return [];
  const dateColumn = headers.indexOf(dateHeader) + 1;
  if (dateColumn < 1) throw new Error(sheetName + ' 시트의 날짜 열을 찾을 수 없습니다.');

  const dateValues = sheet
    .getRange(2, dateColumn, sheet.getLastRow() - 1, 1)
    .getValues();
  const rowNumbers = [];
  dateValues.forEach(function(row, index) {
    if (normalizeSheetDate_(row[0]).slice(0, 7) === monthKey) {
      rowNumbers.push(index + 2);
    }
  });
  if (rowNumbers.length === 0) return [];

  const groups = [];
  rowNumbers.forEach(function(rowNumber) {
    const last = groups[groups.length - 1];
    if (last && last.end + 1 === rowNumber) {
      last.end = rowNumber;
    } else {
      groups.push({ start: rowNumber, end: rowNumber });
    }
  });

  const result = [];
  groups.forEach(function(group) {
    const values = sheet
      .getRange(group.start, 1, group.end - group.start + 1, headers.length)
      .getValues();
    values.forEach(function(row, rowIndex) {
      const object = {};
      headers.forEach(function(header, index) { object[header] = row[index]; });
      object.__rowNumber = group.start + rowIndex;
      result.push(object);
    });
  });
  return result;
}

function getStaffDashboardEvents_(limit) {
  const events = [];
  appendStaffEventsFromSheet_(
    events,
    SHEET_NAMES.VISIT_LOGS,
    HEADERS.visitLogs,
    'visit',
    '방문',
    limit
  );
  appendStaffEventsFromSheet_(
    events,
    SHEET_NAMES.CAREER_LOGS,
    HEADERS.roomLogs,
    'career',
    '커리어존',
    limit
  );
  appendStaffEventsFromSheet_(
    events,
    SHEET_NAMES.CONNECT_LOGS,
    HEADERS.roomLogs,
    'connect',
    '커넥트룸',
    limit
  );
  appendStaffAccessAlerts_(events, limit);
  return events.sort(function(a, b) {
    return str_(b.occurredAt).localeCompare(str_(a.occurredAt));
  }).slice(0, limit);
}

function appendStaffAccessAlerts_(target, limit) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.ACCESS_ALERTS, HEADERS.accessAlerts);
  if (sheet.getLastRow() < 2) return;
  const rowCount = Math.min(limit, sheet.getLastRow() - 1);
  const startRow = sheet.getLastRow() - rowCount + 1;
  const values = sheet.getRange(startRow, 1, rowCount, HEADERS.accessAlerts.length).getValues();
  const index = {};
  HEADERS.accessAlerts.forEach(function(header, column) { index[header] = column; });

  values.forEach(function(row, offset) {
    const alertId = str_(row[index.alertId], 'row-' + String(startRow + offset));
    const date = normalizeSheetDate_(row[index.date]);
    const phoneLastDigits = digits_(row[index.phoneLastDigits]).slice(-4);
    target.push({
      id: 'restricted:' + alertId,
      category: 'restricted',
      categoryLabel: '이용 제한',
      notificationEligible: true,
      origin: 'onsite',
      name: str_(row[index.name], '회원'),
      date: date,
      startTime: '',
      endTime: '',
      purpose: str_(row[index.reason]) + (phoneLastDigits ? ' · 끝자리 ' + phoneLastDigits : ''),
      headcount: 1,
      occurredAt: normalizeDashboardDateTime_(row[index.attemptedAt], date)
    });
  });
}

function appendStaffEventsFromSheet_(target, sheetName, headers, category, categoryLabel, limit) {
  const sheet = getOrCreateSheet_(sheetName, headers);
  if (sheet.getLastRow() < 2) return;
  const rowCount = Math.min(limit, sheet.getLastRow() - 1);
  const startRow = sheet.getLastRow() - rowCount + 1;
  const values = sheet.getRange(startRow, 1, rowCount, headers.length).getValues();
  const index = {};
  headers.forEach(function(header, column) { index[header] = column; });

  values.forEach(function(row, offset) {
    const storedLogId = str_(row[index.logId]);
    const date = normalizeSheetDate_(row[index.date]);
    if (!storedLogId && !date) return;
    const logId = str_(storedLogId, sheetName + ':' + String(startRow + offset));
    const createdAt = normalizeDashboardDateTime_(row[index.createdAt], date);
    const name = str_(row[index.name], category === 'visit' ? '방문자' : '이용자');
    const startTime = category === 'visit' ? '' : normalizeSheetTime_(row[index.startTime]);
    const endTime = category === 'visit' ? '' : normalizeSheetTime_(row[index.endTime]);
    const purpose = category === 'visit' ? '' : str_(row[index.purpose]);
    const headcount = category === 'visit' ? 1 : safePositiveInteger_(row[index.headcount], 1);
    const createdBy = str_(row[index.createdBy]);
    const linkedReservationId = category === 'visit' ? '' : str_(row[index.linkedReservationId]);
    target.push({
      id: category + ':' + logId,
      category: category,
      categoryLabel: categoryLabel,
      notificationEligible: category !== 'visit' &&
        createdBy === 'self' &&
        !linkedReservationId,
      origin: linkedReservationId ? 'naver' : (createdBy === 'admin' ? 'admin' : 'onsite'),
      name: name,
      date: date,
      startTime: startTime,
      endTime: endTime,
      purpose: purpose,
      headcount: headcount,
      occurredAt: createdAt
    });
  });
}

function getStaffTodayReservations_(calendarEntries) {
  const date = today_();
  return (calendarEntries || [])
    .filter(function(row) {
      return (row.source === 'naver' || row.source === 'reservationRequest') && row.date === date;
    })
    .map(function(row) {
      return {
        id: row.source === 'naver' ? str_(row.id).replace(/^naver:/, '') : str_(row.id),
        source: str_(row.source),
        status: str_(row.status),
        bookerName: str_(row.name),
        spaceType: str_(row.spaceType),
        placeName: str_(row.placeName),
        date: date,
        startTime: normalizeSheetTime_(row.startTime),
        endTime: normalizeSheetTime_(row.endTime),
        headcount: safePositiveInteger_(row.headcount, 1),
        purpose: str_(row.purpose),
        cancelled: Boolean(row.cancelled),
        submitted: Boolean(row.submitted)
      };
    })
    .sort(function(a, b) {
      return str_(a.startTime).localeCompare(str_(b.startTime));
    });
}

function getStaffReservationRequests_(calendarEntries) {
  return (calendarEntries || [])
    .filter(function(row) {
      return row.source === 'web' || row.source === 'reservationRequest';
    })
    .map(function(row) {
      return {
        id: str_(row.id),
        requestId: str_(row.id).replace(/^(web|reservationRequest|webReservation):/, ''),
        memberId: str_(row.memberId),
        status: str_(row.status),
        bookerName: str_(row.name),
        date: normalizeSheetDate_(row.date),
        startTime: normalizeSheetTime_(row.startTime),
        endTime: normalizeSheetTime_(row.endTime),
        placeName: str_(row.placeName),
        spaceType: str_(row.spaceType),
        headcount: safePositiveInteger_(row.headcount, 1),
        purpose: str_(row.purpose)
      };
    });
}

function getStaffMonthCheckIns_(requestedMonth) {
  const monthKey = /^\d{4}-\d{2}$/.test(str_(requestedMonth))
    ? str_(requestedMonth)
    : today_().slice(0, 7);
  const cache = CacheService.getScriptCache();
  const cacheKey = 'staff-dashboard-checkins-v1:' + monthKey;
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (error) {
      cache.remove(cacheKey);
    }
  }
  const seen = {};
  const checkIns = getRowsAsObjectsForMonth_(SHEET_NAMES.VISIT_LOGS, HEADERS.visitLogs, 'date', monthKey)
    .map(function(row) {
      return {
        date: normalizeSheetDate_(row.date),
        name: str_(row.name),
        memberId: str_(row.userId)
      };
    })
    .filter(function(checkIn) {
      if (!checkIn.date || !checkIn.name) return false;
      const key = checkIn.date + '|' + (checkIn.memberId || normalizeNameKey_(checkIn.name));
      if (seen[key]) return false;
      seen[key] = true;
      return true;
    });
  try {
    cache.put(cacheKey, JSON.stringify(checkIns), 30);
  } catch (error) {
    Logger.log('직원 대시보드 출석 캐시 생략: ' + (error.message || error));
  }
  return checkIns;
}

function getStaffCalendarEntries_(requestedMonth) {
  const fallbackMonth = today_().slice(0, 7);
  const monthKey = /^\d{4}-\d{2}$/.test(str_(requestedMonth))
    ? str_(requestedMonth)
    : fallbackMonth;
  const cache = CacheService.getScriptCache();
  const cacheKey = 'staff-dashboard-calendar-v1:' + monthKey;
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (error) {
      cache.remove(cacheKey);
    }
  }

  const entries = getRowsAsObjects_(SHEET_NAMES.NAVER_SYNC, HEADERS.naverSync)
    .filter(function(row) {
      return normalizeSheetDate_(row.usageDate).slice(0, 7) === monthKey;
    })
    .map(function(row) {
      return {
        id: 'naver:' + str_(row.reservationId),
        source: 'naver',
        memberId: str_(row.submittedByMemberId),
        date: normalizeSheetDate_(row.usageDate),
        startTime: normalizeSheetTime_(row.startTime),
        endTime: normalizeSheetTime_(row.endTime),
        placeName: getReservationSpaceLabel_(str_(row.spaceType)),
        spaceType: str_(row.spaceType),
        name: str_(row.bookerName),
        headcount: safePositiveInteger_(row.headcount, 1),
        purpose: str_(row.purpose),
        status: str_(row.status),
        cancelled: isCancelledNaverStatus_(row.status),
        submitted: Boolean(str_(row.usageLogId))
      };
    });

  getRowsAsObjects_(SHEET_NAMES.RESERVATION_REQUESTS, HEADERS.reservationRequests)
    .filter(function(row) {
      return normalizeSheetDate_(row.date).slice(0, 7) === monthKey &&
        visibleReservationRequestStatus_(row.status);
    })
    .forEach(function(row) {
      entries.push({
        id: 'web:' + str_(row.requestId),
        source: 'reservationRequest',
        memberId: str_(row.memberId),
        date: normalizeSheetDate_(row.date),
        startTime: normalizeSheetTime_(row.startTime),
        endTime: normalizeSheetTime_(row.endTime),
        placeName: str_(row.placeName, getReservationSpaceLabel_(row.spaceType)),
        spaceType: str_(row.spaceType),
        name: str_(row.name),
        headcount: safePositiveInteger_(row.headcount, 1),
        purpose: str_(row.purpose),
        status: str_(row.status),
        cancelled: str_(row.status) === '취소' || str_(row.status) === '반려',
        submitted: Boolean(str_(row.usageLogId)) || str_(row.status) === '일지제출완료'
      });
    });

  [
    { sheetName: SHEET_NAMES.CAREER_LOGS, label: '커리어존' },
    { sheetName: SHEET_NAMES.CONNECT_LOGS, label: '커넥트룸' }
  ].forEach(function(source) {
    getRowsAsObjects_(source.sheetName, HEADERS.roomLogs)
      .filter(function(row) {
        return normalizeSheetDate_(row.date).slice(0, 7) === monthKey;
      })
      .forEach(function(row) {
        entries.push({
          id: 'usage:' + str_(row.logId),
          source: 'usage',
          date: normalizeSheetDate_(row.date),
          startTime: normalizeSheetTime_(row.startTime),
          endTime: normalizeSheetTime_(row.endTime),
          placeName: source.label,
          spaceType: str_(row.spaceType),
          name: str_(row.name),
          headcount: safePositiveInteger_(row.headcount, 1),
          purpose: str_(row.purpose),
          status: '이용완료',
          cancelled: false,
          submitted: true
        });
      });
  });

  const sorted = entries.sort(function(a, b) {
    return (a.date + ' ' + a.startTime).localeCompare(b.date + ' ' + b.startTime);
  });
  try {
    cache.put(cacheKey, JSON.stringify(sorted), 30);
  } catch (error) {
    Logger.log('직원 대시보드 달력 캐시 생략: ' + (error.message || error));
  }
  return sorted;
}

function getStaffDashboardSummary_() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'staff-dashboard-summary-v2';
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (error) {
      cache.remove(cacheKey);
    }
  }

  const today = today_();
  const monthStart = today.substring(0, 7) + '-01';
  const todayParts = today.split('-').map(Number);
  const todayDate = new Date(todayParts[0], todayParts[1] - 1, todayParts[2]);
  const mondayOffset = (todayDate.getDay() + 6) % 7;
  todayDate.setDate(todayDate.getDate() - mondayOffset);
  const weekStart = Utilities.formatDate(todayDate, TZ, 'yyyy-MM-dd');
  const visits = getRowsAsObjects_(SHEET_NAMES.VISIT_LOGS, HEADERS.visitLogs);
  const visitorStats = visits.reduce(function(stats, row) {
    const date = normalizeSheetDate_(row.date);
    if (date === today) stats.daily += 1;
    if (date >= weekStart && date <= today) stats.weekly += 1;
    if (date >= monthStart && date <= today) stats.monthly += 1;
    return stats;
  }, { daily: 0, weekly: 0, monthly: 0 });

  const memberSheet = getOrCreateSheet_(SHEET_NAMES.MEMBERS, HEADERS.members);
  const memberContext = memberHeaderContext_(memberSheet);
  const memberIdColumn = memberColumn_(memberSheet, 'memberId', memberContext);
  const phoneColumn = memberColumn_(memberSheet, 'phone', memberContext);
  const statusColumn = memberColumn_(memberSheet, 'status', memberContext);
  let activeMemberCount = 0;
  if (memberSheet.getLastRow() >= 2) {
    const firstColumn = Math.min(memberIdColumn, phoneColumn, statusColumn);
    const lastColumn = Math.max(memberIdColumn, phoneColumn, statusColumn);
    const width = lastColumn - firstColumn + 1;
    const memberIdOffset = memberIdColumn - firstColumn;
    const phoneOffset = phoneColumn - firstColumn;
    const statusOffset = statusColumn - firstColumn;
    const activeMemberKeys = memberSheet
      .getRange(2, firstColumn, memberSheet.getLastRow() - 1, width)
      .getValues()
      .reduce(function(keys, row) {
        const memberId = str_(row[memberIdOffset]);
        const active = str_(row[statusOffset], 'approved') === 'approved';
        if (!memberId || !active) return keys;
        const phone = normalizeStoredMemberPhone_(row[phoneOffset]);
        const key = isValidPhone_(phone) ? 'phone:' + phone : 'member:' + memberId;
        keys[key] = true;
        return keys;
      }, {});
    activeMemberCount = Object.keys(activeMemberKeys).length;
  }

  const summary = {
    visitorStats: visitorStats,
    activeMemberCount: activeMemberCount,
    weekStart: weekStart,
    monthStart: monthStart
  };
  cache.put(cacheKey, JSON.stringify(summary), 60);
  return summary;
}

function normalizeDashboardDateTime_(value, fallbackDate) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd HH:mm:ss');
  }
  const text = str_(value);
  if (/^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}(:\d{2})?$/.test(text)) {
    return text.replace('T', ' ') + (text.length === 16 ? ':00' : '');
  }
  return str_(fallbackDate, today_()) + ' 00:00:00';
}

function safePositiveInteger_(value, fallback) {
  const number = Number(value);
  return Number.isFinite(number) && number >= 1
    ? Math.floor(number)
    : Math.max(1, Number(fallback) || 1);
}

function handleAdminLookupMember_(payload) {
  const query = str_(payload.query || payload.name || payload.phone).trim();
  const phone = digits_(payload.phone || query);
  let members = [];
  if (isValidPhone_(phone)) {
    members = findMembersByField_('phone', phone);
  } else {
    const nameKey = normalizeNameKey_(query);
    if (nameKey.length < 2) throw new Error('회원 이름을 2글자 이상 입력해 주세요.');
    members = findMembersByName_(query);
  }

  if (members.length === 0) throw new Error('등록된 회원을 찾을 수 없습니다.');
  const activeMembers = members.filter(isActiveMember_);
  if (activeMembers.length === 0) throw new Error('현재 이용할 수 없는 회원입니다.');
  const member = activeMembers[0];
  return {
    ok: true,
    member: adminMemberResponse_(member),
    members: activeMembers.map(adminMemberResponse_)
  };
}

function findMembersByName_(name) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.MEMBERS, HEADERS.members);
  const headerContext = memberHeaderContext_(sheet);
  const target = normalizeNameKey_(name);
  const lastRow = sheet.getLastRow();
  const matches = [];
  for (let rowNumber = 2; rowNumber <= lastRow; rowNumber++) {
    const member = memberAtRow_(sheet, rowNumber, headerContext);
    if (normalizeNameKey_(member.name) === target) matches.push(member);
  }
  return matches;
}

function adminMemberResponse_(member) {
  return {
    id: member.id,
    name: member.name,
    gender: member.gender,
    age: member.age,
    birthDate: member.birthDate,
    phone: member.phone,
    phoneLastDigits: member.phoneLastDigits
  };
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
    return facilityEntryResponseForMember_(activeExact[0], 'full-phone');
  }

  assertPhoneTailLookupAllowed_();
  const tail = findMembersByField_('phoneLastDigits', digits).filter(isActiveMember_);
  if (tail.length === 0) return { ok: true, mode: 'signup', suggestedPhone: digits };
  if (tail.length === 1) return facilityEntryResponseForMember_(tail[0], 'phone-tail');

  return {
    ok: true,
    mode: 'ambiguous',
    message: '같은 끝자리 4자리를 사용하는 회원이 있습니다. 전체 전화번호를 다시 입력해 주세요.'
  };
}

function facilityEntryResponseForMember_(member, inputType) {
  if (!member.birthDate || isFacilityAgeEligible_(member, today_())) {
    return { ok: true, mode: 'member', member: publicMember_(member) };
  }
  recordAgeRestrictionAlert_(member, inputType);
  return ageRestrictionResponse_(member);
}

function isFacilityAgeEligible_(member, referenceDate) {
  if (!member || !member.birthDate) return false;
  const age = calcAgeOnDate_(member.birthDate, referenceDate);
  return age >= FACILITY_MIN_AGE && age <= FACILITY_MAX_AGE;
}

function ageRestrictionMessage_(age) {
  return '청년이봄 시설은 만 19세 이상 만 39세 이하만 이용할 수 있습니다. 현재 만 ' +
    age + '세로 이용이 불가능합니다.';
}

function assertMemberCanUseSpace_(member, spaceType, referenceDate) {
  const age = calcAgeOnDate_(member.birthDate, referenceDate);
  if (age < FACILITY_MIN_AGE || age > FACILITY_MAX_AGE) {
    throw new Error(ageRestrictionMessage_(age));
  }
}

function ageRestrictionResponse_(member) {
  const age = calcAgeOnDate_(member.birthDate, today_());
  return {
    ok: true,
    mode: 'restricted',
    restricted: true,
    age: age,
    message: ageRestrictionMessage_(age)
  };
}

function recordAgeRestrictionAlert_(member, inputType) {
  return withScriptLock_(function() {
    return recordAgeRestrictionAlertLocked_(member, inputType);
  });
}

function recordAgeRestrictionAlertLocked_(member, inputType) {
  const date = today_();
  const sheet = getOrCreateSheet_(SHEET_NAMES.ACCESS_ALERTS, HEADERS.accessAlerts);
  ensureSheetHeaders_(sheet, HEADERS.accessAlerts);
  const alreadyRecorded = findObjectsByColumnValue_(
    sheet,
    HEADERS.accessAlerts,
    'userId',
    member.id
  ).some(function(row) {
    return normalizeSheetDate_(row.date) === date;
  });
  if (alreadyRecorded) return false;

  const age = calcAgeOnDate_(member.birthDate, date);
  appendObjectRow_(SHEET_NAMES.ACCESS_ALERTS, HEADERS.accessAlerts, {
    alertId: Utilities.getUuid(),
    userId: member.id,
    name: member.name,
    phoneLastDigits: member.phoneLastDigits,
    age: age,
    reason: ageRestrictionMessage_(age),
    date: date,
    attemptedAt: nowText_(),
    inputType: str_(inputType, 'facility-qr')
  });
  return true;
}

function handleRegisterFacilityMember_(payload) {
  const member = normalizeFacilityMember_(payload.member || {});
  if (!member.name || !isValidGender_(member.gender) || !isValidBirthDate_(member.birthDate) || !isValidPhone_(member.phone) || !member.privacyConsent) {
    throw new Error('필수 정보와 개인정보 수집·이용 동의를 모두 확인해 주세요.');
  }
  const registrationAge = calcAgeOnDate_(member.birthDate, today_());

  return withScriptLock_(function() {
    const existing = findMembersByField_('phone', member.phone);
    if (existing.length > 0) {
      const activeExisting = existing.filter(isActiveMember_);
      if (activeExisting.length === 0) throw new Error('현재 이용할 수 없는 회원입니다.');
      if (activeExisting[0].birthDate && !isFacilityAgeEligible_(activeExisting[0], today_())) {
        recordAgeRestrictionAlertLocked_(activeExisting[0], 'signup-existing');
        return ageRestrictionResponse_(activeExisting[0]);
      }
      return { ok: true, existing: true, member: publicMember_(activeExisting[0]) };
    }
    if (registrationAge < FACILITY_MIN_AGE || registrationAge > FACILITY_MAX_AGE) {
      throw new Error(ageRestrictionMessage_(registrationAge) + ' 회원가입을 진행할 수 없습니다.');
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

function setupQrEntryToken() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    '현장 QR 토큰 새로 발급',
    '새 토큰을 발급하면 지금 붙어 있는 QR 코드는 더 이상 작동하지 않습니다.\n새 QR 코드를 출력해 교체해야 합니다. 계속할까요?',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  const token = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty(QR_ENTRY_TOKEN_PROPERTY, token);
  ui.alert(
    '새 QR 토큰을 발급했습니다.',
    '아래 주소로 QR 코드를 새로 만들어 현장에 붙여 주세요.\n\n' +
      QR_ENTRY_PAGE_URL + '?entry=' + token +
      '\n\n이 주소는 GitHub 등 공개된 곳에 올리지 마세요.',
    ui.ButtonSet.OK
  );
}

function getQrEntryToken_() {
  const token = str_(
    PropertiesService.getScriptProperties().getProperty(QR_ENTRY_TOKEN_PROPERTY)
  );
  if (!token) {
    throw new Error('현장 QR 토큰이 설정되지 않았습니다. 직원에게 문의해 주세요.');
  }
  return token;
}

function requireEntryToken_(payload) {
  if (str_(payload && payload.entryToken) !== getQrEntryToken_()) {
    throw new Error('현장 QR 코드로 다시 접속해 주세요.');
  }
}

function assertPhoneTailLookupAllowed_() {
  const cache = CacheService.getScriptCache();
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const bucket = Math.floor(Date.now() / (PHONE_TAIL_LOOKUP_WINDOW_SECONDS * 1000));
    const key = 'phone-tail-lookups:' + bucket;
    const count = Number(cache.get(key) || 0) + 1;
    cache.put(key, String(count), PHONE_TAIL_LOOKUP_WINDOW_SECONDS * 2);
    if (count > PHONE_TAIL_LOOKUP_LIMIT) {
      throw new Error('조회 요청이 많습니다. 전체 휴대폰 번호로 입력하거나 잠시 후 다시 시도해 주세요.');
    }
  } finally {
    lock.releaseLock();
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
    if (!member.birthDate) {
      sheet.getRange(rowNumber, memberColumn_(sheet, 'birthDate')).setValue(birthDate);
      sheet.getRange(rowNumber, memberColumn_(sheet, 'updatedAt')).setValue(nowText_());
    }
    const updatedMember = memberAtRow_(sheet, rowNumber);
    if (!isFacilityAgeEligible_(updatedMember, date)) {
      recordAgeRestrictionAlertLocked_(updatedMember, 'birthdate-update');
      return Object.assign(ageRestrictionResponse_(updatedMember), {
        member: publicMember_(updatedMember)
      });
    }
    const visit = ensureVisitLogLocked_(updatedMember, date, source, []);
    return {
      ok: true,
      member: publicMember_(updatedMember),
      visit: visit
    };
  });
}

function handleSubmitVisitLog_(payload) {
  const memberEntry = requireActiveMemberEntry_(payload.memberId);
  const member = memberEntry.member;
  const date = today_();
  const source = normalizeVisitSource_(payload.source);

  if (!member.birthDate) throw new Error('생년월일을 먼저 입력해 주세요.');
  if (!isFacilityAgeEligible_(member, date)) {
    recordAgeRestrictionAlert_(member, 'visit-submit');
    throw new Error(ageRestrictionMessage_(calcAgeOnDate_(member.birthDate, date)));
  }

  return withScriptLock_(function() {
    return ensureVisitLogLocked_(member, date, source, [], 'self', memberEntry);
  });
}

function handleGetTodayReservations_(payload) {
  const member = requireActiveMember_(payload.memberId);
  return {
    ok: true,
    reservations: getTodayReservationsForMember_(member, today_())
  };
}

function handleGetSpaceAvailability_(payload) {
  const spaceType = normalizeReservationSpaceType_(payload.spaceType, payload.placeId);
  const date = normalizeSheetDate_(payload.date);
  if (!isValidReservationDate_(date)) throw new Error('예약 날짜를 확인해 주세요.');
  return {
    ok: true,
    date: date,
    spaceType: spaceType,
    reservations: getPublicSpaceReservations_(spaceType, date)
  };
}

function handleSubmitSpaceReservationRequest_(payload) {
  const member = requireActiveMember_(payload.memberId);
  const spaceType = normalizeReservationSpaceType_(payload.spaceType, payload.placeId);
  const date = normalizeSheetDate_(payload.date);
  const startTime = normalizeSheetTime_(payload.startTime);
  const endTime = normalizeSheetTime_(payload.endTime);
  const purpose = str_(payload.purpose);
  const headcount = Number(payload.headcount || 1);
  const maxHeadcount = maxHeadcountForSpace_(spaceType);
  const expectedCompanionCount = Math.max(0, headcount - 1);
  const companions = normalizeCompanions_(payload.companions, expectedCompanionCount);

  if (!isValidReservationDate_(date)) throw new Error('예약 날짜를 확인해 주세요.');
  if (!startTime || !endTime || endTime <= startTime) throw new Error('예약 시간을 확인해 주세요.');
  if (!purpose) throw new Error('이용 목적을 입력해 주세요.');
  if (purpose.length > 200) throw new Error('이용 목적은 200자 이내로 입력해 주세요.');
  if (!Number.isInteger(headcount) || headcount < 1 || headcount > maxHeadcount) {
    throw new Error(getReservationSpaceLabel_(spaceType) + '은 1~' + maxHeadcount + '명까지 신청할 수 있습니다.');
  }
  if (companions.length !== expectedCompanionCount) {
    throw new Error('동반 이용자의 성별과 나이를 모두 입력해 주세요.');
  }
  assertMemberCanUseSpace_(member, spaceType, date);

  return withScriptLock_(function() {
    assertNoReservationRequestConflict_(spaceType, date, startTime, endTime, '');
    const requestId = 'web-' + Utilities.getUuid();
    appendObjectRow_(SHEET_NAMES.RESERVATION_REQUESTS, HEADERS.reservationRequests, {
      requestId: requestId,
      status: '신청',
      memberId: member.id,
      name: member.name,
      memberNameKey: normalizeNameKey_(member.name),
      gender: member.gender,
      age: member.birthDate ? calcAgeOnDate_(member.birthDate, date) : member.age,
      phone: member.phone,
      phoneLastDigits: member.phoneLastDigits,
      spaceType: spaceType,
      placeId: spaceType,
      placeName: getReservationSpaceLabel_(spaceType),
      date: date,
      startTime: startTime,
      endTime: endTime,
      purpose: purpose,
      headcount: headcount,
      companionsJson: JSON.stringify(companions),
      companionCount: companions.length,
      requestedAt: nowText_(),
      decidedAt: '',
      decidedBy: '',
      decisionNote: '',
      usageLogId: '',
      submittedAt: '',
      submittedByMemberId: '',
      updatedAt: nowText_()
    });
    clearReservationCaches_(date);
    return {
      ok: true,
      requestId: requestId,
      status: '신청',
      reservation: publicReservationRequest_({
        requestId: requestId,
        status: '신청',
        memberId: member.id,
        name: member.name,
        spaceType: spaceType,
        date: date,
        startTime: startTime,
        endTime: endTime,
        purpose: purpose,
        headcount: headcount,
        companionsJson: JSON.stringify(companions),
        usageLogId: ''
      })
    };
  });
}

function ensureVisitLogLocked_(member, date, source, preparedReservations, createdBy, memberEntry) {
  const reservations = Array.isArray(preparedReservations)
    ? preparedReservations
    : [];
  if (hasVisitLog_(member.id, date)) {
    if (member.role !== 'guest') updateMemberLastUsedAt_(member.id, date, memberEntry);
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
    age: member.birthDate ? calcAgeOnDate_(member.birthDate, date) : member.age,
    phone: member.phone,
    phoneLastDigits: member.phoneLastDigits,
    date: date,
    source: source,
    createdBy: str_(createdBy, 'self'),
    createdAt: nowText_()
  });
  if (member.role !== 'guest') updateMemberLastUsedAt_(member.id, date, memberEntry);
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
  assertMemberCanUseSpace_(member, type, date);

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
        age: member.birthDate ? calcAgeOnDate_(member.birthDate, date) : member.age,
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
    markOperationCalendarMonthPending_(roomData.date);
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
    if (reservationId.indexOf('web:') === 0) {
      return submitWebReservationUsageLocked_(member, reservationId.replace(/^web:/, ''), payload.companions);
    }
    const sheet = ensureNaverReservationSchema_();
    const entry = findNaverReservationEntry_(sheet, reservationId);
    if (!entry) throw new Error('예약 정보를 찾을 수 없습니다. 시트에서 예약 동기화를 다시 실행해 주세요.');
    assertMemberCanUseSpace_(member, str_(entry.values.spaceType), date);
    return submitReservationUsageLocked_(member, sheet, entry, payload.companions, {
      duplicateAsError: true,
      repairedAsSuccess: true,
      source: 'reservation-card'
    });
  });
}

function handleAdminCompleteReservationUsage_(payload) {
  const source = str_(payload.source);
  const reservationId = str_(payload.reservationId).replace(/^(naverSmartPlace:|naver:|reservationRequest:|webReservation:|web:|app:)/, '');
  if (!reservationId) throw new Error('예약 정보를 다시 선택해 주세요.');

  return withScriptLock_(function() {
    if (source === 'reservationRequest' || source === 'webReservation' || source === 'web' || source === 'app') {
      const sheet = getOrCreateSheet_(SHEET_NAMES.RESERVATION_REQUESTS, HEADERS.reservationRequests);
      ensureSheetHeaders_(sheet, HEADERS.reservationRequests);
      const entry = findReservationRequestEntry_(sheet, reservationId);
      if (!entry) throw new Error('예약 신청 정보를 찾을 수 없습니다.');
      const member = requireActiveMember_(str_(payload.memberId, entry.values.memberId));
      const result = submitWebReservationUsageLocked_(member, reservationId, payload.companions, {
        skipSubmitWindowCheck: true
      });
      return {
        ok: true,
        reservationId: result.reservationId,
        status: '이용완료',
        cancelled: false,
        usageLogged: true
      };
    }

    const sheet = ensureNaverReservationSchema_();
    const entry = findNaverReservationEntry_(sheet, reservationId);
    if (!entry) {
      const member = requireActiveMember_(payload.memberId);
      const spaceType = str_(payload.spaceType);
      const date = normalizeSheetDate_(payload.date);
      const roomData = {
        date: date,
        startTime: normalizeSheetTime_(payload.startTime),
        endTime: normalizeSheetTime_(payload.endTime),
        purpose: str_(payload.purpose, '네이버 예약 이용'),
        companions: normalizeCompanions_(payload.companions, Math.max(0, Number(payload.headcount || 1) - 1))
      };
      if (!date || !roomData.startTime || !roomData.endTime || !spaceType) {
        throw new Error('예약 정보를 찾을 수 없습니다. 시트에서 예약 동기화를 다시 실행해 주세요.');
      }
      const roomLogSheet = getRoomLogSheetName_(spaceType);
      assertNoRoomLogConflict_(roomLogSheet, member, spaceType, roomData);
      const roomRow = buildRoomLogRow_(member, spaceType, roomData, {
        createdBy: 'naver-smartplace',
        linkedReservationId: reservationId,
        dedupeStatus: 'staff-smartplace-complete'
      });
      appendObjectRow_(roomLogSheet, HEADERS.roomLogs, roomRow);
      markOperationCalendarMonthPending_(date);
      return {
        ok: true,
        reservationId: 'naverSmartPlace:' + reservationId,
        status: '이용완료',
        cancelled: false,
        usageLogged: true
      };
    }
    const member = requireActiveMember_(payload.memberId || entry.values.submittedByMemberId);
    const result = submitReservationUsageLocked_(member, sheet, entry, payload.companions, {
      duplicateAsError: false,
      repairedAsSuccess: true,
      source: 'staff-status-complete',
      skipSubmitWindowCheck: true,
      date: normalizeSheetDate_(entry.values.usageDate)
    });
    return {
      ok: true,
      reservationId: result.reservationId,
      status: '이용완료',
      cancelled: false,
      usageLogged: true
    };
  });
}

function submitReservationUsageLocked_(member, sheet, entry, companionsValue, options) {
  const reservation = entry.values;
  const reservationId = str_(reservation.reservationId);
  const settings = options || {};
  const date = str_(settings.date, today_());
  if (normalizeSheetDate_(reservation.usageDate) !== date) {
    throw new Error('선택한 날짜의 예약과 일치하지 않습니다.');
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
  if (!settings.skipSubmitWindowCheck) {
    assertReservationSubmitWindow_(reservation);
  }

  const linkedLog = findLinkedReservationLog_(reservationId);
  if (linkedLog) {
    if (str_(linkedLog.row.userId) !== str_(member.id)) {
      throw new Error('이미 다른 회원의 이용일지와 연결된 예약입니다. 관리자에게 확인해 주세요.');
    }
    updateNaverReservationSubmission_(sheet, entry.rowNumber, member, reservation, linkedLog.row.logId);
    markOperationCalendarMonthPending_(date);
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
    startTime: normalizeSheetTime_(reservation.startTime),
    endTime: normalizeSheetTime_(reservation.endTime),
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
  markOperationCalendarMonthPending_(date);
  return { ok: true, created: true, reservationId: reservationId, linkedReservation: true };
}

function submitWebReservationUsageLocked_(member, requestId, companionsValue, options) {
  const settings = options || {};
  const sheet = getOrCreateSheet_(SHEET_NAMES.RESERVATION_REQUESTS, HEADERS.reservationRequests);
  ensureSheetHeaders_(sheet, HEADERS.reservationRequests);
  const entry = findReservationRequestEntry_(sheet, requestId);
  if (!entry) throw new Error('예약 신청 정보를 찾을 수 없습니다.');

  const reservation = entry.values;
  const status = str_(reservation.status);
  if (!activeReservationRequestStatus_(status)) {
    throw new Error('이용할 수 없는 예약 상태입니다.');
  }
  if (str_(reservation.memberId) !== str_(member.id)) {
    throw new Error('로그인한 회원의 예약과 일치하지 않습니다.');
  }
  if (str_(reservation.usageLogId)) {
    throw new Error('이미 동일한 시간대로 제출하였습니다.');
  }
  if (!settings.skipSubmitWindowCheck) {
    assertReservationSubmitWindow_({
      usageDate: reservation.date,
      startTime: reservation.startTime
    });
  }

  const spaceType = str_(reservation.spaceType);
  const roomLogSheet = getRoomLogSheetName_(spaceType);
  const expectedCompanionCount = Math.max(0, Number(reservation.headcount || 1) - 1);
  const storedCompanions = parseCompanionsJson_(reservation.companionsJson);
  const companions = storedCompanions.length === expectedCompanionCount
    ? storedCompanions
    : normalizeCompanions_(companionsValue, expectedCompanionCount);
  if (companions.length !== expectedCompanionCount) {
    throw new Error('예약 인원에 맞게 동반 이용자의 성별과 나이를 모두 입력해 주세요.');
  }

  const roomData = {
    date: normalizeSheetDate_(reservation.date),
    startTime: normalizeSheetTime_(reservation.startTime),
    endTime: normalizeSheetTime_(reservation.endTime),
    purpose: str_(reservation.purpose, '웹 예약 이용'),
    companions: companions
  };
  if (!roomData.startTime || !roomData.endTime || roomData.endTime <= roomData.startTime) {
    throw new Error('예약 시간이 올바르지 않습니다.');
  }
  assertNoRoomLogConflict_(roomLogSheet, member, spaceType, roomData);

  const roomRow = buildRoomLogRow_(member, spaceType, roomData, {
    createdBy: 'web-reservation',
    linkedReservationId: requestId,
    dedupeStatus: 'web-reservation'
  });
  appendObjectRow_(roomLogSheet, HEADERS.roomLogs, roomRow);
  updateReservationRequestRow_(sheet, entry.rowNumber, {
    status: '일지제출완료',
    usageLogId: roomRow.logId,
    submittedAt: nowText_(),
    submittedByMemberId: member.id,
    updatedAt: nowText_()
  });
  clearReservationCaches_(roomData.date);
  markOperationCalendarMonthPending_(roomData.date);
  return { ok: true, created: true, reservationId: 'web:' + requestId, linkedReservation: true };
}

function assertReservationSubmitWindow_(reservation) {
  const usageDate = normalizeSheetDate_(reservation.usageDate);
  const nowDate = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (usageDate !== nowDate) {
    throw new Error('예약 당일에 일지를 제출할 수 있습니다.');
  }

  // 2026-11-12 운영 정책: 예약 당일이면 시간과 관계없이 제출할 수 있다.
  const startTime = normalizeSheetTime_(reservation.startTime);
  if (!/^\d{2}:\d{2}$/.test(startTime)) {
    throw new Error('예약 시간을 확인한 뒤 일지를 제출할 수 있습니다.');
  }
}

function buildRoomLogRow_(member, spaceType, roomData, metadata) {
  const details = metadata || {};
  const memberAge = member.birthDate
    ? calcAgeOnDate_(member.birthDate, roomData.date)
    : member.age;
  const row = {
    logId: Utilities.getUuid(),
    userId: member.id,
    spaceType: spaceType,
    name: member.name,
    memberNameKey: normalizeNameKey_(member.name),
    gender: member.gender,
    age: memberAge,
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
  const summary = buildRoomSummary_(Object.assign({}, member, { age: memberAge }), roomData);
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
    accessAlertCount: anonymizeLogSheet_(
      SHEET_NAMES.ACCESS_ALERTS,
      HEADERS.accessAlerts,
      ids,
      ['userId', 'name', 'phoneLastDigits']
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
    reservationLinkCount: logCounts.reservationLinkCount || 0,
    accessAlertCount: logCounts.accessAlertCount || 0
  });
}

function retentionResult_(memberCount, logCounts) {
  const result = {
    memberCount: memberCount,
    visitLogCount: logCounts.visitLogCount || 0,
    printerLogCount: logCounts.printerLogCount || 0,
    careerLogCount: logCounts.careerLogCount || 0,
    connectLogCount: logCounts.connectLogCount || 0,
    reservationLinkCount: logCounts.reservationLinkCount || 0,
    accessAlertCount: logCounts.accessAlertCount || 0
  };
  result.totalLogCount = result.visitLogCount + result.printerLogCount + result.careerLogCount +
    result.connectLogCount + result.reservationLinkCount + result.accessAlertCount;
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
    const formulaCell = sheet.getRange('A2');
    const formula = str_(formulaCell.getFormula());
    const formulaError = /^#(?:REF|VALUE|ERROR|N\/A)!?$/i.test(str_(formulaCell.getDisplayValue()));
    const sortedFormulaCurrent = formula.indexOf('SORT(') !== -1 &&
      formula.indexOf('CHOOSECOLS(') !== -1;
    const roomFormulaCurrent = item.name !== SHEET_NAMES.CAREER && item.name !== SHEET_NAMES.CONNECT
      ? true
      : formula.indexOf('!AC2:INDEX(') !== -1 &&
        (item.name !== SHEET_NAMES.CAREER || formula.indexOf('!C2:INDEX(') !== -1);
    if (!headersMatch || !formula || formulaError || !sortedFormulaCurrent || !roomFormulaCurrent) {
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
      'filtered,FILTER(HSTACK(dates,APP_VISIT_LOGS!E2:INDEX(APP_VISIT_LOGS!E:E,last),APP_VISIT_LOGS!F2:INDEX(APP_VISIT_LOGS!F:F,last),' +
      'APP_VISIT_LOGS!L2:INDEX(APP_VISIT_LOGS!L:L,last),SEQUENCE(ROWS(dates))),' +
      'IFERROR(TEXT(dates,"yyyy-mm"),LEFT(dates&"",7))="' + monthKey + '"),' +
      'sorted,SORT(filtered,1,TRUE,4,TRUE,5,TRUE),data,CHOOSECOLS(sorted,1,2,3),HSTACK(SEQUENCE(ROWS(data)),data)),"")';
  }
  if (sheetName === SHEET_NAMES.PRINTER) {
    return '=IFERROR(LET(last,MAX(2,COUNTA(APP_PRINTER_LOGS!A:A)),dates,APP_PRINTER_LOGS!H2:INDEX(APP_PRINTER_LOGS!H:H,last),' +
      'filtered,FILTER(HSTACK(dates,APP_PRINTER_LOGS!D2:INDEX(APP_PRINTER_LOGS!D:D,last),APP_PRINTER_LOGS!E2:INDEX(APP_PRINTER_LOGS!E:E,last),' +
      'APP_PRINTER_LOGS!I2:INDEX(APP_PRINTER_LOGS!I:I,last),APP_PRINTER_LOGS!K2:INDEX(APP_PRINTER_LOGS!K:K,last),SEQUENCE(ROWS(dates))),' +
      'IFERROR(TEXT(dates,"yyyy-mm"),LEFT(dates&"",7))="' + monthKey + '"),' +
      'sorted,SORT(filtered,1,TRUE,5,TRUE,6,TRUE),data,CHOOSECOLS(sorted,1,2,3,4),HSTACK(SEQUENCE(ROWS(data)),data)),"")';
  }

  const source = sheetName === SHEET_NAMES.CAREER ? 'APP_CAREER_LOGS' : 'APP_CONNECT_LOGS';
  if (sheetName === SHEET_NAMES.CAREER) {
    return '=IFERROR(LET(last,MAX(2,COUNTA(' + source + '!A:A)),dates,' + source + '!J2:INDEX(' + source + '!J:J,last),' +
      'spaces,IF(' + source + '!C2:INDEX(' + source + '!C:C,last)="careerRoom1","Room1",IF(' +
      source + '!C2:INDEX(' + source + '!C:C,last)="careerRoom2","Room2",' + source + '!C2:INDEX(' + source + '!C:C,last))),' +
      'filtered,FILTER(HSTACK(dates,spaces,' + source + '!M2:INDEX(' + source + '!M:M,last),' + source + '!T2:INDEX(' + source + '!T:T,last),' +
      source + '!U2:INDEX(' + source + '!U:U,last),' + source + '!K2:INDEX(' + source + '!K:K,last),' + source + '!L2:INDEX(' + source + '!L:L,last),' +
      source + '!V2:INDEX(' + source + '!V:V,last),' + source + '!W2:INDEX(' + source + '!W:W,last),' + source + '!X2:INDEX(' + source + '!X:X,last),' +
      source + '!Y2:INDEX(' + source + '!Y:Y,last),' + source + '!Z2:INDEX(' + source + '!Z:Z,last),' + source + '!AA2:INDEX(' + source + '!AA:AA,last),' +
      source + '!AB2:INDEX(' + source + '!AB:AB,last),' + source + '!AC2:INDEX(' + source + '!AC:AC,last),' +
      source + '!Q2:INDEX(' + source + '!Q:Q,last),SEQUENCE(ROWS(dates))),' +
      'IFERROR(TEXT(dates,"yyyy-mm"),LEFT(dates&"",7))="' + monthKey + '"),' +
      'sorted,SORT(filtered,1,TRUE,16,TRUE,17,TRUE),data,CHOOSECOLS(sorted,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15),' +
      'HSTACK(SEQUENCE(ROWS(data)),data)),"")';
  }
  return '=IFERROR(LET(last,MAX(2,COUNTA(' + source + '!A:A)),dates,' + source + '!J2:INDEX(' + source + '!J:J,last),' +
    'filtered,FILTER(HSTACK(dates,' + source + '!M2:INDEX(' + source + '!M:M,last),' + source + '!T2:INDEX(' + source + '!T:T,last),' +
    source + '!U2:INDEX(' + source + '!U:U,last),' + source + '!K2:INDEX(' + source + '!K:K,last),' + source + '!L2:INDEX(' + source + '!L:L,last),' +
    source + '!V2:INDEX(' + source + '!V:V,last),' + source + '!W2:INDEX(' + source + '!W:W,last),' + source + '!X2:INDEX(' + source + '!X:X,last),' +
    source + '!Y2:INDEX(' + source + '!Y:Y,last),' + source + '!Z2:INDEX(' + source + '!Z:Z,last),' + source + '!AA2:INDEX(' + source + '!AA:AA,last),' +
    source + '!AB2:INDEX(' + source + '!AB:AB,last),' + source + '!AC2:INDEX(' + source + '!AC:AC,last),' +
    source + '!Q2:INDEX(' + source + '!Q:Q,last),SEQUENCE(ROWS(dates))),' +
    'IFERROR(TEXT(dates,"yyyy-mm"),LEFT(dates&"",7))="' + monthKey + '"),' +
    'sorted,SORT(filtered,1,TRUE,15,TRUE,16,TRUE),data,CHOOSECOLS(sorted,1,2,3,4,5,6,7,8,9,10,11,12,13,14),' +
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
  const entry = getMemberEntryById_(memberId);
  return entry ? entry.member : null;
}

function getMemberEntryById_(memberId) {
  if (!memberId) return null;
  const sheet = getOrCreateSheet_(SHEET_NAMES.MEMBERS, HEADERS.members);
  const headerContext = memberHeaderContext_(sheet);
  const rowNumber = findFirstRowByColumnValue_(
    sheet,
    memberColumn_(sheet, 'memberId', headerContext),
    memberId
  );
  return rowNumber ? {
    sheet: sheet,
    rowNumber: rowNumber,
    headerContext: headerContext,
    member: memberAtRow_(sheet, rowNumber, headerContext)
  } : null;
}

function updateMemberLastUsedAt_(memberId, date, memberEntry) {
  const entry = memberEntry || getMemberEntryById_(memberId);
  if (!entry) throw new Error('회원 정보를 찾을 수 없습니다.');
  const headerContext = entry.headerContext || memberHeaderContext_(entry.sheet);
  if (!headerContext.columnByHeader.lastusedat) {
    ensureSheetHeaders_(entry.sheet, HEADERS.members);
    entry.headerContext = memberHeaderContext_(entry.sheet);
  }
  const lastUsedRange = entry.sheet.getRange(
    entry.rowNumber,
    memberColumn_(entry.sheet, 'lastUsedAt', entry.headerContext)
  );
  const currentDate = dateOnly_(lastUsedRange.getValue());
  const nextDate = dateOnly_(date);
  if (currentDate && (!nextDate || currentDate >= nextDate)) return;
  const updatedAtColumn = memberColumn_(entry.sheet, 'updatedAt', entry.headerContext);
  const lastUsedAtColumn = memberColumn_(entry.sheet, 'lastUsedAt', entry.headerContext);
  if (lastUsedAtColumn === updatedAtColumn + 1) {
    entry.sheet.getRange(entry.rowNumber, updatedAtColumn, 1, 2)
      .setValues([[nowText_(), nextDate]]);
  } else {
    entry.sheet.getRange(entry.rowNumber, lastUsedAtColumn).setValue(nextDate);
    entry.sheet.getRange(entry.rowNumber, updatedAtColumn).setValue(nowText_());
  }
}

function findMembersByField_(field, value) {
  const sheet = getOrCreateSheet_(SHEET_NAMES.MEMBERS, HEADERS.members);
  const target = digits_(value);
  const headerContext = memberHeaderContext_(sheet);
  const column = memberColumn_(sheet, field, headerContext);
  const isPhoneSearch = field === 'phone';
  const isPhoneTailSearch = field === 'phoneLastDigits';
  if (!isPhoneSearch && !isPhoneTailSearch) {
    return findRowsByColumnValue_(sheet, column, value).map(function(rowNumber) {
      return memberAtRow_(sheet, rowNumber, headerContext);
    });
  }

  const phoneColumn = memberColumn_(sheet, 'phone', headerContext);
  const phoneTailColumn = memberColumn_(sheet, 'phoneLastDigits', headerContext);
  const pattern = isPhoneSearch
    ? buildPhoneRegex_(target)
    : buildMemberPhoneTailLookupRegex_(target);

  return findRowsByColumnsRegex_(sheet, [phoneColumn, phoneTailColumn], pattern)
    .map(function(rowNumber) { return memberAtRow_(sheet, rowNumber, headerContext); })
    .filter(function(member) {
      if (isPhoneSearch) return member.phone === target;
      return member.phoneLastDigits === target || member.phone.slice(-4) === target;
    });
}

function memberHeaderContext_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), HEADERS.members.length);
  const currentHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const columnByHeader = {};
  currentHeaders.forEach(function(value, index) {
    const key = str_(value).trim().toLowerCase();
    if (key && columnByHeader[key] === undefined) columnByHeader[key] = index + 1;
  });
  return {
    width: lastColumn,
    columnByHeader: columnByHeader
  };
}

function memberColumn_(sheet, header, headerContext) {
  const context = headerContext || memberHeaderContext_(sheet);
  const column = context.columnByHeader[str_(header).trim().toLowerCase()];
  if (!column) throw new Error('APP_MEMBERS 시트에서 ' + header + ' 열을 찾을 수 없습니다.');
  return column;
}

function memberAtRow_(sheet, rowNumber, headerContext) {
  const context = headerContext || memberHeaderContext_(sheet);
  const values = sheet.getRange(rowNumber, 1, 1, context.width).getValues()[0];
  const row = {};
  HEADERS.members.forEach(function(header) {
    const column = context.columnByHeader[header.toLowerCase()];
    row[header] = column ? values[column - 1] : '';
  });
  return memberRowToObject_(row);
}

function requireActiveMember_(memberId) {
  return requireActiveMemberEntry_(memberId).member;
}

function requireActiveMemberEntry_(memberId) {
  const entry = getMemberEntryById_(str_(memberId));
  if (!entry) throw new Error('회원 정보를 찾을 수 없습니다. 처음부터 다시 진행해 주세요.');
  if (!isActiveMember_(entry.member)) throw new Error('현재 이용할 수 없는 회원입니다.');
  return entry;
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

function findRowsByColumnsRegex_(sheet, columns, pattern) {
  if (!pattern || sheet.getLastRow() < 2) return [];
  const firstColumn = Math.min.apply(null, columns);
  const lastColumn = Math.max.apply(null, columns);
  const matchingRows = {};
  sheet.getRange(2, firstColumn, sheet.getLastRow() - 1, lastColumn - firstColumn + 1)
    .createTextFinder(pattern)
    .useRegularExpression(true)
    .matchEntireCell(true)
    .findAll()
    .forEach(function(range) {
      if (columns.indexOf(range.getColumn()) !== -1) matchingRows[range.getRow()] = true;
    });
  return Object.keys(matchingRows).map(Number).sort(function(a, b) { return a - b; });
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

function buildMemberPhoneTailLookupRegex_(phoneTail) {
  const storedTailPattern = buildPhoneTailRegex_(phoneTail);
  const phoneSuffixPattern = buildPhoneSuffixRegex_(phoneTail);
  if (!storedTailPattern || !phoneSuffixPattern) return '';
  return '(' + storedTailPattern + ')|(' + phoneSuffixPattern + ')';
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
  if (exact && exact.getLastRow() > 1) return exact;

  const matching = spreadsheet.getSheets().filter(function(sheet) {
    return str_(sheet.getName()).trim().toLowerCase() === expectedName;
  });

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
    { name: SHEET_NAMES.RESERVATION_REQUESTS, headers: HEADERS.reservationRequests },
    { name: SHEET_NAMES.RETENTION_LOGS, headers: HEADERS.retentionLogs },
    { name: SHEET_NAMES.ACCESS_ALERTS, headers: HEADERS.accessAlerts }
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
  const touchedMonthKeys = {};

  Object.keys(latestRecords).sort().forEach(function(reservationId) {
    const record = latestRecords[reservationId];
    const existing = existingMap[reservationId];
    markReservationMonth_(touchedMonthKeys, existing && existing.values && existing.values.usageDate);
    markReservationMonth_(touchedMonthKeys, record.usageDate);
    const result = upsertNaverReservation_(sheet, existing, record);
    summary[result.changeType] += 1;
    if (isNoShowNaverStatus_(record.status)) summary.noShow += 1;
    else if (isCancelledNaverStatus_(record.status)) summary.cancelled += 1;
  });
  sortNaverReservationsByUsageDate_(sheet);
  let calendarUpdatedCount = 0;
  let calendarErrorCount = 0;
  const calendarErrors = [];
  Object.keys(touchedMonthKeys).sort().forEach(function(monthKey) {
    try {
      renderMonthlyOperationCalendarByMonthKey_(monthKey);
      calendarUpdatedCount += 1;
    } catch (error) {
      calendarErrorCount += 1;
      const message = monthKey + ': ' + (error.message || error);
      calendarErrors.push(message);
      Logger.log('월별 예약 달력 갱신 오류 (' + message + ')');
    }
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
    '월별 달력 갱신: ' + calendarUpdatedCount + '개' +
    (calendarErrorCount ? ' / 오류: ' + calendarErrorCount + '개' : '') + '\n' +
    (calendarErrors.length ? '달력 오류 내용: ' + calendarErrors.join(' | ') + '\n' : '') +
    'Drive 휴지통 이동: ' + trashedFileCount + '개'
  );
  return summary;
}

function sortNaverReservationsByUsageDate_(sheet) {
  const rowCount = sheet.getLastRow() - 1;
  if (rowCount < 2) return;

  sheet.getRange(2, 1, rowCount, HEADERS.naverSync.length).sort([
    { column: HEADERS.naverSync.indexOf('usageDate') + 1, ascending: true },
    { column: HEADERS.naverSync.indexOf('startTime') + 1, ascending: true },
    { column: HEADERS.naverSync.indexOf('reservationId') + 1, ascending: true }
  ]);
}

function markReservationMonth_(monthKeys, value) {
  const date = normalizeSheetDate_(value);
  if (date) monthKeys[date.slice(0, 7)] = true;
}

function readPendingOperationCalendarMonths_() {
  const raw = str_(
    PropertiesService.getScriptProperties().getProperty(PENDING_CALENDAR_MONTHS_PROPERTY)
  );
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (error) {
    Logger.log('운영 달력 대기 목록 읽기 오류: ' + (error.message || error));
    return {};
  }
}

function markOperationCalendarMonthPending_(value) {
  const date = normalizeSheetDate_(value);
  const monthKey = date.slice(0, 7);
  if (!isValidMonthKey_(monthKey)) return;

  const pending = readPendingOperationCalendarMonths_();
  pending[monthKey] = Date.now();
  PropertiesService.getScriptProperties().setProperty(
    PENDING_CALENDAR_MONTHS_PROPERTY,
    JSON.stringify(pending)
  );
}

function clearReservationCaches_(date) {
  const monthKey = normalizeSheetDate_(date).slice(0, 7);
  if (!isValidMonthKey_(monthKey)) return;
  const cache = CacheService.getScriptCache();
  cache.remove('staff-dashboard-calendar-v1:' + monthKey);
  cache.remove('staff-dashboard-summary-v2');
}

function clearPendingOperationCalendarMonth_(monthKey, expectedVersion) {
  const pending = readPendingOperationCalendarMonths_();
  if (expectedVersion !== undefined && pending[monthKey] !== expectedVersion) return;
  if (pending[monthKey] === undefined) return;

  delete pending[monthKey];
  const properties = PropertiesService.getScriptProperties();
  if (Object.keys(pending).length > 0) {
    properties.setProperty(PENDING_CALENDAR_MONTHS_PROPERTY, JSON.stringify(pending));
  } else {
    properties.deleteProperty(PENDING_CALENDAR_MONTHS_PROPERTY);
  }
}

function refreshPendingOperationCalendars_() {
  const pending = readPendingOperationCalendarMonths_();
  Object.keys(pending).sort().forEach(function(monthKey) {
    try {
      renderMonthlyOperationCalendarByMonthKey_(monthKey);
      clearPendingOperationCalendarMonth_(monthKey, pending[monthKey]);
    } catch (error) {
      Logger.log('대기 중인 운영 달력 갱신 오류 (' + monthKey + '): ' + (error.message || error));
    }
  });
}

function generateMonthlyOperationCalendar() {
  const monthKey = Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  const pendingVersion = readPendingOperationCalendarMonths_()[monthKey];
  renderMonthlyOperationCalendarByMonthKey_(monthKey);
  if (pendingVersion !== undefined) {
    clearPendingOperationCalendarMonth_(monthKey, pendingVersion);
  }
  SpreadsheetApp.getUi().alert(monthKey + ' 운영 달력을 갱신했습니다.');
}

function refreshSelectedOperationCalendar() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('달력을 표시할 월을 입력하세요.\n형식: YYYY-MM\n예: 2026-07');
  if (!response || response.getSelectedButton() === ui.Button.CANCEL) return;

  const monthKey = str_(response.getResponseText());
  if (!isValidMonthKey_(monthKey)) {
    ui.alert('형식이 잘못되었습니다. YYYY-MM 형식으로 입력해 주세요.');
    return;
  }

  const pendingVersion = readPendingOperationCalendarMonths_()[monthKey];
  renderMonthlyOperationCalendarByMonthKey_(monthKey);
  if (pendingVersion !== undefined) {
    clearPendingOperationCalendarMonth_(monthKey, pendingVersion);
  }
  ui.alert(monthKey + ' 운영 달력을 갱신했습니다.');
}

function renderMonthlyOperationCalendarByMonthKey_(monthKey) {
  if (!isValidMonthKey_(monthKey)) return;

  const parts = monthKey.split('-').map(Number);
  ensureNaverReservationSchema_();
  const reservations = getRowsAsObjects_(SHEET_NAMES.NAVER_SYNC, HEADERS.naverSync).filter(function(reservation) {
    return normalizeSheetDate_(reservation.usageDate).slice(0, 7) === monthKey;
  }).concat(getRowsAsObjects_(SHEET_NAMES.RESERVATION_REQUESTS, HEADERS.reservationRequests)
    .filter(function(reservation) {
      return normalizeSheetDate_(reservation.date).slice(0, 7) === monthKey &&
        visibleReservationRequestStatus_(reservation.status);
    })
    .map(function(reservation) {
      return {
        reservationId: str_(reservation.requestId),
        status: str_(reservation.status),
        bookerName: str_(reservation.name),
        usageDate: normalizeSheetDate_(reservation.date),
        startTime: normalizeSheetTime_(reservation.startTime),
        endTime: normalizeSheetTime_(reservation.endTime),
        spaceType: str_(reservation.spaceType),
        headcount: reservation.headcount,
        purpose: str_(reservation.purpose, '웹 예약 이용'),
        usageLogId: str_(reservation.usageLogId),
        source: 'reservationRequest'
      };
    }));
  const roomLogs = [
    { sheetName: SHEET_NAMES.CAREER_LOGS, headers: HEADERS.roomLogs },
    { sheetName: SHEET_NAMES.CONNECT_LOGS, headers: HEADERS.roomLogs }
  ].reduce(function(rows, source) {
    return rows.concat(getRowsAsObjects_(source.sheetName, source.headers));
  }, []).filter(function(row) {
    return normalizeSheetDate_(row.date).slice(0, 7) === monthKey;
  });
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(monthKey) || spreadsheet.insertSheet(monthKey);
  createOperationCalendarView_(sheet, reservations, roomLogs, parts[0], parts[1] - 1);
}

function createOperationCalendarView_(sheet, reservations, roomLogs, year, month) {
  sheet.getDataRange().breakApart();
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(3);

  sheet.getRange(1, 1, 1, 7).merge()
    .setValue(year + '년 ' + (month + 1) + '월 운영 현황')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF');

  const days = ['일', '월', '화', '수', '목', '금', '토'];
  sheet.getRange(3, 1, 1, 7).setValues([days])
    .setFontWeight('bold')
    .setBackground('#E8F0FE')
    .setHorizontalAlignment('center');
  sheet.getRange(3, 1).setFontColor('#D93025');
  sheet.getRange(3, 7).setFontColor('#1A73E8');

  const eventMap = {};
  const linkedReservationIds = {};
  const actualLogIds = {};
  (roomLogs || []).forEach(function(row) {
    const linkedReservationId = str_(row.linkedReservationId);
    const logId = str_(row.logId);
    if (linkedReservationId) linkedReservationIds[linkedReservationId] = true;
    if (logId) actualLogIds[logId] = true;
  });

  (reservations || []).forEach(function(reservation) {
    if (linkedReservationIds[str_(reservation.reservationId)] ||
        actualLogIds[str_(reservation.usageLogId)]) {
      return;
    }
    const date = normalizeSheetDate_(reservation.usageDate);
    const startTime = normalizeSheetTime_(reservation.startTime);
    const endTime = normalizeSheetTime_(reservation.endTime);
    if (!date || !startTime || !endTime) return;

    const statusLabel = isNoShowNaverStatus_(reservation.status)
      ? '[노쇼]'
      : (isCancelledNaverStatus_(reservation.status) ? '[취소]' : '[예약]');
    if (!eventMap[date]) eventMap[date] = [];
    eventMap[date].push({
      time: startTime + '~' + endTime,
      text: statusLabel + ' ' + getReservationSpaceLabel_(reservation.spaceType) +
        ' / ' + str_(reservation.bookerName),
      note: [
        '예약번호: ' + str_(reservation.reservationId),
        '상태: ' + str_(reservation.status),
        '인원: ' + Math.max(1, Number(reservation.headcount || 1)) + '명',
        '이용 목적: ' + str_(reservation.purpose, '네이버 예약 이용')
      ].join('\n'),
      isCancelled: isCancelledNaverStatus_(reservation.status),
      kind: 'reservation'
    });
  });

  (roomLogs || []).forEach(function(row) {
    const date = normalizeSheetDate_(row.date);
    const startTime = normalizeSheetTime_(row.startTime);
    const endTime = normalizeSheetTime_(row.endTime);
    if (!date || !startTime || !endTime) return;

    const linkedReservationId = str_(row.linkedReservationId);
    const headcount = Math.max(
      1,
      Number(row.headcount || 0),
      Number(row.companionCount || 0) + 1
    );
    if (!eventMap[date]) eventMap[date] = [];
    eventMap[date].push({
      time: startTime + '~' + endTime,
      text: (linkedReservationId ? '[이용완료] ' : '[현장이용] ') +
        getReservationSpaceLabel_(row.spaceType) + ' / ' + str_(row.name, '비회원'),
      note: [
        '기록 구분: ' + (linkedReservationId ? '네이버 예약 연계 이용' : '현장 이용'),
        linkedReservationId ? '예약번호: ' + linkedReservationId : '',
        '인원: ' + headcount + '명',
        '이용 목적: ' + str_(row.purpose, '시설 이용')
      ].filter(function(value) { return value; }).join('\n'),
      isCancelled: false,
      kind: 'usage'
    });
  });
  Object.keys(eventMap).forEach(function(date) {
    eventMap[date].sort(function(a, b) { return a.time.localeCompare(b.time); });
  });

  const firstDay = new Date(year, month, 1).getDay();
  const lastDay = new Date(year, month + 1, 0).getDate();
  let currentDay = 1 - firstDay;
  let currentRow = 4;

  while (currentDay <= lastDay) {
    const counts = [];
    for (let dayIndex = 0; dayIndex < 7; dayIndex += 1) {
      const day = currentDay + dayIndex;
      const date = monthKeyFromParts_(year, month + 1) + '-' + pad2_(day);
      counts.push(day > 0 && day <= lastDay ? (eventMap[date] || []).length : 0);
    }
    const eventRows = Math.max(3, Math.max.apply(null, counts));

    for (let dayIndex = 0; dayIndex < 7; dayIndex += 1) {
      const column = dayIndex + 1;
      const day = currentDay + dayIndex;
      const block = sheet.getRange(currentRow, column, eventRows + 1, 1);
      block.setBorder(true, true, true, true, null, null);

      if (day <= 0 || day > lastDay) {
        block.setBackground('#F5F5F5');
        continue;
      }

      const date = monthKeyFromParts_(year, month + 1) + '-' + pad2_(day);
      const events = eventMap[date] || [];
      sheet.getRange(currentRow, column)
        .setValue(day)
        .setFontWeight('bold')
        .setHorizontalAlignment('left')
        .setBackground('#FFFFFF')
        .setFontColor(dayIndex === 0 ? '#D93025' : (dayIndex === 6 ? '#1A73E8' : '#202124'));

      events.forEach(function(event, index) {
        const cell = sheet.getRange(currentRow + index + 1, column);
        cell.setValue(event.time + ' ' + event.text)
          .setNote(event.note)
          .setWrap(true)
          .setVerticalAlignment('middle')
          .setFontSize(9)
          .setBackground(event.isCancelled ? '#F1F3F4' : (event.kind === 'usage' ? '#E6F4EA' : '#FFF8E1'))
          .setFontColor(event.isCancelled ? '#80868B' : (event.kind === 'usage' ? '#137333' : '#202124'))
          .setFontLine(event.isCancelled ? 'line-through' : 'none');
      });
      if (events.length < eventRows) {
        sheet.getRange(currentRow + events.length + 1, column, eventRows - events.length, 1)
          .setBackground('#FFFFFF');
      }
    }

    sheet.setRowHeight(currentRow, 25);
    for (let offset = 1; offset <= eventRows; offset += 1) {
      sheet.setRowHeight(currentRow + offset, 45);
    }
    currentRow += eventRows + 1;
    currentDay += 7;
  }

  for (let column = 1; column <= 7; column += 1) {
    sheet.setColumnWidth(column, 150);
  }
}

function isValidMonthKey_(monthKey) {
  return /^\d{4}-(0[1-9]|1[0-2])$/.test(str_(monthKey));
}

function monthKeyFromParts_(year, month) {
  return year + '-' + pad2_(month);
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
  const rawBookerName = pickFirstValue_(values, NAVER_CSV_FIELDS.bookerName);
  const bookerName = cleanNaverBookerName_(rawBookerName) || rawBookerName;
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

// automation/naver-booker.js의 cleanNaverBookerName과 같은 규칙이다. 바꿀 때는 같이 바꾼다.
// 예) "김예약 [대리예약] 방문자: 이방문" → "이방문", "최관리 [관리자등록]" → "최관리"
function cleanNaverBookerName_(value) {
  let text = String(value || '').replace(/\s+/g, ' ').trim();
  text = text.replace(/01[0-9]-?\d{3,4}-?\d{4}.*$/, '').replace(/\d{9,}.*$/, '');
  const parts = text.split(/(?:예약자|방문자)\s*[:：]?/);
  for (let index = parts.length - 1; index >= 0; index -= 1) {
    const name = parts[index]
      .replace(/\[[^\]]*\]/g, ' ')
      .replace(/대리\s*예약|관리자\s*등록/g, ' ')
      .replace(/[^가-힣A-Za-z\s]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
    if (name) return name;
  }
  return '';
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
  const naver = findObjectsByColumnValue_(sheet, HEADERS.naverSync, 'usageDate', date)
    .filter(function(reservation) {
      return !isCancelledNaverStatus_(reservation.status) &&
        reservationBelongsToMember_(reservation, member);
    })
    .map(publicNaverReservation_);
  const web = getRowsAsObjects_(SHEET_NAMES.RESERVATION_REQUESTS, HEADERS.reservationRequests)
    .filter(function(reservation) {
      return normalizeSheetDate_(reservation.date) === date &&
        str_(reservation.memberId) === str_(member.id) &&
        visibleReservationRequestStatus_(reservation.status);
    })
    .map(publicReservationRequest_);
  return naver.concat(web)
    .sort(function(a, b) {
      return str_(a.startTime).localeCompare(str_(b.startTime));
    });
}

function reservationBelongsToMember_(reservation, member) {
  if (normalizeNameKey_(reservation.bookerName) !== normalizeNameKey_(member.name)) return false;
  const reservationPhoneTail = digits_(reservation.maskedPhone).slice(-4);
  return !reservationPhoneTail || reservationPhoneTail === str_(member.phoneLastDigits);
}

function publicNaverReservation_(reservation) {
  return {
    id: str_(reservation.reservationId),
    source: 'naver',
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

function publicReservationRequest_(reservation) {
  return {
    id: 'web:' + str_(reservation.requestId),
    source: 'reservationRequest',
    status: str_(reservation.status),
    spaceType: str_(reservation.spaceType),
    placeName: str_(reservation.placeName, getReservationSpaceLabel_(reservation.spaceType)),
    date: normalizeSheetDate_(reservation.date),
    startTime: normalizeSheetTime_(reservation.startTime),
    endTime: normalizeSheetTime_(reservation.endTime),
    headcount: Math.max(1, Number(reservation.headcount || 1)),
    purpose: str_(reservation.purpose, '웹 예약 이용'),
    companions: parseCompanionsJson_(reservation.companionsJson),
    submitted: Boolean(str_(reservation.usageLogId)) || str_(reservation.status) === '일지제출완료'
  };
}

function normalizeReservationSpaceType_(spaceType, placeId) {
  const candidate = str_(placeId) || str_(spaceType);
  if (candidate === 'careerRoom1' || candidate === 'careerRoom2' || candidate === 'connectRoom') return candidate;
  throw new Error('예약 공간을 확인해 주세요.');
}

function maxHeadcountForSpace_(spaceType) {
  if (spaceType === 'careerRoom1') return 1;
  if (spaceType === 'careerRoom2') return 2;
  if (spaceType === 'connectRoom') return 8;
  throw new Error('예약 공간을 확인해 주세요.');
}

function isValidReservationDate_(value) {
  const date = str_(value);
  const match = date.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match || date < today_()) return false;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const parsed = new Date(Date.UTC(year, month - 1, day));
  return year >= 2000 &&
    parsed.getUTCFullYear() === year &&
    parsed.getUTCMonth() === month - 1 &&
    parsed.getUTCDate() === day;
}

function visibleReservationRequestStatus_(status) {
  const value = str_(status);
  return value === '신청' || value === '승인' || value === '일지제출완료';
}

function activeReservationRequestStatus_(status) {
  const value = str_(status);
  return value === '신청' || value === '승인';
}

function blockingReservationRequestStatus_(status) {
  const value = str_(status);
  return value === '신청' || value === '승인';
}

function getPublicSpaceReservations_(spaceType, date) {
  const naverSheet = ensureNaverReservationSchema_();
  const naver = findObjectsByColumnValue_(naverSheet, HEADERS.naverSync, 'usageDate', date)
    .filter(function(reservation) {
      return roomResourcesConflict_(reservation.spaceType, spaceType) &&
        !isCancelledNaverStatus_(reservation.status);
    })
    .map(function(reservation) {
      return {
        id: 'naver:' + str_(reservation.reservationId),
        source: 'naver',
        status: str_(reservation.status),
        spaceType: str_(reservation.spaceType),
        placeName: getReservationSpaceLabel_(reservation.spaceType),
        date: normalizeSheetDate_(reservation.usageDate),
        startTime: normalizeSheetTime_(reservation.startTime),
        endTime: normalizeSheetTime_(reservation.endTime),
        headcount: safePositiveInteger_(reservation.headcount, 1)
      };
    });
  const web = getRowsAsObjects_(SHEET_NAMES.RESERVATION_REQUESTS, HEADERS.reservationRequests)
    .filter(function(reservation) {
      return normalizeSheetDate_(reservation.date) === date &&
        roomResourcesConflict_(reservation.spaceType, spaceType) &&
        blockingReservationRequestStatus_(reservation.status);
    })
    .map(function(reservation) {
      return {
        id: 'web:' + str_(reservation.requestId),
        source: 'reservationRequest',
        status: str_(reservation.status),
        spaceType: str_(reservation.spaceType),
        placeName: str_(reservation.placeName, getReservationSpaceLabel_(reservation.spaceType)),
        date: normalizeSheetDate_(reservation.date),
        startTime: normalizeSheetTime_(reservation.startTime),
        endTime: normalizeSheetTime_(reservation.endTime),
        headcount: safePositiveInteger_(reservation.headcount, 1)
      };
    });
  return naver.concat(web).sort(function(a, b) {
    return str_(a.startTime).localeCompare(str_(b.startTime));
  });
}

function findNaverReservationEntry_(sheet, reservationId) {
  const entries = findEntriesByColumnValue_(sheet, HEADERS.naverSync, 'reservationId', reservationId);
  return entries.length > 0 ? entries[0] : null;
}

function findReservationRequestEntry_(sheet, requestId) {
  const entries = findEntriesByColumnValue_(sheet, HEADERS.reservationRequests, 'requestId', requestId);
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

function updateReservationRequestRow_(sheet, rowNumber, changes) {
  const values = sheet.getRange(rowNumber, 1, 1, HEADERS.reservationRequests.length).getValues()[0];
  Object.keys(changes).forEach(function(header) {
    const column = HEADERS.reservationRequests.indexOf(header);
    if (column >= 0) values[column] = changes[header] || '';
  });
  sheet.getRange(rowNumber, 1, 1, HEADERS.reservationRequests.length).setValues([values]);
}

function assertNoReservationRequestConflict_(spaceType, date, startTime, endTime, ignoredRequestId) {
  const naverSheet = ensureNaverReservationSchema_();
  const naverConflicts = findOverlappingNaverReservationEntries_(naverSheet, spaceType, date, startTime, endTime);
  if (naverConflicts.length > 0) {
    throw new Error('해당 시간대에는 이미 예약이 있습니다.');
  }

  const roomLogSheetName = getRoomLogSheetName_(spaceType);
  const roomLogConflicts = getRowsAsObjects_(roomLogSheetName, HEADERS.roomLogs)
    .filter(function(row) {
      return normalizeSheetDate_(row.date) === date &&
        roomResourcesConflict_(row.spaceType, spaceType) &&
        timeRangesOverlap_(row.startTime, row.endTime, startTime, endTime);
    });
  if (roomLogConflicts.length > 0) {
    throw new Error('해당 시간대에는 이미 이용 기록이 있습니다.');
  }

  const requestConflicts = getRowsAsObjects_(SHEET_NAMES.RESERVATION_REQUESTS, HEADERS.reservationRequests)
    .filter(function(row) {
      return str_(row.requestId) !== str_(ignoredRequestId) &&
        normalizeSheetDate_(row.date) === date &&
        blockingReservationRequestStatus_(row.status) &&
        roomResourcesConflict_(row.spaceType, spaceType) &&
        timeRangesOverlap_(row.startTime, row.endTime, startTime, endTime);
    });
  if (requestConflicts.length > 0) {
    throw new Error('해당 시간대에는 이미 예약 신청이 있습니다.');
  }
}

function parseCompanionsJson_(value) {
  try {
    const parsed = JSON.parse(str_(value, '[]'));
    return normalizeCompanions_(Array.isArray(parsed) ? parsed : [], 10);
  } catch (error) {
    return [];
  }
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
  const allowed = [
    'facility-qr',
    'facility-phone-recheck',
    'facility-signup',
    'facility-signup-existing',
    'facility-birthdate-update',
    'admin-manual'
  ];
  return allowed.indexOf(source) >= 0 ? source : 'facility-qr';
}

function sha256Hex_(value) {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    str_(value),
    Utilities.Charset.UTF_8
  ).map(function(byte) {
    const unsigned = byte < 0 ? byte + 256 : byte;
    return ('0' + unsigned.toString(16)).slice(-2);
  }).join('');
}

function isValidGender_(value) {
  return value === '남성' || value === '여성';
}

function isValidPhone_(value) {
  return /^010\d{8}$/.test(digits_(value));
}

function calcAge_(birthDate) {
  return calcAgeOnDate_(birthDate, today_());
}

function calcAgeOnDate_(birthDate, referenceDate) {
  const birth = normalizeBirthDate_(birthDate).split('-').map(Number);
  const reference = normalizeSheetDate_(referenceDate).split('-').map(Number);
  if (birth.length !== 3 || reference.length !== 3 ||
      birth.some(function(value) { return !Number.isFinite(value); }) ||
      reference.some(function(value) { return !Number.isFinite(value); })) return 0;
  let age = reference[0] - birth[0];
  if (reference[1] < birth[1] ||
      (reference[1] === birth[1] && reference[2] < birth[2])) age -= 1;
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
