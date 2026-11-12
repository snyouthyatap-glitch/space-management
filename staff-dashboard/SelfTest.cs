using System.Text.Json;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.IO.Compression;
using System.Xml.Linq;

namespace YouthSpring.StaffDashboard;

internal static class SelfTest
{
    public static void Run()
    {
        TestWindowsPasswordProtection();
        TestDashboardResponseParsing();
        TestSessionExpiryRecoveryAsync().GetAwaiter().GetResult();
        TestReservationStatusUpdateAsync().GetAwaiter().GetResult();
        TestReservationRequestStatusUpdateAsync().GetAwaiter().GetResult();
        TestMemberLookupAsync().GetAwaiter().GetResult();
        TestMonthlyPerformanceAsync().GetAwaiter().GetResult();
        TestNaverBookingParserSeparatesReservationRows();
        TestNaverBookerNameCleanup();
        TestNaverProxyBookingRow();
        TestNaverTodayCancelRows();
        TestNaverSheetSyncItems();
        TestNaverChangedRowAnd24HourTime();
        TestNaverOtherProductsAreIgnored();
        TestReservationMergeByRepresentativeDateAndTime();
        TestMissingUsageLogStatus();
        TestOverdueUsageLogAlert();
        TestMergeDailyRows();
        TestPerformanceRangeRequestAsync().GetAwaiter().GetResult();
        TestPerformanceWorkbook();
    }

    private static void TestOverdueUsageLogAlert()
    {
        var now = new DateTime(2026, 11, 12, 14, 0, 0);
        const string today = "2026-11-12";
        CalendarEntry Entry(string id, string start, string status = "확정", bool submitted = false,
            bool cancelled = false, string source = "naverSmartPlace", string date = today) =>
            new() { Id = id, Source = source, Date = date, StartTime = start, EndTime = "23:00", Name = id, Status = status, Submitted = submitted, Cancelled = cancelled };

        var overdue = ReservationService.FindOverdueUsageLogs(
            [
                Entry("exactly-30", "13:30"),
                Entry("29-minutes", "13:31"),
                Entry("submitted", "10:00", submitted: true),
                Entry("usage-log", "10:00", source: "usage"),
                Entry("cancelled", "10:00", cancelled: true),
                Entry("marked-noshow", "10:00", status: "노쇼"),
                Entry("yesterday", "10:00", date: "2026-11-11"),
                Entry("web-request", "11:00", status: "승인", source: "reservationRequest")
            ],
            now);

        var ids = string.Join(",", overdue.Select(item => item.Id));
        if (ids != "web-request,exactly-30")
            throw new InvalidOperationException($"일지 미제출 30분 알림 대상 판별 테스트 실패: {ids}");
    }

    private static void TestMissingUsageLogStatus()
    {
        var date = DateTime.Today.AddDays(-1).ToString("yyyy-MM-dd");
        var checkedIn = new CalendarEntry
        {
            Id = "smart:1", Source = "naverSmartPlace", Date = date,
            StartTime = "10:00", EndTime = "11:00", Name = "김 봄", Status = "확정"
        };
        var absent = new CalendarEntry
        {
            Id = "smart:2", Source = "naverSmartPlace", Date = date,
            StartTime = "12:00", EndTime = "13:00", Name = "이청년", Status = "확정"
        };
        var webEntry = new CalendarEntry
        {
            Id = "web:1", Source = "reservationRequest", MemberId = "member-web", Date = date,
            StartTime = "14:00", EndTime = "15:00", Name = "박이봄", Status = "승인"
        };
        ReservationService.ApplyCheckIns(
            [checkedIn, absent, webEntry],
            [
                new StaffCheckIn { Date = date, Name = "김봄", MemberId = "member-1" },
                new StaffCheckIn { Date = date, Name = "박이봄", MemberId = "member-web" }
            ]);

        if (ReservationService.GetUnifiedStatus(checkedIn) != ReservationService.MissingUsageLogStatus ||
            checkedIn.MemberId != "member-1" ||
            ReservationService.GetUnifiedStatus(absent) != "노쇼" ||
            ReservationService.GetUnifiedStatus(webEntry) != ReservationService.MissingUsageLogStatus)
        {
            throw new InvalidOperationException("일지 미제출 상태 판별 테스트에 실패했습니다.");
        }
    }

    private static void TestWindowsPasswordProtection()
    {
        var storage = new AppStorage();
        const string expected = "test-password-1234";
        var protectedValue = storage.ProtectPassword(expected);
        var actual = storage.UnprotectPassword(protectedValue);
        if (actual != expected)
            throw new InvalidOperationException("Windows 비밀번호 보호 테스트에 실패했습니다.");
    }

    private static void TestDashboardResponseParsing()
    {
        const string json = """
        {
          "ok": true,
          "serverTime": "2026-07-28 16:00:00",
          "events": [{
            "id": "career:test",
            "category": "career",
            "categoryLabel": "커리어존",
            "notificationEligible": true,
            "origin": "onsite",
            "name": "테스트",
            "date": "2026-07-28",
            "startTime": "16:00",
            "endTime": "17:00",
            "purpose": "테스트",
            "headcount": 2,
            "occurredAt": "2026-07-28 15:59:00"
          }, {
            "id": "restricted:test",
            "category": "restricted",
            "categoryLabel": "이용 제한",
            "notificationEligible": true,
            "origin": "onsite",
            "name": "제한회원",
            "date": "2026-07-28",
            "purpose": "만 40세 · 시설 이용 가능 연령 초과",
            "headcount": 1,
            "occurredAt": "2026-07-28 15:58:00"
          }],
          "todayReservations": [],
          "summary": {
            "visitorStats": { "daily": 1, "weekly": 2, "monthly": 3 },
            "activeMemberCount": 4,
            "weekStart": "2026-07-27",
            "monthStart": "2026-07-01"
          }
        }
        """;
        var snapshot = JsonSerializer.Deserialize<DashboardSnapshot>(json);
        if (snapshot is null ||
            !snapshot.Ok ||
            snapshot.Events.Count != 2 ||
            !snapshot.Events[0].NotificationEligible ||
            snapshot.Events[1].Category != "restricted" ||
            snapshot.Summary.ActiveMemberCount != 4)
        {
            throw new InvalidOperationException("대시보드 응답 처리 테스트에 실패했습니다.");
        }
    }

    private static async Task TestSessionExpiryRecoveryAsync()
    {
        var handler = new FakeAppsScriptHandler();
        using var httpClient = new HttpClient(handler);
        var client = new AppsScriptClient(httpClient);
        var snapshot = await client.GetSnapshotAsync(
            new DashboardSettings { AppsScriptUrl = "https://example.test/exec" },
            FakeAppsScriptHandler.Password,
            "2026-07",
            100,
            CancellationToken.None);

        if (!snapshot.Ok ||
            handler.ChallengeCount != 2 ||
            handler.LoginCount != 2 ||
            handler.SnapshotCount != 2)
        {
            throw new InvalidOperationException("관리자 세션 만료 복구 테스트에 실패했습니다.");
        }
    }

    private static async Task TestReservationStatusUpdateAsync()
    {
        var handler = new FakeAppsScriptHandler();
        using var httpClient = new HttpClient(handler);
        var client = new AppsScriptClient(httpClient);
        var result = await client.UpdateReservationStatusAsync(
            new DashboardSettings { AppsScriptUrl = "https://example.test/exec" },
            FakeAppsScriptHandler.Password,
            "reservation-123",
            "노쇼",
            CancellationToken.None);

        if (!result.Ok ||
            result.ReservationId != "reservation-123" ||
            result.Status != "노쇼" ||
            !result.Cancelled ||
            handler.StatusUpdateCount != 1)
        {
            throw new InvalidOperationException("예약 상태 변경 요청 테스트에 실패했습니다.");
        }
    }

    private static async Task TestReservationRequestStatusUpdateAsync()
    {
        var handler = new FakeAppsScriptHandler();
        using var httpClient = new HttpClient(handler);
        var client = new AppsScriptClient(httpClient);
        var result = await client.UpdateReservationRequestStatusAsync(
            new DashboardSettings { AppsScriptUrl = "https://example.test/exec" },
            FakeAppsScriptHandler.Password,
            "request-123",
            "승인",
            CancellationToken.None);

        if (!result.Ok ||
            result.ReservationId != "request-123" ||
            result.Status != "승인" ||
            handler.ReservationRequestStatusUpdateCount != 1)
        {
            throw new InvalidOperationException("웹 예약 신청 상태 변경 요청 테스트에 실패했습니다.");
        }
    }

    private static async Task TestMemberLookupAsync()
    {
        var handler = new FakeAppsScriptHandler();
        using var httpClient = new HttpClient(handler);
        var client = new AppsScriptClient(httpClient);
        var settings = new DashboardSettings { AppsScriptUrl = "https://example.test/exec" };
        var member = await client.LookupMemberAsync(
            settings,
            FakeAppsScriptHandler.Password,
            "01012345678",
            CancellationToken.None);

        if (member.Id != "member-1" || handler.MemberLookupCount != 1)
            throw new InvalidOperationException("관리자 회원 조회 테스트에 실패했습니다.");
    }

    private static async Task TestMonthlyPerformanceAsync()
    {
        var handler = new FakeAppsScriptHandler();
        using var httpClient = new HttpClient(handler);
        var client = new AppsScriptClient(httpClient);
        var report = await client.GetMonthlyPerformanceAsync(
            new DashboardSettings { AppsScriptUrl = "https://example.test/exec" },
            FakeAppsScriptHandler.Password,
            "2026-07",
            CancellationToken.None);

        if (!report.Ok || report.Month != "2026-07" || report.Summary.Count != 4 ||
            report.TotalUsers != 12 || report.TotalUniqueUsers != 9 ||
            report.Daily.Visit.Count != 1 || report.Daily.Connect[0].Male20To24 != 2 ||
            report.Simplified.Career.Count != 2 || report.Simplified.Career[0].Space != "Room1" ||
            handler.MonthlyPerformanceCount != 1)
        {
            throw new InvalidOperationException("월간 실적 응답 처리 테스트에 실패했습니다.");
        }
    }

    private static void TestNaverBookingParserSeparatesReservationRows()
    {
        var path = Path.Combine(Path.GetTempPath(), "naver-booking-parser-test.json");
        const string rawText =
            "확정안정준010-2352-3615131865724526. 8. 14.(금) 오후 6:00~9:00청년 커넥트룸" +
            "확정최소연010-6258-7033131486876926. 8. 14.(금) 오후 4:00~6:00AI 커리어존 (청취Room1)";
        try
        {
            File.WriteAllText(path, $$"""
            {
              "results": {
                "today_use": [{
                  "reservationNo": "1318657245",
                  "bookerName": "안정준",
                  "productName": "청년 커넥트룸",
                  "status": "확정",
                  "rawText": "{{rawText}}"
                }, {
                  "reservationNo": "1314868769",
                  "bookerName": "안정준",
                  "productName": "청년 커넥트룸",
                  "status": "확정",
                  "rawText": "{{rawText}}"
                }],
                "pending_confirm": []
              }
            }
            """);

            var snapshot = NaverBookingParser.ReadSnapshot(path);
            if (snapshot.TodayUse.Count != 2)
                throw new InvalidOperationException("네이버 예약 파서 테스트에 실패했습니다: 예약 수가 다릅니다.");

            var first = NaverBookingParser.ConvertTodayUseReservation(snapshot.TodayUse[0]);
            var second = NaverBookingParser.ConvertTodayUseReservation(snapshot.TodayUse[1]);
            if (first.Name != "안정준" || first.PlaceName != "커넥트룸" || first.StartTime != "18:00" ||
                second.Name != "최소연" || second.PlaceName != "커리어존 Room1" || second.StartTime != "16:00")
            {
                throw new InvalidOperationException("네이버 예약번호별 파싱 회귀 테스트에 실패했습니다.");
            }
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // automation/naver-booker.test.js의 NAME_CASES와 같은 사례다.
    private static void TestNaverBookerNameCleanup()
    {
        (string Input, string Expected)[] cases =
        [
            ("한예시", "한예시"),
            ("김예약 [대리예약] 방문자: 이방문", "이방문"),
            ("김예약[대리예약]방문자:이방문", "이방문"),
            ("김예약대리예약방문자 이방문", "이방문"),
            ("김예약대리예약방문자: 이방문", "이방문"),
            ("이방문(010-1234-5678)", "이방문"),
            ("김예약대리예약", "김예약"),
            ("최관리 [관리자등록]", "최관리"),
            ("최관리관리자등록", "최관리"),
            ("예약자: 홍길동", "홍길동"),
            ("방문자：김봄", "김봄"),
            ("홍길동!!", "홍길동"),
            ("김 봄 010-1234-5678", "김 봄"),
            ("John Smith", "John Smith"),
            ("[관리자등록]", ""),
            ("", "")
        ];
        foreach (var (input, expected) in cases)
        {
            var actual = NaverBookingParser.NormalizeBookerName(input);
            if (actual != expected)
                throw new InvalidOperationException($"네이버 예약자명 정리 테스트 실패: '{input}' → '{actual}' (기대값 '{expected}')");
        }
    }

    // 실제 화면 형태: '변경' 표시, '26.11.11.수', 24시간제 '17:30~19:30'
    private static void TestNaverChangedRowAnd24HourTime()
    {
        var path = Path.Combine(Path.GetTempPath(), "naver-booking-changed-test.json");
        try
        {
            File.WriteAllText(path, """
            {
              "results": {
                "today_use": [{
                  "reservationNo": "1400000496", "status": "완료",
                  "rawText": "완료정변경010-0000-12341400000496변경26.11.11.수17:30~19:30AI 커리어존 (청취Room1)1-0원26.11.8.일00:30"
                }, {
                  "reservationNo": "1400000141", "status": "확정",
                  "rawText": "확정김예약대리예약방문자: 이방문010-1111-2222010-3333-4444140000014113:30~15:30AI 커리어존 (청취Room1)1-0원"
                }],
                "pending_confirm": [],
                "today_cancel": []
              }
            }
            """);
            var snapshot = NaverBookingParser.ReadSnapshot(path);
            var changed = NaverBookingParser.ConvertTodayUseReservation(snapshot.TodayUse[0]);
            var today = NaverBookingParser.ConvertTodayUseReservation(snapshot.TodayUse[1]);
            if (!snapshot.TodayUse[0].Changed || snapshot.TodayUse[1].Changed ||
                changed.Date != "2026-11-11" || changed.StartTime != "17:30" || changed.EndTime != "19:30" ||
                changed.Name != "정변경" || changed.PlaceName != "커리어존 Room1" ||
                today.StartTime != "13:30" || today.EndTime != "15:30" || today.Name != "이방문")
            {
                throw new InvalidOperationException(
                    $"네이버 변경·24시간제 파싱 테스트 실패: {changed.Date} {changed.StartTime}~{changed.EndTime} {changed.Name} / {today.StartTime}~{today.EndTime} {today.Name}");
            }
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // 실제 목록에는 [동아리실] 같은 다른 상품도 섞여 나온다(2026-11-12 화면 캡처 형태, 이름·번호는 가짜).
    private static void TestNaverOtherProductsAreIgnored()
    {
        var path = Path.Combine(Path.GetTempPath(), "naver-booking-other-product-test.json");
        try
        {
            File.WriteAllText(path, """
            {
              "results": {
                "today_use": [{
                  "reservationNo": "1400000452", "status": "완료", "productName": "[동아리실] 3층 회의실",
                  "rawText": "완료김예약대리예약방문자: 이방문010-1111-2222010-3333-4444140000045226.11.11.수16:00~16:30[동아리실] 3층 회의실1-0원26.11.10.화16:47"
                }, {
                  "reservationNo": "1400000496", "status": "완료",
                  "rawText": "완료정변경010-0000-12341400000496변경26.11.11.수17:30~19:30AI 커리어존 (청취Room1)1-0원26.11.8.일00:30"
                }],
                "pending_confirm": [{
                  "reservationNo": "1400000999", "status": "신청", "productName": "[동아리실] 3층 회의실",
                  "rawText": "신청홍길동010-5555-6666140000099926.11.20.금10:00~12:00[동아리실] 3층 회의실1-0원"
                }],
                "today_cancel": []
              }
            }
            """);
            var snapshot = NaverBookingParser.ReadSnapshot(path);
            var entries = NaverBookingParser.BuildSmartPlaceEntries(snapshot);
            var syncItems = NaverBookingParser.BuildSheetSyncItems(snapshot);
            if (entries.Count != 1 || entries[0].Name != "정변경" ||
                syncItems.Count != 1 || snapshot.PendingConfirm.Count != 0)
            {
                throw new InvalidOperationException(
                    $"다른 상품 제외 테스트 실패: entries={entries.Count}, sync={syncItems.Count}, pending={snapshot.PendingConfirm.Count}");
            }
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static void TestNaverSheetSyncItems()
    {
        var path = Path.Combine(Path.GetTempPath(), "naver-booking-sheet-sync-test.json");
        try
        {
            File.WriteAllText(path, """
            {
              "results": {
                "today_use": [{
                  "reservationNo": "3000000001", "bookerName": "김확정", "status": "확정", "phone": "010-1111-2222",
                  "productName": "청년 커넥트룸", "useDateTime": "26. 11. 12.(목) 오후 2:00~3:00", "headcount": "3"
                }, {
                  "reservationNo": "3000000002", "bookerName": "이신청", "status": "신청",
                  "productName": "청년 커넥트룸", "useDateTime": "26. 11. 12.(목) 오후 4:00~5:00"
                }, {
                  "reservationNo": "3000000003", "bookerName": "시간없음", "status": "확정",
                  "productName": "청년 커넥트룸", "useDateTime": ""
                }],
                "pending_confirm": [],
                "today_cancel": [{
                  "reservationNo": "3000000004", "bookerName": "박취소", "status": "취소",
                  "productName": "AI 커리어존 (청취Room2)", "useDateTime": "26. 11. 13.(금) 오전 10:00~11:00"
                }]
              }
            }
            """);
            var items = NaverBookingParser.BuildSheetSyncItems(NaverBookingParser.ReadSnapshot(path));
            var summary = string.Join(",", items.Select(item => $"{item.Reservation.ReservationNo}|{item.Status}"));
            var confirmed = items.FirstOrDefault(item => item.Status == "확정");
            if (summary != "3000000001|확정,3000000004|취소" ||
                confirmed is null || confirmed.Reservation.Phone != "010-1111-2222" ||
                confirmed.Entry.SpaceType != "connectRoom" || confirmed.Entry.Headcount != 3 ||
                items[1].Entry.Date != "2026-11-13" || items[1].Entry.SpaceType != "careerRoom2")
            {
                throw new InvalidOperationException($"네이버 예약 시트 반영 대상 테스트 실패: {summary}");
            }
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static void TestNaverTodayCancelRows()
    {
        var path = Path.Combine(Path.GetTempPath(), "naver-booking-cancel-test.json");
        try
        {
            // 김봄: 취소 후 같은 시간에 다시 예약(새 예약번호) → 새 예약은 '확정'으로 남아야 한다.
            // 이취소: 취소만 있음 → '취소'로 보여야 한다.
            File.WriteAllText(path, """
            {
              "results": {
                "today_use": [{
                  "reservationNo": "2000000002", "bookerName": "김봄", "status": "확정",
                  "productName": "청년 커넥트룸", "useDateTime": "26. 11. 12.(목) 오후 2:00~3:00"
                }],
                "pending_confirm": [],
                "today_cancel": [{
                  "reservationNo": "2000000001", "bookerName": "김봄", "status": "취소",
                  "productName": "청년 커넥트룸", "useDateTime": "26. 11. 12.(목) 오후 2:00~3:00"
                }, {
                  "reservationNo": "2000000003", "bookerName": "이취소", "status": "취소",
                  "productName": "AI 커리어존 (청취Room1)", "useDateTime": "26. 11. 12.(목) 오후 4:00~5:00"
                }]
              }
            }
            """);
            var entries = NaverBookingParser.BuildSmartPlaceEntries(NaverBookingParser.ReadSnapshot(path));
            var merged = ReservationService.MergeCalendarEntries([], entries, []);
            var kim = merged.SingleOrDefault(item => item.Name == "김봄");
            var lee = merged.SingleOrDefault(item => item.Name == "이취소");
            if (merged.Count != 2 ||
                kim is null || kim.Cancelled || kim.Id != "naverSmartPlace:2000000002" ||
                lee is null || !lee.Cancelled || ReservationService.GetUnifiedStatus(lee) != "취소" ||
                lee.StartTime != "16:00" || lee.Date != "2026-11-12")
            {
                throw new InvalidOperationException("네이버 오늘취소 표시 테스트에 실패했습니다.");
            }
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static void TestNaverProxyBookingRow()
    {
        var parsed = NaverBookingParser.ParseReservationText(new NaverBookingReservation
        {
            ReservationNo = "1318657245",
            // 실제 목록 화면 형태(2026-11-10 수집본, 번호는 가짜)
            RawText = "확정김예약대리예약방문자: 이방문010-1111-2222010-3333-4444131865724526. 8. 14.(금) 오후 6:00~9:00청년 커넥트룸"
        });
        var entry = NaverBookingParser.ConvertTodayUseReservation(new NaverBookingReservation
        {
            ReservationNo = "1318657245",
            RawText = "확정김예약[대리예약]방문자: 이방문010-1111-2222010-3333-4444131865724526. 8. 14.(금) 오후 6:00~9:00청년 커넥트룸"
        });
        if (parsed.Status != "확정" || entry.Name != "이방문" || entry.PlaceName != "커넥트룸" || entry.StartTime != "18:00")
            throw new InvalidOperationException($"네이버 대리예약 행 파싱 테스트 실패: '{entry.Name}' {entry.PlaceName} {entry.StartTime}");
    }

    private static void TestReservationMergeByRepresentativeDateAndTime()
    {
        var merged = ReservationService.MergeCalendarEntries(
            [
                new CalendarEntry
                {
                    Id = "naver:1",
                    Source = "naverSmartPlace",
                    Date = "2026-08-14",
                    StartTime = "16:00",
                    EndTime = "18:00",
                    Name = "안정준",
                    PlaceName = "커넥트룸",
                    Status = "확정"
                }
            ],
            [],
            [
                new CalendarEntry
                {
                    Id = "web:1",
                    Source = "reservationRequest",
                    Date = "2026-08-14",
                    StartTime = "16:00",
                    EndTime = "18:00",
                    Name = "안정준",
                    PlaceName = "커넥트룸",
                    Status = "승인"
                }
            ]);

        if (merged.Count != 1 ||
            merged[0].Source != "reservationRequest" ||
            merged[0].Name != "안정준" ||
            merged[0].PlaceName != "커넥트룸")
        {
            throw new InvalidOperationException("동일 대표자·날짜·시간 예약 병합 테스트에 실패했습니다.");
        }
    }

    private static void TestMergeDailyRows()
    {
        var merged = PerformanceWorkbookWriter.MergeDailyRows(
            [
                new() { Date = "2026-07-02", Male20To24 = 1, Total = 1 },
                new() { Date = "2026-07-01", Female25To29 = 2, Total = 2 }
            ],
            [new() { Date = "2026-07-01", Female25To29 = 1, Unknown = 1, Total = 2 }],
            [],
            [new() { Date = "2026-07-03", Male40Plus = 1, Total = 1 }]);
        if (merged.Count != 3 ||
            merged[0].Date != "2026-07-01" || merged[0].Female25To29 != 3 || merged[0].Unknown != 1 || merged[0].Total != 4 ||
            merged[1].Date != "2026-07-02" || merged[1].Total != 1 ||
            merged[2].Date != "2026-07-03" || merged[2].Male40Plus != 1)
        {
            throw new InvalidOperationException("전체 일별 현황 합산 테스트에 실패했습니다.");
        }
    }

    private static async Task TestPerformanceRangeRequestAsync()
    {
        var handler = new FakeAppsScriptHandler();
        using var httpClient = new HttpClient(handler);
        var client = new AppsScriptClient(httpClient);
        var report = await client.GetPerformanceForRangeAsync(
            new DashboardSettings { AppsScriptUrl = "https://example.test/exec" },
            FakeAppsScriptHandler.Password,
            "2026-06-15",
            "2026-07-14",
            CancellationToken.None);
        if (!report.Ok || handler.LastPerformanceStartDate != "2026-06-15" ||
            handler.LastPerformanceEndDate != "2026-07-14" ||
            report.PeriodTitle != "2026-06-15 ~ 2026-07-14")
        {
            throw new InvalidOperationException("기간 지정 실적 요청 테스트에 실패했습니다.");
        }
    }

    private static void TestPerformanceWorkbook()
    {
        var path = Path.Combine(Path.GetTempPath(), "youth-spring-performance-test.xlsx");
        try
        {
            PerformanceWorkbookWriter.Write(path, CreatePerformanceReport());
            using var archive = ZipFile.OpenRead(path);
            foreach (var entry in archive.Entries.Where(entry =>
                         entry.FullName.EndsWith(".xml", StringComparison.Ordinal)))
            {
                using var stream = entry.Open();
                _ = XDocument.Load(stream);
            }
            var worksheetCount = archive.Entries.Count(entry =>
                entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal));
            var workbookEntry = archive.GetEntry("xl/workbook.xml")
                ?? throw new InvalidOperationException("엑셀 통합 문서 정보가 없습니다.");
            using var reader = new StreamReader(workbookEntry.Open());
            var workbookXml = reader.ReadToEnd();
            var careerSimplifiedEntry = archive.GetEntry("xl/worksheets/sheet9.xml")
                ?? throw new InvalidOperationException("커리어존 간소화 시트가 없습니다.");
            using var careerReader = new StreamReader(careerSimplifiedEntry.Open());
            var careerXml = careerReader.ReadToEnd();
            var totalDailyEntry = archive.GetEntry("xl/worksheets/sheet2.xml")
                ?? throw new InvalidOperationException("전체 일별 현황 시트가 없습니다.");
            using var totalReader = new StreamReader(totalDailyEntry.Open());
            var totalXml = totalReader.ReadToEnd();
            // 4개 부문 2026-07-01 합계: 방문자 1 + 커넥트룸 2 + 커리어존 1 + 프린터 1 = 5
            if (!workbookXml.Contains("전체 일별 현황", StringComparison.Ordinal) ||
                workbookXml.IndexOf("전체 일별 현황", StringComparison.Ordinal) >
                workbookXml.IndexOf("\"방문자\"", StringComparison.Ordinal) ||
                !totalXml.Contains("2026-07-01", StringComparison.Ordinal) ||
                !totalXml.Contains("<v>5</v>", StringComparison.Ordinal))
            {
                throw new InvalidOperationException("전체 일별 현황 시트 테스트에 실패했습니다.");
            }
            if (worksheetCount != 10 || !workbookXml.Contains("월간 요약", StringComparison.Ordinal) ||
                !workbookXml.Contains("커넥트룸", StringComparison.Ordinal) ||
                !workbookXml.Contains("커리어존 간소화", StringComparison.Ordinal) ||
                !careerXml.Contains("사용공간", StringComparison.Ordinal) ||
                !careerXml.Contains("Room1", StringComparison.Ordinal) ||
                !careerXml.Contains("Room2", StringComparison.Ordinal))
            {
                throw new InvalidOperationException("월간 실적 엑셀 생성 테스트에 실패했습니다.");
            }
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static MonthlyPerformanceReport CreatePerformanceReport() => new()
    {
        Ok = true,
        Month = "2026-07",
        GeneratedAt = "2026-08-01 12:00:00",
        Summary =
        [
            new() { Key = "visit", Label = "방문자", UniqueUsers = 3, TotalUsers = 4 },
            new() { Key = "connect", Label = "커넥트룸", UniqueUsers = 2, TotalUsers = 3 },
            new() { Key = "career", Label = "커리어존", UniqueUsers = 2, TotalUsers = 3 },
            new() { Key = "printer", Label = "프린터", UniqueUsers = 2, TotalUsers = 2 }
        ],
        TotalUsers = 12,
        TotalUniqueUsers = 9,
        UniqueBasis = "회원 ID 월간 중복 제거, 식별 불가 비회원 및 동반자는 기록별 인원",
        Daily = new PerformanceDailyGroups
        {
            Visit = CreateDailyRows(1),
            Connect = CreateDailyRows(2),
            Career = CreateDailyRows(1),
            Printer = CreateDailyRows(1)
        },
        Simplified = new PerformanceSimplifiedGroups
        {
            Visit =
            [
                new() { Date = "2026-07-01", Gender = "남성", Age = 24 }
            ],
            Printer =
            [
                new() { Date = "2026-07-01", Gender = "여성", Age = 29, Count = 2 }
            ],
            Career =
            [
                new()
                {
                    Date = "2026-07-01", Space = "Room1", Purpose = "상담",
                    Headcount = 1, Note = "남성 / 24", StartTime = "10:00", EndTime = "11:00",
                    Male20s = 1
                },
                new()
                {
                    Date = "2026-07-02", Space = "Room2", Purpose = "회의",
                    Headcount = 2, Note = "여성 / 28, 여성 / 18", StartTime = "13:00", EndTime = "14:00",
                    Female20s = 1, FemaleUnder19 = 1
                }
            ],
            Connect =
            [
                new()
                {
                    Date = "2026-07-03", Purpose = "모임", Headcount = 1,
                    Note = "남성 / 31", StartTime = "15:00", EndTime = "16:00", Male30s = 1
                }
            ]
        }
    };

    private static List<PerformanceDailyRow> CreateDailyRows(int firstDayMale20To24) =>
        [new PerformanceDailyRow
        {
            Date = "2026-07-01",
            Male20To24 = firstDayMale20To24,
            Total = firstDayMale20To24
        }];

    private sealed class FakeAppsScriptHandler : HttpMessageHandler
    {
        public const string Password = "admin-test-password";
        public int ChallengeCount { get; private set; }
        public int LoginCount { get; private set; }
        public int SnapshotCount { get; private set; }
        public int StatusUpdateCount { get; private set; }
        public int ReservationRequestStatusUpdateCount { get; private set; }
        public int MemberLookupCount { get; private set; }
        public int MonthlyPerformanceCount { get; private set; }
        public string LastPerformanceStartDate { get; private set; } = "";
        public string LastPerformanceEndDate { get; private set; } = "";

        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken)
        {
            var query = ParseQuery(request.RequestUri!);
            var action = query.GetValueOrDefault("action", "");
            var payload = JsonDocument.Parse(query.GetValueOrDefault("payload", "{}")).RootElement;
            string json;

            if (action == "adminChallenge")
            {
                ChallengeCount++;
                json = JsonSerializer.Serialize(new
                {
                    ok = true,
                    challenge = "challenge-" + ChallengeCount
                });
            }
            else if (action == "adminLogin")
            {
                LoginCount++;
                var challenge = payload.GetProperty("challenge").GetString() ?? "";
                var proof = payload.GetProperty("proof").GetString() ?? "";
                var expected = Sha256(challenge + ":" + Sha256(Password));
                if (proof != expected)
                    throw new InvalidOperationException("관리자 인증 증명 생성이 올바르지 않습니다.");
                json = JsonSerializer.Serialize(new
                {
                    ok = true,
                    sessionToken = "session-" + LoginCount
                });
            }
            else if (action == "adminDashboardSnapshot")
            {
                SnapshotCount++;
                json = SnapshotCount == 1
                    ? """{"ok":false,"error":"관리자 인증이 만료되었습니다. 다시 로그인해 주세요."}"""
                    : """
                      {
                        "ok": true,
                        "serverTime": "2026-07-28 16:00:00",
                        "events": [],
                        "todayReservations": [],
                        "calendarEntries": [],
                        "reservationRequests": [
                          {
                            "id": "request-1",
                            "status": "신청",
                            "bookerName": "테스트 회원",
                            "date": "2026-07-28",
                            "startTime": "14:00",
                            "endTime": "15:00",
                            "placeName": "커리어존 Room1",
                            "headcount": 1,
                            "purpose": "상담"
                          }
                        ],
                        "summary": {
                          "visitorStats": { "daily": 1, "weekly": 2, "monthly": 3 },
                          "activeMemberCount": 4
                        }
                      }
                      """;
            }
            else if (action == "adminUpdateReservationStatus")
            {
                StatusUpdateCount++;
                var reservationId = payload.GetProperty("reservationId").GetString() ?? "";
                var status = payload.GetProperty("status").GetString() ?? "";
                json = JsonSerializer.Serialize(new
                {
                    ok = true,
                    reservationId,
                    status,
                    cancelled = status is "취소" or "노쇼"
                });
            }
            else if (action == "adminUpdateReservationRequestStatus")
            {
                ReservationRequestStatusUpdateCount++;
                var reservationRequestId = payload.GetProperty("reservationRequestId").GetString() ?? "";
                var status = payload.GetProperty("status").GetString() ?? "";
                json = JsonSerializer.Serialize(new
                {
                    ok = true,
                    reservationId = reservationRequestId,
                    status,
                    cancelled = status == "반려"
                });
            }
            else if (action == "adminLookupMember")
            {
                MemberLookupCount++;
                json = """
                  {
                    "ok": true,
                    "member": {
                      "id": "member-1",
                      "name": "테스트 회원",
                      "gender": "여성",
                      "age": 24,
                      "birthDate": "2002-01-01",
                      "phone": "01012345678"
                    }
                  }
                  """;
            }
            else if (action == "adminMonthlyPerformance")
            {
                MonthlyPerformanceCount++;
                var report = CreatePerformanceReport();
                if (payload.TryGetProperty("startDate", out var startDate))
                {
                    LastPerformanceStartDate = startDate.GetString() ?? "";
                    LastPerformanceEndDate = payload.GetProperty("endDate").GetString() ?? "";
                    report.Month = "";
                    report.StartDate = LastPerformanceStartDate;
                    report.EndDate = LastPerformanceEndDate;
                }
                json = JsonSerializer.Serialize(report);
            }
            else
            {
                json = """{"ok":false,"error":"unexpected action"}""";
            }

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(json, Encoding.UTF8, "application/json")
            });
        }

        private static Dictionary<string, string> ParseQuery(Uri uri) =>
            uri.Query.TrimStart('?')
                .Split('&', StringSplitOptions.RemoveEmptyEntries)
                .Select(part => part.Split('=', 2))
                .ToDictionary(
                    part => Uri.UnescapeDataString(part[0]),
                    part => Uri.UnescapeDataString(part.Length > 1 ? part[1] : ""),
                    StringComparer.Ordinal);

        private static string Sha256(string value) =>
            Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
    }
}
