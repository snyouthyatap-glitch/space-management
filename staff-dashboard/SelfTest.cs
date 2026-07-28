using System.Text.Json;
using System.Net;
using System.Security.Cryptography;
using System.Text;

namespace YouthSpring.StaffDashboard;

internal static class SelfTest
{
    public static void Run()
    {
        TestWindowsPasswordProtection();
        TestDashboardResponseParsing();
        TestSessionExpiryRecoveryAsync().GetAwaiter().GetResult();
        TestReservationStatusUpdateAsync().GetAwaiter().GetResult();
        TestManualLogFlowAsync().GetAwaiter().GetResult();
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
            snapshot.Events.Count != 1 ||
            !snapshot.Events[0].NotificationEligible ||
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

    private static async Task TestManualLogFlowAsync()
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
        var result = await client.SubmitManualLogAsync(
            settings,
            FakeAppsScriptHandler.Password,
            new ManualLogRequest
            {
                SubjectType = "guest",
                Guest = new ManualGuest
                {
                    Token = "guest-test-token",
                    Gender = "여성",
                    Age = 24
                },
                LogType = "connectRoom",
                Date = "2026-07-28",
                Data = new ManualLogData
                {
                    StartTime = "16:00",
                    EndTime = "17:00",
                    Purpose = "테스트",
                    Companions = [new ManualCompanion { Gender = "남성", Age = 25 }]
                },
                RequestId = "manual-test-request"
            },
            CancellationToken.None);

        if (member.Id != "member-1" ||
            !result.Ok ||
            !result.LogCreated ||
            handler.MemberLookupCount != 1 ||
            handler.ManualLogCount != 1 ||
            handler.LastManualLogType != "connectRoom" ||
            handler.LastManualCompanionCount != 1)
        {
            throw new InvalidOperationException("관리자 수동 일지 흐름 테스트에 실패했습니다.");
        }
    }

    private sealed class FakeAppsScriptHandler : HttpMessageHandler
    {
        public const string Password = "admin-test-password";
        public int ChallengeCount { get; private set; }
        public int LoginCount { get; private set; }
        public int SnapshotCount { get; private set; }
        public int StatusUpdateCount { get; private set; }
        public int MemberLookupCount { get; private set; }
        public int ManualLogCount { get; private set; }
        public string LastManualLogType { get; private set; } = "";
        public int LastManualCompanionCount { get; private set; }

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
            else if (action == "adminSubmitManualLog")
            {
                ManualLogCount++;
                LastManualLogType = payload.GetProperty("logType").GetString() ?? "";
                LastManualCompanionCount = payload
                    .GetProperty("data")
                    .GetProperty("companions")
                    .GetArrayLength();
                json = """{"ok":true,"date":"2026-07-28","logCreated":true,"linkedReservation":false}""";
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
