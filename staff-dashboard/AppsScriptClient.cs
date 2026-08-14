using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace YouthSpring.StaffDashboard;

internal sealed class AppsScriptClient
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private readonly HttpClient _httpClient;

    private string _sessionToken = "";

    public AppsScriptClient(HttpClient? httpClient = null)
    {
        _httpClient = httpClient ?? new HttpClient();
        _httpClient.Timeout = TimeSpan.FromSeconds(90);
    }

    public async Task<DashboardSnapshot> GetSnapshotAsync(
        DashboardSettings settings,
        string password,
        string calendarMonth,
        int eventLimit,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.AppsScriptUrl))
            throw new InvalidOperationException("Apps Script 주소를 설정해 주세요.");
        if (string.IsNullOrWhiteSpace(password))
            throw new InvalidOperationException("관리자 비밀번호를 설정해 주세요.");

        if (string.IsNullOrWhiteSpace(_sessionToken))
            await LoginAsync(settings.AppsScriptUrl, password, cancellationToken);

        var snapshot = await GetAsync<DashboardSnapshot>(
            settings.AppsScriptUrl,
            "adminDashboardSnapshot",
            new { adminSession = _sessionToken, limit = eventLimit, calendarMonth },
            cancellationToken);

        if (!snapshot.Ok && snapshot.Error.Contains("인증", StringComparison.Ordinal))
        {
            _sessionToken = "";
            await LoginAsync(settings.AppsScriptUrl, password, cancellationToken);
            snapshot = await GetAsync<DashboardSnapshot>(
                settings.AppsScriptUrl,
                "adminDashboardSnapshot",
                new { adminSession = _sessionToken, limit = eventLimit, calendarMonth },
                cancellationToken);
        }

        if (!snapshot.Ok)
            throw new InvalidOperationException(snapshot.Error);
        return snapshot;
    }

    public async Task<ReservationStatusUpdate> UpdateReservationStatusAsync(
        DashboardSettings settings,
        string password,
        string reservationId,
        string status,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.AppsScriptUrl))
            throw new InvalidOperationException("Apps Script 주소를 설정해 주세요.");
        if (string.IsNullOrWhiteSpace(password))
            throw new InvalidOperationException("관리자 비밀번호를 설정해 주세요.");

        if (string.IsNullOrWhiteSpace(_sessionToken))
            await LoginAsync(settings.AppsScriptUrl, password, cancellationToken);

        var result = await GetAsync<ReservationStatusUpdate>(
            settings.AppsScriptUrl,
            "adminUpdateReservationStatus",
            new { adminSession = _sessionToken, reservationId, status },
            cancellationToken);

        if (!result.Ok && result.Error.Contains("인증", StringComparison.Ordinal))
        {
            _sessionToken = "";
            await LoginAsync(settings.AppsScriptUrl, password, cancellationToken);
            result = await GetAsync<ReservationStatusUpdate>(
                settings.AppsScriptUrl,
                "adminUpdateReservationStatus",
                new { adminSession = _sessionToken, reservationId, status },
                cancellationToken);
        }

        if (!result.Ok)
            throw new InvalidOperationException(result.Error);
        return result;
    }

    public Task<ReservationStatusUpdate> UpdateReservationRequestStatusAsync(
        DashboardSettings settings,
        string password,
        string reservationRequestId,
        string status,
        CancellationToken cancellationToken) =>
        ExecuteAdminActionAsync<ReservationStatusUpdate>(
            settings,
            password,
            "adminUpdateReservationRequestStatus",
            session => new { adminSession = session, reservationRequestId, status },
            cancellationToken,
            response => response.Ok,
            response => response.Error);

    public Task<ReservationStatusUpdate> CompleteReservationUsageAsync(
        DashboardSettings settings,
        string password,
        CalendarEntry entry,
        string memberId,
        CancellationToken cancellationToken) =>
        ExecuteAdminActionAsync<ReservationStatusUpdate>(
            settings,
            password,
            "adminCompleteReservationUsage",
            session => new
            {
                adminSession = session,
                reservationId = StripKnownEntryPrefix(entry.Id),
                source = entry.Source,
                memberId,
                date = entry.Date,
                startTime = entry.StartTime,
                endTime = entry.EndTime,
                spaceType = entry.SpaceType,
                placeName = entry.PlaceName,
                purpose = entry.Purpose,
                headcount = entry.Headcount,
                companions = Array.Empty<ManualCompanion>()
            },
            cancellationToken,
            response => response.Ok,
            response => response.Error);

    public async Task<AdminMember> LookupMemberAsync(
        DashboardSettings settings,
        string password,
        string query,
        CancellationToken cancellationToken)
    {
        var members = await LookupMembersAsync(settings, password, query, cancellationToken);
        if (members.Count == 0) throw new InvalidOperationException("등록된 회원을 찾을 수 없습니다.");
        if (members.Count > 1) throw new InvalidOperationException("동명이인이 있습니다. 회원을 선택해 주세요.");
        return members[0];
    }

    public async Task<List<AdminMember>> LookupMembersAsync(
        DashboardSettings settings,
        string password,
        string query,
        CancellationToken cancellationToken)
    {
        var result = await ExecuteAdminActionAsync<AdminMemberLookup>(
            settings,
            password,
            "adminLookupMember",
            session => new { adminSession = session, query, name = query },
            cancellationToken,
            response => response.Ok,
            response => response.Error);
        if (result.Members.Count > 0) return result.Members;
        return string.IsNullOrWhiteSpace(result.Member.Id) ? [] : [result.Member];
    }

    public Task<ManualLogResult> SubmitManualLogAsync(
        DashboardSettings settings,
        string password,
        ManualLogRequest request,
        CancellationToken cancellationToken) =>
        ExecuteAdminActionAsync<ManualLogResult>(
            settings,
            password,
            "adminSubmitManualLog",
            session => new
            {
                adminSession = session,
                subjectType = request.SubjectType,
                memberId = request.MemberId,
                guest = request.Guest,
                logType = request.LogType,
                date = request.Date,
                data = request.Data,
                requestId = request.RequestId
            },
            cancellationToken,
            response => response.Ok,
            response => response.Error);

    public Task<MonthlyPerformanceReport> GetMonthlyPerformanceAsync(
        DashboardSettings settings,
        string password,
        string month,
        CancellationToken cancellationToken) =>
        ExecuteAdminActionAsync<MonthlyPerformanceReport>(
            settings,
            password,
            "adminMonthlyPerformance",
            session => new { adminSession = session, month },
            cancellationToken,
            response => response.Ok,
            response => response.Error);

    public void ResetSession() => _sessionToken = "";

    private static string StripKnownEntryPrefix(string id)
    {
        foreach (var prefix in new[] { "naver:", "naverSmartPlace:", "reservationRequest:", "webReservation:", "web:", "app:" })
        {
            if (id.StartsWith(prefix, StringComparison.Ordinal)) return id[prefix.Length..];
        }

        return id;
    }

    private async Task<T> ExecuteAdminActionAsync<T>(
        DashboardSettings settings,
        string password,
        string action,
        Func<string, object> payloadFactory,
        CancellationToken cancellationToken,
        Func<T, bool> isSuccess,
        Func<T, string> getError)
    {
        if (string.IsNullOrWhiteSpace(settings.AppsScriptUrl))
            throw new InvalidOperationException("Apps Script 주소를 설정해 주세요.");
        if (string.IsNullOrWhiteSpace(password))
            throw new InvalidOperationException("관리자 비밀번호를 설정해 주세요.");

        if (string.IsNullOrWhiteSpace(_sessionToken))
            await LoginAsync(settings.AppsScriptUrl, password, cancellationToken);

        var result = await GetAsync<T>(
            settings.AppsScriptUrl,
            action,
            payloadFactory(_sessionToken),
            cancellationToken);
        if (!isSuccess(result) && getError(result).Contains("인증", StringComparison.Ordinal))
        {
            _sessionToken = "";
            await LoginAsync(settings.AppsScriptUrl, password, cancellationToken);
            result = await GetAsync<T>(
                settings.AppsScriptUrl,
                action,
                payloadFactory(_sessionToken),
                cancellationToken);
        }

        if (!isSuccess(result))
            throw new InvalidOperationException(getError(result));
        return result;
    }

    private async Task LoginAsync(string url, string password, CancellationToken cancellationToken)
    {
        var challenge = await GetAsync<AdminChallenge>(
            url,
            "adminChallenge",
            new { },
            cancellationToken);
        if (!challenge.Ok) throw new InvalidOperationException(challenge.Error);

        var passwordHash = Sha256(password);
        var proof = Sha256(challenge.Challenge + ":" + passwordHash);
        var login = await GetAsync<AdminLogin>(
            url,
            "adminLogin",
            new { challenge = challenge.Challenge, proof },
            cancellationToken);
        if (!login.Ok) throw new InvalidOperationException(login.Error);
        _sessionToken = login.SessionToken;
    }

    private async Task<T> GetAsync<T>(
        string baseUrl,
        string action,
        object payload,
        CancellationToken cancellationToken)
    {
        var separator = baseUrl.Contains('?') ? '&' : '?';
        var requestUrl = baseUrl.Trim() + separator +
            "action=" + Uri.EscapeDataString(action) +
            "&payload=" + Uri.EscapeDataString(JsonSerializer.Serialize(payload));
        using var response = await _httpClient.GetAsync(requestUrl, cancellationToken);
        response.EnsureSuccessStatusCode();
        var json = await response.Content.ReadAsStringAsync(cancellationToken);
        return JsonSerializer.Deserialize<T>(json, JsonOptions)
            ?? throw new InvalidOperationException("Apps Script 응답을 읽지 못했습니다.");
    }

    private static string Sha256(string value)
    {
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(value));
        return Convert.ToHexString(hash).ToLowerInvariant();
    }
}
