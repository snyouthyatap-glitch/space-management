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
        _httpClient.Timeout = TimeSpan.FromSeconds(25);
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

    public async Task<AdminMember> LookupMemberAsync(
        DashboardSettings settings,
        string password,
        string phone,
        CancellationToken cancellationToken)
    {
        var result = await ExecuteAdminActionAsync<AdminMemberLookup>(
            settings,
            password,
            "adminLookupMember",
            session => new { adminSession = session, phone },
            cancellationToken,
            response => response.Ok,
            response => response.Error);
        return result.Member;
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

    public void ResetSession() => _sessionToken = "";

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
