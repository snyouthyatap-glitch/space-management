using System.Text.Json.Serialization;

namespace YouthSpring.StaffDashboard;

internal sealed class DashboardSettings
{
    public string AppsScriptUrl { get; set; } =
        "https://script.google.com/macros/s/AKfycbyYkhXruzoj25GmeWRqlKwz9kcLlHPk5BOWxsfHc_gcsEyVYpiXXlgkwIrtAWaGTizk/exec";
    public string ProtectedPassword { get; set; } = "";
    public int PollSeconds { get; set; } = 30;
}

internal sealed class LocalState
{
    public bool InitialSyncCompleted { get; set; }
    public List<StaffEvent> Events { get; set; } = [];
}

internal sealed class StaffEvent
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = "";

    [JsonPropertyName("category")]
    public string Category { get; set; } = "";

    [JsonPropertyName("categoryLabel")]
    public string CategoryLabel { get; set; } = "";

    [JsonPropertyName("notificationEligible")]
    public bool NotificationEligible { get; set; }

    [JsonPropertyName("origin")]
    public string Origin { get; set; } = "";

    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("date")]
    public string Date { get; set; } = "";

    [JsonPropertyName("startTime")]
    public string StartTime { get; set; } = "";

    [JsonPropertyName("endTime")]
    public string EndTime { get; set; } = "";

    [JsonPropertyName("purpose")]
    public string Purpose { get; set; } = "";

    [JsonPropertyName("headcount")]
    public int Headcount { get; set; } = 1;

    [JsonPropertyName("occurredAt")]
    public string OccurredAt { get; set; } = "";

    public bool IsRead { get; set; }
}

internal sealed class Reservation
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = "";

    [JsonPropertyName("status")]
    public string Status { get; set; } = "";

    [JsonPropertyName("bookerName")]
    public string BookerName { get; set; } = "";

    [JsonPropertyName("placeName")]
    public string PlaceName { get; set; } = "";

    [JsonPropertyName("startTime")]
    public string StartTime { get; set; } = "";

    [JsonPropertyName("endTime")]
    public string EndTime { get; set; } = "";

    [JsonPropertyName("headcount")]
    public int Headcount { get; set; } = 1;

    [JsonPropertyName("purpose")]
    public string Purpose { get; set; } = "";

    [JsonPropertyName("cancelled")]
    public bool Cancelled { get; set; }

    [JsonPropertyName("submitted")]
    public bool Submitted { get; set; }
}

internal sealed class CalendarEntry
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = "";

    [JsonPropertyName("source")]
    public string Source { get; set; } = "";

    [JsonPropertyName("date")]
    public string Date { get; set; } = "";

    [JsonPropertyName("startTime")]
    public string StartTime { get; set; } = "";

    [JsonPropertyName("endTime")]
    public string EndTime { get; set; } = "";

    [JsonPropertyName("placeName")]
    public string PlaceName { get; set; } = "";

    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("purpose")]
    public string Purpose { get; set; } = "";

    [JsonPropertyName("status")]
    public string Status { get; set; } = "";

    [JsonPropertyName("cancelled")]
    public bool Cancelled { get; set; }

    [JsonPropertyName("submitted")]
    public bool Submitted { get; set; }
}

internal sealed class VisitorStats
{
    [JsonPropertyName("daily")]
    public int Daily { get; set; }

    [JsonPropertyName("weekly")]
    public int Weekly { get; set; }

    [JsonPropertyName("monthly")]
    public int Monthly { get; set; }
}

internal sealed class DashboardSummary
{
    [JsonPropertyName("visitorStats")]
    public VisitorStats VisitorStats { get; set; } = new();

    [JsonPropertyName("activeMemberCount")]
    public int ActiveMemberCount { get; set; }

    [JsonPropertyName("weekStart")]
    public string WeekStart { get; set; } = "";

    [JsonPropertyName("monthStart")]
    public string MonthStart { get; set; } = "";
}

internal sealed class DashboardSnapshot
{
    [JsonPropertyName("ok")]
    public bool Ok { get; set; }

    [JsonPropertyName("error")]
    public string Error { get; set; } = "";

    [JsonPropertyName("serverTime")]
    public string ServerTime { get; set; } = "";

    [JsonPropertyName("events")]
    public List<StaffEvent> Events { get; set; } = [];

    [JsonPropertyName("todayReservations")]
    public List<Reservation> TodayReservations { get; set; } = [];

    [JsonPropertyName("calendarEntries")]
    public List<CalendarEntry> CalendarEntries { get; set; } = [];

    [JsonPropertyName("summary")]
    public DashboardSummary Summary { get; set; } = new();
}

internal sealed class AdminChallenge
{
    [JsonPropertyName("ok")]
    public bool Ok { get; set; }

    [JsonPropertyName("error")]
    public string Error { get; set; } = "";

    [JsonPropertyName("challenge")]
    public string Challenge { get; set; } = "";
}

internal sealed class AdminLogin
{
    [JsonPropertyName("ok")]
    public bool Ok { get; set; }

    [JsonPropertyName("error")]
    public string Error { get; set; } = "";

    [JsonPropertyName("sessionToken")]
    public string SessionToken { get; set; } = "";
}
