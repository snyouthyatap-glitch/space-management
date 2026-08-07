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

internal sealed class ReservationStatusUpdate
{
    [JsonPropertyName("ok")]
    public bool Ok { get; set; }

    [JsonPropertyName("error")]
    public string Error { get; set; } = "";

    [JsonPropertyName("reservationId")]
    public string ReservationId { get; set; } = "";

    [JsonPropertyName("status")]
    public string Status { get; set; } = "";

    [JsonPropertyName("cancelled")]
    public bool Cancelled { get; set; }
}

internal sealed class AdminMemberLookup
{
    [JsonPropertyName("ok")]
    public bool Ok { get; set; }

    [JsonPropertyName("error")]
    public string Error { get; set; } = "";

    [JsonPropertyName("member")]
    public AdminMember Member { get; set; } = new();
}

internal sealed class AdminMember
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = "";

    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("gender")]
    public string Gender { get; set; } = "";

    [JsonPropertyName("age")]
    public int Age { get; set; }

    [JsonPropertyName("birthDate")]
    public string BirthDate { get; set; } = "";

    [JsonPropertyName("phone")]
    public string Phone { get; set; } = "";
}

internal sealed class ManualLogRequest
{
    [JsonPropertyName("subjectType")]
    public string SubjectType { get; set; } = "member";

    [JsonPropertyName("memberId")]
    public string MemberId { get; set; } = "";

    [JsonPropertyName("guest")]
    public ManualGuest? Guest { get; set; }

    [JsonPropertyName("logType")]
    public string LogType { get; set; } = "visit";

    [JsonPropertyName("date")]
    public string Date { get; set; } = "";

    [JsonPropertyName("data")]
    public ManualLogData Data { get; set; } = new();

    [JsonPropertyName("requestId")]
    public string RequestId { get; set; } = "";
}

internal sealed class ManualGuest
{
    [JsonPropertyName("token")]
    public string Token { get; set; } = "";

    [JsonPropertyName("gender")]
    public string Gender { get; set; } = "";

    [JsonPropertyName("age")]
    public int Age { get; set; }
}

internal sealed class ManualLogData
{
    [JsonPropertyName("count")]
    public int Count { get; set; }

    [JsonPropertyName("startTime")]
    public string StartTime { get; set; } = "";

    [JsonPropertyName("endTime")]
    public string EndTime { get; set; } = "";

    [JsonPropertyName("purpose")]
    public string Purpose { get; set; } = "";

    [JsonPropertyName("companions")]
    public List<ManualCompanion> Companions { get; set; } = [];
}

internal sealed class ManualCompanion
{
    [JsonPropertyName("gender")]
    public string Gender { get; set; } = "";

    [JsonPropertyName("age")]
    public int Age { get; set; }
}

internal sealed class ManualLogResult
{
    [JsonPropertyName("ok")]
    public bool Ok { get; set; }

    [JsonPropertyName("error")]
    public string Error { get; set; } = "";

    [JsonPropertyName("date")]
    public string Date { get; set; } = "";

    [JsonPropertyName("logCreated")]
    public bool LogCreated { get; set; }

    [JsonPropertyName("linkedReservation")]
    public bool LinkedReservation { get; set; }
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

internal sealed class MonthlyPerformanceReport
{
    [JsonPropertyName("ok")]
    public bool Ok { get; set; }

    [JsonPropertyName("error")]
    public string Error { get; set; } = "";

    [JsonPropertyName("month")]
    public string Month { get; set; } = "";

    [JsonPropertyName("generatedAt")]
    public string GeneratedAt { get; set; } = "";

    [JsonPropertyName("summary")]
    public List<PerformanceSummaryRow> Summary { get; set; } = [];

    [JsonPropertyName("totalUsers")]
    public int TotalUsers { get; set; }

    [JsonPropertyName("totalUniqueUsers")]
    public int TotalUniqueUsers { get; set; }

    [JsonPropertyName("daily")]
    public PerformanceDailyGroups Daily { get; set; } = new();

    [JsonPropertyName("simplified")]
    public PerformanceSimplifiedGroups Simplified { get; set; } = new();

    [JsonPropertyName("uniqueBasis")]
    public string UniqueBasis { get; set; } = "";
}

internal sealed class PerformanceSummaryRow
{
    [JsonPropertyName("key")]
    public string Key { get; set; } = "";

    [JsonPropertyName("label")]
    public string Label { get; set; } = "";

    [JsonPropertyName("uniqueUsers")]
    public int UniqueUsers { get; set; }

    [JsonPropertyName("totalUsers")]
    public int TotalUsers { get; set; }
}

internal sealed class PerformanceDailyGroups
{
    [JsonPropertyName("visit")]
    public List<PerformanceDailyRow> Visit { get; set; } = [];

    [JsonPropertyName("connect")]
    public List<PerformanceDailyRow> Connect { get; set; } = [];

    [JsonPropertyName("career")]
    public List<PerformanceDailyRow> Career { get; set; } = [];

    [JsonPropertyName("printer")]
    public List<PerformanceDailyRow> Printer { get; set; } = [];
}

internal sealed class PerformanceDailyRow
{
    [JsonPropertyName("date")]
    public string Date { get; set; } = "";

    [JsonPropertyName("maleUnder19")]
    public int MaleUnder19 { get; set; }

    [JsonPropertyName("femaleUnder19")]
    public int FemaleUnder19 { get; set; }

    [JsonPropertyName("male20To24")]
    public int Male20To24 { get; set; }

    [JsonPropertyName("female20To24")]
    public int Female20To24 { get; set; }

    [JsonPropertyName("male25To29")]
    public int Male25To29 { get; set; }

    [JsonPropertyName("female25To29")]
    public int Female25To29 { get; set; }

    [JsonPropertyName("male30To34")]
    public int Male30To34 { get; set; }

    [JsonPropertyName("female30To34")]
    public int Female30To34 { get; set; }

    [JsonPropertyName("male35To39")]
    public int Male35To39 { get; set; }

    [JsonPropertyName("female35To39")]
    public int Female35To39 { get; set; }

    [JsonPropertyName("male40Plus")]
    public int Male40Plus { get; set; }

    [JsonPropertyName("female40Plus")]
    public int Female40Plus { get; set; }

    [JsonPropertyName("unknown")]
    public int Unknown { get; set; }

    [JsonPropertyName("total")]
    public int Total { get; set; }
}

internal sealed class PerformanceSimplifiedGroups
{
    [JsonPropertyName("visit")]
    public List<PerformanceVisitSimplifiedRow> Visit { get; set; } = [];

    [JsonPropertyName("printer")]
    public List<PerformancePrinterSimplifiedRow> Printer { get; set; } = [];

    [JsonPropertyName("career")]
    public List<PerformanceRoomSimplifiedRow> Career { get; set; } = [];

    [JsonPropertyName("connect")]
    public List<PerformanceRoomSimplifiedRow> Connect { get; set; } = [];
}

internal class PerformanceVisitSimplifiedRow
{
    [JsonPropertyName("date")]
    public string Date { get; set; } = "";

    [JsonPropertyName("gender")]
    public string Gender { get; set; } = "";

    [JsonPropertyName("age")]
    public int Age { get; set; }
}

internal sealed class PerformancePrinterSimplifiedRow : PerformanceVisitSimplifiedRow
{
    [JsonPropertyName("count")]
    public int Count { get; set; }
}

internal sealed class PerformanceRoomSimplifiedRow
{
    [JsonPropertyName("date")]
    public string Date { get; set; } = "";

    [JsonPropertyName("space")]
    public string Space { get; set; } = "";

    [JsonPropertyName("purpose")]
    public string Purpose { get; set; } = "";

    [JsonPropertyName("headcount")]
    public int Headcount { get; set; }

    [JsonPropertyName("note")]
    public string Note { get; set; } = "";

    [JsonPropertyName("startTime")]
    public string StartTime { get; set; } = "";

    [JsonPropertyName("endTime")]
    public string EndTime { get; set; } = "";

    [JsonPropertyName("male20s")]
    public int Male20s { get; set; }

    [JsonPropertyName("female20s")]
    public int Female20s { get; set; }

    [JsonPropertyName("male30s")]
    public int Male30s { get; set; }

    [JsonPropertyName("female30s")]
    public int Female30s { get; set; }

    [JsonPropertyName("maleUnder19")]
    public int MaleUnder19 { get; set; }

    [JsonPropertyName("femaleUnder19")]
    public int FemaleUnder19 { get; set; }

    [JsonPropertyName("male40Plus")]
    public int Male40Plus { get; set; }

    [JsonPropertyName("female40Plus")]
    public int Female40Plus { get; set; }
}
