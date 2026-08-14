using System.Text.Json;
using System.Text.RegularExpressions;

namespace YouthSpring.StaffDashboard;

internal static class NaverBookingParser
{
    public static NaverBookingSnapshot ReadSnapshot(string outputPath)
    {
        if (!File.Exists(outputPath)) return new NaverBookingSnapshot();
        using var stream = File.OpenRead(outputPath);
        using var document = JsonDocument.Parse(stream);
        if (!document.RootElement.TryGetProperty("results", out var results))
        {
            return new NaverBookingSnapshot();
        }

        return new NaverBookingSnapshot
        {
            PendingConfirm = ReadResultRows(results, "pending_confirm")
                .Where(IsPendingApplication)
                .ToList(),
            TodayUse = ReadResultRows(results, "today_use")
        };
    }

    public static CalendarEntry ConvertTodayUseReservation(NaverBookingReservation item)
    {
        var parsed = ParseReservationText(item);
        var useDateTime = FirstNonEmpty(parsed.UseDateTime, item.UseDateTime);
        var timeRange = ParseTimeRange(useDateTime);
        return new CalendarEntry
        {
            Id = "naverSmartPlace:" + item.ReservationNo,
            Source = "naverSmartPlace",
            Date = ParseUsageDate(useDateTime) ?? DateTime.Today.ToString("yyyy-MM-dd"),
            StartTime = timeRange.StartTime,
            EndTime = timeRange.EndTime,
            PlaceName = ShortProductName(FirstNonEmpty(parsed.ProductName, item.ProductName, "네이버 예약")),
            SpaceType = SpaceTypeFromProductName(FirstNonEmpty(parsed.ProductName, item.ProductName)),
            Name = CleanBookerName(FirstNonEmpty(parsed.BookerName, item.BookerName, "예약자")),
            Purpose = FirstNonEmpty(item.Purpose, "네이버 오늘 이용"),
            Headcount = Math.Max(1, item.Headcount),
            Status = FirstNonEmpty(parsed.Status, item.Status, "확정")
        };
    }

    public static string BuildPendingReservationNotice(List<NaverBookingReservation> pending)
    {
        var counts = pending
            .GroupBy(item => string.IsNullOrWhiteSpace(item.ProductName) ? "예약" : item.ProductName)
            .Select(group => $"{ShortProductName(group.Key)} {group.Count()}건");
        return string.Join(" · ", counts) + "\n프로그램에서 내용을 확인해 주세요.";
    }

    public static bool IsPendingApplication(NaverBookingReservation item) =>
        string.Equals(NormalizeStatus(item.Status), "신청", StringComparison.Ordinal);

    public static NaverBookingReservation ParseReservationText(NaverBookingReservation item)
    {
        var rawText = item.RawText ?? "";
        if (string.IsNullOrWhiteSpace(rawText) || string.IsNullOrWhiteSpace(item.ReservationNo))
            return new NaverBookingReservation();

        var statusName = Regex.Matches(
                rawText,
                @"(확정|신청|취소|완료|노쇼)([^0-9]+?)(01\d-\d{3,4}-\d{4})" + Regex.Escape(item.ReservationNo))
            .Cast<Match>()
            .LastOrDefault();
        var useDateTime = Regex.Match(
            rawText,
            Regex.Escape(item.ReservationNo) + @"(?:변경)?((?:\d{2}\.\s*\d{1,2}\.\s*\d{1,2}\.\([^)]+\)\s*)?(?:오전|오후)\s*\d{1,2}:\d{2}~(?:(?:오전|오후)\s*)?\d{1,2}:\d{2})");

        var productName = "";
        if (useDateTime.Success)
        {
            var tail = rawText[useDateTime.Index..];
            productName = new[] { "AI 커리어존 (청취Room1)", "AI 커리어존 (청취Room2)", "청년 커넥트룸" }
                .Select(name => new { Name = name, Index = tail.IndexOf(name, StringComparison.Ordinal) })
                .Where(item => item.Index >= 0)
                .OrderBy(item => item.Index)
                .Select(item => item.Name)
                .FirstOrDefault() ?? "";
        }

        return new NaverBookingReservation
        {
            Status = statusName?.Groups[1].Value ?? "",
            BookerName = statusName is null ? "" : statusName.Groups[2].Value.Trim(),
            UseDateTime = useDateTime.Success ? useDateTime.Groups[1].Value.Trim() : "",
            ProductName = productName
        };
    }

    public static string ShortProductName(string productName)
    {
        if (productName.Contains("Room1", StringComparison.OrdinalIgnoreCase)) return "커리어존 Room1";
        if (productName.Contains("Room2", StringComparison.OrdinalIgnoreCase)) return "커리어존 Room2";
        if (productName.Contains("커넥트", StringComparison.Ordinal)) return "커넥트룸";
        return productName;
    }

    private static string SpaceTypeFromProductName(string productName)
    {
        if (productName.Contains("Room1", StringComparison.OrdinalIgnoreCase)) return "careerRoom1";
        if (productName.Contains("Room2", StringComparison.OrdinalIgnoreCase)) return "careerRoom2";
        if (productName.Contains("커넥트", StringComparison.Ordinal)) return "connectRoom";
        return "";
    }

    private static List<NaverBookingReservation> ReadResultRows(JsonElement results, string propertyName)
    {
        if (!results.TryGetProperty(propertyName, out var rowsElement) ||
            rowsElement.ValueKind != JsonValueKind.Array) return [];

        var rows = new List<NaverBookingReservation>();
        foreach (var item in rowsElement.EnumerateArray())
        {
            var reservationNo = GetJsonString(item, "reservationNo");
            var rawText = string.Join(" ", new[]
            {
                GetJsonString(item, "rawText"),
                GetJsonString(item, "detailText")
            }.Where(value => !string.IsNullOrWhiteSpace(value)));
            var parsed = ParseReservationText(new NaverBookingReservation
            {
                ReservationNo = reservationNo,
                RawText = rawText
            });
            rows.Add(new NaverBookingReservation
            {
                ReservationNo = reservationNo,
                ProductName = FirstNonEmpty(parsed.ProductName, GetJsonString(item, "productName")),
                BookerName = FirstNonEmpty(parsed.BookerName, GetJsonString(item, "bookerName")),
                Status = FirstNonEmpty(parsed.Status, GetJsonString(item, "status")),
                UseDateTime = FirstNonEmpty(parsed.UseDateTime, GetJsonString(item, "useDateTime")),
                Purpose = GetJsonString(item, "purpose"),
                Headcount = GetJsonInt(item, "headcount", 1),
                RawText = rawText
            });
        }
        return rows;
    }

    private static string CleanBookerName(string value)
    {
        var cleaned = Regex.Replace(value ?? "", @"\s+", " ").Trim();
        cleaned = Regex.Replace(cleaned, @"01[0-9]-?\d{3,4}-?\d{4}.*$", "").Trim();
        cleaned = Regex.Replace(cleaned, @"\d{9,}.*$", "").Trim();
        cleaned = Regex.Replace(cleaned, @"^(예약자|방문자)\s*[:：]?\s*", "").Trim();
        return string.IsNullOrWhiteSpace(cleaned) ? "예약자" : cleaned;
    }

    private static string NormalizeStatus(string value)
    {
        var status = Regex.Replace(value ?? "", @"\s+", "").Trim();
        if (status.Contains("신청", StringComparison.Ordinal)) return "신청";
        if (status.Contains("확정", StringComparison.Ordinal)) return "확정";
        if (status.Contains("취소", StringComparison.Ordinal)) return "취소";
        if (status.Contains("완료", StringComparison.Ordinal)) return "완료";
        if (status.Contains("노쇼", StringComparison.Ordinal)) return "노쇼";
        return status;
    }

    private static string? ParseUsageDate(string value)
    {
        var match = Regex.Match(value ?? "", @"(?:(\d{2,4})\.\s*)?(\d{1,2})\.\s*(\d{1,2})\.\([^)]+\)");
        if (!match.Success) return null;
        var yearText = match.Groups[1].Value;
        var year = string.IsNullOrWhiteSpace(yearText) ? DateTime.Today.Year : int.Parse(yearText);
        if (year < 100) year += 2000;
        var month = int.Parse(match.Groups[2].Value);
        var day = int.Parse(match.Groups[3].Value);
        return $"{year:0000}-{month:00}-{day:00}";
    }

    private static (string StartTime, string EndTime) ParseTimeRange(string value)
    {
        var match = Regex.Match(
            value ?? "",
            @"(?:(?:\d{2,4}\.\s*)?\d{1,2}\.\s*\d{1,2}\.\([^)]+\)\s*)?(오전|오후)\s*(\d{1,2}):(\d{2})~(?:(오전|오후)\s*)?(\d{1,2}):(\d{2})");
        if (!match.Success) return ("", "");

        var startMeridiem = match.Groups[1].Value;
        var endMeridiem = string.IsNullOrWhiteSpace(match.Groups[4].Value)
            ? startMeridiem
            : match.Groups[4].Value;
        return (
            ToTwentyFourHour(startMeridiem, match.Groups[2].Value, match.Groups[3].Value),
            ToTwentyFourHour(endMeridiem, match.Groups[5].Value, match.Groups[6].Value));
    }

    private static string ToTwentyFourHour(string meridiem, string hourText, string minuteText)
    {
        var hour = int.TryParse(hourText, out var parsedHour) ? parsedHour : 0;
        if (meridiem == "오후" && hour < 12) hour += 12;
        if (meridiem == "오전" && hour == 12) hour = 0;
        return $"{hour:00}:{minuteText}";
    }

    private static string GetJsonString(JsonElement element, string propertyName) =>
        element.TryGetProperty(propertyName, out var property) && property.ValueKind == JsonValueKind.String
            ? property.GetString() ?? ""
            : "";

    private static int GetJsonInt(JsonElement element, string propertyName, int fallback)
    {
        if (!element.TryGetProperty(propertyName, out var property)) return fallback;
        if (property.ValueKind == JsonValueKind.Number && property.TryGetInt32(out var number)) return number;
        if (property.ValueKind == JsonValueKind.String)
        {
            var match = Regex.Match(property.GetString() ?? "", @"\d+");
            if (match.Success && int.TryParse(match.Value, out number)) return number;
        }
        return fallback;
    }

    private static string FirstNonEmpty(params string[] values) =>
        values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
}
