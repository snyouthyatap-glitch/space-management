using System.Text.Json;
using System.Text.RegularExpressions;

namespace YouthSpring.StaffDashboard;

internal sealed record NaverSheetSyncItem(NaverBookingReservation Reservation, CalendarEntry Entry, string Status)
{
    // 날짜·시간·공간·인원·예약자까지 넣어서, 네이버에서 예약이 바뀌면 다시 보내게 한다.
    public string Key =>
        $"{Reservation.ReservationNo}|{Status}|{Entry.Date}|{Entry.StartTime}|{Entry.EndTime}|" +
        $"{Entry.SpaceType}|{Entry.Headcount}|{Entry.Name}|{Reservation.Phone}";
}

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
                .Where(item => IsPendingApplication(item) && !IsOtherProduct(item))
                .ToList(),
            TodayUse = ReadResultRows(results, "today_use"),
            TodayCancel = ReadResultRows(results, "today_cancel")
        };
    }

    // 네이버 '오늘이용'과 '오늘취소'를 합쳐 직원 프로그램 예약 목록으로 바꾼다.
    // 같은 사람이 취소 후 같은 시간에 다시 예약했으면, 새 예약이 취소로 덮이지 않게 취소 건을 뺀다.
    public static List<CalendarEntry> BuildSmartPlaceEntries(NaverBookingSnapshot snapshot)
    {
        var active = snapshot.TodayUse
            .Where(item => !IsOtherProduct(item))
            .Select(ConvertTodayUseReservation)
            .Where(item => !string.IsNullOrWhiteSpace(item.Id))
            .ToList();
        var activeSlots = active
            .Where(item => !item.Cancelled && item.Status is not ("취소" or "노쇼"))
            .Select(item => (Id: item.Id, Slot: SlotKey(item)))
            .ToList();
        var cancelled = snapshot.TodayCancel
            .Where(item => !IsOtherProduct(item))
            .Select(ConvertCancelledReservation)
            .Where(item => !string.IsNullOrWhiteSpace(item.Id))
            .Where(item => !activeSlots.Any(slot => slot.Slot == SlotKey(item) && slot.Id != item.Id));
        return active.Concat(cancelled).ToList();
    }

    // 서버 '네이버 예약' 시트에 맞출 대상: 오늘이용 중 확정·승인·이용완료 건과 오늘취소 건.
    // 날짜·시간·공간을 알 수 없는 행은 서버가 거절하므로 미리 뺀다.
    public static List<NaverSheetSyncItem> BuildSheetSyncItems(NaverBookingSnapshot snapshot)
    {
        var items = new List<NaverSheetSyncItem>();
        foreach (var reservation in snapshot.TodayUse.Where(item => !IsOtherProduct(item)))
        {
            var status = NormalizeStatus(reservation.Status);
            if (status is not ("확정" or "승인" or "완료")) continue;
            var entry = ConvertTodayUseReservation(reservation);
            if (IsSyncable(reservation, entry))
                items.Add(new NaverSheetSyncItem(reservation, entry, status == "완료" ? "이용완료" : status));
        }
        foreach (var reservation in snapshot.TodayCancel.Where(item => !IsOtherProduct(item)))
        {
            var entry = ConvertCancelledReservation(reservation);
            if (IsSyncable(reservation, entry))
                items.Add(new NaverSheetSyncItem(reservation, entry, "취소"));
        }
        return items;
    }

    // 네이버 예약자관리 목록에는 다른 상품([동아리실] 등)도 섞여 나온다.
    // 상품명을 알고, 그게 커리어존·커넥트룸이 아니면 뺀다. 상품명을 모르면 남긴다.
    public static bool IsOtherProduct(NaverBookingReservation item)
    {
        var productName = FirstNonEmpty(ParseReservationText(item).ProductName, item.ProductName);
        return !string.IsNullOrWhiteSpace(productName) && SpaceTypeFromProductName(productName) == "";
    }

    private static bool IsSyncable(NaverBookingReservation reservation, CalendarEntry entry) =>
        !string.IsNullOrWhiteSpace(reservation.ReservationNo) &&
        !string.IsNullOrWhiteSpace(entry.StartTime) &&
        !string.IsNullOrWhiteSpace(entry.SpaceType);

    public static CalendarEntry ConvertCancelledReservation(NaverBookingReservation item)
    {
        var entry = ConvertTodayUseReservation(item);
        entry.Status = "취소";
        entry.Cancelled = true;
        entry.Purpose = FirstNonEmpty(item.Purpose, "네이버 예약 취소");
        return entry;
    }

    private static string SlotKey(CalendarEntry item) =>
        $"{Regex.Replace(item.Name ?? "", @"\s+", "")}|{item.Date}|{item.StartTime}|{item.EndTime}";

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
                @"(확정|신청|취소|완료|노쇼)([^0-9]+?)(01\d-\d{3,4}-\d{4})(?:01\d-\d{3,4}-\d{4})?" + Regex.Escape(item.ReservationNo))
            .Cast<Match>()
            .LastOrDefault();
        // 예약번호 바로 뒤의 이용일시. 두 가지 화면 형태를 모두 읽는다.
        //   26. 8. 14.(금) 오후 6:00~9:00   /   26.11.11.수 17:30~19:30 (24시간제, 날짜 생략 가능)
        var useDateTime = Regex.Match(
            rawText,
            Regex.Escape(item.ReservationNo) +
            @"(?:\s*변경)?\s*((?:\d{2}\.\s*\d{1,2}\.\s*\d{1,2}\.\s*(?:\([^)]+\)|[월화수목금토일])\s*)?(?:(?:오전|오후)\s*)?\d{1,2}:\d{2}\s*~\s*(?:(?:오전|오후)\s*)?\d{1,2}:\d{2})");

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
                BookerName = FirstNonEmpty(
                    NormalizeBookerName(parsed.BookerName),
                    NormalizeBookerName(GetJsonString(item, "bookerName"))),
                Status = FirstNonEmpty(parsed.Status, GetJsonString(item, "status")),
                UseDateTime = FirstNonEmpty(parsed.UseDateTime, GetJsonString(item, "useDateTime")),
                Phone = GetJsonString(item, "phone"),
                // 네이버는 바뀐 예약의 예약번호 아래에 '변경'을 붙인다.
                Changed = GetJsonBool(item, "changed") ||
                    (!string.IsNullOrWhiteSpace(reservationNo) &&
                     Regex.IsMatch(rawText, Regex.Escape(reservationNo) + @"\s*변경")),
                Purpose = GetJsonString(item, "purpose"),
                Headcount = GetJsonInt(item, "headcount", 1),
                RawText = rawText
            });
        }
        return rows;
    }

    private static string CleanBookerName(string value)
    {
        var cleaned = NormalizeBookerName(value);
        return string.IsNullOrWhiteSpace(cleaned) ? "예약자" : cleaned;
    }

    // automation/naver-booker.js의 cleanNaverBookerName과 같은 규칙이다. 바꿀 때는 같이 바꾼다.
    // 예) "김예약 [대리예약] 방문자: 이방문" → "이방문", "최관리 [관리자등록]" → "최관리"
    public static string NormalizeBookerName(string? value)
    {
        var text = Regex.Replace(value ?? "", @"\s+", " ").Trim();
        text = Regex.Replace(text, @"01[0-9]-?\d{3,4}-?\d{4}.*$", "");
        text = Regex.Replace(text, @"\d{9,}.*$", "");
        var parts = Regex.Split(text, @"(?:예약자|방문자)\s*[:：]?");
        for (var index = parts.Length - 1; index >= 0; index--)
        {
            var name = Regex.Replace(parts[index], @"\[[^\]]*\]", " ");
            name = Regex.Replace(name, @"대리\s*예약|관리자\s*등록", " ");
            name = Regex.Replace(name, @"[^가-힣A-Za-z\s]", "");
            name = Regex.Replace(name, @"\s+", " ").Trim();
            if (name.Length > 0) return name;
        }
        return "";
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
        var match = Regex.Match(value ?? "", @"(?:(\d{2,4})\.\s*)?(\d{1,2})\.\s*(\d{1,2})\.\s*(?:\([^)]+\)|[월화수목금토일])");
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
        // 오전/오후가 없으면 24시간제(17:30~19:30)로 본다.
        var match = Regex.Match(
            value ?? "",
            @"(?:(오전|오후)\s*)?(\d{1,2}):(\d{2})\s*~\s*(?:(오전|오후)\s*)?(\d{1,2}):(\d{2})");
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

    private static bool GetJsonBool(JsonElement element, string propertyName) =>
        element.TryGetProperty(propertyName, out var property) && property.ValueKind == JsonValueKind.True;

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
