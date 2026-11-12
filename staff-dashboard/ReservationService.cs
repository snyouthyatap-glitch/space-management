using System.Text.RegularExpressions;

namespace YouthSpring.StaffDashboard;

internal static class ReservationService
{
    public static List<CalendarEntry> MergeCalendarEntries(
        IEnumerable<CalendarEntry> calendarEntries,
        IEnumerable<CalendarEntry> todayEntries,
        IEnumerable<CalendarEntry> requestEntries)
    {
        return calendarEntries
            .Concat(todayEntries)
            .Concat(requestEntries)
            .Where(item => !string.IsNullOrWhiteSpace(item.Id))
            .GroupBy(ReservationMergeKey, StringComparer.Ordinal)
            .Select(MergeDuplicateReservationEntries)
            .OrderBy(item => item.Date, StringComparer.Ordinal)
            .ThenBy(item => item.StartTime, StringComparer.Ordinal)
            .ToList();
    }

    public const string MissingUsageLogStatus = "일지 미제출";

    public static void ApplyCheckIns(IEnumerable<CalendarEntry> entries, IReadOnlyCollection<StaffCheckIn> checkIns)
    {
        var byDateAndName = checkIns
            .Where(item => !string.IsNullOrWhiteSpace(item.Date) && !string.IsNullOrWhiteSpace(item.Name))
            .GroupBy(item => $"{item.Date}|{NormalizeNameKey(item.Name)}", StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
        var memberDates = checkIns
            .Where(item => !string.IsNullOrWhiteSpace(item.MemberId))
            .Select(item => $"{item.Date}|{item.MemberId}")
            .ToHashSet(StringComparer.Ordinal);

        foreach (var entry in entries)
        {
            if (entry.Source == "usage") continue;
            if (!string.IsNullOrWhiteSpace(entry.MemberId))
            {
                entry.CheckedIn = memberDates.Contains($"{entry.Date}|{entry.MemberId}");
                continue;
            }

            if (!byDateAndName.TryGetValue($"{entry.Date}|{NormalizeNameKey(entry.Name)}", out var matches))
            {
                entry.CheckedIn = false;
                continue;
            }

            entry.CheckedIn = true;
            var memberIds = matches
                .Select(item => item.MemberId)
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Distinct(StringComparer.Ordinal)
                .ToList();
            if (memberIds.Count == 1) entry.MemberId = memberIds[0];
        }
    }

    private static string NormalizeNameKey(string? value) =>
        Regex.Replace(value ?? "", @"\s+", "").Trim().ToLowerInvariant();

    public static string GetUnifiedStatus(CalendarEntry entry)
    {
        if (entry.Cancelled || entry.Status is "취소" or "반려") return "취소";
        if (entry.Status == "노쇼") return "노쇼";

        var now = DateTime.Now;
        var start = ParseReservationDateTime(entry.Date, entry.StartTime);
        var end = ParseReservationDateTime(entry.Date, entry.EndTime);
        var hasUsageLog = entry.Submitted || entry.Source == "usage" || entry.Status == "일지제출완료";

        if (hasUsageLog)
        {
            if (start.HasValue && end.HasValue && now >= start.Value && now < end.Value) return "이용중";
            if (end.HasValue && now >= end.Value) return "이용완료";
            if (entry.Date.CompareTo(DateTime.Today.ToString("yyyy-MM-dd")) < 0) return "이용완료";
            return "승인 완료";
        }

        if (start.HasValue && now >= start.Value.AddMinutes(10))
            return entry.CheckedIn ? MissingUsageLogStatus : "노쇼";
        if (entry.Status == "신청") return "승인 대기";
        return "승인 완료";
    }

    public const int OverdueUsageLogAlertMinutes = 30;

    // 오늘 예약 중 시작 시각에서 30분이 지났는데 일지가 없는 예약.
    // 직원이 노쇼 처리하거나 일지를 대리 제출해야 하는 대상이다.
    public static List<CalendarEntry> FindOverdueUsageLogs(IEnumerable<CalendarEntry> entries, DateTime now)
    {
        return entries
            .Where(entry =>
                entry.Source != "usage" &&
                !entry.Submitted &&
                !entry.Cancelled &&
                entry.Status is not ("취소" or "반려" or "노쇼" or "이용완료" or "일지제출완료") &&
                IsSameCalendarDate(entry.Date, now))
            .Where(entry =>
                ParseReservationDateTime(entry.Date, entry.StartTime) is { } start &&
                now >= start.AddMinutes(OverdueUsageLogAlertMinutes))
            .OrderBy(entry => entry.StartTime, StringComparer.Ordinal)
            .ToList();
    }

    public static bool IsSameCalendarDate(string value, DateTime date)
    {
        if (DateTime.TryParse(value, out var parsed)) return parsed.Date == date.Date;
        return string.Equals(value, date.ToString("yyyy-MM-dd"), StringComparison.Ordinal);
    }

    public static bool IsPlaceMatch(CalendarEntry item, string filter)
    {
        var value = $"{item.PlaceName} {item.SpaceType}";
        return filter switch
        {
            "커리어존 Room1" => value.Contains("Room1", StringComparison.OrdinalIgnoreCase) ||
                                value.Contains("careerRoom1", StringComparison.OrdinalIgnoreCase),
            "커리어존 Room2" => value.Contains("Room2", StringComparison.OrdinalIgnoreCase) ||
                                value.Contains("careerRoom2", StringComparison.OrdinalIgnoreCase),
            "커넥트룸" => value.Contains("커넥트", StringComparison.OrdinalIgnoreCase) ||
                         value.Contains("connectRoom", StringComparison.OrdinalIgnoreCase),
            _ => true
        };
    }

    public static string NormalizeGridTime(string value)
    {
        var time = Regex.Match(value ?? "", @"(\d{1,2}):(\d{2})");
        return time.Success ? $"{time.Groups[1].Value.PadLeft(2, '0')}:{time.Groups[2].Value}" : (value ?? "").Trim();
    }

    private static string ReservationMergeKey(CalendarEntry item)
    {
        var name = Regex.Replace(item.Name ?? "", @"\s+", "").Trim().ToLowerInvariant();
        var date = item.Date ?? "";
        var start = NormalizeGridTime(item.StartTime);
        var end = NormalizeGridTime(item.EndTime);
        if (!string.IsNullOrWhiteSpace(name) &&
            !string.IsNullOrWhiteSpace(date) &&
            !string.IsNullOrWhiteSpace(start) &&
            !string.IsNullOrWhiteSpace(end))
        {
            return $"reservation:{name}:{date}:{start}:{end}";
        }
        return $"id:{item.Source}:{item.Id}";
    }

    private static CalendarEntry MergeDuplicateReservationEntries(IEnumerable<CalendarEntry> group)
    {
        var entries = group.ToList();
        var primary = entries
            .OrderByDescending(ReservationProgressRank)
            .ThenByDescending(item => item.Source == "reservationRequest")
            .First();

        return new CalendarEntry
        {
            Id = primary.Id,
            Source = primary.Source,
            MemberId = FirstNonEmpty(entries.Select(item => item.MemberId).Prepend(primary.MemberId).ToArray()),
            Date = primary.Date,
            StartTime = primary.StartTime,
            EndTime = primary.EndTime,
            PlaceName = FirstNonEmpty(entries.Select(item => item.PlaceName).Prepend(primary.PlaceName).ToArray()),
            SpaceType = FirstNonEmpty(entries.Select(item => item.SpaceType).Prepend(primary.SpaceType).ToArray()),
            Name = primary.Name,
            Purpose = FirstNonEmpty(entries.Select(item => item.Purpose).Prepend(primary.Purpose).ToArray()),
            Headcount = Math.Max(1, entries.Max(item => item.Headcount)),
            Status = primary.Status,
            Cancelled = entries.Any(item => item.Cancelled || item.Status is "취소" or "반려"),
            Submitted = entries.Any(item => item.Submitted || item.Source == "usage" || item.Status == "일지제출완료"),
            CheckedIn = entries.Any(item => item.CheckedIn)
        };
    }

    private static int ReservationProgressRank(CalendarEntry item)
    {
        if (item.Cancelled || item.Status is "취소" or "반려") return 10;
        if (item.Submitted || item.Source == "usage" || item.Status == "일지제출완료") return 50;
        if (item.Status is "승인" or "확정" or "완료") return 40;
        if (item.Status == "신청") return 20;
        return 30;
    }

    private static DateTime? ParseReservationDateTime(string date, string time)
    {
        if (DateTime.TryParse($"{date} {time}", out var value)) return value;
        return null;
    }

    private static string FirstNonEmpty(params string[] values) =>
        values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
}
