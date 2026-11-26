using System.Text.RegularExpressions;

namespace YouthSpring.StaffDashboard;

internal static class ReservationService
{
    public static List<CalendarEntry> MergeCalendarEntries(
        IEnumerable<CalendarEntry> calendarEntries,
        IEnumerable<CalendarEntry> todayEntries,
        IEnumerable<CalendarEntry> requestEntries)
    {
        var combined = CombineNaverSources(calendarEntries
            .Concat(todayEntries)
            .Concat(requestEntries)
            .Where(item => !string.IsNullOrWhiteSpace(item.Id)));
        return combined
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

            // 네이버가 가린 이름(박*예)은 같은 날 방문한 회원 중 모양이 맞는 사람을 찾는다.
            List<StaffCheckIn>? matches;
            if ((entry.Name ?? "").Contains('*'))
            {
                matches = checkIns
                    .Where(item => item.Date == entry.Date && NaverBookingParser.MaskedNameMatches(entry.Name, item.Name))
                    .ToList();
                if (matches.Count == 0) matches = null;
            }
            else if (!byDateAndName.TryGetValue($"{entry.Date}|{NormalizeNameKey(entry.Name)}", out matches))
            {
                matches = null;
            }
            if (matches is null)
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
        // 노쇼를 먼저 본다. 예전 서버는 노쇼에도 cancelled를 붙여 보냈다.
        if (entry.Status == "노쇼") return "노쇼";
        if (entry.Cancelled || entry.Status is "취소" or "반려") return "취소";

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

        // 승인하지 않은 신청은 시작 시간이 되면 취소된다(서버도 같은 규칙으로 시트를 '취소'로 바꾼다).
        if (entry.Status == "신청")
            return start.HasValue && now >= start.Value ? "취소" : "승인 대기";
        // 노쇼 판정은 '일지 없음' 알림과 같은 시점(시작 30분 뒤)으로 맞춘다.
        if (start.HasValue && now >= start.Value.AddMinutes(OverdueUsageLogAlertMinutes))
            return entry.CheckedIn ? MissingUsageLogStatus : "노쇼";
        return "승인 완료";
    }

    public const int OverdueUsageLogAlertMinutes = 30;
    public const int UnapprovedRequestReminderMinutes = 10;

    // 웹 예약 신청 후 10분이 지나도 승인·반려하지 않은 신청(시작 시간 전인 것만).
    // 직원이 승인을 깜빡했을 때 확인창을 띄우는 대상이다.
    public static List<CalendarEntry> FindUnapprovedRequests(IEnumerable<CalendarEntry> entries, DateTime now)
    {
        return entries
            .Where(entry => entry.Source == "reservationRequest" && entry.Status == "신청" && !entry.Cancelled)
            .Where(entry =>
                DateTime.TryParse(entry.RequestedAt, out var requestedAt) &&
                now >= requestedAt.AddMinutes(UnapprovedRequestReminderMinutes) &&
                ParseReservationDateTime(entry.Date, entry.StartTime) is { } start &&
                now < start)
            .OrderBy(entry => entry.Date, StringComparer.Ordinal)
            .ThenBy(entry => entry.StartTime, StringComparer.Ordinal)
            .ToList();
    }

    // 오늘 예약 중 시작 시각에서 30분이 지났는데 일지가 없는 예약.
    // 직원이 노쇼 처리하거나 일지를 대리 제출해야 하는 대상이다.
    public static List<CalendarEntry> FindOverdueUsageLogs(IEnumerable<CalendarEntry> entries, DateTime now)
    {
        return entries
            .Where(entry =>
                entry.Source != "usage" &&
                !entry.Submitted &&
                !entry.Cancelled &&
                entry.Status is not ("신청" or "취소" or "반려" or "노쇼" or "이용완료" or "일지제출완료") &&
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

    private static bool IsNaverSource(CalendarEntry item) => item.Source is "naver" or "naverSmartPlace";

    private static string NaverReservationNo(CalendarEntry item) =>
        Regex.Replace(item.Id ?? "", @"^(naver|naverSmartPlace):", "");

    // 같은 네이버 예약번호의 시트 줄(naver)과 네이버에서 읽은 줄(naverSmartPlace)을 한 줄로 합친다.
    // 기준은 시트 줄이다. 직원이 처리한 상태(노쇼·취소·이용완료)와 일지 연결은 시트에만 있다.
    // 네이버에서 취소된 것은 시트 반영 전이라도 바로 취소로 보여 준다.
    public static List<CalendarEntry> CombineNaverSources(IEnumerable<CalendarEntry> entries)
    {
        var list = entries.ToList();
        var result = list.Where(item => !IsNaverSource(item)).ToList();
        foreach (var group in list.Where(IsNaverSource).GroupBy(NaverReservationNo, StringComparer.Ordinal))
        {
            var sheet = group.FirstOrDefault(item => item.Source == "naver");
            var live = group.LastOrDefault(item => item.Source == "naverSmartPlace");
            if (sheet is null || live is null)
            {
                result.Add(sheet ?? live ?? group.First());
                continue;
            }

            var sheetHandled = sheet.Status is "노쇼" or "취소" or "반려" or "이용완료" || sheet.Submitted;
            var liveCancelled = live.Cancelled || live.Status == "취소";
            result.Add(new CalendarEntry
            {
                Id = sheet.Id,
                Source = sheet.Source,
                RequestedAt = sheet.RequestedAt,
                MemberId = FirstNonEmpty(sheet.MemberId, live.MemberId),
                Date = FirstNonEmpty(sheet.Date, live.Date),
                StartTime = FirstNonEmpty(sheet.StartTime, live.StartTime),
                EndTime = FirstNonEmpty(sheet.EndTime, live.EndTime),
                PlaceName = FirstNonEmpty(sheet.PlaceName, live.PlaceName),
                SpaceType = FirstNonEmpty(sheet.SpaceType, live.SpaceType),
                Name = FirstNonEmpty(sheet.Name, live.Name),
                Purpose = FirstNonEmpty(sheet.Purpose, live.Purpose),
                Headcount = Math.Max(1, sheet.Headcount),
                Status = sheetHandled ? sheet.Status : liveCancelled ? "취소" : FirstNonEmpty(sheet.Status, live.Status),
                Cancelled = sheetHandled ? sheet.Cancelled : sheet.Cancelled || liveCancelled,
                Submitted = sheet.Submitted || live.Submitted,
                CheckedIn = sheet.CheckedIn || live.CheckedIn,
                PhoneTail = sheet.PhoneTail,
                NaverPhoneTail = FirstNonEmpty(live.NaverPhoneTail, sheet.NaverPhoneTail)
            });
        }
        return result;
    }

    private static string ReservationMergeKey(CalendarEntry item)
    {
        // 취소된 예약은 같은 시간에 새로 잡은 예약과 합치지 않는다(취소 후 재예약이 취소로 보이지 않게).
        if (item.Cancelled || item.Status is "취소" or "반려")
            return $"id:{item.Source}:{item.Id}";
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
        // 직원이 시트에서 노쇼로 바꾼 예약은, 네이버에서 읽은 같은 예약이 '확정'이어도 노쇼로 보여 준다.
        var cancelled = entries.Any(item => item.Status is "취소" or "반려" || (item.Cancelled && item.Status != "노쇼"));
        var noShow = !cancelled && entries.Any(item => item.Status == "노쇼");

        // 일지(usage)와 합쳐져도 예약 번호를 유지해야 상태 변경·상세보기가 그 예약을 가리킨다.
        var identity = primary.Source == "usage"
            ? entries.FirstOrDefault(item => item.Source is "naver" or "reservationRequest") ?? primary
            : primary;
        return new CalendarEntry
        {
            Id = identity.Id,
            Source = identity.Source,
            MemberId = FirstNonEmpty(entries.Select(item => item.MemberId).Prepend(primary.MemberId).ToArray()),
            Date = primary.Date,
            StartTime = primary.StartTime,
            EndTime = primary.EndTime,
            PlaceName = FirstNonEmpty(entries.Select(item => item.PlaceName).Prepend(primary.PlaceName).ToArray()),
            SpaceType = FirstNonEmpty(entries.Select(item => item.SpaceType).Prepend(primary.SpaceType).ToArray()),
            Name = primary.Name,
            Purpose = FirstNonEmpty(entries.Select(item => item.Purpose).Prepend(primary.Purpose).ToArray()),
            Headcount = Math.Max(1, entries.Max(item => item.Headcount)),
            Status = noShow ? "노쇼" : primary.Status,
            Cancelled = cancelled,
            Submitted = entries.Any(item => item.Submitted || item.Source == "usage" || item.Status == "일지제출완료"),
            CheckedIn = entries.Any(item => item.CheckedIn),
            RequestedAt = FirstNonEmpty(entries.Select(item => item.RequestedAt).Prepend(primary.RequestedAt).ToArray()),
            PhoneTail = FirstNonEmpty(entries.Select(item => item.PhoneTail).Prepend(primary.PhoneTail).ToArray()),
            NaverPhoneTail = FirstNonEmpty(entries.Select(item => item.NaverPhoneTail).Prepend(primary.NaverPhoneTail).ToArray())
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
