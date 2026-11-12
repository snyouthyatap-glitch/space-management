using System.Drawing;

namespace YouthSpring.StaffDashboard;

internal static class CalendarEntryPresenter
{
    public static CalendarDayCounts CountReservationsForDate(
        DateTime date,
        IEnumerable<CalendarEntry> calendarEntries,
        Func<IEnumerable<CalendarEntry>, IEnumerable<CalendarEntry>> applyFilters)
    {
        var entries = applyFilters(calendarEntries.Where(item => ReservationService.IsSameCalendarDate(item.Date, date)));
        var pending = 0;
        var reserved = 0;
        foreach (var item in entries)
        {
            var status = ReservationService.GetUnifiedStatus(item);
            if (status == "승인 대기") pending++;
            else if (status is "승인 완료" or "이용중" or "이용완료" or ReservationService.MissingUsageLogStatus) reserved++;
        }

        return new CalendarDayCounts(pending, reserved);
    }

    public static List<CalendarPreviewItem> BuildPreviewItems(
        DateTime date,
        IEnumerable<CalendarEntry> calendarEntries,
        Func<IEnumerable<CalendarEntry>, IEnumerable<CalendarEntry>> applyFilters)
    {
        var items = new List<CalendarPreviewItem>();
        var entries = applyFilters(calendarEntries.Where(item => ReservationService.IsSameCalendarDate(item.Date, date)))
            .OrderBy(item => item.StartTime, StringComparer.Ordinal);

        foreach (var entry in entries)
        {
            var status = ReservationService.GetUnifiedStatus(entry);
            var color = status switch
            {
                "승인 대기" => Color.FromArgb(154, 91, 25),
                ReservationService.MissingUsageLogStatus => Color.FromArgb(196, 84, 20),
                "취소" or "노쇼" => Color.FromArgb(120, 115, 112),
                _ => Color.FromArgb(48, 120, 56)
            };
            var time = string.IsNullOrWhiteSpace(entry.StartTime) ? "" : entry.StartTime + " ";
            var place = ShortCalendarPlace(entry.PlaceName);
            items.Add(new CalendarPreviewItem($"{time}{place}", color));
        }

        return items;
    }

    private static string ShortCalendarPlace(string placeName)
    {
        if (placeName.Contains("Room1", StringComparison.OrdinalIgnoreCase)) return "커리어 R1";
        if (placeName.Contains("Room2", StringComparison.OrdinalIgnoreCase)) return "커리어 R2";
        if (placeName.Contains("커넥트", StringComparison.OrdinalIgnoreCase)) return "커넥트룸";
        if (placeName.Contains("커리어", StringComparison.OrdinalIgnoreCase)) return "커리어존";
        return string.IsNullOrWhiteSpace(placeName) ? "예약" : placeName;
    }
}

internal readonly record struct CalendarDayCounts(int Pending, int Reserved);

internal readonly record struct CalendarPreviewItem(string Text, Color Color);
