using System.Diagnostics;

namespace YouthSpring.StaffDashboard;

internal sealed class MainForm : Form
{
    private const string SheetUrl =
        "https://docs.google.com/spreadsheets/d/1sPVvNPMl0KqQHDpCqNnM4RuOUw73DpgzK7IqcMDlVYs/edit?usp=sharing";
    private const string SmartPlaceUrl =
        "https://new.smartplace.naver.com/bizes/place/6800416?bookingBusinessId=1002115";
    private const string ChromePath = @"C:\Program Files\Google\Chrome\Application\chrome.exe";
    private const string ChromeProfile = "Profile 6";

    private readonly AppStorage _storage = new();
    private readonly AppsScriptClient _client = new();
    private readonly System.Windows.Forms.Timer _pollTimer = new();
    private readonly System.Windows.Forms.Timer _naverBookingTimer = new();
    private readonly NotifyIcon _notifyIcon = new();
    private readonly ComboBox _categoryFilter = new();
    private readonly CheckBox _unreadOnly = new();
    private readonly DataGridView _eventsGrid = new();
    private readonly DataGridView _reservationsGrid = new();
    private readonly TableLayoutPanel _monthGrid = new();
    private readonly Label _calendarMonthLabel = new();
    private readonly Button _previousMonthButton = new();
    private readonly Button _nextMonthButton = new();
    private readonly ComboBox _reservationStatusFilter = new();
    private readonly ComboBox _reservationSpaceFilter = new();
    private readonly Label _reservationsTitle = CreateSectionTitle("오늘 예약");
    private readonly Label _statusLabel = new();
    private readonly Label _naverBookingStatusLabel = new();
    private readonly Label _summaryStatusLabel = new();
    private readonly Label _todayVisitMetricLabel = new();
    private readonly Label _weeklyVisitMetricLabel = new();
    private readonly Label _monthlyVisitMetricLabel = new();
    private readonly Label _memberMetricLabel = new();
    private readonly Label _unreadLabel = new();
    private readonly Label _reservationSummaryLabel = new();
    private readonly Label _reservationFilterSummaryLabel = new();
    private readonly FlowLayoutPanel _reservationActionPanel = new();
    private readonly Button _requestApproveButton = new();
    private readonly Button _requestRejectButton = new();
    private readonly Button _proxyUsageLogButton = new();
    private readonly Button _naverPendingButton = new();
    private readonly Button _missingUsageLogButton = new();
    private readonly Button _moreActionsButton = new();
    private DashboardSettings _settings;
    private LocalState _state;
    private string _password;
    private string _naverId;
    private string _naverPassword;
    private string _loadedCalendarMonth = "";
    private DateTime _selectedCalendarDate = DateTime.Today;
    private DateTime _visibleCalendarMonth = new(DateTime.Today.Year, DateTime.Today.Month, 1);
    private List<CalendarEntry> _calendarEntries = [];
    private readonly Dictionary<string, List<CalendarEntry>> _calendarCacheByMonth = new(StringComparer.Ordinal);
    private readonly Dictionary<string, DateTime> _calendarCacheUpdatedAt = new(StringComparer.Ordinal);
    private readonly Dictionary<string, List<StaffCheckIn>> _checkInsByMonth = new(StringComparer.Ordinal);
    private DateTime? _pendingCalendarMonthRefresh;
    private bool _refreshing;
    private bool _updatingReservationStatus;
    private bool _runtimeSyncCompleted;
    private bool _naverBookingSyncRunning;
    private bool _allowClose;
    private bool _trayNoticeShown;
    private string _lastNaverPendingSignature = "";
    private string _lastNaverDialogSignature = "";
    private NaverBookingSnapshot _lastNaverBookingSnapshot = new();
    private readonly bool _demoMode;
    private readonly string _previewPath;

    public MainForm(bool demoMode = false, string previewPath = "")
    {
        _demoMode = demoMode;
        _previewPath = previewPath;
        _settings = _storage.LoadSettings();
        _state = _storage.LoadState();
        _password = _storage.UnprotectPassword(_settings.ProtectedPassword);
        _naverId = _storage.UnprotectPassword(_settings.ProtectedNaverId);
        _naverPassword = _storage.UnprotectPassword(_settings.ProtectedNaverPassword);

        Text = "청년이봄 직원 운영";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(1100, 680);
        Size = new Size(1380, 860);
        Font = new Font("맑은 고딕", 9.5F);
        BackColor = Color.FromArgb(247, 247, 246);
        Icon = SystemIcons.Application;

        BuildLayout();
        ConfigureTray();

        _pollTimer.Interval = Math.Max(15, _settings.PollSeconds) * 1000;
        _pollTimer.Tick += async (_, _) => await RefreshDashboardAsync();
        _naverBookingTimer.Interval = 5 * 60 * 1000;
        _naverBookingTimer.Tick += async (_, _) => await RefreshNaverBookingStatusAsync();

        Shown += async (_, _) =>
        {
            // 첫 버튼에 포커스 테두리가 생기지 않도록 예약 표에 포커스를 둔다.
            ActiveControl = _reservationsGrid;
            if (_demoMode)
            {
                LoadDemoData();
                if (_previewPath.EndsWith("|naver", StringComparison.Ordinal))
                {
                    // 예약신청 창 미리보기 캡처: --preview=파일.png|naver
                    using var naverDialog = BuildNaverBookingDialog(CreateDemoNaverSnapshot(), DateTime.Now);
                    naverDialog.StartPosition = FormStartPosition.CenterScreen;
                    naverDialog.Show(this);
                    await Task.Delay(600);
                    using var bitmap = new Bitmap(naverDialog.Width, naverDialog.Height);
                    naverDialog.DrawToBitmap(bitmap, new Rectangle(0, 0, naverDialog.Width, naverDialog.Height));
                    bitmap.Save(_previewPath[..^"|naver".Length]);
                    naverDialog.Close();
                    _allowClose = true;
                    Close();
                    return;
                }
                if (!string.IsNullOrWhiteSpace(_previewPath))
                {
                    await Task.Delay(600);
                    CapturePreview(_previewPath);
                    _allowClose = true;
                    Close();
                }
                return;
            }

            if (string.IsNullOrWhiteSpace(_password))
            {
                if (!ShowSettings()) return;
            }
            await RefreshDashboardAsync();
            _pollTimer.Start();
            _naverBookingTimer.Start();
            _ = RefreshNaverBookingStatusAsync();
        };

        Resize += (_, _) =>
        {
            if (WindowState == FormWindowState.Minimized)
            {
                Hide();
                _notifyIcon.ShowBalloonTip(1500, "청년이봄 직원 운영",
                    "백그라운드에서 현장예약을 확인하고 있습니다.", ToolTipIcon.Info);
            }
        };
        FormClosing += HandleFormClosingToTray;
    }

    private void HandleFormClosingToTray(object? sender, FormClosingEventArgs eventArgs)
    {
        if (_allowClose || eventArgs.CloseReason == CloseReason.WindowsShutDown)
        {
            _notifyIcon.Visible = false;
            return;
        }

        eventArgs.Cancel = true;
        HideToTray(showBalloonTip: !_trayNoticeShown);
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(18),
            BackColor = BackColor
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 96));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
        Controls.Add(root);

        root.Controls.Add(BuildHeader(), 0, 0);
        root.Controls.Add(BuildMainContent(), 0, 1);
        root.Controls.Add(BuildStatusBar(), 0, 2);
    }

    private Control BuildHeader()
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        var titlePanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(0, 2, 0, 4)
        };
        titlePanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
        titlePanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        var title = new Label
        {
            Text = "청년이봄 운영 캘린더",
            Font = new Font("맑은 고딕", 18F, FontStyle.Bold),
            ForeColor = Color.FromArgb(52, 48, 46),
            Dock = DockStyle.Fill,
            AutoSize = false,
            TextAlign = ContentAlignment.MiddleLeft
        };
        var metricsPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(0, 1, 0, 0)
        };
        metricsPanel.Controls.AddRange([
            CreateMetricBadge(_todayVisitMetricLabel, "오늘 방문", "0"),
            CreateMetricBadge(_weeklyVisitMetricLabel, "주간", "0"),
            CreateMetricBadge(_monthlyVisitMetricLabel, "월간", "0"),
            CreateMetricBadge(_memberMetricLabel, "회원", "0")
        ]);
        titlePanel.Controls.Add(title, 0, 0);
        titlePanel.Controls.Add(metricsPanel, 0, 1);

        var actions = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(0, 20, 0, 0)
        };
        actions.Controls.Add(CreateHeaderButton("네이버", (_, _) => OpenChrome(SmartPlaceUrl)));
        ConfigureNaverPendingButton();
        actions.Controls.Add(_naverPendingButton);
        ConfigureMissingUsageLogButton();
        actions.Controls.Add(_missingUsageLogButton);
        actions.Controls.Add(CreateHeaderButton("새로고침", async (_, _) => await RefreshAllAsync()));
        ConfigureMoreActionsButton();
        actions.Controls.Add(_moreActionsButton);
        panel.Controls.Add(titlePanel, 0, 0);
        panel.Controls.Add(actions, 1, 0);
        return panel;
    }

    private static Label CreateMetricBadge(Label label, string caption, string value)
    {
        label.Text = $"{caption} {value}";
        label.AutoSize = true;
        label.Margin = new Padding(0, 0, 8, 0);
        label.Padding = new Padding(9, 4, 9, 5);
        label.BackColor = Color.FromArgb(241, 246, 241);
        label.ForeColor = Color.FromArgb(52, 92, 57);
        label.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
        return label;
    }

    private void ConfigureMissingUsageLogButton()
    {
        _missingUsageLogButton.Text = "일지 미제출";
        ApplyHeaderButtonStyle(_missingUsageLogButton);
        _missingUsageLogButton.Click += (_, _) => ShowMissingUsageLogs();
    }

    private List<CalendarEntry> GetMissingUsageLogEntries() =>
        _calendarEntries
            .Where(item => ReservationService.GetUnifiedStatus(item) == ReservationService.MissingUsageLogStatus)
            .OrderBy(item => item.Date, StringComparer.Ordinal)
            .ThenBy(item => item.StartTime, StringComparer.Ordinal)
            .ToList();

    private void UpdateMissingUsageLogButton()
    {
        if (_missingUsageLogButton.IsDisposed) return;
        var count = GetMissingUsageLogEntries().Count;
        _missingUsageLogButton.Text = count > 0 ? $"일지 미제출 {count}" : "일지 미제출";
        ApplyHeaderButtonStyle(_missingUsageLogButton, count > 0 ? MissingUsageLogColor : null);
    }

    private readonly HashSet<string> _overdueAlertedKeys = new(StringComparer.Ordinal);
    private readonly System.Windows.Forms.Timer _overdueCheckTimer = new() { Interval = 60 * 1000 };
    private bool _overdueBalloonShown;

    // 오늘 예약 시작 30분이 지났는데 일지가 없으면 한 번씩 알린다(노쇼 처리·일지 대리 제출용).
    private void CheckOverdueUsageLogs()
    {
        if (!string.IsNullOrWhiteSpace(_previewPath)) return; // 화면 캡처용 실행에서는 알리지 않는다.
        var todayEntries = _visibleCalendarMonth.ToString("yyyy-MM") == DateTime.Today.ToString("yyyy-MM")
            ? _calendarEntries
            : _calendarCacheByMonth.GetValueOrDefault(DateTime.Today.ToString("yyyy-MM")) ?? [];
        var fresh = ReservationService.FindOverdueUsageLogs(todayEntries, DateTime.Now)
            .Where(entry => _overdueAlertedKeys.Add($"{entry.Date}|{entry.Id}"))
            .ToList();
        if (fresh.Count == 0) return;

        string Describe(CalendarEntry entry) =>
            $"{entry.StartTime} {entry.PlaceName} · {entry.Name} " +
            (entry.CheckedIn ? "(출석함 → 일지 대리 제출)" : "(출석 기록 없음 → 노쇼 확인)");
        var message = fresh.Count == 1
            ? Describe(fresh[0])
            : string.Join("\n", fresh.Take(3).Select(Describe)) +
              (fresh.Count > 3 ? $"\n외 {fresh.Count - 3}건" : "");

        _overdueBalloonShown = true;
        System.Media.SystemSounds.Exclamation.Play();
        _notifyIcon.ShowBalloonTip(
            10000,
            fresh.Count == 1
                ? $"일지 미제출 · 시작 {ReservationService.OverdueUsageLogAlertMinutes}분 지남"
                : $"일지 미제출 {fresh.Count}건 · 시작 {ReservationService.OverdueUsageLogAlertMinutes}분 지남",
            message + "\n눌러서 확인",
            ToolTipIcon.Warning);
        _statusLabel.Text = $"시작 {ReservationService.OverdueUsageLogAlertMinutes}분이 지났는데 일지가 없는 예약 {fresh.Count}건 · 노쇼 처리 또는 일지 대리 제출이 필요합니다.";
    }

    private void ShowTodayOverdueReservations()
    {
        RestoreFromTray();
        _selectedCalendarDate = DateTime.Today;
        _visibleCalendarMonth = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        _reservationStatusFilter.SelectedIndex = 0;
        _ = SelectCalendarDateAsync(DateTime.Today);
    }

    private void ShowMissingUsageLogs()
    {
        var missing = GetMissingUsageLogEntries();
        if (missing.Count == 0)
        {
            MessageBox.Show(
                this,
                "이번 달에 출석은 했지만 일지를 제출하지 않은 예약이 없습니다.",
                "일지 미제출",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        var target = missing.FirstOrDefault(item => ReservationService.IsSameCalendarDate(item.Date, DateTime.Today))
            ?? missing[^1];
        if (DateTime.TryParse(target.Date, out var targetDate)) _selectedCalendarDate = targetDate.Date;
        _reservationStatusFilter.SelectedItem = ReservationService.MissingUsageLogStatus;
        RenderCalendarEntries();
        _statusLabel.Text = $"일지 미제출 {missing.Count}건 · 예약을 선택하고 '일지 대리 제출'을 눌러 주세요.";
    }

    private void ConfigureNaverPendingButton()
    {
        _naverPendingButton.Text = "예약신청";
        ApplyHeaderButtonStyle(_naverPendingButton);
        _naverPendingButton.Click += (_, _) => ShowNaverBookingDialog();
        UpdateNaverPendingButton(0);
    }

    private Control BuildMainContent()
    {
        return BuildReservationsPanel();
    }

    private Control BuildEventsPanel()
    {
        var panel = CreateSectionPanel();
        var header = new Panel { Dock = DockStyle.Top, Height = 58 };
        var title = CreateSectionTitle("알림 이력");
        title.Location = new Point(16, 16);
        header.Controls.Add(title);

        _unreadLabel.AutoSize = true;
        _unreadLabel.ForeColor = Color.FromArgb(233, 63, 111);
        _unreadLabel.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
        _unreadLabel.Location = new Point(104, 21);
        header.Controls.Add(_unreadLabel);

        var filterPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Right,
            Width = 450,
            Padding = new Padding(0, 12, 12, 0),
            FlowDirection = FlowDirection.LeftToRight
        };
        _categoryFilter.DropDownStyle = ComboBoxStyle.DropDownList;
        _categoryFilter.Width = 120;
        _categoryFilter.Items.AddRange(["전체", "방문", "커넥트룸", "커리어존", "이용 제한"]);
        _categoryFilter.SelectedIndex = 0;
        _categoryFilter.SelectedIndexChanged += (_, _) => RenderEvents();
        filterPanel.Controls.Add(_categoryFilter);

        _unreadOnly.Text = "미확인만";
        _unreadOnly.AutoSize = true;
        _unreadOnly.Padding = new Padding(8, 5, 8, 0);
        _unreadOnly.CheckedChanged += (_, _) => RenderEvents();
        filterPanel.Controls.Add(_unreadOnly);
        filterPanel.Controls.Add(CreateSmallButton("모두 확인", (_, _) => MarkAllRead()));
        header.Controls.Add(filterPanel);
        panel.Controls.Add(header);

        ConfigureGrid(_eventsGrid);
        _eventsGrid.Columns.Add("read", "상태");
        _eventsGrid.Columns.Add("category", "분류");
        _eventsGrid.Columns.Add("time", "등록 시각");
        _eventsGrid.Columns.Add("name", "이름");
        _eventsGrid.Columns.Add("detail", "내용");
        _eventsGrid.Columns["read"]!.Width = 68;
        _eventsGrid.Columns["category"]!.Width = 92;
        _eventsGrid.Columns["time"]!.Width = 132;
        _eventsGrid.Columns["name"]!.Width = 92;
        _eventsGrid.Columns["detail"]!.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        _eventsGrid.CellDoubleClick += (_, eventArgs) =>
        {
            if (eventArgs.RowIndex < 0) return;
            if (_eventsGrid.Rows[eventArgs.RowIndex].Tag is StaffEvent item)
            {
                item.IsRead = true;
                _storage.SaveState(_state);
                RenderEvents();
            }
        };
        panel.Controls.Add(_eventsGrid);
        _eventsGrid.BringToFront();
        return panel;
    }

    private Control BuildReservationsPanel()
    {
        var panel = CreateSectionPanel();
        var shell = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = Color.White
        };
        shell.RowStyles.Add(new RowStyle(SizeType.Absolute, 82));
        shell.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        panel.Controls.Add(shell);

        var headerPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(16, 12, 16, 8),
            ColumnCount = 2,
            RowCount = 2
        };
        headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        headerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        headerPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
        headerPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        _reservationsTitle.AutoSize = false;
        _reservationsTitle.Dock = DockStyle.Fill;
        _reservationsTitle.Height = 28;
        headerPanel.Controls.Add(_reservationsTitle, 0, 0);

        _reservationSummaryLabel.Dock = DockStyle.Fill;
        _reservationSummaryLabel.ForeColor = Color.FromArgb(105, 100, 97);
        _reservationSummaryLabel.TextAlign = ContentAlignment.MiddleLeft;
        headerPanel.Controls.Add(_reservationSummaryLabel, 0, 1);

        shell.Controls.Add(headerPanel, 0, 0);

        var bodySplit = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,
            SplitterWidth = 8,
            BackColor = Color.FromArgb(232, 230, 228)
        };
        var bodySplitInitialized = false;
        bodySplit.SizeChanged += (_, _) =>
        {
            if (bodySplitInitialized) return;
            var availableWidth = bodySplit.Width - bodySplit.SplitterWidth;
            const int calendarMinWidth = 520;
            const int detailMinWidth = 520;
            if (availableWidth <= calendarMinWidth + detailMinWidth) return;
            bodySplit.SplitterDistance = Math.Min((int)(availableWidth * 0.58), availableWidth - detailMinWidth);
            bodySplitInitialized = true;
        };

        var calendarPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(14, 12, 14, 12),
            BackColor = Color.White,
            ColumnCount = 1,
            RowCount = 4
        };
        calendarPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
        calendarPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        calendarPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 74));
        calendarPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));

        var calendarHeader = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 1
        };
        calendarHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 38));
        calendarHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        calendarHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 38));
        ConfigureCalendarNavButton(_previousMonthButton, "<");
        ConfigureCalendarNavButton(_nextMonthButton, ">");
        _previousMonthButton.Click += async (_, _) => await ShiftCalendarMonthAsync(-1);
        _nextMonthButton.Click += async (_, _) => await ShiftCalendarMonthAsync(1);
        _calendarMonthLabel.Dock = DockStyle.Fill;
        _calendarMonthLabel.Font = new Font("맑은 고딕", 12F, FontStyle.Bold);
        _calendarMonthLabel.ForeColor = Color.FromArgb(52, 48, 46);
        _calendarMonthLabel.TextAlign = ContentAlignment.MiddleCenter;
        calendarHeader.Controls.Add(_previousMonthButton, 0, 0);
        calendarHeader.Controls.Add(_calendarMonthLabel, 1, 0);
        calendarHeader.Controls.Add(_nextMonthButton, 2, 0);
        calendarPanel.Controls.Add(calendarHeader, 0, 0);

        ConfigureMonthGrid();
        calendarPanel.Controls.Add(_monthGrid, 0, 1);

        var filterPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1
        };
        filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        filterPanel.Controls.Add(BuildReservationFilterBlock("상태", _reservationStatusFilter, [
            "전체 상태",
            "승인 대기",
            "승인 완료",
            "이용중",
            "이용완료",
            ReservationService.MissingUsageLogStatus,
            "노쇼",
            "취소"
        ]), 0, 0);
        filterPanel.Controls.Add(BuildReservationFilterBlock("공간", _reservationSpaceFilter, [
            "전체 공간",
            "커리어존 Room1",
            "커리어존 Room2",
            "커넥트룸"
        ]), 1, 0);
        calendarPanel.Controls.Add(filterPanel, 0, 2);

        _reservationFilterSummaryLabel.Dock = DockStyle.Fill;
        _reservationFilterSummaryLabel.ForeColor = Color.FromArgb(105, 100, 97);
        _reservationFilterSummaryLabel.Padding = new Padding(4, 4, 0, 4);
        _reservationFilterSummaryLabel.TextAlign = ContentAlignment.MiddleLeft;
        _reservationFilterSummaryLabel.AutoSize = false;
        calendarPanel.Controls.Add(_reservationFilterSummaryLabel, 0, 3);
        bodySplit.Panel1.Controls.Add(calendarPanel);

        var gridPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12, 8, 12, 12),
            BackColor = Color.White
        };
        ConfigureGrid(_reservationsGrid);
        _reservationsGrid.Columns.Add("time", "시간");
        _reservationsGrid.Columns.Add("place", "장소");
        _reservationsGrid.Columns.Add("representative", "대표자");
        _reservationsGrid.Columns.Add("headcount", "인원");
        _reservationsGrid.Columns.Add("status", "상태");
        // 표 너비에 맞춰 늘고 줄게 해서 '상태' 칸이 잘리거나 가로 스크롤이 생기지 않게 한다.
        foreach (var (name, weight, minimumWidth) in new[]
                 {
                     ("time", 24F, 100),
                     ("place", 30F, 112),
                     ("representative", 20F, 72),
                     ("headcount", 10F, 52),
                     ("status", 22F, 112)
                 })
        {
            var column = _reservationsGrid.Columns[name]!;
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            column.FillWeight = weight;
            column.MinimumWidth = minimumWidth;
            column.Resizable = DataGridViewTriState.True;
        }
        _reservationsGrid.Columns["status"]!.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        _reservationsGrid.Columns["headcount"]!.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        _reservationsGrid.RowTemplate.Height = 44;
        _reservationsGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
        _reservationsGrid.DefaultCellStyle.Padding = new Padding(8, 0, 8, 0);
        _reservationsGrid.ScrollBars = ScrollBars.Vertical;
        _reservationsGrid.SelectionChanged += (_, _) => UpdateReservationActionButtons();
        _reservationsGrid.CellDoubleClick += (_, eventArgs) =>
        {
            if (eventArgs.RowIndex < 0) return;
            if (_reservationsGrid.Rows[eventArgs.RowIndex].Tag is CalendarEntry entry)
                ShowReservationDetailDialog(entry);
        };
        ConfigureReservationContextMenu();
        ConfigureReservationActions();
        gridPanel.Controls.Add(_reservationActionPanel);
        gridPanel.Controls.Add(_reservationsGrid);
        bodySplit.Panel2.Controls.Add(gridPanel);
        shell.Controls.Add(bodySplit, 0, 1);
        _reservationsGrid.BringToFront();
        return panel;
    }

    private void ConfigureCalendarNavButton(Button button, string text)
    {
        button.Text = text;
        button.Dock = DockStyle.Fill;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderColor = Color.FromArgb(218, 214, 210);
        button.BackColor = Color.White;
        button.ForeColor = Color.FromArgb(52, 48, 46);
        button.Font = new Font("맑은 고딕", 10F, FontStyle.Bold);
        button.Margin = new Padding(0, 2, 0, 2);
    }

    private void ConfigureMonthGrid()
    {
        _monthGrid.Dock = DockStyle.Fill;
        _monthGrid.BackColor = Color.FromArgb(232, 230, 228);
        _monthGrid.ColumnCount = 7;
        _monthGrid.RowCount = 7;
        _monthGrid.Margin = new Padding(0, 6, 0, 8);
        _monthGrid.Padding = new Padding(1);
        _monthGrid.GrowStyle = TableLayoutPanelGrowStyle.FixedSize;
        _monthGrid.ColumnStyles.Clear();
        _monthGrid.RowStyles.Clear();
        for (var column = 0; column < 7; column++)
            _monthGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F / 7F));
        _monthGrid.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
        for (var row = 1; row < 7; row++)
            _monthGrid.RowStyles.Add(new RowStyle(SizeType.Percent, 100F / 6F));
        _monthGrid.SizeChanged += (_, _) =>
        {
            var lines = ComputeCalendarPreviewLines();
            if (lines == _calendarPreviewLines) return;
            _calendarPreviewLines = lines;
            RenderMonthGrid();
        };
    }

    private const int CalendarDayLabelHeight = 18;
    private const int CalendarPreviewLineHeight = 14;
    private const int CalendarCellVerticalPadding = 5;
    private int _calendarPreviewLines = 2;

    // 날짜 칸 높이에 실제로 들어가는 미리보기 줄 수(날짜 줄 제외)
    private int ComputeCalendarPreviewLines()
    {
        var gridHeight = _monthGrid.ClientSize.Height - _monthGrid.Padding.Vertical - 28;
        if (gridHeight <= 0) return _calendarPreviewLines;
        var cellHeight = gridHeight / 6 - 1;
        var available = cellHeight - CalendarCellVerticalPadding - CalendarDayLabelHeight;
        return Math.Clamp(available / CalendarPreviewLineHeight, 0, 4);
    }

    private async Task ShiftCalendarMonthAsync(int delta)
    {
        var month = _visibleCalendarMonth.AddMonths(delta);
        await SelectCalendarDateAsync(month);
    }

    private async Task SelectCalendarDateAsync(DateTime date)
    {
        _selectedCalendarDate = date.Date;
        _visibleCalendarMonth = new DateTime(date.Year, date.Month, 1);
        var month = _visibleCalendarMonth.ToString("yyyy-MM");
        if (_demoMode)
        {
            RenderCalendarEntries();
            return;
        }

        if (_calendarCacheByMonth.TryGetValue(month, out var cachedEntries))
        {
            _calendarEntries = cachedEntries.ToList();
            _loadedCalendarMonth = month;
        }

        RenderCalendarEntries();

        if (ShouldRefreshCalendarMonth(month))
        {
            _statusLabel.Text = IsPastCalendarMonth(_visibleCalendarMonth)
                ? "저장된 월 데이터를 확인하는 중..."
                : "예약 변동을 확인하는 중...";
            QueueCalendarMonthRefresh(_visibleCalendarMonth);
        }
        await Task.CompletedTask;
    }

    private void QueueCalendarMonthRefresh(DateTime calendarMonth)
    {
        if (_refreshing)
        {
            _pendingCalendarMonthRefresh = calendarMonth;
            return;
        }

        _ = RefreshDashboardAsync(calendarMonth, renderWhenComplete: true);
    }

    private bool ShouldRefreshCalendarMonth(string month)
    {
        if (!_calendarCacheByMonth.ContainsKey(month)) return true;
        if (!DateTime.TryParse(month + "-01", out var monthStart)) return true;
        if (IsPastCalendarMonth(monthStart)) return false;
        if (!_calendarCacheUpdatedAt.TryGetValue(month, out var updatedAt)) return true;
        var refreshInterval = TimeSpan.FromSeconds(Math.Max(30, _settings.PollSeconds));
        return DateTime.Now - updatedAt > refreshInterval;
    }

    private static bool IsPastCalendarMonth(DateTime monthStart)
    {
        var normalizedMonth = new DateTime(monthStart.Year, monthStart.Month, 1);
        var currentMonth = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        return normalizedMonth < currentMonth;
    }

    private void CacheVisibleCalendarEntries()
    {
        var month = _visibleCalendarMonth.ToString("yyyy-MM");
        _calendarCacheByMonth[month] = _calendarEntries.ToList();
        _calendarCacheUpdatedAt[month] = DateTime.Now;
        _loadedCalendarMonth = month;
    }

    private void RenderMonthGrid()
    {
        if (_monthGrid.ColumnCount == 0) return;

        _calendarPreviewLines = ComputeCalendarPreviewLines();
        _monthGrid.SuspendLayout();
        _monthGrid.Controls.Clear();
        _calendarMonthLabel.Text = _visibleCalendarMonth.ToString("yyyy년 M월");

        var weekdays = new[] { "월", "화", "수", "목", "금", "토", "일" };
        for (var column = 0; column < weekdays.Length; column++)
        {
            _monthGrid.Controls.Add(new Label
            {
                Text = weekdays[column],
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(120, 115, 112),
                Font = new Font("맑은 고딕", 8F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = new Padding(0, 0, 1, 1)
            }, column, 0);
        }

        var firstDay = _visibleCalendarMonth;
        var mondayOffset = ((int)firstDay.DayOfWeek + 6) % 7;
        var startDate = firstDay.AddDays(-mondayOffset);
        var today = DateTime.Today;

        for (var index = 0; index < 42; index++)
        {
            var date = startDate.AddDays(index);
            var counts = CountReservationsForDate(date);
            var previewItems = BuildCalendarPreviewItems(date);
            var isCurrentMonth = date.Month == _visibleCalendarMonth.Month;
            var isSelected = date.Date == _selectedCalendarDate.Date;
            var hasCounts = counts.Pending > 0 || counts.Reserved > 0;

            var cell = CreateCalendarDateCell(
                date,
                counts,
                previewItems,
                isCurrentMonth,
                isSelected,
                date.Date == today,
                hasCounts);
            _monthGrid.Controls.Add(cell, index % 7, (index / 7) + 1);
        }

        _monthGrid.ResumeLayout();
    }

    private Panel CreateCalendarDateCell(
        DateTime date,
        CalendarDayCounts counts,
        IReadOnlyList<CalendarPreviewItem> previewItems,
        bool isCurrentMonth,
        bool isSelected,
        bool isToday,
        bool hasCounts)
    {
        var cell = new Panel
        {
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 0, 1, 1),
            Padding = new Padding(7, 3, 6, 2),
            BackColor = isSelected
                ? Color.FromArgb(232, 245, 230)
                : hasCounts
                    ? Color.FromArgb(250, 250, 248)
                    : Color.White,
            BorderStyle = isSelected || isToday ? BorderStyle.FixedSingle : BorderStyle.None,
            Cursor = Cursors.Hand,
            Tag = date.Date
        };
        cell.Click += async (sender, _) => await SelectCalendarDateFromControlAsync(sender);

        var content = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoScroll = false,
            BackColor = Color.Transparent,
            Margin = new Padding(0),
            Padding = new Padding(0),
            Cursor = Cursors.Hand,
            Tag = date.Date
        };
        content.Click += async (sender, _) => await SelectCalendarDateFromControlAsync(sender);

        var dayLabel = new Label
        {
            Text = BuildCalendarHeaderText(date, counts),
            Width = 140,
            Height = CalendarDayLabelHeight - 1,
            AutoSize = false,
            ForeColor = isCurrentMonth ? Color.FromArgb(52, 48, 46) : Color.FromArgb(173, 168, 164),
            Font = new Font("맑은 고딕", 8.5F, hasCounts ? FontStyle.Bold : FontStyle.Regular),
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 0, 0, 1),
            Tag = date.Date
        };
        dayLabel.Click += async (sender, _) => await SelectCalendarDateFromControlAsync(sender);
        content.Controls.Add(dayLabel);

        if (false && counts.Pending > 0)
            content.Controls.Add(CreateCalendarCountLabel($"신청 {counts.Pending}", Color.FromArgb(154, 91, 25), date));
        if (false && counts.Reserved > 0)
            content.Controls.Add(CreateCalendarCountLabel($"예약 {counts.Reserved}", Color.FromArgb(48, 120, 56), date));

        // 칸에 들어가는 줄 수만큼만 보여 준다. 넘치면 마지막 줄을 '외 N건'으로 쓰고,
        // 한 줄밖에 없으면 첫 예약만 보여 준다(건수는 날짜 줄에 이미 있다).
        var maxPreviewLines = _calendarPreviewLines;
        if (maxPreviewLines > 0 && previewItems.Count > 0)
        {
            var overflow = previewItems.Count > maxPreviewLines && maxPreviewLines > 1;
            var shownCount = overflow
                ? maxPreviewLines - 1
                : Math.Min(previewItems.Count, maxPreviewLines);
            foreach (var item in previewItems.Take(shownCount))
                content.Controls.Add(CreateCalendarPreviewLabel(item, date));
            if (overflow)
                content.Controls.Add(CreateCalendarPreviewLabel(
                    new CalendarPreviewItem($"외 {previewItems.Count - shownCount}건", Color.FromArgb(120, 115, 112)),
                    date));
        }

        cell.Controls.Add(content);

        return cell;
    }

    private static string BuildCalendarHeaderText(DateTime date, CalendarDayCounts counts)
    {
        if (counts.Pending > 0) return $"{date.Day}  신청 {counts.Pending}";
        if (counts.Reserved > 0) return $"{date.Day}  예약 {counts.Reserved}";
        return date.Day.ToString();
    }

    private Label CreateCalendarCountLabel(string text, Color color, DateTime date)
    {
        var label = new Label
        {
            Text = text,
            Width = 120,
            Height = 16,
            AutoSize = false,
            ForeColor = color,
            Font = new Font("맑은 고딕", 7.5F, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Margin = new Padding(0),
            Tag = date.Date
        };
        label.Click += async (sender, _) => await SelectCalendarDateFromControlAsync(sender);
        return label;
    }

    private Label CreateCalendarPreviewLabel(CalendarPreviewItem item, DateTime date)
    {
        var label = new Label
        {
            Text = item.Text,
            Width = 140,
            Height = CalendarPreviewLineHeight,
            AutoSize = false,
            ForeColor = item.Color,
            Font = new Font("맑은 고딕", 7.5F),
            Cursor = Cursors.Hand,
            Margin = new Padding(0),
            Tag = date.Date
        };
        label.Click += async (sender, _) => await SelectCalendarDateFromControlAsync(sender);
        return label;
    }

    private async Task SelectCalendarDateFromControlAsync(object? sender)
    {
        if (sender is Control { Tag: DateTime date })
            await SelectCalendarDateAsync(date);
    }

    private CalendarDayCounts CountReservationsForDate(DateTime date)
        => CalendarEntryPresenter.CountReservationsForDate(
            date,
            _calendarEntries,
            ApplyReservationFilters);

    private List<CalendarPreviewItem> BuildCalendarPreviewItems(DateTime date)
        => CalendarEntryPresenter.BuildPreviewItems(
            date,
            _calendarEntries,
            ApplyReservationFilters);

    private void ConfigureReservationActions()
    {
        // 선택한 예약에 대한 버튼은 표 바로 아래에 둔다(예전에는 오른쪽 위에 있어 잘렸다).
        _reservationActionPanel.Dock = DockStyle.Bottom;
        _reservationActionPanel.Height = 50;
        _reservationActionPanel.Padding = new Padding(0, 10, 0, 0);
        _reservationActionPanel.FlowDirection = FlowDirection.LeftToRight;
        _reservationActionPanel.WrapContents = false;

        ConfigureReservationActionButton(_requestApproveButton, "\uC2B9\uC778", Color.FromArgb(74, 156, 68));
        ConfigureReservationActionButton(_requestRejectButton, "\uBC18\uB824", Color.FromArgb(233, 63, 111));
        ConfigureReservationActionButton(_proxyUsageLogButton, "이용완료", Color.FromArgb(196, 84, 20));
        _proxyUsageLogButton.Text = "일지 대리 제출";
        UpdateReservationActionButtons();
    }

    private void ConfigureReservationActionButton(Button button, string status, Color color)
    {
        button.Text = status;
        button.AutoSize = true;
        button.MinimumSize = new Size(72, 40);
        button.Height = 40;
        button.Padding = new Padding(10, 0, 10, 0);
        button.Margin = new Padding(0, 0, 8, 0);
        button.Cursor = Cursors.Hand;
        // 지금 누를 수 없는 버튼은 회색으로 보여 준다.
        button.EnabledChanged += (_, _) =>
            button.BackColor = button.Enabled ? color : Color.FromArgb(226, 223, 220);
        button.BackColor = color;
        button.ForeColor = Color.White;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 0;
        button.Click += async (_, _) => await UpdateSelectedReservationStatusAsync(status);
        _reservationActionPanel.Controls.Add(button);
    }

    private Control BuildReservationFilterBlock(string labelText, ComboBox comboBox, string[] items)
    {
        var block = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(0, 6, 0, 8)
        };
        block.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
        block.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var label = new Label
        {
            Text = labelText,
            Dock = DockStyle.Fill,
            ForeColor = Color.FromArgb(105, 100, 97),
            Font = new Font("맑은 고딕", 8.5F, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleLeft
        };
        comboBox.Dock = DockStyle.Fill;
        comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
        comboBox.Items.Clear();
        comboBox.Items.AddRange(items);
        comboBox.SelectedIndex = 0;
        comboBox.SelectedIndexChanged += (_, _) => RenderCalendarEntries();

        block.Controls.Add(label, 0, 0);
        block.Controls.Add(comboBox, 0, 1);
        return block;
    }

    private void ConfigureReservationContextMenu()
    {
        var menu = new ContextMenuStrip();
        var detailItem = new ToolStripMenuItem("상세 보기");
        detailItem.Click += (_, _) =>
        {
            if (_reservationsGrid.SelectedRows.Count > 0 &&
                _reservationsGrid.SelectedRows[0].Tag is CalendarEntry entry)
            {
                ShowReservationDetailDialog(entry);
            }
        };
        menu.Items.Add(detailItem);
        menu.Items.Add(new ToolStripSeparator());

        menu.Opening += (_, eventArgs) =>
        {
            while (menu.Items.Count > 2) menu.Items.RemoveAt(2);
            var mouse = _reservationsGrid.PointToClient(Cursor.Position);
            var hit = _reservationsGrid.HitTest(mouse.X, mouse.Y);
            if (hit.RowIndex < 0 ||
                _reservationsGrid.Rows[hit.RowIndex].Tag is not CalendarEntry entry)
            {
                eventArgs.Cancel = true;
                return;
            }

            _reservationsGrid.ClearSelection();
            _reservationsGrid.Rows[hit.RowIndex].Selected = true;
            _reservationsGrid.CurrentCell = _reservationsGrid.Rows[hit.RowIndex].Cells[
                Math.Max(0, hit.ColumnIndex)];
            detailItem.Enabled = true;

            var actions = ReservationStatusActionsForEntry(entry);
            if (actions.Length == 0)
            {
                var disabled = new ToolStripMenuItem("변경 가능한 상태 없음") { Enabled = false };
                menu.Items.Add(disabled);
                return;
            }

            foreach (var action in actions)
            {
                var item = new ToolStripMenuItem(action.Label) { Tag = action.Status };
                item.Click += async (_, _) => await UpdateSelectedReservationStatusAsync(action.Status);
                item.Enabled = IsAllowedStatusForEntry(entry, action.Status);
                item.Checked = StatusMatches(entry, action.Status);
                menu.Items.Add(item);
            }
        };
        _reservationsGrid.ContextMenuStrip = menu;
    }

    private void ShowReservationDetailDialog(CalendarEntry entry)
    {
        using var dialog = new Form
        {
            Text = "예약 상세정보",
            StartPosition = FormStartPosition.CenterParent,
            Size = new Size(520, 520),
            MinimizeBox = false,
            MaximizeBox = false,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            Font = Font,
            BackColor = Color.White
        };

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(18)
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
        dialog.Controls.Add(root);

        var title = new Label
        {
            Text = $"{entry.StartTime}~{entry.EndTime} · {entry.PlaceName}",
            Dock = DockStyle.Fill,
            Font = new Font("맑은 고딕", 13F, FontStyle.Bold),
            ForeColor = Color.FromArgb(52, 48, 46),
            TextAlign = ContentAlignment.MiddleLeft
        };
        root.Controls.Add(title, 0, 0);

        var details = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 0,
            BackColor = Color.White
        };
        details.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
        details.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        root.Controls.Add(details, 0, 1);

        AddDetailRow(details, "날짜", entry.Date);
        AddDetailRow(details, "시간", $"{entry.StartTime}~{entry.EndTime}");
        AddDetailRow(details, "장소", entry.PlaceName);
        AddDetailRow(details, "대표자", entry.Name);
        AddDetailRow(details, "인원수", $"{Math.Max(1, entry.Headcount)}명");
        AddDetailRow(details, "상태", ReservationService.GetUnifiedStatus(entry));
        AddDetailRow(details, "구분", FormatReservationSource(entry.Source));
        AddDetailRow(details, "목적", string.IsNullOrWhiteSpace(entry.Purpose) ? "-" : entry.Purpose);

        var closeButton = new Button
        {
            Text = "닫기",
            Dock = DockStyle.Right,
            Width = 90,
            DialogResult = DialogResult.OK,
            BackColor = Color.FromArgb(70, 66, 64),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        closeButton.FlatAppearance.BorderSize = 0;
        root.Controls.Add(closeButton, 0, 2);
        dialog.AcceptButton = closeButton;
        dialog.ShowDialog(this);
    }

    private static void AddDetailRow(TableLayoutPanel details, string labelText, string value)
    {
        var row = details.RowCount;
        details.RowCount++;
        details.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
        details.Controls.Add(new Label
        {
            Text = labelText,
            Dock = DockStyle.Fill,
            ForeColor = Color.FromArgb(105, 100, 97),
            Font = new Font("맑은 고딕", 9F, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleLeft
        }, 0, row);
        details.Controls.Add(new Label
        {
            Text = value,
            Dock = DockStyle.Fill,
            ForeColor = Color.FromArgb(52, 48, 46),
            TextAlign = ContentAlignment.MiddleLeft,
            AutoEllipsis = true
        }, 1, row);
    }

    private static string FormatReservationSource(string source) => source switch
    {
        "naverSmartPlace" or "naver" => "네이버 예약",
        "reservationRequest" or "webReservation" or "web" or "app" => "웹 예약",
        "usage" => "이용일지",
        _ => string.IsNullOrWhiteSpace(source) ? "-" : source
    };

    private void ConfigureMoreActionsButton()
    {
        _moreActionsButton.Text = "설정";
        ApplyHeaderButtonStyle(_moreActionsButton);

        var menu = new ContextMenuStrip();
        menu.Items.Add("알림 이력", null, (_, _) => ShowEventHistoryDialog());
        menu.Items.Add("실적 다운로드", null, async (_, _) => await DownloadMonthlyPerformanceAsync());
        menu.Items.Add("시트 열기", null, (_, _) => OpenChrome(SheetUrl));
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("연결 설정", null, async (_, _) =>
        {
            if (ShowSettings())
            {
                _pollTimer.Start();
                await RefreshDashboardAsync();
            }
        });
        _moreActionsButton.Click += (_, _) => menu.Show(_moreActionsButton, new Point(0, _moreActionsButton.Height));
    }

    private void ShowEventHistoryDialog()
    {
        using var dialog = new Form
        {
            Text = "알림 이력",
            StartPosition = FormStartPosition.CenterParent,
            MinimumSize = new Size(900, 520),
            Size = new Size(1080, 680),
            Font = Font,
            BackColor = BackColor
        };

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(16)
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        dialog.Controls.Add(root);

        var toolbar = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false
        };
        var categoryFilter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 140 };
        categoryFilter.Items.AddRange(["전체", "방문", "커넥트룸", "커리어존", "이용 제한"]);
        categoryFilter.SelectedIndex = 0;
        var unreadOnly = new CheckBox { Text = "미확인만", AutoSize = true, Padding = new Padding(8, 6, 8, 0) };
        var markAllReadButton = CreateSmallButton("모두 확인", (_, _) =>
        {
            foreach (var item in _state.Events) item.IsRead = true;
            _storage.SaveState(_state);
        });
        toolbar.Controls.Add(categoryFilter);
        toolbar.Controls.Add(unreadOnly);
        toolbar.Controls.Add(markAllReadButton);
        root.Controls.Add(toolbar, 0, 0);

        var grid = new DataGridView();
        ConfigureGrid(grid);
        grid.Columns.Add("read", "상태");
        grid.Columns.Add("category", "분류");
        grid.Columns.Add("time", "등록 시각");
        grid.Columns.Add("name", "이름");
        grid.Columns.Add("detail", "내용");
        grid.Columns["read"]!.Width = 80;
        grid.Columns["category"]!.Width = 110;
        grid.Columns["time"]!.Width = 150;
        grid.Columns["name"]!.Width = 120;
        grid.Columns["detail"]!.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        root.Controls.Add(grid, 0, 1);

        void RenderHistory()
        {
            var selected = categoryFilter.SelectedItem?.ToString() ?? "전체";
            var category = selected switch
            {
                "방문" => "visit",
                "커넥트룸" => "connect",
                "커리어존" => "career",
                "이용 제한" => "restricted",
                _ => ""
            };
            grid.Rows.Clear();
            foreach (var item in _state.Events
                         .Where(item => string.IsNullOrEmpty(category) || item.Category == category)
                         .Where(item => !unreadOnly.Checked || !item.IsRead))
            {
                var detail = item.Category switch
                {
                    "visit" => "방문 등록",
                    "restricted" => item.Purpose,
                    _ => $"{item.StartTime}~{item.EndTime} · {item.Headcount}명 · {item.Purpose}"
                };
                var rowIndex = grid.Rows.Add(
                    item.IsRead ? "확인" : "미확인",
                    item.CategoryLabel,
                    item.OccurredAt,
                    item.Name,
                    detail);
                grid.Rows[rowIndex].Tag = item;
                if (!item.IsRead)
                {
                    grid.Rows[rowIndex].DefaultCellStyle.Font = new Font(grid.Font, FontStyle.Bold);
                    grid.Rows[rowIndex].DefaultCellStyle.BackColor = Color.FromArgb(255, 244, 247);
                }
            }
        }

        categoryFilter.SelectedIndexChanged += (_, _) => RenderHistory();
        unreadOnly.CheckedChanged += (_, _) => RenderHistory();
        markAllReadButton.Click += (_, _) => RenderHistory();
        grid.CellDoubleClick += (_, eventArgs) =>
        {
            if (eventArgs.RowIndex < 0) return;
            if (grid.Rows[eventArgs.RowIndex].Tag is StaffEvent item)
            {
                item.IsRead = true;
                _storage.SaveState(_state);
                RenderHistory();
            }
        };
        RenderHistory();
        dialog.ShowDialog(this);
    }

    private Control BuildStatusBar()
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Padding = new Padding(0, 4, 0, 0)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _statusLabel.Text = "연결 대기";
        _statusLabel.ForeColor = Color.FromArgb(105, 100, 97);
        _statusLabel.AutoSize = false;
        _statusLabel.Dock = DockStyle.Fill;
        _statusLabel.Padding = new Padding(3, 0, 16, 0);
        _statusLabel.TextAlign = ContentAlignment.MiddleLeft;
        _statusLabel.AutoEllipsis = true;
        panel.Controls.Add(_statusLabel, 0, 0);

        _naverBookingStatusLabel.Text = "네이버 예약 확인 대기";
        _naverBookingStatusLabel.ForeColor = Color.FromArgb(105, 100, 97);
        _naverBookingStatusLabel.AutoSize = false;
        _naverBookingStatusLabel.Dock = DockStyle.Fill;
        _naverBookingStatusLabel.Padding = new Padding(16, 0, 3, 0);
        _naverBookingStatusLabel.TextAlign = ContentAlignment.MiddleRight;
        _naverBookingStatusLabel.AutoEllipsis = true;
        panel.Controls.Add(_naverBookingStatusLabel, 1, 0);

        return panel;
    }

    private async Task RefreshDashboardAsync()
    {
        await RefreshDashboardAsync(_visibleCalendarMonth, renderWhenComplete: true);
    }

    private async Task RefreshDashboardAsync(DateTime calendarMonth, bool renderWhenComplete)
    {
        if (_refreshing || _updatingReservationStatus || _demoMode) return;
        _refreshing = true;
        _statusLabel.Text = "운영 정보를 확인하는 중...";
        try
        {
            using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            var requestedMonth = new DateTime(calendarMonth.Year, calendarMonth.Month, 1).ToString("yyyy-MM");
            var eventLimit = _runtimeSyncCompleted ? 100 : 500;
            var snapshot = await _client.GetSnapshotAsync(
                _settings,
                _password,
                requestedMonth,
                eventLimit,
                cancellation.Token);
            MergeEvents(snapshot.Events);
            RenderSummary(snapshot.Summary);
            _checkInsByMonth[requestedMonth] = snapshot.CheckIns ?? [];
            var smartPlaceEntries = _calendarEntries.Where(IsSmartPlaceNaverEntry).ToList();
            var mergedEntries = ReservationService.MergeCalendarEntries(
                snapshot.CalendarEntries.Where(item => !IsSheetNaverEntry(item)),
                ConvertTodayReservations(snapshot.TodayReservations).Where(item => !IsSheetNaverEntry(item)),
                ConvertReservationRequests(snapshot.ReservationRequests).Concat(smartPlaceEntries));
            _calendarCacheByMonth[requestedMonth] = mergedEntries.ToList();
            _calendarCacheUpdatedAt[requestedMonth] = DateTime.Now;

            if (requestedMonth == _visibleCalendarMonth.ToString("yyyy-MM"))
            {
                _calendarEntries = mergedEntries;
                MergeLatestNaverBookingOutput();
                _calendarCacheByMonth[requestedMonth] = _calendarEntries.ToList();
                _calendarCacheUpdatedAt[requestedMonth] = DateTime.Now;
                _loadedCalendarMonth = requestedMonth;
                if (renderWhenComplete) RenderCalendarEntries();
            }
            RenderEvents();
            _runtimeSyncCompleted = true;
            _statusLabel.Text = $"정상 연결 · 마지막 확인 {snapshot.ServerTime}";
        }
        catch (Exception exception)
        {
            _statusLabel.Text = "연결 오류 · " + exception.Message;
        }
        finally
        {
            _refreshing = false;
            if (_pendingCalendarMonthRefresh is { } pendingMonth)
            {
                _pendingCalendarMonthRefresh = null;
                if (ShouldRefreshCalendarMonth(pendingMonth.ToString("yyyy-MM")))
                    QueueCalendarMonthRefresh(pendingMonth);
            }
        }
    }

    private void MergeEvents(IEnumerable<StaffEvent> incoming)
    {
        var originalCount = _state.Events.Count;
        _state.Events = _state.Events
            .Where(item => !string.IsNullOrWhiteSpace(item.Id))
            .GroupBy(item => item.Id, StringComparer.Ordinal)
            .Select(group => group.First())
            .ToList();
        var existingById = _state.Events.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var firstSync = !_state.InitialSyncCompleted;
        var stateChanged = firstSync || originalCount != _state.Events.Count;
        var newAlertEvents = new List<StaffEvent>();

        foreach (var item in incoming
                     .Where(item => !string.IsNullOrWhiteSpace(item.Id))
                     .GroupBy(item => item.Id, StringComparer.Ordinal)
                     .Select(group => group.First()))
        {
            if (existingById.ContainsKey(item.Id)) continue;

            item.IsRead = firstSync;
            _state.Events.Add(item);
            existingById[item.Id] = item;
            stateChanged = true;
            if (!firstSync && item.NotificationEligible) newAlertEvents.Add(item);
        }

        _state.InitialSyncCompleted = true;
        _state.Events = _state.Events
            .OrderByDescending(item => item.OccurredAt, StringComparer.Ordinal)
            .Take(5000)
            .ToList();
        if (stateChanged) _storage.SaveState(_state);

        if (newAlertEvents.Count == 1)
        {
            var item = newAlertEvents[0];
            if (item.Category == "restricted")
            {
                _notifyIcon.ShowBalloonTip(
                    7000,
                    "이용 제한 회원 확인",
                    $"{item.Name}\n{item.Purpose}",
                    ToolTipIcon.Warning);
                return;
            }
            var time = string.IsNullOrWhiteSpace(item.StartTime)
                ? ""
                : $" {item.StartTime}~{item.EndTime}";
            _notifyIcon.ShowBalloonTip(
                5000,
                $"{item.CategoryLabel} 현장예약",
                $"{item.Name} · {item.Headcount}명{time}\n{item.Purpose}",
                ToolTipIcon.Info);
        }
        else if (newAlertEvents.Count > 1)
        {
            var careerCount = newAlertEvents.Count(item => item.Category == "career");
            var connectCount = newAlertEvents.Count(item => item.Category == "connect");
            var restrictedCount = newAlertEvents.Count(item => item.Category == "restricted");
            var parts = new List<string>();
            if (careerCount > 0) parts.Add($"커리어존 {careerCount}건");
            if (connectCount > 0) parts.Add($"커넥트룸 {connectCount}건");
            if (restrictedCount > 0) parts.Add($"이용 제한 {restrictedCount}건");
            _notifyIcon.ShowBalloonTip(
                6000,
                $"새 알림 {newAlertEvents.Count}건",
                string.Join(" · ", parts) + "\n프로그램에서 내용을 확인해 주세요.",
                restrictedCount > 0 ? ToolTipIcon.Warning : ToolTipIcon.Info);
        }
    }

    private void RenderEvents()
    {
        if (_eventsGrid.Columns.Count == 0) return;
        var selected = _categoryFilter.SelectedItem?.ToString() ?? "전체";
        var category = selected switch
        {
            "방문" => "visit",
            "커넥트룸" => "connect",
            "커리어존" => "career",
            "이용 제한" => "restricted",
            _ => ""
        };
        var filtered = _state.Events
            .Where(item => string.IsNullOrEmpty(category) || item.Category == category)
            .Where(item => !_unreadOnly.Checked || !item.IsRead)
            .ToList();

        _eventsGrid.Rows.Clear();
        foreach (var item in filtered)
        {
            var detail = item.Category switch
            {
                "visit" => "방문 등록",
                "restricted" => item.Purpose,
                _ => $"{item.StartTime}~{item.EndTime} · {item.Headcount}명 · {item.Purpose}"
            };
            var rowIndex = _eventsGrid.Rows.Add(
                item.IsRead ? "확인" : "미확인",
                item.CategoryLabel,
                item.OccurredAt,
                item.Name,
                detail);
            var row = _eventsGrid.Rows[rowIndex];
            row.Tag = item;
            if (!item.IsRead)
            {
                row.DefaultCellStyle.Font = new Font(_eventsGrid.Font, FontStyle.Bold);
                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 244, 247);
            }
        }

        var unreadCount = _state.Events.Count(item => !item.IsRead);
        _unreadLabel.Text = unreadCount > 0 ? $"미확인 {unreadCount}" : "모두 확인";
    }

    private void RenderCalendarEntries()
    {
        ReservationService.ApplyCheckIns(
            _calendarEntries,
            _checkInsByMonth.TryGetValue(_visibleCalendarMonth.ToString("yyyy-MM"), out var checkIns) ? checkIns : []);
        UpdateMissingUsageLogButton();
        CheckOverdueUsageLogs();
        var selectedDate = _selectedCalendarDate.ToString("yyyy-MM-dd");
        _reservationsTitle.Text = selectedDate == DateTime.Today.ToString("yyyy-MM-dd")
            ? "오늘 운영 목록"
            : $"{_selectedCalendarDate:M월 d일} 운영 목록";
        RenderMonthGrid();

        var dayItems = _calendarEntries.Where(item => ReservationService.IsSameCalendarDate(item.Date, _selectedCalendarDate)).ToList();
        var statusCounts = dayItems
            .Select(ReservationService.GetUnifiedStatus)
            .GroupBy(status => status)
            .ToDictionary(group => group.Key, group => group.Count());
        _reservationSummaryLabel.Text = dayItems.Count == 0
            ? "선택한 날짜의 예약이 없습니다."
            : string.Join(" · ", new[] { "승인 대기", "승인 완료", "이용중", "이용완료", ReservationService.MissingUsageLogStatus, "노쇼", "취소" }
                .Where(status => statusCounts.ContainsKey(status))
                .Select(status => $"{status} {statusCounts[status]}건"));

        var selectedItems = ApplyReservationFilters(dayItems).ToList();
        _reservationFilterSummaryLabel.Text = dayItems.Count == selectedItems.Count
            ? $"전체 {dayItems.Count}건 표시 중"
            : $"전체 {dayItems.Count}건 중 {selectedItems.Count}건 표시";

        _reservationsGrid.Rows.Clear();
        if (dayItems.Count > 0 && selectedItems.Count == 0)
        {
            _reservationSummaryLabel.Text += " · 현재 필터로 숨겨진 예약이 있습니다.";
            var rowIndex = _reservationsGrid.Rows.Add(
                "-",
                "필터를 전체로 변경해 주세요.",
                "현재 필터로 숨겨진 예약이 있습니다.",
                "-",
                "숨김");
            var row = _reservationsGrid.Rows[rowIndex];
            row.DefaultCellStyle.ForeColor = Color.FromArgb(120, 115, 112);
            row.DefaultCellStyle.BackColor = Color.FromArgb(250, 250, 249);
        }

        foreach (var item in selectedItems)
        {
            var state = ReservationService.GetUnifiedStatus(item);
            var rowIndex = _reservationsGrid.Rows.Add(
                $"{item.StartTime}~{item.EndTime}",
                item.PlaceName,
                item.Name,
                $"{Math.Max(1, item.Headcount)}명",
                state);
            _reservationsGrid.Rows[rowIndex].Tag = item;
            ApplyReservationRowStyle(_reservationsGrid.Rows[rowIndex], state);
        }
    }

    private IEnumerable<CalendarEntry> ApplyReservationFilters(IEnumerable<CalendarEntry> entries)
    {
        var statusFilter = _reservationStatusFilter.SelectedItem?.ToString() ?? "전체 상태";
        var spaceFilter = _reservationSpaceFilter.SelectedItem?.ToString() ?? "전체 공간";

        return entries
            .Where(item => statusFilter == "전체 상태" || ReservationService.GetUnifiedStatus(item) == statusFilter)
            .Where(item => spaceFilter == "전체 공간" || ReservationService.IsPlaceMatch(item, spaceFilter));
    }

    private void UpdateReservationActionButtons()
    {
        var selectedEntry = _reservationsGrid.SelectedRows.Count > 0
            ? _reservationsGrid.SelectedRows[0].Tag as CalendarEntry
            : null;
        var canChangeRequestStatus = selectedEntry != null &&
            CanChangeEntryStatus(selectedEntry) &&
            !_updatingReservationStatus;

        foreach (var button in new[]
                 {
                     (_requestApproveButton, "승인"),
                     (_requestRejectButton, "반려")
                 })
        {
            button.Item1.Enabled = canChangeRequestStatus &&
                selectedEntry != null &&
                IsAllowedStatusForEntry(selectedEntry, button.Item2) &&
                !StatusMatches(selectedEntry, button.Item2);
        }

        _proxyUsageLogButton.Enabled = canChangeRequestStatus &&
            selectedEntry != null &&
            !selectedEntry.Submitted &&
            !selectedEntry.Cancelled &&
            IsAllowedStatusForEntry(selectedEntry, "이용완료");
    }

    private static List<CalendarEntry> ConvertTodayReservations(IEnumerable<Reservation> reservations) =>
        reservations.Select(item => new CalendarEntry
        {
            Id = item.Id,
            Source = string.IsNullOrWhiteSpace(item.Source) ? "naver" : item.Source,
            MemberId = item.MemberId,
            Date = string.IsNullOrWhiteSpace(item.Date) ? DateTime.Today.ToString("yyyy-MM-dd") : item.Date,
            StartTime = item.StartTime,
            EndTime = item.EndTime,
            PlaceName = item.PlaceName,
            SpaceType = item.SpaceType,
            Name = item.BookerName,
            Purpose = item.Purpose,
            Headcount = item.Headcount,
            Status = item.Status,
            Cancelled = item.Cancelled,
            Submitted = item.Submitted
        }).ToList();

    private static List<CalendarEntry> ConvertReservationRequests(IEnumerable<ReservationRequest> requests) =>
        requests.Select(item => new CalendarEntry
        {
            Id = !string.IsNullOrWhiteSpace(item.Id) ? item.Id : item.RequestId,
            Source = "reservationRequest",
            MemberId = item.MemberId,
            Date = item.Date,
            StartTime = item.StartTime,
            EndTime = item.EndTime,
            PlaceName = FirstNonEmpty(item.PlaceName, item.SpaceLabel, FormatSpaceLabel(item.SpaceType, item.Room)),
            Name = FirstNonEmpty(item.BookerName, item.Name, "예약자"),
            Purpose = item.Purpose,
            Headcount = item.Headcount,
            Status = string.IsNullOrWhiteSpace(item.Status) ? "신청" : item.Status,
            Submitted = item.Submitted
        })
        .Where(item => !string.IsNullOrWhiteSpace(item.Id) && !string.IsNullOrWhiteSpace(item.Date))
        .ToList();

    private async Task UpdateSelectedReservationStatusAsync(string status)
    {
        if (_updatingReservationStatus ||
            _reservationsGrid.SelectedRows.Count == 0 ||
            _reservationsGrid.SelectedRows[0].Tag is not CalendarEntry entry ||
            !CanChangeEntryStatus(entry))
        {
            return;
        }

        if (StatusMatches(entry, status))
        {
            _statusLabel.Text = $"예약 {entry.Id}은(는) 이미 {status} 상태입니다.";
            return;
        }

        if (!IsAllowedStatusForEntry(entry, status)) return;

        if (status is "취소" or "노쇼" or "반려")
        {
            var confirmed = MessageBox.Show(
                this,
                $"{entry.PlaceName} · {entry.Name}\n{entry.Date} {entry.StartTime}~{entry.EndTime}\n\n" +
                $"예약 상태를 '{status}'로 변경할까요?",
                "예약 상태 변경",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);
            if (confirmed != DialogResult.Yes) return;
        }

        _updatingReservationStatus = true;
        var updated = false;
        _statusLabel.Text = $"예약 {entry.Id} 상태를 {status}(으)로 변경하는 중...";
        try
        {
            using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            var result = status == "이용완료"
                ? await CompleteReservationUsageAsync(entry, cancellation.Token)
                : IsReservationRequestSource(entry.Source)
                ? await _client.UpdateReservationRequestStatusAsync(
                    _settings,
                    _password,
                    StripKnownEntryPrefix(entry.Id),
                    status,
                    cancellation.Token)
                : await _client.UpdateReservationStatusAsync(
                    _settings,
                    _password,
                    StripKnownEntryPrefix(entry.Id),
                    status,
                    cancellation.Token);
            entry.Status = result.Status;
            entry.Cancelled = result.Cancelled;
            entry.Submitted = entry.Submitted || result.UsageLogged || status == "이용완료";
            RenderCalendarEntries();
            updated = true;
            _statusLabel.Text = $"예약 {result.ReservationId} 상태를 {result.Status}(으)로 변경했습니다.";
        }
        catch (Exception exception)
        {
            _statusLabel.Text = "상태 변경 오류 · " + exception.Message;
            MessageBox.Show(
                this,
                exception.Message,
                "예약 상태 변경 실패",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            _updatingReservationStatus = false;
        }

        if (updated) await RefreshDashboardAsync();
    }

    private async Task<ReservationStatusUpdate> CompleteReservationUsageAsync(
        CalendarEntry entry,
        CancellationToken cancellationToken)
    {
        var memberId = entry.MemberId;
        if (string.IsNullOrWhiteSpace(memberId))
        {
            var members = await _client.LookupMembersAsync(_settings, _password, entry.Name, cancellationToken);
            if (members.Count == 0)
                throw new InvalidOperationException("예약자 이름과 일치하는 회원을 찾을 수 없습니다.");
            var member = members.Count == 1 ? members[0] : ShowMemberSelectionDialog(members, entry.Name);
            if (member is null) throw new InvalidOperationException("회원 선택을 취소했습니다.");
            memberId = member.Id;
        }

        var companions = new List<ManualCompanion>();
        var companionCount = Math.Max(0, entry.Headcount - 1);
        if (companionCount > 0 && !IsReservationRequestSource(entry.Source))
        {
            companions = ShowCompanionDialog(entry, companionCount)
                ?? throw new InvalidOperationException("동반 이용자 입력을 취소했습니다.");
        }

        return await _client.CompleteReservationUsageAsync(
            _settings,
            _password,
            entry,
            memberId,
            companions,
            cancellationToken);
    }

    private List<ManualCompanion>? ShowCompanionDialog(CalendarEntry entry, int companionCount)
    {
        using var dialog = new Form
        {
            Text = "동반 이용자 정보",
            StartPosition = FormStartPosition.CenterParent,
            Size = new Size(420, Math.Min(580, 214 + companionCount * 40)),
            MinimizeBox = false,
            MaximizeBox = false,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            Font = Font,
            BackColor = Color.White
        };

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(16)
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        root.Controls.Add(new Label
        {
            AutoSize = true,
            MaximumSize = new Size(370, 0),
            Margin = new Padding(0, 0, 0, 10),
            Text = $"{entry.Name} · {entry.PlaceName} {entry.StartTime}~{entry.EndTime}\n" +
                   $"예약 인원 {entry.Headcount}명 · 동반 이용자 {companionCount}명의 성별과 만 나이를 입력해 주세요."
        }, 0, 0);

        var rows = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            AutoScroll = true
        };
        rows.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
        rows.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
        rows.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
        var inputs = new List<(ComboBox Gender, NumericUpDown Age)>();
        for (var index = 0; index < companionCount; index++)
        {
            var gender = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 100 };
            gender.Items.AddRange(["남성", "여성"]);
            var age = new NumericUpDown { Minimum = 1, Maximum = 100, Value = 25, Width = 100 };
            rows.Controls.Add(new Label { Text = $"동반 {index + 1}", AutoSize = true, Margin = new Padding(0, 7, 0, 0) }, 0, index);
            rows.Controls.Add(gender, 1, index);
            rows.Controls.Add(age, 2, index);
            inputs.Add((gender, age));
        }
        root.Controls.Add(rows, 0, 1);

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            FlowDirection = FlowDirection.RightToLeft
        };
        // 기본 크기(약 23px)는 누르기 어려워 다른 창과 같은 40px 버튼을 쓴다.
        var ok = CreateCompactDialogButton("제출", MissingUsageLogColor, DialogResult.OK);
        var cancel = CreateCompactDialogButton("취소", Color.FromArgb(117, 111, 108), DialogResult.Cancel);
        buttons.Controls.Add(ok);
        buttons.Controls.Add(cancel);
        root.Controls.Add(buttons, 0, 2);
        dialog.Controls.Add(root);
        dialog.AcceptButton = ok;
        dialog.CancelButton = cancel;

        while (dialog.ShowDialog(this) == DialogResult.OK)
        {
            if (inputs.All(input => input.Gender.SelectedItem is string))
            {
                return inputs
                    .Select(input => new ManualCompanion
                    {
                        Gender = (string)input.Gender.SelectedItem!,
                        Age = (int)input.Age.Value
                    })
                    .ToList();
            }

            MessageBox.Show(
                dialog,
                "동반 이용자의 성별을 모두 선택해 주세요.",
                "동반 이용자 정보",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        return null;
    }

    private AdminMember? ShowMemberSelectionDialog(List<AdminMember> members, string reservationName)
    {
        using var dialog = new Form
        {
            Text = "예약자 회원 선택",
            StartPosition = FormStartPosition.CenterParent,
            Size = new Size(540, 360),
            MinimizeBox = false,
            MaximizeBox = false,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            Font = Font,
            BackColor = Color.White
        };

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(18)
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
        dialog.Controls.Add(root);

        root.Controls.Add(new Label
        {
            Text = $"{reservationName} 예약과 연결할 회원을 선택해 주세요.",
            Dock = DockStyle.Fill,
            ForeColor = Color.FromArgb(52, 48, 46),
            Font = new Font("맑은 고딕", 10F, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleLeft
        }, 0, 0);

        var list = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false };
        foreach (var member in members.OrderBy(item => item.Name).ThenBy(item => item.Phone))
        {
            list.Items.Add(new MemberSelectionItem(member));
        }
        if (list.Items.Count > 0) list.SelectedIndex = 0;
        root.Controls.Add(list, 0, 1);

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false
        };
        var selectButton = CreateCompactDialogButton("선택", Color.FromArgb(74, 156, 68), DialogResult.OK);
        var cancelButton = CreateCompactDialogButton("취소", Color.FromArgb(117, 111, 108), DialogResult.Cancel);
        buttons.Controls.Add(selectButton);
        buttons.Controls.Add(cancelButton);
        root.Controls.Add(buttons, 0, 2);

        dialog.AcceptButton = selectButton;
        dialog.CancelButton = cancelButton;
        list.DoubleClick += (_, _) => dialog.DialogResult = DialogResult.OK;

        return dialog.ShowDialog(this) == DialogResult.OK &&
            list.SelectedItem is MemberSelectionItem selected
                ? selected.Member
                : null;
    }

    private static Button CreateCompactDialogButton(string text, Color color, DialogResult result)
    {
        var button = new Button
        {
            Text = text,
            Width = 88,
            Height = 40,
            Margin = new Padding(8, 4, 0, 0),
            DialogResult = result,
            BackColor = color,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        button.FlatAppearance.BorderSize = 0;
        return button;
    }

    private sealed class MemberSelectionItem(AdminMember member)
    {
        public AdminMember Member { get; } = member;

        public override string ToString() =>
            $"{Member.Name} · {Member.Gender} · 만 {Member.Age}세 · {FormatPhone(Member.Phone)}";
    }

    private static bool CanChangeEntryStatus(CalendarEntry entry) =>
        IsReservationRequestSource(entry.Source) || IsSheetNaverEntry(entry) || IsSmartPlaceNaverEntry(entry);

    private static bool IsAllowedStatusForEntry(CalendarEntry entry, string status) =>
        CanChangeEntryStatus(entry) &&
        ReservationStatusActionsForEntry(entry).Any(action => action.Status == status);

    private static bool StatusMatches(CalendarEntry entry, string status)
    {
        if (string.Equals(entry.Status, status, StringComparison.Ordinal)) return true;
        return status switch
        {
            "승인" => entry.Status is "승인 완료" or "승인",
            "확정" => entry.Status is "승인 완료" or "확정",
            _ => false
        };
    }

    private static (string Label, string Status)[] ReservationStatusActionsForEntry(CalendarEntry entry)
    {
        if (IsReservationRequestSource(entry.Source))
        {
            return
            [
                ("승인 대기로 변경", "신청"),
                ("승인 완료로 변경", "승인"),
                ("반려로 변경", "반려"),
                ("취소로 변경", "취소"),
                ("이용완료로 변경 (일지 대리 제출)", "이용완료")
            ];
        }

        if (IsSheetNaverEntry(entry) || IsSmartPlaceNaverEntry(entry))
        {
            return
            [
                ("승인 완료로 변경", "확정"),
                ("이용완료로 변경 (일지 대리 제출)", "이용완료"),
                ("노쇼로 변경", "노쇼"),
                ("취소로 변경", "취소")
            ];
        }

        return [];
    }

    private static bool IsReservationRequestSource(string source) =>
        source is "reservationRequest" or "webReservation" or "web" or "app";

    private static bool IsSheetNaverEntry(CalendarEntry entry) =>
        string.Equals(entry.Source, "naver", StringComparison.Ordinal);

    private static bool IsSmartPlaceNaverEntry(CalendarEntry entry) =>
        string.Equals(entry.Source, "naverSmartPlace", StringComparison.Ordinal);

    private void ApplyReservationRowStyle(DataGridViewRow row, string state)
    {
        row.DefaultCellStyle.ForeColor = state switch
        {
            "승인 대기" => Color.FromArgb(154, 91, 25),
            "이용중" => Color.FromArgb(36, 88, 122),
            "이용완료" => Color.FromArgb(48, 120, 56),
            ReservationService.MissingUsageLogStatus => Color.FromArgb(196, 84, 20),
            "노쇼" => Color.FromArgb(168, 48, 48),
            "취소" => Color.FromArgb(120, 115, 112),
            _ => Color.FromArgb(52, 48, 46)
        };
        row.Cells["status"].Style.BackColor = state switch
        {
            "승인 대기" => Color.FromArgb(255, 247, 235),
            "승인 완료" => Color.FromArgb(246, 246, 244),
            "이용중" => Color.FromArgb(236, 244, 249),
            "이용완료" => Color.FromArgb(238, 247, 237),
            ReservationService.MissingUsageLogStatus => Color.FromArgb(255, 241, 230),
            "노쇼" => Color.FromArgb(255, 239, 239),
            "취소" => Color.FromArgb(241, 240, 238),
            _ => Color.White
        };
        row.Cells["status"].Style.ForeColor = row.DefaultCellStyle.ForeColor;
        row.Cells["status"].Style.Font = new Font(_reservationsGrid.Font, FontStyle.Bold);
        row.Cells["status"].Style.SelectionBackColor = row.Cells["status"].Style.BackColor;
        row.Cells["status"].Style.SelectionForeColor = row.Cells["status"].Style.ForeColor;
        if (state is "승인 대기" or "이용중" or "노쇼")
            row.DefaultCellStyle.Font = new Font(_reservationsGrid.Font, FontStyle.Bold);
    }

    private static string StripKnownEntryPrefix(string id)
    {
        foreach (var prefix in new[] { "naver:", "naverSmartPlace:", "reservationRequest:", "webReservation:", "web:", "app:" })
        {
            if (id.StartsWith(prefix, StringComparison.Ordinal)) return id[prefix.Length..];
        }

        return id;
    }

    private static string FirstNonEmpty(params string[] values) =>
        values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";

    private static string FormatPhone(string value)
    {
        var digits = new string(value.Where(char.IsDigit).ToArray());
        return digits.Length == 11
            ? $"{digits[..3]}-{digits.Substring(3, 4)}-{digits[7..]}"
            : digits;
    }

    private static string FormatSpaceLabel(string spaceType, string room)
    {
        if (spaceType == "connectRoom") return "커넥트룸";
        if (room == "careerRoom1" || spaceType == "careerRoom1") return "커리어존 Room1";
        if (room == "careerRoom2" || spaceType == "careerRoom2") return "커리어존 Room2";
        return string.IsNullOrWhiteSpace(spaceType) ? "공간 예약" : spaceType;
    }

    private void RenderSummary(DashboardSummary summary)
    {
        _todayVisitMetricLabel.Text = $"오늘 방문 {summary.VisitorStats.Daily:N0}";
        _weeklyVisitMetricLabel.Text = $"주간 {summary.VisitorStats.Weekly:N0}";
        _monthlyVisitMetricLabel.Text = $"월간 {summary.VisitorStats.Monthly:N0}";
        _memberMetricLabel.Text = $"회원 {summary.ActiveMemberCount:N0}";
    }

    private void MarkAllRead()
    {
        foreach (var item in _state.Events) item.IsRead = true;
        _storage.SaveState(_state);
        RenderEvents();
    }

    private bool ShowSettings()
    {
        using var dialog = new SettingsForm(
            _settings,
            _password,
            _naverId,
            _naverPassword,
            LaunchNaverLoginBrowserAsync,
            () => _storage.ReadDiagnosticLog());
        if (dialog.ShowDialog(this) != DialogResult.OK) return false;
        _password = dialog.Password;
        _naverId = dialog.NaverId;
        _naverPassword = dialog.NaverPassword;
        _settings.AppsScriptUrl = dialog.AppsScriptUrl;
        _settings.PollSeconds = dialog.PollSeconds;
        _settings.ProtectedPassword = _storage.ProtectPassword(_password);
        _settings.ProtectedNaverId = _storage.ProtectPassword(_naverId);
        _settings.ProtectedNaverPassword = _storage.ProtectPassword(_naverPassword);
        _storage.SaveSettings(_settings);
        _client.ResetSession();
        _pollTimer.Interval = _settings.PollSeconds * 1000;
        return true;
    }

    private async Task DownloadMonthlyPerformanceAsync()
    {
        using var monthDialog = new MonthSelectionForm();
        if (monthDialog.ShowDialog(this) != DialogResult.OK) return;

        using var saveDialog = new SaveFileDialog
        {
            Title = "실적 저장",
            Filter = "Excel 통합 문서 (*.xlsx)|*.xlsx",
            DefaultExt = "xlsx",
            AddExtension = true,
            FileName = monthDialog.IsMonthMode
                ? $"청년이봄_{monthDialog.SelectedMonth}_월간실적.xlsx"
                : $"청년이봄_{monthDialog.StartDate}_{monthDialog.EndDate}_실적.xlsx"
        };
        if (saveDialog.ShowDialog(this) != DialogResult.OK) return;

        var periodName = monthDialog.PeriodName;
        _pollTimer.Stop();
        _statusLabel.Text = $"{periodName} 실적을 정리하고 있습니다...";
        Cursor = Cursors.WaitCursor;
        try
        {
            // 긴 기간은 서버 집계가 오래 걸려 시간을 넉넉히 둔다.
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(monthDialog.IsMonthMode ? 90 : 240));
            var report = monthDialog.IsMonthMode
                ? await _client.GetMonthlyPerformanceAsync(_settings, _password, monthDialog.SelectedMonth, timeout.Token)
                : await _client.GetPerformanceForRangeAsync(
                    _settings, _password, monthDialog.StartDate, monthDialog.EndDate, timeout.Token);
            PerformanceWorkbookWriter.Write(saveDialog.FileName, report);
            _statusLabel.Text = $"{periodName} 실적 저장 완료";
            MessageBox.Show(this, "실적 파일을 저장했습니다.\n\n" + saveDialog.FileName,
                "실적 다운로드 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (OperationCanceledException)
        {
            _statusLabel.Text = "실적 다운로드 시간 초과";
            MessageBox.Show(this, "실적 집계 시간이 초과되었습니다. 잠시 후 다시 시도해 주세요.",
                "실적 다운로드 실패", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        catch (Exception exception)
        {
            _statusLabel.Text = "실적 다운로드 오류 · " + exception.Message;
            MessageBox.Show(this, exception.Message, "실적 다운로드 실패",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            Cursor = Cursors.Default;
            _pollTimer.Start();
        }
    }

    private async Task RefreshAllAsync()
    {
        await RefreshDashboardAsync();
        await RefreshNaverBookingStatusAsync();
    }

    private async Task RefreshNaverBookingStatusAsync()
    {
        if (_naverBookingSyncRunning || _demoMode) return;
        var repoRoot = FindRepositoryRoot();
        if (string.IsNullOrWhiteSpace(repoRoot))
        {
            _naverBookingStatusLabel.Text = "네이버 확인 불가";
            return;
        }

        _naverBookingSyncRunning = true;
        _naverBookingStatusLabel.Text = "네이버 확인 중...";
        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(90));
            var result = await RunNaverBookingSyncAsync(repoRoot, _naverId, _naverPassword, timeout.Token);
            LogNaverRun("네이버 예약 확인", result.ExitCode, result.OutputText, result.ErrorText);
            if (result.ExitCode != 0)
            {
                var message = FirstMeaningfulLine(result.ErrorText, result.OutputText);
                if ((result.ErrorText + result.OutputText).Contains("NAVER_LOGIN_REQUIRED", StringComparison.OrdinalIgnoreCase))
                {
                    _naverBookingStatusLabel.Text = "네이버 로그인 필요";
                    return;
                }
                _naverBookingStatusLabel.Text = "네이버 확인 실패" +
                    (string.IsNullOrWhiteSpace(message) ? "" : " · " + message);
                return;
            }

            var outputPath = Path.Combine(repoRoot, "outputs", "naver-booking-sync.json");
            var naverSnapshot = NaverBookingParser.ReadSnapshot(outputPath);
            _lastNaverBookingSnapshot = naverSnapshot;
            _naverBookingCheckedAt = DateTime.Now;
            var displayedTodayUseCount = MergeNaverTodayUseReservations(naverSnapshot);
            var sheetChanges = await SyncNaverReservationsToSheetAsync(naverSnapshot);
            var sheetText = sheetChanges > 0 ? $" · 시트 반영 {sheetChanges}" : "";
            var pending = naverSnapshot.PendingConfirm;
            UpdateNaverPendingButton(pending.Count);
            var checkedAt = DateTime.Now.ToString("HH:mm");
            if (pending.Count == 0)
            {
                var todayUseText = naverSnapshot.TodayUse.Count > 0
                    ? $" · 오늘이용 {displayedTodayUseCount}"
                    : "";
                _naverBookingStatusLabel.Text = $"네이버 신청 없음{todayUseText}{sheetText} · {checkedAt}";
                _lastNaverPendingSignature = "";
                return;
            }

            _naverBookingStatusLabel.Text =
                $"네이버 신청 {pending.Count} · 오늘이용 {displayedTodayUseCount}{sheetText} · {checkedAt}";
            var signature = string.Join("|", pending.Select(item => item.ReservationNo).Order(StringComparer.Ordinal));
            if (!string.Equals(signature, _lastNaverPendingSignature, StringComparison.Ordinal))
            {
                _lastNaverPendingSignature = signature;
                _notifyIcon.ShowBalloonTip(
                    7000,
                    $"네이버 예약 신청 {pending.Count}건",
                    NaverBookingParser.BuildPendingReservationNotice(pending),
                    ToolTipIcon.Info);
                if (!string.Equals(signature, _lastNaverDialogSignature, StringComparison.Ordinal))
                {
                    _lastNaverDialogSignature = signature;
                    BeginInvoke(new Action(ShowNaverBookingDialog));
                }
            }
        }
        catch (OperationCanceledException)
        {
            _naverBookingStatusLabel.Text = "네이버 예약 확인 시간 초과";
            LogNaverRun("네이버 예약 확인", -1, "", "90초 안에 끝나지 않아 중단했습니다.");
        }
        catch (Exception exception)
        {
            _naverBookingStatusLabel.Text = "네이버 예약 확인 오류 · " + exception.Message;
            LogNaverRun("네이버 예약 확인", -1, "", exception.ToString());
        }
        finally
        {
            _naverBookingSyncRunning = false;
        }
    }

    // 이미 서버에 맞춘 '예약번호|상태'. 같은 예약을 5분마다 다시 보내지 않게 한다.
    private readonly HashSet<string> _naverSheetSyncedKeys = new(StringComparer.Ordinal);

    // 오늘 확정 예약은 시트에 없으면 추가하고, 오늘 취소된 예약은 시트에서 '취소'로 바꾼다.
    // 반환값: 실제로 시트가 바뀐 건수
    private async Task<int> SyncNaverReservationsToSheetAsync(NaverBookingSnapshot snapshot)
    {
        if (_demoMode || string.IsNullOrWhiteSpace(_password)) return 0;
        var changed = 0;
        foreach (var item in NaverBookingParser.BuildSheetSyncItems(snapshot))
        {
            if (_naverSheetSyncedKeys.Contains(item.Key)) continue;
            try
            {
                using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                var result = await _client.EnsureNaverReservationAsync(_settings, _password, item, timeout.Token);
                _naverSheetSyncedKeys.Add(item.Key);
                if (!result.Created && !result.Updated) continue;
                changed++;
                var action = result.Created
                    ? "추가"
                    : item.Status == "취소"
                        ? "취소로 변경"
                        : "변경 반영(" + string.Join(", ", result.ChangedFields.Select(DescribeNaverSheetField)) + ")";
                _storage.AppendDiagnosticLog(
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 네이버 예약 시트 반영 · {item.Reservation.ReservationNo} {action}" +
                    $" → {item.Entry.Date} {item.Entry.StartTime}~{item.Entry.EndTime} {item.Entry.PlaceName}");
            }
            catch (Exception exception)
            {
                // 다음 확인 때 다시 시도한다.
                LogNaverRun($"네이버 예약 시트 반영 · {item.Reservation.ReservationNo}", -1, "", exception.Message);
            }
        }
        if (changed > 0) _ = RefreshDashboardAsync();
        return changed;
    }

    private static string DescribeNaverSheetField(string field) => field switch
    {
        "usageDate" => "날짜",
        "startTime" => "시작",
        "endTime" => "종료",
        "spaceType" or "placeName" => "공간",
        "headcount" => "인원",
        "purpose" => "목적",
        "bookerName" => "예약자",
        "maskedPhone" => "연락처",
        "status" or "confirmedAt" => "상태",
        _ => field
    };

    // 성공은 한 줄로, 실패는 출력 전체를 남긴다.
    private void LogNaverRun(string title, int exitCode, string outputText, string errorText)
    {
        var header = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {title} · " + (exitCode == 0 ? "성공" : $"실패 (종료 코드 {exitCode})");
        if (exitCode == 0)
        {
            _storage.AppendDiagnosticLog(header);
            return;
        }
        _storage.AppendDiagnosticLog(
            header + "\r\n" +
            "[출력]\r\n" + (string.IsNullOrWhiteSpace(outputText) ? "(없음)" : outputText.Trim()) + "\r\n" +
            "[오류]\r\n" + (string.IsNullOrWhiteSpace(errorText) ? "(없음)" : errorText.Trim()));
    }

    private static string GetNaverAutomationProfileDir(string repoRoot) =>
        Path.Combine(repoRoot, "automation", ".naver-booking-state", "chrome-profile");

    // 수집 스크립트가 쓰는 크롬 프로필로 실제 크롬을 열어, 직원이 직접 네이버에 로그인해 두게 한다.
    private Task<string?> LaunchNaverLoginBrowserAsync()
    {
        if (_naverBookingSyncRunning)
            return Task.FromResult<string?>("지금 네이버 예약을 확인하는 중입니다. 잠시 뒤 다시 눌러 주세요.");

        var repoRoot = FindRepositoryRoot();
        if (string.IsNullOrWhiteSpace(repoRoot))
            return Task.FromResult<string?>("프로그램 폴더(SPACE_MANAGEMENT_REPO)를 찾지 못했습니다.");
        if (!File.Exists(ChromePath))
            return Task.FromResult<string?>("크롬이 설치된 위치를 찾지 못했습니다: " + ChromePath);

        try
        {
            var profileDir = GetNaverAutomationProfileDir(repoRoot);
            Directory.CreateDirectory(profileDir);
            Process.Start(new ProcessStartInfo
            {
                FileName = ChromePath,
                Arguments = $"--new-window --user-data-dir=\"{profileDir}\" \"{SmartPlaceUrl}\"",
                UseShellExecute = true
            });
            _naverBookingStatusLabel.Text = "네이버 로그인 창을 열었습니다";
            _storage.AppendDiagnosticLog(
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 네이버 로그인 설정 · 로그인용 크롬 창 열림\r\n프로필: {profileDir}");
            return Task.FromResult<string?>(null);
        }
        catch (Exception exception)
        {
            _naverBookingStatusLabel.Text = "네이버 로그인 창을 열지 못했습니다";
            LogNaverRun("네이버 로그인 설정", -1, "", exception.ToString());
            return Task.FromResult<string?>(exception.Message);
        }
    }

    private void UpdateNaverPendingButton(int pendingCount)
    {
        if (_naverPendingButton.IsDisposed) return;
        _naverPendingButton.Text = pendingCount > 0 ? $"예약신청 {pendingCount}" : "예약신청";
        ApplyHeaderButtonStyle(_naverPendingButton, pendingCount > 0 ? Color.FromArgb(214, 108, 31) : null);
    }

    private void ShowNaverBookingDialog()
    {
        var snapshot = _lastNaverBookingSnapshot;
        if (snapshot.PendingConfirm.Count == 0 && snapshot.TodayUse.Count == 0)
        {
            var repoRoot = FindRepositoryRoot();
            if (!string.IsNullOrWhiteSpace(repoRoot))
            {
                var outputPath = Path.Combine(repoRoot, "outputs", "naver-booking-sync.json");
                snapshot = NaverBookingParser.ReadSnapshot(outputPath);
                _lastNaverBookingSnapshot = snapshot;
            }
        }

        using var dialog = BuildNaverBookingDialog(snapshot, _naverBookingCheckedAt);
        dialog.ShowDialog(this);
    }

    private DateTime? _naverBookingCheckedAt;
    private static readonly System.Globalization.CultureInfo KoreanCulture = new("ko-KR");

    private Form BuildNaverBookingDialog(NaverBookingSnapshot snapshot, DateTime? checkedAt)
    {
        var pending = snapshot.PendingConfirm
            .Select(item => (Item: item, Entry: NaverBookingParser.ConvertTodayUseReservation(item)))
            .OrderBy(row => row.Entry.Date, StringComparer.Ordinal)
            .ThenBy(row => row.Entry.StartTime, StringComparer.Ordinal)
            .ToList();
        var todayCancelCount = snapshot.TodayCancel.Count;

        var dialog = new Form
        {
            Text = "네이버 예약 신청",
            StartPosition = FormStartPosition.CenterParent,
            Size = new Size(860, 540),
            MinimumSize = new Size(720, 420),
            Font = Font,
            BackColor = Color.White,
            MinimizeBox = false,
            ShowInTaskbar = false
        };

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(22, 18, 22, 16),
            BackColor = Color.White
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
        dialog.Controls.Add(root);

        // 머리말: 건수를 크게, 할 일과 마지막 확인 시각을 작게
        var header = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Margin = new Padding(0) };
        header.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
        header.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        header.Controls.Add(new Label
        {
            Text = pending.Count > 0 ? $"승인 대기 {pending.Count}건" : "승인 대기 없음",
            Dock = DockStyle.Fill,
            Font = new Font("맑은 고딕", 15F, FontStyle.Bold),
            ForeColor = pending.Count > 0 ? Color.FromArgb(178, 86, 18) : Color.FromArgb(52, 48, 46),
            TextAlign = ContentAlignment.MiddleLeft,
            Margin = new Padding(0)
        }, 0, 0);
        var subtitleParts = new List<string>();
        if (pending.Count > 0) subtitleParts.Add("승인·거절은 네이버 스마트플레이스에서 처리해 주세요");
        if (todayCancelCount > 0) subtitleParts.Add($"오늘 취소 {todayCancelCount}건");
        subtitleParts.Add(checkedAt is { } at ? $"마지막 확인 {at:HH:mm}" : "아직 네이버를 확인하지 않았습니다");
        header.Controls.Add(new Label
        {
            Text = string.Join(" · ", subtitleParts),
            Dock = DockStyle.Fill,
            ForeColor = Color.FromArgb(105, 100, 97),
            TextAlign = ContentAlignment.TopLeft,
            Margin = new Padding(1, 2, 0, 0)
        }, 0, 1);
        root.Controls.Add(header, 0, 0);

        Control body;
        if (pending.Count == 0)
        {
            body = new Label
            {
                Text = "승인을 기다리는 네이버 예약 신청이 없습니다.\n새 신청이 들어오면 상단 '예약신청' 버튼이 주황색으로 바뀝니다.",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.FromArgb(120, 115, 112),
                BackColor = Color.FromArgb(250, 250, 249),
                BorderStyle = BorderStyle.FixedSingle
            };
        }
        else
        {
            var grid = new DataGridView();
            ConfigureGrid(grid);
            grid.ScrollBars = ScrollBars.Vertical;
            grid.RowTemplate.Height = 44;
            grid.DefaultCellStyle.Padding = new Padding(6, 0, 6, 0);
            foreach (var (name, title, weight, minimumWidth) in new[]
                     {
                         ("date", "이용일", 18F, 104),
                         ("time", "시간", 16F, 96),
                         ("place", "공간", 18F, 100),
                         ("name", "예약자", 14F, 72),
                         ("headcount", "인원", 8F, 48),
                         ("purpose", "이용 목적", 26F, 110)
                     })
            {
                var column = new DataGridViewTextBoxColumn
                {
                    Name = name,
                    HeaderText = title,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                    FillWeight = weight,
                    MinimumWidth = minimumWidth,
                    SortMode = DataGridViewColumnSortMode.NotSortable
                };
                grid.Columns.Add(column);
            }
            grid.Columns["headcount"]!.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.Columns["headcount"]!.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            foreach (var (item, entry) in pending)
            {
                var dateText = DateTime.TryParse(entry.Date, out var date)
                    ? date.ToString("M월 d일 (ddd)", KoreanCulture)
                    : FirstNonEmpty(entry.Date, "-");
                var timeText = string.IsNullOrWhiteSpace(entry.StartTime)
                    ? FirstNonEmpty(item.UseDateTime, "-")
                    : $"{entry.StartTime}~{entry.EndTime}";
                var rowIndex = grid.Rows.Add(
                    dateText,
                    timeText,
                    FirstNonEmpty(entry.PlaceName, "-"),
                    FirstNonEmpty(entry.Name, "-"),
                    $"{Math.Max(1, item.Headcount)}명",
                    FirstNonEmpty(item.Purpose, "-"));
                grid.Rows[rowIndex].Cells["date"].ToolTipText = $"예약번호 {item.ReservationNo}";
                if (DateTime.TryParse(entry.Date, out var rowDate) && rowDate.Date == DateTime.Today)
                    grid.Rows[rowIndex].DefaultCellStyle.Font = new Font(grid.DefaultCellStyle.Font!, FontStyle.Bold);
            }
            grid.ShowCellToolTips = true;
            grid.CellDoubleClick += (_, eventArgs) =>
            {
                if (eventArgs.RowIndex >= 0) OpenChrome(SmartPlaceUrl);
            };
            // 진한 검은 테두리 대신 옅은 1px 테두리로 표 영역만 구분한다.
            var frame = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(1),
                BackColor = Color.FromArgb(222, 218, 214),
                Margin = new Padding(0)
            };
            frame.Controls.Add(grid);
            body = frame;
        }
        root.Controls.Add(body, 0, 1);

        // 하단: 왼쪽에 도움말, 오른쪽에 버튼(주 동작이 맨 오른쪽)
        var footer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = new Padding(0, 12, 0, 0) };
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        footer.Controls.Add(new Label
        {
            Text = pending.Count > 0 ? "오늘 이용 예약은 굵게 표시 · 행을 두 번 누르면 네이버가 열립니다" : "",
            Dock = DockStyle.Fill,
            ForeColor = Color.FromArgb(140, 135, 132),
            TextAlign = ContentAlignment.MiddleLeft
        }, 0, 0);
        var buttons = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Margin = new Padding(0)
        };
        var closeButton = CreateHeaderButton("닫기", (_, _) => dialog.Close());
        var openButton = CreateHeaderButton(
            pending.Count > 0 ? "네이버에서 승인하기" : "네이버 열기",
            (_, _) => OpenChrome(SmartPlaceUrl));
        if (pending.Count > 0) ApplyHeaderButtonStyle(openButton, Color.FromArgb(44, 126, 75));
        buttons.Controls.Add(closeButton);
        buttons.Controls.Add(openButton);
        footer.Controls.Add(buttons, 1, 0);
        root.Controls.Add(footer, 0, 2);
        dialog.CancelButton = closeButton;

        return dialog;
    }

    private static async Task<(int ExitCode, string OutputText, string ErrorText)> RunNaverBookingSyncAsync(
        string repoRoot,
        string naverId,
        string naverPassword,
        CancellationToken cancellationToken)
    {
        using var process = new Process();
        process.StartInfo = new ProcessStartInfo
        {
            FileName = "cmd.exe",
            Arguments = "/c npm.cmd run naver:sync:headless",
            WorkingDirectory = repoRoot,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        if (!string.IsNullOrWhiteSpace(naverId))
            process.StartInfo.Environment["NAVER_BOOKING_ID"] = naverId;
        if (!string.IsNullOrWhiteSpace(naverPassword))
            process.StartInfo.Environment["NAVER_BOOKING_PASSWORD"] = naverPassword;
        process.Start();
        var outputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var errorTask = process.StandardError.ReadToEndAsync(cancellationToken);
        await process.WaitForExitAsync(cancellationToken);
        return (process.ExitCode, await outputTask, await errorTask);
    }

    // 반환값: 화면에 반영한 '오늘이용' 건수(취소 건 제외)
    private int MergeNaverTodayUseReservations(NaverBookingSnapshot snapshot)
    {
        var entries = NaverBookingParser.BuildSmartPlaceEntries(snapshot);

        var withoutPreviousSmartPlace = _calendarEntries.Where(item => !IsSmartPlaceNaverEntry(item));
        _calendarEntries = ReservationService.MergeCalendarEntries(withoutPreviousSmartPlace, entries, []);
        CacheVisibleCalendarEntries();
        RenderCalendarEntries();
        return entries.Count(item => !item.Cancelled);
    }

    private void MergeLatestNaverBookingOutput()
    {
        var repoRoot = FindRepositoryRoot();
        if (string.IsNullOrWhiteSpace(repoRoot)) return;
        var outputPath = Path.Combine(repoRoot, "outputs", "naver-booking-sync.json");
        try
        {
            var snapshot = NaverBookingParser.ReadSnapshot(outputPath);
            var entries = NaverBookingParser.BuildSmartPlaceEntries(snapshot);
            if (entries.Count == 0) return;
            var withoutPreviousSmartPlace = _calendarEntries.Where(item => !IsSmartPlaceNaverEntry(item));
            _calendarEntries = ReservationService.MergeCalendarEntries(withoutPreviousSmartPlace, entries, []);
            CacheVisibleCalendarEntries();
        }
        catch
        {
            // The live SmartPlace refresh will surface detailed errors; stale local output is only a fast display hint.
        }
    }

    private static string FirstMeaningfulLine(params string[] values) =>
        values.SelectMany(value => (value ?? "").Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
            .Select(line => line.Trim())
            .FirstOrDefault(line =>
                !string.IsNullOrWhiteSpace(line) &&
                !line.StartsWith(">", StringComparison.Ordinal) &&
                !line.StartsWith("npm notice", StringComparison.OrdinalIgnoreCase) &&
                !line.StartsWith("npm warn", StringComparison.OrdinalIgnoreCase)) ?? "";

    private static string FindRepositoryRoot()
    {
        var configuredRoot = Environment.GetEnvironmentVariable("SPACE_MANAGEMENT_REPO");
        var defaultUserRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "ChatGPT",
            "청년이봄프로그램");
        foreach (var start in new[] { configuredRoot, Directory.GetCurrentDirectory(), AppContext.BaseDirectory, defaultUserRoot })
        {
            if (string.IsNullOrWhiteSpace(start)) continue;
            var directory = new DirectoryInfo(start);
            while (directory is not null)
            {
                if (File.Exists(Path.Combine(directory.FullName, "package.json")) &&
                    File.Exists(Path.Combine(directory.FullName, "automation", "naver-booking-sync.js")))
                {
                    return directory.FullName;
                }
                directory = directory.Parent;
            }
        }
        return "";
    }

    private void ConfigureTray()
    {
        var menu = new ContextMenuStrip();
        menu.Items.Add("열기", null, (_, _) => RestoreFromTray());
        menu.Items.Add("새로고침", null, async (_, _) => await RefreshAllAsync());
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("종료", null, (_, _) =>
        {
            _allowClose = true;
            Close();
        });
        _notifyIcon.Icon = SystemIcons.Information;
        _notifyIcon.Text = "청년이봄 직원 운영";
        _notifyIcon.Visible = true;
        _notifyIcon.ContextMenuStrip = menu;
        _notifyIcon.DoubleClick += (_, _) => RestoreFromTray();
        _notifyIcon.BalloonTipClicked += (_, _) =>
        {
            if (!_overdueBalloonShown) return;
            _overdueBalloonShown = false;
            ShowTodayOverdueReservations();
        };
        _notifyIcon.BalloonTipClosed += (_, _) => _overdueBalloonShown = false;

        // 서버 새로고침이 없어도 시간이 지나 30분을 넘기는 예약을 놓치지 않도록 1분마다 확인한다.
        _overdueCheckTimer.Tick += (_, _) => CheckOverdueUsageLogs();
        _overdueCheckTimer.Start();
    }

    private void RestoreFromTray()
    {
        Show();
        WindowState = FormWindowState.Normal;
        Activate();
    }

    private void HideToTray(bool showBalloonTip)
    {
        Hide();
        if (!showBalloonTip) return;
        const string trayTitle = "\uCCAD\uB144\uC774\uBD04 \uC9C1\uC6D0 \uC6B4\uC601";
        const string trayMessage = "\uBC31\uADF8\uB77C\uC6B4\uB4DC\uC5D0\uC11C \uD604\uC7A5 \uC608\uC57D\uC744 \uD655\uC778\uD558\uACE0 \uC788\uC2B5\uB2C8\uB2E4.";
        _notifyIcon.ShowBalloonTip(1500, trayTitle, trayMessage, ToolTipIcon.Info);
        _trayNoticeShown = true;
    }

    private void OpenChrome(string url)
    {
        try
        {
            if (File.Exists(ChromePath))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = ChromePath,
                    Arguments = $"--profile-directory=\"{ChromeProfile}\" \"{url}\"",
                    UseShellExecute = true
                });
                return;
            }
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        }
        catch (Exception exception)
        {
            MessageBox.Show(this, "바로가기를 열지 못했습니다.\n" + exception.Message,
                "바로가기 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private static NaverBookingSnapshot CreateDemoNaverSnapshot()
    {
        string Use(int dayOffset, string time)
        {
            var date = DateTime.Today.AddDays(dayOffset);
            return $"{date:yy}. {date.Month}. {date.Day}.({date.ToString("ddd", KoreanCulture)}) {time}";
        }

        return new NaverBookingSnapshot
        {
            PendingConfirm =
            [
                new NaverBookingReservation { ReservationNo = "1400000001", Status = "신청", BookerName = "김신청", ProductName = "청년 커넥트룸", UseDateTime = Use(0, "오후 4:00~6:00"), Headcount = 4, Purpose = "스터디 모임" },
                new NaverBookingReservation { ReservationNo = "1400000002", Status = "신청", BookerName = "이봄", ProductName = "AI 커리어존 (청취Room1)", UseDateTime = Use(1, "오전 10:00~11:00"), Headcount = 1, Purpose = "온라인 면접" },
                new NaverBookingReservation { ReservationNo = "1400000003", Status = "신청", BookerName = "박청년", ProductName = "AI 커리어존 (청취Room2)", UseDateTime = Use(3, "오후 2:00~4:00"), Headcount = 2, Purpose = "자기소개서 첨삭" }
            ],
            TodayCancel =
            [
                new NaverBookingReservation { ReservationNo = "1400000009", Status = "취소", BookerName = "최취소", ProductName = "청년 커넥트룸", UseDateTime = Use(0, "오후 7:00~8:00") }
            ]
        };
    }

    private void LoadDemoData()
    {
        _state = new LocalState
        {
            InitialSyncCompleted = true,
            Events =
            [
                new StaffEvent { Id = "career:1", Category = "career", CategoryLabel = "커리어존", Name = "김봄", Date = "2026-07-28", StartTime = "16:00", EndTime = "17:00", Purpose = "진로 상담 준비", Headcount = 2, OccurredAt = "2026-07-28 15:42:10", IsRead = false },
                new StaffEvent { Id = "restricted:1", Category = "restricted", CategoryLabel = "이용 제한", Name = "최회원", Date = "2026-07-28", Purpose = "만 40세 · 시설 이용 가능 연령 초과 · 끝자리 0926", Headcount = 1, OccurredAt = "2026-07-28 15:38:20", IsRead = false },
                new StaffEvent { Id = "connect:1", Category = "connect", CategoryLabel = "커넥트룸", Name = "이청년", Date = "2026-07-28", StartTime = "15:30", EndTime = "17:30", Purpose = "팀 프로젝트 회의", Headcount = 4, OccurredAt = "2026-07-28 15:31:04", IsRead = false },
                new StaffEvent { Id = "visit:1", Category = "visit", CategoryLabel = "방문", Name = "박이봄", Date = "2026-07-28", OccurredAt = "2026-07-28 14:55:21", IsRead = true }
            ]
        };
        RenderSummary(new DashboardSummary
        {
            VisitorStats = new VisitorStats { Daily = 18, Weekly = 84, Monthly = 326 },
            ActiveMemberCount = 1248
        });
        // 오늘 날짜를 기준으로 만든다. 지금보다 이른 시간대에 '일지 미제출'·'노쇼'·'이용완료'가 보이게 한다.
        var today = DateTime.Today;
        var month = today.ToString("yyyy-MM");
        string Day(int offset) => today.AddDays(offset).ToString("yyyy-MM-dd");
        string Hour(int offsetFromNow) =>
            $"{Math.Clamp(DateTime.Now.Hour + offsetFromNow, 0, 22):00}:00";
        _selectedCalendarDate = today;
        _visibleCalendarMonth = new DateTime(today.Year, today.Month, 1);
        _loadedCalendarMonth = month;
        _calendarEntries =
        [
            new CalendarEntry { Id = "smart:1", Source = "naverSmartPlace", Date = Day(0), StartTime = Hour(-4), EndTime = Hour(-3), PlaceName = "커리어존 Room1", Name = "김청년", Headcount = 2, Status = "확정" },
            new CalendarEntry { Id = "smart:2", Source = "naverSmartPlace", Date = Day(0), StartTime = Hour(-3), EndTime = Hour(-2), PlaceName = "커리어존 Room2", Name = "박노쇼", Status = "확정" },
            new CalendarEntry { Id = "usage:1", Source = "usage", Date = Day(0), StartTime = Hour(-2), EndTime = Hour(0), PlaceName = "커넥트룸", Name = "이이봄", Headcount = 4, Status = "이용완료", Submitted = true },
            new CalendarEntry { Id = "web:1", Source = "reservationRequest", Date = Day(0), StartTime = Hour(1), EndTime = Hour(2), PlaceName = "커넥트룸", Name = "정신청", Headcount = 3, Status = "신청" },
            new CalendarEntry { Id = "smart:3", Source = "naverSmartPlace", Date = Day(0), StartTime = Hour(2), EndTime = Hour(3), PlaceName = "커리어존 Room2", Name = "최청년", Status = "취소", Cancelled = true },
            new CalendarEntry { Id = "smart:4", Source = "naverSmartPlace", Date = Day(2), StartTime = "13:00", EndTime = "14:00", PlaceName = "커넥트룸", Name = "정이봄", Status = "확정" },
            new CalendarEntry { Id = "smart:5", Source = "naverSmartPlace", Date = Day(2), StartTime = "15:00", EndTime = "16:00", PlaceName = "커리어존 Room1", Name = "한봄", Status = "확정" },
            new CalendarEntry { Id = "smart:6", Source = "naverSmartPlace", Date = Day(2), StartTime = "17:00", EndTime = "18:00", PlaceName = "커리어존 Room2", Name = "윤청년", Status = "확정" }
        ];
        _checkInsByMonth[month] = [new StaffCheckIn { Date = Day(0), Name = "김청년", MemberId = "demo-member-1" }];
        RenderCalendarEntries();
        RenderEvents();
        _statusLabel.Text = $"미리보기 데이터 · 마지막 확인 {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
    }

    private void CapturePreview(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var bitmap = new Bitmap(Width, Height);
        DrawToBitmap(bitmap, new Rectangle(0, 0, Width, Height));
        bitmap.Save(path);
    }

    private static Panel CreateSectionPanel() => new()
    {
        Dock = DockStyle.Fill,
        BackColor = Color.White,
        BorderStyle = BorderStyle.FixedSingle,
        Padding = new Padding(0)
    };

    private static Label CreateSectionTitle(string text) => new()
    {
        Text = text,
        Font = new Font("맑은 고딕", 13F, FontStyle.Bold),
        ForeColor = Color.FromArgb(52, 48, 46),
        AutoSize = true
    };

    private static readonly Color MissingUsageLogColor = Color.FromArgb(196, 84, 20);
    private static readonly Font HeaderButtonFont = new("맑은 고딕", 9.5F);
    private static readonly Font HeaderButtonBoldFont = new("맑은 고딕", 9.5F, FontStyle.Bold);

    private static Button CreateHeaderButton(string text, EventHandler click)
    {
        var button = new Button { Text = text };
        ApplyHeaderButtonStyle(button);
        button.Click += click;
        return button;
    }

    // 상단 버튼은 모두 같은 모양으로 두고, 처리할 일이 있을 때만 accent 색을 채운다.
    private static void ApplyHeaderButtonStyle(Button button, Color? accent = null)
    {
        button.AutoSize = true;
        button.MinimumSize = new Size(76, 40);
        button.Height = 40;
        button.Padding = new Padding(12, 0, 12, 0);
        button.Margin = new Padding(6, 0, 0, 0);
        button.FlatStyle = FlatStyle.Flat;
        button.Cursor = Cursors.Hand;
        if (accent is { } color)
        {
            button.BackColor = color;
            button.ForeColor = Color.White;
            button.Font = HeaderButtonBoldFont;
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = ControlPaint.Light(color, 0.15F);
            return;
        }

        button.BackColor = Color.White;
        button.ForeColor = Color.FromArgb(52, 48, 46);
        button.Font = HeaderButtonFont;
        button.FlatAppearance.BorderSize = 1;
        button.FlatAppearance.BorderColor = Color.FromArgb(214, 210, 206);
        button.FlatAppearance.MouseOverBackColor = Color.FromArgb(245, 245, 243);
    }

    private static Button CreateSmallButton(string text, EventHandler click)
    {
        var button = new Button
        {
            Text = text,
            Width = 86,
            Height = 30,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.White
        };
        button.Click += click;
        return button;
    }

    private static void ConfigureGrid(DataGridView grid)
    {
        grid.Dock = DockStyle.Fill;
        grid.BackgroundColor = Color.White;
        grid.BorderStyle = BorderStyle.None;
        grid.ShowCellToolTips = false;
        grid.ReadOnly = true;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.AllowUserToResizeColumns = true;
        grid.AllowUserToResizeRows = false;
        grid.MultiSelect = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.RowHeadersVisible = false;
        grid.AutoGenerateColumns = false;
        grid.ColumnHeadersHeight = 38;
        grid.RowTemplate.Height = 38;
        grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
        grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
        grid.EnableHeadersVisualStyles = false;
        grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 243);
        grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(245, 245, 243);
        grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = Color.FromArgb(70, 66, 64);
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(70, 66, 64);
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
        grid.DefaultCellStyle.Font = new Font("맑은 고딕", 9.5F);
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(238, 242, 239);
        grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(52, 48, 46);
        grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 249);
        grid.GridColor = Color.FromArgb(232, 230, 228);
    }
}
