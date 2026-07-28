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
    private readonly NotifyIcon _notifyIcon = new();
    private readonly ComboBox _categoryFilter = new();
    private readonly CheckBox _unreadOnly = new();
    private readonly DataGridView _eventsGrid = new();
    private readonly DataGridView _reservationsGrid = new();
    private readonly MonthCalendar _operationCalendar = new();
    private readonly Label _reservationsTitle = CreateSectionTitle("오늘 예약");
    private readonly Label _dailyValue = CreateStatValue();
    private readonly Label _weeklyValue = CreateStatValue();
    private readonly Label _monthlyValue = CreateStatValue();
    private readonly Label _memberValue = CreateStatValue();
    private readonly Label _statusLabel = new();
    private readonly Label _unreadLabel = new();
    private DashboardSettings _settings;
    private LocalState _state;
    private string _password;
    private string _loadedCalendarMonth = "";
    private List<CalendarEntry> _calendarEntries = [];
    private bool _refreshing;
    private bool _runtimeSyncCompleted;
    private bool _allowClose;
    private readonly bool _demoMode;
    private readonly string _previewPath;

    public MainForm(bool demoMode = false, string previewPath = "")
    {
        _demoMode = demoMode;
        _previewPath = previewPath;
        _settings = _storage.LoadSettings();
        _state = _storage.LoadState();
        _password = _storage.UnprotectPassword(_settings.ProtectedPassword);

        Text = "청년이봄 직원 운영";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(1100, 700);
        Size = new Size(1380, 860);
        Font = new Font("맑은 고딕", 9.5F);
        BackColor = Color.FromArgb(247, 247, 246);
        Icon = SystemIcons.Application;

        BuildLayout();
        ConfigureTray();

        _pollTimer.Interval = Math.Max(15, _settings.PollSeconds) * 1000;
        _pollTimer.Tick += async (_, _) => await RefreshDashboardAsync();

        Shown += async (_, _) =>
        {
            if (_demoMode)
            {
                LoadDemoData();
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
        FormClosing += OnFormClosing;
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(22),
            BackColor = BackColor
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 112));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        Controls.Add(root);

        root.Controls.Add(BuildHeader(), 0, 0);
        root.Controls.Add(BuildStats(), 0, 1);
        root.Controls.Add(BuildMainContent(), 0, 2);
        root.Controls.Add(BuildStatusBar(), 0, 3);
    }

    private Control BuildHeader()
    {
        var panel = new Panel { Dock = DockStyle.Fill };
        var title = new Label
        {
            Text = "청년이봄 직원 운영",
            Font = new Font("맑은 고딕", 20F, FontStyle.Bold),
            ForeColor = Color.FromArgb(52, 48, 46),
            AutoSize = true,
            Location = new Point(0, 5)
        };
        var subtitle = new Label
        {
            Text = "현장예약 알림과 오늘 운영 현황",
            ForeColor = Color.FromArgb(120, 115, 112),
            AutoSize = true,
            Location = new Point(3, 44)
        };
        panel.Controls.Add(title);
        panel.Controls.Add(subtitle);

        var actions = new FlowLayoutPanel
        {
            Dock = DockStyle.Right,
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(0, 12, 0, 0)
        };
        actions.Controls.Add(CreateActionButton("시트 열기", (_, _) => OpenChrome(SheetUrl), Color.FromArgb(74, 156, 68)));
        actions.Controls.Add(CreateActionButton("네이버 스마트플레이스", (_, _) => OpenChrome(SmartPlaceUrl), Color.FromArgb(3, 199, 90)));
        actions.Controls.Add(CreateActionButton("새로고침", async (_, _) => await RefreshDashboardAsync(), Color.FromArgb(70, 66, 64)));
        actions.Controls.Add(CreateActionButton("설정", async (_, _) =>
        {
            if (ShowSettings())
            {
                _pollTimer.Start();
                await RefreshDashboardAsync();
            }
        }, Color.FromArgb(117, 111, 108)));
        panel.Controls.Add(actions);
        return panel;
    }

    private Control BuildStats()
    {
        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 4,
            Padding = new Padding(0, 8, 0, 12)
        };
        for (var i = 0; i < 4; i++) table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
        table.Controls.Add(CreateStatPanel("오늘 방문", _dailyValue, Color.FromArgb(233, 63, 111)), 0, 0);
        table.Controls.Add(CreateStatPanel("이번 주 방문", _weeklyValue, Color.FromArgb(74, 156, 68)), 1, 0);
        table.Controls.Add(CreateStatPanel("이번 달 방문", _monthlyValue, Color.FromArgb(242, 153, 48)), 2, 0);
        table.Controls.Add(CreateStatPanel("가입 회원", _memberValue, Color.FromArgb(55, 108, 142)), 3, 0);
        return table;
    }

    private Control BuildMainContent()
    {
        var split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,
            SplitterDistance = 680,
            SplitterWidth = 10,
            BackColor = BackColor
        };
        split.Panel1.Padding = new Padding(0, 0, 4, 0);
        split.Panel2.Padding = new Padding(4, 0, 0, 0);
        split.Panel1.Controls.Add(BuildEventsPanel());
        split.Panel2.Controls.Add(BuildReservationsPanel());
        return split;
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
        _categoryFilter.Items.AddRange(["전체", "방문", "커넥트룸", "커리어존"]);
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
        _reservationsTitle.Dock = DockStyle.Top;
        _reservationsTitle.Height = 52;
        _reservationsTitle.Padding = new Padding(16, 14, 0, 0);
        panel.Controls.Add(_reservationsTitle);

        var calendarPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 206,
            Padding = new Padding(12, 0, 12, 10),
            BackColor = Color.White
        };
        _operationCalendar.Dock = DockStyle.Fill;
        _operationCalendar.MaxSelectionCount = 1;
        _operationCalendar.ShowTodayCircle = true;
        _operationCalendar.TitleBackColor = Color.FromArgb(74, 156, 68);
        _operationCalendar.DateSelected += async (_, _) =>
        {
            var month = _operationCalendar.SelectionStart.ToString("yyyy-MM");
            if (month != _loadedCalendarMonth && !_demoMode)
                await RefreshDashboardAsync();
            else
                RenderCalendarEntries();
        };
        calendarPanel.Controls.Add(_operationCalendar);
        panel.Controls.Add(calendarPanel);

        ConfigureGrid(_reservationsGrid);
        _reservationsGrid.Columns.Add("time", "시간");
        _reservationsGrid.Columns.Add("reservation", "예약 정보");
        _reservationsGrid.Columns.Add("status", "상태");
        _reservationsGrid.Columns["time"]!.Width = 96;
        _reservationsGrid.Columns["reservation"]!.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        _reservationsGrid.Columns["status"]!.Width = 82;
        panel.Controls.Add(_reservationsGrid);
        _reservationsGrid.BringToFront();
        return panel;
    }

    private Control BuildStatusBar()
    {
        var panel = new Panel { Dock = DockStyle.Fill };
        _statusLabel.Text = "연결 대기";
        _statusLabel.ForeColor = Color.FromArgb(105, 100, 97);
        _statusLabel.AutoSize = true;
        _statusLabel.Location = new Point(3, 9);
        panel.Controls.Add(_statusLabel);
        return panel;
    }

    private async Task RefreshDashboardAsync()
    {
        if (_refreshing || _demoMode) return;
        _refreshing = true;
        _statusLabel.Text = "운영 정보를 확인하는 중...";
        try
        {
            using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            var requestedMonth = _operationCalendar.SelectionStart.ToString("yyyy-MM");
            var eventLimit = _runtimeSyncCompleted ? 100 : 500;
            var snapshot = await _client.GetSnapshotAsync(
                _settings,
                _password,
                requestedMonth,
                eventLimit,
                cancellation.Token);
            MergeEvents(snapshot.Events);
            RenderSummary(snapshot.Summary);
            _calendarEntries = snapshot.CalendarEntries.Count > 0
                ? snapshot.CalendarEntries
                : ConvertTodayReservations(snapshot.TodayReservations);
            _loadedCalendarMonth = requestedMonth;
            RenderCalendarEntries();
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
            _notifyIcon.ShowBalloonTip(
                6000,
                $"새 현장예약 {newAlertEvents.Count}건",
                $"커리어존 {careerCount}건 · 커넥트룸 {connectCount}건\n프로그램에서 내용을 확인해 주세요.",
                ToolTipIcon.Info);
        }
    }

    private void RenderEvents()
    {
        var selected = _categoryFilter.SelectedItem?.ToString() ?? "전체";
        var category = selected switch
        {
            "방문" => "visit",
            "커넥트룸" => "connect",
            "커리어존" => "career",
            _ => ""
        };
        var filtered = _state.Events
            .Where(item => string.IsNullOrEmpty(category) || item.Category == category)
            .Where(item => !_unreadOnly.Checked || !item.IsRead)
            .ToList();

        _eventsGrid.Rows.Clear();
        foreach (var item in filtered)
        {
            var detail = item.Category == "visit"
                ? "방문 등록"
                : $"{item.StartTime}~{item.EndTime} · {item.Headcount}명 · {item.Purpose}";
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
        var selectedDate = _operationCalendar.SelectionStart.ToString("yyyy-MM-dd");
        _reservationsTitle.Text = selectedDate == DateTime.Today.ToString("yyyy-MM-dd")
            ? "오늘 예약·이용"
            : $"{_operationCalendar.SelectionStart:M월 d일} 예약·이용";
        var boldDates = _calendarEntries
            .Select(item => DateTime.TryParse(item.Date, out var date) ? date.Date : DateTime.MinValue)
            .Where(date => date != DateTime.MinValue)
            .Distinct()
            .ToArray();
        _operationCalendar.BoldedDates = boldDates;
        _operationCalendar.UpdateBoldedDates();

        _reservationsGrid.Rows.Clear();
        foreach (var item in _calendarEntries.Where(item => item.Date == selectedDate))
        {
            var state = item.Cancelled
                ? "취소/노쇼"
                : item.Source == "usage"
                    ? "이용완료"
                    : item.Submitted ? "이용 제출" : item.Status;
            var rowIndex = _reservationsGrid.Rows.Add(
                $"{item.StartTime}~{item.EndTime}",
                $"{item.PlaceName} · {item.Name}",
                state);
            if (item.Cancelled)
                _reservationsGrid.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.Gray;
            else if (item.Submitted)
                _reservationsGrid.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.FromArgb(74, 156, 68);
        }
    }

    private static List<CalendarEntry> ConvertTodayReservations(IEnumerable<Reservation> reservations) =>
        reservations.Select(item => new CalendarEntry
        {
            Id = item.Id,
            Source = "naver",
            Date = DateTime.Today.ToString("yyyy-MM-dd"),
            StartTime = item.StartTime,
            EndTime = item.EndTime,
            PlaceName = item.PlaceName,
            Name = item.BookerName,
            Purpose = item.Purpose,
            Status = item.Status,
            Cancelled = item.Cancelled,
            Submitted = item.Submitted
        }).ToList();

    private void RenderSummary(DashboardSummary summary)
    {
        _dailyValue.Text = summary.VisitorStats.Daily.ToString("N0");
        _weeklyValue.Text = summary.VisitorStats.Weekly.ToString("N0");
        _monthlyValue.Text = summary.VisitorStats.Monthly.ToString("N0");
        _memberValue.Text = summary.ActiveMemberCount.ToString("N0");
    }

    private void MarkAllRead()
    {
        foreach (var item in _state.Events) item.IsRead = true;
        _storage.SaveState(_state);
        RenderEvents();
    }

    private bool ShowSettings()
    {
        using var dialog = new SettingsForm(_settings, _password);
        if (dialog.ShowDialog(this) != DialogResult.OK) return false;
        _password = dialog.Password;
        _settings.AppsScriptUrl = dialog.AppsScriptUrl;
        _settings.PollSeconds = dialog.PollSeconds;
        _settings.ProtectedPassword = _storage.ProtectPassword(_password);
        _storage.SaveSettings(_settings);
        _client.ResetSession();
        _pollTimer.Interval = _settings.PollSeconds * 1000;
        return true;
    }

    private void ConfigureTray()
    {
        var menu = new ContextMenuStrip();
        menu.Items.Add("열기", null, (_, _) => RestoreFromTray());
        menu.Items.Add("새로고침", null, async (_, _) => await RefreshDashboardAsync());
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
    }

    private void RestoreFromTray()
    {
        Show();
        WindowState = FormWindowState.Normal;
        Activate();
    }

    private void OnFormClosing(object? sender, FormClosingEventArgs eventArgs)
    {
        if (_allowClose || eventArgs.CloseReason == CloseReason.WindowsShutDown)
        {
            _notifyIcon.Visible = false;
            return;
        }

        var result = MessageBox.Show(this,
            "프로그램을 종료하면 현장예약 알림도 중지됩니다.\n종료하시겠습니까?",
            "프로그램 종료",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);
        if (result == DialogResult.Yes)
        {
            _allowClose = true;
            _notifyIcon.Visible = false;
            return;
        }
        eventArgs.Cancel = true;
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

    private void LoadDemoData()
    {
        _state = new LocalState
        {
            InitialSyncCompleted = true,
            Events =
            [
                new StaffEvent { Id = "career:1", Category = "career", CategoryLabel = "커리어존", Name = "김봄", Date = "2026-07-28", StartTime = "16:00", EndTime = "17:00", Purpose = "진로 상담 준비", Headcount = 2, OccurredAt = "2026-07-28 15:42:10", IsRead = false },
                new StaffEvent { Id = "connect:1", Category = "connect", CategoryLabel = "커넥트룸", Name = "이청년", Date = "2026-07-28", StartTime = "15:30", EndTime = "17:30", Purpose = "팀 프로젝트 회의", Headcount = 4, OccurredAt = "2026-07-28 15:31:04", IsRead = false },
                new StaffEvent { Id = "visit:1", Category = "visit", CategoryLabel = "방문", Name = "박이봄", Date = "2026-07-28", OccurredAt = "2026-07-28 14:55:21", IsRead = true }
            ]
        };
        RenderSummary(new DashboardSummary
        {
            VisitorStats = new VisitorStats { Daily = 18, Weekly = 84, Monthly = 326 },
            ActiveMemberCount = 1248
        });
        _operationCalendar.SetDate(new DateTime(2026, 7, 28));
        _loadedCalendarMonth = "2026-07";
        _calendarEntries =
        [
            new CalendarEntry { Id = "naver:1", Source = "naver", Date = "2026-07-28", StartTime = "14:00", EndTime = "15:00", PlaceName = "커리어존 Room1", Name = "김청년", Status = "확정" },
            new CalendarEntry { Id = "usage:1", Source = "usage", Date = "2026-07-28", StartTime = "16:00", EndTime = "18:00", PlaceName = "커넥트룸", Name = "이이봄", Status = "이용완료", Submitted = true },
            new CalendarEntry { Id = "naver:3", Source = "naver", Date = "2026-07-28", StartTime = "19:00", EndTime = "20:00", PlaceName = "커리어존 Room2", Name = "최청년", Status = "취소", Cancelled = true },
            new CalendarEntry { Id = "naver:4", Source = "naver", Date = "2026-07-30", StartTime = "13:00", EndTime = "14:00", PlaceName = "커넥트룸", Name = "정이봄", Status = "확정" }
        ];
        RenderCalendarEntries();
        RenderEvents();
        _statusLabel.Text = "미리보기 데이터 · 마지막 확인 2026-07-28 16:10:00";
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
        BorderStyle = BorderStyle.FixedSingle
    };

    private static Label CreateSectionTitle(string text) => new()
    {
        Text = text,
        Font = new Font("맑은 고딕", 13F, FontStyle.Bold),
        ForeColor = Color.FromArgb(52, 48, 46),
        AutoSize = true
    };

    private static Label CreateStatValue() => new()
    {
        Text = "-",
        Font = new Font("맑은 고딕", 22F, FontStyle.Bold),
        ForeColor = Color.FromArgb(52, 48, 46),
        AutoSize = true,
        Location = new Point(16, 34)
    };

    private static Control CreateStatPanel(string title, Label value, Color accent)
    {
        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            Margin = new Padding(0, 0, 10, 0),
            BorderStyle = BorderStyle.FixedSingle
        };
        var accentLine = new Panel { Dock = DockStyle.Left, Width = 5, BackColor = accent };
        var label = new Label
        {
            Text = title,
            ForeColor = Color.FromArgb(105, 100, 97),
            AutoSize = true,
            Location = new Point(18, 12)
        };
        panel.Controls.Add(accentLine);
        panel.Controls.Add(label);
        panel.Controls.Add(value);
        return panel;
    }

    private static Button CreateActionButton(string text, EventHandler click, Color color)
    {
        var button = new Button
        {
            Text = text,
            AutoSize = true,
            Height = 38,
            Padding = new Padding(12, 0, 12, 0),
            Margin = new Padding(6, 0, 0, 0),
            BackColor = color,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        button.FlatAppearance.BorderSize = 0;
        button.Click += click;
        return button;
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
        grid.ReadOnly = true;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.AllowUserToResizeRows = false;
        grid.MultiSelect = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.RowHeadersVisible = false;
        grid.AutoGenerateColumns = false;
        grid.ColumnHeadersHeight = 38;
        grid.RowTemplate.Height = 38;
        grid.EnableHeadersVisualStyles = false;
        grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 243);
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(70, 66, 64);
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(235, 242, 235);
        grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(52, 48, 46);
        grid.GridColor = Color.FromArgb(232, 230, 228);
    }
}
