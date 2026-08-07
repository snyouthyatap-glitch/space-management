using System.Diagnostics;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace YouthSpring.StaffDashboard;

internal sealed class MainForm : Form
{
    private const string SheetUrl =
        "https://docs.google.com/spreadsheets/d/1sPVvNPMl0KqQHDpCqNnM4RuOUw73DpgzK7IqcMDlVYs/edit?usp=sharing";
    private const string AdminPageUrl =
        "https://snyouthyatap-glitch.github.io/space-management/admin.html";
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
    private readonly MonthCalendar _operationCalendar = new();
    private readonly Label _reservationsTitle = CreateSectionTitle("오늘 예약");
    private readonly Label _dailyValue = CreateStatValue();
    private readonly Label _weeklyValue = CreateStatValue();
    private readonly Label _monthlyValue = CreateStatValue();
    private readonly Label _memberValue = CreateStatValue();
    private readonly Label _statusLabel = new();
    private readonly Label _naverBookingStatusLabel = new();
    private readonly Label _unreadLabel = new();
    private readonly FlowLayoutPanel _reservationActionPanel = new();
    private readonly Button _requestApproveButton = new();
    private readonly Button _requestRejectButton = new();
    private readonly Button _moreActionsButton = new();
    private DashboardSettings _settings;
    private LocalState _state;
    private string _password;
    private string _loadedCalendarMonth = "";
    private List<CalendarEntry> _calendarEntries = [];
    private bool _refreshing;
    private bool _updatingReservationStatus;
    private bool _runtimeSyncCompleted;
    private bool _naverBookingSyncRunning;
    private bool _allowClose;
    private bool _trayNoticeShown;
    private string _lastNaverPendingSignature = "";
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
        MinimumSize = new Size(1280, 700);
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
        actions.Controls.Add(CreateActionButton("실적 다운로드", async (_, _) =>
            await DownloadMonthlyPerformanceAsync(), Color.FromArgb(196, 112, 37)));
        actions.Controls.Add(CreateActionButton("수동 일지", async (_, _) =>
        {
            _pollTimer.Stop();
            try
            {
                using var dialog = new ManualLogForm(_client, _settings, _password);
                dialog.ShowDialog(this);
            }
            finally
            {
                _pollTimer.Start();
            }
            await RefreshDashboardAsync();
        }, Color.FromArgb(55, 108, 142)));
        actions.Controls.Add(CreateActionButton("네이버 스마트플레이스", (_, _) => OpenChrome(SmartPlaceUrl), Color.FromArgb(3, 199, 90)));
        actions.Controls.Add(CreateActionButton("전체 새로고침", async (_, _) => await RefreshAllAsync(), Color.FromArgb(70, 66, 64)));
        ConfigureMoreActionsButton();
        actions.Controls.Add(_moreActionsButton);
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
        var headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 96,
            Padding = new Padding(16, 14, 16, 8)
        };
        _reservationsTitle.AutoSize = false;
        _reservationsTitle.Dock = DockStyle.Top;
        _reservationsTitle.Height = 28;
        headerPanel.Controls.Add(_reservationsTitle);
        ConfigureReservationActions();
        headerPanel.Controls.Add(_reservationActionPanel);
        panel.Controls.Add(headerPanel);

        var bodySplit = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            SplitterDistance = 206,
            SplitterWidth = 8,
            Panel1MinSize = 130,
            Panel2MinSize = 120,
            BackColor = Color.FromArgb(232, 230, 228)
        };

        var calendarPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12, 0, 12, 10),
            BackColor = Color.White
        };
        _operationCalendar.Dock = DockStyle.Fill;
        _operationCalendar.CalendarDimensions = new Size(1, 1);
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
        bodySplit.Panel1.Controls.Add(calendarPanel);

        ConfigureGrid(_reservationsGrid);
        _reservationsGrid.Columns.Add("time", "시간");
        _reservationsGrid.Columns.Add("representative", "대표자");
        _reservationsGrid.Columns.Add("place", "장소");
        _reservationsGrid.Columns.Add("headcount", "인원수");
        _reservationsGrid.Columns.Add("status", "상태");
        _reservationsGrid.Columns["time"]!.Width = 150;
        _reservationsGrid.Columns["representative"]!.Width = 170;
        _reservationsGrid.Columns["place"]!.Width = 220;
        _reservationsGrid.Columns["headcount"]!.Width = 90;
        _reservationsGrid.Columns["status"]!.Width = 130;
        foreach (DataGridViewColumn column in _reservationsGrid.Columns)
        {
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            column.Resizable = DataGridViewTriState.True;
        }
        _reservationsGrid.SelectionChanged += (_, _) => UpdateReservationActionButtons();
        ConfigureReservationContextMenu();
        bodySplit.Panel2.Controls.Add(_reservationsGrid);
        panel.Controls.Add(bodySplit);
        _reservationsGrid.BringToFront();
        return panel;
    }

    private void ConfigureReservationActions()
    {
        _reservationActionPanel.Dock = DockStyle.Bottom;
        _reservationActionPanel.Height = 40;
        _reservationActionPanel.FlowDirection = FlowDirection.LeftToRight;
        _reservationActionPanel.WrapContents = false;

        ConfigureReservationActionButton(_requestApproveButton, "\uC2B9\uC778", Color.FromArgb(74, 156, 68));
        ConfigureReservationActionButton(_requestRejectButton, "\uBC18\uB824", Color.FromArgb(233, 63, 111));
        UpdateReservationActionButtons();
    }

    private void ConfigureReservationActionButton(Button button, string status, Color color)
    {
        button.Text = status;
        button.AutoSize = true;
        button.Height = 32;
        button.Padding = new Padding(10, 0, 10, 0);
        button.Margin = new Padding(0, 0, 8, 0);
        button.BackColor = color;
        button.ForeColor = Color.White;
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 0;
        button.Click += async (_, _) => await UpdateSelectedReservationStatusAsync(status);
        _reservationActionPanel.Controls.Add(button);
    }

    private void ConfigureReservationContextMenu()
    {
        var menu = new ContextMenuStrip();
        foreach (var status in new[] { "승인", "반려" })
        {
            var item = new ToolStripMenuItem(status + "로 변경") { Tag = status };
            item.Click += async (_, _) => await UpdateSelectedReservationStatusAsync(status);
            menu.Items.Add(item);
        }

        menu.Opening += (_, eventArgs) =>
        {
            var mouse = _reservationsGrid.PointToClient(Cursor.Position);
            var hit = _reservationsGrid.HitTest(mouse.X, mouse.Y);
            if (hit.RowIndex < 0 ||
                _reservationsGrid.Rows[hit.RowIndex].Tag is not CalendarEntry entry ||
                !CanChangeEntryStatus(entry))
            {
                eventArgs.Cancel = true;
                return;
            }

            _reservationsGrid.ClearSelection();
            _reservationsGrid.Rows[hit.RowIndex].Selected = true;
            _reservationsGrid.CurrentCell = _reservationsGrid.Rows[hit.RowIndex].Cells[
                Math.Max(0, hit.ColumnIndex)];
            foreach (ToolStripMenuItem menuItem in menu.Items)
            {
                var status = menuItem.Tag?.ToString() ?? "";
                menuItem.Enabled = IsAllowedStatusForEntry(entry, status);
                menuItem.Checked = string.Equals(
                    status,
                    entry.Status,
                    StringComparison.Ordinal);
            }
        };
        _reservationsGrid.ContextMenuStrip = menu;
    }

    private void ConfigureMoreActionsButton()
    {
        _moreActionsButton.Text = "더보기";
        _moreActionsButton.AutoSize = true;
        _moreActionsButton.Height = 38;
        _moreActionsButton.Padding = new Padding(12, 0, 12, 0);
        _moreActionsButton.Margin = new Padding(6, 0, 0, 0);
        _moreActionsButton.BackColor = Color.FromArgb(117, 111, 108);
        _moreActionsButton.ForeColor = Color.White;
        _moreActionsButton.FlatStyle = FlatStyle.Flat;
        _moreActionsButton.Cursor = Cursors.Hand;
        _moreActionsButton.FlatAppearance.BorderSize = 0;

        var menu = new ContextMenuStrip();
        menu.Items.Add("알림 이력", null, (_, _) => ShowEventHistoryDialog());
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("실적 다운로드", null, async (_, _) => await DownloadMonthlyPerformanceAsync());
        menu.Items.Add("관리자 페이지", null, (_, _) => OpenChrome(AdminPageUrl));
        menu.Items.Add("시트 열기", null, (_, _) => OpenChrome(SheetUrl));
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("설정", null, async (_, _) =>
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
        var panel = new Panel { Dock = DockStyle.Fill };
        _statusLabel.Text = "연결 대기";
        _statusLabel.ForeColor = Color.FromArgb(105, 100, 97);
        _statusLabel.AutoSize = true;
        _statusLabel.Dock = DockStyle.Left;
        _statusLabel.Padding = new Padding(3, 9, 16, 0);
        panel.Controls.Add(_statusLabel);

        _naverBookingStatusLabel.Text = "네이버 예약 확인 대기";
        _naverBookingStatusLabel.ForeColor = Color.FromArgb(105, 100, 97);
        _naverBookingStatusLabel.AutoSize = true;
        _naverBookingStatusLabel.Dock = DockStyle.Right;
        _naverBookingStatusLabel.Padding = new Padding(16, 9, 3, 0);
        panel.Controls.Add(_naverBookingStatusLabel);

        return panel;
    }

    private async Task RefreshDashboardAsync()
    {
        if (_refreshing || _updatingReservationStatus || _demoMode) return;
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
            var smartPlaceEntries = _calendarEntries.Where(IsSmartPlaceNaverEntry).ToList();
            _calendarEntries = MergeCalendarEntries(
                snapshot.CalendarEntries.Where(item => !IsSheetNaverEntry(item)),
                ConvertTodayReservations(snapshot.TodayReservations).Where(item => !IsSheetNaverEntry(item)),
                ConvertReservationRequests(snapshot.ReservationRequests).Concat(smartPlaceEntries));
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
            var state = GetUnifiedReservationStatus(item);
            var rowIndex = _reservationsGrid.Rows.Add(
                $"{item.StartTime}~{item.EndTime}",
                item.Name,
                item.PlaceName,
                $"{Math.Max(1, item.Headcount)}명",
                state);
            _reservationsGrid.Rows[rowIndex].Tag = item;
            ApplyReservationRowStyle(_reservationsGrid.Rows[rowIndex], state);
        }
    }

    private void UpdateReservationActionButtons()
    {
        var selectedEntry = _reservationsGrid.SelectedRows.Count > 0
            ? _reservationsGrid.SelectedRows[0].Tag as CalendarEntry
            : null;
        var canChangeRequestStatus = selectedEntry != null &&
            IsReservationRequestSource(selectedEntry.Source) &&
            !_updatingReservationStatus;

        foreach (var button in new[]
                 {
                     (_requestApproveButton, "승인"),
                     (_requestRejectButton, "반려")
                 })
        {
            button.Item1.Enabled = canChangeRequestStatus &&
                !string.Equals(selectedEntry?.Status, button.Item2, StringComparison.Ordinal);
        }
    }

    private static List<CalendarEntry> ConvertTodayReservations(IEnumerable<Reservation> reservations) =>
        reservations.Select(item => new CalendarEntry
        {
            Id = item.Id,
            Source = string.IsNullOrWhiteSpace(item.Source) ? "naver" : item.Source,
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

    private static List<CalendarEntry> MergeCalendarEntries(
        IEnumerable<CalendarEntry> calendarEntries,
        IEnumerable<CalendarEntry> todayEntries,
        IEnumerable<CalendarEntry> requestEntries)
    {
        var merged = calendarEntries
            .Concat(todayEntries)
            .Concat(requestEntries)
            .Where(item => !string.IsNullOrWhiteSpace(item.Id))
            .GroupBy(ReservationMergeKey, StringComparer.Ordinal)
            .Select(MergeDuplicateReservationEntries)
            .OrderBy(item => item.Date, StringComparer.Ordinal)
            .ThenBy(item => item.StartTime, StringComparer.Ordinal)
            .ToList();
        return merged;
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
            Submitted = entries.Any(item => item.Submitted || item.Source == "usage" || item.Status == "일지제출완료")
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

    private static string NormalizeGridTime(string value)
    {
        var time = Regex.Match(value ?? "", @"(\d{1,2}):(\d{2})");
        return time.Success ? $"{time.Groups[1].Value.PadLeft(2, '0')}:{time.Groups[2].Value}" : (value ?? "").Trim();
    }

    private async Task UpdateSelectedReservationStatusAsync(string status)
    {
        if (_updatingReservationStatus ||
            _reservationsGrid.SelectedRows.Count == 0 ||
            _reservationsGrid.SelectedRows[0].Tag is not CalendarEntry entry ||
            !CanChangeEntryStatus(entry))
        {
            return;
        }

        if (string.Equals(entry.Status, status, StringComparison.Ordinal))
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
            var result = IsReservationRequestSource(entry.Source)
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

    private static bool CanChangeEntryStatus(CalendarEntry entry) =>
        IsReservationRequestSource(entry.Source);

    private static bool IsAllowedStatusForEntry(CalendarEntry entry, string status) =>
        IsReservationRequestSource(entry.Source) && (status is "승인" or "반려");

    private static bool IsReservationRequestSource(string source) =>
        source is "reservationRequest" or "webReservation" or "web" or "app";

    private static bool IsSheetNaverEntry(CalendarEntry entry) =>
        string.Equals(entry.Source, "naver", StringComparison.Ordinal);

    private static bool IsSmartPlaceNaverEntry(CalendarEntry entry) =>
        string.Equals(entry.Source, "naverSmartPlace", StringComparison.Ordinal);

    private static string GetUnifiedReservationStatus(CalendarEntry entry)
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

        if (start.HasValue && now >= start.Value.AddMinutes(10)) return "노쇼";
        if (entry.Status == "신청") return "승인 대기";
        return "승인 완료";
    }

    private static DateTime? ParseReservationDateTime(string date, string time)
    {
        if (DateTime.TryParse($"{date} {time}", out var value)) return value;
        return null;
    }

    private void ApplyReservationRowStyle(DataGridViewRow row, string state)
    {
        row.DefaultCellStyle.ForeColor = state switch
        {
            "승인 대기" => Color.FromArgb(196, 112, 37),
            "이용중" => Color.FromArgb(55, 108, 142),
            "이용완료" => Color.FromArgb(74, 156, 68),
            "노쇼" or "취소" => Color.Gray,
            _ => Color.FromArgb(52, 48, 46)
        };
        if (state is "승인 대기" or "이용중")
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

    private static string FormatSpaceLabel(string spaceType, string room)
    {
        if (spaceType == "connectRoom") return "커넥트룸";
        if (room == "careerRoom1" || spaceType == "careerRoom1") return "커리어존 Room1";
        if (room == "careerRoom2" || spaceType == "careerRoom2") return "커리어존 Room2";
        return string.IsNullOrWhiteSpace(spaceType) ? "공간 예약" : spaceType;
    }

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

    private async Task DownloadMonthlyPerformanceAsync()
    {
        using var monthDialog = new MonthSelectionForm();
        if (monthDialog.ShowDialog(this) != DialogResult.OK) return;

        using var saveDialog = new SaveFileDialog
        {
            Title = "월간 실적 저장",
            Filter = "Excel 통합 문서 (*.xlsx)|*.xlsx",
            DefaultExt = "xlsx",
            AddExtension = true,
            FileName = $"청년이봄_{monthDialog.SelectedMonth}_월간실적.xlsx"
        };
        if (saveDialog.ShowDialog(this) != DialogResult.OK) return;

        _pollTimer.Stop();
        _statusLabel.Text = $"{monthDialog.SelectedMonth} 실적을 정리하고 있습니다...";
        Cursor = Cursors.WaitCursor;
        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(90));
            var report = await _client.GetMonthlyPerformanceAsync(
                _settings,
                _password,
                monthDialog.SelectedMonth,
                timeout.Token);
            PerformanceWorkbookWriter.Write(saveDialog.FileName, report);
            _statusLabel.Text = $"{monthDialog.SelectedMonth} 실적 저장 완료";
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
            _naverBookingStatusLabel.Text = "네이버 예약 확인 불가 · 프로젝트 폴더 없음";
            return;
        }

        _naverBookingSyncRunning = true;
        _naverBookingStatusLabel.Text = "네이버 예약 확인 중...";
        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(90));
            var result = await RunNaverBookingSyncAsync(repoRoot, timeout.Token);
            if (result.ExitCode != 0)
            {
                var message = FirstMeaningfulLine(result.ErrorText, result.OutputText);
                _naverBookingStatusLabel.Text = "네이버 예약 확인 실패" +
                    (string.IsNullOrWhiteSpace(message) ? "" : " · " + message);
                return;
            }

            var outputPath = Path.Combine(repoRoot, "outputs", "naver-booking-sync.json");
            var naverSnapshot = ReadNaverBookingSnapshot(outputPath);
            var displayedTodayUseCount = MergeNaverTodayUseReservations(naverSnapshot.TodayUse);
            var pending = naverSnapshot.PendingConfirm;
            var checkedAt = DateTime.Now.ToString("HH:mm");
            if (pending.Count == 0)
            {
                var todayUseText = naverSnapshot.TodayUse.Count > 0
                    ? $" · 오늘 이용 {naverSnapshot.TodayUse.Count}건/표시 {displayedTodayUseCount}건"
                    : "";
                _naverBookingStatusLabel.Text = $"네이버 예약 신청 없음{todayUseText} · {checkedAt}";
                _lastNaverPendingSignature = "";
                return;
            }

            _naverBookingStatusLabel.Text =
                $"네이버 예약 신청 있음 {pending.Count}건 · 오늘 이용 {naverSnapshot.TodayUse.Count}건/표시 {displayedTodayUseCount}건 · {checkedAt}";
            var signature = string.Join("|", pending.Select(item => item.ReservationNo).Order(StringComparer.Ordinal));
            if (!string.Equals(signature, _lastNaverPendingSignature, StringComparison.Ordinal))
            {
                _lastNaverPendingSignature = signature;
                _notifyIcon.ShowBalloonTip(
                    7000,
                    $"네이버 예약 신청 {pending.Count}건",
                    BuildPendingReservationNotice(pending),
                    ToolTipIcon.Info);
            }
        }
        catch (OperationCanceledException)
        {
            _naverBookingStatusLabel.Text = "네이버 예약 확인 시간 초과";
        }
        catch (Exception exception)
        {
            _naverBookingStatusLabel.Text = "네이버 예약 확인 오류 · " + exception.Message;
        }
        finally
        {
            _naverBookingSyncRunning = false;
        }
    }

    private static async Task<(int ExitCode, string OutputText, string ErrorText)> RunNaverBookingSyncAsync(
        string repoRoot,
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
        process.Start();
        var outputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var errorTask = process.StandardError.ReadToEndAsync(cancellationToken);
        await process.WaitForExitAsync(cancellationToken);
        return (process.ExitCode, await outputTask, await errorTask);
    }

    private int MergeNaverTodayUseReservations(List<NaverBookingReservation> reservations)
    {
        var entries = reservations
            .Select(ConvertNaverTodayUseReservation)
            .Where(item => !string.IsNullOrWhiteSpace(item.Id))
            .ToList();

        var withoutPreviousSmartPlace = _calendarEntries.Where(item => !IsSmartPlaceNaverEntry(item));
        _calendarEntries = MergeCalendarEntries(withoutPreviousSmartPlace, entries, []);
        RenderCalendarEntries();
        return entries.Count;
    }

    private static CalendarEntry ConvertNaverTodayUseReservation(NaverBookingReservation item)
    {
        var parsed = ParseNaverReservationText(item);
        var timeRange = ParseNaverTimeRange(FirstNonEmpty(item.UseDateTime, parsed.UseDateTime));
        return new CalendarEntry
        {
            Id = "naverSmartPlace:" + item.ReservationNo,
            Source = "naverSmartPlace",
            Date = DateTime.Today.ToString("yyyy-MM-dd"),
            StartTime = timeRange.StartTime,
            EndTime = timeRange.EndTime,
            PlaceName = ShortNaverProductName(FirstNonEmpty(item.ProductName, parsed.ProductName, "네이버 예약")),
            Name = FirstNonEmpty(parsed.BookerName, item.BookerName, "예약자"),
            Purpose = FirstNonEmpty(item.Purpose, "네이버 오늘 이용"),
            Headcount = Math.Max(1, item.Headcount),
            Status = FirstNonEmpty(parsed.Status, item.Status, "확정")
        };
    }

    private static NaverBookingSnapshot ReadNaverBookingSnapshot(string outputPath)
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
            PendingConfirm = ReadNaverResultRows(results, "pending_confirm"),
            TodayUse = ReadNaverResultRows(results, "today_use")
        };
    }

    private static List<NaverBookingReservation> ReadNaverResultRows(JsonElement results, string propertyName)
    {
        if (!results.TryGetProperty(propertyName, out var rowsElement) ||
            rowsElement.ValueKind != JsonValueKind.Array) return [];

        var rows = new List<NaverBookingReservation>();
        foreach (var item in rowsElement.EnumerateArray())
        {
            rows.Add(new NaverBookingReservation
            {
                ReservationNo = GetJsonString(item, "reservationNo"),
                ProductName = GetJsonString(item, "productName"),
                BookerName = GetJsonString(item, "bookerName"),
                Status = GetJsonString(item, "status"),
                UseDateTime = GetJsonString(item, "useDateTime"),
                Purpose = GetJsonString(item, "purpose"),
                Headcount = GetJsonInt(item, "headcount", 1),
                RawText = GetJsonString(item, "rawText")
            });
        }
        return rows;
    }

    private static string BuildPendingReservationNotice(List<NaverBookingReservation> pending)
    {
        var counts = pending
            .GroupBy(item => string.IsNullOrWhiteSpace(item.ProductName) ? "예약" : item.ProductName)
            .Select(group => $"{ShortNaverProductName(group.Key)} {group.Count()}건");
        return string.Join(" · ", counts) + "\n프로그램에서 내용을 확인해 주세요.";
    }

    private static string ShortNaverProductName(string productName)
    {
        if (productName.Contains("Room1", StringComparison.OrdinalIgnoreCase)) return "커리어존 Room1";
        if (productName.Contains("Room2", StringComparison.OrdinalIgnoreCase)) return "커리어존 Room2";
        if (productName.Contains("커넥트", StringComparison.Ordinal)) return "커넥트룸";
        return productName;
    }

    private static (string StartTime, string EndTime) ParseNaverTimeRange(string value)
    {
        var match = Regex.Match(
            value ?? "",
            @"(?:(?:\d{2,4}\.\s*)?\d{1,2}\.\s*\d{1,2}\.\([^)]+\)\s*)?(오전|오후)\s*(\d{1,2}):(\d{2})~(?:(오전|오후)\s*)?(\d{1,2}):(\d{2})");
        if (!match.Success) return ("", "");

        var startMeridiem = match.Groups[1].Value;
        var endMeridiem = string.IsNullOrWhiteSpace(match.Groups[4].Value)
            ? startMeridiem
            : match.Groups[4].Value;
        var start = ToTwentyFourHour(startMeridiem, match.Groups[2].Value, match.Groups[3].Value);
        var end = ToTwentyFourHour(endMeridiem, match.Groups[5].Value, match.Groups[6].Value);
        return (start, end);
    }

    private static string ToTwentyFourHour(string meridiem, string hourText, string minuteText)
    {
        var hour = int.TryParse(hourText, out var parsedHour) ? parsedHour : 0;
        if (meridiem == "오후" && hour < 12) hour += 12;
        if (meridiem == "오전" && hour == 12) hour = 0;
        return $"{hour:00}:{minuteText}";
    }

    private static NaverBookingReservation ParseNaverReservationText(NaverBookingReservation item)
    {
        var rawText = item.RawText ?? "";
        if (string.IsNullOrWhiteSpace(rawText) || string.IsNullOrWhiteSpace(item.ReservationNo)) return new NaverBookingReservation();

        var statusName = Regex.Matches(
                rawText,
                @"(확정|신청|취소|완료|노쇼)(.+?)(01\d-\d{3,4}-\d{4})" + Regex.Escape(item.ReservationNo))
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
                .FirstOrDefault(tail.Contains) ?? "";
        }

        return new NaverBookingReservation
        {
            Status = statusName?.Groups[1].Value ?? "",
            BookerName = statusName is null ? "" : statusName.Groups[2].Value.Trim(),
            UseDateTime = useDateTime.Success ? useDateTime.Groups[1].Value.Trim() : "",
            ProductName = productName
        };
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

    private static string FirstMeaningfulLine(params string[] values) =>
        values.SelectMany(value => (value ?? "").Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
            .Select(line => line.Trim())
            .FirstOrDefault(line => !string.IsNullOrWhiteSpace(line) && !line.StartsWith(">", StringComparison.Ordinal)) ?? "";

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

    private sealed class NaverBookingSnapshot
    {
        public List<NaverBookingReservation> PendingConfirm { get; init; } = [];
        public List<NaverBookingReservation> TodayUse { get; init; } = [];
    }

    private sealed class NaverBookingReservation
    {
        public string ReservationNo { get; init; } = "";
        public string ProductName { get; init; } = "";
        public string BookerName { get; init; } = "";
        public string Status { get; init; } = "";
        public string UseDateTime { get; init; } = "";
        public string Purpose { get; init; } = "";
        public int Headcount { get; init; } = 1;
        public string RawText { get; init; } = "";
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
        grid.EnableHeadersVisualStyles = false;
        grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 243);
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(70, 66, 64);
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 9F, FontStyle.Bold);
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(235, 242, 235);
        grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(52, 48, 46);
        grid.GridColor = Color.FromArgb(232, 230, 228);
    }
}
