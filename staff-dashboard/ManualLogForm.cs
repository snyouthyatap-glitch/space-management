using System.Text.RegularExpressions;

namespace YouthSpring.StaffDashboard;

internal sealed class ManualLogForm : Form
{
    private static readonly Color Ink = Color.FromArgb(52, 48, 46);
    private static readonly Color Muted = Color.FromArgb(105, 100, 97);
    private static readonly Color Green = Color.FromArgb(74, 156, 68);
    private static readonly Color Pink = Color.FromArgb(233, 63, 111);

    private readonly AppsScriptClient _client;
    private readonly DashboardSettings _settings;
    private readonly string _password;
    private readonly RadioButton _memberMode = new();
    private readonly RadioButton _guestMode = new();
    private readonly Panel _memberPanel = new();
    private readonly Panel _guestPanel = new();
    private readonly TextBox _phone = new();
    private readonly Button _lookupButton = new();
    private readonly Label _memberSummary = new();
    private readonly Button _changeMemberButton = new();
    private readonly ComboBox _guestGender = new();
    private readonly NumericUpDown _guestAge = new();
    private readonly DateTimePicker _usageDate = new();
    private readonly ComboBox _logType = new();
    private readonly Panel _printerPanel = new();
    private readonly Panel _roomPanel = new();
    private readonly NumericUpDown _printerCount = new();
    private readonly DateTimePicker _startTime = new();
    private readonly DateTimePicker _endTime = new();
    private readonly TextBox _purpose = new();
    private readonly NumericUpDown _companionCount = new();
    private readonly DataGridView _companionGrid = new();
    private readonly Button _submitButton = new();
    private readonly Label _statusLabel = new();
    private readonly DataGridView _recentGrid = new();

    private AdminMember? _member;
    private string _guestToken = NewRequestId();
    private string _pendingPrinterSignature = "";
    private string _pendingPrinterRequestId = "";
    private bool _normalizingPhone;
    private bool _busy;

    public ManualLogForm(
        AppsScriptClient client,
        DashboardSettings settings,
        string password,
        string previewPath = "")
    {
        _client = client;
        _settings = settings;
        _password = password;

        Text = "청년이봄 수동 일지";
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(920, 790);
        MinimumSize = new Size(820, 700);
        Font = new Font("맑은 고딕", 9.5F);
        BackColor = Color.FromArgb(247, 247, 246);
        Icon = SystemIcons.Application;

        BuildLayout();
        SetSubjectMode(false);
        UpdateLogTypeFields();
        if (!string.IsNullOrWhiteSpace(previewPath))
        {
            Shown += async (_, _) =>
            {
                _guestMode.Checked = true;
                _logType.SelectedIndex = 4;
                _companionCount.Value = 3;
                await Task.Delay(400);
                using var bitmap = new Bitmap(Width, Height);
                DrawToBitmap(bitmap, new Rectangle(Point.Empty, Size));
                bitmap.Save(previewPath, System.Drawing.Imaging.ImageFormat.Png);
                Close();
            };
        }
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            RowCount = 3,
            ColumnCount = 1,
            Padding = new Padding(20),
            BackColor = BackColor
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 62));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
        Controls.Add(root);

        var header = new Panel { Dock = DockStyle.Fill };
        header.Controls.Add(new Label
        {
            Text = "수동 일지 제출",
            Font = new Font("맑은 고딕", 19F, FontStyle.Bold),
            ForeColor = Ink,
            AutoSize = true,
            Location = new Point(2, 1)
        });
        header.Controls.Add(new Label
        {
            Text = DateTime.Now.ToString("yyyy.MM.dd"),
            ForeColor = Muted,
            AutoSize = true,
            Location = new Point(5, 39)
        });
        root.Controls.Add(header, 0, 0);

        var tabs = new TabControl { Dock = DockStyle.Fill };
        var submitTab = new TabPage("일지 제출")
        {
            BackColor = Color.White,
            Padding = new Padding(16)
        };
        var recentTab = new TabPage("이번 실행 내역")
        {
            BackColor = Color.White,
            Padding = new Padding(12)
        };
        tabs.TabPages.Add(submitTab);
        tabs.TabPages.Add(recentTab);
        root.Controls.Add(tabs, 0, 1);

        var submitLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2
        };
        submitLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 174));
        submitLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        submitTab.Controls.Add(submitLayout);
        submitLayout.Controls.Add(BuildSubjectGroup(), 0, 0);
        submitLayout.Controls.Add(BuildLogGroup(), 0, 1);

        ConfigureRecentGrid();
        recentTab.Controls.Add(_recentGrid);

        _statusLabel.Dock = DockStyle.Fill;
        _statusLabel.Text = "준비";
        _statusLabel.ForeColor = Muted;
        _statusLabel.TextAlign = ContentAlignment.MiddleLeft;
        root.Controls.Add(_statusLabel, 0, 2);
    }

    private Control BuildSubjectGroup()
    {
        var group = new GroupBox
        {
            Text = "1. 이용자",
            Dock = DockStyle.Fill,
            ForeColor = Ink,
            Padding = new Padding(14)
        };

        _memberMode.Text = "회원 조회";
        _memberMode.AutoSize = true;
        _memberMode.Location = new Point(20, 28);
        _memberMode.CheckedChanged += (_, _) =>
        {
            if (_memberMode.Checked) SetSubjectMode(false);
        };
        group.Controls.Add(_memberMode);

        _guestMode.Text = "비회원";
        _guestMode.AutoSize = true;
        _guestMode.Location = new Point(120, 28);
        _guestMode.CheckedChanged += (_, _) =>
        {
            if (_guestMode.Checked) SetSubjectMode(true);
        };
        group.Controls.Add(_guestMode);

        _memberPanel.SetBounds(20, 58, 820, 94);
        _memberPanel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        ConfigureMemberPanel();
        group.Controls.Add(_memberPanel);

        _guestPanel.SetBounds(20, 58, 820, 94);
        _guestPanel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        ConfigureGuestPanel();
        group.Controls.Add(_guestPanel);
        return group;
    }

    private void ConfigureMemberPanel()
    {
        _memberPanel.Controls.Add(CreateFieldLabel("휴대폰 번호 전체", 0, 0));
        _phone.SetBounds(0, 24, 320, 32);
        _phone.MaxLength = 11;
        _phone.PlaceholderText = "01012341234";
        _phone.TextChanged += (_, _) => NormalizePhoneInput();
        _phone.KeyDown += async (_, eventArgs) =>
        {
            if (eventArgs.KeyCode != Keys.Enter) return;
            eventArgs.SuppressKeyPress = true;
            await LookupMemberAsync();
        };
        _memberPanel.Controls.Add(_phone);

        _lookupButton.Text = "회원 찾기";
        _lookupButton.SetBounds(330, 23, 104, 34);
        StyleCommandButton(_lookupButton, Ink);
        _lookupButton.Click += async (_, _) => await LookupMemberAsync();
        _memberPanel.Controls.Add(_lookupButton);

        _memberSummary.SetBounds(0, 22, 650, 40);
        _memberSummary.Font = new Font(Font, FontStyle.Bold);
        _memberSummary.ForeColor = Green;
        _memberSummary.TextAlign = ContentAlignment.MiddleLeft;
        _memberSummary.Visible = false;
        _memberPanel.Controls.Add(_memberSummary);

        _changeMemberButton.Text = "다른 회원 조회";
        _changeMemberButton.SetBounds(660, 25, 128, 32);
        _changeMemberButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        StyleSecondaryButton(_changeMemberButton);
        _changeMemberButton.Visible = false;
        _changeMemberButton.Click += (_, _) => ResetMember();
        _memberPanel.Controls.Add(_changeMemberButton);
    }

    private void ConfigureGuestPanel()
    {
        _guestPanel.Controls.Add(CreateFieldLabel("성별", 0, 0));
        _guestGender.SetBounds(0, 24, 190, 32);
        _guestGender.DropDownStyle = ComboBoxStyle.DropDownList;
        _guestGender.Items.AddRange(["선택", "남성", "여성"]);
        _guestGender.SelectedIndex = 0;
        _guestPanel.Controls.Add(_guestGender);

        _guestPanel.Controls.Add(CreateFieldLabel("만 나이", 212, 0));
        _guestAge.SetBounds(212, 24, 140, 32);
        _guestAge.Minimum = 1;
        _guestAge.Maximum = 100;
        _guestAge.Value = 20;
        _guestPanel.Controls.Add(_guestAge);
    }

    private Control BuildLogGroup()
    {
        var group = new GroupBox
        {
            Text = "2. 이용 기록",
            Dock = DockStyle.Fill,
            ForeColor = Ink,
            Padding = new Padding(14)
        };

        group.Controls.Add(CreateFieldLabel("이용 날짜", 20, 28));
        _usageDate.SetBounds(20, 51, 220, 32);
        _usageDate.Format = DateTimePickerFormat.Custom;
        _usageDate.CustomFormat = "yyyy-MM-dd";
        _usageDate.MaxDate = DateTime.Today;
        _usageDate.Value = DateTime.Today;
        group.Controls.Add(_usageDate);

        group.Controls.Add(CreateFieldLabel("일지 종류", 262, 28));
        _logType.SetBounds(262, 51, 240, 32);
        _logType.DropDownStyle = ComboBoxStyle.DropDownList;
        _logType.DisplayMember = nameof(LogTypeChoice.Label);
        _logType.Items.AddRange(
        [
            new LogTypeChoice("visit", "방문일지"),
            new LogTypeChoice("printer", "프린터"),
            new LogTypeChoice("careerRoom1", "커리어존 Room1"),
            new LogTypeChoice("careerRoom2", "커리어존 Room2"),
            new LogTypeChoice("connectRoom", "커넥트룸")
        ]);
        _logType.SelectedIndex = 0;
        _logType.SelectedIndexChanged += (_, _) => UpdateLogTypeFields();
        group.Controls.Add(_logType);

        var conditionalHost = new Panel
        {
            Location = new Point(20, 98),
            Size = new Size(820, 260),
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        };
        ConfigurePrinterPanel();
        ConfigureRoomPanel();
        conditionalHost.Controls.Add(_printerPanel);
        conditionalHost.Controls.Add(_roomPanel);
        group.Controls.Add(conditionalHost);

        _submitButton.Text = "일지 제출";
        _submitButton.Size = new Size(150, 38);
        _submitButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        _submitButton.Location = new Point(690, 374);
        StyleCommandButton(_submitButton, Green);
        _submitButton.Click += async (_, _) => await SubmitManualLogAsync();
        group.Controls.Add(_submitButton);
        group.Resize += (_, _) =>
        {
            _submitButton.Location = new Point(
                group.ClientSize.Width - _submitButton.Width - 20,
                group.ClientSize.Height - _submitButton.Height - 18);
        };
        return group;
    }

    private void ConfigurePrinterPanel()
    {
        _printerPanel.Dock = DockStyle.Fill;
        _printerPanel.Controls.Add(CreateFieldLabel("사용 매수", 0, 0));
        _printerCount.SetBounds(0, 24, 150, 32);
        _printerCount.Minimum = 1;
        _printerCount.Maximum = 10;
        _printerCount.Value = 1;
        _printerPanel.Controls.Add(_printerCount);
    }

    private void ConfigureRoomPanel()
    {
        _roomPanel.Dock = DockStyle.Fill;
        _roomPanel.Controls.Add(CreateFieldLabel("시작 시간", 0, 0));
        ConfigureTimePicker(_startTime, 0, 24, new TimeSpan(9, 0, 0));
        _roomPanel.Controls.Add(_startTime);

        _roomPanel.Controls.Add(CreateFieldLabel("종료 시간", 172, 0));
        ConfigureTimePicker(_endTime, 172, 24, new TimeSpan(10, 0, 0));
        _roomPanel.Controls.Add(_endTime);

        _roomPanel.Controls.Add(CreateFieldLabel("이용 목적", 354, 0));
        _purpose.SetBounds(354, 24, 430, 32);
        _purpose.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        _purpose.MaxLength = 200;
        _roomPanel.Controls.Add(_purpose);

        _roomPanel.Controls.Add(CreateFieldLabel("추가 인원", 0, 69));
        _companionCount.SetBounds(0, 92, 120, 32);
        _companionCount.Minimum = 0;
        _companionCount.Maximum = 10;
        _companionCount.ValueChanged += (_, _) => RenderCompanionRows();
        _roomPanel.Controls.Add(_companionCount);

        ConfigureCompanionGrid();
        _companionGrid.SetBounds(145, 69, 639, 170);
        _companionGrid.Anchor =
            AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        _roomPanel.Controls.Add(_companionGrid);
    }

    private void ConfigureCompanionGrid()
    {
        _companionGrid.BackgroundColor = Color.White;
        _companionGrid.BorderStyle = BorderStyle.FixedSingle;
        _companionGrid.AllowUserToAddRows = false;
        _companionGrid.AllowUserToDeleteRows = false;
        _companionGrid.AllowUserToResizeRows = false;
        _companionGrid.RowHeadersVisible = false;
        _companionGrid.AutoGenerateColumns = false;
        _companionGrid.SelectionMode = DataGridViewSelectionMode.CellSelect;
        _companionGrid.RowTemplate.Height = 32;
        _companionGrid.ColumnHeadersHeight = 34;
        _companionGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "number",
            HeaderText = "추가 인원",
            ReadOnly = true,
            Width = 100
        });
        _companionGrid.Columns.Add(new DataGridViewComboBoxColumn
        {
            Name = "gender",
            HeaderText = "성별",
            DataSource = new[] { "남성", "여성" },
            Width = 220,
            FlatStyle = FlatStyle.Flat
        });
        _companionGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "age",
            HeaderText = "만 나이",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        });
    }

    private void ConfigureRecentGrid()
    {
        ConfigureReadOnlyGrid(_recentGrid);
        _recentGrid.Columns.Add("time", "시간");
        _recentGrid.Columns.Add("subject", "이용자");
        _recentGrid.Columns.Add("date", "날짜");
        _recentGrid.Columns.Add("type", "종류");
        _recentGrid.Columns.Add("result", "처리");
        _recentGrid.Columns["time"]!.Width = 90;
        _recentGrid.Columns["subject"]!.Width = 150;
        _recentGrid.Columns["date"]!.Width = 120;
        _recentGrid.Columns["type"]!.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        _recentGrid.Columns["result"]!.Width = 130;
    }

    private async Task LookupMemberAsync()
    {
        if (_busy) return;
        var phone = Regex.Replace(_phone.Text, @"\D", "");
        if (!Regex.IsMatch(phone, @"^010\d{8}$"))
        {
            SetStatus("휴대폰 번호 11자리를 정확히 입력해 주세요.", true);
            return;
        }

        SetBusy(true, "조회 중...");
        try
        {
            using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            _member = await _client.LookupMemberAsync(
                _settings,
                _password,
                phone,
                cancellation.Token);
            _phone.Visible = false;
            _lookupButton.Visible = false;
            _memberSummary.Text =
                $"{_member.Name} · {_member.Gender} · " +
                (string.IsNullOrWhiteSpace(_member.BirthDate)
                    ? "생년월일 미등록"
                    : $"만 {_member.Age}세") +
                $" · {FormatPhone(_member.Phone)}";
            _memberSummary.Visible = true;
            _changeMemberButton.Visible = true;
            UpdateSubmitEnabled();
            SetStatus("회원을 확인했습니다.");
        }
        catch (Exception exception)
        {
            SetStatus(exception.Message, true);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async Task SubmitManualLogAsync()
    {
        if (_busy) return;
        try
        {
            var request = BuildManualLogRequest();
            SetBusy(true, "저장 중...");
            using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            var result = await _client.SubmitManualLogAsync(
                _settings,
                _password,
                request,
                cancellation.Token);

            if (request.LogType == "printer")
            {
                _pendingPrinterSignature = "";
                _pendingPrinterRequestId = "";
            }
            AddRecentRow(request, result);
            ResetLogDetails();
            if (request.SubjectType == "guest")
            {
                _guestToken = NewRequestId();
                _guestGender.SelectedIndex = 0;
                _guestAge.Value = 20;
            }

            SetStatus(result.LinkedReservation
                ? "네이버 예약과 연결하여 일지를 저장했습니다."
                : result.LogCreated
                    ? "수동 일지를 저장했습니다."
                    : "이미 등록된 기록을 확인했습니다.");
        }
        catch (Exception exception)
        {
            SetStatus(exception.Message, true);
            MessageBox.Show(
                this,
                exception.Message,
                "수동 일지 저장 실패",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private ManualLogRequest BuildManualLogRequest()
    {
        var guest = _guestMode.Checked;
        if (!guest && _member is null) throw new InvalidOperationException("회원을 먼저 조회해 주세요.");
        if (guest && _guestGender.SelectedItem?.ToString() is not ("남성" or "여성"))
            throw new InvalidOperationException("비회원 성별을 선택해 주세요.");

        var logType = SelectedLogType();
        var date = _usageDate.Value.ToString("yyyy-MM-dd");
        var data = new ManualLogData();
        var requestId = NewRequestId();

        if (logType == "printer")
        {
            data.Count = (int)_printerCount.Value;
            var subjectKey = guest ? _guestToken : _member!.Id;
            var signature = $"{subjectKey}:{date}:{data.Count}";
            if (_pendingPrinterSignature != signature ||
                string.IsNullOrWhiteSpace(_pendingPrinterRequestId))
            {
                _pendingPrinterSignature = signature;
                _pendingPrinterRequestId = NewRequestId();
            }
            requestId = _pendingPrinterRequestId;
        }
        else if (IsRoomType(logType))
        {
            data.StartTime = _startTime.Value.ToString("HH:mm");
            data.EndTime = _endTime.Value.ToString("HH:mm");
            data.Purpose = _purpose.Text.Trim();
            if (_endTime.Value.TimeOfDay <= _startTime.Value.TimeOfDay)
                throw new InvalidOperationException("종료 시간은 시작 시간보다 늦어야 합니다.");
            if (string.IsNullOrWhiteSpace(data.Purpose))
                throw new InvalidOperationException("이용 목적을 입력해 주세요.");
            data.Companions = CollectCompanions();
        }

        return new ManualLogRequest
        {
            SubjectType = guest ? "guest" : "member",
            MemberId = guest ? "" : _member!.Id,
            Guest = guest
                ? new ManualGuest
                {
                    Token = _guestToken,
                    Gender = _guestGender.SelectedItem!.ToString()!,
                    Age = (int)_guestAge.Value
                }
                : null,
            LogType = logType,
            Date = date,
            Data = data,
            RequestId = requestId
        };
    }

    private List<ManualCompanion> CollectCompanions()
    {
        _companionGrid.EndEdit();
        var result = new List<ManualCompanion>();
        foreach (DataGridViewRow row in _companionGrid.Rows)
        {
            var gender = row.Cells["gender"].Value?.ToString() ?? "";
            var ageText = row.Cells["age"].Value?.ToString() ?? "";
            if (gender is not ("남성" or "여성") ||
                !int.TryParse(ageText, out var age) ||
                age is < 1 or > 100)
            {
                throw new InvalidOperationException("추가 인원의 성별과 만 나이를 모두 입력해 주세요.");
            }
            result.Add(new ManualCompanion { Gender = gender, Age = age });
        }
        return result;
    }

    private void AddRecentRow(ManualLogRequest request, ManualLogResult result)
    {
        var subject = request.SubjectType == "guest" ? "비회원" : _member?.Name ?? "회원";
        var resultText = result.LinkedReservation
            ? "예약 연결"
            : result.LogCreated ? "저장 완료" : "기존 기록";
        _recentGrid.Rows.Insert(
            0,
            DateTime.Now.ToString("HH:mm"),
            subject,
            request.Date,
            LogTypeLabel(request.LogType),
            resultText);
        const int rowIndex = 0;
        _recentGrid.Rows[rowIndex].Cells["result"].Style.ForeColor = Green;
        _recentGrid.Rows[rowIndex].Cells["result"].Style.Font =
            new Font(_recentGrid.Font, FontStyle.Bold);
    }

    private void SetSubjectMode(bool guest)
    {
        if (_guestMode.Checked != guest) _guestMode.Checked = guest;
        if (_memberMode.Checked == guest) _memberMode.Checked = !guest;
        _guestPanel.Visible = guest;
        _memberPanel.Visible = !guest;
        _pendingPrinterSignature = "";
        _pendingPrinterRequestId = "";
        if (guest) _guestToken = NewRequestId();
        UpdateSubmitEnabled();
        SetStatus("준비");
    }

    private void ResetMember()
    {
        _member = null;
        _phone.Text = "";
        _phone.Visible = true;
        _lookupButton.Visible = true;
        _memberSummary.Visible = false;
        _changeMemberButton.Visible = false;
        UpdateSubmitEnabled();
        _phone.Focus();
        SetStatus("준비");
    }

    private void UpdateLogTypeFields()
    {
        var logType = SelectedLogType();
        _printerPanel.Visible = logType == "printer";
        _roomPanel.Visible = IsRoomType(logType);
        _pendingPrinterSignature = "";
        _pendingPrinterRequestId = "";
    }

    private void RenderCompanionRows()
    {
        var previous = _companionGrid.Rows
            .Cast<DataGridViewRow>()
            .Select(row => new
            {
                Gender = row.Cells["gender"].Value,
                Age = row.Cells["age"].Value
            })
            .ToList();
        _companionGrid.Rows.Clear();
        for (var index = 0; index < (int)_companionCount.Value; index++)
        {
            var rowIndex = _companionGrid.Rows.Add();
            _companionGrid.Rows[rowIndex].Cells["number"].Value = index + 1;
            if (index >= previous.Count) continue;
            _companionGrid.Rows[rowIndex].Cells["gender"].Value = previous[index].Gender;
            _companionGrid.Rows[rowIndex].Cells["age"].Value = previous[index].Age;
        }
    }

    private void ResetLogDetails()
    {
        _printerCount.Value = 1;
        _purpose.Text = "";
        _companionCount.Value = 0;
        _startTime.Value = DateTime.Today.AddHours(9);
        _endTime.Value = DateTime.Today.AddHours(10);
    }

    private void SetBusy(bool busy, string buttonText = "")
    {
        _busy = busy;
        _lookupButton.Enabled = !busy;
        _submitButton.Enabled = !busy && (_guestMode.Checked || _member is not null);
        _memberMode.Enabled = !busy;
        _guestMode.Enabled = !busy;
        _submitButton.Text = busy && !string.IsNullOrWhiteSpace(buttonText)
            ? buttonText
            : "일지 제출";
    }

    private void UpdateSubmitEnabled() =>
        _submitButton.Enabled = !_busy && (_guestMode.Checked || _member is not null);

    private void SetStatus(string message, bool error = false)
    {
        _statusLabel.Text = message;
        _statusLabel.ForeColor = error ? Color.FromArgb(190, 55, 55) : Muted;
    }

    private void NormalizePhoneInput()
    {
        if (_normalizingPhone) return;
        var digits = Regex.Replace(_phone.Text, @"\D", "");
        if (digits.Length > 11) digits = digits[..11];
        if (_phone.Text == digits) return;
        _normalizingPhone = true;
        _phone.Text = digits;
        _phone.SelectionStart = digits.Length;
        _normalizingPhone = false;
    }

    private string SelectedLogType() =>
        (_logType.SelectedItem as LogTypeChoice)?.Value ?? "visit";

    private static bool IsRoomType(string value) =>
        value is "careerRoom1" or "careerRoom2" or "connectRoom";

    private static string LogTypeLabel(string value) => value switch
    {
        "visit" => "방문일지",
        "printer" => "프린터",
        "careerRoom1" => "커리어존 Room1",
        "careerRoom2" => "커리어존 Room2",
        "connectRoom" => "커넥트룸",
        _ => value
    };

    private static string FormatPhone(string value)
    {
        var digits = Regex.Replace(value, @"\D", "");
        return digits.Length == 11
            ? $"{digits[..3]}-{digits.Substring(3, 4)}-{digits[7..]}"
            : digits;
    }

    private static string NewRequestId() => Guid.NewGuid().ToString("N");

    private static Label CreateFieldLabel(string text, int x, int y) => new()
    {
        Text = text,
        AutoSize = true,
        ForeColor = Muted,
        Font = new Font("맑은 고딕", 9F, FontStyle.Bold),
        Location = new Point(x, y)
    };

    private static void ConfigureTimePicker(
        DateTimePicker picker,
        int x,
        int y,
        TimeSpan initialTime)
    {
        picker.SetBounds(x, y, 150, 32);
        picker.Format = DateTimePickerFormat.Custom;
        picker.CustomFormat = "HH:mm";
        picker.ShowUpDown = true;
        picker.Value = DateTime.Today.Add(initialTime);
    }

    private static void StyleCommandButton(Button button, Color color)
    {
        button.BackColor = color;
        button.ForeColor = Color.White;
        button.FlatStyle = FlatStyle.Flat;
        button.Cursor = Cursors.Hand;
        button.FlatAppearance.BorderSize = 0;
    }

    private static void StyleSecondaryButton(Button button)
    {
        button.BackColor = Color.White;
        button.ForeColor = Ink;
        button.FlatStyle = FlatStyle.Flat;
        button.Cursor = Cursors.Hand;
        button.FlatAppearance.BorderColor = Color.FromArgb(210, 205, 201);
    }

    private static void ConfigureReadOnlyGrid(DataGridView grid)
    {
        grid.Dock = DockStyle.Fill;
        grid.BackgroundColor = Color.White;
        grid.BorderStyle = BorderStyle.None;
        grid.ReadOnly = true;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.AllowUserToResizeRows = false;
        grid.RowHeadersVisible = false;
        grid.AutoGenerateColumns = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.ColumnHeadersHeight = 38;
        grid.RowTemplate.Height = 36;
    }

    private sealed record LogTypeChoice(string Value, string Label);
}
