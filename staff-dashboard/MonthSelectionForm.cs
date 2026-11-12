namespace YouthSpring.StaffDashboard;

// 실적 기간 선택: 월별(한 달) 또는 기간 지정(시작일~종료일, 최대 366일)
internal sealed class MonthSelectionForm : Form
{
    public const int MaxRangeDays = 366;

    private readonly RadioButton _monthMode = new();
    private readonly RadioButton _rangeMode = new();
    private readonly DateTimePicker _monthPicker = new();
    private readonly DateTimePicker _startPicker = new();
    private readonly DateTimePicker _endPicker = new();
    private readonly Label _rangeHint = new();

    public bool IsMonthMode => _monthMode.Checked;
    public string SelectedMonth => _monthPicker.Value.ToString("yyyy-MM");
    public string StartDate => _startPicker.Value.ToString("yyyy-MM-dd");
    public string EndDate => _endPicker.Value.ToString("yyyy-MM-dd");

    // 파일 이름·안내에 쓰는 기간 이름. 예) 2026-11, 2026-09-01~2026-11-12
    public string PeriodName => IsMonthMode ? SelectedMonth : $"{StartDate}~{EndDate}";

    public MonthSelectionForm()
    {
        Text = "실적 다운로드";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(420, 300);
        Font = new Font("맑은 고딕", 10F);
        BackColor = Color.White;

        var title = new Label
        {
            Text = "실적을 받을 기간을 선택하세요.",
            Font = new Font("맑은 고딕", 13F, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(24, 20)
        };

        _monthMode.Text = "월별";
        _monthMode.AutoSize = true;
        _monthMode.Checked = true;
        _monthMode.Location = new Point(28, 62);
        _monthMode.CheckedChanged += (_, _) => UpdateMode();

        _monthPicker.Format = DateTimePickerFormat.Custom;
        _monthPicker.CustomFormat = "yyyy년 MM월";
        _monthPicker.ShowUpDown = true;
        _monthPicker.MaxDate = DateTime.Today;
        _monthPicker.Value = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        _monthPicker.Location = new Point(48, 90);
        _monthPicker.Width = 160;

        _rangeMode.Text = "기간 지정";
        _rangeMode.AutoSize = true;
        _rangeMode.Location = new Point(28, 132);
        _rangeMode.CheckedChanged += (_, _) => UpdateMode();

        foreach (var picker in new[] { _startPicker, _endPicker })
        {
            picker.Format = DateTimePickerFormat.Custom;
            picker.CustomFormat = "yyyy-MM-dd";
            picker.MaxDate = DateTime.Today;
            picker.Width = 140;
            picker.ValueChanged += (_, _) => UpdateRangeHint();
        }
        _startPicker.Value = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        _endPicker.Value = DateTime.Today;
        _startPicker.Location = new Point(48, 160);
        _endPicker.Location = new Point(220, 160);
        var tilde = new Label { Text = "~", AutoSize = true, Location = new Point(197, 163) };

        _rangeHint.AutoSize = false;
        _rangeHint.Size = new Size(360, 22);
        _rangeHint.Location = new Point(48, 194);
        _rangeHint.ForeColor = Color.FromArgb(110, 106, 103);

        var cancel = new Button
        {
            Text = "취소",
            DialogResult = DialogResult.Cancel,
            Size = new Size(80, 40),
            Location = new Point(222, 244),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.White
        };
        var confirm = new Button
        {
            Text = "다운로드",
            DialogResult = DialogResult.OK,
            Size = new Size(100, 40),
            Location = new Point(308, 244),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(55, 108, 142),
            ForeColor = Color.White
        };
        confirm.FlatAppearance.BorderSize = 0;
        confirm.Click += (_, _) =>
        {
            var error = ValidateRange();
            if (error is null) return;
            MessageBox.Show(this, error, "기간 확인", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            DialogResult = DialogResult.None;
        };

        Controls.AddRange([title, _monthMode, _monthPicker, _rangeMode, _startPicker, tilde, _endPicker, _rangeHint, cancel, confirm]);
        AcceptButton = confirm;
        CancelButton = cancel;
        UpdateMode();
    }

    private void UpdateMode()
    {
        _monthPicker.Enabled = _monthMode.Checked;
        _startPicker.Enabled = _rangeMode.Checked;
        _endPicker.Enabled = _rangeMode.Checked;
        UpdateRangeHint();
    }

    private void UpdateRangeHint()
    {
        if (_monthMode.Checked)
        {
            _rangeHint.Text = "선택한 월의 전체 날짜가 엑셀로 정리됩니다.";
            return;
        }
        var days = (_endPicker.Value.Date - _startPicker.Value.Date).Days + 1;
        _rangeHint.Text = days > 0 ? $"{days}일 · 최대 {MaxRangeDays}일까지" : "종료일이 시작일보다 빠릅니다.";
    }

    private string? ValidateRange()
    {
        if (_monthMode.Checked) return null;
        var days = (_endPicker.Value.Date - _startPicker.Value.Date).Days + 1;
        if (days < 1) return "종료일이 시작일보다 빠릅니다.";
        if (days > MaxRangeDays) return $"기간은 최대 {MaxRangeDays}일까지 고를 수 있습니다.";
        return null;
    }
}
