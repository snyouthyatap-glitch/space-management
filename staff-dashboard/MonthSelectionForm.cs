namespace YouthSpring.StaffDashboard;

internal sealed class MonthSelectionForm : Form
{
    private readonly DateTimePicker _monthPicker = new();

    public string SelectedMonth => _monthPicker.Value.ToString("yyyy-MM");

    public MonthSelectionForm()
    {
        Text = "실적 다운로드";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(360, 170);
        Font = new Font("맑은 고딕", 10F);
        BackColor = Color.White;

        var title = new Label
        {
            Text = "실적을 받을 월을 선택하세요.",
            Font = new Font("맑은 고딕", 13F, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(24, 22)
        };
        var hint = new Label
        {
            Text = "선택한 월의 전체 날짜가 엑셀로 정리됩니다.",
            ForeColor = Color.FromArgb(110, 106, 103),
            AutoSize = true,
            Location = new Point(26, 55)
        };
        _monthPicker.Format = DateTimePickerFormat.Custom;
        _monthPicker.CustomFormat = "yyyy년 MM월";
        _monthPicker.ShowUpDown = true;
        _monthPicker.MaxDate = DateTime.Today;
        _monthPicker.Value = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        _monthPicker.Location = new Point(28, 85);
        _monthPicker.Width = 145;

        var cancel = new Button
        {
            Text = "취소",
            DialogResult = DialogResult.Cancel,
            Size = new Size(72, 34),
            Location = new Point(184, 126),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.White
        };
        var confirm = new Button
        {
            Text = "다운로드",
            DialogResult = DialogResult.OK,
            Size = new Size(92, 34),
            Location = new Point(264, 126),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(55, 108, 142),
            ForeColor = Color.White
        };
        confirm.FlatAppearance.BorderSize = 0;

        Controls.AddRange([title, hint, _monthPicker, cancel, confirm]);
        AcceptButton = confirm;
        CancelButton = cancel;
    }
}
