namespace YouthSpring.StaffDashboard;

internal sealed class SettingsForm : Form
{
    private readonly TextBox _urlTextBox = new();
    private readonly TextBox _passwordTextBox = new();
    private readonly NumericUpDown _pollSeconds = new();

    public string AppsScriptUrl => _urlTextBox.Text.Trim();
    public string Password => _passwordTextBox.Text;
    public int PollSeconds => (int)_pollSeconds.Value;

    public SettingsForm(DashboardSettings settings, string currentPassword)
    {
        Text = "연결 설정";
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(560, 270);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Font = new Font("맑은 고딕", 10F);
        BackColor = Color.White;

        var title = new Label
        {
            Text = "직원 알림 프로그램 연결",
            Font = new Font("맑은 고딕", 16F, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(28, 22)
        };
        Controls.Add(title);

        Controls.Add(CreateLabel("Apps Script 웹앱 주소", 30, 70));
        _urlTextBox.SetBounds(30, 94, 500, 30);
        _urlTextBox.Text = settings.AppsScriptUrl;
        Controls.Add(_urlTextBox);

        Controls.Add(CreateLabel("웹 관리자 비밀번호", 30, 136));
        _passwordTextBox.SetBounds(30, 160, 330, 30);
        _passwordTextBox.UseSystemPasswordChar = true;
        _passwordTextBox.Text = currentPassword;
        Controls.Add(_passwordTextBox);

        Controls.Add(CreateLabel("확인 간격(초)", 380, 136));
        _pollSeconds.SetBounds(380, 160, 150, 30);
        _pollSeconds.Minimum = 15;
        _pollSeconds.Maximum = 120;
        _pollSeconds.Value = Math.Clamp(settings.PollSeconds, 15, 120);
        Controls.Add(_pollSeconds);

        var cancelButton = new Button
        {
            Text = "취소",
            DialogResult = DialogResult.Cancel,
            FlatStyle = FlatStyle.Flat
        };
        cancelButton.SetBounds(350, 215, 85, 34);

        var saveButton = new Button
        {
            Text = "저장",
            DialogResult = DialogResult.OK,
            BackColor = Color.FromArgb(233, 63, 111),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        saveButton.FlatAppearance.BorderSize = 0;
        saveButton.SetBounds(445, 215, 85, 34);
        saveButton.Click += (sender, eventArgs) =>
        {
            if (!Uri.TryCreate(AppsScriptUrl, UriKind.Absolute, out _) ||
                !AppsScriptUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(this, "올바른 HTTPS Apps Script 주소를 입력해 주세요.", "주소 확인",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }
            if (Password.Length < 8)
            {
                MessageBox.Show(this, "관리자 비밀번호를 8자 이상 입력해 주세요.", "비밀번호 확인",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
            }
        };

        Controls.Add(cancelButton);
        Controls.Add(saveButton);
        AcceptButton = saveButton;
        CancelButton = cancelButton;
    }

    private static Label CreateLabel(string text, int x, int y) => new()
    {
        Text = text,
        AutoSize = true,
        Location = new Point(x, y),
        ForeColor = Color.FromArgb(70, 66, 64)
    };
}
