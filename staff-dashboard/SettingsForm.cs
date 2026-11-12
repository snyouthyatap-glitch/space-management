namespace YouthSpring.StaffDashboard;

internal sealed class SettingsForm : Form
{
    private readonly TextBox _urlTextBox = new();
    private readonly TextBox _passwordTextBox = new();
    private readonly TextBox _naverIdTextBox = new();
    private readonly TextBox _naverPasswordTextBox = new();
    private readonly NumericUpDown _pollSeconds = new();
    private readonly Button _naverLoginSetupButton = new();
    private readonly Button _viewErrorLogButton = new();
    private readonly Func<Task<string?>>? _naverLoginSetupAsync;
    private readonly Func<string>? _errorLogProvider;

    public string AppsScriptUrl => _urlTextBox.Text.Trim();
    public string Password => _passwordTextBox.Text;
    public string NaverId => _naverIdTextBox.Text.Trim();
    public string NaverPassword => _naverPasswordTextBox.Text;
    public int PollSeconds => (int)_pollSeconds.Value;

    public SettingsForm(
        DashboardSettings settings,
        string currentPassword,
        string currentNaverId,
        string currentNaverPassword,
        Func<Task<string?>>? naverLoginSetupAsync = null,
        Func<string>? errorLogProvider = null)
    {
        _naverLoginSetupAsync = naverLoginSetupAsync;
        _errorLogProvider = errorLogProvider;

        Text = "연결 설정";
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(560, 450);
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

        Controls.Add(CreateLabel("관리자 비밀번호", 30, 136));
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

        Controls.Add(CreateLabel("네이버 아이디", 30, 202));
        _naverIdTextBox.SetBounds(30, 226, 235, 30);
        _naverIdTextBox.Text = currentNaverId;
        Controls.Add(_naverIdTextBox);

        Controls.Add(CreateLabel("네이버 비밀번호", 295, 202));
        _naverPasswordTextBox.SetBounds(295, 226, 235, 30);
        _naverPasswordTextBox.UseSystemPasswordChar = true;
        _naverPasswordTextBox.Text = currentNaverPassword;
        Controls.Add(_naverPasswordTextBox);

        var naverNotice = new Label
        {
            Text = "네이버 계정은 이 Windows 사용자 계정에서만 복호화되도록 저장됩니다. " +
                   "자동 로그인이 막히면(추가 인증 등) '네이버 로그인 설정'으로 직접 한 번 로그인해 두세요.",
            AutoSize = false,
            Location = new Point(30, 264),
            Size = new Size(500, 44),
            ForeColor = Color.FromArgb(105, 100, 97)
        };
        Controls.Add(naverNotice);

        _naverLoginSetupButton.Text = "네이버 로그인 설정";
        _naverLoginSetupButton.FlatStyle = FlatStyle.Flat;
        _naverLoginSetupButton.BackColor = Color.FromArgb(44, 126, 75);
        _naverLoginSetupButton.ForeColor = Color.White;
        _naverLoginSetupButton.FlatAppearance.BorderSize = 0;
        _naverLoginSetupButton.Cursor = Cursors.Hand;
        _naverLoginSetupButton.SetBounds(30, 318, 170, 40);
        _naverLoginSetupButton.Enabled = _naverLoginSetupAsync is not null;
        _naverLoginSetupButton.Click += async (_, _) => await RunNaverLoginSetupAsync();
        Controls.Add(_naverLoginSetupButton);

        _viewErrorLogButton.Text = "오류 로그 보기";
        _viewErrorLogButton.FlatStyle = FlatStyle.Flat;
        _viewErrorLogButton.BackColor = Color.White;
        _viewErrorLogButton.ForeColor = Color.FromArgb(52, 48, 46);
        _viewErrorLogButton.FlatAppearance.BorderColor = Color.FromArgb(214, 210, 206);
        _viewErrorLogButton.Cursor = Cursors.Hand;
        _viewErrorLogButton.SetBounds(210, 318, 140, 40);
        _viewErrorLogButton.Enabled = _errorLogProvider is not null;
        _viewErrorLogButton.Click += (_, _) => ShowErrorLogDialog();
        Controls.Add(_viewErrorLogButton);

        var cancelButton = new Button
        {
            Text = "취소",
            DialogResult = DialogResult.Cancel,
            FlatStyle = FlatStyle.Flat
        };
        cancelButton.SetBounds(350, 390, 85, 40);

        var saveButton = new Button
        {
            Text = "저장",
            DialogResult = DialogResult.OK,
            BackColor = Color.FromArgb(233, 63, 111),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat
        };
        saveButton.FlatAppearance.BorderSize = 0;
        saveButton.SetBounds(445, 390, 85, 40);
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

    private async Task RunNaverLoginSetupAsync()
    {
        if (_naverLoginSetupAsync is null) return;

        var originalText = _naverLoginSetupButton.Text;
        try
        {
            UseWaitCursor = true;
            _naverLoginSetupButton.Enabled = false;
            _naverLoginSetupButton.Text = "여는 중...";

            var errorMessage = await _naverLoginSetupAsync();
            if (string.IsNullOrWhiteSpace(errorMessage))
            {
                MessageBox.Show(
                    this,
                    "네이버 로그인용 크롬 창을 열었습니다.\n\n" +
                    "1. 열린 창에서 네이버에 로그인하고 예약자관리 화면이 보이는지 확인해 주세요.\n" +
                    "2. 로그인이 끝나면 그 크롬 창을 닫아 주세요. 창이 열려 있으면 자동 확인이 실패합니다.\n" +
                    "3. 직원 프로그램에서 '새로고침'을 누르면 바로 다시 확인합니다.",
                    "네이버 로그인 설정",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            MessageBox.Show(this, errorMessage, "네이버 로그인 설정", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        finally
        {
            UseWaitCursor = false;
            _naverLoginSetupButton.Enabled = _naverLoginSetupAsync is not null;
            _naverLoginSetupButton.Text = originalText;
        }
    }

    private void ShowErrorLogDialog()
    {
        var text = _errorLogProvider?.Invoke();
        using var dialog = new Form
        {
            Text = "오류 로그",
            StartPosition = FormStartPosition.CenterParent,
            Size = new Size(860, 560),
            MinimumSize = new Size(700, 420),
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
            Padding = new Padding(16)
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        dialog.Controls.Add(root);

        root.Controls.Add(new Label
        {
            Text = "최근 네이버 확인·로그인 기록 (최신이 아래)",
            AutoSize = true,
            Font = new Font(Font, FontStyle.Bold),
            Padding = new Padding(0, 0, 0, 8)
        }, 0, 0);

        var logBox = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(250, 250, 249),
            Font = new Font("Consolas", 9.5F),
            Text = string.IsNullOrWhiteSpace(text) ? "아직 기록이 없습니다." : text
        };
        root.Controls.Add(logBox, 0, 1);

        var closeButton = new Button
        {
            Text = "닫기",
            DialogResult = DialogResult.OK,
            Size = new Size(88, 40),
            Anchor = AnchorStyles.Right,
            Margin = new Padding(0, 10, 0, 0)
        };
        root.Controls.Add(closeButton, 0, 2);
        dialog.AcceptButton = closeButton;
        dialog.Shown += (_, _) =>
        {
            logBox.SelectionStart = logBox.TextLength;
            logBox.ScrollToCaret();
        };
        dialog.ShowDialog(this);
    }

    private static Label CreateLabel(string text, int x, int y) => new()
    {
        Text = text,
        AutoSize = true,
        Location = new Point(x, y),
        ForeColor = Color.FromArgb(70, 66, 64)
    };
}
