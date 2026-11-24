namespace YouthSpring.StaffDashboard;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Any(argument => argument.Equals("--self-test", StringComparison.OrdinalIgnoreCase)))
        {
            SelfTest.Run();
            return 0;
        }

        ApplicationConfiguration.Initialize();

        // 자동 업데이트로 다시 켜진 경우: 이전 프로그램이 끝날 때까지 기다린 뒤 실행 잠금을 잡는다.
        var justUpdated = AppUpdater.WaitForPreviousProcess(args);
        AppUpdater.CleanupOldCopy(Environment.ProcessPath);

        // 가상 시트 데모: --virtual-sheet=http://127.0.0.1:8787/exec (demo/demo-server.js)
        var virtualSheetUrl = args
            .FirstOrDefault(argument => argument.StartsWith("--virtual-sheet=", StringComparison.OrdinalIgnoreCase))
            ?.Substring("--virtual-sheet=".Length) ?? "";
        if (!string.IsNullOrWhiteSpace(virtualSheetUrl) && !IsLoopback(virtualSheetUrl))
        {
            MessageBox.Show("가상 시트 데모 주소는 내 컴퓨터(127.0.0.1) 주소만 쓸 수 있습니다.", "가상 시트 데모");
            return 1;
        }

        // 데모는 실제 프로그램과 동시에 켤 수 있게 실행 잠금을 따로 쓴다.
        using var singleInstance = new Mutex(
            true,
            string.IsNullOrWhiteSpace(virtualSheetUrl)
                ? @"Local\YouthSpringStaffDashboard"
                : @"Local\YouthSpringStaffDashboard-VirtualDemo",
            out var isFirstInstance);
        if (!isFirstInstance)
        {
            if (!args.Any(argument =>
                    argument.StartsWith("--preview=", StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show(
                    "청년이봄 직원 알림 프로그램이 이미 실행 중입니다.\n작업 표시줄의 알림 영역을 확인해 주세요.",
                    "이미 실행 중",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            return 1;
        }

        var demoMode = args.Any(argument =>
            argument.Equals("--demo", StringComparison.OrdinalIgnoreCase) ||
            argument.StartsWith("--preview=", StringComparison.OrdinalIgnoreCase));
        var previewPath = args
            .FirstOrDefault(argument => argument.StartsWith("--preview=", StringComparison.OrdinalIgnoreCase))
            ?.Substring("--preview=".Length) ?? "";
        Application.Run(new MainForm(demoMode, previewPath, virtualSheetUrl, justUpdated));
        return 0;
    }

    private static bool IsLoopback(string url) =>
        Uri.TryCreate(url, UriKind.Absolute, out var uri) &&
        uri.Scheme == Uri.UriSchemeHttp &&
        (uri.IsLoopback || uri.Host.EndsWith(".localhost", StringComparison.OrdinalIgnoreCase));
}
