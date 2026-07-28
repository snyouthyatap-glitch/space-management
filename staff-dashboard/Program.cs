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
        using var singleInstance = new Mutex(
            true,
            @"Local\YouthSpringStaffDashboard",
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
        Application.Run(new MainForm(demoMode, previewPath));
        return 0;
    }
}
