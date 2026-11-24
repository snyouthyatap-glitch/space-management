using System.Diagnostics;

namespace YouthSpring.StaffDashboard;

// 자동 업데이트: 작업 폴더의 staff-dashboard\release에 새 실행 파일이 놓이면,
// 설치된 프로그램이 그 파일로 자기 자신을 바꾸고 다시 켠다.
// 실행 중인 exe는 덮어쓸 수 없지만 이름은 바꿀 수 있어서, 지금 파일을 .old로 옮기고 새 파일을 복사한다.
internal static class AppUpdater
{
    public const string ExeName = "청년이봄 직원알림.exe";
    // 복사 중인 파일을 집어 가지 않도록, 마지막으로 바뀐 뒤 이만큼 지난 파일만 쓴다.
    public static readonly TimeSpan SettleTime = TimeSpan.FromMinutes(2);
    private const long MinimumExeBytes = 10 * 1024 * 1024;

    public static string ReleaseExePath(string repoRoot) =>
        Path.Combine(repoRoot, "staff-dashboard", "release", ExeName);

    // 설치 폴더(%LOCALAPPDATA%\YouthSpringStaffDashboard\App)에서 실행 중일 때만 스스로 바꾼다.
    public static bool IsInstalledCopy(string? processPath) =>
        !string.IsNullOrWhiteSpace(processPath) &&
        processPath.Contains(Path.Combine("YouthSpringStaffDashboard", "App"), StringComparison.OrdinalIgnoreCase);

    public static bool ShouldUpdate(
        DateTime currentWriteUtc,
        DateTime releaseWriteUtc,
        long releaseBytes,
        DateTime utcNow) =>
        releaseBytes >= MinimumExeBytes &&
        releaseWriteUtc > currentWriteUtc.AddMinutes(1) &&
        releaseWriteUtc <= utcNow - SettleTime;

    // 새 파일로 바꾸고 새 프로세스를 띄운다. 새 프로세스는 이 프로세스가 끝나길 기다렸다가 시작한다.
    // 실패하면 원래 파일로 되돌리고 예외를 던진다.
    public static void ReplaceAndRestart(string currentExe, string releaseExe)
    {
        var oldExe = currentExe + ".old";
        if (File.Exists(oldExe)) File.Delete(oldExe);
        File.Move(currentExe, oldExe);
        try
        {
            File.Copy(releaseExe, currentExe, overwrite: false);
            File.SetLastWriteTimeUtc(currentExe, File.GetLastWriteTimeUtc(releaseExe));
            var start = new ProcessStartInfo
            {
                FileName = currentExe,
                WorkingDirectory = Path.GetDirectoryName(currentExe) ?? "",
                UseShellExecute = false
            };
            // 처음 켤 때 받은 옵션은 그대로 넘긴다(예전 --after-update는 뺀다).
            foreach (var argument in Environment.GetCommandLineArgs().Skip(1)
                         .Where(item => !item.StartsWith("--after-update=", StringComparison.OrdinalIgnoreCase)))
                start.ArgumentList.Add(argument);
            start.ArgumentList.Add($"--after-update={Environment.ProcessId}");
            Process.Start(start);
        }
        catch
        {
            if (File.Exists(currentExe)) File.Delete(currentExe);
            File.Move(oldExe, currentExe);
            throw;
        }
    }

    // 업데이트 뒤 남은 예전 파일을 지운다(이전 프로세스가 끝난 뒤라 지울 수 있다).
    public static void CleanupOldCopy(string? processPath)
    {
        if (string.IsNullOrWhiteSpace(processPath)) return;
        try
        {
            var oldExe = processPath + ".old";
            if (File.Exists(oldExe)) File.Delete(oldExe);
        }
        catch
        {
            // 다음 실행 때 다시 지운다.
        }
    }

    // --after-update=<이전 프로세스 번호>: 이전 프로세스가 완전히 끝날 때까지(최대 30초) 기다린다.
    public static bool WaitForPreviousProcess(string[] args)
    {
        var argument = args.FirstOrDefault(item => item.StartsWith("--after-update=", StringComparison.OrdinalIgnoreCase));
        if (argument is null) return false;
        if (int.TryParse(argument["--after-update=".Length..], out var pid))
        {
            try
            {
                using var previous = Process.GetProcessById(pid);
                previous.WaitForExit(30_000);
            }
            catch (ArgumentException)
            {
                // 이미 끝났다.
            }
        }
        return true;
    }
}
