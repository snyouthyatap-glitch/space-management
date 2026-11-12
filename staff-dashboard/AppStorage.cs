using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace YouthSpring.StaffDashboard;

internal sealed class AppStorage
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    private readonly string _directory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "YouthSpringStaffDashboard");

    private string SettingsPath => Path.Combine(_directory, "settings.json");
    private string StatePath => Path.Combine(_directory, "events.json");

    public DashboardSettings LoadSettings() => Load(SettingsPath, new DashboardSettings());

    public LocalState LoadState() => Load(StatePath, new LocalState());

    public void SaveSettings(DashboardSettings settings) => Save(SettingsPath, settings);

    public void SaveState(LocalState state) => Save(StatePath, state);

    // 네이버 확인·로그인 기록. 프로그램을 다시 켜도 '오류 로그 보기'에서 볼 수 있게 파일로 남긴다.
    private const int MaxDiagnosticLogBytes = 256 * 1024;
    private string DiagnosticLogPath => Path.Combine(_directory, "naver-sync.log");

    public void AppendDiagnosticLog(string entry)
    {
        try
        {
            Directory.CreateDirectory(_directory);
            File.AppendAllText(DiagnosticLogPath, entry.TrimEnd() + Environment.NewLine + Environment.NewLine, Encoding.UTF8);
            var info = new FileInfo(DiagnosticLogPath);
            if (info.Length <= MaxDiagnosticLogBytes) return;
            // 너무 커지면 뒤쪽(최근) 절반만 남긴다.
            var text = File.ReadAllText(DiagnosticLogPath, Encoding.UTF8);
            File.WriteAllText(DiagnosticLogPath, text[(text.Length / 2)..], Encoding.UTF8);
        }
        catch
        {
            // 기록 실패가 프로그램 동작을 막지 않게 한다.
        }
    }

    public string ReadDiagnosticLog(int maxChars = 60_000)
    {
        try
        {
            if (!File.Exists(DiagnosticLogPath)) return "";
            var text = File.ReadAllText(DiagnosticLogPath, Encoding.UTF8);
            return text.Length <= maxChars ? text : "…(앞부분 생략)\r\n" + text[^maxChars..];
        }
        catch (Exception exception)
        {
            return "기록 파일을 읽지 못했습니다: " + exception.Message;
        }
    }

    public string ProtectPassword(string password)
    {
        if (string.IsNullOrEmpty(password)) return "";
        var plain = Encoding.UTF8.GetBytes(password);
        var protectedBytes = Protect(plain);
        return Convert.ToBase64String(protectedBytes);
    }

    public string UnprotectPassword(string protectedPassword)
    {
        if (string.IsNullOrWhiteSpace(protectedPassword)) return "";
        try
        {
            var protectedBytes = Convert.FromBase64String(protectedPassword);
            return Encoding.UTF8.GetString(Unprotect(protectedBytes));
        }
        catch
        {
            return "";
        }
    }

    private static T Load<T>(string path, T fallback)
    {
        try
        {
            if (!File.Exists(path)) return fallback;
            return JsonSerializer.Deserialize<T>(File.ReadAllText(path), JsonOptions) ?? fallback;
        }
        catch
        {
            return fallback;
        }
    }

    private void Save<T>(string path, T value)
    {
        Directory.CreateDirectory(_directory);
        var temporaryPath = path + ".tmp";
        File.WriteAllText(temporaryPath, JsonSerializer.Serialize(value, JsonOptions));
        File.Move(temporaryPath, path, true);
    }

    private static byte[] Protect(byte[] input) => Crypt(input, true);

    private static byte[] Unprotect(byte[] input) => Crypt(input, false);

    private static byte[] Crypt(byte[] input, bool protect)
    {
        var inputBlob = new DataBlob();
        var outputBlob = new DataBlob();
        try
        {
            inputBlob.Size = input.Length;
            inputBlob.Data = Marshal.AllocHGlobal(input.Length);
            Marshal.Copy(input, 0, inputBlob.Data, input.Length);

            var ok = protect
                ? CryptProtectData(ref inputBlob, null, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, 0, ref outputBlob)
                : CryptUnprotectData(ref inputBlob, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, 0, ref outputBlob);
            if (!ok) throw new InvalidOperationException("Windows 암호화 처리를 완료하지 못했습니다.");

            var result = new byte[outputBlob.Size];
            Marshal.Copy(outputBlob.Data, result, 0, outputBlob.Size);
            return result;
        }
        finally
        {
            if (inputBlob.Data != IntPtr.Zero) Marshal.FreeHGlobal(inputBlob.Data);
            if (outputBlob.Data != IntPtr.Zero) LocalFree(outputBlob.Data);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DataBlob
    {
        public int Size;
        public IntPtr Data;
    }

    [DllImport("crypt32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CryptProtectData(
        ref DataBlob dataIn,
        string? description,
        IntPtr optionalEntropy,
        IntPtr reserved,
        IntPtr promptStruct,
        int flags,
        ref DataBlob dataOut);

    [DllImport("crypt32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CryptUnprotectData(
        ref DataBlob dataIn,
        IntPtr description,
        IntPtr optionalEntropy,
        IntPtr reserved,
        IntPtr promptStruct,
        int flags,
        ref DataBlob dataOut);

    [DllImport("kernel32.dll")]
    private static extern IntPtr LocalFree(IntPtr memory);
}
