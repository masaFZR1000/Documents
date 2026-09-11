using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace TemperatureProfileVerifier;

internal static class Program
{
    static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;

        if (args.Length == 0)
        {
            PrintUsage();
            return 1;
        }

        string targetDir = args[0];
        if (!Directory.Exists(targetDir))
        {
            Console.Error.WriteLine($"[ERROR] Directory not found: {targetDir}");
            return 2;
        }

        // 1) 温度プロファイルファイル探索
        //    例: 温度プロファイル.TXT / 温度ﾌﾟﾛﾌｧｲﾙ.TXT / profile*.txt
        string? profilePath = FindProfileFile(targetDir);
        if (profilePath is null)
        {
            Console.Error.WriteLine("[ERROR] 温度プロファイル.TXT が見つかりません。");
            return 3;
        }

        // 2) LOGファイル探索（再帰）
        var logFiles = Directory.EnumerateFiles(targetDir, "*.LOG", SearchOption.AllDirectories)
                                .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                                .ToList();

        if (logFiles.Count == 0)
        {
            Console.Error.WriteLine("[ERROR] *.LOG ファイルが見つかりません。");
            return 4;
        }

        Console.WriteLine($"Profile : {profilePath}");
        Console.WriteLine($"LogFiles: {logFiles.Count}");

        // 3) プロファイル読込
        var profile = TemperatureProfileParser.Parse(profilePath);
        Console.WriteLine($"Profile steps: {profile.Steps.Count}");

        // 4) LOG読込（時系列連結）
        var allRecords = new List<LogRecord>(capacity: 200_000);
        foreach (var file in logFiles)
        {
            var records = LogParser.Parse(file);
            allRecords.AddRange(records);
        }

        allRecords.Sort((a, b) => a.Timestamp.CompareTo(b.Timestamp));

        if (allRecords.Count == 0)
        {
            Console.Error.WriteLine("[ERROR] 有効なログレコードがありません。");
            return 5;
        }

        // 5) 検証
        var verifier = new ProfileVerifier(
            toleranceCelsius: 1.0,     // 許容偏差
            settleMinutes: 3,          // 到達猶予
            requiredHitRate: 0.8       // 区間内一致率
        );

        VerificationResult result = verifier.Verify(profile, allRecords);

        // 6) 結果出力（console + report）
        Console.WriteLine();
        Console.WriteLine("=== Verification Summary ===");
        Console.WriteLine($"Overall : {(result.IsSuccess ? "PASS" : "FAIL")}");
        Console.WriteLine($"Steps   : {result.StepResults.Count}");
        Console.WriteLine($"Pass    : {result.StepResults.Count(x => x.IsPass)}");
        Console.WriteLine($"Fail    : {result.StepResults.Count(x => !x.IsPass)}");

        string reportPath = Path.Combine(targetDir, $"verification_report_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
        File.WriteAllText(reportPath, ReportWriter.BuildText(result, profilePath, logFiles), Encoding.UTF8);

        Console.WriteLine($"Report  : {reportPath}");
        return result.IsSuccess ? 0 : 10;
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  TemperatureProfileVerifier.exe <target-directory>");
        Console.WriteLine();
        Console.WriteLine("Example:");
        Console.WriteLine("  TemperatureProfileVerifier.exe \"C:\\work\\20260910\"");
    }

    private static string? FindProfileFile(string root)
    {
        var files = Directory.EnumerateFiles(root, "*.TXT", SearchOption.AllDirectories).ToList();

        // 優先: 温度プロファイル
        var exact = files.FirstOrDefault(f =>
        {
            string name = Path.GetFileName(f);
            return name.Contains("温度プロファイル", StringComparison.OrdinalIgnoreCase)
                || name.Contains("温度ﾌﾟﾛﾌｧｲﾙ", StringComparison.OrdinalIgnoreCase);
        });
        if (!string.IsNullOrEmpty(exact)) return exact;

        // 次点: profile
        var profile = files.FirstOrDefault(f =>
            Path.GetFileName(f).Contains("profile", StringComparison.OrdinalIgnoreCase));
        return profile;
    }
}

public sealed record ProfileStep(
    int StepNo,
    double TargetTemperature,
    TimeSpan Duration
);

public sealed class TemperatureProfile
{
    public List<ProfileStep> Steps { get; } = new();
}

public sealed record LogRecord(
    DateTime Timestamp,
    double? MeasuredTemperature,
    double? SetTemperature,
    string RawLine
);

public static class TemperatureProfileParser
{
    // 想定入力（ゆるく対応）:
    // step,target,durationMin
    // 1,30.0,10
    // 2,40.0,20
    // または 日本語列名混在
    public static TemperatureProfile Parse(string path)
    {
        var profile = new TemperatureProfile();
        var lines = File.ReadAllLines(path, Encoding.UTF8);

        int stepNo = 1;
        foreach (string raw in lines)
        {
            string line = raw.Trim();
            if (string.IsNullOrWhiteSpace(line)) continue;
            if (line.StartsWith("#")) continue;

            // 区切り: , \t 空白
            string[] parts = line.Split([',', '\t'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            if (parts.Length < 2) continue;

            // ヘッダ行スキップ
            if (parts.Any(p => p.Contains("step", StringComparison.OrdinalIgnoreCase) ||
                               p.Contains("target", StringComparison.OrdinalIgnoreCase) ||
                               p.Contains("温度", StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            // 柔軟パース: 数値を抽出
            var nums = ExtractDoubles(line);
            if (nums.Count < 2) continue;

            // パターン:
            // [step, temp, min] or [temp, min]
            double target, durationMin;
            if (nums.Count >= 3)
            {
                stepNo = (int)Math.Round(nums[0]);
                target = nums[1];
                durationMin = nums[2];
            }
            else
            {
                target = nums[0];
                durationMin = nums[1];
            }

            if (durationMin <= 0) continue;

            profile.Steps.Add(new ProfileStep(
                StepNo: stepNo,
                TargetTemperature: target,
                Duration: TimeSpan.FromMinutes(durationMin)
            ));

            stepNo++;
        }

        if (profile.Steps.Count == 0)
            throw new InvalidOperationException("温度プロファイルのステップを解析できませんでした。");

        return profile;
    }

    private static List<double> ExtractDoubles(string text)
    {
        var list = new List<double>();
        foreach (Match m in Regex.Matches(text, @"[-+]?\d+(\.\d+)?"))
        {
            if (double.TryParse(m.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var v))
                list.Add(v);
        }
        return list;
    }
}

public static class LogParser
{
    // 例ログ:
    // 2026/09/01 14:00:07,8185525,01,MON?
    // 2026/09/01 14:00:07,8185579,29.3,,CONSTANT,0  <- 測定温度
    // 2026/09/01 14:00:26,8204791,01,TEMP,S28.6    <- 設定温度コマンド
    static readonly Regex TempCommandRegex = new(@"TEMP,S(?<temp>-?\d+(\.\d+)?)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public static List<LogRecord> Parse(string path)
    {
        var result = new List<LogRecord>();
        foreach (string raw in File.ReadLines(path, Encoding.UTF8))
        {
            if (string.IsNullOrWhiteSpace(raw)) continue;

            var parts = raw.Split(',');
            if (parts.Length < 3) continue;

            if (!DateTime.TryParseExact(parts[0].Trim(),
                    "yyyy/MM/dd HH:mm:ss",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out DateTime ts))
            {
                continue;
            }

            double? measured = null;
            double? setTemp = null;

            // 測定温度候補: 3列目が数値（例 29.3）
            if (double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out double m))
            {
                measured = m;
            }

            // 設定温度候補: TEMP,Sxx.x
            var m2 = TempCommandRegex.Match(raw);
            if (m2.Success &&
                double.TryParse(m2.Groups["temp"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double st))
            {
                setTemp = st;
            }

            // どちらも無ければ捨てる
            if (measured is null && setTemp is null) continue;

            result.Add(new LogRecord(ts, measured, setTemp, raw));
        }

        return result;
    }
}

public sealed record StepVerificationResult(
    int StepNo,
    DateTime StartTime,
    DateTime EndTime,
    double TargetTemperature,
    double AverageMeasured,
    double HitRate,
    bool ReachedWithinSettleWindow,
    bool IsPass,
    string Reason
);

public sealed class VerificationResult
{
    public bool IsSuccess { get; set; }
    public List<StepVerificationResult> StepResults { get; } = new();
}

public sealed class ProfileVerifier
{
    private readonly double _tolerance;
    private readonly TimeSpan _settleWindow;
    private readonly double _requiredHitRate;

    public ProfileVerifier(double toleranceCelsius, int settleMinutes, double requiredHitRate)
    {
        _tolerance = toleranceCelsius;
        _settleWindow = TimeSpan.FromMinutes(settleMinutes);
        _requiredHitRate = requiredHitRate;
    }

    public VerificationResult Verify(TemperatureProfile profile, List<LogRecord> records)
    {
        var result = new VerificationResult();

        // 解析起点:
        // 最初の設定温度イベントを開始点にする（無ければ最初の計測時刻）
        DateTime cursor = records.FirstOrDefault(r => r.SetTemperature.HasValue)?.Timestamp
                          ?? records.First(r => r.MeasuredTemperature.HasValue).Timestamp;

        foreach (var step in profile.Steps)
        {
            DateTime start = cursor;
            DateTime end = start + step.Duration;

            var section = records.Where(r => r.Timestamp >= start && r.Timestamp < end && r.MeasuredTemperature.HasValue)
                                 .Select(r => r.MeasuredTemperature!.Value)
                                 .ToList();

            bool reached = records.Any(r =>
                r.Timestamp >= start &&
                r.Timestamp <= start + _settleWindow &&
                r.MeasuredTemperature.HasValue &&
                Math.Abs(r.MeasuredTemperature.Value - step.TargetTemperature) <= _tolerance);

            double avg = section.Count > 0 ? section.Average() : double.NaN;
            double hitRate = section.Count > 0
                ? section.Count(v => Math.Abs(v - step.TargetTemperature) <= _tolerance) / (double)section.Count
                : 0.0;

            bool pass = section.Count > 0
                        && reached
                        && hitRate >= _requiredHitRate;

            string reason = pass
                ? "OK"
                : BuildReason(section.Count, reached, hitRate);

            result.StepResults.Add(new StepVerificationResult(
                step.StepNo, start, end, step.TargetTemperature, avg, hitRate, reached, pass, reason));

            cursor = end;
        }

        result.IsSuccess = result.StepResults.All(x => x.IsPass);
        return result;
    }

    private string BuildReason(int count, bool reached, double hitRate)
    {
        if (count == 0) return "区間内に測定値なし";
        if (!reached) return "到達猶予時間内に目標温度へ到達せず";
        if (hitRate < _requiredHitRate) return $"一致率不足({hitRate:P1} < {_requiredHitRate:P1})";
        return "不明";
    }
}

public static class ReportWriter
{
    public static string BuildText(VerificationResult result, string profilePath, List<string> logs)
    {
        var sb = new StringBuilder();
        sb.AppendLine("温度制御プロファイル検証レポート");
        sb.AppendLine($"生成日時: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine($"プロファイル: {profilePath}");
        sb.AppendLine("ログファイル:");
        foreach (var l in logs) sb.AppendLine($"  - {l}");
        sb.AppendLine();
        sb.AppendLine($"総合判定: {(result.IsSuccess ? "PASS" : "FAIL")}");
        sb.AppendLine();

        sb.AppendLine("Step,Start,End,Target,AvgMeasured,HitRate,ReachedInSettle,Pass,Reason");
        foreach (var s in result.StepResults)
        {
            sb.AppendLine($"{s.StepNo},{s.StartTime:yyyy/MM/dd HH:mm:ss},{s.EndTime:yyyy/MM/dd HH:mm:ss},{s.TargetTemperature:F1},{s.AverageMeasured:F2},{s.HitRate:P1},{s.ReachedWithinSettleWindow},{s.IsPass},{s.Reason}");
        }

        return sb.ToString();
    }
}