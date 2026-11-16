using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace NeisAnalyzer
{
    /// <summary>명령줄 모드: 나이스성적분석.exe --cli 파일1.xlsx 파일2.xlsx ... [--out 결과.xlsx]</summary>
    internal static class Cli
    {
        public static int Run(string[] args)
        {
            AttachConsole();
            Console.OutputEncoding = Encoding.UTF8;
            var files = new List<string>();
            string outPath = null;
            for (int i = 1; i < args.Length; i++)
            {
                if (args[i] == "--out" && i + 1 < args.Length) { outPath = args[++i]; continue; }
                if (Directory.Exists(args[i])) files.AddRange(Directory.GetFiles(args[i], "*.xlsx").Where(f => !Path.GetFileName(f).StartsWith("~$")));
                else files.Add(args[i]);
            }
            var sw = Stopwatch.StartNew();
            var detected = files.Select(NeisReader.Detect).ToList();
            Console.WriteLine("== 파일 판별 ==");
            foreach (var d in detected) Console.WriteLine($"  [{Path.GetFileName(d.Path)}] {d.Label}");

            var settings = Analyzer.DefaultSettings(detected);
            Console.WriteLine("== 기본 설정 ==");
            foreach (var a in settings.PerfAreas) Console.WriteLine($"  수행 {a.Name}: 만점 {a.Max}, 비율 {a.Ratio}%");
            foreach (var e in settings.ExamRatios) Console.WriteLine($"  지필 {e.Key}: 비율 {e.Value}%");

            var r = Analyzer.Analyze(detected, settings);
            Console.WriteLine("== 경고 ==");
            if (r.Warnings.Count == 0) Console.WriteLine("  없음");
            foreach (var w in r.Warnings) Console.WriteLine("  ! " + w);

            Console.WriteLine("== 검산 ==");
            foreach (var ea in r.ExamAnalyses)
            {
                if (ea.Table == null) { Console.WriteLine($"  {ea.ExamName}: 일람표 없음(검산 생략)"); continue; }
                Console.WriteLine($"  {ea.ExamName}: {(ea.VerifyAllMatch == true ? "모두 일치" : "불일치 있음")}");
                foreach (var v in ea.Verify)
                    Console.WriteLine($"    {v.ClassNo}{r.Unit}  인원 {v.CalcCount}/{v.TableCount}  총점 {v.CalcSum}/{v.TableSum}  평균 {v.CalcMean}/{v.TableMean}  → {v.Result}");
                foreach (var i in ea.VerifyStudentIssues) Console.WriteLine("    - " + i);
            }
            Console.WriteLine("== 요약 ==");
            foreach (var h in r.Highlights) Console.WriteLine("  • " + h);

            if (outPath == null && detected.Count > 0)
                outPath = Path.Combine(Path.GetDirectoryName(detected[0].Path), $"성적분석결과_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");
            if (outPath != null)
            {
                ReportWriter.Write(r, outPath);
                Console.WriteLine($"== 저장: {outPath} ({sw.Elapsed.TotalSeconds:0.0}초) ==");
            }
            return r.ExamAnalyses.Any(e => e.VerifyAllMatch == false) ? 2 : 0;
        }

        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        static extern bool AttachConsole(int pid);
        static void AttachConsole() { try { AttachConsole(-1); } catch { } }
    }
}
