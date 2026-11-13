using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;

namespace NeisAnalyzer
{
    /// <summary>분석 결과를 시트 14개짜리 엑셀 파일로 저장한다.</summary>
    public static class ReportWriter
    {
        static readonly Color Blue = Color.FromArgb(31, 78, 121);
        static readonly Color LightBlue = Color.FromArgb(221, 235, 247);
        static readonly Color Yellow = Color.FromArgb(255, 242, 204);
        static readonly Color Red = Color.FromArgb(252, 228, 228);
        static readonly Color Green = Color.FromArgb(226, 239, 218);
        static readonly Color Gray = Color.FromArgb(242, 242, 242);

        public static void Write(AnalysisResult r, string path)
        {
            using (var pkg = new ExcelPackage())
            {
                WriteGlossary(pkg, r);
                WriteSummary(pkg, r);
                WriteStudents(pkg, r);
                WriteRanking(pkg, r);
                WriteAchievement(pkg, r);
                WriteClassCompare(pkg, r);
                WriteDistribution(pkg, r);
                if (r.HasPerf) WritePerf(pkg, r);
                foreach (var ea in r.ExamAnalyses) WriteItems(pkg, r, ea);
                if (r.HasPerf) WritePerfVsWritten(pkg, r);
                WriteWatch(pkg, r);
                if (r.Exams.Count >= 2) WriteChange(pkg, r);
                WriteIndividual(pkg, r);
                if (r.ExamAnalyses.Any(e => e.Table != null)) WriteVerify(pkg, r);
                WriteCriteria(pkg, r);
                var fi = new FileInfo(path);
                if (fi.Exists) fi.Delete();
                pkg.SaveAs(fi);
            }
        }

        // ===================== 공통 도우미 =====================
        class Sheet
        {
            public ExcelWorksheet Ws;
            public int Row = 1;
            public Sheet(ExcelPackage pkg, string name, string title, string howToRead)
            {
                Ws = pkg.Workbook.Worksheets.Add(name);
                Ws.Cells[1, 1].Value = title;
                Ws.Cells[1, 1].Style.Font.Size = 16; Ws.Cells[1, 1].Style.Font.Bold = true; Ws.Cells[1, 1].Style.Font.Color.SetColor(Blue);
                Ws.Row(1).Height = 26;
                Ws.Cells[2, 1].Value = "▶ " + howToRead;
                Ws.Cells[2, 1].Style.Font.Color.SetColor(Color.FromArgb(89, 89, 89));
                Ws.Cells[2, 1].Style.WrapText = false;
                Row = 4;
                Ws.Cells.Style.Font.Name = "맑은 고딕";
                Ws.Cells.Style.Font.Size = 10;
                Ws.Cells[1, 1].Style.Font.Size = 16;
                Ws.View.ShowGridLines = false;
                Ws.PrinterSettings.FitToPage = true; Ws.PrinterSettings.FitToWidth = 1; Ws.PrinterSettings.FitToHeight = 0;
                Ws.PrinterSettings.Orientation = eOrientation.Landscape;
                Ws.PrinterSettings.PaperSize = ePaperSize.A4;
            }
            public void Section(string text)
            {
                Row++;
                Ws.Cells[Row, 1].Value = text;
                Ws.Cells[Row, 1].Style.Font.Bold = true; Ws.Cells[Row, 1].Style.Font.Size = 12; Ws.Cells[Row, 1].Style.Font.Color.SetColor(Blue);
                Row++;
            }
            public void Text(string text, bool bold = false, Color? color = null)
            {
                Ws.Cells[Row, 1].Value = text;
                if (bold) Ws.Cells[Row, 1].Style.Font.Bold = true;
                if (color.HasValue) Ws.Cells[Row, 1].Style.Font.Color.SetColor(color.Value);
                Row++;
            }
            public int Header(params string[] cols)
            {
                for (int i = 0; i < cols.Length; i++)
                {
                    var c = Ws.Cells[Row, i + 1];
                    c.Value = cols[i];
                    c.Style.Font.Bold = true; c.Style.Font.Color.SetColor(Color.White);
                    c.Style.Fill.PatternType = ExcelFillStyle.Solid; c.Style.Fill.BackgroundColor.SetColor(Blue);
                    c.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; c.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    c.Style.WrapText = true;
                    Border(c);
                }
                Ws.Row(Row).Height = 30;
                return Row++;
            }
            public void Cells(params object[] vals) => Cells(null, vals);
            public void Cells(Color? fill, params object[] vals)
            {
                for (int i = 0; i < vals.Length; i++)
                {
                    var c = Ws.Cells[Row, i + 1];
                    var v = vals[i];
                    if (v is double d) { c.Value = d; c.Style.Numberformat.Format = Math.Abs(d - Math.Round(d)) < 1e-9 ? "0" : "0.0#"; }
                    else if (v is Pct p) { c.Value = p.V; c.Style.Numberformat.Format = "0.0%"; }
                    else if (v is Fmt f) { c.Value = f.V; c.Style.Numberformat.Format = f.F; }
                    else c.Value = v;
                    if (v is string) c.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left; else c.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    if (fill.HasValue) { c.Style.Fill.PatternType = ExcelFillStyle.Solid; c.Style.Fill.BackgroundColor.SetColor(fill.Value); }
                    Border(c);
                }
                Row++;
            }
            public void Fill(int row, int col, Color color)
            {
                var c = Ws.Cells[row, col];
                c.Style.Fill.PatternType = ExcelFillStyle.Solid; c.Style.Fill.BackgroundColor.SetColor(color);
            }
            public void Bold(int row, int col) => Ws.Cells[row, col].Style.Font.Bold = true;
            public void Freeze(int headerRow, int cols = 0) => Ws.View.FreezePanes(headerRow + 1, cols + 1);
            public void Widths(params double[] w) { for (int i = 0; i < w.Length; i++) Ws.Column(i + 1).Width = w[i]; }
            public void AutoWidths(int maxCols, double min = 8, double max = 40)
            {
                for (int c = 1; c <= maxCols; c++)
                {
                    double best = min;
                    for (int rr = 3; rr <= Ws.Dimension.End.Row; rr++)
                    {
                        var v = Ws.Cells[rr, c].Value; if (v == null) continue;
                        var s = v is double d ? d.ToString("0.0#") : v.ToString();
                        if (s.Contains("\n")) s = s.Split('\n').OrderByDescending(x => x.Length).First();
                        double len = s.Sum(ch => ch > 0x7F ? 1.9 : 1.1) + 2;
                        if (len > best) best = len;
                    }
                    Ws.Column(c).Width = Math.Min(max, best);
                }
            }
            public void RepeatHeader(int headerRow) { try { Ws.PrinterSettings.RepeatRows = Ws.Cells[$"{headerRow}:{headerRow}"]; } catch { } }
            static void Border(ExcelRange c)
            {
                c.Style.Border.Top.Style = c.Style.Border.Bottom.Style = c.Style.Border.Left.Style = c.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                c.Style.Border.Top.Color.SetColor(Color.FromArgb(191, 191, 191)); c.Style.Border.Bottom.Color.SetColor(Color.FromArgb(191, 191, 191));
                c.Style.Border.Left.Color.SetColor(Color.FromArgb(191, 191, 191)); c.Style.Border.Right.Color.SetColor(Color.FromArgb(191, 191, 191));
            }
            /// <summary>표 아래에 "이 표의 용어 읽는 법" 블록을 붙인다. 열 너비를 정한 뒤에 부를 것.</summary>
            public void Explain(params string[] keys)
            {
                Row++;
                Section("📘 이 표의 용어 읽는 법  (숫자가 무엇을 뜻하고 어떻게 읽는지)");
                int lastCol = Math.Max(8, Ws.Dimension == null ? 8 : Math.Min(14, Ws.Dimension.End.Column));
                double textWidth = 0; for (int c = 3; c <= lastCol; c++) textWidth += Ws.Column(c).Width;
                if (Ws.Column(1).Width + Ws.Column(2).Width < 18) { Ws.Column(1).Width = Math.Max(Ws.Column(1).Width, 8); Ws.Column(2).Width = Math.Max(Ws.Column(2).Width, 10); }
                foreach (var key in keys)
                {
                    var t = Glossary.Get(key); if (t == null) continue;
                    var name = Ws.Cells[Row, 1, Row, 2]; name.Merge = true; name.Value = t.Name;
                    name.Style.Font.Bold = true; name.Style.Font.Color.SetColor(Blue); name.Style.VerticalAlignment = ExcelVerticalAlignment.Top; name.Style.WrapText = true;
                    name.Style.Fill.PatternType = ExcelFillStyle.Solid; name.Style.Fill.BackgroundColor.SetColor(LightBlue);
                    var text = t.What + "\n▸ 값의 범위: " + t.Range + "   ▸ 방향: " + t.Better + "\n▸ 읽는 법: " + t.HowToRead + (string.IsNullOrEmpty(t.Caution) ? "" : "\n▸ 주의: " + t.Caution);
                    var body = Ws.Cells[Row, 3, Row, lastCol]; body.Merge = true; body.Value = text;
                    body.Style.WrapText = true; body.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    body.Style.Fill.PatternType = ExcelFillStyle.Solid; body.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(248, 250, 253));
                    int lines = 0;
                    foreach (var line in text.Split('\n')) lines += Math.Max(1, (int)Math.Ceiling(line.Sum(ch => ch > 0x7F ? 2.0 : 1.0) / Math.Max(20, textWidth * 0.95)));
                    Ws.Row(Row).Height = Math.Min(409, lines * 15.5 + 6);
                    Row++;
                }
            }
            /// <summary>그래프를 현재 줄 아래에 놓고, 그래프 높이만큼 줄을 내린다. 그래프끼리 겹치지 않게 항상 세로로 차례로 배치한다.</summary>
            public ExcelChart ChartBelow(string name, eChartType type, string title, int width = 720, int height = 320)
            {
                Row++;
                var ch = Chart(name, type, title, Row, 1, width, height);
                Row += (int)Math.Ceiling(height / 20.0) + 2;
                return ch;
            }
            public ExcelChart Chart(string name, eChartType type, string title, int row, int col, int width = 640, int height = 320)
            {
                var ch = Ws.Drawings.AddChart(name, type);
                ch.Title.Text = title; ch.Title.Font.Size = 12; ch.Title.Font.Bold = true;
                ch.SetPosition(row - 1, 0, col - 1, 0);
                ch.SetSize(width, height);
                ch.Style = eChartStyle.Style2;
                return ch;
            }
        }
        class Pct { public double V; public Pct(double v) { V = v; } }
        class Fmt { public double V; public string F; public Fmt(double v, string f) { V = v; F = f; } }
        static Pct P(double v) => new Pct(v);
        static Fmt F1(double v) => new Fmt(v, "0.0");
        static Fmt F2(double v) => new Fmt(v, "0.00");
        static object N(double? v) => v.HasValue ? (object)v.Value : "";
        static object N1(double? v) => v.HasValue ? (object)F1(v.Value) : "";
        static object N2(double? v) => v.HasValue ? (object)F2(v.Value) : "";
        static string SubjectTitle(AnalysisResult r) => r.Header == null ? "" : $"{r.Header.Year}학년도 {r.Header.Semester}학기 {r.Header.Grade}학년 {r.Header.Subject} {string.Join("·", r.Exams)}";
        static string ClassLabel(int c) => c + "반";

        // ===================== 0. 읽는 법 (용어 설명) =====================
        static void WriteGlossary(ExcelPackage pkg, AnalysisResult r)
        {
            var s = new Sheet(pkg, "0.읽는 법", "이 결과 파일 읽는 법 — 모든 지표의 뜻과 해석 기준", "표에 나오는 숫자가 무엇을 뜻하고, 높을수록 좋은지, 어느 정도면 좋은 값인지 정리했습니다. 각 시트 맨 아래에도 그 표에 쓰인 용어 설명이 다시 나옵니다.");
            s.Text("• 시트 순서대로 보시면 됩니다: 1.요약(전체 그림) → 2~3(학생별 점수·등수) → 4~6(성취도·반별·분포) → 7(수행평가) → 8(문항 분석) → 9~11(비교) → 12(개인별, 인쇄용) → 13(검산) → 14(기준).", false, Color.FromArgb(64, 64, 64));
            s.Text("• 빨간색 칸 = 점검이 필요함, 노란색 칸 = 경계·주의, 초록색 칸 = 좋음/일치, 회색 = 통계에서 제외.", false, Color.FromArgb(64, 64, 64));
            s.Row++;
            int hr = s.Header("용어", "무엇인가", "값의 범위", "방향 (높고 낮음의 뜻)", "수치 읽는 법 (어느 정도면 좋은 값인지)", "주의할 점");
            foreach (var t in Glossary.Terms)
            {
                s.Cells(t.Name, t.What, t.Range, t.Better, t.HowToRead, t.Caution);
                int row = s.Row - 1;
                for (int c = 1; c <= 6; c++) { s.Ws.Cells[row, c].Style.WrapText = true; s.Ws.Cells[row, c].Style.VerticalAlignment = ExcelVerticalAlignment.Top; s.Ws.Cells[row, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left; }
                s.Bold(row, 1); s.Ws.Cells[row, 1].Style.Font.Color.SetColor(Blue);
                int lines = Math.Max(Lines(t.What, 40), Math.Max(Lines(t.HowToRead, 70), Lines(t.Caution, 46)));
                s.Ws.Row(row).Height = Math.Min(409, lines * 15.5 + 8);
            }
            s.Freeze(hr);
            s.RepeatHeader(hr);
            s.Widths(24, 42, 12, 18, 72, 48);
        }
        static int Lines(string text, double widthChars) => string.IsNullOrEmpty(text) ? 1 : Math.Max(1, (int)Math.Ceiling(text.Sum(ch => ch > 0x7F ? 2.0 : 1.0) / widthChars));

        // ===================== 1. 요약 =====================
        static void WriteSummary(ExcelPackage pkg, AnalysisResult r)
        {
            var s = new Sheet(pkg, "1.요약", $"성적 분석 요약 - {SubjectTitle(r)}", "이번 시험 전체 그림을 한눈에 보는 시트입니다. 자세한 내용은 각 시트를 보세요.");
            s.Ws.PrinterSettings.Orientation = eOrientation.Portrait;
            s.Section("기본 정보");
            s.Header("항목", "값");
            s.Cells("학년도 / 학기", r.Header == null ? "" : $"{r.Header.Year}학년도 {r.Header.Semester}학기");
            s.Cells("학년 / 과목", r.Header == null ? "" : $"{r.Header.Grade}학년 {r.Header.Subject}");
            s.Cells("지필고사", string.Join(", ", r.Exams));
            s.Cells("반", string.Join(", ", r.Classes.Select(ClassLabel)));
            s.Cells("명단 인원 / 분석 인원 / 제외 인원", $"{r.Students.Count}명 / {r.Count}명 / {r.Students.Count - r.Count}명");
            s.Cells("분석 일시", DateTime.Now.ToString("yyyy-MM-dd HH:mm"));

            s.Section("학년 전체 최종 점수");
            s.Header("응시 인원", "평균", "중앙값", "표준편차", "최고점", "최저점");
            s.Cells((double)r.Count, F1(r.Mean), F1(r.Median), F1(r.Sd), N(r.Max), N(r.Min));

            s.Section("성취도 분포");
            s.Header("성취도", "인원", "비율");
            foreach (var g in Analyzer.Grades)
                s.Cells(g, (double)(r.AchievementCounts.ContainsKey(g) ? r.AchievementCounts[g] : 0), P(r.Count == 0 ? 0 : (double)(r.AchievementCounts.ContainsKey(g) ? r.AchievementCounts[g] : 0) / r.Count));

            s.Section("반별 평균 순위");
            s.Header("순위", "반", "인원", "평균", "학년 평균과 차이");
            foreach (var c in r.ClassSummaries.Where(c => c.Count > 0).OrderBy(c => c.MeanRank))
                s.Cells((double)c.MeanRank, ClassLabel(c.ClassNo), (double)c.Count, F1(c.Mean), new Fmt(c.Mean - r.Mean, "+0.0;-0.0;0.0"));

            s.Section("주요 지표 한눈에 (값 + 해석)");
            s.Header("지표", "값", "해석", "이 숫자의 뜻");
            if (r.Count > 0)
            {
                s.Cells("학년 평균(최종 점수)", F1(r.Mean), r.Mean >= 75 ? "높은 편" : r.Mean >= 60 ? "보통" : "낮은 편", "학생 전체 점수의 평균. 중학교 시험은 보통 55~75점 사이입니다.");
                s.Cells("표준편차", F1(r.Sd), Glossary.Sd(r.Sd, 100), "점수가 평균에서 얼마나 퍼져 있는지. 10 이하 고름, 15~20 보통, 25 이상 차이 큼.");
                s.Cells("성취도 E 비율", P((double)r.AchievementCounts["E"] / r.Count), (double)r.AchievementCounts["E"] / r.Count >= 0.3 ? "매우 높음(학습 결손 점검)" : (double)r.AchievementCounts["E"] / r.Count >= 0.2 ? "높은 편" : "보통", "60점 미만 학생의 비율. 10~20%가 흔한 수준입니다.");
                s.Cells("성취도 경계 학생 수", (double)r.Ranked.Count(x => x.Boundary), r.Ranked.Count(x => x.Boundary) > 0 ? "점수 입력 재확인 권장" : "없음", $"기준 점수 ±{r.Settings.BoundaryRange:0.##}점 안에 있어 1점 차이로 성취도가 바뀔 수 있는 학생 수.");
            }
            foreach (var ea in r.ExamAnalyses)
            {
                string pre = r.Exams.Count > 1 ? $"[{ea.ExamName}] " : "";
                var items = ea.ItemStats.Where(i => i.Correct.HasValue).ToList();
                if (items.Count == 0) continue;
                if (ea.KR20.HasValue) s.Cells(pre + "시험 신뢰도 KR-20", F2(ea.KR20.Value), Glossary.KR20(ea.KR20.Value), "시험이 실력을 일관되게 재는 정도(0~1). 0.7 이상 보통, 0.8 이상 양호.");
                double avgP = items.Average(i => i.CorrectRate);
                s.Cells(pre + "평균 정답률", P(avgP), avgP >= 0.8 ? "쉬운 시험" : avgP >= 0.6 ? "적절한 난이도" : avgP >= 0.45 ? "어려운 편" : "매우 어려움", "선택형 문항 정답률의 평균. 60~70%가 적절합니다.");
                int low = items.Count(i => (i.Discrimination ?? 0) < 0.2), neg = items.Count(i => (i.Discrimination ?? 0) < 0);
                s.Cells(pre + "변별도 낮은 문항 수(0.2 미만)", (double)low, low == 0 ? "없음(좋음)" : low <= items.Count * 0.15 ? "일부(점검)" : "많음(문항 검토 필요)", $"선택형 {items.Count}문항 중 잘하는 학생과 못하는 학생을 가르지 못하는 문항 수. 이 중 음수 {neg}개.");
                int inv = items.Count(i => i.LevelInversion != null);
                s.Cells(pre + "성취도 수준 역전 문항 수", (double)inv, inv == 0 ? "없음(좋음)" : "문항 검토 필요('성취도별 문항' 시트)", "낮은 성취도 집단이 높은 집단보다 15%p 이상 많이 맞힌 문항 수. 정답 오류나 찍기 쉬운 문항일 수 있습니다.");
            }
            if (r.PerfWrittenCorr.HasValue) s.Cells("수행 vs 지필 상관계수", F2(r.PerfWrittenCorr.Value), Glossary.Correlation(r.PerfWrittenCorr.Value), "두 평가가 같은 방향으로 움직이는 정도(-1~1). 0.4~0.7이 흔한 값.");

            s.Section("이번 시험의 특징 (자동 요약)");
            foreach (var h in r.Highlights) s.Text("• " + h);

            if (r.Warnings.Count > 0)
            {
                s.Section("확인이 필요한 사항 (경고)");
                foreach (var w in r.Warnings) s.Text("⚠ " + w, false, Color.FromArgb(192, 0, 0));
            }
            s.Widths(34, 44, 26, 70, 12, 12);
            s.Explain("평균", "중앙값", "표준편차", "성취도", "KR20");
        }

        // ===================== 2. 학생별 전체 성적 =====================
        static void WriteStudents(ExcelPackage pkg, AnalysisResult r)
        {
            var s = new Sheet(pkg, "2.학생별 성적", "학생별 전체 성적", "학생 한 명이 한 줄입니다. 노란색은 성취도 경계(±" + r.Settings.BoundaryRange.ToString("0.##") + "점), 회색은 등수·통계에서 제외된 학생입니다.");
            var cols = new List<string> { "반", "번호", "이름" };
            foreach (var a in r.Settings.PerfAreas) cols.Add($"{a.Name}\n({a.Max:0.##}점)");
            if (r.HasPerf) cols.Add($"수행 합계\n({r.PerfMax:0.##}점)");
            foreach (var e in r.Exams) { cols.Add($"{e}\n원점수"); cols.Add($"{e}\n환산({(r.Settings.ExamRatios.ContainsKey(e) ? r.Settings.ExamRatios[e] : 0):0.##}%)"); }
            cols.AddRange(new[] { "최종 점수", "학년 등수", "반 등수", "성취도", "성취도 경계", "비고" });
            int hr = s.Header(cols.ToArray());
            foreach (var st in r.Students)
            {
                var vals = new List<object> { (double)st.ClassNo, (double)st.Number, st.Name };
                for (int i = 0; i < r.Settings.PerfAreas.Count; i++)
                {
                    if (st.Perf == null || i >= st.Perf.Scores.Length) vals.Add("");
                    else if (st.Perf.Exempt[i]) vals.Add("면제");
                    else vals.Add(N(st.Perf.Scores[i]));
                }
                if (r.HasPerf) vals.Add(N(st.PerfTotal));
                foreach (var e in r.Exams)
                {
                    AnswerRow row;
                    if (st.Written.TryGetValue(e, out row)) { vals.Add(row.Total); vals.Add(N2(st.WrittenConverted.ContainsKey(e) ? st.WrittenConverted[e] : (double?)null)); }
                    else { vals.Add(st.WrittenExemptByTable.Contains(e) ? "면제" : "결시"); vals.Add(""); }
                }
                vals.Add(N2(st.Final));
                vals.Add(st.GradeRank.HasValue ? (object)(double)st.GradeRank.Value : "");
                vals.Add(st.ClassRank.HasValue ? (object)(double)st.ClassRank.Value : "");
                vals.Add(st.Achievement ?? "");
                vals.Add(st.BoundaryNote ?? "");
                vals.Add(st.Included ? (st.Status == "정상" ? "" : st.Status) : "제외: " + st.Status);
                s.Cells(!st.Included ? Gray : st.Boundary ? Yellow : (Color?)null, vals.ToArray());
            }
            s.Freeze(hr, 3);
            s.RepeatHeader(hr);
            s.AutoWidths(cols.Count, 7, 30);
            s.Ws.Column(cols.Count).Width = 36;
            s.Explain("성취도", "성취도경계", "등수");
        }

        // ===================== 3. 등수표 =====================
        static void WriteRanking(ExcelPackage pkg, AnalysisResult r)
        {
            var s = new Sheet(pkg, "3.등수표", "등수표 (최종 점수 순)", "최종 점수가 높은 순서입니다. 동점이면 같은 등수이고 다음 등수는 건너뜁니다.");
            s.Ws.PrinterSettings.Orientation = eOrientation.Portrait;
            int hr = s.Header("학년 등수", "반", "번호", "이름", "수행 합계", "지필 환산", "최종 점수", "성취도", "반 등수");
            foreach (var st in r.Ranked)
                s.Cells(st.Boundary ? Yellow : (Color?)null, (double)st.GradeRank.Value, (double)st.ClassNo, (double)st.Number, st.Name, N(st.PerfTotal), F2(st.WrittenConverted.Values.Sum()), F2(st.Final.Value), st.Achievement, (double)st.ClassRank.Value);
            s.Freeze(hr);
            s.RepeatHeader(hr);
            s.Widths(10, 6, 6, 12, 10, 10, 10, 8, 8);
            s.Explain("등수", "성취도");
        }

        // ===================== 4. 성취도 분석 =====================
        static void WriteAchievement(ExcelPackage pkg, AnalysisResult r)
        {
            var s = new Sheet(pkg, "4.성취도 분석", "성취도 분석", $"학년·반별 A~E 인원과 비율입니다. 기준: A {r.Settings.A:0.##} / B {r.Settings.B:0.##} / C {r.Settings.C:0.##} / D {r.Settings.D:0.##} 이상, 그 미만 E.");
            s.Section("학년·반별 성취도 인원");
            int hr = s.Header("구분", "인원", "A", "B", "C", "D", "E", "A 비율", "B 비율", "C 비율", "D 비율", "E 비율");
            int firstData = s.Row;
            Action<string, int, Dictionary<string, int>> line = (label, n, d) =>
            {
                var v = new List<object> { label, (double)n };
                foreach (var g in Analyzer.Grades) v.Add((double)(d.ContainsKey(g) ? d[g] : 0));
                foreach (var g in Analyzer.Grades) v.Add(P(n == 0 ? 0 : (double)(d.ContainsKey(g) ? d[g] : 0) / n));
                s.Cells(v.ToArray());
            };
            line("학년 전체", r.Count, r.AchievementCounts);
            for (int c = 1; c <= 12; c++) s.Bold(s.Row - 1, c);
            foreach (var c in r.ClassSummaries) line(ClassLabel(c.ClassNo), c.Count, c.AchievementCounts);
            int lastData = s.Row - 1;
            s.Widths(12, 8, 7, 7, 7, 7, 7, 9, 9, 9, 9, 9);

            if (r.ClassSummaries.Count > 0)
            {
                var ch = s.ChartBelow("ach", eChartType.ColumnClustered, "반별 성취도 인원", 720, 340);
                for (int i = 0; i < 5; i++)
                {
                    var ser = ch.Series.Add(s.Ws.Cells[firstData + 1, 3 + i, lastData, 3 + i], s.Ws.Cells[firstData + 1, 1, lastData, 1]);
                    ser.Header = Analyzer.Grades[i];
                }
                var pie = s.ChartBelow("achpie", eChartType.Pie, "학년 전체 성취도 비율", 420, 340);
                var ps = pie.Series.Add(s.Ws.Cells[firstData, 3, firstData, 7], s.Ws.Cells[hr, 3, hr, 7]);
                ((ExcelPieChart)pie).DataLabel.ShowPercent = true; ((ExcelPieChart)pie).DataLabel.ShowCategory = true;
            }

            // 성취도별 평균·표준편차 (KICE 성취도분석표 양식)
            s.Section("성취도별 인원·평균·표준편차  (교육과정평가원 성취도분석표와 같은 구성)");
            var lcols = new List<string> { "성취도", "인원", "비율", "최종 점수\n평균", "최종 점수\n표준편차", "최저", "최고" };
            if (r.HasPerf) lcols.Add($"수행 평균\n(만점 {r.PerfMax:0.##})");
            foreach (var e in r.Exams) lcols.Add($"{e}\n원점수 평균");
            s.Header(lcols.ToArray());
            foreach (var ls in r.LevelSummaries)
            {
                var v = new List<object> { ls.Grade, (double)ls.Count, P(r.Count == 0 ? 0 : (double)ls.Count / r.Count) };
                if (ls.Count > 0) { v.Add(F1(ls.Mean)); v.Add(F1(ls.Sd)); v.Add(F1(ls.Min)); v.Add(F1(ls.Max)); } else { v.Add("-"); v.Add("-"); v.Add("-"); v.Add("-"); }
                if (r.HasPerf) v.Add(N1(ls.PerfMean));
                foreach (var e in r.Exams) v.Add(ls.WrittenMeans.ContainsKey(e) ? (object)F1(ls.WrittenMeans[e]) : "-");
                s.Cells(v.ToArray());
            }
            {
                var v = new List<object> { "전체", (double)r.Count, P(1), F1(r.Mean), F1(r.Sd), N1(r.Count > 0 ? r.Min : (double?)null), N1(r.Count > 0 ? r.Max : (double?)null) };
                if (r.HasPerf) { var pt = r.Ranked.Where(x => x.PerfTotal.HasValue).Select(x => x.PerfTotal.Value).ToList(); v.Add(pt.Count > 0 ? (object)F1(pt.Average()) : "-"); }
                foreach (var e in r.Exams) { var w = r.Ranked.Where(x => x.Written.ContainsKey(e)).Select(x => x.Written[e].Total).ToList(); v.Add(w.Count > 0 ? (object)F1(w.Average()) : "-"); }
                s.Cells(LightBlue, v.ToArray());
            }
            for (int c = 1; c <= lcols.Count; c++) if (s.Ws.Column(c).Width < 11) s.Ws.Column(c).Width = 11;

            s.Section($"성취도 경계 학생 (기준 점수 ±{r.Settings.BoundaryRange:0.##}점)");
            s.Header("반", "번호", "이름", "최종 점수", "성취도", "경계", "학년 등수");
            foreach (var st in r.Ranked.Where(x => x.Boundary).OrderByDescending(x => x.Final))
                s.Cells(Yellow, (double)st.ClassNo, (double)st.Number, st.Name, F2(st.Final.Value), st.Achievement, st.BoundaryNote, (double)st.GradeRank.Value);
            if (!r.Ranked.Any(x => x.Boundary)) s.Text("경계에 있는 학생이 없습니다.");
            s.Explain("성취도", "성취도별통계", "성취도경계");
        }

        // ===================== 5. 반별 비교 =====================
        static void WriteClassCompare(ExcelPackage pkg, AnalysisResult r)
        {
            var s = new Sheet(pkg, "5.반별 비교", "반별 비교", "반별 최종 점수 통계와 수행·지필 평균입니다. 표준편차가 크면 반 안의 점수 차이가 크다는 뜻입니다.");
            var cols = new List<string> { "반", "명단", "분석 인원", "평균", "중앙값", "표준편차", "최고점", "최저점", "평균 순위" };
            foreach (var a in r.Settings.PerfAreas) cols.Add($"{a.Name}\n평균({a.Max:0.##})");
            if (r.HasPerf) cols.Add("수행 합계\n평균");
            foreach (var e in r.Exams) cols.Add($"{e}\n원점수 평균");
            int hr = s.Header(cols.ToArray());
            int first = s.Row;
            foreach (var c in r.ClassSummaries)
            {
                var v = new List<object> { ClassLabel(c.ClassNo), (double)c.Enrolled, (double)c.Count, F1(c.Mean), F1(c.Median), F1(c.Sd), N(c.Max), N(c.Min), c.MeanRank == 0 ? "" : (object)(double)c.MeanRank };
                foreach (var m in c.PerfAreaMeans) v.Add(F1(m));
                if (r.HasPerf) v.Add(F1(c.PerfMean));
                foreach (var e in r.Exams) v.Add(F1(c.WrittenMeans.ContainsKey(e) ? c.WrittenMeans[e] : 0));
                s.Cells(v.ToArray());
            }
            int last = s.Row - 1;
            {
                var v = new List<object> { "학년 전체", (double)r.Students.Count, (double)r.Count, F1(r.Mean), F1(r.Median), F1(r.Sd), N(r.Max), N(r.Min), "" };
                for (int i = 0; i < r.Settings.PerfAreas.Count; i++)
                {
                    var vals = r.Ranked.Where(x => x.Perf != null && i < x.Perf.Scores.Length && x.Perf.Scores[i].HasValue).Select(x => x.Perf.Scores[i].Value).ToList();
                    v.Add(F1(vals.Count > 0 ? vals.Average() : 0));
                }
                if (r.HasPerf) { var pt = r.Ranked.Where(x => x.PerfTotal.HasValue).Select(x => x.PerfTotal.Value).ToList(); v.Add(F1(pt.Count > 0 ? pt.Average() : 0)); }
                foreach (var ea in r.ExamAnalyses) v.Add(F1(ea.Mean));
                s.Cells(LightBlue, v.ToArray());
                for (int c = 1; c <= cols.Count; c++) s.Bold(s.Row - 1, c);
            }
            s.AutoWidths(cols.Count, 8, 16);
            if (r.ClassSummaries.Count > 0)
            {
                var ch = s.ChartBelow("cls", eChartType.ColumnClustered, "반별 최종 점수 평균", 640, 320);
                var ser = ch.Series.Add(s.Ws.Cells[first, 4, last, 4], s.Ws.Cells[first, 1, last, 1]); ser.Header = "평균";
                ch.Legend.Remove();
                int wc = 10 + r.Settings.PerfAreas.Count + (r.HasPerf ? 1 : 0);
                if (r.HasPerf && r.Exams.Count > 0)
                {
                    var ch2 = s.ChartBelow("cls2", eChartType.ColumnClustered, "반별 수행 합계 평균 · 지필 원점수 평균", 640, 320);
                    var s1 = ch2.Series.Add(s.Ws.Cells[first, wc - 1, last, wc - 1], s.Ws.Cells[first, 1, last, 1]); s1.Header = "수행 합계 평균";
                    var s2 = ch2.Series.Add(s.Ws.Cells[first, wc, last, wc], s.Ws.Cells[first, 1, last, 1]); s2.Header = r.Exams[0] + " 평균";
                }
            }
            s.Explain("평균", "중앙값", "표준편차", "최고최저");
        }

        // ===================== 6. 점수 분포 =====================
        static void WriteDistribution(ExcelPackage pkg, AnalysisResult r)
        {
            var s = new Sheet(pkg, "6.점수 분포", "점수 분포 (10점 구간)", "최종 점수를 10점 구간으로 나눈 인원입니다. 90~100 구간에는 100점이 포함됩니다.");
            var cols = new List<string> { "점수 구간", "학년 전체", "비율" };
            cols.AddRange(r.Classes.Select(ClassLabel));
            int hr = s.Header(cols.ToArray());
            int first = s.Row;
            for (int b = 9; b >= 0; b--)
            {
                var v = new List<object> { Analyzer.BinLabel(b), (double)r.Histogram[b], P(r.Count == 0 ? 0 : (double)r.Histogram[b] / r.Count) };
                foreach (var c in r.Classes) v.Add((double)r.ClassHistogram[c][b]);
                s.Cells(v.ToArray());
            }
            int last = s.Row - 1;
            s.AutoWidths(cols.Count, 8, 14);
            var ch = s.ChartBelow("hist", eChartType.ColumnClustered, "학년 전체 점수 분포", 640, 320);
            var ser = ch.Series.Add(s.Ws.Cells[first, 2, last, 2], s.Ws.Cells[first, 1, last, 1]); ser.Header = "인원";
            ch.Legend.Remove();
            if (r.Classes.Count > 1)
            {
                var ch2 = s.ChartBelow("hist2", eChartType.ColumnClustered, "반별 점수 분포", 800, 340);
                for (int i = 0; i < r.Classes.Count; i++)
                {
                    var sr = ch2.Series.Add(s.Ws.Cells[first, 4 + i, last, 4 + i], s.Ws.Cells[first, 1, last, 1]); sr.Header = ClassLabel(r.Classes[i]);
                }
            }
            s.Explain("점수분포");
        }

        // ===================== 7. 수행평가 분석 =====================
        static void WritePerf(ExcelPackage pkg, AnalysisResult r)
        {
            var s = new Sheet(pkg, "7.수행평가 분석", "수행평가 분석", "영역별 평균과 만점·최저점 학생 비율입니다. 만점자 비율이 아주 높으면 변별이 잘 안 되는 영역일 수 있습니다.");
            s.Section("영역별 통계 (분석 대상 학생 기준)");
            s.Header("영역", "만점", "반영 비율", "인원", "평균", "평균 비율", "표준편차", "만점자 수", "만점자 비율", "최저점", "최저점 인원", "최저점 비율", "면제 인원");
            for (int i = 0; i < r.Settings.PerfAreas.Count; i++)
            {
                var a = r.Settings.PerfAreas[i];
                var vals = r.Ranked.Where(x => x.Perf != null && i < x.Perf.Scores.Length && x.Perf.Scores[i].HasValue).Select(x => x.Perf.Scores[i].Value).ToList();
                int exempt = r.Students.Count(x => x.Perf != null && i < x.Perf.Exempt.Length && x.Perf.Exempt[i]);
                if (vals.Count == 0) { s.Cells(a.Name, a.Max, P(a.Ratio / 100), 0.0, "", "", "", "", "", "", "", "", (double)exempt); continue; }
                int full = vals.Count(v => Math.Abs(v - a.Max) < 1e-9);
                double min = vals.Min(); int minN = vals.Count(v => Math.Abs(v - min) < 1e-9);
                s.Cells(a.Name, a.Max, P(a.Ratio / 100), (double)vals.Count, F1(vals.Average()), P(a.Max == 0 ? 0 : vals.Average() / a.Max), F1(Analyzer.Sd(vals)), (double)full, P((double)full / vals.Count), min, (double)minN, P((double)minN / vals.Count), (double)exempt);
            }
            s.Section("반별 영역 평균");
            var cols = new List<string> { "반", "인원" }; cols.AddRange(r.Settings.PerfAreas.Select(a => a.Name)); cols.Add("수행 합계");
            int hr = s.Header(cols.ToArray());
            int first = s.Row;
            foreach (var c in r.ClassSummaries)
            {
                var v = new List<object> { ClassLabel(c.ClassNo), (double)c.Count }; foreach (var m in c.PerfAreaMeans) v.Add(F1(m)); v.Add(F1(c.PerfMean));
                s.Cells(v.ToArray());
            }
            int last = s.Row - 1;
            s.Section("영역별 점수 분포 (점수별 인원)");
            for (int i = 0; i < r.Settings.PerfAreas.Count; i++)
            {
                var a = r.Settings.PerfAreas[i];
                var vals = r.Ranked.Where(x => x.Perf != null && i < x.Perf.Scores.Length && x.Perf.Scores[i].HasValue).Select(x => x.Perf.Scores[i].Value).ToList();
                var groups = vals.GroupBy(v => v).OrderByDescending(g => g.Key).ToList();
                s.Text(a.Name, true);
                s.Header("점수", "인원", "비율");
                foreach (var g in groups) s.Cells(g.Key, (double)g.Count(), P(vals.Count == 0 ? 0 : (double)g.Count() / vals.Count));
            }
            s.AutoWidths(13, 8, 20);
            if (r.ClassSummaries.Count > 0 && r.Settings.PerfAreas.Count > 0)
            {
                var ch = s.ChartBelow("perf", eChartType.ColumnClustered, "반별 수행 영역 평균", 640, 320);
                for (int i = 0; i < r.Settings.PerfAreas.Count; i++)
                {
                    var sr = ch.Series.Add(s.Ws.Cells[first, 3 + i, last, 3 + i], s.Ws.Cells[first, 1, last, 1]); sr.Header = r.Settings.PerfAreas[i].Name;
                }
            }
            s.Explain("만점자비율", "평균", "표준편차");
        }

        // ===================== 8. 문항 분석 =====================
        static void WriteItems(ExcelPackage pkg, AnalysisResult r, ExamAnalysis ea)
        {
            string suffix = r.Exams.Count > 1 ? "-" + ea.ExamName.Replace("학년", "").Replace(" ", "") : "";
            var s = new Sheet(pkg, Trunc("8.문항 분석" + suffix), $"문항 분석 - {ea.ExamName} (응시 {ea.N}명)", "정답률이 낮을수록 어려운 문항, 변별도가 0.3 이상이면 잘 가르는 문항입니다. 빨간색은 변별도가 0.1 미만이거나 음수인 문항입니다.");
            if (ea.KR20.HasValue)
                s.Text($"시험 신뢰도 KR-20 = {ea.KR20:0.00}  → {Glossary.KR20(ea.KR20.Value)}.  (시험이 실력을 일관되게 재는 정도, 0~1. 0.90 이상 매우 높음 / 0.80 양호 / 0.70 보통 / 0.60 미만 낮음)", true, Blue);
            s.Text($"지필 원점수 평균 {ea.Mean:0.0}점, 표준편차 {ea.Sd:0.0}점 → {Glossary.Sd(ea.Sd, ea.TotalMax)}  (만점 {ea.TotalMax:0.##}점, 선택형 {ea.ChoiceMax:0.##}점, 서답형 {ea.EssayMax:0.##}점)");
            var choiceItems = ea.ItemStats.Where(i => i.Correct.HasValue).ToList();
            if (choiceItems.Count > 0)
            {
                double avgP = choiceItems.Average(i => i.CorrectRate), avgD = choiceItems.Average(i => i.Discrimination ?? 0);
                s.Text($"평균 정답률 {avgP:P0} → {(avgP >= 0.8 ? "쉬운 시험" : avgP >= 0.6 ? "적절한 난이도" : avgP >= 0.45 ? "어려운 편" : "매우 어려운 시험")} (60~70%가 적절)   /   평균 변별도 {avgD:0.00} → {Glossary.Discrimination(avgD)}   /   변별도 0.2 미만 문항 {choiceItems.Count(i => (i.Discrimination ?? 0) < 0.2)}개, 음수 {choiceItems.Count(i => (i.Discrimination ?? 0) < 0)}개");
                var hardest = choiceItems.OrderBy(i => i.CorrectRate).Take(3).Select(i => $"{i.Number}번({i.CorrectRate:P0})");
                var easiest = choiceItems.OrderByDescending(i => i.CorrectRate).Take(3).Select(i => $"{i.Number}번({i.CorrectRate:P0})");
                int inv = choiceItems.Count(i => i.LevelInversion != null);
                s.Text($"정답률이 가장 낮은 문항: {string.Join(", ", hardest)}   /   가장 높은 문항: {string.Join(", ", easiest)}   /   성취도 수준 역전 문항 {inv}개" + (inv > 0 ? " → '성취도별 문항' 시트에서 확인" : ""));
            }
            s.Text("색: 빨강 = 변별도 0.1 미만(점검), 노랑 = 어려운 문항(정답률 50% 미만). 변별도는 0.40 이상 매우 좋음 / 0.30 좋음 / 0.20 보통 / 0.10 낮음 / 음수 점검.", false, Color.FromArgb(89, 89, 89));
            s.Row++;
            int hr = s.Header("문항", "정답", "배점", "정답률", "오답률", "무응답률", "복수답안", "난이도", "상위 27%\n정답률", "하위 27%\n정답률", "변별도\n(상-하)", "변별도\n해석", "문항-총점\n상관", "상관\n해석", "매력적 오답\n(번호)", "매력적 오답\n선택 비율", "종합 판정");
            foreach (var it in ea.ItemStats)
            {
                if (!it.Correct.HasValue) { s.Cells((double)it.Number, it.Points > 0 ? "서답형" : "-", it.Points, "", "", "", "", "", "", "", "", "", "", "", "", "", "문항별 자료 없음(서답형)"); continue; }
                Color? fill = it.Discrimination < 0.1 ? Red : it.Difficulty == "어려움" ? Yellow : (Color?)null;
                s.Cells(fill, (double)it.Number, (double)it.Correct.Value, it.Points, P(it.CorrectRate), P(it.WrongRate), P(it.BlankRate), P(it.MultiRate), it.Difficulty,
                    P(it.UpperRate ?? 0), P(it.LowerRate ?? 0), F2(it.Discrimination ?? 0), Glossary.Discrimination(it.Discrimination ?? 0),
                    it.PointBiserial.HasValue ? (object)F2(it.PointBiserial.Value) : "-", it.PointBiserial.HasValue ? Glossary.PointBiserial(it.PointBiserial.Value) : "계산 불가",
                    it.AttractiveWrong.HasValue ? (object)(double)it.AttractiveWrong.Value : "-", P(it.AttractiveWrongRate), Glossary.ItemVerdict(it));
            }
            s.Freeze(hr, 1);
            s.Widths(6, 6, 6, 9, 9, 9, 9, 8, 10, 10, 10, 16, 10, 10, 11, 11, 40);
            int first = hr + 1, last = s.Row - 1;
            var ch = s.ChartBelow("items", eChartType.ColumnClustered, "문항별 정답률", 900, 320);
            var sr = ch.Series.Add(s.Ws.Cells[first, 4, last, 4], s.Ws.Cells[first, 1, last, 1]); sr.Header = "정답률";
            ch.Legend.Remove();
            var ch2 = s.ChartBelow("disc", eChartType.ColumnClustered, "문항별 변별도 (0.3 이상이면 좋음, 음수는 점검)", 900, 320);
            var sr2 = ch2.Series.Add(s.Ws.Cells[first, 11, last, 11], s.Ws.Cells[first, 1, last, 1]); sr2.Header = "변별도";
            ch2.Legend.Remove();
            s.Explain("정답률", "난이도구분", "상하위정답률", "변별도", "점이연상관", "매력적오답", "KR20", "오답률");

            // ---- 답지 반응 분포 ----
            var s2 = new Sheet(pkg, Trunc("8-2.답지 반응" + suffix), $"답지 반응 분포 - {ea.ExamName}", "각 문항에서 ①~⑤를 고른 학생 비율입니다. 파란색이 정답, 빨간색은 가장 많이 고른 오답(매력적 오답)입니다.");
            var cols = new List<string> { "문항", "정답" };
            for (int o = 1; o <= ea.OptionCount; o++) cols.Add(Circled(o));
            cols.AddRange(new[] { "복수답안", "무표기", "정답률" });
            int hr2 = s2.Header(cols.ToArray());
            foreach (var it in ea.ItemStats.Where(i => i.Correct.HasValue))
            {
                var v = new List<object> { (double)it.Number, (double)it.Correct.Value };
                for (int o = 1; o <= ea.OptionCount; o++) v.Add(P(it.N == 0 ? 0 : (double)it.OptionCounts[o] / it.N));
                v.Add(P(it.MultiRate)); v.Add(P(it.BlankRate)); v.Add(P(it.CorrectRate));
                s2.Cells(v.ToArray());
                int row = s2.Row - 1;
                s2.Fill(row, 2 + it.Correct.Value, LightBlue); s2.Bold(row, 2 + it.Correct.Value);
                if (it.AttractiveWrong.HasValue && it.AttractiveWrongRate >= 0.15) s2.Fill(row, 2 + it.AttractiveWrong.Value, Red);
            }
            s2.Row++;
            s2.Text("인원수 기준", true);
            var cols2 = new List<string> { "문항", "정답" };
            for (int o = 1; o <= ea.OptionCount; o++) cols2.Add(Circled(o));
            cols2.AddRange(new[] { "복수답안", "무표기", "합계" });
            s2.Header(cols2.ToArray());
            foreach (var it in ea.ItemStats.Where(i => i.Correct.HasValue))
            {
                var v = new List<object> { (double)it.Number, (double)it.Correct.Value };
                for (int o = 1; o <= ea.OptionCount; o++) v.Add((double)it.OptionCounts[o]);
                v.Add((double)it.MultiCount); v.Add((double)it.BlankCount); v.Add((double)it.N);
                s2.Cells(v.ToArray());
            }
            s2.Freeze(hr2, 2);
            s2.Widths(6, 6, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9);
            s2.Explain("답지반응", "매력적오답", "오답률");

            // ---- 반별 문항 정답률 ----
            var s3 = new Sheet(pkg, Trunc("8-3.반별 문항" + suffix), $"반별 문항 정답률 비교 - {ea.ExamName}", $"반별 정답률이 학년 정답률보다 {r.Settings.ClassItemGap:P0} 이상 낮으면 빨간색, {r.Settings.ClassItemGap:P0} 이상 높으면 초록색입니다. 수업 보완이 필요한 반-단원을 찾는 데 쓰세요.");
            var classes = ea.ItemStats.SelectMany(i => i.ClassN.Keys).Distinct().OrderBy(c => c).ToList();
            var cols3 = new List<string> { "문항", "정답", "학년 정답률" }; cols3.AddRange(classes.Select(ClassLabel)); cols3.Add("최저 반"); cols3.Add("최고-최저 차이");
            int hr3 = s3.Header(cols3.ToArray());
            foreach (var it in ea.ItemStats.Where(i => i.Correct.HasValue))
            {
                var v = new List<object> { (double)it.Number, (double)it.Correct.Value, P(it.CorrectRate) };
                foreach (var c in classes) v.Add(it.ClassCorrectRate.ContainsKey(c) ? (object)P(it.ClassCorrectRate[c]) : "");
                var rates = classes.Where(c => it.ClassCorrectRate.ContainsKey(c)).Select(c => (c, it.ClassCorrectRate[c])).ToList();
                if (rates.Count > 0) { var mn = rates.OrderBy(x => x.Item2).First(); v.Add(ClassLabel(mn.Item1)); v.Add(P(rates.Max(x => x.Item2) - mn.Item2)); } else { v.Add(""); v.Add(""); }
                s3.Cells(v.ToArray());
                int row = s3.Row - 1;
                for (int i = 0; i < classes.Count; i++)
                {
                    if (!it.ClassCorrectRate.ContainsKey(classes[i])) continue;
                    double diff = it.ClassCorrectRate[classes[i]] - it.CorrectRate;
                    if (diff <= -r.Settings.ClassItemGap) s3.Fill(row, 4 + i, Red);
                    else if (diff >= r.Settings.ClassItemGap) s3.Fill(row, 4 + i, Green);
                }
            }
            {
                var v = new List<object> { "반 평균 정답률", "", P(ea.ItemStats.Where(i => i.Correct.HasValue).Select(i => i.CorrectRate).DefaultIfEmpty(0).Average()) };
                foreach (var c in classes) v.Add(P(ea.ItemStats.Where(i => i.Correct.HasValue && i.ClassCorrectRate.ContainsKey(c)).Select(i => i.ClassCorrectRate[c]).DefaultIfEmpty(0).Average()));
                v.Add(""); v.Add("");
                s3.Cells(LightBlue, v.ToArray());
            }
            s3.Freeze(hr3, 3);
            s3.Widths(6, 6, 11, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9);
            s3.Ws.Column(cols3.Count - 1).Width = 9; s3.Ws.Column(cols3.Count).Width = 13;
            s3.Explain("반별문항", "정답률");

            // ---- 성취도별 문항 정답률 ----
            var s4 = new Sheet(pkg, Trunc("8-4.성취도별 문항" + suffix), $"성취도별 문항 정답률 - {ea.ExamName}", "성취도 A~E 집단이 각 문항을 맞힌 비율입니다. A→E로 갈수록 낮아지는 것이 정상이며, 낮은 집단이 높은 집단보다 15%p 이상 높으면 '수준 역전'(노란색)으로 표시해 문항 검토가 필요함을 알립니다.");
            s4.Text("정답률 50% 미만 칸은 연한 빨간색입니다. 성취도는 최종 점수(수행+지필) 기준이며, 인원 5명 미만인 집단은 역전 판정에서 제외합니다.", false, Color.FromArgb(89, 89, 89));
            s4.Row++;
            var levels = Analyzer.Grades.ToList();
            var cols4 = new List<string> { "문항", "정답", "학년\n정답률" };
            foreach (var g in levels) { int n = r.LevelSummaries.First(x => x.Grade == g).Count; cols4.Add($"{g}\n({n}명)"); }
            cols4.Add("A−E 차이"); cols4.Add("비고");
            int hr4 = s4.Header(cols4.ToArray());
            foreach (var it in ea.ItemStats.Where(i => i.Correct.HasValue))
            {
                var v = new List<object> { (double)it.Number, (double)it.Correct.Value, P(it.CorrectRate) };
                foreach (var g in levels) v.Add(it.LevelN.ContainsKey(g) && it.LevelN[g] > 0 ? (object)P(it.LevelCorrectRate[g]) : "-");
                bool hasA = it.LevelN.ContainsKey("A") && it.LevelN["A"] > 0, hasE = it.LevelN.ContainsKey("E") && it.LevelN["E"] > 0;
                v.Add(hasA && hasE ? (object)new Fmt(it.LevelCorrectRate["A"] - it.LevelCorrectRate["E"], "+0%;-0%;0%") : "-");
                v.Add(it.LevelInversion ?? "");
                s4.Cells(it.LevelInversion != null ? Yellow : (Color?)null, v.ToArray());
                int row = s4.Row - 1;
                for (int i = 0; i < levels.Count; i++)
                    if (it.LevelN.ContainsKey(levels[i]) && it.LevelN[levels[i]] > 0 && it.LevelCorrectRate[levels[i]] < 0.5 && it.LevelInversion == null) s4.Fill(row, 4 + i, Red);
            }
            {
                var v = new List<object> { "평균 정답률", "", P(ea.ItemStats.Where(i => i.Correct.HasValue).Select(i => i.CorrectRate).DefaultIfEmpty(0).Average()) };
                foreach (var g in levels)
                {
                    var rates = ea.ItemStats.Where(i => i.Correct.HasValue && i.LevelN.ContainsKey(g) && i.LevelN[g] > 0).Select(i => i.LevelCorrectRate[g]).ToList();
                    v.Add(rates.Count > 0 ? (object)P(rates.Average()) : "-");
                }
                v.Add(""); v.Add("");
                s4.Cells(LightBlue, v.ToArray());
            }
            s4.Freeze(hr4, 3);
            s4.Widths(6, 6, 10, 10, 10, 10, 10, 10, 11, 60);
            s4.Explain("성취도별문항", "성취도", "정답률");
        }

        // ===================== 9. 수행 vs 지필 =====================
        static void WritePerfVsWritten(ExcelPackage pkg, AnalysisResult r)
        {
            var s = new Sheet(pkg, "9.수행 vs 지필", "수행평가 vs 지필평가 비교", $"두 평가를 100점 만점 비율로 바꿔 비교합니다. 수행은 대체로 지필보다 높으므로 학년 평균 차이를 뺀 '상대 차이'를 보세요. 상대 차이가 ±{Analyzer.GapCut}점 이상이면 노란색입니다.");
            s.Text($"상관계수(피어슨) = {(r.PerfWrittenCorr.HasValue ? r.PerfWrittenCorr.Value.ToString("0.00") + " → " + Glossary.Correlation(r.PerfWrittenCorr.Value) : "계산 불가")}   (0.7 이상 강함 / 0.4 중간 / 0.2 약함 / 그 미만 거의 없음)   /   학년 평균 차이(수행 비율 − 지필 비율) = {r.AvgPerfWrittenGap:+0.0;-0.0}점", true, Blue);
            s.Row++;
            int chartRow = s.Row; s.Row += 23;   // 산점도 자리를 표 위에 비워 둔다
            int hr = s.Header("반", "번호", "이름", "수행 합계", "수행 비율(%)", "지필 원점수", "지필 비율(%)", "차이\n(수행-지필)", "상대 차이\n(학년 평균 차이 제외)", "해석");
            int first = s.Row;
            foreach (var st in r.Ranked.Where(x => x.PerfWrittenGapAdj.HasValue).OrderByDescending(x => Math.Abs(x.PerfWrittenGapAdj.Value)))
            {
                double gap = st.PerfWrittenGap.Value, adj = st.PerfWrittenGapAdj.Value;
                string note = adj >= Analyzer.GapCut ? "수행에 비해 지필이 유난히 낮음" : adj <= -Analyzer.GapCut ? "지필에 비해 수행이 유난히 낮음" : adj >= Analyzer.GapCut * 0.6 ? "지필이 다소 낮은 편" : adj <= -Analyzer.GapCut * 0.6 ? "수행이 다소 낮은 편" : "";
                s.Cells(Math.Abs(adj) >= Analyzer.GapCut ? Yellow : (Color?)null, (double)st.ClassNo, (double)st.Number, st.Name, N(st.PerfTotal), F1(st.PerfPercent.Value), N(st.WrittenRawSum), F1(st.WrittenPercent.Value), new Fmt(gap, "+0.0;-0.0;0.0"), new Fmt(adj, "+0.0;-0.0;0.0"), note);
            }
            int last = s.Row - 1;
            s.Freeze(hr);
            s.Widths(6, 6, 12, 10, 12, 11, 12, 12, 16, 28);
            if (last >= first)
            {
                var ch = s.Chart("scatter", eChartType.XYScatter, "수행 비율(가로) vs 지필 비율(세로) — 점 하나가 학생 한 명", chartRow, 1, 560, 420);
                var sr = ch.Series.Add(s.Ws.Cells[first, 7, last, 7], s.Ws.Cells[first, 5, last, 5]); sr.Header = "학생";
                ch.Legend.Remove();
                ((ExcelScatterChartSerie)sr).Marker = eMarkerStyle.Circle;
                ch.XAxis.Title.Text = "수행 비율(%)"; ch.YAxis.Title.Text = "지필 비율(%)";
                ch.XAxis.MinValue = 0; ch.XAxis.MaxValue = 100; ch.YAxis.MinValue = 0; ch.YAxis.MaxValue = 100;
            }
            s.Explain("상관계수", "상대차이");
        }

        // ===================== 10. 관심 학생 =====================
        static void WriteWatch(ExcelPackage pkg, AnalysisResult r)
        {
            var s = new Sheet(pkg, "10.관심 학생", "관심 학생", "상담이나 확인이 필요한 학생 목록입니다. 사유를 보고 필요한 학생만 골라 보세요.");
            int hr = s.Header("반", "번호", "이름", "최종 점수", "성취도", "학년 등수", "무표기", "사유");
            foreach (var st in r.Students.Where(x => x.Flags.Count > 0))
            {
                Color? fill = !st.Included ? Gray : st.Achievement == "E" ? Red : st.Boundary ? Yellow : (Color?)null;
                s.Cells(fill, (double)st.ClassNo, (double)st.Number, st.Name, N2(st.Final), st.Achievement ?? "", st.GradeRank.HasValue ? (object)(double)st.GradeRank.Value : "", (double)st.BlankCount, string.Join(" / ", st.Flags));
            }
            if (!r.Students.Any(x => x.Flags.Count > 0)) s.Text("해당 학생이 없습니다.");
            s.Section("제외 학생 (등수·통계에 넣지 않은 학생)");
            s.Header("반", "번호", "이름", "제외 사유", "수행 합계", "지필 원점수");
            foreach (var st in r.Students.Where(x => !x.Included))
                s.Cells(Gray, (double)st.ClassNo, (double)st.Number, st.Name, st.Status, st.PerfAllExempt ? "면제" : N(st.PerfTotal), st.Written.Count == 0 ? "없음" : string.Join(", ", st.Written.Select(w => $"{w.Key} {w.Value.Total:0.##}")));
            if (!r.Students.Any(x => !x.Included)) s.Text("제외된 학생이 없습니다.");
            s.Freeze(hr);
            s.Widths(6, 6, 12, 10, 8, 10, 8, 70);
            s.Explain("성취도경계", "상대차이", "오답률");
        }

        // ===================== 11. 중간→기말 변화 =====================
        static void WriteChange(ExcelPackage pkg, AnalysisResult r)
        {
            string e1 = r.Exams[0], e2 = r.Exams[r.Exams.Count - 1];
            var s = new Sheet(pkg, "11.점수 변화", $"{e1} → {e2} 점수 변화", "두 지필고사의 원점수 변화입니다. 15점 이상 오르면 초록색, 15점 이상 내리면 빨간색입니다.");
            var both = r.Students.Where(x => x.Written.ContainsKey(e1) && x.Written.ContainsKey(e2)).ToList();
            if (both.Count > 0)
            {
                double m1 = both.Average(x => x.Written[e1].Total), m2 = both.Average(x => x.Written[e2].Total);
                s.Text($"두 시험 모두 응시 {both.Count}명. {e1} 평균 {m1:0.0}점 → {e2} 평균 {m2:0.0}점 ({m2 - m1:+0.0;-0.0}점)", true, Blue);
                var corr = Analyzer.Pearson(both.Select(x => x.Written[e1].Total).ToList(), both.Select(x => x.Written[e2].Total).ToList());
                if (corr.HasValue) s.Text($"두 시험 점수의 상관계수 {corr:0.00} → {Glossary.Correlation(corr.Value)} (높을수록 중간에 잘한 학생이 기말도 잘 봄. 0.7 이상이면 순위가 거의 유지됨)");
            }
            s.Row++;
            int hr = s.Header("반", "번호", "이름", e1, e2, "변화", "비고");
            foreach (var st in both.OrderByDescending(x => x.Written[e2].Total - x.Written[e1].Total))
            {
                double d = st.Written[e2].Total - st.Written[e1].Total;
                s.Cells(d >= 15 ? Green : d <= -15 ? Red : (Color?)null, (double)st.ClassNo, (double)st.Number, st.Name, st.Written[e1].Total, st.Written[e2].Total, new Fmt(d, "+0.#;-0.#;0"), d >= 15 ? "크게 상승" : d <= -15 ? "크게 하락" : "");
            }
            s.Freeze(hr);
            s.Widths(6, 6, 12, 12, 12, 10, 12);
            s.Explain("점수변화", "상관계수");
        }

        // ===================== 12. 개인별 결과표 =====================
        static void WriteIndividual(ExcelPackage pkg, AnalysisResult r)
        {
            var s = new Sheet(pkg, "12.개인별 결과표", "학생 개인별 결과표", "학생 한 명이 한 블록입니다. 상담·인쇄용으로 쓰세요. 틀린 문항 뒤 괄호 안은 학생이 고른 번호입니다.");
            s.Ws.PrinterSettings.Orientation = eOrientation.Portrait;
            int count = 0;
            foreach (var st in r.Students)
            {
                int top = s.Row;
                s.Ws.Cells[s.Row, 1].Value = $"{st.ClassNo}반 {st.Number}번  {st.Name}";
                s.Ws.Cells[s.Row, 1, s.Row, 8].Merge = true;
                s.Ws.Cells[s.Row, 1].Style.Font.Bold = true; s.Ws.Cells[s.Row, 1].Style.Font.Size = 12;
                s.Ws.Cells[s.Row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid; s.Ws.Cells[s.Row, 1].Style.Fill.BackgroundColor.SetColor(LightBlue);
                s.Row++;
                var hdr = new List<string>(); var val = new List<object>();
                foreach (var a in r.Settings.PerfAreas) hdr.Add(a.Name);
                if (r.HasPerf) hdr.Add("수행 합계");
                foreach (var e in r.Exams) { hdr.Add(e + " 원점수"); hdr.Add(e + " 환산"); }
                hdr.AddRange(new[] { "최종 점수", "성취도", "학년 등수", "반 등수" });
                for (int i = 0; i < r.Settings.PerfAreas.Count; i++) val.Add(st.Perf == null || i >= st.Perf.Scores.Length ? "" : st.Perf.Exempt[i] ? "면제" : N(st.Perf.Scores[i]));
                if (r.HasPerf) val.Add(st.PerfAllExempt ? "면제" : N(st.PerfTotal));
                foreach (var e in r.Exams) { val.Add(st.Written.ContainsKey(e) ? (object)st.Written[e].Total : "결시/면제"); val.Add(st.WrittenConverted.ContainsKey(e) ? (object)F2(st.WrittenConverted[e]) : ""); }
                val.Add(N2(st.Final)); val.Add(st.Achievement ?? "-"); val.Add(st.GradeRank.HasValue ? (object)$"{st.GradeRank}/{r.Count}" : "-"); val.Add(st.ClassRank.HasValue ? (object)$"{st.ClassRank}/{r.ClassSummaries.First(c => c.ClassNo == st.ClassNo).Count}" : "-");
                s.Header(hdr.ToArray());
                s.Cells(val.ToArray());
                foreach (var e in r.Exams)
                {
                    AnswerRow row;
                    if (!st.Written.TryGetValue(e, out row)) continue;
                    var ea = r.ExamAnalyses.First(x => x.ExamName == e);
                    var wrong = new List<string>(); var blank = new List<string>();
                    for (int i = 0; i < ea.Items.Count && i < row.Answers.Length; i++)
                    {
                        var a = row.Answers[i];
                        if (!ea.Items[i].IsChoice || a == ".") continue;
                        if (a == "-" || a.Length == 0) blank.Add(ea.Items[i].Number.ToString());
                        else wrong.Add($"{ea.Items[i].Number}({a})");
                    }
                    s.Ws.Cells[s.Row, 1].Value = $"{e} 틀린 문항({wrong.Count}개): " + (wrong.Count == 0 ? "없음" : string.Join(", ", wrong)) + (blank.Count > 0 ? $"   무표기: {string.Join(", ", blank)}" : "");
                    s.Ws.Cells[s.Row, 1, s.Row, 8].Merge = true; s.Ws.Cells[s.Row, 1].Style.WrapText = true;
                    s.Ws.Row(s.Row).Height = wrong.Count > 12 ? 30 : 16;
                    s.Row++;
                }
                if (st.Flags.Count > 0) { s.Ws.Cells[s.Row, 1].Value = "참고: " + string.Join(" / ", st.Flags); s.Ws.Cells[s.Row, 1, s.Row, 8].Merge = true; s.Row++; }
                s.Row++;
                count++;
                if (count % 6 == 0) s.Ws.Row(s.Row - 1).PageBreak = true;
            }
            s.Widths(14, 14, 14, 14, 14, 14, 14, 14, 14, 14);
        }

        // ===================== 13. 검산 =====================
        static void WriteVerify(ExcelPackage pkg, AnalysisResult r)
        {
            var s = new Sheet(pkg, "13.검산", "검산 (정오표 계산값 vs 나이스 일람표)", "정오표에서 계산한 반별 응시생수·총점·평균이 나이스 일람표와 같은지 비교합니다. 모두 '일치'이면 계산이 정확하다는 뜻입니다.");
            foreach (var ea in r.ExamAnalyses.Where(e => e.Table != null))
            {
                s.Section($"{ea.ExamName} — 결과: {(ea.VerifyAllMatch == true ? "모두 일치 ✔" : "불일치 항목 있음 ✘")}");
                s.Header("반", "응시생수(계산)", "응시생수(일람표)", "총점(계산)", "총점(일람표)", "평균(계산)", "평균(일람표)", "결과");
                foreach (var v in ea.Verify)
                    s.Cells(v.Match ? Green : v.Result == "불일치" ? Red : Gray, ClassLabel(v.ClassNo), v.Result.StartsWith("정오표 없음") ? "-" : (object)(double)v.CalcCount, v.TableCount.HasValue ? (object)(double)v.TableCount.Value : "-", v.Result.StartsWith("정오표 없음") ? "-" : (object)v.CalcSum, N(v.TableSum), v.Result.StartsWith("정오표 없음") ? "-" : (object)F1(v.CalcMean), N1(v.TableMean), v.Result);
                if (ea.VerifyStudentIssues.Count > 0)
                {
                    s.Text("학생별 비교에서 발견된 차이:", true);
                    foreach (var i in ea.VerifyStudentIssues) s.Text("• " + i, false, Color.FromArgb(192, 0, 0));
                }
                else s.Text("학생별 점수도 모두 일치합니다.", false, Color.FromArgb(0, 112, 60));
            }
            s.Widths(8, 14, 15, 12, 13, 12, 13, 10);
            s.Explain("검산");
        }

        // ===================== 14. 분석 기준 안내 =====================
        static void WriteCriteria(ExcelPackage pkg, AnalysisResult r)
        {
            var s = new Sheet(pkg, "14.분석 기준", "분석 기준 안내", "이 결과 파일을 만들 때 사용한 비율·기준·반올림 방식과 제외 학생 사유입니다.");
            s.Ws.PrinterSettings.Orientation = eOrientation.Portrait;
            s.Section("계산 기준");
            foreach (var n in r.Notes) s.Text("• " + n);
            s.Section("제외한 학생과 이유");
            if (r.Students.Any(x => !x.Included))
                foreach (var st in r.Students.Where(x => !x.Included)) s.Text($"• {st.Label}: {st.Status}");
            else s.Text("없음");
            s.Section("사용한 파일");
            foreach (var f in r.FileList) s.Text("• " + f);
            if (r.Warnings.Count > 0)
            {
                s.Section("경고");
                foreach (var w in r.Warnings) s.Text("⚠ " + w, false, Color.FromArgb(192, 0, 0));
            }
            s.Section("프로그램 정보");
            s.Text("나이스 성적 분석 프로그램 (인터넷 연결 없이 이 컴퓨터 안에서만 계산했습니다. 학생 자료는 어디에도 전송되지 않습니다.)");
            s.Widths(120);
        }

        static string Circled(int o) => o >= 1 && o <= 9 ? ((char)('①' + o - 1)).ToString() : o.ToString();
        static string Trunc(string name) => name.Length > 31 ? name.Substring(0, 31) : name;
    }
}
