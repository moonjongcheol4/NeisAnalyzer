using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace NeisAnalyzer
{
    /// <summary>나이스에서 내려받은 엑셀 파일을 열어 종류를 알아내고 내용을 읽는다.</summary>
    public static class NeisReader
    {
        static readonly Regex RxClassNum = new Regex(@"^\s*(\d+)\s*/\s*(\d+)\s*$");

        public static DetectedFile Detect(string path)
        {
            var d = new DetectedFile { Path = path };
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xlsm")
            {
                d.Kind = NeisFileKind.Unknown;
                d.Error = ext == ".xls"
                    ? "옛 엑셀 형식(.xls)은 읽을 수 없습니다. 나이스 저장 화면에서 'XLS data'(.xlsx)를 골라 다시 저장하세요."
                    : $"엑셀(.xlsx) 파일이 아니라서 읽지 않습니다({ext}). PDF는 쓸 수 없고, 나이스에서 엑셀(XLS data)로 저장한 파일을 넣으세요.";
                return d;
            }
            try
            {
                using (var pkg = OpenReadOnly(path))
                {
                    var ws = pkg.Workbook.Worksheets.FirstOrDefault(w => w.Dimension != null);
                    if (ws == null) { d.Kind = NeisFileKind.Unknown; return d; }
                    var grid = new Grid(ws);
                    string title = grid.FindTextInRows(1, 8, t => t.Contains("정오표") || t.Contains("일람표") || Compact(t).Contains("문항정보표"));
                    if (title == null) { d.Kind = NeisFileKind.Unknown; return d; }
                    if (Compact(title).Contains("문항정보표")) { d.Kind = NeisFileKind.ItemInfoSheet; d.Info = ReadItemInfo(grid, path); d.Header = d.Info.Header; }
                    else if (title.Contains("정오표")) { d.Kind = NeisFileKind.AnswerSheet; d.Answer = ReadAnswerSheet(grid, path); d.Header = d.Answer.Header; }
                    else if (title.Contains("수행평가")) { d.Kind = NeisFileKind.PerfSheet; d.Perf = ReadPerf(grid, path); d.Header = d.Perf.Header; }
                    else if (title.Contains("일람표")) { d.Kind = NeisFileKind.ScoreTable; d.Table = ReadScoreTable(grid, path); d.Header = d.Table.Header; }
                    else d.Kind = NeisFileKind.Unknown;
                }
            }
            catch (Exception ex)
            {
                d.Kind = NeisFileKind.Unknown;
                d.Error = ex.Message.Contains("Package") || ex.Message.Contains("encrypted") || ex is InvalidDataException
                    ? "엑셀 파일로 열 수 없습니다(손상되었거나 암호가 걸린 파일). 나이스에서 엑셀(XLS data)로 다시 저장해 보세요."
                    : ex.Message;
            }
            return d;
        }

        static ExcelPackage OpenReadOnly(string path)
        {
            // OneDrive 동기화 중에도 열 수 있도록 공유 읽기로 복사해 연다
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var ms = new MemoryStream();
                fs.CopyTo(ms);
                ms.Position = 0;
                return new ExcelPackage(ms);
            }
        }

        // ---------- 공통: 머리말 해석 ----------
        static ExamHeader ParseHeader(Grid g, int maxRow)
        {
            var h = new ExamHeader();
            for (int r = 1; r <= maxRow; r++)
            {
                for (int c = 1; c <= g.MaxCol; c++)
                {
                    var t = g.Text(r, c);
                    if (t == null) continue;
                    if (t.Contains("학년도"))
                    {
                        var tokens = Regex.Split(t.Trim(), @"\s{2,}").Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
                        foreach (var tok in tokens)
                        {
                            Match m;
                            if ((m = Regex.Match(tok, @"(\d{4})\s*학년도")).Success) h.Year = int.Parse(m.Groups[1].Value);
                            else if ((m = Regex.Match(tok, @"(\d+)\s*학기")).Success) h.Semester = int.Parse(m.Groups[1].Value);
                            else if (h.ExamName == null && (tok.Contains("고사") || tok.Contains("시험") || Regex.IsMatch(tok, @"^\d+\s*차"))) h.ExamName = tok;   // "3학년 기말고사", "1차 시험", "2차 지필평가" 등
                            else if ((m = Regex.Match(tok, @"(\d+)\s*학년")).Success) h.Grade = int.Parse(m.Groups[1].Value);
                            else if ((m = Regex.Match(tok, @"^(\d+)\s*(반|강의실)$")).Success) { h.ClassNo = int.Parse(m.Groups[1].Value); h.LectureWord = m.Groups[2].Value == "강의실"; }
                            else if (tok == "주간" || tok == "야간") { }
                            else if (h.Subject == null && !tok.Contains("강의실")) h.Subject = tok;
                        }
                    }
                    if (t.Contains("교과목") && t.Contains(":"))
                    {
                        var m = Regex.Match(t, @"교과목\s*:\s*([^\s]+)");
                        if (m.Success) h.Subject = m.Groups[1].Value;
                    }
                    if (t.Contains("고사") && t.Contains(":") && h.ExamName == null)
                    {
                        var m = Regex.Match(t, @"고사\s*:\s*(.+?)(\s{2,}|$)");
                        if (m.Success) h.ExamName = m.Groups[1].Value.Trim();
                    }
                }
            }
            return h;
        }

        static string Compact(string t) => Regex.Replace(t ?? "", @"\s+", "");

        // ---------- 문항정보표 ----------
        /// <summary>출제 교사가 작성한 문항정보표: 문항별 내용영역·성취기준·예상 난이도(어려움/보통/쉬움 ○표)·배점·정답</summary>
        static ItemInfoFile ReadItemInfo(Grid g, string path)
        {
            var f = new ItemInfoFile { Path = path, Header = new ExamHeader() };
            for (int r = 1; r <= Math.Min(12, g.MaxRow); r++)
                for (int c = 1; c <= g.MaxCol; c++)
                {
                    var t = g.Text(r, c); if (t == null) continue;
                    Match m;
                    if ((m = Regex.Match(t, @"(\d{4})\s*년도\s*(\d)\s*학기\s*(.+)")).Success)
                    { f.Header.Year = int.Parse(m.Groups[1].Value); f.Header.Semester = int.Parse(m.Groups[2].Value); f.Header.ExamName = m.Groups[3].Value.Trim(); }
                    else if ((m = Regex.Match(t, @"^\s*(\d+)\s*학년\s*$")).Success) f.Header.Grade = int.Parse(m.Groups[1].Value);
                    else if ((m = Regex.Match(t, @"\(\s*(.+?)\s*\)\s*과목")).Success) f.Header.Subject = m.Groups[1].Value.Trim();
                    else if ((m = Regex.Match(t, @"총\s*문항\s*수\s*(\d+)")).Success) f.TotalItems = int.Parse(m.Groups[1].Value);
                }
            if (f.Header.ExamName == null) throw new Exception("문항정보표 머리말에서 '○○년도 ○학기 ○○고사'를 찾지 못했습니다.");

            int hdrRow = g.FindRow(t => Compact(t) == "문항번호");
            if (hdrRow < 0) throw new Exception("문항정보표에서 '문항번호' 머리글을 찾지 못했습니다.");
            while (hdrRow > 0)
            {
                int numCol = g.FindColInRow(hdrRow, t => Compact(t) == "문항번호");
                int areaCol = g.FindColInRow(hdrRow, t => Compact(t).Contains("내용영역"));
                int stdCol = g.FindColInRow(hdrRow, t => Compact(t).Contains("성취기준"));
                int ptsCol = g.FindColInRow(hdrRow, t => Compact(t) == "배점");
                int ansCol = g.FindColInRow(hdrRow, t => Compact(t) == "정답");
                int diffRow = hdrRow + 1;
                int hardCol = g.FindColInRow(diffRow, t => Compact(t) == "어려움");
                int midCol = g.FindColInRow(diffRow, t => Compact(t) == "보통");
                int easyCol = g.FindColInRow(diffRow, t => Compact(t) == "쉬움");
                string section = "선택형";
                for (int r = hdrRow - 1; r >= Math.Max(1, hdrRow - 2); r--)
                {
                    var t = g.Text(r, numCol);
                    if (t != null && t.Contains("서답형")) { section = "서답형"; break; }
                    if (t != null && t.Contains("선택형")) break;
                }
                int r2 = diffRow + 1;
                for (; r2 <= g.MaxRow; r2++)
                {
                    var label = g.Text(r2, numCol);
                    if (label != null && (Compact(label).StartsWith("총합계") || Compact(label).StartsWith("합계"))) break;
                    var n = g.Num(r2, numCol);
                    if (!n.HasValue) continue;   // 쪽 나눔 줄(학교명, 쪽 번호)은 건너뜀
                    var meta = new ItemMeta { Number = (int)n.Value, Section = section };
                    if (areaCol > 0) meta.Area = (g.Text(r2, areaCol) ?? "").Trim();
                    if (stdCol > 0) meta.Standard = Regex.Replace(g.Text(r2, stdCol) ?? "", @"\s+", " ").Trim();
                    if (ptsCol > 0) meta.Points = g.Num(r2, ptsCol) ?? 0;
                    if (ansCol > 0) meta.CorrectRaw = (g.Text(r2, ansCol) ?? "").Trim();
                    bool hard = hardCol > 0 && !string.IsNullOrWhiteSpace(g.Text(r2, hardCol));
                    bool mid = midCol > 0 && !string.IsNullOrWhiteSpace(g.Text(r2, midCol));
                    bool easy = easyCol > 0 && !string.IsNullOrWhiteSpace(g.Text(r2, easyCol));
                    meta.Expected = hard ? "어려움" : mid ? "보통" : easy ? "쉬움" : null;
                    f.Items.Add(meta);
                }
                hdrRow = g.FindRow(t => Compact(t) == "문항번호", r2 + 1);
            }
            if (f.Items.Count == 0) throw new Exception("문항정보표에서 문항 줄을 찾지 못했습니다.");
            return f;
        }

        // ---------- 정오표 ----------
        /// <summary>학생 줄에서 "반/번호"가 든 열과 학번 열을 정한다. 머리글은 "반/번호" 하나만 있고(두 칸 병합), 실제 값의 순서는 학교마다 다르다.</summary>
        static void LocateKeyCols(Grid g, int firstDataRow, int hdrCol, out int classCol, out int idCol)
        {
            classCol = hdrCol; idCol = hdrCol + 1;
            for (int r = firstDataRow; r <= Math.Min(g.MaxRow, firstDataRow + 3); r++)
            {
                var a = g.Text(r, hdrCol); var b = g.Text(r, hdrCol + 1);
                if (a != null && RxClassNum.IsMatch(a)) return;
                if (b != null && RxClassNum.IsMatch(b)) { classCol = hdrCol + 1; idCol = hdrCol; return; }
            }
        }

        static AnswerSheetFile ReadAnswerSheet(Grid g, string path)
        {
            var f = new AnswerSheetFile { Path = path, Header = ParseHeader(g, 6) };
            if (!f.Header.ClassNo.HasValue) throw new Exception("정오표 머리말에서 반을 찾지 못했습니다.");
            f.ClassNo = f.Header.ClassNo.Value;

            int hdrRow = g.FindRow(t => Compact(t) == "반/번호");
            if (hdrRow < 0) throw new Exception("정오표에서 반/번호 머리글을 찾지 못했습니다.");
            int hdrCol = g.FindColInRow(hdrRow, t => Compact(t) == "반/번호");
            int idCol = hdrCol + 1, nameCol = hdrCol + 2;
            int ansRow = g.FindRow(t => t.Trim() == "정답", hdrRow + 1, hdrRow + 3);
            int ptsRow = g.FindRow(t => t.Trim() == "배점", hdrRow + 1, hdrRow + 4);
            if (ansRow < 0 || ptsRow < 0) throw new Exception("정오표에서 정답·배점 행을 찾지 못했습니다.");

            int choiceCol = g.FindColInRow(hdrRow, t => Compact(t).StartsWith("선택형"));
            int essayCol = g.FindColInRow(hdrRow, t => Compact(t).StartsWith("서답형"));
            int etcCol = g.FindColInRow(hdrRow, t => Compact(t).StartsWith("기타"));
            // 총점 열 머리글은 학교·과목 설정에 따라 "영역총점" 또는 "과목총점"으로 나온다. "총점"으로 끝나는 열을 찾는다.
            int totalCol = g.FindColInRow(hdrRow, t => Compact(t).EndsWith("총점") && !Compact(t).StartsWith("선택형") && !Compact(t).StartsWith("서답형"));
            if (choiceCol < 0) throw new Exception("정오표에서 '선택형점수' 열을 찾지 못했습니다. 나이스에서 '교과목별 학생답 정오표'를 엑셀(XLS data)로 저장한 파일인지 확인하세요.");
            if (totalCol < 0) throw new Exception("정오표에서 총점 열(영역총점 또는 과목총점)을 찾지 못했습니다. 나이스에서 '교과목별 학생답 정오표'를 엑셀(XLS data)로 저장한 파일인지 확인하세요.");

            var itemCols = new List<int>();
            for (int c = nameCol + 1; c < choiceCol; c++)
            {
                var v = g.Num(hdrRow, c);
                if (v.HasValue)
                {
                    itemCols.Add(c);
                    var correctRaw = g.Text(ansRow, c);
                    var pts = g.Num(ptsRow, c) ?? 0;
                    int? correct = null;
                    int ci;
                    if (correctRaw != null && int.TryParse(correctRaw.Trim(), out ci)) correct = ci;
                    f.Items.Add(new ItemInfo { Number = (int)v.Value, CorrectRaw = correctRaw, Correct = correct, Points = pts });
                }
            }
            f.ChoiceMax = g.Num(ptsRow, choiceCol) ?? 0;
            f.EssayMax = essayCol > 0 ? (g.Num(ptsRow, essayCol) ?? 0) : 0;
            f.EtcMax = etcCol > 0 ? (g.Num(ptsRow, etcCol) ?? 0) : 0;
            f.TotalMax = g.Num(ptsRow, totalCol) ?? (f.ChoiceMax + f.EssayMax + f.EtcMax);

            double choicePts = f.Items.Where(i => i.IsChoice).Sum(i => i.Points);
            if (Math.Abs(choicePts - f.ChoiceMax) > 0.001)
                f.Warnings.Add($"{f.ClassNo}반 정오표: 선택형 문항 배점 합({choicePts})이 선택형 만점({f.ChoiceMax})과 다릅니다.");

            // 학교에 따라 "반/번호, 학번" 순서이거나 "학번, 반/번호" 순서로 나오므로 첫 학생 줄을 보고 정한다
            int classCol; LocateKeyCols(g, ptsRow + 1, hdrCol, out classCol, out idCol);
            for (int r = ptsRow + 1; r <= g.MaxRow; r++)
            {
                var key = g.Text(r, classCol);
                if (key == null) break;
                var m = RxClassNum.Match(key);
                if (!m.Success) break;
                var row = new AnswerRow
                {
                    ClassNo = int.Parse(m.Groups[1].Value),
                    HomeClassNo = int.Parse(m.Groups[1].Value),
                    Number = int.Parse(m.Groups[2].Value),
                    StudentId = g.Text(r, idCol),
                    Name = (g.Text(r, nameCol) ?? "").Trim(),
                    Answers = itemCols.Select(c => (g.Text(r, c) ?? "").Trim()).ToArray(),
                    ChoiceScore = g.Num(r, choiceCol) ?? 0,
                    EssayScore = essayCol > 0 ? (g.Num(r, essayCol) ?? 0) : 0,
                    EtcScore = etcCol > 0 ? (g.Num(r, etcCol) ?? 0) : 0,
                    Total = g.Num(r, totalCol) ?? 0,
                };
                if (row.HomeClassNo != f.ClassNo) f.MixedCount++;   // 고교학점제: 강의실에 여러 반 학생이 섞여 있음 (경고 아님)
                // 문항별 정오로 선택형 점수를 다시 계산해 파일 값과 비교
                double recomputed = 0;
                for (int i = 0; i < f.Items.Count; i++)
                    if (f.Items[i].IsChoice && row.Answers[i] == ".") recomputed += f.Items[i].Points;
                if (Math.Abs(recomputed - row.ChoiceScore) > 0.001)
                    f.Warnings.Add($"{row.ClassNo}반 {row.Number}번 {row.Name}: 문항 정오로 계산한 선택형 점수({recomputed})가 파일의 선택형 점수({row.ChoiceScore})와 다릅니다.");
                f.Rows.Add(row);
            }
            return f;
        }

        // ---------- 수행평가 일람표 ----------
        static PerfFile ReadPerf(Grid g, string path)
        {
            var f = new PerfFile { Path = path, Header = ParseHeader(g, 6) };
            if (!f.Header.ClassNo.HasValue) throw new Exception("수행평가 일람표 머리말에서 강의실(반)을 찾지 못했습니다.");
            f.ClassNo = f.Header.ClassNo.Value;

            int hdrRow = g.FindRow(t => Compact(t) == "반/번호");
            if (hdrRow < 0) throw new Exception("수행평가 일람표에서 반/번호 머리글을 찾지 못했습니다.");
            int hdrCol = g.FindColInRow(hdrRow, t => Compact(t) == "반/번호");
            int nameCol = g.FindColInRow(hdrRow, t => t.Trim() == "성명");
            int sumCol = g.FindColInRow(hdrRow, t => t.Trim() == "합계");
            int noteCol = g.FindColInRow(hdrRow, t => t.Trim() == "비고");
            if (nameCol < 0 || sumCol < 0) throw new Exception("수행평가 일람표에서 성명·합계 열을 찾지 못했습니다.");
            int idCol = hdrCol + 1;

            var areaCols = new List<int>();
            var rx = new Regex(@"^(.*?)\s*\(\s*만점\s*([\d.]+)\s*,\s*([\d.]+)\s*%\s*\)", RegexOptions.Singleline);
            for (int c = nameCol + 1; c < sumCol; c++)
            {
                var t = g.Text(hdrRow, c);
                if (string.IsNullOrWhiteSpace(t)) continue;
                var norm = Regex.Replace(t, @"\s+", " ").Trim();
                var m = rx.Match(norm);
                if (m.Success)
                {
                    f.Areas.Add(new PerfArea { Name = m.Groups[1].Value.Trim(), Max = double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture), Ratio = double.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture) });
                }
                else
                {
                    f.Areas.Add(new PerfArea { Name = norm, Max = 0, Ratio = 0 });
                    f.Warnings.Add($"{f.ClassNo}반 수행평가: 영역 [{norm}]의 만점·비율을 제목에서 읽지 못했습니다. 설정에서 직접 입력하세요.");
                }
                areaCols.Add(c);
            }

            int classCol = hdrCol;
            for (int r = hdrRow + 1; r <= g.MaxRow; r++)
            {
                if (r == hdrRow + 1 || f.Rows.Count == 0) LocateKeyCols(g, r, hdrCol, out classCol, out idCol);   // 머리글 아래 빈 줄이 있을 수 있어 첫 학생 줄에서 정한다
                var key = g.Text(r, classCol);
                if (key == null) { if (f.Rows.Count == 0) continue; break; }
                var m = RxClassNum.Match(key);
                if (!m.Success) { if (f.Rows.Count == 0) continue; break; }
                var row = new PerfRow
                {
                    ClassNo = int.Parse(m.Groups[1].Value),
                    HomeClassNo = int.Parse(m.Groups[1].Value),
                    Number = int.Parse(m.Groups[2].Value),
                    StudentId = g.Text(r, idCol),
                    Name = (g.Text(r, nameCol) ?? "").Trim(),
                    Scores = new double?[areaCols.Count],
                    Exempt = new bool[areaCols.Count],
                    Total = g.Num(r, sumCol),
                    Note = noteCol > 0 ? g.Text(r, noteCol) : null,
                };
                for (int i = 0; i < areaCols.Count; i++)
                {
                    var n = g.Num(r, areaCols[i]);
                    var t = g.Text(r, areaCols[i]);
                    if (n.HasValue) row.Scores[i] = n;
                    else if (t != null && t.Contains("면제")) row.Exempt[i] = true;
                }
                if (row.HomeClassNo != f.ClassNo) f.MixedCount++;   // 고교학점제: 강의실에 여러 반 학생이 섞임
                f.Rows.Add(row);
            }
            return f;
        }

        // ---------- 정기시험 교과목별 일람표 ----------
        static ScoreTableFile ReadScoreTable(Grid g, string path)
        {
            var f = new ScoreTableFile { Path = path, Header = ParseHeader(g, 7) };
            for (int r = 1; r <= 7; r++)
                for (int c = 1; c <= g.MaxCol; c++)
                {
                    var t = g.Text(r, c);
                    if (t == null) continue;
                    var m = Regex.Match(t, @"과목만점\s*:\s*([\d.]+)");
                    if (m.Success) f.SubjectMax = double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                }

            int hdrRow = g.FindRow(t => Compact(t) == "반번호");
            if (hdrRow < 0) throw new Exception("일람표에서 반/번호 머리글을 찾지 못했습니다.");
            int hdrCol = g.FindColInRow(hdrRow, t => Compact(t) == "반번호");
            var classCols = new Dictionary<int, int>();
            for (int c = hdrCol + 1; c <= g.MaxCol; c++)
            {
                var v = g.Num(hdrRow, c);
                if (v.HasValue)
                {
                    int cls = (int)v.Value;
                    classCols[cls] = c;
                    f.Cells[cls] = new Dictionary<int, ScoreCell>();
                    f.Stats[cls] = new ClassStat { ClassNo = cls };
                }
            }
            for (int r = hdrRow + 1; r <= g.MaxRow; r++)
            {
                var label = g.Text(r, hdrCol);
                var num = g.Num(r, hdrCol);
                if (num.HasValue)
                {
                    foreach (var kv in classCols)
                    {
                        var n = g.Num(r, kv.Value);
                        var t = g.Text(r, kv.Value);
                        if (n.HasValue) f.Cells[kv.Key][(int)num.Value] = new ScoreCell { Score = n };
                        else if (t != null && t.Contains("면제")) f.Cells[kv.Key][(int)num.Value] = new ScoreCell { Exempt = true };
                    }
                    continue;
                }
                if (label == null) continue;
                var norm = Compact(label);
                if (norm == "응시생수")
                {
                    foreach (var kv in classCols)
                    {
                        var t = g.Text(r, kv.Value);
                        var m = t == null ? null : Regex.Match(t, @"(\d+)");
                        if (m != null && m.Success) f.Stats[kv.Key].Count = int.Parse(m.Groups[1].Value);
                    }
                }
                else if (norm == "총점")
                {
                    foreach (var kv in classCols) f.Stats[kv.Key].Sum = g.Num(r, kv.Value);
                }
                else if (norm == "평균")
                {
                    foreach (var kv in classCols) f.Stats[kv.Key].Mean = g.Num(r, kv.Value);
                }
            }
            return f;
        }

        /// <summary>엑셀 시트를 (행,열)로 읽기 쉽게 감싼 도우미</summary>
        class Grid
        {
            readonly ExcelWorksheet _ws;
            public int MaxRow, MaxCol;
            public Grid(ExcelWorksheet ws) { _ws = ws; MaxRow = ws.Dimension.End.Row; MaxCol = ws.Dimension.End.Column; }
            public object Raw(int r, int c) => _ws.Cells[r, c].Value;
            public string Text(int r, int c)
            {
                var v = Raw(r, c);
                if (v == null) return null;
                var s = v is double d ? d.ToString(CultureInfo.InvariantCulture) : v.ToString();
                return s.Length == 0 ? null : s;
            }
            public double? Num(int r, int c)
            {
                var v = Raw(r, c);
                if (v == null) return null;
                if (v is double d) return d;
                if (v is int i) return i;
                if (v is decimal m) return (double)m;
                double p;
                if (v is string s && double.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out p)) return p;
                return null;
            }
            public int FindRow(Func<string, bool> pred, int from = 1, int to = int.MaxValue)
            {
                for (int r = from; r <= Math.Min(to, MaxRow); r++)
                    for (int c = 1; c <= MaxCol; c++)
                    { var t = Text(r, c); if (t != null && pred(t)) return r; }
                return -1;
            }
            public int FindColInRow(int r, Func<string, bool> pred)
            {
                for (int c = 1; c <= MaxCol; c++)
                { var t = Text(r, c); if (t != null && pred(t)) return c; }
                return -1;
            }
            public string FindTextInRows(int from, int to, Func<string, bool> pred)
            {
                for (int r = from; r <= Math.Min(to, MaxRow); r++)
                    for (int c = 1; c <= MaxCol; c++)
                    { var t = Text(r, c); if (t != null && pred(t)) return t; }
                return null;
            }
        }
    }
}
