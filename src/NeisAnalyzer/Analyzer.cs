using System;
using System.Collections.Generic;
using System.Linq;

namespace NeisAnalyzer
{
    // ===================== 결과 자료 구조 =====================

    public class Student
    {
        public int ClassNo, Number;      // ClassNo = 분석 단위(중학교: 반, 고교학점제: 강의실)
        public int HomeClassNo;          // 학생의 행정반 (중학교에서는 ClassNo와 같음)
        public bool Lecture;             // 강의실 모드로 만들어진 학생인지
        public string StudentId, Name;

        public PerfRow Perf;                       // 수행 자료 (없으면 null)
        public double? PerfTotal;                  // 수행 합계
        public bool PerfAllExempt, PerfPartialExempt, PerfMissingArea;

        public Dictionary<string, AnswerRow> Written = new Dictionary<string, AnswerRow>();   // 고사명 → 정오표 행
        public Dictionary<string, double> WrittenConverted = new Dictionary<string, double>(); // 고사명 → 환산 점수
        public HashSet<string> WrittenExemptByTable = new HashSet<string>();                  // 일람표에 면제로 표시된 고사

        public double? Final;            // 최종 점수
        public double? AchievementBasisValue;
        public string Achievement;       // A~E
        public int? GradeRank, ClassRank;
        public bool Boundary;            // 성취도 경계
        public string BoundaryNote;      // 예: "A 경계(-0.4)"
        public bool Included;            // 등수·통계 포함 여부
        public string Status = "정상";   // 정상 / 면제 / 지필 결시 / ...
        public List<string> Flags = new List<string>();   // 관심 학생 사유
        public int BlankCount;           // 무표기 수(모든 고사 합)
        public double? WrittenRawSum;    // 지필 원점수 합(고사가 둘이면 평균 아님, 합)
        public double? PerfPercent, WrittenPercent, PerfWrittenGap, PerfWrittenGapAdj;

        public string Key => HomeClassNo + "/" + Number;   // 행정반/번호 = 학년 안에서 유일한 학생 식별자
        public string Label => Lecture ? $"{HomeClassNo}반 {Number}번 {Name}({ClassNo}강의실)" : $"{ClassNo}반 {Number}번 {Name}";
    }

    public class ItemStat
    {
        public int Number;
        public double Points;
        public int? Correct;
        public int N, CorrectCount, BlankCount, MultiCount;
        public Dictionary<int, int> OptionCounts = new Dictionary<int, int>();
        public double CorrectRate => N == 0 ? 0 : (double)CorrectCount / N;
        public double WrongRate => N == 0 ? 0 : 1 - CorrectRate - BlankRate - MultiRate;
        public double BlankRate => N == 0 ? 0 : (double)BlankCount / N;
        public double MultiRate => N == 0 ? 0 : (double)MultiCount / N;
        public string Difficulty;
        public double? UpperRate, LowerRate, Discrimination, PointBiserial;
        public int? AttractiveWrong;     // 가장 많이 고른 오답 번호
        public double AttractiveWrongRate;
        public Dictionary<int, double> ClassCorrectRate = new Dictionary<int, double>();
        public Dictionary<int, int> ClassN = new Dictionary<int, int>();
        // 성취도(A~E)별 정답률: 낮은 성취도가 높은 성취도보다 많이 맞힌 "수준 역전" 문항을 찾는 데 쓴다
        public Dictionary<string, int> LevelN = new Dictionary<string, int>();
        public Dictionary<string, double> LevelCorrectRate = new Dictionary<string, double>();
        public string LevelInversion;    // 예: "수준 역전: D가 B보다 20%p 높음" (없으면 null)
        // 문항정보표(있을 때)
        public string Area, Standard, Expected;   // 내용영역, 성취기준, 교사 예상 난이도
        public int? ExpectedGap => Expected == null || Difficulty == null ? (int?)null : DifficultyLevel(Difficulty) - DifficultyLevel(Expected);   // +면 예상보다 어려웠음
        public static int DifficultyLevel(string d) => d == "어려움" ? 2 : d == "보통" ? 1 : 0;
    }

    /// <summary>성취도(A~E)별 인원·평균·표준편차 (KICE 성취도분석표와 같은 구성)</summary>
    public class LevelSummary
    {
        public string Grade;
        public int Count;
        public double Mean, Sd, Min, Max;
        public double? PerfMean;
        public Dictionary<string, double> WrittenMeans = new Dictionary<string, double>();
    }

    public class ExamAnalysis
    {
        public string ExamName;
        public List<ItemInfo> Items;
        public double TotalMax, ChoiceMax, EssayMax;
        public int OptionCount = 5;
        public List<ItemStat> ItemStats = new List<ItemStat>();
        public double? KR20;
        public int N;                  // 응시 인원
        public double Mean, Sd;        // 지필 원점수(영역 총점)
        public Dictionary<int, ClassStat> ClassStats = new Dictionary<int, ClassStat>();   // 정오표로 계산한 반별 통계
        public ScoreTableFile Table;   // 검산용 일람표(있으면)
        public ItemInfoFile Info;      // 문항정보표(있으면): 예상 난이도·내용영역·성취기준
        public List<VerifyRow> Verify = new List<VerifyRow>();
        public List<string> VerifyStudentIssues = new List<string>();
        public bool? VerifyAllMatch;
    }

    public class VerifyRow
    {
        public int ClassNo;
        public int CalcCount; public double CalcSum, CalcMean;
        public int? TableCount; public double? TableSum, TableMean;
        public bool Match;
        public string Result;   // 일치 / 불일치 / 정오표 없음(비교 불가)
    }

    public class ClassSummary
    {
        public int ClassNo;
        public int Count, Enrolled;
        public double Mean, Median, Sd, Max, Min;
        public double[] PerfAreaMeans;
        public double PerfMean;
        public Dictionary<string, double> WrittenMeans = new Dictionary<string, double>();
        public Dictionary<string, int> AchievementCounts = new Dictionary<string, int>();
        public int MeanRank;
    }

    public class AnalysisResult
    {
        public AnalysisSettings Settings;
        public ExamHeader Header;              // 대표 머리말(학년도·학기·학년·과목)
        public List<string> Exams = new List<string>();
        public List<int> Classes = new List<int>();
        public List<Student> Students = new List<Student>();
        public List<Student> Ranked = new List<Student>();
        public List<ExamAnalysis> ExamAnalyses = new List<ExamAnalysis>();
        public List<ClassSummary> ClassSummaries = new List<ClassSummary>();
        public List<LevelSummary> LevelSummaries = new List<LevelSummary>();
        public bool LectureMode;                       // 고교학점제: 파일 단위가 행정반이 아니라 강의실(수업 집단)
        public string Unit => LectureMode ? "강의실" : "반";
        public List<string> Warnings = new List<string>();
        public List<string> Notes = new List<string>();     // 분석 기준 안내에 넣을 설명
        public List<string> Highlights = new List<string>(); // "이번 시험의 특징"
        public double PerfMax => Settings.PerfAreas.Sum(a => a.Max);
        public bool HasPerf => Settings.PerfAreas.Count > 0;
        public int Count => Ranked.Count;
        public double Mean, Median, Sd, Max, Min;
        public Dictionary<string, int> AchievementCounts = new Dictionary<string, int>();
        public int[] Histogram = new int[10];   // 0~9, 10~19, ..., 90~100
        public Dictionary<int, int[]> ClassHistogram = new Dictionary<int, int[]>();
        public double? PerfWrittenCorr;
        public double AvgPerfWrittenGap;
        public List<string> FileList = new List<string>();
    }

    // ===================== 계산 엔진 =====================

    public static class Analyzer
    {
        public static readonly string[] Grades = { "A", "B", "C", "D", "E" };
        public const double GapCut = 20;   // 수행-지필 상대 차이가 이만큼 이상이면 관심 학생

        /// <summary>읽은 파일에서 기본 설정값을 만든다 (수행 영역, 지필 비율).</summary>
        public static AnalysisSettings DefaultSettings(IList<DetectedFile> files)
        {
            var s = new AnalysisSettings();
            var perf = files.FirstOrDefault(f => f.Kind == NeisFileKind.PerfSheet);
            if (perf != null)
                s.PerfAreas = perf.Perf.Areas.Select(a => new PerfArea { Name = a.Name, Max = a.Max, Ratio = a.Ratio }).ToList();
            var exams = files.Where(f => f.Kind == NeisFileKind.AnswerSheet).Select(f => f.Header.ExamName ?? "지필고사").Distinct().OrderBy(e => s.ExamOrderKey(e)).ToList();
            double remain = Math.Max(0, 100 - s.PerfRatioSum);
            foreach (var e in exams) s.ExamRatios[e] = exams.Count == 0 ? 0 : Math.Round(remain / exams.Count, 2);
            return s;
        }

        /// <summary>고교학점제 파일인지 판단한다: 정오표 머리말이 "강의실"이거나, 파일의 반 번호와 다른 반 학생이 섞여 있으면 강의실 모드.</summary>
        public static bool IsLectureMode(IEnumerable<DetectedFile> files) =>
            files.Any(f => (f.Kind == NeisFileKind.AnswerSheet && (f.Header.LectureWord || f.Answer.MixedCount > 0)) || (f.Kind == NeisFileKind.PerfSheet && f.Perf.MixedCount > 0));

        /// <summary>파일 묶음을 검사해 빠진 반·중복·섞임을 경고로 만든다.</summary>
        public static List<string> CheckFiles(IList<DetectedFile> files)
        {
            var w = new List<string>();
            var ok = files.Where(f => f.Kind != NeisFileKind.Unknown).ToList();
            string u = IsLectureMode(ok) ? "강의실" : "반";
            if (u == "강의실") w.Add("고교학점제 강의실 파일로 판단했습니다. 반 대신 강의실(수업 집단) 단위로 비교하고, 학생의 반·번호는 그대로 표시합니다. 일람표가 강의실이 아닌 행정반 기준이면 검산이 맞지 않을 수 있습니다.");
            foreach (var f in files.Where(f => f.Kind == NeisFileKind.Unknown))
                w.Add($"[{System.IO.Path.GetFileName(f.Path)}] {(f.Error != null ? "읽는 중 문제: " + f.Error : "나이스 정오표·수행평가·일람표가 아니어서 건너뜁니다.")}");
            if (ok.Count == 0) { w.Add("분석할 수 있는 나이스 파일이 없습니다."); return w; }

            foreach (var g in new[] { "학년도", "학기", "학년", "과목" })
            {
                var vals = ok.Select(f => g == "학년도" ? f.Header.Year.ToString() : g == "학기" ? f.Header.Semester.ToString() : g == "학년" ? f.Header.Grade.ToString() : (f.Header.Subject ?? "")).Distinct().ToList();
                if (vals.Count > 1) w.Add($"{g}이(가) 서로 다른 파일이 섞여 있습니다: {string.Join(", ", vals)}");
            }
            var answers = ok.Where(f => f.Kind == NeisFileKind.AnswerSheet).ToList();
            var perfs = ok.Where(f => f.Kind == NeisFileKind.PerfSheet).ToList();
            var tables = ok.Where(f => f.Kind == NeisFileKind.ScoreTable).ToList();
            var infos = ok.Where(f => f.Kind == NeisFileKind.ItemInfoSheet).ToList();
            foreach (var dup in answers.GroupBy(f => (f.Header.ExamName, f.Header.ClassNo)).Where(g => g.Count() > 1))
                w.Add($"{dup.Key.ClassNo}{u} {dup.Key.ExamName} 정오표가 {dup.Count()}개 들어 있습니다. 첫 번째 파일만 사용합니다.");
            foreach (var dup in perfs.GroupBy(f => f.Header.ClassNo).Where(g => g.Count() > 1))
                w.Add($"{dup.Key}{u} 수행평가 일람표가 {dup.Count()}개 들어 있습니다. 첫 번째 파일만 사용합니다.");
            foreach (var dup in tables.GroupBy(f => f.Header.ExamName).Where(g => g.Count() > 1))
                w.Add($"{dup.Key} 일람표가 {dup.Count()}개 들어 있습니다. 첫 번째 파일만 사용합니다.");

            var classes = answers.Select(f => f.Header.ClassNo.Value).Concat(perfs.Select(f => f.Header.ClassNo.Value)).Distinct().OrderBy(c => c).ToList();
            foreach (var t in tables) classes = classes.Union(t.Table.Cells.Keys).OrderBy(c => c).ToList();
            var exams = answers.Select(f => f.Header.ExamName).Distinct().ToList();
            foreach (var e in exams)
            {
                var missing = classes.Where(c => !answers.Any(f => f.Header.ExamName == e && f.Header.ClassNo == c)).ToList();
                if (missing.Count > 0) w.Add($"{e} 정오표가 없는 {u}: {string.Join(", ", missing.Select(c => c + u))}");
            }
            if (perfs.Count > 0)
            {
                var missing = classes.Where(c => !perfs.Any(f => f.Header.ClassNo == c)).ToList();
                if (missing.Count > 0) w.Add($"수행평가 일람표가 없는 {u}: {string.Join(", ", missing.Select(c => c + u))}");
            }
            else w.Add("수행평가 일람표가 없습니다. 지필 점수만으로 분석합니다(최종 점수 = 지필 원점수).");
            if (answers.Count == 0) w.Add("정오표가 없습니다. 지필 점수·문항 분석을 할 수 없습니다.");
            if (perfs.Count > 1)
            {
                var first = perfs[0].Perf.Areas.Select(a => a.Name).ToList();
                foreach (var p in perfs.Skip(1))
                    if (!p.Perf.Areas.Select(a => a.Name).SequenceEqual(first))
                        w.Add($"{p.Header.ClassNo}{u} 수행평가 영역 구성이 {perfs[0].Header.ClassNo}{u}과 다릅니다.");
            }
            foreach (var t in tables)
                if (!exams.Contains(t.Header.ExamName))
                    w.Add($"일람표의 고사명 [{t.Header.ExamName}]과 일치하는 정오표가 없어 검산에 쓰지 못합니다.");
            foreach (var i in infos)
                if (answers.Count > 0 && MatchExam(i.Header.ExamName, exams) == null)
                    w.Add($"문항정보표의 고사 [{i.Header.ExamName}]에 맞는 정오표를 찾지 못해 예상 난이도 비교를 하지 못합니다.");
            foreach (var dup in infos.GroupBy(i => MatchExam(i.Header.ExamName, exams) ?? i.Header.ExamName).Where(g => g.Count() > 1))
                w.Add($"[{dup.Key}] 문항정보표가 {dup.Count()}개 들어 있습니다. 첫 번째 파일만 사용합니다.");
            foreach (var f in ok)
                foreach (var fw in (f.Answer?.Warnings ?? f.Perf?.Warnings ?? f.Table?.Warnings ?? new List<string>()))
                    w.Add(fw);
            return w;
        }

        /// <summary>문항정보표의 고사명("중간고사")을 정오표의 고사명("3학년 중간고사", "1차 시험")과 맞춘다. 고사가 하나뿐이면 그것으로 본다.</summary>
        public static string MatchExam(string infoExam, IEnumerable<string> exams)
        {
            var list = exams.Where(e => e != null).ToList();
            bool hasKey = false;
            if (infoExam != null)
                foreach (var key in new[] { "중간", "기말", "1차", "2차", "3차", "4차", "학기말" })
                    if (infoExam.Contains(key)) { hasKey = true; var hit = list.FirstOrDefault(e => e.Contains(key)); if (hit != null) return hit; }
            if (hasKey) return null;   // "중간고사" 정보표를 기말고사 정오표에 붙이지 않는다
            return list.Count == 1 ? list[0] : null;
        }

        public static AnalysisResult Analyze(IList<DetectedFile> files, AnalysisSettings s)
        {
            var r = new AnalysisResult { Settings = s };
            r.Warnings.AddRange(CheckFiles(files));
            var ok = files.Where(f => f.Kind != NeisFileKind.Unknown).ToList();
            r.FileList = files.Select(f => $"{System.IO.Path.GetFileName(f.Path)} → {f.Label}").ToList();
            if (ok.Count == 0) return r;
            r.Header = ok.First().Header;

            var answers = ok.Where(f => f.Kind == NeisFileKind.AnswerSheet).GroupBy(f => (f.Header.ExamName ?? "지필고사", f.Header.ClassNo.Value)).ToDictionary(g => g.Key, g => g.First().Answer);
            var perfs = ok.Where(f => f.Kind == NeisFileKind.PerfSheet).GroupBy(f => f.Header.ClassNo.Value).ToDictionary(g => g.Key, g => g.First().Perf);
            var tables = ok.Where(f => f.Kind == NeisFileKind.ScoreTable).GroupBy(f => f.Header.ExamName ?? "").ToDictionary(g => g.Key, g => g.First().Table);

            r.Exams = answers.Keys.Select(k => k.Item1).Distinct().OrderBy(e => s.ExamOrderKey(e)).ToList();
            r.Classes = answers.Keys.Select(k => k.Item2).Concat(perfs.Keys).Distinct().OrderBy(c => c).ToList();
            bool hasPerfFiles = perfs.Count > 0;

            // ---------- 고교학점제 강의실 모드 ----------
            // 중학교: 학생의 반 = 파일의 반. 고교학점제: 파일은 강의실(수업 집단)이고 학생의 반은 행정반.
            // 강의실 모드에서는 모든 "반별" 비교를 강의실 단위로 하고, 학생은 행정반/번호로 식별한다.
            r.LectureMode = IsLectureMode(ok);
            foreach (var a in answers.Values) foreach (var row in a.Rows) row.ClassNo = r.LectureMode ? a.ClassNo : row.HomeClassNo;
            foreach (var p in perfs.Values) foreach (var row in p.Rows) row.ClassNo = r.LectureMode ? p.ClassNo : row.HomeClassNo;

            // ---------- 학생 명단 합치기 ----------
            var students = new Dictionary<string, Student>();
            Student Get(int unit, int home, int num, string id, string name, string source)
            {
                var key = home + "/" + num;
                Student st;
                if (!students.TryGetValue(key, out st))
                {
                    st = new Student { ClassNo = unit, HomeClassNo = home, Number = num, StudentId = id, Name = name, Lecture = r.LectureMode };
                    students[key] = st;
                }
                else
                {
                    if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(st.Name) && st.Name != name)
                        r.Warnings.Add($"{home}반 {num}번 이름이 파일마다 다릅니다: [{st.Name}] / [{name}] ({source}). 반·번호를 확인하세요.");
                    else if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(st.StudentId) && st.StudentId != id)
                        r.Warnings.Add($"{home}반 {num}번 {name} 학번이 파일마다 다릅니다: {st.StudentId} / {id} ({source}).");
                    if (st.ClassNo != unit)
                        r.Warnings.Add($"{home}반 {num}번 {name}: 파일마다 {r.Unit}이 다릅니다({st.ClassNo}{r.Unit} / {unit}{r.Unit}, {source}). 먼저 읽은 파일의 {r.Unit}으로 둡니다.");
                }
                return st;
            }
            foreach (var p in perfs.Values)
                foreach (var row in p.Rows)
                {
                    var st = Get(row.ClassNo, row.HomeClassNo, row.Number, row.StudentId, row.Name, "수행평가");
                    st.Perf = row;
                }
            foreach (var kv in answers)
                foreach (var row in kv.Value.Rows)
                {
                    var st = Get(row.ClassNo, row.HomeClassNo, row.Number, row.StudentId, row.Name, kv.Key.Item1 + " 정오표");
                    st.Written[kv.Key.Item1] = row;
                }
            foreach (var t in tables)
                if (!r.LectureMode)   // 강의실 모드에서는 일람표의 번호가 행정반 번호와 섞여 학생을 특정할 수 없다
                foreach (var cls in t.Value.Cells)
                    foreach (var cell in cls.Value)
                        if (cell.Value.Exempt)
                        {
                            Student st;
                            if (students.TryGetValue(cls.Key + "/" + cell.Key, out st)) st.WrittenExemptByTable.Add(t.Key);
                        }
            r.Students = students.Values.OrderBy(x => x.ClassNo).ThenBy(x => x.HomeClassNo).ThenBy(x => x.Number).ToList();

            // ---------- 학생별 점수 ----------
            int areaCount = s.PerfAreas.Count;
            foreach (var st in r.Students)
            {
                var reasons = new List<string>();
                // 수행
                if (st.Perf != null)
                {
                    int exempt = st.Perf.Exempt.Count(e => e);
                    int blank = 0;
                    for (int i = 0; i < st.Perf.Scores.Length; i++) if (!st.Perf.Scores[i].HasValue && !st.Perf.Exempt[i]) blank++;
                    st.PerfAllExempt = exempt > 0 && exempt == st.Perf.Scores.Length;
                    st.PerfPartialExempt = exempt > 0 && !st.PerfAllExempt;
                    st.PerfMissingArea = blank > 0;
                    double sum = st.Perf.Scores.Where(x => x.HasValue).Sum(x => x.Value);
                    st.PerfTotal = st.PerfAllExempt ? (double?)null : (st.Perf.Total ?? sum);
                    if (st.Perf.Total.HasValue && Math.Abs(st.Perf.Total.Value - sum) > 0.001)
                        r.Warnings.Add($"{st.Label}: 수행 영역 점수 합({sum})이 나이스 합계({st.Perf.Total})와 다릅니다. 나이스 합계를 사용했습니다.");
                    if (st.PerfPartialExempt) st.Flags.Add("수행 일부 영역 면제(합계는 나이스 값 그대로)");
                    if (st.PerfMissingArea) st.Flags.Add("수행 점수가 비어 있는 영역 있음");
                }
                // 지필
                double convSum = 0; bool writtenComplete = true;
                var missingExams = new List<string>();
                foreach (var e in r.Exams)
                {
                    AnswerRow row;
                    if (st.Written.TryGetValue(e, out row))
                    {
                        double ratio = s.ExamRatios.ContainsKey(e) ? s.ExamRatios[e] : 0;
                        var file = answers.First(a => a.Key.Item1 == e).Value;
                        double max = file.TotalMax > 0 ? file.TotalMax : 100;
                        double conv = Rounding.Round(row.Total / max * ratio, s.WrittenDecimals);
                        st.WrittenConverted[e] = conv;
                        convSum += conv;
                        st.WrittenRawSum = (st.WrittenRawSum ?? 0) + row.Total;
                        st.BlankCount += row.Answers.Count(a => a == "-");
                    }
                    else { writtenComplete = false; missingExams.Add(e); }
                }
                // 상태 판정
                bool perfOk = !hasPerfFiles || (st.Perf != null && !st.PerfAllExempt);
                if (st.Perf == null && hasPerfFiles && perfs.ContainsKey(st.ClassNo)) reasons.Add("수행평가 명단에 없음");
                if (st.Perf == null && hasPerfFiles && !perfs.ContainsKey(st.ClassNo)) reasons.Add("해당 반 수행평가 파일 없음");
                if (st.PerfAllExempt) reasons.Add("수행평가 전체 면제");
                foreach (var e in missingExams)
                    reasons.Add(st.WrittenExemptByTable.Contains(e) ? $"{e} 면제(일람표)" : $"{e} 정오표에 없음(결시 또는 면제)");
                st.Included = perfOk && writtenComplete && r.Exams.Count > 0;
                if (!hasPerfFiles && r.Exams.Count > 0) st.Included = writtenComplete;
                if (st.Included)
                {
                    double final = (st.PerfTotal ?? 0) + convSum;
                    st.Final = Rounding.Round(final, s.FinalDecimals);
                    st.AchievementBasisValue = s.Basis == AchievementBasis.RoundedInteger ? Rounding.Round(st.Final.Value, 0) : st.Final.Value;
                    st.Achievement = Grade(st.AchievementBasisValue.Value, s);
                    // 경계
                    foreach (var pair in new[] { ("A", s.A), ("B", s.B), ("C", s.C), ("D", s.D) })
                    {
                        double diff = st.Final.Value - pair.Item2;
                        if (Math.Abs(diff) <= s.BoundaryRange + 1e-9)
                        {
                            st.Boundary = true;
                            st.BoundaryNote = $"{pair.Item1} 경계({(diff >= 0 ? "+" : "")}{diff:0.##})";
                            break;
                        }
                    }
                    st.Status = reasons.Count == 0 ? "정상" : string.Join(", ", reasons);
                    if (st.PerfTotal.HasValue && r.PerfMax > 0) st.PerfPercent = st.PerfTotal.Value / r.PerfMax * 100;
                    if (st.WrittenRawSum.HasValue)
                    {
                        double totalMax = r.Exams.Sum(e => { var f = answers.First(a => a.Key.Item1 == e).Value; return f.TotalMax > 0 ? f.TotalMax : 100; });
                        st.WrittenPercent = st.WrittenRawSum.Value / totalMax * 100;
                    }
                    if (st.PerfPercent.HasValue && st.WrittenPercent.HasValue) st.PerfWrittenGap = st.PerfPercent - st.WrittenPercent;
                }
                else
                {
                    bool everythingExempt = st.Written.Count == 0 && (st.Perf == null || st.PerfAllExempt);
                    st.Status = everythingExempt && st.PerfAllExempt ? "면제" : string.Join(", ", reasons);
                    if (st.Status.Length == 0) st.Status = "제외";
                }
            }

            // ---------- 등수 ----------
            r.Ranked = r.Students.Where(x => x.Included).OrderByDescending(x => x.Final).ThenBy(x => x.ClassNo).ThenBy(x => x.Number).ToList();
            AssignRanks(r.Ranked, x => x.Final.Value, (x, rk) => x.GradeRank = rk);
            foreach (var g in r.Ranked.GroupBy(x => x.ClassNo))
                AssignRanks(g.ToList(), x => x.Final.Value, (x, rk) => x.ClassRank = rk);

            // ---------- 학년 통계 ----------
            if (r.Ranked.Count > 0)
            {
                var finals = r.Ranked.Select(x => x.Final.Value).ToList();
                r.Mean = finals.Average(); r.Median = Median(finals); r.Sd = Sd(finals); r.Max = finals.Max(); r.Min = finals.Min();
                foreach (var g in Grades) r.AchievementCounts[g] = r.Ranked.Count(x => x.Achievement == g);
                foreach (var st in r.Ranked) r.Histogram[Bin(st.Final.Value)]++;
            }

            // ---------- 성취도별 통계 ----------
            foreach (var g in Grades)
            {
                var grp = r.Ranked.Where(x => x.Achievement == g).ToList();
                var ls = new LevelSummary { Grade = g, Count = grp.Count };
                if (grp.Count > 0)
                {
                    var f = grp.Select(x => x.Final.Value).ToList();
                    ls.Mean = f.Average(); ls.Sd = Sd(f); ls.Min = f.Min(); ls.Max = f.Max();
                    var pt = grp.Where(x => x.PerfTotal.HasValue).Select(x => x.PerfTotal.Value).ToList();
                    if (pt.Count > 0) ls.PerfMean = pt.Average();
                    foreach (var e in r.Exams)
                    {
                        var w = grp.Where(x => x.Written.ContainsKey(e)).Select(x => x.Written[e].Total).ToList();
                        if (w.Count > 0) ls.WrittenMeans[e] = w.Average();
                    }
                }
                r.LevelSummaries.Add(ls);
            }

            // ---------- 반별 통계 ----------
            foreach (var cls in r.Classes)
            {
                var all = r.Students.Where(x => x.ClassNo == cls).ToList();
                var inc = all.Where(x => x.Included).ToList();
                var cs = new ClassSummary { ClassNo = cls, Count = inc.Count, Enrolled = all.Count, PerfAreaMeans = new double[areaCount] };
                r.ClassHistogram[cls] = new int[10];
                if (inc.Count > 0)
                {
                    var f = inc.Select(x => x.Final.Value).ToList();
                    cs.Mean = f.Average(); cs.Median = Median(f); cs.Sd = Sd(f); cs.Max = f.Max(); cs.Min = f.Min();
                    for (int i = 0; i < areaCount; i++)
                    {
                        var vals = inc.Where(x => x.Perf != null && i < x.Perf.Scores.Length && x.Perf.Scores[i].HasValue).Select(x => x.Perf.Scores[i].Value).ToList();
                        cs.PerfAreaMeans[i] = vals.Count > 0 ? vals.Average() : 0;
                    }
                    var pt = inc.Where(x => x.PerfTotal.HasValue).Select(x => x.PerfTotal.Value).ToList();
                    cs.PerfMean = pt.Count > 0 ? pt.Average() : 0;
                    foreach (var e in r.Exams)
                    {
                        var vals = inc.Where(x => x.Written.ContainsKey(e)).Select(x => x.Written[e].Total).ToList();
                        cs.WrittenMeans[e] = vals.Count > 0 ? vals.Average() : 0;
                    }
                    foreach (var g in Grades) cs.AchievementCounts[g] = inc.Count(x => x.Achievement == g);
                    foreach (var st in inc) r.ClassHistogram[cls][Bin(st.Final.Value)]++;
                }
                r.ClassSummaries.Add(cs);
            }
            var byMean = r.ClassSummaries.Where(c => c.Count > 0).OrderByDescending(c => c.Mean).ToList();
            for (int i = 0; i < byMean.Count; i++) byMean[i].MeanRank = i + 1;

            // ---------- 수행 vs 지필 상관 ----------
            var pairs = r.Ranked.Where(x => x.PerfPercent.HasValue && x.WrittenPercent.HasValue).ToList();
            if (pairs.Count >= 3) r.PerfWrittenCorr = Pearson(pairs.Select(x => x.PerfPercent.Value).ToList(), pairs.Select(x => x.WrittenPercent.Value).ToList());
            // 수행은 대체로 지필보다 높으므로, 학년 평균 차이를 뺀 "상대 차이"로 특이 학생을 찾는다
            if (pairs.Count > 0)
            {
                r.AvgPerfWrittenGap = pairs.Average(x => x.PerfWrittenGap.Value);
                foreach (var st in pairs) st.PerfWrittenGapAdj = st.PerfWrittenGap.Value - r.AvgPerfWrittenGap;
            }

            // ---------- 문항 분석 (고사별) ----------
            foreach (var e in r.Exams)
            {
                var ea = AnalyzeExam(e, answers.Where(a => a.Key.Item1 == e).ToDictionary(a => a.Key.Item2, a => a.Value), s, r);
                ScoreTableFile t;
                if (tables.TryGetValue(e, out t) || (tables.Count == 1 && r.Exams.Count == 1 && (t = tables.Values.First()) != null))
                    Verify(ea, t, answers.Where(a => a.Key.Item1 == e).ToDictionary(a => a.Key.Item2, a => a.Value), r.LectureMode);
                r.ExamAnalyses.Add(ea);
            }

            // ---------- 문항정보표(예상 난이도) 붙이기 ----------
            foreach (var f in ok.Where(x => x.Kind == NeisFileKind.ItemInfoSheet))
            {
                var exam = MatchExam(f.Info.Header.ExamName, r.Exams);
                var ea = exam == null ? null : r.ExamAnalyses.FirstOrDefault(x => x.ExamName == exam);
                if (ea == null || ea.Info != null) continue;
                ea.Info = f.Info;
                var metas = f.Info.Items.Where(m => m.Section == "선택형").ToList();
                int choiceCount = ea.ItemStats.Count(i => i.Correct.HasValue);
                if (metas.Count != choiceCount)
                    r.Warnings.Add($"[{exam}] 문항정보표의 선택형 문항 수({metas.Count})와 정오표의 선택형 문항 수({choiceCount})가 다릅니다. 번호가 같은 문항끼리만 비교합니다.");
                var mismatch = new List<string>();
                foreach (var m in metas)
                {
                    var st = ea.ItemStats.FirstOrDefault(i => i.Number == m.Number && i.Correct.HasValue);
                    if (st == null) continue;
                    st.Area = m.Area; st.Standard = m.Standard; st.Expected = m.Expected;
                    int ans;
                    if (int.TryParse(m.CorrectRaw ?? "", out ans) && ans != st.Correct.Value) mismatch.Add($"{m.Number}번(정보표 {ans}, 정오표 {st.Correct})");
                    if (m.Points > 0 && Math.Abs(m.Points - st.Points) > 0.001) mismatch.Add($"{m.Number}번 배점(정보표 {m.Points:0.##}, 정오표 {st.Points:0.##})");
                }
                if (mismatch.Count > 0)
                    r.Warnings.Add($"[{exam}] 문항정보표와 정오표의 정답·배점이 다른 문항: {string.Join(", ", mismatch)}. 다른 시험의 문항정보표이거나 정답 정정이 있었는지 확인하세요.");
            }

            // ---------- 관심 학생 표시 ----------
            foreach (var st in r.Students)
            {
                if (st.Achievement == "E") st.Flags.Add("성취도 E");
                if (st.Boundary) st.Flags.Add("성취도 경계: " + st.BoundaryNote);
                int itemCount = r.ExamAnalyses.Sum(x => x.Items.Count(i => i.IsChoice));
                if (itemCount > 0 && (st.BlankCount >= 3 || st.BlankCount >= itemCount * 0.1)) st.Flags.Add($"무표기 {st.BlankCount}개");
                if (st.PerfWrittenGapAdj.HasValue && Math.Abs(st.PerfWrittenGapAdj.Value) >= GapCut)
                    st.Flags.Add(st.PerfWrittenGapAdj > 0 ? $"수행에 비해 지필이 유난히 낮음(상대 차이 {st.PerfWrittenGapAdj:0}점)" : $"지필에 비해 수행이 유난히 낮음(상대 차이 {-st.PerfWrittenGapAdj:0}점)");
                if (!st.Included) st.Flags.Add("통계 제외: " + st.Status);
                if (r.Exams.Count >= 2 && st.Written.Count == r.Exams.Count)
                {
                    double first = st.Written[r.Exams[0]].Total, last = st.Written[r.Exams[r.Exams.Count - 1]].Total;
                    if (Math.Abs(last - first) >= 15) st.Flags.Add(last > first ? $"{r.Exams[0]}→{r.Exams.Last()} {last - first:+0;-0}점 상승" : $"{r.Exams[0]}→{r.Exams.Last()} {last - first:+0;-0}점 하락");
                }
            }

            BuildNotes(r);
            foreach (var ea in r.ExamAnalyses.Where(x => x.Info != null))
                r.Notes.Add($"문항정보표({ea.ExamName}): 출제 교사가 시험 전에 적은 예상 난이도(어려움·보통·쉬움)를 실제 정답률로 정한 난이도와 비교했습니다. 실제 난이도는 정답률 {s.EasyCut:P0} 이상 쉬움, {s.HardCut:P0} 이상 보통, 그 미만 어려움 기준입니다.");
            if (r.LectureMode) r.Notes.Add("고교학점제 강의실 모드: 파일 머리말이 '강의실'이거나 한 파일에 여러 반 학생이 섞여 있어, 반 대신 강의실(수업 집단) 단위로 비교했습니다. 학생은 행정반/번호로 식별하고 '반' 열에 행정반, '강의실' 열에 수업 집단을 적었습니다. 검산의 일람표도 강의실 단위로 비교합니다.");
            r.Notes.Add("표준편차는 모집단 방식(인원수 N으로 나눔)으로 계산했습니다. 교육과정평가원(KICE) 성취평가 분석 프로그램과 같은 방식이며, 엑셀 STDEV.S(표본 방식, N−1로 나눔)보다 조금 작게 나옵니다.");
            r.Notes.Add($"성취도별 문항 정답률의 성취도는 최종 점수(수행+지필) 기준 성취도입니다. 낮은 성취도 집단이 높은 성취도 집단보다 {LevelInversionGap:P0} 이상 많이 맞힌 문항은 '수준 역전'으로 표시합니다(두 집단 모두 {LevelInversionMinN}명 이상일 때).");
            BuildHighlights(r);
            return r;
        }

        static ExamAnalysis AnalyzeExam(string exam, Dictionary<int, AnswerSheetFile> byClass, AnalysisSettings s, AnalysisResult r)
        {
            var first = byClass.Values.First();
            var ea = new ExamAnalysis { ExamName = exam, Items = first.Items, TotalMax = first.TotalMax, ChoiceMax = first.ChoiceMax, EssayMax = first.EssayMax };
            // 문항 구성이 반마다 같은지 확인
            foreach (var kv in byClass)
                if (kv.Value.Items.Count != first.Items.Count || !kv.Value.Items.Select(i => i.CorrectRaw + "/" + i.Points).SequenceEqual(first.Items.Select(i => i.CorrectRaw + "/" + i.Points)))
                    r.Warnings.Add($"{kv.Key}{r.Unit} {exam} 정오표의 문항 수·정답·배점이 다른 {r.Unit}과 다릅니다. 문항 분석은 첫 {r.Unit} 기준으로 합니다.");
            var rows = byClass.Values.SelectMany(f => f.Rows.Where(x => x.Answers.Length == first.Items.Count)).ToList();
            ea.N = rows.Count;
            if (rows.Count == 0) return ea;
            var totals = rows.Select(x => x.Total).ToList();
            ea.Mean = totals.Average(); ea.Sd = Sd(totals);
            foreach (var g in rows.GroupBy(x => x.ClassNo))
                ea.ClassStats[g.Key] = new ClassStat { ClassNo = g.Key, Count = g.Count(), Sum = g.Sum(x => x.Total), Mean = g.Average(x => x.Total) };

            // 상·하위 27%
            var sorted = rows.OrderByDescending(x => x.Total).ToList();
            int k = Math.Max(1, (int)Math.Round(rows.Count * 0.27));
            var upper = new HashSet<AnswerRow>(sorted.Take(k));
            var lower = new HashSet<AnswerRow>(sorted.Skip(rows.Count - k));

            int maxOpt = 5;
            foreach (var row in rows) foreach (var a in row.Answers) { int d; if (int.TryParse(a, out d)) maxOpt = Math.Max(maxOpt, d); }
            foreach (var it in first.Items) if (it.Correct.HasValue) maxOpt = Math.Max(maxOpt, it.Correct.Value);
            ea.OptionCount = Math.Min(maxOpt, 9);

            // 선택형 총점(신뢰도 계산용)
            var choiceTotals = rows.Select(row => { double t = 0; for (int i = 0; i < first.Items.Count; i++) if (first.Items[i].IsChoice && row.Answers[i] == ".") t += first.Items[i].Points; return t; }).ToList();
            double varTotal = Variance(choiceTotals);
            double sumItemVar = 0; int kChoice = 0;
            // 학생별 성취도(최종 점수 기준) — 성취도별 정답률 계산용
            var levelByKey = new Dictionary<string, string>();
            foreach (var st0 in r.Students) if (st0.Achievement != null) levelByKey[st0.Key] = st0.Achievement;

            for (int i = 0; i < first.Items.Count; i++)
            {
                var it = first.Items[i];
                var st = new ItemStat { Number = it.Number, Points = it.Points, Correct = it.Correct, N = rows.Count };
                for (int o = 1; o <= ea.OptionCount; o++) st.OptionCounts[o] = 0;
                var correctFlags = new List<double>();
                int upC = 0, loC = 0;
                foreach (var row in rows)
                {
                    var a = row.Answers[i];
                    bool correct = a == ".";
                    int d;
                    if (correct) { st.CorrectCount++; if (it.Correct.HasValue && st.OptionCounts.ContainsKey(it.Correct.Value)) st.OptionCounts[it.Correct.Value]++; }
                    else if (a == "-" || a.Length == 0) st.BlankCount++;
                    else if (int.TryParse(a, out d)) { if (st.OptionCounts.ContainsKey(d)) st.OptionCounts[d]++; }
                    else st.MultiCount++;
                    correctFlags.Add(correct ? 1 : 0);
                    if (upper.Contains(row) && correct) upC++;
                    if (lower.Contains(row) && correct) loC++;
                    if (!st.ClassN.ContainsKey(row.ClassNo)) { st.ClassN[row.ClassNo] = 0; st.ClassCorrectRate[row.ClassNo] = 0; }
                    st.ClassN[row.ClassNo]++;
                    if (correct) st.ClassCorrectRate[row.ClassNo]++;
                    string lv;
                    if (levelByKey.TryGetValue(row.ClassNo + "/" + row.Number, out lv))
                    {
                        if (!st.LevelN.ContainsKey(lv)) { st.LevelN[lv] = 0; st.LevelCorrectRate[lv] = 0; }
                        st.LevelN[lv]++;
                        if (correct) st.LevelCorrectRate[lv]++;
                    }
                }
                foreach (var c in st.ClassN.Keys.ToList()) st.ClassCorrectRate[c] = st.ClassN[c] == 0 ? 0 : st.ClassCorrectRate[c] / st.ClassN[c];
                foreach (var lv in st.LevelN.Keys.ToList()) st.LevelCorrectRate[lv] = st.LevelN[lv] == 0 ? 0 : st.LevelCorrectRate[lv] / st.LevelN[lv];
                if (it.Correct.HasValue) st.LevelInversion = FindLevelInversion(st);
                st.UpperRate = (double)upC / k; st.LowerRate = (double)loC / k;
                st.Discrimination = st.UpperRate - st.LowerRate;
                st.PointBiserial = Pearson(correctFlags, totals);
                double p = st.CorrectRate;
                st.Difficulty = p >= s.EasyCut ? "쉬움" : p >= s.HardCut ? "보통" : "어려움";
                if (it.Correct.HasValue)
                {
                    var wrong = st.OptionCounts.Where(o => o.Key != it.Correct.Value).OrderByDescending(o => o.Value).FirstOrDefault();
                    if (wrong.Value > 0) { st.AttractiveWrong = wrong.Key; st.AttractiveWrongRate = (double)wrong.Value / rows.Count; }
                    sumItemVar += p * (1 - p) * it.Points * it.Points;
                    kChoice++;
                }
                ea.ItemStats.Add(st);
            }
            if (kChoice >= 2 && varTotal > 0) ea.KR20 = kChoice / (kChoice - 1.0) * (1 - sumItemVar / varTotal);
            return ea;
        }

        /// <summary>낮은 성취도 집단의 정답률이 높은 성취도 집단보다 15%p 이상 높으면 "수준 역전"(두 집단 모두 5명 이상일 때만).</summary>
        public const double LevelInversionGap = 0.15;
        public const int LevelInversionMinN = 5;
        static string FindLevelInversion(ItemStat st)
        {
            var found = new List<string>();
            for (int hi = 0; hi < Grades.Length; hi++)
                for (int lo = hi + 1; lo < Grades.Length; lo++)
                {
                    string h = Grades[hi], l = Grades[lo];
                    if (!st.LevelN.ContainsKey(h) || !st.LevelN.ContainsKey(l) || st.LevelN[h] < LevelInversionMinN || st.LevelN[l] < LevelInversionMinN) continue;
                    double gap = st.LevelCorrectRate[l] - st.LevelCorrectRate[h];
                    if (gap >= LevelInversionGap) found.Add($"{l}({st.LevelCorrectRate[l]:P0})가 {h}({st.LevelCorrectRate[h]:P0})보다 높음");
                }
            return found.Count == 0 ? null : "수준 역전: " + string.Join(", ", found);
        }

        static void Verify(ExamAnalysis ea, ScoreTableFile t, Dictionary<int, AnswerSheetFile> byClass, bool lecture)
        {
            ea.Table = t;
            bool all = true;
            var classes = ea.ClassStats.Keys.Union(t.Stats.Keys).OrderBy(c => c).ToList();
            foreach (var cls in classes)
            {
                var v = new VerifyRow { ClassNo = cls };
                ClassStat cs;
                if (ea.ClassStats.TryGetValue(cls, out cs)) { v.CalcCount = cs.Count ?? 0; v.CalcSum = cs.Sum ?? 0; v.CalcMean = Math.Round(cs.Mean ?? 0, 1, MidpointRounding.AwayFromZero); }
                ClassStat ts;
                if (t.Stats.TryGetValue(cls, out ts)) { v.TableCount = ts.Count; v.TableSum = ts.Sum; v.TableMean = ts.Mean; }
                if (!byClass.ContainsKey(cls)) { v.Match = false; v.Result = "정오표 없음(비교 불가)"; ea.Verify.Add(v); continue; }
                else if (!t.Stats.ContainsKey(cls)) { v.Match = false; v.Result = "일람표에 없는 반"; }
                else
                {
                    v.Match = v.TableCount == v.CalcCount && Math.Abs((v.TableSum ?? -1) - v.CalcSum) < 0.001 && Math.Abs((v.TableMean ?? -1) - v.CalcMean) < 0.051;
                    v.Result = v.Match ? "일치" : "불일치";
                }
                if (!v.Match) all = false;
                ea.Verify.Add(v);
                // 학생별 비교 (강의실 모드에서는 번호가 여러 반에 걸쳐 겹치므로 생략)
                if (!lecture && byClass.ContainsKey(cls) && t.Cells.ContainsKey(cls))
                {
                    var rows = byClass[cls].Rows.ToDictionary(x => x.Number, x => x);
                    foreach (var cell in t.Cells[cls])
                    {
                        AnswerRow row;
                        if (cell.Value.Exempt) { if (rows.ContainsKey(cell.Key)) ea.VerifyStudentIssues.Add($"{cls}반 {cell.Key}번: 일람표는 면제인데 정오표에 답안이 있습니다."); continue; }
                        if (!rows.TryGetValue(cell.Key, out row)) { ea.VerifyStudentIssues.Add($"{cls}반 {cell.Key}번: 일람표에 점수({cell.Value.Score})가 있는데 정오표에 없습니다."); all = false; continue; }
                        if (Math.Abs(row.Total - (cell.Value.Score ?? -1)) > 0.001) { ea.VerifyStudentIssues.Add($"{cls}반 {cell.Key}번 {row.Name}: 정오표 총점 {row.Total} ≠ 일람표 {cell.Value.Score}"); all = false; }
                    }
                    foreach (var row in rows.Values)
                        if (!t.Cells[cls].ContainsKey(row.Number)) { ea.VerifyStudentIssues.Add($"{cls}반 {row.Number}번 {row.Name}: 정오표에는 있는데 일람표에 없습니다."); all = false; }
                }
            }
            ea.VerifyAllMatch = all;
        }

        static void BuildNotes(AnalysisResult r)
        {
            var s = r.Settings;
            if (r.HasPerf)
                r.Notes.Add("수행평가: " + string.Join(", ", s.PerfAreas.Select(a => $"{a.Name} 만점 {a.Max:0.##}점({a.Ratio:0.##}%)")) + $" → 합계 {r.PerfMax:0.##}점({s.PerfRatioSum:0.##}%)");
            foreach (var e in r.Exams)
                r.Notes.Add($"지필평가 {e}: 반영 비율 {(s.ExamRatios.ContainsKey(e) ? s.ExamRatios[e] : 0):0.##}% (환산 점수 = 원점수 ÷ 만점 × 비율, 소수 {s.WrittenDecimals}자리에서 반올림)");
            r.Notes.Add($"최종 점수 = 수행 합계 + 지필 환산 점수 (소수 {s.FinalDecimals}자리, 반올림은 '5 이상 올림' 방식)");
            r.Notes.Add($"성취도 기준: A {s.A:0.##}점 이상, B {s.B:0.##}점 이상, C {s.C:0.##}점 이상, D {s.D:0.##}점 이상, 그 미만 E");
            r.Notes.Add(s.Basis == AchievementBasis.RoundedInteger
                ? "성취도 판정 점수: 최종 점수를 소수 첫째 자리에서 반올림한 정수(학교생활기록부 원점수 방식). 예) 89.5 → 90 → A"
                : "성취도 판정 점수: 반올림하지 않은 최종 점수. 예) 89.5 → B");
            r.Notes.Add($"성취도 경계 표시: 최종 점수가 기준 점수의 ±{s.BoundaryRange:0.##}점 안에 있는 학생");
            r.Notes.Add("등수: 최종 점수가 같으면 같은 등수, 다음 등수는 동점자 수만큼 건너뜀 (예: 1, 1, 3)");
            r.Notes.Add($"문항 난이도: 정답률 {s.EasyCut:P0} 이상 쉬움, {s.HardCut:P0} 이상 보통, 그 미만 어려움");
            r.Notes.Add("변별도: 지필 총점 상위 27% 집단 정답률 − 하위 27% 집단 정답률. 0.3 이상 양호, 0.1 미만 또는 음수는 점검 필요");
            r.Notes.Add("문항-총점 상관(점이연 상관): 문항 정오(맞음 1, 틀림 0)와 지필 총점의 상관계수. 0.2 미만이면 변별력이 약함");
            r.Notes.Add("신뢰도 KR-20: 선택형 문항 전체가 얼마나 일관되게 실력을 재는지(0~1). 0.7 이상이면 보통, 0.8 이상이면 양호");
            r.Notes.Add($"반별 문항 정답률 강조: 학년 정답률보다 {s.ClassItemGap:P0} 이상 낮은 반-문항");
            r.Notes.Add($"수행 vs 지필 차이: 두 점수를 100점 만점 비율로 바꾼 뒤 (수행 비율 − 지필 비율)을 구하고, 학년 평균 차이를 뺀 '상대 차이'가 ±{GapCut}점 이상이면 관심 학생으로 표시");
            r.Notes.Add("답지 반응 분포: 맞은 학생은 정답 번호를 고른 것으로 계산(정오표에는 틀린 학생이 고른 번호만 있음). 복수답안·무표기는 따로 집계");
            r.Notes.Add("통계 제외 학생: 수행 전체 면제, 지필 결시·면제, 명단 불일치 학생은 평균·등수·성취도 분포에서 제외하고 '관심 학생' 시트에 사유와 함께 표시");
        }

        static void BuildHighlights(AnalysisResult r)
        {
            var h = r.Highlights;
            if (r.Count == 0) { h.Add("통계에 포함할 학생이 없습니다."); return; }
            h.Add($"분석 대상 {r.Count}명(명단 {r.Students.Count}명 중 {r.Students.Count - r.Count}명 제외). 학년 평균 {r.Mean:0.0}점, 중앙값 {r.Median:0.0}점, 표준편차 {r.Sd:0.0}점, 최고 {r.Max:0.##}점, 최저 {r.Min:0.##}점.");
            h.Add("성취도 분포: " + string.Join(", ", Grades.Select(g => $"{g} {r.AchievementCounts[g]}명({(double)r.AchievementCounts[g] / r.Count:P0})")) + ".");
            var cs = r.ClassSummaries.Where(c => c.Count > 0).OrderByDescending(c => c.Mean).ToList();
            if (cs.Count >= 2) h.Add($"{r.Unit}별 평균이 가장 높은 {r.Unit}은 {cs.First().ClassNo}{r.Unit}({cs.First().Mean:0.0}점), 가장 낮은 {r.Unit}은 {cs.Last().ClassNo}{r.Unit}({cs.Last().Mean:0.0}점)으로 {cs.First().Mean - cs.Last().Mean:0.0}점 차이가 납니다.");
            foreach (var ea in r.ExamAnalyses)
            {
                var items = ea.ItemStats.Where(i => i.Correct.HasValue).ToList();
                if (items.Count == 0) continue;
                string pre = r.Exams.Count > 1 ? $"[{ea.ExamName}] " : "";
                var hard = items.OrderBy(i => i.CorrectRate).First();
                var easy = items.OrderByDescending(i => i.CorrectRate).First();
                h.Add($"{pre}가장 어려운 문항은 {hard.Number}번(정답률 {hard.CorrectRate:P0}), 가장 쉬운 문항은 {easy.Number}번(정답률 {easy.CorrectRate:P0})입니다.");
                int nHard = items.Count(i => i.Difficulty == "어려움"), nEasy = items.Count(i => i.Difficulty == "쉬움");
                h.Add($"{pre}선택형 {items.Count}문항 중 쉬움 {nEasy}개, 보통 {items.Count - nHard - nEasy}개, 어려움 {nHard}개. 지필 평균 {ea.Mean:0.0}점(만점 {ea.TotalMax:0.##}점).");
                var low = items.Where(i => i.Discrimination.HasValue && i.Discrimination < 0.1).Select(i => i.Number.ToString()).ToList();
                if (low.Count > 0) h.Add($"{pre}변별도가 낮거나 음수인 문항: {string.Join(", ", low)}번. 문항 자체를 점검해 보세요.");
                var attr = items.Where(i => i.AttractiveWrong.HasValue).OrderByDescending(i => i.AttractiveWrongRate).First();
                if (attr.AttractiveWrongRate >= 0.3) h.Add($"{pre}{attr.Number}번에서는 학생 {attr.AttractiveWrongRate:P0}가 오답 {attr.AttractiveWrong}번을 골랐습니다(정답 {attr.Correct}번). 흔한 오개념일 수 있습니다.");
                if (ea.KR20.HasValue) h.Add($"{pre}시험 신뢰도(KR-20)는 {ea.KR20:0.00}로 {(ea.KR20 >= 0.8 ? "양호" : ea.KR20 >= 0.7 ? "보통" : "낮은 편")}입니다.");
                var blanks = ea.ItemStats.Sum(i => i.BlankCount);
                if (blanks > 0) h.Add($"{pre}무표기가 모두 {blanks}건 있습니다.");
                if (ea.Info != null)
                {
                    var cmp = items.Where(i => i.ExpectedGap.HasValue).ToList();
                    int same = cmp.Count(i => i.ExpectedGap == 0), harder = cmp.Count(i => i.ExpectedGap > 0), easier = cmp.Count(i => i.ExpectedGap < 0);
                    h.Add($"{pre}문항정보표의 예상 난이도와 실제가 같은 문항 {same}개, 예상보다 어려웠던 문항 {harder}개, 예상보다 쉬웠던 문항 {easier}개입니다." + (harder + easier > cmp.Count / 2 && cmp.Count > 0 ? " 예상과 다른 문항이 절반을 넘으니 난이도 판단 기준을 되짚어 볼 만합니다." : ""));
                }
                if (ea.VerifyAllMatch == true) h.Add($"{pre}검산: 정오표로 계산한 {r.Unit}별 응시생수·총점·평균이 나이스 일람표와 모두 일치합니다.");
                else if (ea.VerifyAllMatch == false) h.Add($"{pre}검산: 일람표와 맞지 않는 항목이 있습니다. '검산' 시트를 확인하세요.");
            }
            if (r.PerfWrittenCorr.HasValue) h.Add($"수행평가와 지필평가의 상관계수는 {r.PerfWrittenCorr:0.00}입니다{(r.PerfWrittenCorr >= 0.7 ? " (매우 비슷한 경향)" : r.PerfWrittenCorr >= 0.4 ? " (어느 정도 비슷한 경향)" : " (관련이 약함)")}.");
            int boundary = r.Ranked.Count(x => x.Boundary);
            if (boundary > 0) h.Add($"성취도 경계(±{r.Settings.BoundaryRange:0.##}점)에 있는 학생이 {boundary}명 있습니다. 반올림 규정을 한 번 더 확인하세요.");
        }

        // ---------- 도우미 ----------
        public static string Grade(double v, AnalysisSettings s) => v >= s.A ? "A" : v >= s.B ? "B" : v >= s.C ? "C" : v >= s.D ? "D" : "E";
        static int Bin(double v) { int b = (int)Math.Floor(v / 10); return Math.Max(0, Math.Min(9, b)); }
        public static string BinLabel(int b) => b == 9 ? "90~100" : $"{b * 10}~{b * 10 + 9}";

        /// <summary>동점자는 같은 등수, 다음 등수는 동점자 수만큼 건너뛴다 (1, 1, 3).</summary>
        static void AssignRanks(List<Student> list, Func<Student, double> score, Action<Student, int> set)
        {
            var sorted = list.OrderByDescending(score).ToList();
            int rank = 0;
            for (int i = 0; i < sorted.Count; i++)
            {
                if (i == 0 || Math.Abs(score(sorted[i]) - score(sorted[i - 1])) > 1e-9) rank = i + 1;
                set(sorted[i], rank);
            }
        }

        public static double Median(List<double> v) { var s = v.OrderBy(x => x).ToList(); int n = s.Count; if (n == 0) return 0; return n % 2 == 1 ? s[n / 2] : (s[n / 2 - 1] + s[n / 2]) / 2; }
        public static double Variance(List<double> v) { if (v.Count == 0) return 0; double m = v.Average(); return v.Sum(x => (x - m) * (x - m)) / v.Count; }
        public static double Sd(List<double> v) => Math.Sqrt(Variance(v));
        public static double? Pearson(List<double> x, List<double> y)
        {
            int n = Math.Min(x.Count, y.Count); if (n < 2) return null;
            double mx = x.Take(n).Average(), my = y.Take(n).Average();
            double sxy = 0, sxx = 0, syy = 0;
            for (int i = 0; i < n; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) * (x[i] - mx); syy += (y[i] - my) * (y[i] - my); }
            if (sxx == 0 || syy == 0) return null;
            return sxy / Math.Sqrt(sxx * syy);
        }
    }
}
