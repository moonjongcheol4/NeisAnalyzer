using System;
using System.Collections.Generic;
using System.Linq;

namespace NeisAnalyzer
{
    public enum NeisFileKind { Unknown, AnswerSheet, PerfSheet, ScoreTable, ItemInfoSheet }

    /// <summary>파일 상단에서 읽은 학년도·학기·학년·반·고사명·과목</summary>
    public class ExamHeader
    {
        public int Year;          // 학년도
        public int Semester;      // 학기
        public int Grade;         // 학년
        public int? ClassNo;      // 반 또는 강의실 번호 (학년 전체 파일이면 null)
        public bool LectureWord;  // 머리말에 "반"이 아니라 "강의실"로 적혀 있음 (고교학점제 정오표)
        public string ExamName;   // 예: "3학년 기말고사" (수행 파일은 null)
        public string Subject;    // 예: "과학"
        public string Describe()
        {
            var s = $"{Year}학년도 {Semester}학기 {Grade}학년";
            if (ClassNo.HasValue) s += $" {ClassNo}반";
            if (!string.IsNullOrEmpty(ExamName)) s += $" {ExamName}";
            if (!string.IsNullOrEmpty(Subject)) s += $" {Subject}";
            return s;
        }
    }

    /// <summary>문항정보표(출제 교사가 시험 전에 작성)의 문항 한 줄: 내용영역·성취기준·예상 난이도·배점·정답</summary>
    public class ItemMeta
    {
        public int Number;
        public string Section;        // "선택형" 또는 "서답형"
        public string Area;           // 내용영역
        public string Standard;       // 성취기준 (예: "[9과05-01] 물체에 작용하는 힘을 ...")
        public string Expected;       // 예상 난이도: 어려움 / 보통 / 쉬움 (없으면 null)
        public double Points;
        public string CorrectRaw;
        public string StandardCode => Standard == null ? null : (System.Text.RegularExpressions.Regex.Match(Standard, @"\[[^\]]+\]").Value.Length > 0 ? System.Text.RegularExpressions.Regex.Match(Standard, @"\[[^\]]+\]").Value : Standard);
    }

    public class ItemInfoFile
    {
        public string Path;
        public ExamHeader Header;     // Year·Semester·Grade·Subject·ExamName(예: "중간고사")
        public int TotalItems;
        public List<ItemMeta> Items = new List<ItemMeta>();
        public List<string> Warnings = new List<string>();
    }

    public class ItemInfo
    {
        public int Number;            // 문항 번호
        public string CorrectRaw;     // 정답 원문 (예: "3")
        public int? Correct;          // 정답 번호 (숫자로 읽힌 경우)
        public double Points;         // 배점
        public bool IsChoice => Correct.HasValue;   // 선택형(답지 분석 가능) 여부
    }

    public class AnswerRow
    {
        public int ClassNo, Number;   // ClassNo: 분석 단위(중학교=반, 고교 강의실 모드=강의실). HomeClassNo: 학생의 행정반
        public int HomeClassNo;
        public string StudentId, Name;
        public string[] Answers;      // 문항별 원문: "." 맞음, "2" 틀리게 고른 번호, "A" 복수답안코드, "-" 무표기
        public double ChoiceScore, EssayScore, EtcScore, Total;
    }

    public class AnswerSheetFile
    {
        public string Path;
        public ExamHeader Header;
        public int ClassNo;
        public List<ItemInfo> Items = new List<ItemInfo>();
        public double ChoiceMax, EssayMax, EtcMax, TotalMax;
        public List<AnswerRow> Rows = new List<AnswerRow>();
        public List<string> Warnings = new List<string>();
        public int MixedCount;     // 파일의 반(강의실) 번호와 다른 반 학생 수 → 0보다 크면 고교학점제 강의실 파일
    }

    public class PerfArea
    {
        public string Name;
        public double Max;
        public double Ratio;   // 반영 비율(%)
    }

    public class PerfRow
    {
        public int ClassNo, Number;   // ClassNo: 분석 단위(반 또는 강의실). HomeClassNo: 행정반
        public int HomeClassNo;
        public string StudentId, Name;
        public double?[] Scores;   // null = 면제 또는 빈칸
        public bool[] Exempt;
        public double? Total;      // 나이스가 계산한 합계
        public string Note;
    }

    public class PerfFile
    {
        public string Path;
        public ExamHeader Header;
        public int ClassNo;
        public List<PerfArea> Areas = new List<PerfArea>();
        public List<PerfRow> Rows = new List<PerfRow>();
        public List<string> Warnings = new List<string>();
        public int MixedCount;     // 강의실 번호와 다른 반 학생 수
    }

    public class ScoreCell { public double? Score; public bool Exempt; }

    public class ClassStat { public int ClassNo; public int? Count; public double? Sum; public double? Mean; }

    public class ScoreTableFile
    {
        public string Path;
        public ExamHeader Header;
        public double SubjectMax;
        public Dictionary<int, Dictionary<int, ScoreCell>> Cells = new Dictionary<int, Dictionary<int, ScoreCell>>(); // [반][번호]
        public Dictionary<int, ClassStat> Stats = new Dictionary<int, ClassStat>();
        public List<string> Warnings = new List<string>();
    }

    /// <summary>끌어다 놓은 파일 하나의 판별 결과</summary>
    public class DetectedFile
    {
        public string Path;
        public NeisFileKind Kind;
        public ExamHeader Header;
        public string Error;
        public AnswerSheetFile Answer;
        public PerfFile Perf;
        public ScoreTableFile Table;
        public ItemInfoFile Info;
        public string Label
        {
            get
            {
                if (Error != null) return "읽기 실패: " + Error;
                switch (Kind)
                {
                    case NeisFileKind.AnswerSheet: return $"정오표 - {Header.Grade}학년 {Header.ClassNo}{(Header.LectureWord || Answer.MixedCount > 0 ? "강의실" : "반")} {Header.ExamName} ({Header.Subject}, 문항 {Answer.Items.Count}개, 학생 {Answer.Rows.Count}명{(Answer.MixedCount > 0 ? ", 여러 반 섞임" : "")})";
                    case NeisFileKind.PerfSheet: return $"수행평가 - {Header.Grade}학년 {Header.ClassNo}{(Perf.MixedCount > 0 ? "강의실" : "반")} ({Header.Subject}, 영역 {Perf.Areas.Count}개, 학생 {Perf.Rows.Count}명{(Perf.MixedCount > 0 ? ", 여러 반 섞임" : "")})";
                    case NeisFileKind.ScoreTable: return $"일람표(검산용) - {Header.Grade}학년 전체 {Header.ExamName} ({Header.Subject}, {Table.Cells.Count}개 반)";
                    case NeisFileKind.ItemInfoSheet: return $"문항정보표 - {Header.Grade}학년 {Header.ExamName} ({Header.Subject}, 문항 {Info.Items.Count}개, 예상 난이도 어려움 {Info.Items.Count(i => i.Expected == "어려움")}·보통 {Info.Items.Count(i => i.Expected == "보통")}·쉬움 {Info.Items.Count(i => i.Expected == "쉬움")})";
                    default: return "알 수 없는 파일 (나이스 정오표·수행평가·일람표가 아님)";
                }
            }
        }
    }

    public enum AchievementBasis { RoundedInteger, FinalScore }

    public class AnalysisSettings
    {
        public List<PerfArea> PerfAreas = new List<PerfArea>();
        public Dictionary<string, double> ExamRatios = new Dictionary<string, double>(); // 고사명 → 반영 비율(%)
        public double A = 90, B = 80, C = 70, D = 60;
        public int WrittenDecimals = 2;        // 지필 환산 점수 소수 자릿수
        public int FinalDecimals = 2;          // 최종 점수 소수 자릿수
        public AchievementBasis Basis = AchievementBasis.RoundedInteger;
        public double BoundaryRange = 1.0;     // 성취도 경계 표시 범위(±점)
        public double EasyCut = 0.8, HardCut = 0.5;   // 난이도 구분(정답률)
        public double ClassItemGap = 0.15;     // 반별 문항 정답률이 학년보다 이만큼 낮으면 강조
        public double PerfRatioSum => PerfAreas.Sum(a => a.Ratio);
        public string ExamOrderKey(string exam)
        {
            // 중간 → 기말 순으로 정렬되게
            if (exam == null) return "9";
            if (exam.Contains("중간")) return "1" + exam;
            if (exam.Contains("기말")) return "2" + exam;
            return "5" + exam;
        }
    }

    public static class Rounding
    {
        public static double Round(double v, int decimals)
        {
            if (decimals < 0) return v;
            return Math.Round(v, decimals, MidpointRounding.AwayFromZero);
        }
    }
}
