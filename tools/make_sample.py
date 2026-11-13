# -*- coding: utf-8 -*-
"""
가짜 이름으로 된 나이스 모양 예시 파일을 만든다 (테스트·시연용).
실제 학생 자료는 전혀 쓰지 않는다. 구조(셀 위치·머리말·범례)는 실제 나이스 파일과 같게 맞춘다.

사용법:  python make_sample.py <출력폴더> [반 수] [--midterm]
  --midterm 을 주면 중간고사 정오표·일람표도 함께 만든다 (중간·기말 두 고사 테스트용).
"""
import os, random, sys
import openpyxl
from openpyxl.utils import get_column_letter

random.seed(20261112)

SUR = list("김이박최정강조윤장임한오서신권황안송류전홍고문양손배조백허유남심노정하곽성차주우구신임나전민유진지엄채원천방공강현함변염양변여추노도소신석선설마길주연방위표명기반왕금옥육인맹제모장남탁국여진어은편구용")
GIV = ["민준","서준","도윤","예준","시우","하준","지호","주원","지후","준서","서연","서윤","지우","서현","하은","하윤","민서","지유","윤서","채원",
       "수아","지민","지아","다은","은서","예은","수빈","소율","예린","지원","시은","유나","가은","연우","유진","하린","채은","승현","현우","건우",
       "우진","선우","연준","은우","유준","정우","승우","시윤","준혁","지환","태윤","예성","민성","지안","윤우","동현","승민","준영","찬영","재원"]

def name():
    return random.choice(SUR) + random.choice(GIV)

MULTI = {"A": (1, 2), "B": (1, 3), "C": (1, 4), "D": (1, 5), "E": (2, 3), "F": (2, 4), "G": (2, 5), "H": (3, 4), "I": (3, 5), "J": (4, 5)}
LEGEND = ("※ . : 맞음  ,  번호 : 틀림  ,  알파벳 : 복수답안코드  ,  - : 무표기\n"
          "A=1,2 , B=1,3 , C=1,4 , D=1,5 , E=2,3 , F=2,4 , G=2,5 , H=3,4 , I=3,5 , J=4,5 , K=1,2,3 , L=1,2,4 , M=1,2,5 , N=1,3,4 , O=1,3,5 ,\n"
          "P=1,4,5 , Q=2,3,4 , R=2,3,5 , S=2,4,5 , T=3,4,5 , U=1,2,3,4 , V=1,2,3,5 , W=1,2,4,5 , X=1,3,4,5 , Y=2,3,4,5 , Z=1,2,3,4,5\n"
          "※ 기타란의 점수는 기타 점수와 기타2 점수가 합산된 점수입니다.")

YEAR, SEM, GRADE, SUBJECT = 2026, 2, 3, "과학"
AREAS = [("체세포 분열", 15.0, 15.0), ("별과 우리은하", 15.0, 15.0), ("에너지 전환", 10.0, 10.0)]

def make_exam(num_items=26, essay=0):
    """정답·배점. 배점 합 = 100 - 서답형."""
    pts = [4] * num_items
    total = 100 - essay
    diff = total - sum(pts)
    i = 0
    while diff != 0:
        step = 1 if diff > 0 else -1
        pts[i % num_items] += step
        diff -= step
        i += 1
    correct = [random.randint(1, 5) for _ in range(num_items)]
    difficulty = [random.uniform(0.25, 0.95) for _ in range(num_items)]
    return correct, pts, difficulty

def make_students(num_classes):
    """반별 학생 목록: (번호, 학번, 이름, 능력치, 상태)"""
    classes = {}
    for c in range(1, num_classes + 1):
        n = random.choice([24, 25, 26, 27])
        skip = set(random.sample(range(1, n + 1), random.choice([0, 0, 1, 1, 2])))  # 빈 번호
        rows = []
        for num in range(1, n + 1):
            if num in skip:
                continue
            sid = 2024000000 + random.randint(1, 180)
            status = "정상"
            r = random.random()
            if r < 0.02: status = "전체면제"
            elif r < 0.035: status = "지필결시"
            elif r < 0.05: status = "수행일부면제"
            if num == n and random.random() < 0.5:
                sid = 2026000000 + random.randint(100, 199)   # 전입생
            rows.append((num, sid, name(), random.gauss(0.6, 0.18), status))
        classes[c] = rows
    return classes

def answers_for(ability, correct, difficulty):
    out = []
    for k, (ans, d) in enumerate(zip(correct, difficulty)):
        p = min(0.98, max(0.03, d + (ability - 0.6) * 1.2))
        rnd = random.random()
        if rnd < p: out.append(".")
        elif rnd > 0.985: out.append("-")                   # 무표기
        elif rnd > 0.975: out.append(random.choice(list(MULTI.keys())))   # 복수 답안
        else:
            wrong = [o for o in range(1, 6) if o != ans]
            attractive = wrong[k % 4]
            out.append(str(attractive if random.random() < 0.5 else random.choice(wrong)))
    return out

def style_title(ws):
    ws.column_dimensions["A"].width = 2

def write_answer_sheet(path, cls, rows, exam_name, correct, pts, difficulty, essay_max):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "sheet1"
    ws["U2"] = "2026.11.12."
    ws["J3"] = "정기시험교과목별 학생답 정오표"
    ws["B5"] = f"{YEAR}학년도   {SEM}학기   주간 {GRADE}학년   {cls}반   {exam_name}   {SUBJECT}"
    n = len(correct)
    ws.cell(7, 2, "반/번호"); ws.cell(7, 4, "번호")
    for i in range(n):
        ws.cell(7, 5 + i, float(i + 1))
    c_choice, c_essay, c_etc, c_total = 5 + n, 6 + n, 7 + n, 8 + n
    ws.cell(7, c_choice, "선택형점수"); ws.cell(7, c_essay, "서답형점수"); ws.cell(7, c_etc, "기타\n점수"); ws.cell(7, c_total, "영역\n총점")
    ws.cell(8, 4, "정답"); ws.cell(9, 4, "배점")
    for i in range(n):
        ws.cell(8, 5 + i, float(correct[i])); ws.cell(9, 5 + i, float(pts[i]))
    ws.cell(9, c_choice, float(sum(pts))); ws.cell(9, c_essay, float(essay_max)); ws.cell(9, c_etc, 0.0); ws.cell(9, c_total, float(sum(pts) + essay_max))
    r = 10
    results = {}
    for num, sid, nm, ability, status in rows:
        if status in ("전체면제", "지필결시"):
            continue
        ans = answers_for(ability, correct, difficulty)
        choice = sum(p for a, p in zip(ans, pts) if a == ".")
        essay = round(min(essay_max, max(0, random.gauss(ability * essay_max, essay_max * 0.2)))) if essay_max else 0
        ws.cell(r, 2, f"{cls}/{num}"); ws.cell(r, 3, float(sid)); ws.cell(r, 4, nm)
        for i, a in enumerate(ans):
            ws.cell(r, 5 + i, a if not a.isdigit() else float(a))
        ws.cell(r, c_choice, float(choice)); ws.cell(r, c_essay, float(essay)); ws.cell(r, c_etc, 0.0); ws.cell(r, c_total, float(choice + essay))
        results[num] = choice + essay
        r += 1
    ws.cell(r + 1, 2, LEGEND)
    ws.cell(r + 3, 20, 1.0); ws.cell(r + 3, 22, "/"); ws.cell(r + 3, 24, 1.0); ws.cell(r + 3, 30, "예시중학교")
    wb.save(path)
    return results

def write_perf(path, cls, rows):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "sheet1"
    ws["M1"] = "2026.11.12."
    ws["F2"] = "수행평가 강의실별 일람표"
    ws["B3"] = f"{YEAR}학년도   {SEM}학기   주간   {GRADE}학년   {cls} 강의실"
    ws["B4"] = f"교과목 : {SUBJECT}"; ws["J4"] = "교과담당교사 (홍길동,김철수) 인"
    ws.cell(6, 2, "반/번호"); ws.cell(6, 4, "성명")
    cols = [5, 7, 8]
    for (nm, mx, ratio), c in zip(AREAS, cols):
        ws.cell(6, c, f"{nm}\n(만점 {mx:.2f},\n{ratio:.2f}%)")
    ws.cell(6, 11, "합계"); ws.cell(6, 15, " 비고")
    r = 7
    perf = {}
    for num, sid, nm, ability, status in rows:
        ws.cell(r, 2, f"{cls}/{num}"); ws.cell(r, 3, float(sid)); ws.cell(r, 4, nm)
        if status == "전체면제":
            for c in cols: ws.cell(r, c, "면제")
            r += 1; continue
        scores = []
        for k, ((anm, mx, ratio), c) in enumerate(zip(AREAS, cols)):
            if status == "수행일부면제" and k == 1:
                ws.cell(r, c, "면제"); continue
            s = min(mx, max(mx * 0.2, round(random.gauss(mx * (0.55 + ability * 0.5), mx * 0.12))))
            if ability > 0.75 and random.random() < 0.6: s = mx
            ws.cell(r, c, float(s)); scores.append(s)
        ws.cell(r, 11, float(sum(scores)))
        perf[num] = sum(scores)
        r += 1
    ws.cell(r, 2, "응시생수")
    for c in cols: ws.cell(r, c, f"{len(perf)} 명")
    ws.cell(r + 1, 2, "총점"); ws.cell(r + 2, 2, "평균")
    ws.cell(r + 3, 2, "학년응시생수"); ws.cell(r + 4, 2, "학년총점"); ws.cell(r + 5, 2, "학년평균")
    ws.cell(r + 7, 9, 1.0); ws.cell(r + 7, 12, "/"); ws.cell(r + 7, 14, 1.0); ws.cell(r + 7, 16, "예시중학교")
    wb.save(path)
    return perf

def write_table(path, exam_name, classes, results_by_class):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "sheet1"
    ws["L1"] = "2026.11.12."
    ws["F2"] = "정기시험 교과목별 일람표"
    ws["B4"] = f"{YEAR}학년도   {SEM}학기   주간 {GRADE}학년"
    ws["B6"] = f"고사 : {exam_name}   교과목 : {SUBJECT}   과목만점 : 100.00"; ws["L6"] = "교과담당교사 : 홍길동, 김철수"
    ws.cell(8, 2, "반\n번호")
    cls_list = sorted(classes)
    colmap = {}
    c = 3
    for cls in cls_list:
        ws.cell(8, c, float(cls)); colmap[cls] = c
        c += 2 if cls % 3 == 0 else 1          # 일부 열은 병합된 것처럼 한 칸 띄움
    maxnum = max(num for rows in classes.values() for num, *_ in rows)
    for num in range(1, maxnum + 1):
        r = 8 + num
        ws.cell(r, 2, float(num))
        for cls in cls_list:
            status = {row[0]: row[4] for row in classes[cls]}.get(num)
            if status is None: continue
            if status in ("전체면제",):
                ws.cell(r, colmap[cls], "면제")
            elif num in results_by_class[cls]:
                ws.cell(r, colmap[cls], float(results_by_class[cls][num]))
    r = 9 + maxnum
    ws.cell(r, 2, "응시생수"); ws.cell(r + 1, 2, "총      점"); ws.cell(r + 2, 2, "평      균")
    for cls in cls_list:
        vals = list(results_by_class[cls].values())
        ws.cell(r, colmap[cls], f"{len(vals)} 명")
        ws.cell(r + 1, colmap[cls], float(sum(vals)))
        ws.cell(r + 2, colmap[cls], round(sum(vals) / len(vals), 1) if vals else 0.0)
    ws.cell(r + 4, 10, 1.0); ws.cell(r + 4, 12, "/"); ws.cell(r + 4, 13, 1.0); ws.cell(r + 4, 16, "예시중학교")
    wb.save(path)

def main():
    out = sys.argv[1] if len(sys.argv) > 1 else "sample"
    num_classes = int(sys.argv[2]) if len(sys.argv) > 2 and sys.argv[2].isdigit() else 6
    midterm = "--midterm" in sys.argv
    os.makedirs(out, exist_ok=True)
    classes = make_students(num_classes)
    seq = 5432900000000000
    def nxt():
        nonlocal seq
        seq += random.randint(10000000, 99999999)
        return os.path.join(out, f"download-{seq}.xlsx")

    exams = [("3학년 기말고사", 26, 0)]
    if midterm:
        exams.insert(0, ("3학년 중간고사", 25, 10))   # 중간고사는 25문항 + 서답형 10점
    for exam_name, n_items, essay in exams:
        correct, pts, diff = make_exam(n_items, essay)
        results = {}
        for cls, rows in classes.items():
            results[cls] = write_answer_sheet(nxt(), cls, rows, exam_name, correct, pts, diff, essay)
        write_table(nxt(), exam_name, classes, results)
    for cls, rows in classes.items():
        write_perf(nxt(), cls, rows)
    print(f"만든 폴더: {out}  (반 {num_classes}개, 고사 {len(exams)}개, 파일 {len(os.listdir(out))}개)")

if __name__ == "__main__":
    main()
