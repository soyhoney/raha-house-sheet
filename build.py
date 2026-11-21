"""라하 하우스 예약관리 시트(xlsx) 생성 스크립트

사용법:
    python build.py 라하하우스_예약관리_2027.xlsx            # 빈 시트 (예시 1줄)
    python build.py 라하하우스_예약관리_2027_예시데이터.xlsx sample   # 시연용 예시 10건
    python build.py test.xlsx test 20                         # 검증용 테스트 데이터 (2026년 10월)

필요: openpyxl
"""
import sys, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.utils import get_column_letter as CL
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment

OUT = sys.argv[1]
TEST = len(sys.argv) > 2 and sys.argv[2] == "test"

FONT = "Arial"
def F(**k):
    k.setdefault("name", FONT); k.setdefault("size", 10)
    return Font(**k)
fill = lambda c: PatternFill("solid", start_color=c, end_color=c)
thin = Side(style="thin", color="D0D0D0")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WON = '#,##0"원";-#,##0"원";""'
PCT = '0%;-0%;""'
DATE = "yyyy-mm-dd"
INPUT_FILL = fill("FFF8E1")   # 입력칸
AUTO_FILL = fill("EEF3F8")    # 자동칸
HEAD_FILL = fill("37474F")
HEAD_FONT = F(bold=True, color="FFFFFF")

FIRST = 5; LAST = int(sys.argv[3]) if len(sys.argv) > 3 else 1004
def R(col):  # 예약 탭 열 범위
    return f"'예약'!${col}${FIRST}:${col}${LAST}"

wb = Workbook()
# ------------------------------------------------------------------ 설정
st = wb.active; st.title = "설정"
st["A1"] = "설정 — 요금·정원·할인 기본값이 바뀌면 여기만 고치세요"; st["A1"].font = F(bold=True, size=14)
rooms = [("2호점 2인실", "2호점", 30000, 2), ("2호점 3인실", "2호점", 25000, 3),
         ("3호점 2인실 A", "3호점", 30000, 2), ("3호점 2인실 B", "3호점", 30000, 2),
         ("4호점 1인실", "4호점", 40000, 1), ("4호점 2인실 B", "4호점", 30000, 2)]
for i, h in enumerate(["객실", "지점", "지금 요금\n(자동)", "정원(명)"]):
    c = st.cell(4, 1 + i, h); c.font = HEAD_FONT; c.fill = HEAD_FILL; c.alignment = CENTER
st.row_dimensions[4].height = 30
# 요금표 (기간별, 체크인 날짜 기준)
PT0, PT1 = 22, 61   # 요금표 데이터 행
PA = f"$A${PT0}:$A${PT1}"; PB = f"$B${PT0}:$G${PT1}"; PH = "$B$21:$G$21"
def rate_at(date_expr, room_expr, sheet=""):
    """date_expr 날짜에 적용되는 room 요금 (적용 시작일 <= 날짜 인 가장 최근 줄)"""
    a = sheet + PA
    start = f"SUMPRODUCT(MAX(({a}<={date_expr})*{a}))"
    return f"INDEX({sheet}{PB},MATCH({start},{a},0),MATCH({room_expr},{sheet}{PH},0))"
for r, row in enumerate(rooms, start=5):
    for j, v in enumerate(row):
        if j == 2: continue
        c = st.cell(r, 1 + j, v); c.fill = INPUT_FILL; c.border = BOX; c.font = F(color="0000FF")
    st.cell(r, 3, f'=IFERROR({rate_at(f"MAX(TODAY(),MIN({PA}))", f"A{r}")},"")')
    st.cell(r, 3).number_format = WON; st.cell(r, 3).fill = AUTO_FILL; st.cell(r, 3).border = BOX
st["A12"] = "※ 객실 6개 고정(달력·통계가 5~10행을 사용). '지금 요금'은 아래 요금표에서 오늘 날짜 기준으로 자동 표시"
st["A12"].font = F(italic=True, color="777777", size=9)

st["A19"] = "요금표 (1인 1박) — 요금이 바뀌면 새 줄을 추가하세요. 예전 줄은 지우지 마세요(지난 예약 금액 유지)"
st["A19"].font = F(bold=True, size=11)
st["A20"] = ("✔ 체크인 날짜 기준: 체크인 날 적용되는 요금으로 숙박 전체를 계산합니다. "
             "요금이 바뀌는 날에 걸쳐 묵어도 체크인 날 요금이 끝까지 유지돼요. (예: 6/25~7/5 → 10박 모두 6월 요금)")
st["A20"].font = F(color="1B5E20", size=9, bold=True)
st["A21"] = "적용 시작일"
for i in range(6):
    st.cell(21, 2 + i, f"=A{5+i}")
for ci in range(1, 8):
    c = st.cell(21, ci); c.font = HEAD_FONT; c.fill = HEAD_FILL; c.alignment = CENTER; c.border = BOX
st.row_dimensions[21].height = 30
price_rows = [(dt.date(2026 if TEST else 2027, 1, 1), [x[2] for x in rooms])]
if len(sys.argv) > 2 and sys.argv[2] == "sample":
    price_rows.append((dt.date(2027, 2, 1), [32000, 27000, 32000, 32000, 42000, 32000]))
for r in range(PT0, PT1 + 1):
    for ci in range(1, 8):
        c = st.cell(r, ci); c.fill = INPUT_FILL; c.border = BOX; c.font = F(color="0000FF")
        c.number_format = DATE if ci == 1 else WON
for k, (d0, vals) in enumerate(price_rows):
    st.cell(PT0 + k, 1, d0)
    for j, v in enumerate(vals): st.cell(PT0 + k, 2 + j, v)
st.cell(PT0, 9, "← 첫 줄: 안내문 요금(1인 1박)").font = F(italic=True, color="777777", size=9)
if len(price_rows) > 1:
    st.cell(PT0 + 1, 9, "← 예시: 2월부터 인상 (1월 체크인 손님은 2월에도 1월 요금)").font = F(italic=True, color="777777", size=9)

st["F4"], st["G4"] = "할인 기본값", "값"
st["F5"], st["G5"] = "다이빙 연계 할인 (교육 1일당)", 3000
st["F6"], st["G6"] = "장기 할인 기본율 (빈칸일 때)", 0.10
st["F7"], st["G7"] = "장기 할인 기준 (박 이상)", 10
st["F8"], st["G8"] = "재방문 할인 기본율 (빈칸일 때)", 0.20
st["G5"].number_format = WON; st["G6"].number_format = PCT; st["G8"].number_format = PCT
st["F9"] = "계산: (기본금액 − 다이빙) × (1 − 장기% − 재방문%) − 기타조정   ※ 할인율은 합산"
st["F9"].font = F(italic=True, color="777777", size=9)

st["F11"], st["G11"], st["H11"] = "환불 규칙 (체크인 기준)", "며칠 전 이상", "환불율"
for r, (d, g, h) in enumerate([("15일 전까지", 15, 1.0), ("7~14일 전", 7, 0.6), ("6일 전부터", 0, 0)], start=12):
    st.cell(r, 6, d); st.cell(r, 7, g); st.cell(r, 8, h).number_format = '0%'
for c in ["F4", "G4", "F11", "G11", "H11"]:
    st[c].font = HEAD_FONT; st[c].fill = HEAD_FILL; st[c].alignment = CENTER
for r in list(range(5, 9)) + list(range(12, 15)):
    for col in "GH":
        if col == "H" and r < 12: continue
        st[f"{col}{r}"].fill = INPUT_FILL; st[f"{col}{r}"].font = F(color="0000FF"); st[f"{col}{r}"].border = BOX
    st[f"F{r}"].border = BOX

lists = {"J": ("상태", ["입금대기", "확정", "취소"]),
         "K": ("성별", ["남", "여"]),
         "L": ("예약경로", ["인스타", "카톡", "지인소개", "다이빙샵소개", "재방문", "기타"])}
for col, (h, vals) in lists.items():
    st[f"{col}4"] = h; st[f"{col}4"].font = HEAD_FONT; st[f"{col}4"].fill = HEAD_FILL; st[f"{col}4"].alignment = CENTER
    for i, v in enumerate(vals):
        st[f"{col}{5+i}"] = v
    for r in range(5, 16):
        st[f"{col}{r}"].border = BOX; st[f"{col}{r}"].fill = INPUT_FILL
st["L16"] = "↑ 경로 추가는 빈칸에"; st["L16"].font = F(italic=True, color="777777", size=9)
for col, w in zip("ABCDEFGHIJKL", [16, 12, 13, 12, 12, 32, 13, 9, 2, 10, 7, 14]):
    st.column_dimensions[col].width = w

# ------------------------------------------------------------------ 예약
ws = wb.create_sheet("예약", 0)
ws["A1"] = "라하 하우스 예약 — 예약 1건 = 1줄 (여러 달에 걸쳐도 1줄)"; ws["A1"].font = F(bold=True, size=14)
ws["A2"] = ("노란 칸만 입력하세요. 파란 칸은 자동 계산.  ·  할인 칸은 목록에서 O 선택  ·  "
            "할인율을 비워두면 기본값(장기 10%, 재방문 20%)  ·  1박 요금은 설정 탭 요금표에서 체크인 날짜 기준  ·  확정된 예약만 매출에 잡히고, 입금대기도 달력·정원 확인에는 포함됩니다.")
ws["A2"].font = F(color="555555", size=9)
heads = [("상태", 9), ("이름", 9), ("성별", 5), ("연락처", 14), ("예약경로", 11), ("객실", 13),
         ("체크인", 11), ("체크아웃", 11), ("다이빙\n연계", 6), ("교육\n일수", 6), ("장기", 5), ("장기 %\n(빈칸=10%)", 9),
         ("재방문", 6), ("재방문 %\n(빈칸=20%)", 9), ("기타 조정\n(원, 빼는 금액)", 11), ("메모", 18),
         ("입금일", 11), ("입금액", 11), ("취소일", 11),
         ("박수", 5), ("1박 요금\n(체크인일 기준)", 10), ("기본금액", 11), ("다이빙\n할인", 9), ("장기·재방문\n할인", 11),
         ("최종금액", 11), ("미수금", 10), ("환불액", 10), ("확인 필요", 22),
         ("1박당 기본", 9), ("1박당 다이빙", 9), ("1박당 %할인", 9), ("1박당 기타", 9), ("1박당 최종", 9)]
NIN = 19  # A~S 입력
ws.merge_cells("A3:S3"); ws["A3"] = "▼ 입력"; ws.merge_cells("T3:AB3"); ws["T3"] = "▼ 자동 계산"
ws.merge_cells("AC3:AG3"); ws["AC3"] = "▼ 통계용 (수정 금지)"
for c, f_ in (("A3", "F9A825"), ("T3", "5C7A99"), ("AC3", "9E9E9E")):
    ws[c].fill = fill(f_); ws[c].font = HEAD_FONT; ws[c].alignment = CENTER
for i, (h, w) in enumerate(heads, start=1):
    c = ws.cell(4, i, h); c.font = HEAD_FONT; c.fill = HEAD_FILL; c.alignment = CENTER; c.border = BOX
    ws.column_dimensions[CL(i)].width = w
ws.row_dimensions[4].height = 42
for col in ["AC", "AD", "AE", "AF", "AG"]:
    ws.column_dimensions[col].hidden = True

S = "'설정'!"
for r in range(FIRST, LAST + 1):
    f = {}
    f["T"] = f'=IF(OR(G{r}="",H{r}=""),"",H{r}-G{r})'
    f["U"] = f'=IF(OR(F{r}="",G{r}=""),"",IFERROR({rate_at(f"G{r}", f"F{r}", S)},""))'
    f["V"] = f'=IF(OR(T{r}="",U{r}=""),"",T{r}*U{r})'
    f["W"] = f'=IF(V{r}="","",IF(I{r}="O",MIN(N(J{r}),T{r})*{S}$G$5,0))'
    f["X"] = (f'=IF(V{r}="","",(V{r}-W{r})*(IF(K{r}="O",IF(L{r}="",{S}$G$6,L{r}),0)'
              f'+IF(M{r}="O",IF(N{r}="",{S}$G$8,N{r}),0)))')
    f["Y"] = f'=IF(V{r}="","",V{r}-W{r}-X{r}-N(O{r}))'
    f["Z"] = f'=IF(OR(Y{r}="",A{r}="취소",A{r}=""),"",Y{r}-N(R{r}))'
    f["AA"] = (f'=IF(AND(A{r}="취소",S{r}<>"",G{r}<>""),N(R{r})*IF(G{r}-S{r}>={S}$G$12,{S}$H$12,'
               f'IF(G{r}-S{r}>={S}$G$13,{S}$H$13,{S}$H$14)),0)')
    occb = (f'$F${FIRST}:$F${LAST},F{r},$G${FIRST}:$G${LAST},"<="&G{r},$H${FIRST}:$H${LAST},">"&G{r},$A${FIRST}:$A${LAST},')
    occ = lambda extra='': f'(COUNTIFS({occb}"확정"{extra})+COUNTIFS({occb}"입금대기"{extra}))'
    opp = occ(f',$C${FIRST}:$C${LAST},IF(C{r}="남","여","남")')
    live = f'AND(A{r}<>"",A{r}<>"취소",F{r}<>"",G{r}<>"")'
    f["AB"] = (f'=IF(B{r}="","",_xlfn.TEXTJOIN(", ",TRUE,'
               f'IF(AND(G{r}<>"",H{r}<>"",H{r}<=G{r}),"체크아웃 날짜 오류",""),'
               f'IF(OR(A{r}="",C{r}="",F{r}="",G{r}="",H{r}=""),"필수칸 비어있음",""),'
               f'IF({live},IF({occ()}>IFERROR(INDEX({S}$D$5:$D$10,MATCH(F{r},{S}$A$5:$A$10,0)),99),"정원 초과",""),""),'
               f'IF(AND({live},C{r}<>""),IF({opp}>0,"혼성 배정",""),""),'
               f'IF(AND(K{r}="O",N(T{r})<{S}$G$7),"10박 미만 장기할인",""),'
               f'IF(AND(I{r}="O",N(J{r})=0),"교육일수 없음",""),'
               f'IF(AND(F{r}<>"",G{r}<>"",N(U{r})=0),"요금표에 요금 없음",""),'
               f'IF(AND(A{r}="확정",N(Z{r})>0,G{r}<>"",G{r}<=TODAY()),"미입금",""),'
               f'IF(AND(A{r}="취소",S{r}=""),"취소일 없음","")))')
    f["AC"] = f'=IFERROR(V{r}/T{r},0)'
    f["AD"] = f'=IFERROR(W{r}/T{r},0)'
    f["AE"] = f'=IFERROR(X{r}/T{r},0)'
    f["AF"] = f'=IFERROR(N(O{r})/T{r},0)'
    f["AG"] = f'=IFERROR(Y{r}/T{r},0)'
    for col, v in f.items():
        ws[f"{col}{r}"] = v
    for i in range(1, 34):
        c = ws.cell(r, i); c.font = F(); c.border = BOX
        c.fill = INPUT_FILL if i <= NIN else AUTO_FILL
        if i <= NIN: c.font = F(color="0000FF")
    for col in "GHQS": ws[f"{col}{r}"].number_format = DATE
    for col in ["O", "R", "U", "V", "W", "X", "Y", "Z", "AA"]: ws[f"{col}{r}"].number_format = WON
    for col in "LN": ws[f"{col}{r}"].number_format = PCT
    for col in "ACEIJKMT": ws[f"{col}{r}"].alignment = Alignment(horizontal="center")
    ws[f"AB{r}"].font = F(color="C62828", bold=True)

def dv(formula, rng, **k):
    d = DataValidation(type="list", formula1=formula, allow_blank=True, **k); ws.add_data_validation(d)
    for x in rng.split(): d.add(x)
dv(f"{S}$J$5:$J$7", f"A{FIRST}:A{LAST}")
dv(f"{S}$K$5:$K$6", f"C{FIRST}:C{LAST}")
dv(f"{S}$L$5:$L$15", f"E{FIRST}:E{LAST}")
dv(f"{S}$A$5:$A$10", f"F{FIRST}:F{LAST}")
dv('"O"', f"I{FIRST}:I{LAST} K{FIRST}:K{LAST} M{FIRST}:M{LAST}")
dd = DataValidation(type="date", operator="greaterThan", formula1="43831", allow_blank=True,
                    error="날짜로 입력하세요 (예: 2027-01-05)", showErrorMessage=True)
ws.add_data_validation(dd)
for x in [f"G{FIRST}:H{LAST}", f"Q{FIRST}:Q{LAST}", f"S{FIRST}:S{LAST}"]: dd.add(x)
dp = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True,
                    error="퍼센트로 입력하세요 (예: 20%)", showErrorMessage=True)
ws.add_data_validation(dp)
for x in [f"L{FIRST}:L{LAST}", f"N{FIRST}:N{LAST}"]: dp.add(x)

rng = f"A{FIRST}:AB{LAST}"
ws.conditional_formatting.add(rng, FormulaRule(formula=[f'$A{FIRST}="취소"'], font=Font(color="9E9E9E", strike=True), fill=fill("F2F2F2")))
ws.conditional_formatting.add(f"A{FIRST}:A{LAST}", FormulaRule(formula=[f'$A{FIRST}="입금대기"'], fill=fill("FFE0B2")))
ws.conditional_formatting.add(f"A{FIRST}:A{LAST}", FormulaRule(formula=[f'$A{FIRST}="확정"'], fill=fill("C8E6C9")))
ws.conditional_formatting.add(f"AB{FIRST}:AB{LAST}", FormulaRule(formula=[f'LEN($AB{FIRST})>0'], fill=fill("FFEBEE")))
ws.freeze_panes = "C5"
ws.auto_filter.ref = f"A4:AB{LAST}"

# 예시 1줄 (지우고 쓰세요)
example = ["확정", "예시손님", "여", "카톡 abc123", "인스타", "2호점 3인실", dt.date(2027, 1, 3), dt.date(2027, 1, 15),
           "O", 4, "O", None, None, None, None, "예시 줄 — 확인 후 지우세요", dt.date(2026, 12, 20), 259200, None]
tests = [
    # (설명) 이름, 성별, 객실, 체크인, 체크아웃, 다이빙, 교육일, 장기, 장기%, 재방문, 재방문%, 기타, 상태, 입금액, 취소일, 경로
    ["확정", "임하은", "여", "", "재방문", "2호점 3인실", dt.date(2026, 9, 30), dt.date(2026, 12, 2), None, None, "O", 0.2, None, None, None, "63박 20%", None, 1260000, None],
    ["확정", "노지아", "여", "", "인스타", "2호점 3인실", dt.date(2026, 10, 1), dt.date(2026, 10, 14), None, None, "O", None, None, None, None, "", None, 0, None],
    ["확정", "송유나", "여", "", "카톡", "2호점 2인실", dt.date(2026, 10, 15), dt.date(2026, 10, 27), None, None, None, None, "O", None, None, "", None, 288000, None],
    ["확정", "권태양", "남", "", "카톡", "2호점 2인실", dt.date(2026, 10, 27), dt.date(2026, 10, 31), None, None, None, None, None, None, None, "혼성 아님(날짜 안 겹침)", None, 0, None],
    ["확정", "문건우", "남", "", "카톡", "2호점 2인실", dt.date(2026, 10, 20), dt.date(2026, 10, 25), None, None, None, None, None, None, None, "혼성 테스트", None, 0, None],
    ["입금대기", "테스트A", "여", "", "기타", "2호점 3인실", dt.date(2026, 10, 5), dt.date(2026, 10, 7), "O", 1, None, None, None, None, None, "정원초과 테스트", None, 0, None],
    ["확정", "테스트B", "여", "", "기타", "2호점 3인실", dt.date(2026, 10, 6), dt.date(2026, 10, 8), None, None, None, None, None, None, None, "4번째 손님", None, 50000, None],
    ["확정", "다이버", "남", "", "다이빙샵소개", "3호점 2인실 B", dt.date(2026, 10, 3), dt.date(2026, 10, 15), "O", 4, "O", None, None, None, 5000, "A안 예시", None, 0, None],
    ["취소", "취소손님", "여", "", "인스타", "4호점 1인실", dt.date(2026, 10, 20), dt.date(2026, 10, 25), None, None, None, None, None, None, None, "10일 전 취소", None, 200000, dt.date(2026, 10, 10)],
]
D = dt.date
samples = [
    ["확정", "박지민", "여", "카톡 jimin_p", "인스타", "2호점 3인실", D(2027,1,2), D(2027,1,14), "O", 5, "O", None, None, None, None, "오픈워터+어드밴스", D(2026,12,20), 256500, None],
    ["확정", "이수진", "여", "010-1234-5678", "카톡", "2호점 3인실", D(2027,1,5), D(2027,1,12), None, None, None, None, "O", None, None, "작년 3월 방문", D(2026,12,22), 140000, None],
    ["확정", "최민호", "남", "카톡 minho_c", "다이빙샵소개", "3호점 2인실 B", D(2027,1,3), D(2027,1,10), "O", 4, None, None, None, None, None, "프리다이빙 4일", D(2026,12,18), 198000, None],
    ["확정", "한도윤", "남", "010-2222-3333", "지인소개", "3호점 2인실 B", D(2027,1,4), D(2027,1,9), None, None, None, None, None, None, None, "잔금 체크인 때", D(2026,12,26), 100000, None],
    ["확정", "오하늘", "여", "인스타 @haneul.o", "인스타", "2호점 2인실", D(2027,1,10), D(2027,2,20), None, None, "O", 0.2, None, None, None, "41박 장기 — 할인율 20%로 조정", D(2026,12,28), 984000, None],
    ["입금대기", "강서준", "남", "카톡 seojun_k", "카톡", "4호점 1인실", D(2027,1,15), D(2027,1,20), None, None, None, None, None, None, None, "입금 안내 보냄", None, None, None],
    ["확정", "윤가은", "여", "010-4444-5555", "재방문", "3호점 2인실 A", D(2027,1,8), D(2027,1,22), None, None, "O", None, "O", None, 4000, "장기+재방문 중복, 끝자리 정리", D(2026,12,30), 290000, None],
    ["취소", "문지후", "남", "카톡 jihu_m", "인스타", "4호점 2인실 B", D(2027,1,18), D(2027,1,25), None, None, None, None, None, None, None, "체크인 10일 전 취소", D(2026,12,1), 210000, D(2027,1,8)],
    ["확정", "서예린", "여", "인스타 @yerin.s", "다이빙샵소개", "2호점 3인실", D(2027,1,10), D(2027,1,16), "O", 6, None, None, None, None, None, "", D(2027,1,2), 132000, None],
    ["확정", "배준혁", "남", "010-6666-7777", "카톡", "2호점 2인실", D(2027,2,1), D(2027,2,8), None, None, None, None, None, None, None, "혼성 경고 예시 — 방 옮겨야 함 · 2월 인상 요금 적용", D(2027,1,20), 224000, None],
]
SAMPLE = len(sys.argv) > 2 and sys.argv[2] == "sample"
rows = samples if SAMPLE else [example] + (tests if TEST else [])
for i, row in enumerate(rows):
    for j, v in enumerate(row):
        if v is not None and v != "":
            ws.cell(FIRST + i, 1 + j, v)

# ------------------------------------------------------------------ 달력
cal = wb.create_sheet("달력", 1)
cal["A1"] = "예약 달력"; cal["A1"].font = F(bold=True, size=14)
cal["A2"], cal["B2"], cal["C2"], cal["D2"] = "연도", 2026 if TEST else 2027, "월", 10 if TEST else 1
for c in ("B2", "D2"): cal[c].fill = INPUT_FILL; cal[c].font = F(bold=True, color="0000FF", size=12); cal[c].border = BOX; cal[c].alignment = CENTER
cdv = DataValidation(type="list", formula1='"1,2,3,4,5,6,7,8,9,10,11,12"'); cal.add_data_validation(cdv); cdv.add("D2")
cal["F2"] = "⬜ 빈방   🟨 일부   🟩 만석   🟥 정원 초과   🟪 혼성   ·  (·대기) = 입금대기"; cal["F2"].font = F(size=9, color="555555")
cal["A4"], cal["B4"] = "객실", "정원"
cal["A5"] = ""
DC = list(range(3, 34))  # C~AG = 1~31일
for k, ci in enumerate(DC, start=1):
    col = CL(ci)
    cal[f"{col}4"] = f'=IF(MONTH(DATE($B$2,$D$2,{k}))=$D$2,DATE($B$2,$D$2,{k}),"")'
    cal[f"{col}4"].number_format = "d"
    cal[f"{col}5"] = f'=IF({col}4="","",CHOOSE(WEEKDAY({col}4),"일","월","화","수","목","금","토"))'
    cal.column_dimensions[col].width = 9.5
for c in ["A4", "B4"] + [f"{CL(ci)}4" for ci in DC]:
    cal[c].font = HEAD_FONT; cal[c].fill = HEAD_FILL; cal[c].alignment = CENTER
for ci in DC:
    cal[f"{CL(ci)}5"].alignment = CENTER; cal[f"{CL(ci)}5"].font = F(size=9, color="777777")
cal.column_dimensions["A"].width = 14; cal.column_dimensions["B"].width = 5

CNT0, MIX0 = 27, 35  # 계산용 행
for i in range(6):
    r = 6 + i; rc = CNT0 + i; rm = MIX0 + i
    cal[f"A{r}"] = f"={S}A{5+i}"; cal[f"B{r}"] = f"={S}D{5+i}"
    cal[f"A{rc}"] = f"=A{r}"; cal[f"B{rc}"] = f"={S}B{5+i}"
    cal[f"A{rm}"] = f"=A{r}"
    cal.row_dimensions[r].height = 46
    for c in (f"A{r}", f"B{r}"):
        cal[c].font = F(bold=True); cal[c].alignment = CENTER; cal[c].border = BOX
    for ci in DC:
        col = CL(ci); d = f"{col}$4"
        cond = (f"({R('F')}=$A{r})*({R('G')}<={d})*({R('H')}>{d})*({R('A')}<>\"취소\")*({R('A')}<>\"\")")
        label = f'{R("B")}&"("&{R("C")}&IF({R("A")}="입금대기","·대기","")&")"'
        text = f'=IF({d}="","",_xlfn.TEXTJOIN(CHAR(10),TRUE,IF({cond},{label},"")))'
        cal[f"{col}{r}"] = ArrayFormula(f"{col}{r}", text)
        cal[f"{col}{r}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cal[f"{col}{r}"].font = F(size=9); cal[f"{col}{r}"].border = BOX
        base = (f"{R('F')},$A{rc},{R('G')},\"<=\"&{d},{R('H')},\">\"&{d},{R('A')},")
        cnt = lambda extra='': f'(COUNTIFS({base}"확정"{extra})+COUNTIFS({base}"입금대기"{extra}))'
        cal[f"{col}{rc}"] = f'=IF({d}="","",{cnt()})'
        cm = cnt("," + R("C") + ',"남"'); cw = cnt("," + R("C") + ',"여"')
        cal[f"{col}{rm}"] = f'=IF({d}="","",IF(AND({cm}>0,{cw}>0),1,0))'

cal["A13"] = "남은 침대"; cal["A13"].font = F(bold=True)
# 객실별 → 지점 소계 → 전체 (14~23행)
layout = [("room", 0), ("room", 1), ("sub", "2호점"), ("room", 2), ("room", 3), ("sub", "3호점"),
          ("room", 4), ("room", 5), ("sub", "4호점"), ("all", None)]
sub_rows = []
for k, (kind, v) in enumerate(layout):
    r = 14 + k
    if kind == "room":
        cal[f"A{r}"] = f"=A{6+v}"; cal[f"B{r}"] = f"=B{6+v}"
        cal[f"A{r}"].font = F(); cal[f"B{r}"].font = F(color="777777")
    elif kind == "sub":
        cal[f"A{r}"] = f"{v} 합계"; sub_rows.append(r)
    else:
        cal[f"A{r}"] = "전체"
    for c in (f"A{r}", f"B{r}"):
        cal[c].alignment = CENTER; cal[c].border = BOX
    if kind != "room":
        cal[f"A{r}"].font = F(bold=True)
        for ci in range(1, 34): cal.cell(r, ci).fill = fill("ECEFF1")
    for ci in DC:
        col = CL(ci)
        if kind == "room":
            cal[f"{col}{r}"] = f'=IF({col}$4="","",$B{r}-{col}{CNT0+v})'
        elif kind == "sub":
            cal[f"{col}{r}"] = f'=IF({col}$4="","",{col}{r-2}+{col}{r-1})'
        else:
            cal[f"{col}{r}"] = f'=IF({col}$4="","",' + "+".join(f"{col}{x}" for x in sub_rows) + ")"
        cal[f"{col}{r}"].alignment = CENTER; cal[f"{col}{r}"].border = BOX
        if kind != "room": cal[f"{col}{r}"].font = F(bold=True)
cal.conditional_formatting.add("C14:AG23", CellIsRule(operator="lessThan", formula=["0"], fill=fill("EF9A9A"), font=Font(bold=True, color="B71C1C")))
cal.conditional_formatting.add("C14:AG23", FormulaRule(formula=['AND(C14<>"",C14=0)'], font=Font(bold=True, color="B71C1C")))
cal.conditional_formatting.add("C4:AG5", FormulaRule(formula=['OR(C$5="토",C$5="일")'], font=Font(color="FFCDD2", bold=True)))

cal[f"A{CNT0-2}"] = "▼ 계산용 (수정 금지) — 객실별 인원 / 혼성 여부"; cal[f"A{CNT0-2}"].font = F(italic=True, color="999999")
for r in range(CNT0 - 2, MIX0 + 6):
    cal.row_dimensions[r].hidden = True

grid = "C6:AG11"
cal.conditional_formatting.add(grid, FormulaRule(formula=[f"C{MIX0}=1"], fill=fill("CE93D8"), stopIfTrue=True))
cal.conditional_formatting.add(grid, FormulaRule(formula=[f'AND(C$4<>"",N(C{CNT0})>$B6)'], fill=fill("EF9A9A"), stopIfTrue=True))
cal.conditional_formatting.add(grid, FormulaRule(formula=[f'AND(C$4<>"",N(C{CNT0})=$B6)'], fill=fill("A5D6A7"), stopIfTrue=True))
cal.conditional_formatting.add(grid, FormulaRule(formula=[f'AND(C$4<>"",N(C{CNT0})>0)'], fill=fill("FFF59D"), stopIfTrue=True))
cal.conditional_formatting.add(grid, FormulaRule(formula=['C$4=""'], fill=fill("EEEEEE"), stopIfTrue=True))
cal.freeze_panes = "C6"

# ------------------------------------------------------------------ 지출
st["F16"], st["G16"] = "기본 환율 (1 EGP = 원)", 28
st["F16"].border = BOX
st["G16"].fill = INPUT_FILL; st["G16"].font = F(color="0000FF"); st["G16"].border = BOX; st["G16"].number_format = '#,##0.0"원"'
st["F17"] = "※ 28원은 예시값 — 실제 환율로 고치세요. 지출 탭에서 달마다 따로 넣을 수도 있어요"
st["F17"].font = F(italic=True, color="777777", size=9)
RATE = f"{S}$G$16"

ex = wb.create_sheet("지출")
ex["A1"] = "운영 지출 (월별)"; ex["A1"].font = F(bold=True, size=14)
ex["A2"] = ("이집트 파운드(EGP)로 입력하세요. 환율 칸을 비우면 설정 탭의 기본 환율로 원화 환산 → "
            "통계 탭의 지출·순이익에 반영됩니다.")
ex["A2"].font = F(color="555555", size=9)
ITEMS = ["가스", "전기", "수도", "인터넷", "수리", "비품", "기타"]
eh = (["월"] + [f"{x}\n(EGP)" for x in ITEMS] + ["환율\n(원/EGP, 빈칸=기본)", "합계\n(EGP)", "합계\n(원)", "메모", "적용 환율"])
for j, h in enumerate(eh, start=1):
    c = ex.cell(4, j, h); c.font = HEAD_FONT; c.fill = HEAD_FILL; c.alignment = CENTER; c.border = BOX
ex.row_dimensions[4].height = 34
EX0, EX1 = 5, 200
start_year = 2026 if TEST else 2027
for k in range(36):
    r = EX0 + k
    y, m = start_year + k // 12, k % 12 + 1
    ex[f"A{r}"] = dt.date(y, m, 1); ex[f"A{r}"].number_format = "yyyy-mm"
    ex[f"A{r}"].font = F(bold=True); ex[f"A{r}"].alignment = CENTER
for r in range(EX0, EX0 + 36):
    for ci in range(2, 10):
        c = ex.cell(r, ci); c.fill = INPUT_FILL; c.font = F(color="0000FF"); c.number_format = '#,##0;-#,##0;""'
    ex[f"I{r}"].number_format = '#,##0.0;-#,##0.0;""'
    ex[f"J{r}"] = f'=IF(COUNT(B{r}:H{r})=0,"",SUM(B{r}:H{r}))'; ex[f"J{r}"].number_format = '#,##0'
    ex[f"K{r}"] = f'=IF(J{r}="","",J{r}*M{r})'; ex[f"K{r}"].number_format = WON
    ex[f"M{r}"] = f'=IF(I{r}="",{RATE},I{r})'
    ex[f"L{r}"].fill = INPUT_FILL; ex[f"L{r}"].font = F(color="0000FF")
    for col in "JK": ex[f"{col}{r}"].fill = AUTO_FILL; ex[f"{col}{r}"].font = F(bold=True)
    for ci in range(1, 13): ex.cell(r, ci).border = BOX
ex.column_dimensions["M"].hidden = True
for col, w in zip("ABCDEFGHIJKL", [10, 9, 9, 9, 9, 9, 9, 9, 13, 11, 13, 24]):
    ex.column_dimensions[col].width = w
ex.freeze_panes = "B5"
ex[f"A{EX0+37}"] = "↑ 더 필요하면 아래 줄에 월(예: 2030-01-01)을 넣고 이어서 쓰세요."
ex[f"A{EX0+37}"].font = F(italic=True, color="777777", size=9)
for r in range(EX0 + 36, EX1 + 1):  # 여분 줄도 계산되게
    ex[f"J{r}"] = f'=IF(COUNT(B{r}:H{r})=0,"",SUM(B{r}:H{r}))'
    ex[f"K{r}"] = f'=IF(J{r}="","",J{r}*M{r})'
    ex[f"M{r}"] = f'=IF(I{r}="",{RATE},I{r})'
    ex[f"K{r}"].number_format = WON
if SAMPLE:
    demo = {0: [350, 2600, 420, 600, 0, 850, 150, None, "1월 고지서"],
            1: [380, 2900, 450, 600, 1200, 300, 0, 29.5, "2월: 에어컨 수리, 환율 직접 입력"]}
    cols = [2, 3, 4, 5, 6, 7, 8, 9, 12]  # B~H 항목, I 환율, L 메모
    for k, vals in demo.items():
        for j, v in enumerate(vals):
            if v is not None: ex.cell(EX0 + k, cols[j], v)

def EX(col): return f"'지출'!${col}${EX0}:${col}${EX1}"

# ------------------------------------------------------------------ 통계
sx = wb.create_sheet("통계", 2)
sx["A1"] = "통계"; sx["A1"].font = F(bold=True, size=14)
sx["A2"] = ("매출은 '묵은 날' 기준으로 달마다 나눠 잡습니다 (확정 예약만). 지출은 지출 탭의 원화 합계. "
            "순이익 = 순매출(할인 뺀 매출) − 지출.")
sx["A2"].font = F(size=9, color="555555")
sx["A3"], sx["B3"], sx["C3"], sx["D3"] = "연도", 2026 if TEST else 2027, "월", 10 if TEST else 1
for c in ("B3", "D3"): sx[c].fill = INPUT_FILL; sx[c].font = F(bold=True, color="0000FF", size=12); sx[c].border = BOX; sx[c].alignment = CENTER
sdv = DataValidation(type="list", formula1='"1,2,3,4,5,6,7,8,9,10,11,12"'); sx.add_data_validation(sdv); sdv.add("D3")
sx["F3"] = "=DATE($B$3,$D$3,1)"; sx["G3"] = "=EDATE(F3,1)"
sx["F3"].number_format = DATE; sx["G3"].number_format = DATE
sx["E3"] = "기간"; sx["H3"] = "까지(미포함)"
for c in ("E3", "F3", "G3", "H3"): sx[c].font = F(size=9, color="999999")
MLABEL = '$B$3&"년 "&$D$3&"월"'

def ov(ms, me):
    G, H = R("G"), R("H")
    return (f"(({G}<{me})*({H}>{ms})*((({H}<{me})*{H}+({H}>={me})*{me})-((({G}>{ms})*{G}+({G}<={ms})*{ms}))))")
OK = f'({R("A")}="확정")'
def sp(*parts): return "SUMPRODUCT(" + "*".join(parts) + ")"
def exp_total(a, b): return f'SUMIFS({EX("K")},{EX("A")},">="&{a},{EX("A")},"<"&{b})'
def exp_item(col, a, b): return sp(f'({EX("A")}>={a})', f'({EX("A")}<{b})', EX(col), EX("M"))
def exp_egp(col, a, b): return sp(f'({EX("A")}>={a})', f'({EX("A")}<{b})', EX(col))

ms, me = "$F$3", "$G$3"
OV = ov(ms, me)
cap_all = f"SUM({S}$D$5:$D$10)"

def title(cell, text):
    sx[cell] = f'={MLABEL}&" {text}"'; sx[cell].font = F(bold=True, size=11)

title("A5", "요약")
cards = [
    ("순매출", "=" + sp(OK, OV, R("AG")), WON),
    ("지출", "=" + exp_total(ms, me), WON),
    ("순이익", "=A7-C7", WON),
    ("할인 전 매출", "=" + sp(OK, OV, R("AC")), WON),
    ("할인 합계", "=B10+B11+B12", WON),
    ("가동률", f"=IFERROR({sp(OK, OV)}/({cap_all}*({me}-{ms})),0)", "0.0%"),
    ("투숙 손님 수", f"={sp(OK, f'({OV}>0)')}", '0"명"'),
    ("평균 숙박일 (이 달 체크인)", f'=IFERROR(AVERAGEIFS({R("T")},{R("A")},"확정",{R("G")},">="&{ms},{R("G")},"<"&{me}),0)', '0.0"박"'),
]
for i, (lab, fm, nf) in enumerate(cards):
    col = CL(1 + i * 2); col2 = CL(2 + i * 2)
    sx.merge_cells(f"{col}6:{col2}6"); sx.merge_cells(f"{col}7:{col2}7")
    sx[f"{col}6"] = lab; sx[f"{col}6"].font = F(size=9, color="666666"); sx[f"{col}6"].alignment = CENTER
    sx[f"{col}7"] = fm; sx[f"{col}7"].number_format = nf
    sx[f"{col}7"].font = F(bold=True, size=14, color="1B5E20" if lab == "순이익" else "000000"); sx[f"{col}7"].alignment = CENTER
    for c in (f"{col}6", f"{col2}6", f"{col}7", f"{col2}7"):
        sx[c].fill = fill("E8F5E9" if lab == "순이익" else "F5F7FA"); sx[c].border = BOX
sx.conditional_formatting.add("E7", CellIsRule(operator="lessThan", formula=["0"], font=Font(bold=True, color="C62828")))
sx.row_dimensions[7].height = 30

def head(r, c0, labels):
    for j, h in enumerate(labels):
        c = sx.cell(r, c0 + j, h); c.font = HEAD_FONT; c.fill = HEAD_FILL; c.alignment = CENTER

head(9, 1, ["할인 종류별", "금액"])
rowsd = [("다이빙 연계", sp(OK, OV, R("AD"))), ("장기·재방문", sp(OK, OV, R("AE"))), ("기타 조정", sp(OK, OV, R("AF")))]
for i, (lab, fm) in enumerate(rowsd):
    sx[f"A{10+i}"] = lab; sx[f"B{10+i}"] = "=" + fm; sx[f"B{10+i}"].number_format = WON
    for col in "AB": sx[f"{col}{10+i}"].border = BOX
head(9, 4, ["돈 관리", "금액"])
money = [("입금대기 예상 매출 (이 달)", "=" + sp(f'({R("A")}="입금대기")', OV, R("AG"))),
         ("미수금 (확정 전체)", f'=SUMIFS({R("Z")},{R("A")},"확정")'),
         ("취소 위약금 (이 달 취소)", "=" + sp(f'({R("A")}="취소")', f'({R("S")}>={ms})', f'({R("S")}<{me})', f'({R("R")}-{R("AA")})'))]
for i, (lab, fm) in enumerate(money):
    sx[f"D{10+i}"] = lab; sx[f"E{10+i}"] = fm; sx[f"E{10+i}"].number_format = WON
    for col in "DE": sx[f"{col}{10+i}"].border = BOX

head(9, 7, ["지출 항목", "EGP", "원화"])
for i, it in enumerate(ITEMS):
    r = 10 + i; col = CL(2 + i)
    sx[f"G{r}"] = it
    sx[f"H{r}"] = "=" + exp_egp(col, ms, me); sx[f"H{r}"].number_format = '#,##0" E£";-#,##0;"-"'
    sx[f"I{r}"] = "=" + exp_item(col, ms, me); sx[f"I{r}"].number_format = WON
    for c in "GHI": sx[f"{c}{r}"].border = BOX
sx["G17"] = "합계"; sx["H17"] = "=SUM(H10:H16)"; sx["I17"] = "=SUM(I10:I16)"
sx["H17"].number_format = '#,##0" E£"'; sx["I17"].number_format = WON
for c in "GHI": sx[f"{c}17"].font = F(bold=True); sx[f"{c}17"].border = BOX

# 객실별
RB = 20
title(f"A{RB}", "객실별")
head(RB + 1, 1, ["객실", "정원", "숙박(박)", "가동률", "할인 전 매출", "할인", "순매출"])
for i in range(6):
    r = RB + 2 + i; RM = f'({R("F")}=$A{r})'
    sx[f"A{r}"] = f"={S}A{5+i}"; sx[f"B{r}"] = f"={S}D{5+i}"
    sx[f"C{r}"] = "=" + sp(OK, RM, OV)
    sx[f"D{r}"] = f"=IFERROR(C{r}/(B{r}*({me}-{ms})),0)"
    sx[f"E{r}"] = "=" + sp(OK, RM, OV, R("AC"))
    sx[f"F{r}"] = f"=E{r}-G{r}"
    sx[f"G{r}"] = "=" + sp(OK, RM, OV, R("AG"))
rt = RB + 8
sx[f"A{rt}"] = "합계"; sx[f"B{rt}"] = f"=SUM(B{RB+2}:B{rt-1})"; sx[f"C{rt}"] = f"=SUM(C{RB+2}:C{rt-1})"
sx[f"D{rt}"] = f"=IFERROR(C{rt}/(B{rt}*({me}-{ms})),0)"
for col in "EFG": sx[f"{col}{rt}"] = f"=SUM({col}{RB+2}:{col}{rt-1})"
for r in range(RB + 2, rt + 1):
    sx[f"D{r}"].number_format = "0.0%"
    for col in "EFG": sx[f"{col}{r}"].number_format = WON
    for col in "ABCDEFG": sx[f"{col}{r}"].border = BOX
for col in "ABCDEFG": sx[f"{col}{rt}"].font = F(bold=True)

# 예약경로별
title(f"I{RB}", "예약경로별")
head(RB + 1, 9, ["경로", "손님 수", "순매출"])
for i in range(11):
    r = RB + 2 + i; RT = f'({R("E")}=$I{r})'
    sx[f"I{r}"] = f'=IF({S}L{5+i}="","",{S}L{5+i})'
    sx[f"J{r}"] = f'=IF(I{r}="","",{sp(OK, RT, f"({OV}>0)")})'
    sx[f"K{r}"] = f'=IF(I{r}="","",{sp(OK, RT, OV, R("AG"))})'
    sx[f"J{r}"].number_format = '0"명"'; sx[f"K{r}"].number_format = WON
    for col in "IJK": sx[f"{col}{r}"].border = BOX

# 연간
YA = 35
sx[f"A{YA}"] = '=$B$3&"년 연간 추이"'; sx[f"A{YA}"].font = F(bold=True, size=11)
head(YA + 1, 1, ["월", "할인 전 매출", "할인", "순매출", "지출", "순이익", "가동률", "투숙 손님"])
for m in range(1, 13):
    r = YA + 1 + m
    a, b = f"DATE($B$3,{m},1)", f"EDATE(DATE($B$3,{m},1),1)"
    OVm = ov(a, b)
    sx[f"A{r}"] = f"{m}월"
    sx[f"B{r}"] = "=" + sp(OK, OVm, R("AC"))
    sx[f"C{r}"] = f"=B{r}-D{r}"
    sx[f"D{r}"] = "=" + sp(OK, OVm, R("AG"))
    sx[f"E{r}"] = "=" + exp_total(a, b)
    sx[f"F{r}"] = f"=D{r}-E{r}"
    sx[f"G{r}"] = f"=IFERROR({sp(OK, OVm)}/({cap_all}*({b}-{a})),0)"
    sx[f"H{r}"] = f"={sp(OK, f'({OVm}>0)')}"
    for col in "BCDEF": sx[f"{col}{r}"].number_format = WON
    sx[f"G{r}"].number_format = "0.0%"; sx[f"H{r}"].number_format = '0"명"'
    for col in "ABCDEFGH": sx[f"{col}{r}"].border = BOX
YT = YA + 14
sx[f"A{YT}"] = "합계"
for col in "BCDEF": sx[f"{col}{YT}"] = f"=SUM({col}{YA+2}:{col}{YT-1})"; sx[f"{col}{YT}"].number_format = WON
sx[f"G{YT}"] = f"=AVERAGE(G{YA+2}:G{YT-1})"; sx[f"G{YT}"].number_format = "0.0%"
for col in "ABCDEFGH": sx[f"{col}{YT}"].font = F(bold=True); sx[f"{col}{YT}"].border = BOX
sx[f"A{YT+1}"] = "※ 손님 수는 달마다 세므로 합계 없음 · 가동률 합계는 12개월 평균"; sx[f"A{YT+1}"].font = F(size=8, color="999999")
sx.conditional_formatting.add(f"F{YA+2}:F{YT}", CellIsRule(operator="lessThan", formula=["0"], font=Font(bold=True, color="C62828")))

ch = BarChart(); ch.type = "col"; ch.title = "월별 순매출 · 지출 · 순이익"; ch.style = 10; ch.grouping = "clustered"
ch.add_data(Reference(sx, min_col=4, max_col=6, min_row=YA + 1, max_row=YT - 1), titles_from_data=True)
ch.set_categories(Reference(sx, min_col=1, min_row=YA + 2, max_row=YT - 1))
ch.height, ch.width = 8.5, 20
ch.roundedCorners = False; ch.legend.position = "b"; ch.legend.overlay = False
ch.x_axis.delete = False; ch.y_axis.delete = False; ch.title.overlay = False
ch.y_axis.numFmt = '#,##0'
sx.add_chart(ch, f"J{YA}")
for col, w in zip("ABCDEFGHIJKLMNOP", [16, 14, 12, 26, 14, 14, 12, 12, 14, 10, 14, 12, 12, 12, 12, 12]):
    sx.column_dimensions[col].width = w

wb._sheets = [wb[n] for n in ["예약", "달력", "통계", "지출", "설정"]]


for sh in wb.worksheets:
    sh.sheet_view.showGridLines = False if sh.title != "예약" else True
    for row in sh.iter_rows():
        for c in row:
            if c.font and c.font.name != FONT:
                c.font = Font(name=FONT, size=c.font.size, bold=c.font.bold, italic=c.font.italic, color=c.font.color, strike=c.font.strike)
wb.save(OUT)
print("saved", OUT)
