"""產生「試跑用」假資料：把收集表範本複製一份，填入 71 班的假資料。

全部是虛構的老師代號（例：英語甲），不含任何真實姓名。
用法：python engine/make_demo_data.py  →  產生 試跑/試跑資料_排課需求收集表.xlsx
"""
import random
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "收集表" / "115_10_排課需求收集表.xlsx"
OUT = ROOT / "試跑" / "試跑資料_排課需求收集表.xlsx"

DAYS = ["一", "二", "三", "四", "五"]
GANZHI = "甲乙丙丁戊己庚辛壬癸"
CLASSES = {1: 12, 2: 11, 3: 12, 4: 12, 5: 12, 6: 12}

# 每個年級的科任課（科目, 每週節數, 地點, 連堂節數）
PLAN = {
    1: [("英語", 1, "原班教室", None), ("體育", 1, "操場／體育館", None), ("音樂", 1, "音樂教室", None)],
    3: [("英語", 2, "原班教室", None), ("自然", 3, "自然教室", 2), ("體育", 2, "操場／體育館", None),
        ("音樂", 2, "音樂教室", None), ("資訊", 1, "電腦教室", None), ("健康", 1, "原班教室", None)],
    5: [("英語", 2, "原班教室", None), ("自然", 3, "自然教室", 2), ("社會", 3, "原班教室", None),
        ("體育", 2, "操場／體育館", None), ("音樂", 2, "音樂教室", None), ("資訊", 1, "電腦教室", None),
        ("健康", 1, "原班教室", None)],
}
PLAN[2], PLAN[4], PLAN[6] = PLAN[1], PLAN[3], PLAN[5]
TARGET_LOAD = {"英語": 20, "自然": 21, "社會": 18, "體育": 20, "音樂": 20, "資訊": 24, "健康": 16}

# 假設的作息（真資料請 Bing 填）
COMMON = [
    ("全校", "三", "1", "綠色迪士尼", "導師", "是", None),
    ("三、四年級", "一", "5–7", "綜合", "導師", "是", "年段共同時段"),
    ("五、六年級", "四", "5–7", "綜合", "導師", "是", "年段共同時段"),
    ("全校", "三", "下午", "不上課", "—", "是", "（試跑假設）星期三下午全校不上課"),
] + [("一、二年級", d, "下午", "不上課", "—", "是", "（試跑假設）低年級只有星期二整天")
     for d in ["一", "三", "四", "五"]]


def class_name(g, i):
    return f"{g}年{i}班"


def class_slots(g):
    """回傳這個年級學生在校的 (星期, 節次)，扣掉共同鎖定時段。"""
    out = []
    for d in DAYS:
        for p in range(1, 8):
            if d == "三" and p >= 5:
                continue
            if g <= 2 and p >= 5 and d != "二":
                continue
            if d == "三" and p == 1:
                continue
            if g in (3, 4) and d == "一" and p >= 5:
                continue
            if g in (5, 6) and d == "四" and p >= 5:
                continue
            out.append((d, p))
    return out


NEW_RULES = [  # 2026-10-21 新增（R17 待 Bing 確認，可改「否」關閉）
    ["R17", "3 盡量", "同科分散", "所有科任", "同一班同一科一天最多 1 次（連堂算 1 次）", "高", "是", "避免同一天上兩次同科"],
    ["R18", "3 盡量", "間隔天數", "自然", "同一班自然各次上課至少隔 1 天（例：星期一上連堂，原班那節排星期三以後）", "中",
     "是", "Bing 2026-10-21：原班那節要不同天，最好差一天"],
]
R11_TEXT = "每班每週3節：2節連堂在自然教室（可跨大下課）＋1節在原班，原班那節排不同天"


def add_rules(wb):
    """規則表加上新規則、更新 R11 說明，並讓「規則積木」下拉選單多出新積木。"""
    ws = wb["規則表"]
    for r in range(4, ws.max_row + 1):
        if ws.cell(r, 1).value == "R11":
            ws.cell(r, 5, R11_TEXT)
    for k, row in enumerate(NEW_RULES):
        for j, v in enumerate(row):
            ws.cell(20 + k, 1 + j, v)
    opts = wb["選項"]
    for k, block in enumerate(["同科分散", "間隔天數"]):
        opts.cell(11 + k, 13, block)
    for dv in ws.data_validations.dataValidation:
        if "M$2:$M$10" in (dv.formula1 or ""):
            dv.formula1 = "選項!$M$2:$M$12"
        dv.sqref = openpyxl.worksheet.cell_range.MultiCellRange(str(dv.sqref).replace("19", "40"))


def main(seed=115, out=OUT):
    random.seed(seed)
    wb = openpyxl.load_workbook(TEMPLATE)

    # 1. 科任配課：每科依目標節數切給老師（同一位老師教連續的班，比較貼近實際）
    rows_by_subject = {}
    for g, n in CLASSES.items():
        for i in range(1, n + 1):
            for subj, hrs, loc, block in PLAN[g]:
                rows_by_subject.setdefault(subj, []).append([class_name(g, i), subj, hrs, loc, block])
    assign = []
    for subj, rows in rows_by_subject.items():
        k, load = 0, 0
        for r in rows:
            if load + r[2] > TARGET_LOAD[subj]:
                k, load = k + 1, 0
            load += r[2]
            note = "另1節在原班" if subj == "自然" else None
            assign.append([f"{subj}{GANZHI[k]}", subj, r[0], r[2], r[3], r[4], note])
    ws = wb["科任配課"]
    for i, r in enumerate(assign):
        for j, v in enumerate(r):
            ws.cell(5 + i, 1 + j, v)
    teachers = sorted({r[0] for r in assign})

    # 2. 共同時段
    ws = wb["共同時段"]
    for i, r in enumerate(COMMON):
        for j, v in enumerate(r):
            ws.cell(4 + i, 1 + j, v)

    # 3. 本土語預排：每班 1 節閩南語，閩南語甲 23 班、乙 22 班、丙 21 班、丁 3 班、科任英語乙兼 2 班；
    #    少數班有客語／族語／新住民語學生，同一節到別處分組上課（同一位老師同一節帶多班＝合班，正常）
    ws = wb["本土語預排"]
    # 跟學校實際預排一樣：每位閩南語老師的班級跳班、跳年級，隨機打散
    all_classes = [(g, i) for g, n in CLASSES.items() for i in range(1, n + 1)]
    random.shuffle(all_classes)
    load = [("閩南語甲", 23), ("閩南語乙", 22), ("閩南語丙", 21), ("閩南語丁", 3), ("英語乙", 2)]
    busy = {}
    rows = []
    k = 0
    for t, n in load:
        for g, i in all_classes[k:k + n]:
            options = [s2 for s2 in class_slots(g) if s2[1] != 7 and (t, s2) not in busy
                       and sum(1 for (t2, s3) in busy if t2 == t and s3[0] == s2[0]) < 6]
            d, p = random.choice(options)
            busy[(t, (d, p))] = class_name(g, i)
            rows.append([class_name(g, i), d, str(p), "閩南語", t, "原班教室", "（試跑）科任兼本土語" if t == "英語乙" else None])
        k += n
    for idx in random.sample(range(len(rows)), 10):
        lang, t = random.choice([("客語", "客語甲"), ("客語", "客語甲"), ("族語", "族語甲"), ("新住民語", "新住民語甲")])
        r0 = rows[idx]
        rows.append([r0[0], r0[1], r0[2], lang, t, "本土語教室", "分組上課"])
    for r, vals in enumerate(rows):
        for j, v in enumerate(vals):
            ws.cell(5 + r, 1 + j, v)

    # 4. 特教抽離：10 筆
    ws = wb["特教抽離"]
    picks = random.sample([(g, i) for g in range(2, 7) for i in range(1, CLASSES[g] + 1)], 10)
    for k, (g, i) in enumerate(picks):
        d, p = random.choice([s for s in class_slots(g) if s[1] <= 4])
        vals = ["資源班抽離", class_name(g, i), random.choice(["國語", "數學"]), d, str(p),
                f"資源班{GANZHI[k % 3]}", random.randint(1, 3), None]
        for j, v in enumerate(vals):
            ws.cell(5 + k, 1 + j, v)

    # 5. 職務限制（虛構職務與文號）
    ws = wb["職務限制"]
    duty_rules = [("（試跑）英語輔導團團員", "四", "上午", "高市教○字第000001號（假）"),
                  ("（試跑）自然輔導團團員", "二", "下午", "高市教○字第000002號（假）")]
    for k, r in enumerate(duty_rules):
        for j, v in enumerate(r):
            ws.cell(4 + k, 1 + j, v)
    for k, (t, duty) in enumerate([("英語丙", duty_rules[0][0]), ("自然丁", duty_rules[1][0])]):
        ws.cell(4 + k, 7, t)
        ws.cell(4 + k, 8, duty)

    # 6. 班級教室：每年級 12 班分 2 層，每層 6 班
    ws = wb["班級教室"]
    bld = {1: ("甲棟", 1), 4: ("甲棟", 3), 2: ("乙棟", 1), 5: ("乙棟", 3), 3: ("丙棟", 1), 6: ("丙棟", 3)}
    r = 4
    for g, n in CLASSES.items():
        b, f0 = bld[g]
        for i in range(1, n + 1):
            f = f0 + (i - 1) // 6
            ws.cell(r, 2, b)
            ws.cell(r, 3, f)
            ws.cell(r, 4, f"{b[0]}{f}{(i - 1) % 6 + 1:02d}")
            r += 1
    for k, (name, n, cap, pos) in enumerate([("自然教室", 4, 4, "丁棟2-3F"), ("音樂教室", 5, 5, "丁棟1F"),
                                             ("電腦教室", 2, 2, "丁棟4F"), ("操場／體育館", 1, 8, "操場")]):
        ws.cell(4 + k, 7, name)
        ws.cell(4 + k, 8, n)
        ws.cell(4 + k, 9, cap)
        ws.cell(4 + k, 10, pos)

    # 7. 老師需求：約 1/3 老師有需求
    ws = wb["老師需求"]
    r = 5
    reasons = ["接送小孩", "進修", "身體因素", "兼任行政會議", None]
    for t in random.sample(teachers, len(teachers) // 3):
        n_must = random.randint(0, 3)
        for _ in range(n_must):
            d = random.choice(DAYS)
            p = random.choice(["1", "2", "3", "4", "6"])
            for j, v in enumerate([t, d, p, "必須", random.choice(reasons)]):
                ws.cell(r, 1 + j, v)
            r += 1
        if random.random() < 0.6:
            for j, v in enumerate([t, "五", "下午", "希望", "接送小孩"]):
                ws.cell(r, 1 + j, v)
            r += 1
        if random.random() < 0.4:
            for j, v in enumerate([t, random.choice(DAYS), "1", "希望", None]):
                ws.cell(r, 1 + j, v)
            r += 1

    add_rules(wb)

    out.parent.mkdir(exist_ok=True)
    wb.save(out)
    print(f"已產生 {out.name}：{len(teachers)} 位科任、{len(assign)} 筆配課")


if __name__ == "__main__":
    import sys
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 115, Path(sys.argv[2]) if len(sys.argv) > 2 else OUT)
