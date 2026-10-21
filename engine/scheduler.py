"""楠梓國小科任排課引擎。

讀「排課需求收集表」（xlsx，之後改讀 Google 試算表匯出的同格式檔案），
依「規則表」的 0 法規 → 1 鎖定 → 2 必須 → 3 盡量 排出科任課表。

用法：python engine/scheduler.py 試跑/試跑資料_排課需求收集表.xlsx [輸出資料夾]
"""
import json
import re
import sys
import time
from datetime import datetime
from zoneinfo import ZoneInfo
from collections import defaultdict
from pathlib import Path

import openpyxl
from ortools.sat.python import cp_model

DAYS = ["一", "二", "三", "四", "五"]
PERIODS = [1, 2, 3, 4, 5, 6, 7]
# 可以連堂的兩節（不跨午休）
PAIRS = [(1, 2), (2, 3), (3, 4), (5, 6), (6, 7)]
# 跑班距離只算「中間只有 10 分鐘下課」的連續兩節（大下課、午休算休息）
WALK_PAIRS = [(1, 2), (3, 4), (5, 6), (6, 7)]
WEIGHT = {"高": 10, "中": 4, "低": 1}
SUBJECT_ABBR = {"國": "國語", "數": "數學", "英": "英語", "社": "社會", "自": "自然"}
CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6}
IN_CLASS = "原班教室"


# ---------- 讀資料 ----------

def cell_str(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def is_example(row):
    return any("範例" in cell_str(v) for v in row)


def read_rows(table, first_row, ncols, first_col=1):
    """table＝一張工作表的所有儲存格（二維清單）；回傳去掉空白列與範例列的字串列。"""
    out = []
    for r in table[first_row - 1:]:
        r = [cell_str(v) for v in list(r[first_col - 1:ncols]) + [None] * ncols][:ncols - first_col + 1]
        if not any(r) or is_example(r):
            continue
        out.append(r)
    return out


def parse_periods(text):
    """'1'、'5–7'、'4、5'、'上午'、'下午'、'整天'、'不排第7節' → 節次集合"""
    t = text.replace("第", "").replace("節", "")
    if "整天" in t:
        return set(PERIODS)
    out = set()
    if "上午" in t:
        out |= {1, 2, 3, 4}
    if "下午" in t:
        out |= {5, 6, 7}
    for a, b in re.findall(r"(\d)\s*[–\-~～至]\s*(\d)", t):
        out |= set(range(int(a), int(b) + 1))
    t = re.sub(r"(\d)\s*[–\-~～至]\s*(\d)", "", t)
    out |= {int(x) for x in re.findall(r"\d", t) if 1 <= int(x) <= 7}
    return out


def parse_days(text):
    m = re.findall(r"星期([一二三四五])", text)
    if m:
        return set(m)
    if text in DAYS:
        return {text}
    return set(DAYS)


def parse_grades(target):
    """'全校' → 1–6；'三、四年級' → {3,4}；'3年1班' → None（單一班級）"""
    if target == "全校":
        return {1, 2, 3, 4, 5, 6}
    if "班" in target:
        return None
    g = {CN_NUM[c] for c in target if c in CN_NUM}
    g |= {int(c) for c in target if c.isdigit() and 1 <= int(c) <= 6}
    return g


SHEETS = ["規則表", "科任配課", "本土語預排", "特教抽離", "共同時段", "職務限制", "班級教室", "老師需求"]


def xlsx_to_tables(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    return {n: [list(r) for r in wb[n].iter_rows(values_only=True)] for n in SHEETS}


def load(path):
    return load_tables(xlsx_to_tables(path))


def load_tables(wb):
    """wb＝{工作表名稱: 二維清單}；可來自 xlsx，也可來自 Google 試算表（Apps Script 傳來的 JSON）。"""
    missing = [n for n in SHEETS if n not in wb]
    if missing:
        raise ValueError(f"試算表缺少工作表：{'、'.join(missing)}")
    d = {}
    d["rules"] = [dict(id=r[0], level=int(r[1][:1]) if r[1][:1].isdigit() else 3, block=r[2],
                       target=r[3], content=r[4], importance=r[5], enabled=(r[6] != "否"), reason=r[7])
                  for r in read_rows(wb["規則表"], 4, 8) if r[0]]
    d["assign"] = [dict(teacher=r[0], subject=r[1], cls=r[2], hours=int(float(r[3] or 0)),
                        room=r[4] or IN_CLASS, block=int(float(r[5])) if r[5] else 0, note=r[6])
                   for r in read_rows(wb["科任配課"], 5, 7) if r[0] and r[2]]
    d["lang"] = [dict(cls=r[0], day=r[1], period=int(r[2]), lang=r[3], teacher=r[4])
                 for r in read_rows(wb["本土語預排"], 5, 7) if r[0] and r[1] and r[2]]
    d["sped"] = [dict(kind=r[0], cls=r[1], subject=r[2], day=r[3], period=int(r[4]), teacher=r[5])
                 for r in read_rows(wb["特教抽離"], 5, 8) if r[1] and r[3] and r[4]]
    d["common"] = [dict(target=r[0], day=r[1], periods=parse_periods(r[2]), activity=r[3], who=r[4],
                        lock=(r[5] != "否"), note=r[6])
                   for r in read_rows(wb["共同時段"], 4, 7) if r[0] and r[1] and r[2]]
    ws = wb["職務限制"]
    d["duty_rules"] = [dict(duty=r[0], day=r[1], periods=parse_periods(r[2]), ref=r[3])
                       for r in read_rows(ws, 4, 5) if r[0] and r[1]]
    d["duty_members"] = [dict(teacher=r[6], duty=r[7]) for r in read_rows(ws, 4, 8) if r[6] and r[7]]
    ws = wb["班級教室"]
    d["classes"] = {}
    for r in read_rows(ws, 4, 4):
        if r[0]:
            floor = int(float(r[2])) if r[2] else 0
            d["classes"][r[0]] = dict(building=r[1], floor=floor, room=r[3],
                                      grade=int(r[0].split("年")[0]))
    d["rooms"] = {}
    for r in read_rows(ws, 4, 11, first_col=7):
        if r[0]:
            cap = r[2] or r[1] or "1"
            d["rooms"][r[0]] = dict(cap=int(float(cap)), where=r[3], subjects=r[4])
    d["needs"] = [dict(teacher=r[0], day=r[1], periods=parse_periods(r[2]), level=r[3], reason=r[4],
                       raw=r[2])
                  for r in read_rows(wb["老師需求"], 5, 5) if r[0] and r[1] and r[2]]
    return d


# ---------- 排課 ----------

def slot_key(day, p):
    return f"{day}{p}"


class Scheduler:
    def __init__(self, data):
        self.d = data
        self.rules = {r["id"]: r for r in data["rules"]}
        self.on = lambda block: [r for r in data["rules"] if r["block"] == block and r["enabled"]]
        self.issues = []  # 事前檢查
        self.build_grid()

    def rule_enabled_for_target(self, target):
        for r in self.d["rules"]:
            if r["block"] == "鎖定時段" and r["target"] == target:
                return r["enabled"]
        return True

    def build_grid(self):
        """每班每格先標上：不上課／鎖定／本土語／特教抽離，剩下的才給系統排。"""
        d = self.d
        self.grid = {c: {} for c in d["classes"]}
        for row in d["common"]:
            grades = parse_grades(row["target"])
            if not row["lock"]:
                continue
            off = row["activity"] == "不上課"
            if not off and not self.rule_enabled_for_target(row["target"]):
                continue
            for c, info in d["classes"].items():
                if (grades is None and c != row["target"]) or (grades is not None and info["grade"] not in grades):
                    continue
                for p in row["periods"]:
                    k = slot_key(row["day"], p)
                    if off:
                        self.grid[c][k] = dict(type="off", label="不上課")
                    elif k not in self.grid[c]:
                        self.grid[c][k] = dict(type="lock", label=row["activity"], who=row["who"])
        self.teacher_block = defaultdict(dict)  # 老師 → 格 → 原因（不能排）
        self.lang_slots = defaultdict(lambda: defaultdict(list))  # 本土語老師 → 格 → [(班級, 語別)]
        if self.rule_enabled_for_target("本土語"):
            for row in d["lang"]:
                k = slot_key(row["day"], row["period"])
                if row["cls"] in self.grid:
                    self.grid[row["cls"]][k] = dict(type="lang", label=row["lang"], who=row["teacher"])
                self.teacher_block[row["teacher"]][k] = f"本土語（{row['cls']}）"
                self.lang_slots[row["teacher"]][k].append((row["cls"], row["lang"]))
        if self.rule_enabled_for_target("特教抽離"):
            for row in d["sped"]:
                k = slot_key(row["day"], row["period"])
                if row["cls"] in self.grid and self.grid[row["cls"]].get(k, {}).get("type") != "off":
                    self.grid[row["cls"]][k] = dict(type="sped", label=row["subject"], who="導師",
                                                    note=f"特教抽離：原班上{row['subject']}")

        # 第 0 層：職務不排時段
        self.duty_slots = defaultdict(dict)
        if self.on("職務不排時段"):
            rules_by_duty = defaultdict(list)
            for r in d["duty_rules"]:
                rules_by_duty[r["duty"]].append(r)
            for m in d["duty_members"]:
                for r in rules_by_duty.get(m["duty"], []):
                    for p in r["periods"]:
                        self.duty_slots[m["teacher"]][slot_key(r["day"], p)] = f"{m['duty']}（{r['ref']}）"

        # 老師需求
        self.must = defaultdict(dict)
        self.wish = defaultdict(dict)
        for n in d["needs"]:
            target = self.must if n["level"] == "必須" else self.wish
            for p in n["periods"]:
                target[n["teacher"]][slot_key(n["day"], p)] = n["reason"] or n["raw"]

        # 科目不排時段（分硬性／盡量）
        self.subj_forbid = defaultdict(set)   # 科目 → 格（硬性）
        self.subj_avoid = []                  # (規則, 科目集合, 格集合)
        for r in self.on("科目不排時段"):
            subjects = set()
            for part in re.split(r"[、,，\s]+", r["target"]):
                subjects.add(SUBJECT_ABBR.get(part, part))
            days = parse_days(r["content"])
            slots = {slot_key(dd, p) for dd in days for p in parse_periods(r["content"])}
            if r["level"] <= 2:
                for s in subjects:
                    self.subj_forbid[s] |= slots
            else:
                self.subj_avoid.append((r, subjects, slots))

    def free(self, cls, k):
        return k not in self.grid.get(cls, {})

    def teacher_hard_block(self, t):
        out = dict(self.teacher_block.get(t, {}))
        out.update(self.duty_slots.get(t, {}))
        if any(r["level"] <= 2 for r in self.on("老師不排時段")):
            out.update(self.must.get(t, {}))
        return out

    # ----- 事前檢查（白話說明可能排不出來的原因） -----
    def precheck(self):
        d = self.d
        all_slots = [slot_key(dd, p) for dd in DAYS for p in PERIODS]
        teachers = defaultdict(list)
        for a in d["assign"]:
            teachers[a["teacher"]].append(a)
            if a["cls"] not in d["classes"]:
                self.issues.append(dict(level="錯誤", text=f"配課的班級「{a['cls']}」不在「班級教室」頁，請檢查班級名稱。"))
            if a["room"] != IN_CLASS and a["room"] not in d["rooms"]:
                self.issues.append(dict(level="錯誤", text=f"上課地點「{a['room']}」不在專科教室清單（{a['teacher']}／{a['cls']}）。"))
        for t, rows in teachers.items():
            hrs = sum(a["hours"] for a in rows)
            blocked = self.teacher_hard_block(t)
            avail = len([k for k in all_slots if k not in blocked])
            # 只算「至少有一個任教班級在校、沒鎖定」的格子（例如只教低年級的老師，下午幾乎都不能排）
            usable = len([k for k in all_slots if k not in blocked
                          and any(self.free(a["cls"], k) and k not in self.subj_forbid.get(a["subject"], set())
                                  for a in rows)])
            if hrs > avail:
                self.issues.append(dict(level="錯誤", text=f"{t} 要上 {hrs} 節，但扣掉不能排的時段只剩 {avail} 格。"))
            elif hrs > usable:
                self.issues.append(dict(level="錯誤", text=f"{t} 要上 {hrs} 節，但任教班級的在校時段加起來只有 {usable} 格（例如低年級下午不上課），一定排不下，請調整配課。"))
            n_must = len(self.must.get(t, {}))
            if n_must > 3:
                self.issues.append(dict(level="提醒", text=f"{t} 勾了 {n_must} 節「必須」不排，超過每人 3 節的上限。"))
        self.precheck_lang()
        for c in d["classes"]:
            need = sum(a["hours"] for a in d["assign"] if a["cls"] == c)
            free = len([k for k in all_slots if self.free(c, k)])
            if need > free:
                self.issues.append(dict(level="錯誤", text=f"{c} 科任要 {need} 節，但扣掉鎖定時段只剩 {free} 格。"))
        for room, info in d["rooms"].items():
            need = sum(a["block"] or a["hours"] for a in d["assign"] if a["room"] == room)
            slots = len([k for k in all_slots if any(self.free(c, k) for c in d["classes"])])
            if need > info["cap"] * slots:
                self.issues.append(dict(level="錯誤", text=f"{room} 一週要用 {need} 節，但 {info['cap']} 間 × {slots} 格只夠 {info['cap'] * slots} 節。"))
            elif need > 0.85 * info["cap"] * slots:
                self.issues.append(dict(level="提醒", text=f"{room} 使用率 {need * 100 // (info['cap'] * slots)}%，很擠，可能讓其他規則難以做到。"))

    def precheck_lang(self):
        """檢查「本土語預排」本身有沒有排錯：老師撞堂、排在不上課或鎖定時段、單日節數過多。"""
        d = self.d
        lim = next((int(re.search(r"\d+", r["content"]).group()) for r in self.on("每日節數上限")), 6)
        base = {c: {} for c in d["classes"]}  # 只看共同時段（不含本土語自己）
        for row in d["common"]:
            grades = parse_grades(row["target"])
            for c, info in d["classes"].items():
                if (grades is None and c == row["target"]) or (grades is not None and info["grade"] in grades):
                    for p in row["periods"]:
                        base[c][slot_key(row["day"], p)] = row["activity"]
        for row in d["lang"]:
            k = slot_key(row["day"], row["period"])
            where = f"{row['cls']} 星期{row['day']}第{row['period']}節"
            if row["cls"] not in d["classes"]:
                self.issues.append(dict(level="錯誤", text=f"本土語預排的班級「{row['cls']}」不在「班級教室」頁。"))
            elif k in base[row["cls"]]:
                self.issues.append(dict(level="錯誤", text=f"本土語預排：{where}（{row['teacher']}）排在「{base[row['cls']][k]}」的時段，請改時間。"))
        for t, slots in self.lang_slots.items():
            for k, rows in slots.items():
                classes = sorted({c for c, _ in rows})
                langs = {lg for _, lg in rows}
                if len(classes) > 1 and "閩南語" in langs:
                    self.issues.append(dict(level="錯誤", text=f"本土語預排：{t} 星期{k[0]}第{k[1:]}節同時排了 {'、'.join(classes)}，閩南語是整班上課，會撞堂。（如果真的是合班上課，可以忽略）"))
            for dd in DAYS:
                n = len([k for k in slots if k[0] == dd])
                if n > lim:
                    self.issues.append(dict(level="提醒", text=f"本土語預排：{t} 星期{dd} 就有 {n} 節本土語，超過每天 {lim} 節。"))
            total = len(slots)
            busy = sum(a["hours"] for a in d["assign"] if a["teacher"] == t)
            if busy:
                self.issues.append(dict(level="提醒", text=f"{t} 同時有本土語 {total} 節＋科任 {busy} 節，系統已自動避開本土語時段。"))

    # ----- 建模 -----
    def solve(self, time_limit=60, relax=False):
        d = self.d
        m = cp_model.CpModel()
        block_rule = any(r["target"] in ("自然",) or r["block"] == "連堂" for r in self.on("連堂"))
        spread = self.on("同科分散")
        daily = self.on("每日節數上限")
        walk = self.on("跑班距離")
        wish_rules = [r for r in self.on("老師不排時段") if r["level"] == 3]
        penalties = []  # (權重, 變數或運算式, 說明代碼)

        self.vars = []  # 每筆配課的變數
        occ_teacher = defaultdict(lambda: defaultdict(list))
        occ_class = defaultdict(lambda: defaultdict(list))
        occ_room = defaultdict(lambda: defaultdict(list))
        relax_vars = []

        for i, a in enumerate(d["assign"]):
            t, c = a["teacher"], a["cls"]
            if c not in d["classes"]:
                continue
            hard = self.teacher_hard_block(t)
            forbid = self.subj_forbid.get(a["subject"], set())

            def ok(k):
                if not self.free(c, k) or k in forbid:
                    return False
                if k in hard:
                    return relax and k not in self.teacher_block.get(t, {})
                return True

            use_block = block_rule and a["block"] >= 2 and a["room"] != IN_CLASS
            v = dict(a=a, singles={}, blocks={}, i=i)
            if use_block:
                for dd in DAYS:
                    for p, q in PAIRS:
                        k1, k2 = slot_key(dd, p), slot_key(dd, q)
                        if ok(k1) and ok(k2):
                            v["blocks"][(dd, p)] = m.NewBoolVar(f"b{i}{dd}{p}")
                m.AddExactlyOne(v["blocks"].values()) if v["blocks"] else m.AddBoolOr([])
                n_single = a["hours"] - a["block"]
            else:
                n_single = a["hours"]
            for dd in DAYS:
                for p in PERIODS:
                    k = slot_key(dd, p)
                    if ok(k) and n_single > 0:
                        v["singles"][(dd, p)] = m.NewBoolVar(f"x{i}{dd}{p}")
            m.Add(sum(v["singles"].values()) == n_single)
            single_room = IN_CLASS if use_block else a["room"]

            # 佔用
            for (dd, p), x in v["singles"].items():
                k = slot_key(dd, p)
                occ_teacher[t][k].append(x)
                occ_class[c][k].append(x)
                if single_room != IN_CLASS:
                    occ_room[single_room][k].append(x)
                if k in hard:
                    relax_vars.append((x, t, k, hard[k]))
            for (dd, p), b in v["blocks"].items():
                for q in (p, p + 1):
                    k = slot_key(dd, q)
                    occ_teacher[t][k].append(b)
                    occ_class[c][k].append(b)
                    occ_room[a["room"]][k].append(b)
                    if k in hard:
                        relax_vars.append((b, t, k, hard[k]))
            # 連堂那 2 節與原班那節不相連
            if use_block:
                for (dd, p), b in v["blocks"].items():
                    for q in (p - 1, p + 2):
                        if (dd, q) in v["singles"]:
                            m.AddBoolOr([b.Not(), v["singles"][(dd, q)].Not()])
            # 同科分散：一天最多 1 次（連堂算 1 次）
            for dd in DAYS:
                events = [x for (d2, _), x in v["singles"].items() if d2 == dd] + \
                         [b for (d2, _), b in v["blocks"].items() if d2 == dd]
                if len(events) > 1 and spread:
                    r = spread[0]
                    if r["level"] <= 2:
                        m.Add(sum(events) <= 1)
                    else:
                        ex = m.NewIntVar(0, len(events), f"sp{i}{dd}")
                        m.Add(ex >= sum(events) - 1)
                        penalties.append((WEIGHT.get(r["importance"], 4), ex))
            self.vars.append(v)

        # 必須：同一位老師同一節只能一班；同一班同一節只能一門科任；教室容量
        for t, slots in occ_teacher.items():
            for k, xs in slots.items():
                if len(xs) > 1:
                    m.Add(sum(xs) <= 1)
        for c, slots in occ_class.items():
            for k, xs in slots.items():
                if len(xs) > 1:
                    m.Add(sum(xs) <= 1)
        if self.on("教室容量"):
            for room, slots in occ_room.items():
                cap = d["rooms"].get(room, {}).get("cap", 1)
                for k, xs in slots.items():
                    if len(xs) > cap:
                        m.Add(sum(xs) <= cap)

        # 放寬模式：違反老師時段要付很高代價
        if relax:
            for x, *_ in relax_vars:
                penalties.append((1000, x))

        # 盡量：科目不排時段
        for r, subjects, slots in self.subj_avoid:
            w = WEIGHT.get(r["importance"], 4)
            for v in self.vars:
                if v["a"]["subject"] in subjects:
                    for (dd, p), x in v["singles"].items():
                        if slot_key(dd, p) in slots:
                            penalties.append((w, x))
                    for (dd, p), b in v["blocks"].items():
                        if slot_key(dd, p) in slots or slot_key(dd, p + 1) in slots:
                            penalties.append((w, b))
        # 盡量：老師「希望」不排
        if wish_rules:
            w = WEIGHT.get(wish_rules[0]["importance"], 4)
            for t, slots in self.wish.items():
                for k in slots:
                    for x in occ_teacher[t].get(k, []):
                        penalties.append((w, x))
        # 每日節數上限
        for r in daily:
            lim = int(re.search(r"\d+", r["content"]).group())
            w = WEIGHT.get(r["importance"], 4)
            for t, slots in occ_teacher.items():
                for dd in DAYS:
                    xs = [x for p in PERIODS for x in slots.get(slot_key(dd, p), [])]
                    fixed = len([k for k in self.lang_slots.get(t, {}) if k[0] == dd])  # 已預排的本土語
                    if len(xs) + fixed <= lim:
                        continue
                    if r["level"] <= 2:
                        m.Add(sum(xs) + fixed <= lim)
                    else:
                        ex = m.NewIntVar(0, 7, f"dl{t}{dd}")
                        m.Add(ex >= sum(xs) + fixed - lim)
                        penalties.append((w, ex))
        # 跑班距離：只看到班上課的節
        self.walk_cost = {}
        if walk:
            w = WEIGHT.get(walk[0]["importance"], 4)
            loc_of = {c: (i["building"], i["floor"]) for c, i in d["classes"].items()}
            for t in occ_teacher:
                at = defaultdict(lambda: defaultdict(list))  # 格 → 位置 → 變數
                for v in self.vars:
                    if v["a"]["teacher"] != t:
                        continue
                    single_in_class = (v["a"]["room"] == IN_CLASS) or bool(v["blocks"])
                    if not single_in_class:
                        continue
                    L = loc_of[v["a"]["cls"]]
                    for (dd, p), x in v["singles"].items():
                        at[slot_key(dd, p)][L].append(x)
                for dd in DAYS:
                    for p, q in WALK_PAIRS:
                        k1, k2 = slot_key(dd, p), slot_key(dd, q)
                        for L1, xs1 in at.get(k1, {}).items():
                            for L2, xs2 in at.get(k2, {}).items():
                                dist = distance(L1, L2)
                                if dist == 0:
                                    continue
                                z = m.NewBoolVar("")
                                m.Add(z >= sum(xs1) + sum(xs2) - 1)
                                penalties.append((w * dist, z))

        m.Minimize(sum(wt * x for wt, x in penalties))
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = time_limit
        solver.parameters.num_workers = 8
        t0 = time.time()
        status = solver.Solve(m)
        self.solve_info = dict(status=solver.StatusName(status), seconds=round(time.time() - t0, 1),
                               objective=solver.ObjectiveValue() if status in (cp_model.OPTIMAL, cp_model.FEASIBLE) else None,
                               relax=relax)
        if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            return False
        self.relaxed_hits = [dict(teacher=t, slot=k, reason=why) for x, t, k, why in relax_vars if solver.Value(x)]
        self.extract(solver)
        return True

    def extract(self, solver):
        self.lessons = []  # 每一節科任課
        for v in self.vars:
            a = v["a"]
            for (dd, p), x in v["singles"].items():
                if solver.Value(x):
                    room = IN_CLASS if v["blocks"] else a["room"]
                    self.lessons.append(dict(teacher=a["teacher"], subject=a["subject"], cls=a["cls"],
                                             day=dd, period=p, room=room, block=False))
            for (dd, p), b in v["blocks"].items():
                if solver.Value(b):
                    for q in (p, p + 1):
                        self.lessons.append(dict(teacher=a["teacher"], subject=a["subject"], cls=a["cls"],
                                                 day=dd, period=q, room=a["room"], block=True))


def distance(L1, L2):
    if L1 == L2:
        return 0
    return abs(L1[1] - L2[1]) + (3 if L1[0] != L2[0] else 0)


# ---------- 規則體檢 ----------

def check(s):
    d = s.d
    by_teacher = defaultdict(list)
    for l in s.lessons:
        by_teacher[l["teacher"]].append(l)
    loc_of = {c: (i["building"], i["floor"]) for c, i in d["classes"].items()}
    report = []

    def add(rule, ok, summary, details=()):
        report.append(dict(id=rule["id"], level=rule["level"], block=rule["block"], target=rule["target"],
                           content=rule["content"], importance=rule["importance"], enabled=rule["enabled"],
                           reason=rule["reason"], ok=ok, summary=summary, details=list(details)[:60],
                           n=len(list(details))))

    for r in d["rules"]:
        if not r["enabled"]:
            add(r, None, "已暫停（啟用＝否）")
            continue
        b = r["block"]
        if b == "職務不排時段":
            bad = [f"{l['teacher']} 星期{l['day']}第{l['period']}節" for l in s.lessons
                   if slot_key(l["day"], l["period"]) in s.duty_slots.get(l["teacher"], {})]
            n_t = len(s.duty_slots)
            add(r, not bad, f"{n_t} 位老師有職務限制，全部避開" if not bad else f"有 {len(bad)} 節排在職務時段", bad)
        elif b == "科目不排時段":
            subjects = {SUBJECT_ABBR.get(x, x) for x in re.split(r"[、,，\s]+", r["target"])}
            slots = {slot_key(dd, p) for dd in parse_days(r["content"]) for p in parse_periods(r["content"])}
            bad = [f"{l['cls']} {l['subject']}（{l['teacher']}）星期{l['day']}第{l['period']}節" for l in s.lessons
                   if l["subject"] in subjects and slot_key(l["day"], l["period"]) in slots]
            total = len([l for l in s.lessons if l["subject"] in subjects])
            add(r, not bad, "完全做到" if not bad else f"{total} 節中有 {len(bad)} 節沒避開", bad)
        elif b == "鎖定時段":
            add(r, True, "已先鎖住，系統沒有動")
        elif b == "其他" and "同一節只能在一個班" in r["content"]:
            seen, bad = set(), []
            for l in s.lessons:
                k = (l["teacher"], l["day"], l["period"])
                if k in seen:
                    bad.append(f"{l['teacher']} 星期{l['day']}第{l['period']}節")
                seen.add(k)
            add(r, not bad, f"{len(by_teacher)} 位科任都沒有撞堂" if not bad else "有撞堂", bad)
        elif b == "教室容量":
            use = defaultdict(int)
            for l in s.lessons:
                if l["room"] != IN_CLASS:
                    use[(l["room"], l["day"], l["period"])] += 1
            bad = [f"{rm} 星期{dd}第{p}節 {n} 班（只有 {d['rooms'][rm]['cap']} 間）" for (rm, dd, p), n in use.items()
                   if n > d["rooms"].get(rm, {}).get("cap", 1)]
            peak = {rm: max([n for (r2, _, _), n in use.items() if r2 == rm] or [0]) for rm in d["rooms"]}
            txt = "、".join(f"{rm}最多同時 {n}／{d['rooms'][rm]['cap']} 間" for rm, n in peak.items())
            add(r, not bad, txt, bad)
        elif b == "連堂":
            bad = []
            for v in s.vars:
                a = v["a"]
                if a["subject"] != r["target"] or not a["block"]:
                    continue
                mine = sorted([l for l in s.lessons if l["cls"] == a["cls"] and l["teacher"] == a["teacher"]],
                              key=lambda l: (DAYS.index(l["day"]), l["period"]))
                blocks = [l for l in mine if l["block"]]
                if len(blocks) != a["block"]:
                    bad.append(f"{a['cls']} 沒有排到連堂")
            n = len([v for v in s.vars if v["a"]["subject"] == r["target"] and v["a"]["block"]])
            add(r, not bad, f"{n} 班都是 2 節連堂在專科教室＋1 節在原班，且不相連" if not bad else "有班級沒排好", bad)
        elif b == "老師不排時段":
            src = s.must if r["level"] <= 2 else s.wish
            asked = sum(len(v) for v in src.values())
            bad = [f"{l['teacher']} 星期{l['day']}第{l['period']}節（{src[l['teacher']][slot_key(l['day'], l['period'])]}）"
                   for l in s.lessons if slot_key(l["day"], l["period"]) in src.get(l["teacher"], {})]
            add(r, not bad, f"{asked} 格需求全部做到" if not bad else f"{asked} 格需求中有 {len(bad)} 格沒做到", bad)
        elif b == "每日節數上限":
            lim = int(re.search(r"\d+", r["content"]).group())
            bad = []
            for t in sorted(set(by_teacher) | set(s.lang_slots)):
                for dd in DAYS:
                    n = len([l for l in by_teacher.get(t, []) if l["day"] == dd])
                    fixed = len([k for k in s.lang_slots.get(t, {}) if k[0] == dd])
                    if n + fixed > lim:
                        note = "（全是預排本土語，系統無法調整）" if n == 0 else (f"（含預排本土語 {fixed} 節）" if fixed else "")
                        bad.append(f"{t} 星期{dd} {n + fixed} 節{note}")
            add(r, not bad, f"所有老師每天都在 {lim} 節以內（含本土語）" if not bad else f"{len(bad)} 人次超過", bad)
        elif b == "跑班距離":
            moves, bad = 0, []
            for t, ls in by_teacher.items():
                at = {(l["day"], l["period"]): l for l in ls if l["room"] == IN_CLASS}
                for dd in DAYS:
                    for p, q in WALK_PAIRS:
                        l1, l2 = at.get((dd, p)), at.get((dd, q))
                        if l1 and l2:
                            moves += 1
                            dist = distance(loc_of[l1["cls"]], loc_of[l2["cls"]])
                            if dist > 0:
                                where = "換棟" if loc_of[l1["cls"]][0] != loc_of[l2["cls"]][0] else f"差 {dist} 層樓"
                                bad.append(f"{t} 星期{dd}第{p}→{q}節：{l1['cls']}→{l2['cls']}（{where}）")
            add(r, not bad, f"{moves} 次連續兩節換班，全部在同一層" if not bad else
                f"{moves} 次連續兩節換班，{moves - len(bad)} 次在同一層，{len(bad)} 次要爬樓或換棟", bad)
        elif b == "同科分散":
            cnt = defaultdict(set)
            for l in s.lessons:
                cnt[(l["cls"], l["subject"], l["day"])].add((l["period"], l["block"]))
            bad = []
            for (c, subj, dd), ps in cnt.items():
                n_events = len([p for p, blk in ps if not blk]) + (1 if any(blk for _, blk in ps) else 0)
                if n_events > 1:
                    bad.append(f"{c} {subj} 星期{dd} 上 {n_events} 次")
            add(r, not bad, "每班同一科都分散在不同天" if not bad else f"{len(bad)} 處同一天上兩次", bad)
        else:
            add(r, None, "這條規則系統還不會自動處理，請人工確認")
    return report


# ---------- 輸出 ----------

def build_output(s, src_name):
    d = s.d
    classes = {}
    for c, info in d["classes"].items():
        cells = {}
        for dd in DAYS:
            for p in PERIODS:
                k = slot_key(dd, p)
                if k in s.grid[c]:
                    cells[k] = dict(s.grid[c][k])
        for l in s.lessons:
            if l["cls"] == c:
                cells[slot_key(l["day"], l["period"])] = dict(type="subject", label=l["subject"], who=l["teacher"],
                                                              room=l["room"], block=l["block"])
        n_free = len([1 for dd in DAYS for p in PERIODS if slot_key(dd, p) not in cells])
        n_keep = len([1 for v in cells.values() if v["type"] in ("lock", "sped") and v.get("who") == "導師"])
        classes[c] = dict(grade=info["grade"], building=info["building"], floor=info["floor"], room=info["room"],
                          cells=cells, free=n_free, keep=n_keep,
                          subject_hours=len([l for l in s.lessons if l["cls"] == c]))
    teachers = {}
    for a in d["assign"]:
        t = teachers.setdefault(a["teacher"], dict(subjects=set(), hours=0, cells={}))
        t["subjects"].add(a["subject"])
        t["hours"] += a["hours"]
    for t, slots in s.lang_slots.items():  # 本土語老師也列出來（課是預排鎖定的）
        info = teachers.setdefault(t, dict(subjects=set(), hours=0, cells={}))
        for k, rows in slots.items():
            info["subjects"].update(lg for _, lg in rows)
            info["hours"] += 1
            info["cells"][k] = dict(type="lang", label="、".join(sorted({c for c, _ in rows})),
                                    who="、".join(sorted({lg for _, lg in rows})) + "・預排", note="本土語（預排鎖定）")
    for t, info in teachers.items():
        info["subjects"] = sorted(info["subjects"])
        for k, why in s.duty_slots.get(t, {}).items():
            info["cells"][k] = dict(type="duty", label="職務", note=why)
        for k, why in s.must.get(t, {}).items():
            info["cells"].setdefault(k, dict(type="must", label="必須不排", note=why))
        for k, why in s.wish.get(t, {}).items():
            info["cells"].setdefault(k, dict(type="wish", label="希望不排", note=why))
    for l in s.lessons:
        k = slot_key(l["day"], l["period"])
        prev = teachers[l["teacher"]]["cells"].get(k)
        teachers[l["teacher"]]["cells"][k] = dict(type="subject", label=l["cls"], subject=l["subject"],
                                                  room=l["room"], block=l["block"],
                                                  conflict=prev["type"] if prev and prev["type"] == "wish" else None)
    for t, info in teachers.items():
        per_day = [len([1 for p in PERIODS if info["cells"].get(slot_key(dd, p), {}).get("type") in ("subject", "lang")])
                   for dd in DAYS]
        info["per_day"] = per_day
    rooms = {}
    for rm, info in d["rooms"].items():
        use = defaultdict(list)
        for l in s.lessons:
            if l["room"] == rm:
                use[slot_key(l["day"], l["period"])].append(f"{l['cls']}")
        rooms[rm] = dict(cap=info["cap"], where=info["where"], use=dict(use))
    return dict(
        meta=dict(source=src_name, generated=datetime.now(ZoneInfo("Asia/Taipei")).strftime("%Y-%m-%d %H:%M"), **s.solve_info,
                  n_teachers=len(teachers), n_classes=len(classes), n_lessons=len(s.lessons)),
        days=DAYS, periods=PERIODS, classes=classes, teachers=teachers, rooms=rooms,
        rules=check(s), issues=s.issues, relaxed=getattr(s, "relaxed_hits", []))


def run(data, src_name, time_limit=60, log=print):
    """完整排課流程：事前檢查 → 排課 → 排不出來就放寬重排 → 規則體檢。回傳結果 dict。"""
    s = Scheduler(data)
    s.precheck()
    for i in s.issues:
        log(f"[{i['level']}] {i['text']}")
    ok = s.solve(time_limit)
    if not ok:
        log("照規則排不出來，改用「放寬老師時段」再試一次，找出卡住的地方…")
        s.issues.append(dict(level="錯誤", text="照全部規則排不出來；下面的課表是放寬「老師不排時段（必須）」與「職務」後的結果；被迫排進去的格子列在規則體檢的 R01、R12。"))
        ok = s.solve(time_limit, relax=True)
    log(s.solve_info)
    if not ok:
        return dict(meta=dict(source=src_name, **s.solve_info), issues=s.issues)
    result = build_output(s, src_name)
    for r in result["rules"]:
        mark = {True: "✅", False: "⚠️", None: "—"}[r["ok"]]
        log(f"{mark} {r['id']} {r['block']}｜{r['summary']}")
    return result


def main():
    src = Path(sys.argv[1])
    out_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else src.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    result = run(load(src), src.name)
    (out_dir / "排課結果.json").write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")
    return result


if __name__ == "__main__":
    main()
