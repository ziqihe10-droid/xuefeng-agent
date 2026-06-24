#!/usr/bin/env python3
"""雪峰Agent 单文件服务端：HTML UI、推荐数据、联网补充与流式对话编排。"""

import gzip
import json
import os
import re
import shutil
import sqlite3
import sys
import urllib.error
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(HERE, "admission_clean.db")
GZ_PATH = os.path.join(HERE, "admission_clean.db.gz")
USER_XLSX = os.path.join(HERE, "自定义数据.xlsx")

if not os.path.exists(DB_PATH) and os.path.exists(GZ_PATH):
    with gzip.open(GZ_PATH, "rb") as gz:
        with open(DB_PATH, "wb") as db_file:
            shutil.copyfileobj(gz, db_file)

HAS_DB = os.path.exists(DB_PATH)
USER_DATA = []

PROVINCES = [
    "北京", "天津", "上海", "重庆", "河北", "山西", "辽宁", "吉林", "黑龙江", "江苏", "浙江", "安徽",
    "福建", "江西", "山东", "河南", "湖北", "湖南", "广东", "广西", "海南", "四川", "贵州", "云南",
    "西藏", "陕西", "甘肃", "青海", "宁夏", "新疆", "内蒙古",
]

MAJORS = [
    "计算机", "软件", "人工智能", "大数据", "电气", "机械", "自动化", "电子", "通信", "物联网",
    "土木", "临床", "口腔", "法学", "会计", "金融", "材料", "化工", "生物", "医学", "护理",
    "师范", "英语", "日语", "新闻", "设计", "美术", "音乐", "体育", "汉语言", "思政", "数学",
    "化学", "地理", "航空航天", "能源", "交通", "环境",
]

MODEL_ALIASES = {
    "": "deepseek-v4-flash",
    "deepseek-chat": "deepseek-v4-flash",
    "deepseek-reason": "deepseek-v4-pro",
    "deepseek-reasoner": "deepseek-v4-pro",
}

SUBJECT_DETAIL_KEYWORDS = ["物理", "历史", "化学", "生物", "政治", "地理"]

GAOKAO_PRIORITY = [
    "province",
    "score_or_rank",
    "subject",
    "majors",
    "accept_adjustment",
    "region_pref",
    "career_goal",
    "budget",
]
FUN_PRIORITY = ["topic", "goal", "style", "guardrails"]

GAOKAO_REQUIRED_FIELDS = ["province", "score_or_rank", "subject", "majors", "accept_adjustment"]
GAOKAO_OPTIONAL_FIELDS = ["region_pref", "career_goal", "budget"]
FUN_REQUIRED_FIELDS = ["topic", "goal"]
FUN_OPTIONAL_FIELDS = ["style", "guardrails"]

FIELD_MAX_ATTEMPTS = {
    "province": 2,
    "score_or_rank": 2,
    "subject": 2,
    "majors": 1,
    "accept_adjustment": 2,
    "region_pref": 1,
    "career_goal": 1,
    "budget": 1,
    "topic": 1,
    "goal": 1,
    "style": 1,
    "guardrails": 1,
}

GAOKAO_QUESTION_TEXT = {
    "province": "你是哪个省的考生？",
    "score_or_rank": "告诉我你的分数，或者你的位次，给一个就行。",
    "subject": "你是物理类、历史类，还是像浙江这种综合改革省份？",
    "majors": "你想学什么专业？如果有明确不想学的，也可以一起说。",
    "accept_adjustment": "你能不能接受调剂？能接受、不能接受，直接说就行。",
    "region_pref": "你更想去哪里读书？如果有不想去的城市或省份，也可以顺手告诉我。",
    "career_goal": "你更偏向考研、直接就业，还是考公考编？",
    "budget": "学费大概能接受什么范围？如果家里比较看重性价比，也可以直接说。",
}

GAOKAO_UNKNOWN_MAJOR_PATTERNS = [
    "不知道想学什么",
    "不知道学什么",
    "没想好专业",
    "专业还没想好",
    "我也不知道",
    "还不知道",
    "暂时不知道",
    "没方向",
]

FUN_QUESTION_TEXT = {
    "topic": "你这次主要想聊什么？比如高考、专业、学校、就业，还是单纯想听我点评一下。",
    "goal": "你最想让我帮你做到哪一步？是分析、推荐、吐槽，还是帮你拿主意。",
    "style": "你希望我怎么说？直接一点、温和一点，还是带点雪峰式吐槽。",
    "guardrails": "有没有明确不想碰的话题，或者你不想听的方向？",
}

OFFICIAL_SITES = {
    "北京": "bjeea.cn",
    "天津": "zhaokao.net",
    "河北": "hebeea.edu.cn",
    "山西": "sxkszx.cn",
    "内蒙古": "nm.zsks.cn",
    "辽宁": "lnzsks.com",
    "吉林": "jleea.edu.cn",
    "黑龙江": "lzk.hl.cn",
    "上海": "shmeea.edu.cn",
    "江苏": "jseea.cn",
    "浙江": "zjzs.net",
    "安徽": "ahzsks.cn",
    "福建": "eeafj.cn",
    "江西": "jxeea.cn",
    "山东": "sdzk.cn",
    "河南": "haeea.cn",
    "湖北": "hbea.edu.cn",
    "湖南": "hneeb.cn",
    "广东": "eeagd.edu.cn",
    "广西": "gxeea.cn",
    "海南": "hainanu.edu.cn",
    "重庆": "cqksy.cn",
    "四川": "sceea.cn",
    "贵州": "zsksy.guizhou.gov.cn",
    "云南": "ynzs.cn",
    "西藏": "zsks.edu.xizang.gov.cn",
    "陕西": "sneea.cn",
    "甘肃": "ganseea.cn",
    "青海": "qhjyks.com",
    "宁夏": "nxjyks.cn",
    "新疆": "xjzk.gov.cn",
}

GAOKAO_SYSTEM_PROMPT = """现在是2026年6月，高考刚结束，志愿填报正在进行。
你是雪峰Agent的报考顾问，面对的是几乎没接触过计算机行业的零基础用户。
回复要求：
1. 只用纯文本，不要 Markdown、标题符号、代码块、表格。
2. 先说结论，再说理由，最后告诉用户下一步怎么做。
3. 句子短一点，尽量像正常说话，不要写成技术说明书。
4. 没有数据就明确说没有，不准编造分数、位次、录取结果。
5. 如果信息还不够，就基于现有信息给方向，同时提醒补什么会更准。
6. 如果系统给了本地数据库结果，以本地结果为主；联网信息只做补充参考。
7. 如果本地数据库没有明确命中，就不要直接点名某所学校说“稳”“冲”“大概率能上”。只能明确告诉用户：当前本地数据不够，下面只能给方向性建议，具体学校要等官方位次或补充数据后再判断。"""

FUN_SYSTEM_PROMPT = """你是雪峰Agent里的娱乐聊天顾问，语气可以直接、接地气，但要让零基础用户也能轻松看懂。
回复要求：
1. 只用纯文本，不要 Markdown、代码块、表格。
2. 先直接回答用户想听的，再补一句理由或提醒。
3. 不要装技术，不要堆概念，不要故意复杂化。
4. 如果信息太少，先补问最关键的1到2条。"""

UNIFIED_EXTRACT_PROMPT = """你是一个只负责信息提取的助手。你的任务是把用户消息和已有上下文整理成结构化 JSON。

输出规则：
1. 只能输出一个 JSON 对象，不能输出解释，不能输出 Markdown。
2. 结合 mode、current_questions、collected、history 来理解用户这一轮输入。
3. 如果用户没有否定旧信息，就不要清空旧信息。
4. 如果用户明确表示“不知道想学什么”“专业没想好”“我也不知道学什么”，请返回 major_unknown=true，并把 majors 设为空数组。
5. 如果用户只是说“我也不知道”，要结合 current_questions 判断是在回答哪个问题。
6. accept_adjustment 只能是 true、false 或 null。
7. score 和 rank 只填数字，没有就填 0。
8. subject 只允许：物理类、历史类、综合。
9. career_goal 尽量归一为：考研、直接就业、考公、考编；无法判断就留空。
10. budget 保留成适合展示的简短中文，比如“正常预算”“预算偏谨慎”“不太高就行”。
11. majors、schools、region_pref、region_avoid、guardrails 必须是数组。
12. thinking_mode 固定返回 false。

JSON 模板：
{"province":"","score":0,"rank":0,"subject":"","majors":[],"major_unknown":false,"schools":[],"accept_adjustment":null,"region_pref":[],"region_avoid":[],"career_goal":"","budget":"","topic":"","goal":"","style":"","guardrails":[],"thinking_mode":false}
"""


def clean_num(value):
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "").replace("分", "").replace("位", "").replace("名", "").replace(" ", "")
    try:
        return int(float(text))
    except Exception:
        return None


def normalize_model_name(model):
    model = (model or "").strip()
    return MODEL_ALIASES.get(model, model or "deepseek-v4-flash")


def unique_list(values):
    seen = set()
    items = []
    for value in values or []:
        if not value:
            continue
        if value in seen:
            continue
        seen.add(value)
        items.append(value)
    return items


def safe_json_loads(text, fallback=None):
    try:
        return json.loads(text)
    except Exception:
        return fallback


def load_user_data():
    global USER_DATA
    USER_DATA = []
    if not os.path.exists(USER_XLSX):
        return
    try:
        import openpyxl
    except ImportError:
        return

    try:
        wb = openpyxl.load_workbook(USER_XLSX, data_only=True)
        for row in wb.active.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            school = str(row[0]).strip()
            if len(school) < 2 or school in ["学校名称", "院校名称"]:
                continue
            note = str(row[8]).strip() if len(row) > 8 and row[8] else ""
            if "示例" in note or "不参与排序" in note:
                continue
            major = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            category_raw = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            if "物理" in category_raw:
                category = "物理类"
            elif "历史" in category_raw:
                category = "历史类"
            elif "综合" in category_raw:
                category = "综合"
            else:
                category = category_raw
            province = str(row[3]).strip() if len(row) > 3 and row[3] else ""
            province = province.replace("省", "").replace("市", "").strip()
            score_24 = clean_num(row[4] if len(row) > 4 else None)
            rank_24 = clean_num(row[5] if len(row) > 5 else None)
            score_25 = clean_num(row[6] if len(row) > 6 else None)
            rank_25 = clean_num(row[7] if len(row) > 7 else None)
            if score_24 and score_24 < 100 and rank_24 and rank_24 > 300:
                score_24, rank_24 = rank_24, score_24
            if score_25 and score_25 < 100 and rank_25 and rank_25 > 300:
                score_25, rank_25 = rank_25, score_25
            if score_24 and rank_24:
                USER_DATA.append({
                    "school": school,
                    "major": major,
                    "year": 2024,
                    "category": category,
                    "score": score_24,
                    "rank": rank_24,
                    "province": province,
                })
            if score_25 and rank_25:
                USER_DATA.append({
                    "school": school,
                    "major": major,
                    "year": 2025,
                    "category": category,
                    "score": score_25,
                    "rank": rank_25,
                    "province": province,
                })
        wb.close()
    except Exception as exc:
        print(f"[user data] {exc}")


load_user_data()


def parse_wan_number(text):
    match = re.search(r"(\d+(?:\.\d+)?)\s*万\s*(\d{0,4})", text)
    if not match:
        return None
    head = int(float(match.group(1)) * 10000)
    tail = match.group(2) or ""
    if not tail:
        return head
    if len(tail) == 1:
        return head + int(tail) * 1000
    if len(tail) == 2:
        return head + int(tail) * 100
    if len(tail) == 3:
        return head + int(tail) * 10
    return head + int(tail)


def parse_w_suffix_number(text):
    match = re.search(r"(\d+(?:\.\d+)?)\s*[wW]\s*(\d{0,4})", text)
    if not match:
        return None
    head = int(float(match.group(1)) * 10000)
    tail = match.group(2) or ""
    if not tail:
        return head
    if len(tail) == 1:
        return head + int(tail) * 1000
    if len(tail) == 2:
        return head + int(tail) * 100
    if len(tail) == 3:
        return head + int(tail) * 10
    return head + int(tail)


def extract_rank(text):
    patterns = [
        r"(?:位次|排名|排位|省排|名次)[^\d]{0,4}(\d{4,7})",
        r"(\d{4,7})\s*(?:位次|排名|排位|省排|名次)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return int(match.group(1))

    wan_patterns = [
        r"(?:位次|排名|排位|省排|名次)[^\d万]{0,4}(\d+(?:\.\d+)?\s*万\s*\d{0,4})",
        r"(\d+(?:\.\d+)?\s*万\s*\d{0,4})\s*(?:位次|排名|排位|省排|名次)",
        r"(\d+(?:\.\d+)?)\s*[wW](?:左右|上下)?",
        r"(\d+(?:\.\d+)?)\s*万(?:左右|上下)?",
    ]
    for pattern in wan_patterns:
        match = re.search(pattern, text)
        if match:
            raw = match.group(1)
            value = parse_wan_number(raw) or parse_w_suffix_number(raw)
            if value:
                return value

    loose_match = re.search(r"(\d+(?:\.\d+)?)\s*[wW](?:左右|上下)?", text)
    if loose_match:
        value = parse_w_suffix_number(loose_match.group(0))
        if value:
            return value

    loose_wan_match = re.search(r"(\d+(?:\.\d+)?)\s*万(?:左右|上下)?", text)
    if loose_wan_match:
        value = parse_wan_number(loose_wan_match.group(0))
        if value:
            return value
    return 0


def extract_score(text):
    patterns = [
        r"(\d{3})\s*分",
        r"分数[^\d]{0,3}(\d{3})",
        r"考了[^\d]{0,3}(\d{3})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return int(match.group(1))

    standalone = re.findall(r"(?<!\d)(\d{3})(?!\d)", text)
    for item in standalone:
        value = int(item)
        if 200 <= value <= 750:
            return value
    return 0


def extract_subject(text):
    subject_map = {
        "物化地": "物理类",
        "物化政": "物理类",
        "物化生": "物理类",
        "物理类": "物理类",
        "理科": "物理类",
        "史政地": "历史类",
        "史地政": "历史类",
        "历史类": "历史类",
        "文科": "历史类",
        "综合改革": "综合",
        "综合类": "综合",
        "综合": "综合",
    }
    for keyword, subject in subject_map.items():
        if keyword in text:
            return subject
    return ""


def extract_subject_detail(text):
    found = []
    for subject in SUBJECT_DETAIL_KEYWORDS:
        if subject in (text or ""):
            found.append(subject)
    return unique_list(found)


def extract_requirement_subjects(major_name):
    found = []
    for subject in SUBJECT_DETAIL_KEYWORDS:
        if subject in (major_name or ""):
            found.append(subject)
    return unique_list(found)


def requirement_is_unlimited(major_name):
    return "不限" in (major_name or "")


def is_major_group_text(major_name):
    text = (major_name or "").strip()
    return bool(text) and not any(major in text for major in MAJORS)


def subject_requirement_compatible(subject, subject_detail, major_name):
    major_name = major_name or ""
    if not major_name:
        return True
    if requirement_is_unlimited(major_name):
        return True

    required = extract_requirement_subjects(major_name)
    if not required:
        return True

    detail_set = set(subject_detail or [])
    if detail_set:
        return all(item in detail_set for item in required)

    if subject in {"历史类", "物理类"}:
        return False
    return True


def extract_province(text):
    best_index = len(text) + 1
    best_province = ""
    for province in PROVINCES:
        index = text.find(province)
        if index >= 0 and index < best_index:
            best_index = index
            best_province = province
    return best_province


def extract_majors(text):
    negative_blocks = re.findall(r"(?:不学|不考虑|拒绝|排斥|不接受|别推荐|不想学).*?(?:[。，,；;\n]|$)", text)
    negative_text = "".join(negative_blocks)
    found = []
    for major in MAJORS:
        if major in text and major not in negative_text:
            found.append(major)
    return unique_list(found)


def extract_region_preferences(text):
    region_pref = []
    region_avoid = []
    if "省内" in text or "本省" in text:
        region_pref.append("省内")
    if "省外" in text:
        region_pref.append("省外")
    pref_patterns = [r"(?:想去|想在|优先|最好在)([^。；;\n]*)", r"(?:城市想去|地域倾向)([^。；;\n]*)"]
    avoid_patterns = [r"(?:不想去|不去|避开|别去)([^。；;\n]*)", r"(?:地域避开|不考虑)([^。；;\n]*)"]
    for pattern in pref_patterns:
        for block in re.findall(pattern, text):
            for province in PROVINCES:
                if province in block:
                    region_pref.append(province)
    for pattern in avoid_patterns:
        for block in re.findall(pattern, text):
            for province in PROVINCES:
                if province in block:
                    region_avoid.append(province)
    return unique_list(region_pref), unique_list(region_avoid)


def extract_career_goal(text):
    for keyword, value in [
        ("考研", "考研"),
        ("读研", "考研"),
        ("直接就业", "直接就业"),
        ("找工作", "直接就业"),
        ("就业", "直接就业"),
        ("考公", "考公"),
        ("考编", "考编"),
        ("律师", "法律职业"),
        ("司法", "法律职业"),
    ]:
        if keyword in text:
            return value
    return ""


def extract_budget_text(text, contextual=False):
    budget_match = re.search(r"(?:学费|预算|费用)[^。；;\n]{0,12}?(\d{4,6})", text)
    if budget_match:
        return budget_match.group(1)
    if "中外合作" in text and ("不考虑" in text or "慎重" in text):
        return "预算偏谨慎"
    if "正常价格" in text or "正常预算" in text or "普通价格" in text:
        return "正常预算"
    if "不太高" in text or "便宜点" in text or "性价比" in text:
        return "预算偏谨慎"
    if "无限" in text or "不限" in text:
        return "预算充足"
    if contextual:
        plain = (text or "").strip()
        if plain and len(plain) <= 20:
            return plain
    return ""


def parse_accept_adjustment(text, contextual=False):
    plain = (text or "").strip()
    compact = re.sub(r"\s+", "", plain)
    compact_plain = re.sub(r"[，。、“”‘’！？!?,；;：:\-（）()\[\]【】/\\]", "", compact)
    if not compact:
        return None
    if re.search(r"(不接受|不能接受|不服从|拒绝|不考虑).{0,6}(调剂|调解|调配)", compact_plain) or compact_plain in {"不能", "不接受", "不服从", "不可以"}:
        return False
    if re.search(r"(接受|可以|能|服从).{0,6}(调剂|调解|调配)", compact_plain) or "服从调剂" in compact_plain:
        return True
    if not contextual:
        return None
    positive = {"能", "可以", "可以的", "能接受", "接受", "接受的", "服从", "行", "好的", "好"}
    negative = {"不能", "不行", "不可以", "不接受", "不接受的", "不能接受", "不服从", "拒绝", "不要", "否"}
    if compact_plain in positive:
        return True
    if compact_plain in negative:
        return False
    if any(token in compact_plain for token in ["不能接受", "不接受", "不服从", "拒绝", "不考虑", "不调剂"]):
        return False
    if any(token in compact_plain for token in ["能接受", "可以接受", "可以调剂", "接受调剂", "服从调剂"]):
        return True
    if compact_plain.startswith(("可以", "能接受", "接受", "服从")):
        return True
    return None


def contextual_extract_fields(mode, text, current_questions):
    current_questions = current_questions or []
    if not current_questions or not (text or "").strip():
        return {}

    compact = re.sub(r"\s+", "", text.strip())
    result = {}

    if mode != "gaokao":
        for field in current_questions:
            if field == "guardrails" and compact:
                result["guardrails"] = [text.strip()]
            elif field in {"topic", "goal", "style"} and compact and compact not in {"不知道", "我也不知道", "没想好"}:
                result[field] = text.strip()
        return result

    for field in current_questions:
        if field == "province":
            province = extract_province(text)
            if province:
                result["province"] = province
        elif field == "score_or_rank":
            score = extract_score(text)
            rank = extract_rank(text)
            if score:
                result["score"] = score
            if rank:
                result["rank"] = rank
        elif field == "subject":
            subject = extract_subject(text)
            if subject:
                result["subject"] = subject
        elif field == "majors":
            majors = extract_majors(text)
            if majors:
                result["majors"] = majors
            elif infer_unknown_major(text) or compact in {"不知道", "我也不知道", "没想好", "还没想好", "都行", "随便"}:
                result["majors"] = []
                result["major_unknown"] = True
        elif field == "accept_adjustment":
            accept_adjustment = parse_accept_adjustment(text, contextual=True)
            if accept_adjustment is not None:
                result["accept_adjustment"] = accept_adjustment
        elif field == "region_pref":
            region_pref, region_avoid = extract_region_preferences(text)
            if region_pref:
                result["region_pref"] = region_pref
            elif compact in {"省内", "本省", "省外", "外省"}:
                result["region_pref"] = [compact]
            if region_avoid:
                result["region_avoid"] = region_avoid
        elif field == "career_goal":
            career_goal = extract_career_goal(text)
            if career_goal:
                result["career_goal"] = career_goal
        elif field == "budget":
            budget = extract_budget_text(text, contextual=True)
            if budget:
                result["budget"] = budget
    return result


def extract_info_from_text(text):
    schools = re.findall(r"[\u4e00-\u9fa5]{2,12}(?:大学|学院)", text)
    region_pref, region_avoid = extract_region_preferences(text)

    career_goal = extract_career_goal(text)
    budget = extract_budget_text(text)
    accept_adjustment = parse_accept_adjustment(text)

    style = ""
    for keyword in ["直接一点", "温和一点", "毒舌", "幽默", "轻松", "严肃"]:
        if keyword in text:
            style = keyword
            break

    goal = ""
    if re.search(r"(推荐|分析|规划|拿主意|解释|吐槽|比较|帮我看看)", text):
        goal = "需要明确建议"

    guardrails = []
    for pattern in re.findall(r"(?:不要|别|避免|不能).{0,18}", text):
        guardrails.append(pattern.strip())

    topic = ""
    if any(keyword in text for keyword in ["志愿", "高考", "专业", "学校", "就业"]):
        topic = "升学相关"
    elif len(text.strip()) >= 8:
        topic = text.strip()[:18]

    return {
        "province": extract_province(text),
        "rank": extract_rank(text),
        "score": extract_score(text),
        "subject": extract_subject(text),
        "majors": extract_majors(text),
        "major_unknown": False,
        "schools": unique_list(schools),
        "accept_adjustment": accept_adjustment,
        "region_pref": region_pref,
        "region_avoid": region_avoid,
        "career_goal": career_goal,
        "budget": budget,
        "topic": topic,
        "goal": goal,
        "style": style,
        "guardrails": unique_list(guardrails),
    }


def fallback_extract_info(mode, text, current_questions=None):
    current_questions = current_questions or []
    base = extract_info_from_text(text)
    contextual = contextual_extract_fields(mode, text, current_questions)
    if mode != "gaokao":
        result = {
            "topic": base.get("topic", ""),
            "goal": base.get("goal", ""),
            "style": base.get("style", ""),
            "guardrails": base.get("guardrails", []),
        }
        result.update(contextual)
        if current_questions:
            focused = {}
            for field in current_questions:
                value = result.get(field)
                if isinstance(value, list) and value:
                    focused[field] = value
                elif isinstance(value, str) and value.strip():
                    focused[field] = value.strip()
            if focused:
                return focused
        return result

    result = {
        "province": base.get("province", ""),
        "score": base.get("score", 0),
        "rank": base.get("rank", 0),
        "subject": base.get("subject", ""),
        "majors": base.get("majors", []),
        "major_unknown": infer_unknown_major(text),
        "schools": base.get("schools", []),
        "accept_adjustment": base.get("accept_adjustment"),
        "region_pref": base.get("region_pref", []),
        "region_avoid": base.get("region_avoid", []),
        "career_goal": base.get("career_goal", ""),
        "budget": base.get("budget", ""),
    }
    result = merge_collected(result, contextual)
    if current_questions:
        focused = {}
        for field in current_questions:
            if field == "score_or_rank":
                if result.get("score"):
                    focused["score"] = result["score"]
                if result.get("rank"):
                    focused["rank"] = result["rank"]
                continue
            value = result.get(field)
            if isinstance(value, list) and value:
                focused[field] = value
            elif isinstance(value, str) and value.strip():
                focused[field] = value.strip()
            elif field == "accept_adjustment" and value is not None:
                focused[field] = value
        if result.get("major_unknown"):
            focused["major_unknown"] = True
        if focused:
            return focused
    return result


def infer_unknown_major(text):
    content = (text or "").strip()
    if not content:
        return False
    return any(pattern in content for pattern in GAOKAO_UNKNOWN_MAJOR_PATTERNS)


def normalize_extracted_fields(mode, parsed):
    parsed = parsed if isinstance(parsed, dict) else {}
    defaults = get_default_collected(mode)
    normalized = {}

    if "province" in parsed and isinstance(parsed.get("province"), str):
        normalized["province"] = parsed.get("province", "").strip()
    if "score" in parsed:
        normalized["score"] = clean_num(parsed.get("score")) or 0
    if "rank" in parsed:
        normalized["rank"] = clean_num(parsed.get("rank")) or 0
    if "subject" in parsed and isinstance(parsed.get("subject"), str):
        subject = parsed.get("subject", "").strip()
        if subject in {"物理类", "历史类", "综合"}:
            normalized["subject"] = subject
    if "majors" in parsed and isinstance(parsed.get("majors"), list):
        normalized["majors"] = unique_list([str(item).strip() for item in parsed.get("majors") or [] if str(item).strip()])
    if "major_unknown" in parsed:
        normalized["major_unknown"] = bool(parsed.get("major_unknown"))
    if "schools" in parsed and isinstance(parsed.get("schools"), list):
        normalized["schools"] = unique_list([str(item).strip() for item in parsed.get("schools") or [] if str(item).strip()])
    if "accept_adjustment" in parsed:
        value = parsed.get("accept_adjustment")
        normalized["accept_adjustment"] = value if isinstance(value, bool) or value is None else None
    for key in ["region_pref", "region_avoid", "guardrails"]:
        if key in parsed and isinstance(parsed.get(key), list):
            normalized[key] = unique_list([str(item).strip() for item in parsed.get(key) or [] if str(item).strip()])
    for key in ["career_goal", "budget", "topic", "goal", "style"]:
        if key in parsed and isinstance(parsed.get(key), str):
            normalized[key] = parsed.get(key, "").strip()

    filtered = {}
    for key in defaults:
        if key in normalized:
            filtered[key] = normalized[key]
    return filtered


def ai_extract_fields(mode, message, history, current_questions, collected, config):
    api_key = (config.get("key") or "").strip()
    if not api_key or not (message or "").strip():
        return {}

    api_url = (config.get("url") or "https://api.deepseek.com").strip().rstrip("/")
    model = normalize_model_name(config.get("model"))
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": UNIFIED_EXTRACT_PROMPT},
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "mode": mode,
                        "message": message,
                        "current_questions": current_questions or [],
                        "collected": collected or {},
                        "history": history[-8:] if isinstance(history, list) else [],
                    },
                    ensure_ascii=False,
                ),
            },
        ],
        "temperature": 0,
        "max_tokens": 320,
        "thinking": {"type": "disabled"},
    }
    request = urllib.request.Request(
        api_url + "/v1/chat/completions",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=35) as response:
        data = json.loads(response.read().decode("utf-8"))
    content = ((data.get("choices") or [{}])[0].get("message") or {}).get("content", "")
    content = content.replace("```json", "").replace("```", "").strip()
    parsed = safe_json_loads(content, {})
    return normalize_extracted_fields(mode, parsed)


def get_default_collected(mode):
    if mode == "gaokao":
        return {
            "province": "",
            "rank": 0,
            "score": 0,
            "subject": "",
            "majors": [],
            "major_unknown": False,
            "schools": [],
            "accept_adjustment": None,
            "region_pref": [],
            "region_avoid": [],
            "career_goal": "",
            "budget": "",
        }
    return {
        "topic": "",
        "goal": "",
        "style": "",
        "guardrails": [],
    }


def merge_collected(current, incoming):
    merged = dict(current or {})
    for key, value in (incoming or {}).items():
        if value is None:
            continue
        if isinstance(value, list):
            if value:
                merged[key] = unique_list((merged.get(key) or []) + value)
            elif key not in merged:
                merged[key] = []
        elif isinstance(value, bool):
            merged[key] = value
        elif isinstance(value, int):
            if value > 0:
                merged[key] = value
        elif isinstance(value, str):
            if value.strip():
                merged[key] = value.strip()
        else:
            merged[key] = value
    return merged


def merge_collected_for_mode(mode, current, incoming):
    allowed = set(get_default_collected(mode).keys())
    base = {}
    for key, value in (current or {}).items():
        if key in allowed:
            base[key] = value
    filtered = {}
    for key, value in (incoming or {}).items():
        if key in allowed:
            filtered[key] = value
    return merge_collected(base, filtered)


def ensure_plan_state(mode, planner):
    planner = planner or {}
    state = planner.get("planState") or {}
    collected = get_default_collected(mode)
    collected = merge_collected_for_mode(mode, collected, state.get("collected") or {})
    asked_fields = unique_list(state.get("asked_fields") or [])
    current_questions = state.get("current_questions") or []
    asked_count = state.get("asked_count") or {}
    return {
        "mode": mode,
        "collected": collected,
        "asked_fields": asked_fields,
        "asked_count": {str(key): int(value) for key, value in asked_count.items() if str(key)},
        "current_questions": current_questions,
        "last_user_message": state.get("last_user_message") or "",
        "ready": bool(state.get("ready")),
    }


def sanitize_plan_state(mode, state):
    state = dict(state or {})
    defaults = get_default_collected(mode)
    collected = state.get("collected") or {}
    clean_collected = {}
    for key, default_value in defaults.items():
        value = collected.get(key, default_value)
        if isinstance(default_value, list):
            clean_collected[key] = unique_list(value or [])
        elif isinstance(default_value, bool):
            clean_collected[key] = bool(value)
        else:
            clean_collected[key] = value
    return {
        "mode": mode,
        "collected": clean_collected,
        "asked_fields": unique_list(state.get("asked_fields") or []),
        "asked_count": {
            str(key): max(0, int(value))
            for key, value in (state.get("asked_count") or {}).items()
            if str(key)
        },
        "current_questions": state.get("current_questions") or [],
        "last_user_message": state.get("last_user_message") or "",
        "ready": bool(state.get("ready")),
    }


def determine_missing_fields(mode, collected):
    if mode == "gaokao":
        missing = []
        if not collected.get("province"):
            missing.append("province")
        if not (collected.get("score") or collected.get("rank")):
            missing.append("score_or_rank")
        if not collected.get("subject"):
            missing.append("subject")
        if not collected.get("majors") and not collected.get("major_unknown"):
            missing.append("majors")
        if collected.get("accept_adjustment") is None:
            missing.append("accept_adjustment")
        if not collected.get("region_pref") and not collected.get("region_avoid"):
            missing.append("region_pref")
        if not collected.get("career_goal"):
            missing.append("career_goal")
        if not collected.get("budget"):
            missing.append("budget")
        return missing

    missing = []
    if not collected.get("topic"):
        missing.append("topic")
    if not collected.get("goal"):
        missing.append("goal")
    if not collected.get("style"):
        missing.append("style")
    if not collected.get("guardrails"):
        missing.append("guardrails")
    return missing


def prune_answered_questions(mode, state):
    state = sanitize_plan_state(mode, state)
    missing = set(determine_missing_fields(mode, state["collected"]))
    state["current_questions"] = [field for field in (state.get("current_questions") or []) if field in missing]
    return state


def required_fields_for_mode(mode):
    return GAOKAO_REQUIRED_FIELDS if mode == "gaokao" else FUN_REQUIRED_FIELDS


def optional_fields_for_mode(mode):
    return GAOKAO_OPTIONAL_FIELDS if mode == "gaokao" else FUN_OPTIONAL_FIELDS


def should_skip_field(mode, state, field):
    attempts = int((state.get("asked_count") or {}).get(field, 0))
    limit = FIELD_MAX_ATTEMPTS.get(field, 1)
    if attempts < limit:
        return False
    return field in optional_fields_for_mode(mode) or field == "majors"


def planner_should_ask(mode, plan_enabled, state, message):
    if not plan_enabled:
        return False
    missing = determine_missing_fields(mode, state["collected"])
    effective_missing = [field for field in missing if not should_skip_field(mode, state, field)]
    if mode == "gaokao":
        core_missing = [field for field in required_fields_for_mode(mode) if field in effective_missing]
        optional_missing = [field for field in optional_fields_for_mode(mode) if field in effective_missing]
        return bool(core_missing or optional_missing)
    stripped = message.strip()
    if len(stripped) < 8:
        return bool(effective_missing)
    return bool(effective_missing and not state["collected"].get("goal"))


def select_questions(mode, state):
    priority = GAOKAO_PRIORITY if mode == "gaokao" else FUN_PRIORITY
    missing = determine_missing_fields(mode, state["collected"])
    available = [field for field in missing if not should_skip_field(mode, state, field)]
    selected = []
    for field in priority:
        if field not in available:
            continue
        if field in state["asked_fields"]:
            continue
        selected.append(field)
        if len(selected) >= 2:
            break
    if not selected:
        for field in priority:
            if field in available:
                selected.append(field)
                if len(selected) >= 1:
                    break
    return selected


def build_question_text(mode, fields):
    mapping = GAOKAO_QUESTION_TEXT if mode == "gaokao" else FUN_QUESTION_TEXT
    intro = "先补两条最关键的信息，我再继续给你排。" if len(fields) >= 2 else "先补一条最关键的信息，我再继续。"
    if mode == "fun":
        intro = "我先问你一两句，这样后面的回答会更贴你的意思。"
    lines = [intro]
    for index, field in enumerate(fields, start=1):
        lines.append(f"{index}. {mapping[field]}")
    return "\n".join(lines)


def build_question_response(mode, state):
    state = sanitize_plan_state(mode, state)
    fields = select_questions(mode, state)
    state["asked_fields"] = unique_list(state["asked_fields"] + fields)
    asked_count = dict(state.get("asked_count") or {})
    for field in fields:
        asked_count[field] = int(asked_count.get(field, 0)) + 1
    state["asked_count"] = asked_count
    state["current_questions"] = fields
    state["ready"] = False
    return {
        "type": "question",
        "text": build_question_text(mode, fields),
        "planState": sanitize_plan_state(mode, state),
        "missing_fields": determine_missing_fields(mode, state["collected"]),
    }


def recommend_by_user_data(province):
    matched = [row for row in USER_DATA if row.get("province") and row["province"] in province and row.get("score") and row.get("rank")]
    if not matched:
        return None
    grouped = {}
    for row in matched:
        key = f"{row['school']}|{row.get('major', '')}"
        entry = grouped.setdefault(key, {
            "school": row["school"],
            "major": row.get("major", ""),
            "scores": [],
            "ranks": [],
            "years": [],
        })
        entry["scores"].append(row["score"])
        entry["ranks"].append(row["rank"])
        entry["years"].append(row["year"])
    all_rows = []
    for entry in grouped.values():
        score_avg = int(sum(entry["scores"]) / len(entry["scores"]))
        rank_avg = int(sum(entry["ranks"]) / len(entry["ranks"]))
        detail_parts = []
        year_pairs = sorted(zip(entry["years"], entry["scores"], entry["ranks"]))
        for year, score, rank in year_pairs:
            detail_parts.append(f"{year}:{score}分 {rank}位")
        major_text = entry["major"]
        if detail_parts:
            major_text = f"{major_text} [{' '.join(detail_parts)}]".strip()
        all_rows.append({
            "school": entry["school"],
            "major": major_text,
            "score": score_avg,
            "rank": rank_avg,
            "year": "综合",
            "source": "user",
        })
    all_rows.sort(key=lambda row: row["rank"])
    size = len(all_rows)
    split_1 = max(1, size // 3)
    split_2 = max(split_1 + 1, (size * 2) // 3)
    return {
        "rank": 0,
        "score": 0,
        "chong": all_rows[:split_1],
        "wen": all_rows[split_1:split_2],
        "bao": all_rows[split_2:],
        "source": "custom_only",
    }


def recommend_data(province="", major="", keyword="", rank=0, score=0, subject=""):
    if province:
        custom = recommend_by_user_data(province)
        if custom:
            custom["rank"] = rank
            custom["score"] = score
            return custom

    if not HAS_DB or not province or not (rank > 0 or score > 0):
        return {"rank": rank, "score": score, "chong": [], "wen": [], "bao": [], "source": "empty"}

    conn = sqlite3.connect(DB_PATH)
    base = "province LIKE ? AND (score>0 OR rank>0)"
    params = [f"%{province}%"]
    subject_detail = extract_subject_detail(subject)
    if subject:
        base += " AND (category=? OR category='' OR category IS NULL)"
        params.append(subject)
    if major:
        base += " AND major_name LIKE ?"
        params.append(f"%{major}%")
    if keyword:
        keywords = [item for item in keyword.split(",") if item]
        if keywords:
            keyword_conditions = []
            for item in keywords:
                keyword_conditions.append("(major_name LIKE ? OR school_name LIKE ?)")
                params.append(f"%{item}%")
                params.append(f"%{item}%")
            base += " AND (" + " OR ".join(keyword_conditions) + ")"

    def rows_from_sql(sql, sql_params):
        rows = []
        for row in conn.execute(sql, sql_params).fetchall():
            item = {
                "school": row[0],
                "major": row[1],
                "score": row[2],
                "rank": row[3],
                "year": row[4],
                "source": "db",
            }
            if subject and not subject_requirement_compatible(subject, subject_detail, item["major"]):
                continue
            rows.append(item)
        return rows

    def filter_rows(rows):
        filtered = []
        for item in rows:
            if subject and not subject_requirement_compatible(subject, subject_detail, item["major"]):
                continue
            filtered.append(item)
        return filtered

    chong = []
    wen = []
    bao = []

    if rank > 0:
        chong = rows_from_sql(
            f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {base} "
            "AND rank>0 AND rank<? AND rank>=? ORDER BY rank ASC LIMIT 50",
            params + [rank, max(1, int(rank * 0.90))],
        )
        wen = rows_from_sql(
            f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {base} "
            "AND rank>0 AND rank>=? AND rank<=? ORDER BY rank ASC LIMIT 50",
            params + [rank, int(rank * 1.3)],
        )
        bao = rows_from_sql(
            f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {base} "
            "AND rank>0 AND rank>? AND rank<=? ORDER BY rank ASC LIMIT 50",
            params + [int(rank * 1.3), int(rank * 1.6)],
        )

    if not (chong or wen or bao) and keyword and rank > 0:
        plain_params = [f"%{province}%"]
        plain_base = "province LIKE ? AND (score>0 OR rank>0)"
        if subject:
            plain_base += " AND (category=? OR category='' OR category IS NULL)"
            plain_params.append(subject)
        chong = rows_from_sql(
            f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {plain_base} "
            "AND rank>0 AND rank<? AND rank>=? ORDER BY rank ASC LIMIT 50",
            plain_params + [rank, max(1, int(rank * 0.90))],
        )
        wen = rows_from_sql(
            f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {plain_base} "
            "AND rank>0 AND rank>=? AND rank<=? ORDER BY rank ASC LIMIT 50",
            plain_params + [rank, int(rank * 1.3)],
        )
        bao = rows_from_sql(
            f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {plain_base} "
            "AND rank>0 AND rank>? AND rank<=? ORDER BY rank ASC LIMIT 50",
            plain_params + [int(rank * 1.3), int(rank * 1.6)],
        )

    only_major_group_result = False
    if major and province == "江苏":
        group_rows = []
        major_group_base = "province LIKE ? AND (score>0 OR rank>0)"
        major_group_params = [f"%{province}%"]
        if subject:
            major_group_base += " AND (category=? OR category='' OR category IS NULL)"
            major_group_params.append(subject)
        if rank > 0:
            group_rows = rows_from_sql(
                f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {major_group_base} "
                "AND rank>0 AND rank>=? AND rank<=? ORDER BY rank ASC LIMIT 120",
                major_group_params + [max(1, int(rank * 0.90)), int(rank * 1.6)],
            )
        elif score > 0:
            group_rows = rows_from_sql(
                f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {major_group_base} "
                "AND score>=? AND score<=? ORDER BY score ASC LIMIT 120",
                major_group_params + [score - 50, score + 25],
            )
        group_rows = [row for row in group_rows if is_major_group_text(row.get("major"))]
        if group_rows:
            only_major_group_result = True
            chong, wen, bao = [], [], []

    if not only_major_group_result and not (chong or wen or bao) and score > 0:
        chong = rows_from_sql(
            f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {base} "
            "AND score>? AND score<=? ORDER BY score DESC LIMIT 80",
            params + [score, score + 25],
        )
        wen = rows_from_sql(
            f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {base} "
            "AND score>=? AND score<=? ORDER BY score ASC LIMIT 50",
            params + [score - 25, score + 25],
        )
        bao = rows_from_sql(
            f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {base} "
            "AND score>=? AND score<? ORDER BY score ASC LIMIT 50",
            params + [score - 50, score - 25],
        )
        if not (chong or wen or bao):
            plain_base = "province LIKE ? AND (score>0 OR rank>0)"
            plain_params = [f"%{province}%"]
            if subject:
                plain_base += " AND (category=? OR category='' OR category IS NULL)"
                plain_params.append(subject)
            chong = rows_from_sql(
                f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {plain_base} "
                "AND score>? AND score<=? ORDER BY score DESC LIMIT 80",
                plain_params + [score, score + 25],
            )
            wen = rows_from_sql(
                f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {plain_base} "
                "AND score>=? AND score<=? ORDER BY score ASC LIMIT 50",
                plain_params + [score - 25, score + 25],
            )
            bao = rows_from_sql(
                f"SELECT school_name,major_name,score,rank,year FROM admission WHERE {plain_base} "
                "AND score>=? AND score<? ORDER BY score ASC LIMIT 50",
                plain_params + [score - 50, score - 25],
            )

    conn.close()
    return {"rank": rank, "score": score, "chong": chong, "wen": wen, "bao": bao, "source": "db"}


def baidu_search(query, limit=5):
    results = []
    sites = ["gaokao.chsi.com.cn", "eol.cn"]
    for province, domain in OFFICIAL_SITES.items():
        if province in query:
            sites.insert(0, domain)
            break

    for site in sites[:3]:
        if len(results) >= limit:
            break
        try:
            scoped_query = f"{query} site:{site}"
            url = "https://www.baidu.com/s?wd=" + urllib.parse.quote(scoped_query)
            request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(request, timeout=8) as response:
                html = response.read().decode("utf-8", errors="ignore")
            snippets = re.findall(r'class="c-abstract"[^>]*>(.*?)</span>', html)
            snippets += re.findall(r'<span class="content-right_[^"]*">(.*?)</span>', html)
            for snippet in snippets:
                clean = re.sub(r"<[^>]+>", "", snippet).strip()
                if len(clean) < 18:
                    continue
                results.append({"title": site, "summary": clean[:220], "source": "百度搜索"})
                if len(results) >= limit:
                    break
        except Exception:
            continue
    if not results:
        results.append({
            "title": "搜索结果较少",
            "summary": "这次联网补充没有拿到稳定结果，更建议填写 Tavily Key 获取更准的联网信息。",
            "source": "系统提示",
        })
    return results[:limit]


def tavily_search(query, api_key, limit=3):
    if not api_key:
        return []
    payload = json.dumps(
        {
            "api_key": api_key,
            "query": query,
            "search_depth": "basic",
            "max_results": limit,
            "include_answer": True,
        }
    ).encode("utf-8")
    request = urllib.request.Request(
        "https://api.tavily.com/search",
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=18) as response:
        data = json.loads(response.read().decode("utf-8"))
    findings = []
    answer = (data.get("answer") or "").strip()
    if answer:
        findings.append({"title": "Tavily总结", "summary": answer[:220], "source": "Tavily"})
    for item in data.get("results") or []:
        title = (item.get("title") or "联网结果").strip()
        summary = (item.get("content") or "").strip()
        if not summary:
            continue
        findings.append({"title": title[:60], "summary": summary[:220], "source": "Tavily"})
        if len(findings) >= limit + 1:
            break
    return findings


def build_search_queries(info, recommendations):
    queries = []
    province = info.get("province") or ""
    majors = info.get("majors") or []
    schools = info.get("schools") or []

    if province and majors:
        queries.append(f"{province} {majors[0]} 2025 录取位次")
        queries.append(f"{majors[0]} 专业 2026 就业前景")
    if schools:
        queries.append(f"{schools[0]} 招生专业 就业")
    for bucket in ["chong", "wen", "bao"]:
        for row in (recommendations.get(bucket) or [])[:2]:
            queries.append(f"{row['school']} 王牌专业 校园环境")
    return unique_list(queries)[:4]


def build_evidence(info, recommendations, web_findings):
    missing = determine_missing_fields("gaokao", {
        "province": info.get("province", ""),
        "rank": info.get("rank", 0),
        "score": info.get("score", 0),
        "subject": info.get("subject", ""),
        "majors": info.get("majors", []),
        "major_unknown": info.get("major_unknown", False),
        "accept_adjustment": info.get("accept_adjustment"),
        "region_pref": info.get("region_pref", []),
        "region_avoid": info.get("region_avoid", []),
        "career_goal": info.get("career_goal", ""),
        "budget": info.get("budget", ""),
    })

    summary = []
    total_db = sum(len(recommendations.get(bucket) or []) for bucket in ["chong", "wen", "bao"])
    if total_db:
        summary.append(f"本地数据库找到了 {total_db} 条相对匹配的学校和专业。")
    else:
        summary.append("本地数据库里暂时没有命中足够多的匹配结果。")
        if info.get("province") and info.get("majors"):
            summary.append(f"这次没有表格，是因为本地库里暂时没有找到“{info.get('province')} + {','.join(info.get('majors') or [])}”的可直接展示结果。")
    if info.get("province") == "江苏" and total_db:
        summary.append("江苏这批结果很多是专业组数据，表里看到的“不限、思想政治”等是选科要求，不等于已经细化到法学专业明细。")
    if info.get("province") == "江苏" and info.get("majors"):
        summary.append("如果你限定了法学这类具体专业，但本地库只有专业组数据，我会宁可少推，也不会把不确定能报的组硬塞给你。")
    if web_findings:
        summary.append(f"联网补充整理了 {len(web_findings)} 条辅助信息，主要用来看专业趋势和学校环境。")

    return {
        "db_recommendations": recommendations,
        "web_findings": web_findings,
        "missing_fields": missing,
        "evidence_summary": summary,
    }


def format_recommendations_for_model(recommendations):
    lines = []
    for bucket, label in [("chong", "冲"), ("wen", "稳"), ("bao", "保")]:
        rows = recommendations.get(bucket) or []
        if not rows:
            continue
        lines.append(f"{label}档：")
        for row in rows[:6]:
            lines.append(
                f"{row['school']}，{row['major']}，{row['year']}年，最低{row.get('score') or '?'}分，位次{row.get('rank') or '?'}。"
            )
    return "\n".join(lines)


def format_web_findings_for_model(web_findings):
    lines = []
    for item in web_findings[:6]:
        lines.append(f"{item['title']}：{item['summary']}")
    return "\n".join(lines)


def build_messages(mode, info, evidence, history, user_message):
    system_prompt = GAOKAO_SYSTEM_PROMPT if mode == "gaokao" else FUN_SYSTEM_PROMPT
    messages = [{"role": "system", "content": system_prompt}]

    profile_parts = []
    if mode == "gaokao":
        profile_parts.append(f"省份：{info.get('province') or '未提供'}")
        profile_parts.append(f"分数：{info.get('score') or '未提供'}")
        profile_parts.append(f"位次：{info.get('rank') or '未提供'}")
        profile_parts.append(f"选科：{info.get('subject') or '未提供'}")
        if info.get("majors"):
            profile_parts.append(f"意向专业：{','.join(info.get('majors') or [])}")
        elif info.get("major_unknown"):
            profile_parts.append("意向专业：暂时没想好，需要先给方向建议")
        else:
            profile_parts.append("意向专业：未提供")
        profile_parts.append(
            "是否接受调剂：" +
            ("接受" if info.get("accept_adjustment") is True else "不接受" if info.get("accept_adjustment") is False else "未说明")
        )
        if info.get("region_pref"):
            profile_parts.append(f"地域偏好：{','.join(info['region_pref'])}")
        if info.get("region_avoid"):
            profile_parts.append(f"地域避开：{','.join(info['region_avoid'])}")
        if info.get("career_goal"):
            profile_parts.append(f"发展倾向：{info['career_goal']}")
        if info.get("budget"):
            profile_parts.append(f"预算：{info['budget']}")
    else:
        for key, label in [("topic", "主题"), ("goal", "目标"), ("style", "风格"), ("guardrails", "禁区")]:
            value = info.get(key)
            if isinstance(value, list):
                value = ",".join(value)
            profile_parts.append(f"{label}：{value or '未说明'}")

    messages.append({"role": "system", "content": "用户当前已知信息：\n" + "\n".join(profile_parts)})

    if mode == "gaokao":
        messages.append({
            "role": "system",
            "content": "本地数据库整理：\n" + (format_recommendations_for_model(evidence["db_recommendations"]) or "暂时没有明确命中。"),
        })
    if evidence.get("web_findings"):
        messages.append({
            "role": "system",
            "content": "联网补充摘要：\n" + format_web_findings_for_model(evidence["web_findings"]),
        })

    for item in history[-10:]:
        role = item.get("role")
        content = (item.get("displayText") or item.get("content") or "").strip()
        if role not in {"user", "assistant"} or not content:
            continue
        messages.append({"role": role, "content": content})

    messages.append({"role": "user", "content": user_message})
    return messages


def stream_completion(api_url, api_key, model, messages, writer, temperature=0.6):
    payload = json.dumps(
        {
            "model": normalize_model_name(model),
            "messages": messages,
            "temperature": temperature,
            "stream": True,
            "thinking": {"type": "disabled"},
        }
    ).encode("utf-8")
    request = urllib.request.Request(
        api_url.rstrip("/") + "/v1/chat/completions",
        data=payload,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )

    with urllib.request.urlopen(request, timeout=180) as response:
        while True:
            raw_line = response.readline()
            if not raw_line:
                break
            line = raw_line.decode("utf-8", errors="ignore").strip()
            if not line or not line.startswith("data:"):
                continue
            data = line[5:].strip()
            if data == "[DONE]":
                break
            chunk = safe_json_loads(data, {})
            saw_finish = False
            for choice in chunk.get("choices") or []:
                delta = choice.get("delta") or {}
                content = delta.get("content") or ""
                if content:
                    writer({"type": "delta", "text": content})
                if choice.get("finish_reason"):
                    saw_finish = True
            if saw_finish:
                break


class Handler(BaseHTTPRequestHandler):
    def _send(self, data, code=200):
        self.send_response(code)
        self.send_header("Content-Type", "application/json;charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode("utf-8"))

    def _send_sse_headers(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream;charset=utf-8")
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.send_header("Connection", "close")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()

    def _read_json(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            length = 0
        raw = self.rfile.read(length) if length > 0 else b"{}"
        return safe_json_loads(raw.decode("utf-8"), {})

    def _sse_write(self, payload):
        try:
            self.wfile.write(f"data: {json.dumps(payload, ensure_ascii=False)}\n\n".encode("utf-8"))
            self.wfile.flush()
        except (BrokenPipeError, ConnectionResetError):
            raise

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.end_headers()

    def do_GET(self):
        if self.path == "/ping":
            return self._send({"ok": True, "db": HAS_DB})

        for image_name in ["img_suit.png", "img_scifi.png"]:
            if self.path == "/" + image_name:
                image_path = os.path.join(HERE, image_name)
                if os.path.exists(image_path):
                    self.send_response(200)
                    self.send_header("Content-Type", "image/png")
                    self.send_header("Cache-Control", "max-age=3600")
                    self.end_headers()
                    with open(image_path, "rb") as image_file:
                        self.wfile.write(image_file.read())
                    return

        self.send_response(200)
        self.send_header("Content-Type", "text/html;charset=utf-8")
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.end_headers()
        self.wfile.write(load_html_page().encode("utf-8"))

    def do_POST(self):
        if self.path != "/chat/stream":
            return self._send({"error": "not found"}, 404)

        self.close_connection = True
        body = self._read_json()
        mode = (body.get("mode") or "gaokao").strip()
        message = (body.get("message") or "").strip()
        history = body.get("history") or []
        config = body.get("config") or {}
        planner = body.get("planner") or {}

        if not message:
            return self._send({"error": "empty message"}, 400)

        self._send_sse_headers()

        def emit(payload):
            self._sse_write(payload)

        try:
            emit({"type": "phase", "text": "正在整理你刚才提供的信息..."})

            plan_state = ensure_plan_state(mode, planner)
            try:
                extracted = ai_extract_fields(
                    mode,
                    message,
                    history,
                    plan_state.get("current_questions") or [],
                    plan_state.get("collected") or {},
                    config,
                )
            except Exception:
                extracted = {}
            if not extracted:
                extracted = fallback_extract_info(mode, message, plan_state.get("current_questions") or [])
            plan_state["collected"] = merge_collected_for_mode(mode, plan_state["collected"], extracted)
            plan_state["last_user_message"] = message
            plan_state = prune_answered_questions(mode, plan_state)

            plan_enabled = bool(planner.get("enabled"))
            if planner_should_ask(mode, plan_enabled, plan_state, message):
                question_payload = build_question_response(mode, plan_state)
                plan_state = question_payload["planState"]
                emit(question_payload)
                emit({"type": "done", "kind": "question", "planState": plan_state})
                return

            plan_state["current_questions"] = []
            plan_state["ready"] = True
            plan_state = sanitize_plan_state(mode, plan_state)

            collected = plan_state["collected"]
            recommendations = {"rank": 0, "score": 0, "chong": [], "wen": [], "bao": [], "source": "empty"}
            web_findings = []

            if mode == "gaokao":
                emit({"type": "phase", "text": "正在查本地录取数据库..."})
                keyword = ",".join(collected.get("majors") or [])
                recommendations = recommend_data(
                    province=collected.get("province", ""),
                    major=(collected.get("majors") or [""])[0] if collected.get("majors") else "",
                    keyword=keyword,
                    rank=collected.get("rank", 0),
                    score=collected.get("score", 0),
                    subject=collected.get("subject", ""),
                )

                emit({"type": "phase", "text": "正在补充联网参考信息..."})
                queries = build_search_queries(collected, recommendations)
                tavily_key = (config.get("tavily") or "").strip()
                for query in queries:
                    findings = []
                    if tavily_key:
                        try:
                            findings = tavily_search(query, tavily_key, limit=2)
                        except Exception:
                            findings = []
                    if not findings:
                        findings = baidu_search(query, limit=2)
                    web_findings.extend(findings)
                web_findings = unique_list([json.dumps(item, ensure_ascii=False) for item in web_findings])
                web_findings = [json.loads(item) for item in web_findings[:8]]

            evidence = build_evidence(collected, recommendations, web_findings)
            emit({"type": "evidence", "evidence": evidence})

            api_url = (config.get("url") or "https://api.deepseek.com").strip()
            api_key = (config.get("key") or "").strip()
            model = normalize_model_name(config.get("model"))
            if not api_key:
                emit({
                    "type": "error",
                    "text": "还没填写 API Key。点右上角 API 设置，填好以后我就能继续回答。",
                    "planState": sanitize_plan_state(mode, plan_state),
                })
                emit({"type": "done", "kind": "error", "planState": sanitize_plan_state(mode, plan_state)})
                return

            emit({"type": "phase", "text": "正在生成最终建议..."})
            messages = build_messages(mode, collected, evidence, history, message)
            stream_completion(api_url, api_key, model, messages, emit, temperature=0.65 if mode == "gaokao" else 0.8)
            emit({"type": "done", "kind": "answer", "planState": sanitize_plan_state(mode, plan_state)})
        except urllib.error.HTTPError as exc:
            raw = exc.read().decode("utf-8", errors="ignore")
            data = safe_json_loads(raw, {})
            message_text = data.get("error", {}).get("message") if isinstance(data, dict) else ""
            message_text = message_text or f"请求模型失败，HTTP {exc.code}。"
            emit({"type": "error", "text": f"出错了：{message_text}", "planState": sanitize_plan_state(mode, plan_state)})
            emit({"type": "done", "kind": "error", "planState": sanitize_plan_state(mode, plan_state)})
        except Exception as exc:
            emit({"type": "error", "text": f"出错了：{exc}", "planState": sanitize_plan_state(mode, plan_state)})
            emit({"type": "done", "kind": "error", "planState": sanitize_plan_state(mode, plan_state)})

    def log_message(self, format_text, *args):
        message = format_text % args if args else format_text
        if any(path in message for path in ["/ping", "/chat/stream"]):
            print(f"[REQ] {message}")


def load_html_page():
    with open(os.path.join(HERE, "index.html"), "r", encoding="utf-8") as html_file:
        return html_file.read()


def main():
    port = 8765
    server = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print(f"雪峰Agent: http://127.0.0.1:{port}/")
    print(f"数据库: {'已加载' if HAS_DB else '未找到'}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.shutdown()
        print("\n已停止")


if __name__ == "__main__":
    main()
