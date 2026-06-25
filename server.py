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
import zipfile
import xml.etree.ElementTree as ET
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
REQUIREMENT_HINTS = ["选考", "选科", "首选", "再选", "科目要求", "专业要求", "限选", "须选", "要求"]
REPLACE_LIST_FIELDS = {"subject_detail", "majors", "schools", "region_pref", "region_avoid", "guardrails"}
MAJOR_SEARCH_ALIASES = {
    "计算机": ["计算机", "计算机类", "计算机科学与技术", "软件", "软件工程", "人工智能", "大数据", "数据科学", "信息安全", "网络工程", "物联网"],
    "软件": ["软件", "软件工程", "计算机", "程序设计"],
    "人工智能": ["人工智能", "智能科学", "数据科学", "计算机", "软件"],
    "法学": ["法学", "法律", "法治", "涉外法治", "司法", "知识产权"],
    "医学": ["医学", "临床", "口腔", "预防医学", "基础医学"],
    "金融": ["金融", "经济", "经济学", "财政", "会计", "财务"],
}
SUBJECT_DETAIL_ABBR_MAP = {
    "物化地": ["物理", "化学", "地理"],
    "物化政": ["物理", "化学", "政治"],
    "物化生": ["物理", "化学", "生物"],
    "物生地": ["物理", "生物", "地理"],
    "物生政": ["物理", "生物", "政治"],
    "物政地": ["物理", "政治", "地理"],
    "史政地": ["历史", "政治", "地理"],
    "史地政": ["历史", "地理", "政治"],
    "历政地": ["历史", "政治", "地理"],
    "历地政": ["历史", "地理", "政治"],
    "政史地": ["政治", "历史", "地理"],
    "政地史": ["政治", "地理", "历史"],
    "地政史": ["地理", "政治", "历史"],
}

# 当 DB 的 category 为空/综合时，用专业名关键词判断文理倾向
# 历史类用户不应被推荐这些明显仅限物理类的专业
STEM_MAJOR_KEYWORDS = [
    "工程", "电气", "机械", "材料", "化工", "土木", "计算机科学", "软件",
    "人工智能", "大数据", "数据科学", "网络工程", "信息安全", "物联网",
    "自动化", "电子", "通信", "智能制造", "机器人", "微电子", "集成电路",
    "冶金", "采矿", "测绘", "水利", "能源", "地质", "石油", "船舶",
    "车辆", "航空", "航天", "力学", "建筑", "城乡规划", "风景园林",
    "生物科学", "生物技术", "生物工程", "生物医学", "生物制药",
    "物理学", "应用物理", "化学", "应用化学", "数学与应用数学",
    "信息与计算", "信息管理", "统计学", "应用统计",
    "光电", "测控", "给排水", "包装", "印刷", "纺织", "轻化",
    "核工程", "兵器", "弹药", "探测制导",
]
# 如果专业名包含以下关键词，大概率是文科/文理兼收，历史类用户放行
LIBERAL_ARTS_MAJOR_KEYWORDS = [
    "法学", "法律", "知识产权", "社会工作", "社会学",
    "新闻", "传播", "广告", "编辑", "出版", "网络与新媒体",
    "汉语言", "汉语", "英语", "日语", "俄语", "德语", "法语", "韩语",
    "翻译", "商务英语", "外国语言", "中国语言", "古典文献",
    "历史", "考古", "文物", "哲学", "宗教", "逻辑",
    "政治", "行政", "公共管理", "马克思主义", "思想政治教育", "国际政治",
    "教育", "师范", "学前", "小学教育", "特殊教育", "心理学",
    "会计", "财务", "审计", "金融", "经济", "财政", "税务",
    "国际贸易", "工商管理", "市场营销", "人力资源", "物流管理",
    "旅游管理", "酒店管理", "会展", "文化产业管理", "公共事业管理",
    "法学", "知识产权", "监狱学",
    "设计", "美术", "音乐", "舞蹈", "戏剧", "影视", "播音",
    "编导", "表演", "书法", "摄影", "动画", "数字媒体艺术",
    "护理", "助产", "康复治疗", "中药", "针灸", "推拿", "中医",
    "预防医学", "卫生检验", "医学检验", "医学实验",
    "体育", "运动", "武术",
]

def _category_unknown_major_suitable(subject, major_name):
    """当 DB 的 category 为空/综合时，用专业名关键词粗略判断文理兼容性。"""
    if subject not in ("物理类", "历史类"):
        return True
    major_name = major_name or ""
    if subject == "历史类":
        # 师范类标记：即使含 STEM 关键词，加了师范也可能是文理兼收
        has_teacher_tag = "师范" in major_name
        # 先查理工关键词，命中就过滤（除了带师范标记的）
        for kw in STEM_MAJOR_KEYWORDS:
            if kw in major_name:
                if has_teacher_tag:
                    return True  # 数学师范、物理师范等文理兼收
                return False
        # 再看文科关键词，命中就放行
        for kw in LIBERAL_ARTS_MAJOR_KEYWORDS:
            if kw in major_name:
                return True
        # 都没命中 → 不确定，放行（宁可多推不少推）
        return True
    # 物理类：大部分专业都能报，不过滤
    return True

GAOKAO_PRIORITY = [
    "province",
    "score_or_rank",
    "subject",
    "subject_detail",
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
    "province": 3,
    "score_or_rank": 3,
    "subject": 3,
    "subject_detail": 2,
    "majors": 3,
    "accept_adjustment": 3,
    "region_pref": 2,
    "career_goal": 2,
    "budget": 2,
    "topic": 2,
    "goal": 2,
    "style": 1,
    "guardrails": 1,
}

GAOKAO_QUESTION_TEXT = {
    "province": "你是哪个省的考生？",
    "score_or_rank": "告诉我你的分数，或者你的位次，给一个就行。",
    "subject": "你是物理类、历史类，还是像浙江这种综合改革省份？",
    "subject_detail": "如果你是新高考选科，麻烦再告诉我你的具体组合，比如物化地、物化生、史政地这种。",
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
9. subject_detail 必须是数组，优先返回更细的选科，如物理、化学、生物、政治、地理、历史。
10. career_goal 尽量归一为：考研、直接就业、考公、考编；无法判断就留空。
11. budget 保留成适合展示的简短中文，比如“正常预算”“预算偏谨慎”“不太高就行”。
12. majors、schools、region_pref、region_avoid、guardrails 必须是数组。
13. thinking_mode 固定返回 false。

JSON 模板：
{"province":"","score":0,"rank":0,"subject":"","subject_detail":[],"majors":[],"major_unknown":false,"schools":[],"accept_adjustment":null,"region_pref":[],"region_avoid":[],"career_goal":"","budget":"","topic":"","goal":"","style":"","guardrails":[],"thinking_mode":false}
"""

SEMANTIC_RETRIEVAL_PROMPT = """你是一个志愿检索词补全助手。
你的任务不是回答用户问题，也不是推荐学校，而是在本地数据库没有直接命中时，帮系统补一些更合适的中文检索词。

输出规则：
1. 只能输出一个 JSON 对象，不能输出解释，不能输出 Markdown。
2. terms 必须是数组，返回 0 到 8 个中文短词。
3. 这些词只能是专业名称、专业大类、专业组方向、学科方向或常见近义叫法。
4. 不要输出学校名、分数、位次、城市、评价词。
5. 如果用户表达已经很明确，也可以补充更常见的同类叫法。
6. 如果实在无法判断，就返回空数组。

JSON 模板：
{"terms":[],"reason":""}
"""

EXTRACT_VERIFY_PROMPT = """你是一个信息提取验证器。你的任务不是提取新信息，而是检查已有的提取结果是否正确。

你会收到：
- 用户的原始消息
- 系统用正则/字典提取出的字段
- 已有的上下文信息

你需要逐项检查提取结果有没有这些错误：
1. 张冠李戴：把选科当成专业（比如用户说"选了物化生"，不应提取"生物"为专业）、把日常用词当成字段（比如"老师说"不应提取"师范"）
2. 遗漏：用户明确说了的信息没提取到
3. 矛盾：提取结果和已有上下文冲突
4. 数字错误：分数/位次解析错了

输出规则：
1. 只能输出一个 JSON 对象。
2. correct 是 bool：true 表示提取结果没问题可以直接用，false 表示有问题需要修正。
3. issues 是字符串数组，列出每个具体问题。没有问题时为空数组。
4. 不要输出解释或 Markdown。

JSON 模板：
{"correct":true,"issues":[]}
"""

EXTRACT_FIX_PROMPT = """你是一个信息提取修正器。原始提取有错误，你需要根据用户消息和上下文重新提取正确的字段。

输出规则：
1. 只能输出一个 JSON 对象，不能输出解释，不能输出 Markdown。
2. 结合 mode、current_questions、collected、history 来理解用户这一轮输入。
3. 如果用户没有否定旧信息，就不要清空旧信息。
4. 如果用户明确表示"不知道想学什么""专业没想好""我也不知道学什么"，请返回 major_unknown=true，并把 majors 设为空数组。
5. 如果用户只是说"我也不知道"，要结合 current_questions 判断是在回答哪个问题。
6. accept_adjustment 只能是 true、false 或 null。
7. score 和 rank 只填数字，没有就填 0。
8. subject 只允许：物理类、历史类、综合。
9. subject_detail 必须是数组，优先返回更细的选科。
10. career_goal 尽量归一为：考研、直接就业、考公、考编；无法判断就留空。
11. budget 保留成适合展示的简短中文。
12. majors、schools、region_pref、region_avoid、guardrails 必须是数组。
13. thinking_mode 固定返回 false。
14. 修正时特别注意：不要把选科描述当成专业（"选了物化生"中的化学/生物不是专业），不要把日常对话中的职业称呼当成专业意向（"老师说"不是想当老师，"医生说"不是想当医生，除非明确说"想当医生""想做老师"）。

JSON 模板：
{"province":"","score":0,"rank":0,"subject":"","subject_detail":[],"majors":[],"major_unknown":false,"schools":[],"accept_adjustment":null,"region_pref":[],"region_avoid":[],"career_goal":"","budget":"","topic":"","goal":"","style":"","guardrails":[],"thinking_mode":false}
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


def default_subject_detail(subject):
    if subject == "物理类":
        return ["物理"]
    if subject == "历史类":
        return ["历史"]
    return []


def normalize_category_text(category_raw):
    category_raw = (category_raw or "").strip()
    if "物理" in category_raw:
        return "物理类"
    if "历史" in category_raw:
        return "历史类"
    if "综合" in category_raw:
        return "综合"
    return category_raw


def build_major_search_terms(major="", keyword=""):
    raw_terms = []
    for item in [major] + [part.strip() for part in str(keyword or "").split(",")]:
        if item:
            raw_terms.append(str(item).strip())
    raw_terms = unique_list(raw_terms)
    expanded = list(raw_terms)
    for term in raw_terms:
        compact = term.replace("专业", "").replace("方向", "").replace("类", "").strip()
        if compact:
            expanded.append(compact)
        for key, aliases in MAJOR_SEARCH_ALIASES.items():
            if key in term or term in aliases:
                expanded.extend(aliases)
    return unique_list(expanded)


def text_matches_terms(text, terms):
    text = text or ""
    return any(term and term in text for term in terms or [])


def _col_letter_to_index(col_letters):
    index = 0
    for ch in col_letters.upper():
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index - 1


def extract_xlsx_rows(path):
    ns_main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    ns_rel = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheets = workbook.find(f"{ns_main}sheets")
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root:
                texts = []
                for node in item.iter(f"{ns_main}t"):
                    texts.append(node.text or "")
                shared.append("".join(texts))

        first_sheet = None
        for sheet in sheets or []:
            name = sheet.attrib.get("name", "")
            if "说明" in name:
                continue
            first_sheet = sheet
            break
        if first_sheet is None and sheets is not None and len(sheets):
            first_sheet = sheets[0]
        if first_sheet is None:
            return []

        target = rel_map[first_sheet.attrib[ns_rel + "id"]].lstrip("/")
        worksheet = ET.fromstring(archive.read(target))
        rows = []
        for row in worksheet.iter(f"{ns_main}row"):
            row_values = []
            for cell in row.iter(f"{ns_main}c"):
                cell_ref = cell.attrib.get("r", "")
                col_match = re.match(r"([A-Z]+)\d+", cell_ref)
                col_idx = _col_letter_to_index(col_match.group(1)) if col_match else len(row_values)
                while len(row_values) < col_idx:
                    row_values.append("")
                cell_type = cell.attrib.get("t")
                if cell_type == "inlineStr":
                    inline = cell.find(f"{ns_main}is")
                    value = "".join(node.text or "" for node in inline.iter(f"{ns_main}t")) if inline is not None else ""
                else:
                    value_node = cell.find(f"{ns_main}v")
                    value = value_node.text if value_node is not None else ""
                    if cell_type == "s" and value:
                        try:
                            value = shared[int(value)]
                        except (IndexError, ValueError):
                            value = ""
                row_values.append(value)
            if any(v for v in row_values):
                rows.append(row_values)
        return rows


def append_user_data_record(school, major, category, province, score, rank, year):
    if not school or score is None or rank is None:
        return
    USER_DATA.append({
        "school": school,
        "major": major,
        "year": year,
        "category": category,
        "score": score,
        "rank": rank,
        "province": province,
    })


def import_user_rows(rows):
    for row in rows:
        if not row or not row[0]:
            continue
        school = str(row[0]).strip()
        if len(school) < 2 or school in ["学校名称", "院校名称"]:
            continue
        note = str(row[8]).strip() if len(row) > 8 and row[8] else ""
        if "示例" in note or "不参与排序" in note:
            continue
        major = str(row[1]).strip() if len(row) > 1 and row[1] else ""
        category = normalize_category_text(str(row[2]).strip() if len(row) > 2 and row[2] else "")
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
        append_user_data_record(school, major, category, province, score_24, rank_24, 2024)
        append_user_data_record(school, major, category, province, score_25, rank_25, 2025)


def pick_primary_sheet(workbook):
    for sheet in workbook.worksheets:
        if "说明" not in sheet.title:
            return sheet
    return workbook.active


def load_user_data():
    global USER_DATA
    USER_DATA = []
    if not os.path.exists(USER_XLSX):
        print(f"[user data] {USER_XLSX} 不存在，跳过自定义数据加载")
        return
    raw_rows = []
    loader = "none"
    try:
        import openpyxl
        wb = openpyxl.load_workbook(USER_XLSX, data_only=True)
        sheet = pick_primary_sheet(wb)
        raw_rows = list(sheet.iter_rows(min_row=2, values_only=True))
        wb.close()
        loader = "openpyxl"
    except Exception as exc:
        try:
            raw_rows = extract_xlsx_rows(USER_XLSX)[1:]
            loader = "xml_fallback"
        except Exception as fallback_exc:
            print(f"[user data] openpyxl={exc}; xml_fallback={fallback_exc}")
    if loader != "none":
        import_user_rows(raw_rows)
        print(f"[user data] 已从 {USER_XLSX} 加载 {len(USER_DATA)} 条自定义数据 (loader={loader}, raw_rows={len(raw_rows)})")


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
    detail = extract_subject_detail(text)
    if "物理" in detail and "历史" not in detail:
        return "物理类"
    if "历史" in detail and "物理" not in detail:
        return "历史类"
    return ""


def extract_subject_detail(text):
    text = re.sub(r"(男生|女生)", "", text or "")
    found = []
    for keyword, subjects in SUBJECT_DETAIL_ABBR_MAP.items():
        if keyword in text:
            found.extend(subjects)
    for subject in SUBJECT_DETAIL_KEYWORDS:
        if subject in text:
            found.append(subject)
    return unique_list(found)


def normalize_category_subject(category):
    text = (category or "").strip()
    if not text or text == "综合":
        return ""
    if "文科" in text:
        return "历史类"
    if "理科" in text:
        return "物理类"
    if "首选物理" in text or text == "物理类":
        return "物理类"
    if "首选历史" in text or text == "历史类":
        return "历史类"
    if "物理" in text and "历史" not in text:
        return "物理类"
    if "历史" in text and "物理" not in text:
        return "历史类"
    return ""


def extract_requirement_subjects(major_name):
    major_name = major_name or ""
    if not major_name:
        return []
    segments = []
    for segment in re.findall(r"[（(]([^()（）]{1,40})[)）]", major_name):
        if any(hint in segment for hint in REQUIREMENT_HINTS):
            segments.append(segment)
    if any(hint in major_name for hint in REQUIREMENT_HINTS):
        segments.append(major_name)
    if not segments:
        return []
    found = []
    for segment in segments:
        for keyword, subjects in SUBJECT_DETAIL_ABBR_MAP.items():
            if keyword in segment:
                found.extend(subjects)
        for subject in SUBJECT_DETAIL_KEYWORDS:
            if subject in segment:
                found.append(subject)
    return unique_list(found)


def requirement_is_unlimited(major_name):
    major_name = major_name or ""
    if "不限" not in major_name:
        return False
    bracket_text = "".join(re.findall(r"[（(]([^()（）]{1,40})[)）]", major_name))
    return "不限" in bracket_text or any(hint in major_name for hint in REQUIREMENT_HINTS)


def is_major_group_text(major_name):
    text = (major_name or "").strip()
    return bool(text) and not any(major in text for major in MAJORS)


def subject_requirement_compatible(subject, subject_detail, category, major_name):
    category_subject = normalize_category_subject(category)
    if subject in {"历史类", "物理类"} and category_subject and category_subject != subject:
        return False

    major_name = major_name or ""
    if not major_name:
        return True
    if requirement_is_unlimited(major_name):
        return True

    required = extract_requirement_subjects(major_name)
    if not required:
        # category 为空/综合时，用专业名关键词做文理判断
        if subject in ("物理类", "历史类") and category_subject in ("", "综合"):
            return _category_unknown_major_suitable(subject, major_name)
        return True

    detail_set = set(subject_detail or [])
    if detail_set:
        return all(item in detail_set for item in required)

    first_subject = "物理" if subject == "物理类" else "历史" if subject == "历史类" else ""
    if first_subject and all(item == first_subject for item in required):
        return True
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
    text = re.sub(r"(男生|女生)", "", text or "")
    negative_blocks = re.findall(r"(?:不学|不考虑|拒绝|排斥|不接受|别推荐|不想学).*?(?:[。，,；;\n]|$)", text)
    negative_text = "".join(negative_blocks)
    ambiguous_patterns = {
        "化学": r"(?:(?:想|要|准备|考虑|打算)?学|报|读|专业|方向).{0,4}化学|化学(?:专业|类|科学|工程|工艺|师范|教育)",
        "生物": r"(?:(?:想|要|准备|考虑|打算)?学|报|读|专业|方向).{0,4}生物|生物(?:专业|类|科学|工程|技术|医学)",
        "地理": r"(?:(?:想|要|准备|考虑|打算)?学|报|读|专业|方向).{0,4}地理|地理(?:专业|类|科学|信息|师范)",
        "数学": r"(?:(?:想|要|准备|考虑|打算)?学|报|读|专业|方向).{0,4}数学|数学(?:专业|类|科学|师范)",
        "英语": r"(?:(?:想|要|准备|考虑|打算)?学|报|读|专业|方向).{0,4}英语|英语(?:专业|类|师范)",
        "日语": r"(?:(?:想|要|准备|考虑|打算)?学|报|读|专业|方向).{0,4}日语|日语(?:专业|类)",
    }
    found = []
    for major in MAJORS:
        if major in negative_text:
            continue
        if major in ambiguous_patterns:
            if re.search(ambiguous_patterns[major], text):
                found.append(major)
            continue
        if major in text:
            found.append(major)
    for keyword, major in [("律师", "法学"), ("司法", "法学"), ("医生", "医学"), ("医师", "医学"), ("教师", "师范")]:
        if keyword in text and major not in negative_text:
            found.append(major)
    # "老师" 太容易误匹配（"老师说""听老师说"只是指人不是说想当老师）
    # 只在明确的职业意图语境下才映射到师范
    teacher_intent = re.search(r"(?:想当|想做|当个?|成为|报考?).{0,4}老师|老师.{0,4}(?:专业|方向|编)", text)
    if teacher_intent and "师范" not in negative_text:
        found.append("师范")
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
            elif field in {"topic", "goal", "style"} and compact:
                if compact in {"不知道", "我也不知道", "没想好", "都行", "随便"}:
                    result[field] = "未指定"
                else:
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
                details = extract_subject_detail(text)
                if details:
                    result["subject_detail"] = details
                else:
                    result["subject_detail"] = default_subject_detail(subject)
        elif field == "subject_detail":
            details = extract_subject_detail(text)
            if details:
                result["subject_detail"] = details
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
            elif compact in {"不知道", "我也不知道", "没想好", "都行", "随便", "无偏好"}:
                result["region_pref"] = ["无偏好"]
            if region_avoid:
                result["region_avoid"] = region_avoid
        elif field == "career_goal":
            career_goal = extract_career_goal(text)
            if career_goal:
                result["career_goal"] = career_goal
            elif compact in {"不知道", "我也不知道", "没想好", "还没想好", "都行", "随便", "再说"}:
                result["career_goal"] = "未定"
        elif field == "budget":
            budget = extract_budget_text(text, contextual=True)
            if budget:
                result["budget"] = budget
            elif compact in {"不知道", "我也不知道", "没想好", "都行", "随便", "正常", "普通"}:
                result["budget"] = "正常预算"
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
        "subject_detail": extract_subject_detail(text),
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
        "subject_detail": base.get("subject_detail", []),
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
            if field == "subject_detail":
                if result.get("subject_detail"):
                    focused["subject_detail"] = result["subject_detail"]
                continue
            if field == "subject":
                if result.get("subject"):
                    focused["subject"] = result["subject"]
                if result.get("subject_detail"):
                    focused["subject_detail"] = result["subject_detail"]
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
            return merge_collected(result, focused)
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
    if "subject_detail" in parsed and isinstance(parsed.get("subject_detail"), list):
        normalized["subject_detail"] = unique_list([
            str(item).strip()
            for item in parsed.get("subject_detail") or []
            if str(item).strip() in SUBJECT_DETAIL_KEYWORDS
        ])
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


def _ai_call_json(prompt, user_content, config, max_tokens=320, temperature=0):
    """通用 AI JSON 调用：发送 prompt + user_content，返回解析后的 dict。失败返回 None。"""
    api_key = (config.get("key") or "").strip()
    api_url = (config.get("url") or "https://api.deepseek.com").strip().rstrip("/")
    if not api_key:
        return None
    model = normalize_model_name(config.get("model"))
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": prompt},
            {"role": "user", "content": json.dumps(user_content, ensure_ascii=False)},
        ],
        "temperature": temperature,
        "max_tokens": max_tokens,
        "thinking": {"type": "disabled"},
    }
    request = urllib.request.Request(
        api_url + "/v1/chat/completions",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=35) as response:
        data = json.loads(response.read().decode("utf-8"))
    content = ((data.get("choices") or [{}])[0].get("message") or {}).get("content", "")
    content = content.replace("```json", "").replace("```", "").strip()
    return safe_json_loads(content)


def ai_verify_extraction(mode, message, regex_extracted, collected, history, config):
    """AI 验证正则提取结果是否正确。返回 {"correct": bool, "issues": [...]}。失败返回 None 视为通过。"""
    if not (config.get("key") or "").strip():
        return None
    try:
        result = _ai_call_json(
            EXTRACT_VERIFY_PROMPT,
            {
                "mode": mode,
                "message": message,
                "regex_extracted": regex_extracted,
                "collected": collected or {},
                "history": (history or [])[-6:],
            },
            config,
            max_tokens=200,
            temperature=0,
        )
        if isinstance(result, dict) and "correct" in result:
            return {"correct": bool(result.get("correct")), "issues": result.get("issues") or []}
    except Exception:
        pass
    return None


def ai_fix_extraction(mode, message, regex_extracted, issues, collected, history, config):
    """AI 修正提取结果。返回修正后的字段 dict，失败返回 None。"""
    if not (config.get("key") or "").strip():
        return None
    try:
        result = _ai_call_json(
            EXTRACT_FIX_PROMPT,
            {
                "mode": mode,
                "message": message,
                "regex_extracted": regex_extracted,
                "issues": issues or [],
                "current_questions": [],
                "collected": collected or {},
                "history": (history or [])[-8:],
            },
            config,
            max_tokens=400,
            temperature=0,
        )
        if isinstance(result, dict):
            return normalize_extracted_fields(mode, result)
    except Exception:
        pass
    return None


def ai_expand_retrieval_terms(collected, history, config):
    api_key = (config.get("key") or "").strip()
    if not api_key:
        return []

    majors = collected.get("majors") or []
    if not majors and not collected.get("career_goal"):
        return []

    api_url = (config.get("url") or "https://api.deepseek.com").strip().rstrip("/")
    model = normalize_model_name(config.get("model"))
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SEMANTIC_RETRIEVAL_PROMPT},
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "province": collected.get("province", ""),
                        "subject": collected.get("subject", ""),
                        "subject_detail": collected.get("subject_detail", []),
                        "majors": majors,
                        "major_unknown": collected.get("major_unknown", False),
                        "career_goal": collected.get("career_goal", ""),
                        "latest_history": (history or [])[-6:],
                    },
                    ensure_ascii=False,
                ),
            },
        ],
        "temperature": 0,
        "max_tokens": 220,
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
    if not isinstance(parsed, dict) or not isinstance(parsed.get("terms"), list):
        return []
    terms = unique_list([str(item).strip() for item in parsed.get("terms") or [] if str(item).strip()])
    banned = {"学校", "大学", "学院", "冲", "稳", "保", "推荐", "志愿", "就业", "考研"}
    return [term for term in terms if term not in banned][:8]


def get_default_collected(mode):
    if mode == "gaokao":
        return {
            "province": "",
            "rank": 0,
            "score": 0,
            "subject": "",
            "subject_detail": [],
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
    incoming = incoming or {}
    for key, value in incoming.items():
        if value is None:
            continue
        if isinstance(value, list):
            if value:
                if key in REPLACE_LIST_FIELDS:
                    merged[key] = unique_list(value)
                else:
                    merged[key] = unique_list((merged.get(key) or []) + value)
            elif key not in merged:
                merged[key] = []
        elif isinstance(value, bool):
            if key == "major_unknown":
                continue  # 由下方专门的 major/major_unknown 分支处理，避免 False 覆盖已有的 True
            merged[key] = value
        elif isinstance(value, int):
            if value > 0:
                merged[key] = value
        elif isinstance(value, str):
            if value.strip():
                merged[key] = value.strip()
        else:
            merged[key] = value
    if incoming.get("majors") and not incoming.get("major_unknown"):
        # 有专业且用户没有明确说"不知道"→ 采纳
        merged["majors"] = unique_list(incoming.get("majors") or [])
        merged["major_unknown"] = False
    elif incoming.get("major_unknown") is True:
        # 用户明确说不知道学什么 → 清空正则误提的专业（如选了物化生→"生物"）
        merged["major_unknown"] = True
        merged["majors"] = []
    incoming_subject = (incoming.get("subject") or "").strip() if isinstance(incoming.get("subject"), str) else ""
    if incoming.get("subject_detail"):
        merged["subject_detail"] = unique_list(incoming.get("subject_detail") or [])
    elif incoming_subject:
        merged["subject_detail"] = default_subject_detail(incoming_subject)
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
    # 必填字段多给一次机会追问，之后才跳过（避免无限循环）
    if field in required_fields_for_mode(mode) and attempts == limit:
        return False
    return True


def planner_should_ask(mode, plan_enabled, state, message):
    missing = determine_missing_fields(mode, state["collected"])
    effective_missing = [field for field in missing if not should_skip_field(mode, state, field)]
    if mode == "gaokao":
        # 必填字段不全：无视 plan 开关，必须追问
        core_missing = [field for field in required_fields_for_mode(mode) if field in effective_missing]
        if core_missing:
            return True
        # 必填已全，可选字段：看 plan 开关
        if plan_enabled and effective_missing:
            return True
        return False
    if not plan_enabled:
        return False
    stripped = message.strip()
    if len(stripped) < 8:
        return bool(effective_missing)
    return bool(effective_missing and not state["collected"].get("goal"))


def select_questions(mode, state):
    priority = GAOKAO_PRIORITY if mode == "gaokao" else FUN_PRIORITY
    missing = determine_missing_fields(mode, state["collected"])
    available = [field for field in missing if not should_skip_field(mode, state, field)]
    if mode == "gaokao":
        core_available = [field for field in required_fields_for_mode(mode) if field in available]
        if core_available:
            available = core_available
    selected = []
    for field in priority:
        if field not in available:
            continue
        if field in state["asked_fields"]:
            continue
        selected.append(field)
        if len(selected) >= 2:
            break
    # 所有可选字段都问过了 → 允许从 still-missing 的必填字段中重问
    if not selected:
        for field in priority:
            if field not in available:
                continue
            selected.append(field)
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


def ai_followup_question(mode, state, message, history, config):
    api_key = (config.get("key") or "").strip()
    if not api_key:
        return None
    api_url = (config.get("url") or "https://api.deepseek.com").strip().rstrip("/")
    model = normalize_model_name(config.get("model"))
    allowed_fields = (
        "province、score_or_rank、subject、subject_detail、majors、accept_adjustment、region_pref、career_goal、budget"
        if mode == "gaokao"
        else "topic、goal、style、guardrails"
    )
    prompt = f"""你是一个追问助手。你的任务是基于当前已知信息，生成1到2个最关键的中文追问。
输出规则：
1. 只能输出JSON对象。
2. fields 必须是数组，最多2个，只能从 {allowed_fields} 中选择。
3. question 必须是给用户看的中文追问，尽量短，一次只问1到2个问题。
4. 不要重复 collected 里已经明确有值的信息。
5. 如果当前信息已经足够，不要乱问，fields 设为空数组，question 设为空字符串。
JSON模板：{{"fields":[],"question":""}}"""
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": prompt},
            {
                "role": "user",
                "content": json.dumps(
                    {
                        "mode": mode,
                        "message": message,
                        "collected": state.get("collected") or {},
                        "asked_fields": state.get("asked_fields") or [],
                        "current_questions": state.get("current_questions") or [],
                        "missing_fields": determine_missing_fields(mode, state.get("collected") or {}),
                        "history": history[-8:] if isinstance(history, list) else [],
                    },
                    ensure_ascii=False,
                ),
            },
        ],
        "temperature": 0,
        "max_tokens": 220,
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
    fields = []
    if isinstance(parsed, dict) and isinstance(parsed.get("fields"), list):
        fields = [str(item).strip() for item in parsed.get("fields") or [] if str(item).strip()]
    fields = [field for field in fields if field in GAOKAO_PRIORITY or field in FUN_PRIORITY]
    question = ""
    if isinstance(parsed, dict) and isinstance(parsed.get("question"), str):
        question = parsed.get("question", "").strip()
    if not question:
        question = build_question_text(mode, fields[:2] or select_questions(mode, state)[:1])
    return {
        "fields": fields[:2],
        "question": question,
    }


def build_question_response(mode, state, config=None, history=None, message=""):
    state = sanitize_plan_state(mode, state)
    fields = select_questions(mode, state)
    if not fields:
        ai_question = ai_followup_question(mode, state, message, history or [], config or {})
        if ai_question and ai_question.get("question"):
            fields = ai_question.get("fields") or []
            state["asked_fields"] = unique_list(state["asked_fields"] + fields)
            asked_count = dict(state.get("asked_count") or {})
            for field in fields:
                asked_count[field] = int(asked_count.get(field, 0)) + 1
            state["asked_count"] = asked_count
            state["current_questions"] = fields
            state["ready"] = False
            return {
                "type": "question",
                "text": ai_question["question"],
                "planState": sanitize_plan_state(mode, state),
                "missing_fields": determine_missing_fields(mode, state["collected"]),
            }

    if not fields:
        fields = determine_missing_fields(mode, state["collected"])[:1]

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


def recommend_by_user_data(province, major="", keyword="", subject="", subject_detail=None):
    subject_detail = subject_detail or []
    province_rows = []
    for row in USER_DATA:
        if not row.get("province") or row["province"] not in province:
            continue
        if row.get("score") is None or row.get("rank") is None:
            continue
        if subject and not subject_requirement_compatible(subject, subject_detail, row.get("category"), row.get("major")):
            continue
        province_rows.append(row)
    if not province_rows:
        return None

    exact_terms = unique_list([term for term in [major] + str(keyword or "").split(",") if str(term).strip()])
    relaxed_terms = build_major_search_terms(major=major, keyword=keyword)

    matched = [row for row in province_rows if text_matches_terms((row.get("major") or "") + " " + (row.get("school") or ""), exact_terms)]
    match_mode = "exact" if matched else ""
    if not matched and relaxed_terms:
        matched = [row for row in province_rows if text_matches_terms((row.get("major") or "") + " " + (row.get("school") or ""), relaxed_terms)]
        match_mode = "semantic" if matched else ""
    if not matched:
        if exact_terms or relaxed_terms:
            return None
        matched = list(province_rows)
        match_mode = "province"
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
        "match_mode": match_mode or "province",
        "semantic_terms": [],
    }


def bucket_rows_for_target(rows, rank=0, score=0):
    rows = list(rows or [])
    if not rows:
        return [], [], []
    if rank > 0 and any((row.get("rank") or 0) > 0 for row in rows):
        ordered = sorted(rows, key=lambda row: (row.get("rank") or 10**9, row.get("score") or 10**9))
        chong = [row for row in ordered if (row.get("rank") or 10**9) < rank][:50]
        wen = [row for row in ordered if rank <= (row.get("rank") or 10**9) <= int(rank * 1.3)][:50]
        bao = [row for row in ordered if (row.get("rank") or 0) > int(rank * 1.3)][:50]
        if chong or wen or bao:
            return chong, wen, bao
    if score > 0 and any((row.get("score") or 0) > 0 for row in rows):
        ordered = sorted(rows, key=lambda row: (abs((row.get("score") or 0) - score), -(row.get("score") or 0)))
        chong = [row for row in ordered if (row.get("score") or 0) > score][:50]
        wen = [row for row in ordered if score - 25 <= (row.get("score") or 0) <= score + 25][:50]
        bao = [row for row in ordered if 0 < (row.get("score") or 0) < score - 25][:50]
        if chong or wen or bao:
            return chong, wen, bao
    ordered = sorted(rows, key=lambda row: ((row.get("rank") or 10**9), -(row.get("score") or 0)))
    size = len(ordered)
    split_1 = max(1, size // 3)
    split_2 = max(split_1 + 1, (size * 2) // 3)
    return ordered[:split_1], ordered[split_1:split_2], ordered[split_2:]


def normalize_result_row(item):
    normalized = dict(item)
    major_text = (normalized.get("major") or "").strip()
    if re.fullmatch(r"\d{2,4}", major_text):
        normalized["major"] = f"专业组{major_text}"
    elif not major_text:
        normalized["major"] = "未细分专业"
    normalized["school"] = (normalized.get("school") or "").strip()
    return normalized


def dedupe_rows(rows):
    seen = set()
    deduped = []
    for row in rows or []:
        normalized = normalize_result_row(row)
        key = (
            normalized.get("school") or "",
            normalized.get("major") or "",
            normalized.get("year") or "",
            normalized.get("score") or 0,
            normalized.get("rank") or 0,
        )
        if key in seen:
            continue
        seen.add(key)
        deduped.append(normalized)
    return deduped


def recommend_data(province="", major="", keyword="", rank=0, score=0, subject="", subject_detail=None):
    if province:
        subject_detail = subject_detail or extract_subject_detail(subject)
        custom = recommend_by_user_data(province, major=major, keyword=keyword, subject=subject, subject_detail=subject_detail)
        if custom:
            custom["rank"] = rank
            custom["score"] = score
            custom.setdefault("match_mode", "custom")
            custom.setdefault("semantic_terms", [])
            return custom

    if not HAS_DB or not province or not (rank > 0 or score > 0):
        return {"rank": rank, "score": score, "chong": [], "wen": [], "bao": [], "source": "empty", "match_mode": "empty", "semantic_terms": []}

    conn = sqlite3.connect(DB_PATH)
    base = "province LIKE ? AND (score>0 OR rank>0)"
    params = [f"%{province}%"]
    subject_detail = subject_detail or extract_subject_detail(subject)

    def append_subject_clause(base_sql, sql_params):
        if not subject:
            return base_sql, list(sql_params)
        next_sql = base_sql + " AND (category IS NULL OR category='' OR category='综合' OR category LIKE ? OR category LIKE ? OR category LIKE ?)"
        next_params = list(sql_params)
        if subject == "物理类":
            next_params.extend(["%物理%", "%理科%", "%首选物理%"])
        elif subject == "历史类":
            next_params.extend(["%历史%", "%文科%", "%首选历史%"])
        else:
            next_params.extend(["%", "%", "%"])
        return next_sql, next_params

    base, params = append_subject_clause(base, params)

    def append_terms_clause(base_sql, sql_params, terms):
        terms = [term for term in terms or [] if term]
        if not terms:
            return base_sql, list(sql_params)
        next_sql = base_sql
        next_params = list(sql_params)
        keyword_conditions = []
        for item in terms:
            keyword_conditions.append("(major_name LIKE ? OR school_name LIKE ?)")
            next_params.append(f"%{item}%")
            next_params.append(f"%{item}%")
        next_sql += " AND (" + " OR ".join(keyword_conditions) + ")"
        return next_sql, next_params

    exact_terms = unique_list([term for term in [major] + str(keyword or "").split(",") if str(term).strip()])
    relaxed_terms = build_major_search_terms(major=major, keyword=keyword)

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
            if subject and not subject_requirement_compatible(subject, subject_detail, row[5] if len(row) > 5 else "", item["major"]):
                continue
            rows.append(item)
        return dedupe_rows(rows)

    def fetch_bucket_rows(query_base, query_params):
        local_chong = []
        local_wen = []
        local_bao = []
        if rank > 0:
            local_chong = rows_from_sql(
                f"SELECT school_name,major_name,score,rank,year,category FROM admission WHERE {query_base} "
                "AND rank>0 AND rank<? AND rank>=? ORDER BY rank ASC LIMIT 50",
                query_params + [rank, max(1, int(rank * 0.90))],
            )
            local_wen = rows_from_sql(
                f"SELECT school_name,major_name,score,rank,year,category FROM admission WHERE {query_base} "
                "AND rank>0 AND rank>=? AND rank<=? ORDER BY rank ASC LIMIT 50",
                query_params + [rank, int(rank * 1.3)],
            )
            local_bao = rows_from_sql(
                f"SELECT school_name,major_name,score,rank,year,category FROM admission WHERE {query_base} "
                "AND rank>0 AND rank>? AND rank<=? ORDER BY rank ASC LIMIT 50",
                query_params + [int(rank * 1.3), int(rank * 1.6)],
            )
        if not (local_chong or local_wen or local_bao) and score > 0:
            local_chong = rows_from_sql(
                f"SELECT school_name,major_name,score,rank,year,category FROM admission WHERE {query_base} "
                "AND score>? AND score<=? ORDER BY score DESC LIMIT 80",
                query_params + [score, score + 25],
            )
            local_wen = rows_from_sql(
                f"SELECT school_name,major_name,score,rank,year,category FROM admission WHERE {query_base} "
                "AND score>=? AND score<=? ORDER BY score ASC LIMIT 50",
                query_params + [score - 25, score + 25],
            )
            local_bao = rows_from_sql(
                f"SELECT school_name,major_name,score,rank,year,category FROM admission WHERE {query_base} "
                "AND score>=? AND score<? ORDER BY score ASC LIMIT 50",
                query_params + [score - 50, score - 25],
            )
        return local_chong, local_wen, local_bao

    def fetch_nearest_rows(query_base, query_params):
        if rank > 0:
            rows = rows_from_sql(
                f"SELECT school_name,major_name,score,rank,year,category FROM admission WHERE {query_base} "
                "AND rank>0 ORDER BY ABS(rank-?) ASC LIMIT 120",
                query_params + [rank],
            )
            if rows:
                return rows
        if score > 0:
            return rows_from_sql(
                f"SELECT school_name,major_name,score,rank,year,category FROM admission WHERE {query_base} "
                "AND score>0 ORDER BY ABS(score-?) ASC LIMIT 120",
                query_params + [score],
            )
        return []

    chong = []
    wen = []
    bao = []
    match_mode = "province"

    exact_base, exact_params = append_terms_clause(base, params, exact_terms)
    relaxed_base, relaxed_params = append_terms_clause(base, params, relaxed_terms)
    plain_base, plain_params = base, list(params)

    for label, query_base, query_params in [
        ("exact", exact_base, exact_params),
        ("semantic", relaxed_base, relaxed_params),
        ("province", plain_base, plain_params),
    ]:
        chong, wen, bao = fetch_bucket_rows(query_base, query_params)
        if chong or wen or bao:
            match_mode = label
            break
    if not (chong or wen or bao):
        nearest_rows = fetch_nearest_rows(plain_base, plain_params)
        chong, wen, bao = bucket_rows_for_target(nearest_rows, rank=rank, score=score)
        if chong or wen or bao:
            match_mode = "nearest"

    only_major_group_result = False
    if major and province == "江苏":
        group_rows = []
        major_group_base = "province LIKE ? AND (score>0 OR rank>0)"
        major_group_params = [f"%{province}%"]
        major_group_base, major_group_params = append_subject_clause(major_group_base, major_group_params)
        if rank > 0:
            group_rows = rows_from_sql(
                f"SELECT school_name,major_name,score,rank,year,category FROM admission WHERE {major_group_base} "
                "AND rank>0 AND rank>=? AND rank<=? ORDER BY rank ASC LIMIT 120",
                major_group_params + [max(1, int(rank * 0.90)), int(rank * 1.6)],
            )
        elif score > 0:
            group_rows = rows_from_sql(
                f"SELECT school_name,major_name,score,rank,year,category FROM admission WHERE {major_group_base} "
                "AND score>=? AND score<=? ORDER BY score ASC LIMIT 120",
                major_group_params + [score - 50, score + 25],
            )
        group_rows = [row for row in group_rows if is_major_group_text(row.get("major"))]
        if group_rows:
            only_major_group_result = True
            chong, wen, bao = bucket_rows_for_target(group_rows, rank=rank, score=score)
            match_mode = "province"

    conn.close()
    if match_mode == "province" and exact_terms and (chong or wen or bao):
        match_mode = "direct"
    return {
        "rank": rank,
        "score": score,
        "chong": chong,
        "wen": wen,
        "bao": bao,
        "source": "db",
        "match_mode": match_mode,
        "semantic_terms": [],
    }


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
        if recommendations.get("match_mode") == "semantic" and recommendations.get("semantic_terms"):
            summary.append(f"直接关键词没命中后，我又用 AI 补了这些相近检索词：{','.join(recommendations.get('semantic_terms') or [])}。")
        elif recommendations.get("match_mode") in ("province", "nearest", "direct"):
            summary.append("这次结果里包含一部分省内兜底数据，命中精度会比直接专业匹配弱一些。")
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
        if info.get("subject_detail"):
            profile_parts.append(f"选科细项：{','.join(info.get('subject_detail') or [])}")
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
            # 正则/字典优先，匹配不到再调 AI
            extracted = fallback_extract_info(mode, message, plan_state.get("current_questions") or [])
            if mode == "gaokao":
                regex_got_something = bool(
                    extracted.get("province") or extracted.get("score") or extracted.get("rank")
                    or extracted.get("subject") or extracted.get("majors") or extracted.get("major_unknown")
                )
            else:
                regex_got_something = bool(extracted.get("topic") or extracted.get("goal"))
            if not regex_got_something:
                # 正则打不中 → 调 AI 提取
                try:
                    ai_result = ai_extract_fields(
                        mode, message, history,
                        plan_state.get("current_questions") or [],
                        plan_state.get("collected") or {},
                        config,
                    )
                    if ai_result:
                        extracted = ai_result
                except Exception:
                    pass
            else:
                # 正则打中了 → AI 验证，有问题则 AI 修正
                verify_result = ai_verify_extraction(
                    mode, message, extracted,
                    plan_state.get("collected") or {},
                    history, config,
                )
                if verify_result is not None and not verify_result.get("correct"):
                    issues = verify_result.get("issues") or []
                    fixed = ai_fix_extraction(
                        mode, message, extracted, issues,
                        plan_state.get("collected") or {},
                        history, config,
                    )
                    if fixed:
                        extracted = fixed
            plan_state["collected"] = merge_collected_for_mode(mode, plan_state["collected"], extracted)
            plan_state["last_user_message"] = message
            plan_state = prune_answered_questions(mode, plan_state)

            plan_enabled = bool(planner.get("enabled"))
            if planner_should_ask(mode, plan_enabled, plan_state, message):
                question_payload = build_question_response(mode, plan_state, config=config, history=history, message=message)
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
                    subject_detail=collected.get("subject_detail", []),
                )
                total_rows = sum(len(recommendations.get(bucket) or []) for bucket in ["chong", "wen", "bao"])
                should_try_semantic = total_rows == 0 or (
                    recommendations.get("match_mode") in ("province", "nearest", "direct") and bool(collected.get("majors"))
                )
                if should_try_semantic:
                    semantic_terms = []
                    try:
                        semantic_terms = ai_expand_retrieval_terms(collected, history, config)
                    except Exception:
                        semantic_terms = []
                    if semantic_terms:
                        emit({"type": "phase", "text": "正在用 AI 补全相近专业词后再试一次..."})
                        merged_keyword = ",".join(unique_list((collected.get("majors") or []) + semantic_terms))
                        fallback_recommendations = recommendations
                        recommendations = recommend_data(
                            province=collected.get("province", ""),
                            major=(collected.get("majors") or [""])[0] if collected.get("majors") else "",
                            keyword=merged_keyword,
                            rank=collected.get("rank", 0),
                            score=collected.get("score", 0),
                            subject=collected.get("subject", ""),
                            subject_detail=collected.get("subject_detail", []),
                        )
                        semantic_total = sum(len(recommendations.get(bucket) or []) for bucket in ["chong", "wen", "bao"])
                        if semantic_total:
                            recommendations["semantic_terms"] = semantic_terms
                            recommendations["match_mode"] = "semantic"
                        else:
                            recommendations = fallback_recommendations
                            recommendations["semantic_terms"] = semantic_terms
                            recommendations["match_mode"] = recommendations.get("match_mode", "province")

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
