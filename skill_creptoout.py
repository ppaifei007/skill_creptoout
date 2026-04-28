import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET


_RE_NUMBER = re.compile(r"\d[\d,]*([.]\d+)?")
_RE_LIST_MARKER = re.compile(r"^\s*\d+$")
_RE_PAGINATION = re.compile(r"^\s*\d*\s*[\/／]\s*\d+\s*$")
_RE_SHORT_SERIAL = re.compile(r"^\s*0?\d{1,2}\s*$")
_RE_TITLE_PREFIX = re.compile(r"^\s*\d+(?:\.\d+){0,2}\s*")
_RE_TITLE_SERIAL = re.compile(r"^\d+(?:\.\d+){1,2}$")
_TARGET_CONTEXT_KEYWORDS = (
    "目录",
    "目标",
    "工作部署",
    "提价",
    "结算",
    "完成",
    "实现",
    "完善",
    "建强",
    "升级",
    "推进",
    "迭代",
    "争取",
    "规范",
    "自主",
    "纳管",
    "整改",
    "看管",
    "关键人",
    "联系人",
    "责任人",
    "体系",
    "能力",
    "工程",
    "步法",
    "消息转型",
)
_SENSITIVE_METRIC_KEYWORDS = (
    "收入",
    "增幅",
    "比重",
    "贡献",
    "份额",
    "占比",
    "纳管",
    "未纳管",
    "新增",
    "完备率",
    "渗透率",
    "渗透提升",
    "达标率",
    "保有率",
    "转化率",
    "毛利率",
    "签约",
    "中标",
    "储备率",
    "净值",
    "规模",
    "金额",
    "回款",
    "商机",
    "支撑",
    "交付",
    "应验",
    "未验",
    "员工",
    "队伍",
    "完成率",
    "目标值",
    "到达值",
)
_STRATEGIC_SLOGAN_KEYWORDS = (
    "经营发展目标",
    "工作思路",
    "锚定",
    "三个提升",
    "一个原则",
    "1个原则",
    "2大运营",
    "3大主业",
    "4项保障",
    "10项重点工作",
    "六个方面工作",
    "八项重点任务",
)
_RESULT_ORIENTED_VERBS = (
    "实现",
    "完成",
    "达标",
    "达成",
    "累计",
    "已",
    "超",
    "完成率",
    "同比",
    "增幅",
    "收入",
    "签约",
    "回款",
)
_KEEP_NEAR_KEYWORDS = (
    "回顾",
    "奋进",
    "经营回顾",
    "工作部署",
    "工程",
    "体系",
    "步法",
    "关键人",
    "联系人",
    "责任人",
    "自主能力",
    "看管",
    "本月",
    "当月",
    "全年",
    "上半年",
    "下半年",
    "一季度",
    "二季度",
    "三季度",
    "四季度",
    "自主",
)
_DIRECTORY_TITLE_KEYWORDS = (
    "目录",
    "整体回顾",
    "工作部署",
    "智算专题",
    "重点分析及工作部署",
)
_TITLE_TEXT_KEYWORDS = (
    "不足",
    "推进",
    "攻坚",
    "跃升",
    "提升",
    "专题",
)
_TITLE_HINT_KEYWORDS = (
    "竞争情况",
    "收入结构",
    "重点指标",
    "重点工作",
    "下阶段重点工作",
    "市场份额",
    "增量份额",
    "收入完成率",
    "收入增幅",
    "客群效益情况",
    "发展形势",
    "发展先机",
    "核心生产要素",
    "政策利好",
    "三大主业之一",
    "算力办公室",
    "行业第一",
    "行业领先",
    "KPI情况",
    "各地市",
    "工作部署",
    "整体回顾",
)
_TITLE_SUFFIXES = (
    "趋势",
    "增幅",
    "完成率",
    "情况",
    "形势",
    "先机",
    "生产要素",
    "政策利好",
    "算力办公室",
)
_COVER_SLIDE_HINT_KEYWORDS = (
    "培训",
    "巡前培训",
    "客户部",
    "政企客户部",
)
_CONSISTENCY_TABLE_KEYWORDS = (
    "商机类型",
    "数量",
    "项目预算",
    "预计签约",
    "预计IT主营",
)
_CONSISTENCY_IT_RESULT_TABLE_KEYWORDS = (
    "商机类型",
    "数量",
    "项目预算",
    "预计签约",
    "IT主营收入",
)
_CONSISTENCY_PROGRESS_KEYWORDS = (
    "拓展进度",
    "已拓展数",
    "商家总数",
)
_CONSISTENCY_SMART_COMPUTE_KEYWORDS = (
    "智算项目年度签约目标",
    "当前跟进商机",
    "商机金额",
    "储备率",
    "全年签约目标",
    "二季度商机金额",
    "商机数量",
)
_REGIONAL_COMPLETION_SLIDE_KEYWORDS = (
    "各地分公司累计收入及一季度完成情况",
    "一季度收入完成率",
    "政企有效收入",
    "净值情况",
)
_GOAL_TARGET_SLIDE_KEYWORDS = (
    "政企工作思路",
    "经营发展目标",
    "能力建设目标",
    "政企客群收入目标",
)
_LABOR_SCORE_TABLE_KEYWORDS = (
    "劳动竞赛",
    "总得分",
    "组内排名",
    "政企收入",
    "客群运营",
    "通信业务",
    "算力服务",
    "智能服务",
)
_IOT_RESULT_SLIDE_KEYWORDS = (
    "蜂窝物联收入",
    "各地市计费用户情况",
    "卡+模组",
    "各地市物联网应用情况",
)
_COMPUTE_RESULT_SLIDE_KEYWORDS = (
    "智算定制服务项目进展",
    "智算N节点销售/借货情况",
    "已立项项目金额",
    "季度签约金额",
    "1季度预计可落收",
    "线上销售",
    "线下销售",
    "线下借货",
)
_PROJECT_TABLE_SLIDE_KEYWORDS = (
    "千万大单",
    "项目名称",
    "中标方",
    "丢标原因",
)
_PROJECT_TABLE_HEADERS = (
    "序号",
    "地市",
    "项目名称",
    "中标方",
    "丢标金额",
    "友商金额",
    "丢标原因",
    "商机类型",
    "数量",
    "项目预算",
    "预计签约",
    "IT主营收入",
    "合计",
    "总计",
)
_PROJECT_TABLE_STATIC_TEXTS = (
    "友商存量",
    "行业平衡",
    "优势",
    "竞争",
    "劣势",
)
_CITY_LABELS = (
    "省战客",
    "战客",
    "全省",
    "杭州",
    "宁波",
    "温州",
    "金华",
    "台州",
    "嘉兴",
    "绍兴",
    "湖州",
    "丽水",
    "衢州",
    "舟山",
)
_CONSISTENCY_KPI_TABLE_KEYWORDS = (
    "KPI",
    "目标值",
    "到达值",
    "完成率",
)
_KPI_TABLE_EXTRA_KEYWORDS = (
    "完成情况",
    "迁转值",
    "全省",
    "地市",
)
_SENSITIVE_TABLE_RESULT_KEYWORDS = (
    "总规模",
    "机架数",
    "计费机架数",
    "上架率",
)
_TIME_SENSITIVE_VALUE_KEYWORDS = (
    "完成值",
    "计费功率",
    "下订单",
    "新增订单",
)
_TABLE_LINKAGE_RESULT_KEYWORDS = (
    "收入",
    "份额",
    "增幅",
    "同比",
    "完成值",
    "目标值",
    "到达值",
    "完成率",
    "位列",
    "领先",
    "行业第一",
)
_TABLE_LINKAGE_STRUCTURE_KEYWORDS = (
    "单位",
    "目标值",
    "到达值",
    "完成率",
    "达成情况",
    "完成情况",
)
_SENSITIVE_UNITS = (
    "%",
    "％",
    "亿",
    "万",
    "人",
    "条",
    "户",
    "家",
    "倍",
    "朵",
)
_SENSITIVE_WORD_UNITS = (
    "PP",
    "pp",
)
_TABLE_RESULT_HINT_KEYWORDS = (
    "收入",
    "增幅",
    "同比",
    "环比",
    "占比",
    "份额",
    "贡献",
    "纳管",
    "渗透率",
    "完成率",
    "到达值",
    "目标值",
    "净增",
    "拓展",
    "对标",
    "行业",
    "地市",
    "同组",
    "结构",
    "覆盖",
    "走访",
    "全省",
    "战客",
)
_LABOR_WEIGHT_KEEP_KEYWORDS = (
    "总得分",
    "政企收入",
    "客群运营",
    "通信业务",
    "算力服务",
    "智能服务",
)
_SENSITIVE_NAMES = (
    "字节跳动",
    "字节",
    "百度",
    "拼多多",
    "阿里云",
    "阿里",
    "腾讯",
    "京东",
    "金山",
    "快手",
    "哔哩哔哩",
    "蚂蚁金服",
    "蚂蚁",
    "华为",
    "华为云",
)
_ORG_NAME_KEYWORDS = (
    "有限公司",
    "股份有限公司",
    "公司",
    "银行",
    "分行",
    "分局",
    "管理局",
    "工作部",
    "委员会",
    "办公室",
    "开发区",
    "农村商业",
    "集团",
    "政府",
    "学校",
    "学院",
    "大学",
    "医院",
)
_GENERIC_LABEL_KEYWORDS = (
    "集团",
    "成员",
    "低占",
    "低渗",
    "目标",
    "清单",
    "行动",
    "家数",
    "占比",
    "达标",
    "宽带",
    "纳管",
    "迎回",
    "筑堤",
    "头雁",
    "省直管",
    "地市",
)
_GENERIC_ORG_PHRASES = (
    "集团公司",
    "省公司",
    "市公司",
    "总部",
    "分公司",
)


@dataclass
class RedactStats:
    files_touched: int = 0
    text_nodes_changed: int = 0
    chart_numfmt_set: int = 0


def _local(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _redact_sensitive_names(s: str) -> str:
    for name in _SENSITIVE_NAMES:
        s = s.replace(name, "xxx")
    return s


def _looks_like_customer_name(s: str) -> bool:
    text = s.strip()
    if not text:
        return False
    if len(text) < 4:
        return False
    if "客户名称" in text or "集团名称" in text:
        return False
    if any(ch.isdigit() for ch in text):
        return False
    if any(p in text for p in _GENERIC_ORG_PHRASES):
        return False
    if any(sep in text for sep in ("，", "、", ",", "|", " ")):
        return False
    if any(
        k in text
        for k in (
            "有限公司",
            "股份有限公司",
            "公司",
            "银行",
            "分行",
            "分局",
            "管理局",
            "工作部",
            "委员会",
            "办公室",
            "开发区",
            "学校",
            "学院",
            "大学",
            "医院",
        )
    ):
        return True
    if any(k in text for k in _GENERIC_LABEL_KEYWORDS):
        return False
    if any(k in text for k in ("集团", "政府", "农村商业")):
        return True
    return False


def _is_title_text(s: str, slide_text: str = "") -> bool:
    text = s.strip()
    if not text:
        return False
    compact = text.replace(" ", "")
    if len(compact) > 28:
        return False
    if compact.startswith(("•", "-", "1.", "2.", "3.", "4.")):
        return False
    if any(p in compact for p in ("。", "；", ";")):
        return False
    if any(k in compact for k in _TITLE_TEXT_KEYWORDS):
        return True
    if any(k in compact for k in _TITLE_HINT_KEYWORDS):
        return True
    if compact.endswith(_TITLE_SUFFIXES) and len(compact) <= 20:
        return True
    if any(ch.isdigit() for ch in text):
        return False
    if ("：" in compact or ":" in compact) and len(compact) <= 24:
        return True
    if "目录" in slide_text and compact in _DIRECTORY_TITLE_KEYWORDS:
        return True
    return False


def _is_title_paragraph(s: str, slide_text: str = "") -> bool:
    text = s.strip()
    if not text:
        return False
    compact = text.replace(" ", "")
    if len(compact) > 32:
        return False
    if any(p in compact for p in ("。", "；", ";")):
        return False
    core = _RE_TITLE_PREFIX.sub("", compact)
    if not core:
        return False
    if any(ch.isdigit() for ch in core):
        return False
    if any(k in core for k in _TITLE_HINT_KEYWORDS):
        return True
    if core.endswith(_TITLE_SUFFIXES) and len(core) <= 28:
        return True
    if ("：" in core or ":" in core) and len(core) <= 24:
        return True
    if "目录" in slide_text and core in _DIRECTORY_TITLE_KEYWORDS:
        return True
    return False


def _is_cover_slide(slide_text: str) -> bool:
    compact = slide_text.replace(" ", "")
    if len(compact) > 80:
        return False
    if not any(k in compact for k in _COVER_SLIDE_HINT_KEYWORDS):
        return False
    if "年" not in compact or "月" not in compact:
        return False
    return True


def _next_non_space(s: str, start: int) -> str:
    return s[start:].lstrip()


def _prev_non_space_char(s: str, end: int) -> str:
    for i in range(end - 1, -1, -1):
        if not s[i].isspace():
            return s[i]
    return ""


def _is_table_linkage_slide(slide_text: str) -> bool:
    result_hits = sum(1 for k in _TABLE_LINKAGE_RESULT_KEYWORDS if k in slide_text)
    structure_hits = sum(1 for k in _TABLE_LINKAGE_STRUCTURE_KEYWORDS if k in slide_text)
    return result_hits >= 2 and structure_hits >= 3


def _is_result_context(around: str, slide_text: str = "") -> bool:
    context = around + slide_text
    return any(k in context for k in _RESULT_ORIENTED_VERBS)


def _is_business_result_slide(slide_text: str) -> bool:
    hit = sum(1 for k in _TABLE_RESULT_HINT_KEYWORDS if k in slide_text)
    return hit >= 2


def _is_strategic_slogan_percent(around: str) -> bool:
    compact = around.replace(" ", "")
    if "三个" in compact and any(k in compact for k in ("经营发展目标", "三个提升")):
        return True
    if any(k in compact for k in ("1个原则", "2大运营", "3大主业", "4项保障", "10项重点工作")):
        return True
    if any(k in compact for k in ("六个方面工作", "八项重点任务")):
        return True
    return False


def _iter_paragraphs_with_context(root: ET.Element) -> list[tuple[ET.Element, bool]]:
    out: list[tuple[ET.Element, bool]] = []

    def walk(node: ET.Element, in_table: bool) -> None:
        local = _local(node.tag)
        now_in_table = in_table or local == "tc"
        if local == "p":
            out.append((node, now_in_table))
        for ch in list(node):
            walk(ch, now_in_table)

    walk(root, False)
    return out


def _is_project_table_text(s: str, slide_text: str = "") -> bool:
    text = s.strip().strip("“”\"'")
    if not text:
        return False
    if not all(k in slide_text for k in _PROJECT_TABLE_SLIDE_KEYWORDS):
        return False
    if text in _PROJECT_TABLE_HEADERS or text in _PROJECT_TABLE_STATIC_TEXTS or text in _CITY_LABELS:
        return False
    if len(text) > 40:
        return False
    if any(p in text for p in ("：", "；", "。")):
        return False
    if not any("\u4e00" <= ch <= "\u9fff" for ch in text):
        return False
    if any(
        k in text
        for k in (
            "公司",
            "项目",
            "工程",
            "平台",
            "服务",
            "监控",
            "建设",
            "运维",
            "局",
            "分局",
            "电信",
            "联通",
            "华数",
            "数据",
        )
    ):
        return True
    return False


def _is_project_name_cell(s: str, slide_text: str = "", in_table: bool = False) -> bool:
    text = s.strip().strip("“”\"'")
    if not text:
        return False
    if not in_table:
        return False
    if "项目名称" not in slide_text:
        return False
    if any(k in slide_text for k in ("中标金额", "业主单位", "地市", "月份")):
        if text in _PROJECT_TABLE_HEADERS or text in _CITY_LABELS:
            return False
        if any(k in text for k in ("项目名称", "中标金额", "业主单位", "地市", "月份")):
            return False
        if len(text) > 60:
            return False
        if not any("\u4e00" <= ch <= "\u9fff" for ch in text):
            return False
        return True
    return False


def _is_project_name_context_text(s: str, slide_text: str = "", in_table: bool = False) -> bool:
    text = s.strip().strip("“”\"'")
    if not text:
        return False
    if in_table:
        return False
    hint_hits = sum(
        1
        for k in ("项目名称", "中标金额", "业主单位", "千万商机下千万元单商机", "项目")
        if k in slide_text
    )
    if hint_hits < 2:
        return False
    if len(text) < 6 or len(text) > 60:
        return False
    if not any("\u4e00" <= ch <= "\u9fff" for ch in text):
        return False
    if any(k in text for k in ("月份", "地市", "业主单位", "中标金额", "项目名称")):
        return False
    if any(k in text for k in ("项目", "工程", "系统", "平台", "建设", "服务", "改造")):
        return True
    return False


def _should_redact_number(
    s: str,
    start: int,
    end: int,
    slide_text: str = "",
    underlined: bool = False,
    in_table: bool = False,
) -> bool:
    prev_ch = s[start - 1] if start > 0 else ""
    next_ch = s[end] if end < len(s) else ""
    prev_sig = _prev_non_space_char(s, start)
    next_non_space = _next_non_space(s, end)
    next_sig = next_non_space[:1]
    next_sig2 = next_non_space[:2]
    token = s[start:end]
    around = s[max(0, start - 12) : min(len(s), end + 12)]
    left6 = s[max(0, start - 6) : start]

    if _RE_PAGINATION.match(s.strip()):
        return False

    if start == 0 and _RE_TITLE_SERIAL.match(token) and any(
        k in slide_text for k in _TITLE_HINT_KEYWORDS
    ):
        return False

    if (
        start == 0
        and _RE_TITLE_SERIAL.match(token)
        and len(s.strip()) <= 26
        and next_sig not in _SENSITIVE_UNITS
        and next_sig2 not in _SENSITIVE_WORD_UNITS
        and "，" not in s
        and "," not in s
        and "。" not in s
    ):
        return False

    if start == 0 and _RE_SHORT_SERIAL.match(token) and next_ch in (
        ".",
        "、",
        ")",
        "）",
        "．",
        ":",
        "：",
    ):
        return False

    if "目录" in slide_text and any(k in slide_text for k in _DIRECTORY_TITLE_KEYWORDS):
        if _RE_SHORT_SERIAL.match(s.strip()):
            return False

    if _RE_SHORT_SERIAL.match(s.strip()) and "目录" in slide_text:
        return False

    if prev_ch in ("/", "／") or next_ch in ("/", "／"):
        return False

    if left6.endswith("路径") and next_ch in ("：", ":"):
        return False

    if len(token) == 4 and any(k in around for k in ("回顾", "奋进")):
        return False

    if next_sig in ("年", "月", "日", "号", "周", "季", "天"):
        return False

    if next_sig in ("-", "－", "~", "至") and any(
        c in next_non_space[:6] for c in ("月", "年")
    ):
        return False

    if token in ("2", "3", "4", "5") and next_sig == "G":
        return False

    if next_non_space.startswith("KPI") and "." in token:
        return False

    if prev_sig == "注":
        return False

    if underlined:
        return True

    if next_sig in ("%", "％"):
        # Only use local text window for slogan detection to avoid slide-level false keeps.
        if _is_result_context(around, ""):
            return True
        if _is_strategic_slogan_percent(around):
            return False
        if _is_business_result_slide(slide_text):
            return True

    if "PP" in around.upper():
        return True

    if next_sig in ("个", "大", "项") and any(
        k in around for k in ("原则", "运营", "主业", "保障", "重点工作")
    ):
        return False

    if next_sig in ("条", "项") and any(
        k in around for k in ("主线", "路径", "工作规划", "工作计划", "拓展路径")
    ):
        return False

    if (
        in_table
        and next_sig == "分"
        and any(k in around for k in _LABOR_WEIGHT_KEEP_KEYWORDS)
    ):
        return False

    if in_table and _is_business_result_slide(slide_text):
        return True

    if any(k in around for k in _TIME_SENSITIVE_VALUE_KEYWORDS):
        return True

    if all(k in slide_text for k in _LABOR_SCORE_TABLE_KEYWORDS):
        return True

    if all(k in slide_text for k in _IOT_RESULT_SLIDE_KEYWORDS):
        return True

    if all(k in slide_text for k in _COMPUTE_RESULT_SLIDE_KEYWORDS):
        return True

    if all(k in slide_text for k in _REGIONAL_COMPLETION_SLIDE_KEYWORDS):
        return True

    if all(k in slide_text for k in _GOAL_TARGET_SLIDE_KEYWORDS) and any(
        k in around for k in ("目标", "增幅", "贡献", "份额")
    ):
        return True

    if all(k in slide_text for k in _CONSISTENCY_KPI_TABLE_KEYWORDS) and any(
        k in slide_text for k in _KPI_TABLE_EXTRA_KEYWORDS
    ):
        return True

    # Result-oriented slides with tables should redact linked numeric cells consistently.
    if _is_table_linkage_slide(slide_text) and any(
        k in slide_text
        for k in (
            *_SENSITIVE_METRIC_KEYWORDS,
            *_TIME_SENSITIVE_VALUE_KEYWORDS,
        )
    ):
        return True

    if all(k in slide_text for k in _CONSISTENCY_TABLE_KEYWORDS):
        if any(k in slide_text for k in ("商机", "预算", "签约", "IT主营")):
            return True

    if all(k in slide_text for k in _CONSISTENCY_IT_RESULT_TABLE_KEYWORDS):
        return True

    if all(k in slide_text for k in _CONSISTENCY_PROGRESS_KEYWORDS):
        return True

    if any(k in slide_text for k in _CONSISTENCY_SMART_COMPUTE_KEYWORDS):
        if any(k in slide_text for k in _SENSITIVE_TABLE_RESULT_KEYWORDS):
            return True
        if any(k in slide_text for k in ("全年签约目标", "二季度商机金额", "商机数量")):
            return True


    if next_ch in ("个", "位") and any(
        k in around for k in ("关键人", "联系人", "责任人", "看管")
    ):
        return False

    if _is_business_result_slide(slide_text):
        if re.fullmatch(r"\s*[-+]?\d[\d,]*([.]\d+)?\s*", s):
            return True

    if next_sig in ("G", "T") and any(k in around for k in ("带宽", "裁撤", "TOP客户", "新增")):
        return True

    if any(k in around for k in _SENSITIVE_METRIC_KEYWORDS):
        return True

    if next_sig in _SENSITIVE_UNITS or next_sig2 in _SENSITIVE_WORD_UNITS:
        return True

    if prev_sig == "(" or prev_sig == "（":
        bracket_tail = s[end : min(len(s), end + 16)]
        if any(k in bracket_tail for k in ("%", "％", "PP", "pp", "亿", "万", "倍")):
            return True

    if prev_sig in ("年", "月", "日", "号"):
        return False

    if any(k in around for k in _KEEP_NEAR_KEYWORDS):
        return False

    left = s[:start]
    if _RE_LIST_MARKER.match(left) and next_ch in (".", "、", ")", "）", "．", ":", "："):
        return False

    if (prev_ch == "（" and next_ch == "）") or (prev_ch == "(" and next_ch == ")"):
        prefix = s[: start - 1]
        if not prefix or prefix[-1].isspace() or prefix[-1] in (
            "\n",
            "\r",
            "。",
            "；",
            ";",
            "：",
            ":",
        ):
            return False

    if prev_ch == "第" and next_ch in ("章", "节", "条", "款", "项", "页", "步"):
        return False

    left4 = s[max(0, start - 6) : start]
    for p in ("图", "表", "式", "附件", "编号", "序号", "路径", "注", "No.", "NO.", "No", "NO", "ID", "Id", "id"):
        if left4.endswith(p):
            return False

    if prev_ch and prev_ch.isascii() and prev_ch.isalnum():
        return False

    return False


def _redact_text(
    s: str,
    full_text: str | None = None,
    base_offset: int = 0,
    slide_text: str = "",
    underlined: bool = False,
    in_table: bool = False,
) -> str:
    context = full_text if full_text is not None else s

    def repl(m: re.Match[str]) -> str:
        start = base_offset + m.start()
        end = base_offset + m.end()
        if _should_redact_number(
            context,
            start,
            end,
            slide_text=slide_text,
            underlined=underlined,
            in_table=in_table,
        ):
            return "xxx"
        return m.group(0)

    return _RE_NUMBER.sub(repl, s)


def _paragraph_text_nodes(para: ET.Element) -> list[tuple[ET.Element, bool]]:
    nodes: list[tuple[ET.Element, bool]] = []
    for run in para.iter():
        if _local(run.tag) != "r":
            continue
        underlined = False
        for child in list(run):
            if _local(child.tag) != "rPr":
                continue
            if child.attrib.get("u") not in (None, "none"):
                underlined = True
            for sub in child.iter():
                if _local(sub.tag) == "u":
                    underlined = True
        for child in run.iter():
            if _local(child.tag) == "t" and child.text:
                nodes.append((child, underlined))
    if nodes:
        return nodes
    return [(el, False) for el in para.iter() if _local(el.tag) == "t" and el.text]


def _xml_redact_a_t(xml_bytes: bytes) -> tuple[bytes, int]:
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return xml_bytes, 0
    changed = 0
    doc_text_nodes = [el for el in root.iter() if _local(el.tag) == "t" and el.text]
    doc_text = "".join(el.text or "" for el in doc_text_nodes)
    for para, in_table in _iter_paragraphs_with_context(root):

        text_nodes = _paragraph_text_nodes(para)
        if not text_nodes:
            continue

        full_text = "".join(el.text or "" for el, _ in text_nodes)
        paragraph_is_title = _is_title_paragraph(full_text, slide_text=doc_text)
        cover_slide = _is_cover_slide(doc_text)
        offset = 0
        for el, underlined in text_nodes:
            original = el.text or ""
            is_title_like = paragraph_is_title or _is_title_text(original, slide_text=doc_text)
            if cover_slide:
                offset += len(original)
                continue
            if (
                in_table
                and _is_business_result_slide(doc_text)
                and not is_title_like
                and not any(k in full_text for k in _LABOR_WEIGHT_KEEP_KEYWORDS)
            ):
                # Fallback for embedded table containers where context matching is weak:
                # redact all numeric tokens in result tables, while preserving common non-sensitive tech labels.
                if re.search(r"(5G|2G|6\+N|UCU|20\d{2}年|\d{1,2}月)", original):
                    new_text = _redact_text(
                        original,
                        full_text=full_text,
                        base_offset=offset,
                        slide_text=doc_text,
                        underlined=underlined,
                        in_table=in_table,
                    )
                else:
                    new_text = _RE_NUMBER.sub("xxx", original)
            else:
                new_text = _redact_text(
                    original,
                    full_text=full_text,
                    base_offset=offset,
                    slide_text=doc_text,
                    underlined=underlined,
                    in_table=in_table,
                )
            # Keep directory item titles readable: do not whole-mask them as customer/project names.
            if "目录" in doc_text and is_title_like:
                if new_text != original:
                    el.text = new_text
                    changed += 1
                offset += len(original)
                continue
            new_text = _redact_sensitive_names(new_text)
            if _looks_like_customer_name(new_text):
                new_text = "xxx"
            elif in_table and _is_project_table_text(new_text, slide_text=doc_text):
                new_text = "xxx"
            elif _is_project_name_cell(new_text, slide_text=doc_text, in_table=in_table):
                new_text = "xxx"
            elif _is_project_name_context_text(new_text, slide_text=doc_text, in_table=in_table):
                new_text = "xxx"
            if new_text != original:
                el.text = new_text
                changed += 1
            offset += len(original)

    # Force-fix residual PP expressions such as "2.3PP" that may escape title/context branches.
    for el in root.iter():
        if _local(el.tag) != "t" or not el.text:
            continue
        txt = el.text
        if re.search(r"\d[\d,]*([.]\d+)?\s*(PP|pp)", txt):
            if _is_strategic_slogan_percent(txt):
                continue
            new_txt = _RE_NUMBER.sub("xxx", txt)
            if new_txt != txt:
                el.text = new_txt
                changed += 1

    if changed == 0:
        return xml_bytes, 0
    out = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return out, changed


def _chart_set_numfmt(xml_bytes: bytes) -> tuple[bytes, int]:
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return xml_bytes, 0

    changed = 0

    def walk(node: ET.Element, ancestors: list[str]) -> None:
        nonlocal changed
        local = _local(node.tag)
        ancestors2 = [*ancestors, local]

        if local == "t" and ("dLbl" in ancestors2 or "dLbls" in ancestors2) and node.text:
            new_text = _RE_NUMBER.sub("xxx", node.text)
            if new_text != node.text:
                node.text = new_text
                changed += 1

        if local == "numFmt":
            if "catAx" not in ancestors2:
                fmt = node.attrib.get("formatCode")
                if fmt != '"xxx"' or node.attrib.get("sourceLinked") != "0":
                    node.set("formatCode", '"xxx"')
                    node.set("sourceLinked", "0")
                    changed += 1
            return

        for ch in list(node):
            walk(ch, ancestors2)

        if local in ("dLbl", "dLbls", "valAx") and "catAx" not in ancestors2:
            has_numfmt = any(_local(ch.tag) == "numFmt" for ch in list(node))
            if not has_numfmt:
                numfmt = ET.Element(
                    node.tag[: node.tag.find("}") + 1] + "numFmt"
                    if "}" in node.tag
                    else "numFmt"
                )
                numfmt.set("formatCode", '"xxx"')
                numfmt.set("sourceLinked", "0")
                node.append(numfmt)
                changed += 1

    walk(root, [])
    if changed == 0:
        return xml_bytes, 0
    out = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return out, changed


def redact_pptx(in_pptx: Path, out_pptx: Path) -> RedactStats:
    stats = RedactStats()
    with zipfile.ZipFile(in_pptx, "r") as zin, zipfile.ZipFile(out_pptx, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            touched = False

            if info.filename.lower().endswith(".xml"):
                new_data, c = _xml_redact_a_t(data)
                if c:
                    stats.text_nodes_changed += c
                    data = new_data
                    touched = True

                if info.filename.startswith("ppt/charts/") and info.filename.lower().endswith(".xml"):
                    new_data, c2 = _chart_set_numfmt(data)
                    if c2:
                        stats.chart_numfmt_set += c2
                        data = new_data
                        touched = True

            if touched:
                stats.files_touched += 1

            zout.writestr(info, data)
    return stats


def scan_pptx_for_digits(pptx: Path) -> tuple[int, int]:
    digits = 0
    nodes = 0
    with zipfile.ZipFile(pptx, "r") as z:
        for name in z.namelist():
            if not name.lower().endswith(".xml"):
                continue
            data = z.read(name)
            try:
                root = ET.fromstring(data)
            except ET.ParseError:
                continue
            for el in root.iter():
                if _local(el.tag) == "t" and el.text:
                    nodes += 1
                    if re.search(r"\d", el.text):
                        digits += 1
    return digits, nodes


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("in_pptx", type=Path)
    ap.add_argument("out_pptx", type=Path)
    ap.add_argument("--verify", action="store_true")
    args = ap.parse_args()

    stats = redact_pptx(args.in_pptx, args.out_pptx)
    print(
        f"done files_touched={stats.files_touched} "
        f"text_nodes_changed={stats.text_nodes_changed} "
        f"chart_numfmt_set={stats.chart_numfmt_set}"
    )
    if args.verify:
        d, n = scan_pptx_for_digits(args.out_pptx)
        print(f"verify a:t_nodes={n} nodes_with_digits={d}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
