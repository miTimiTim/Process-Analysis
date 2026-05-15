#!/usr/bin/env python3
"""
lzpc analyzer - 廉洁风险排查流程分析工具
解析各单位自查表格，生成应然-实然对照、风险评级、现场核验手册等分析报告。

用法：
    python analyzer.py <输入Excel路径> [输出目录]
    输出目录默认在输入文件同目录下自动生成。
"""

import json
import re
import sys
import difflib
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

try:
    from docx import Document
    from docx.shared import Pt, Inches, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("错误: 需要安装 openpyxl。请执行: pip install openpyxl")
    sys.exit(1)


# ============================================================================
# 配置常量
# ============================================================================

SCRIPT_DIR = Path(__file__).resolve().parent
ALIASES_PATH = SCRIPT_DIR / "aliases.json"

# 风险评分权重
RISK_WEIGHTS = {
    "node_completeness": 0.25,    # 节点完整度
    "position_clarity": 0.20,     # 岗位清晰度
    "power_concentration": 0.20,  # 权力集中度
    "ought_is_match": 0.20,       # 应然-实然匹配度
    "self_awareness": 0.15,       # 自查风险自觉度（反向：自查越空洞风险越高）
}

# 风险等级阈值
RISK_THRESHOLDS = [
    (4.0, "A", "极高风险"),
    (3.0, "B", "高风险"),
    (2.0, "C", "中等风险"),
    (0.0, "D", "低风险"),
]

# 差异类型及优先级
DIFF_PRIORITY = {
    "节点消失": ("红", 1),
    "岗位偏移": ("橙", 2),
    "节点新增": ("黄", 3),
    "名称差异": ("蓝", 4),
    "一致": ("-", 5),
}

# 职位关键词（用于从节点描述中提取岗位）
POSITION_KEYWORDS = [
    "专责", "主管", "经理", "主任", "部长", "班长", "组长", "负责人",
    "分管领导", "领导", "审批人", "经办人", "管理员", "专员", "调度",
    "总工", "工程师", "会计", "出纳", "采购", "仓管", "质检", "安全员",
    "施工员", "监理", "设计", "客服", "文员", "所长", "厂长", "书记",
    "党委", "班子", "纪委", "总监", "副总", "总经理", "董事长",
]

# 权力节点类型标记
AUTHORITY_NODE_PATTERNS = {
    "发起": ["申请", "提出", "报修", "报装", "受理", "接单", "接报", "申报", "需求", "发起", "提交", "登记"],
    "审批": ["审批", "批准", "审核", "核准", "同意", "签字", "决定", "确定", "审定", "核定", "批复"],
    "执行": ["施工", "维修", "安装", "采购", "验收", "勘查", "检查", "实施", "执行", "操作", "处置", "派工", "领用", "发放"],
    "核验": ["验收", "核验", "核查", "复核", "检查", "审计", "监督", "盘点", "审核", "核实", "确认"],
}

# 排查分类 → 最小节点清单对照（用于交叉校验）
INSPECTION_CATEGORY_NODE_MAP = {
    "工程建设与抢修维修（含二供）": {
        "项目立项": ["项目立项/需求提出"],
        "招投标": ["招投标/采购方式确定"],
        "工程变更": ["现场施工/工程变更"],
        "验收结算": ["验收", "结算"],
        "抢修维修派工": ["施工派工"],
    },
    "给水工程验收": {
        "图纸设计审查": ["图纸设计审查"],
        "材料进场": ["材料进场核验"],
        "预验收": ["预验收"],
        "竣工验收": ["竣工验收"],
    },
    "用户报装与接水业务": {
        "报装审批": ["报装申请受理", "审批"],
        "现场勘查": ["现场勘查"],
        "预算编制": ["预算编制"],
        "施工安排": ["施工安排/派工"],
    },
    "物资及服务采购": {
        "招投标": ["采购方式确定", "供应商选择/招标"],
        "供应商选择": ["供应商选择/招标"],
        "询价比价": ["采购方式确定"],
        "验收入库": ["验收入库"],
        "物料申领及退库管理": ["物料申领", "退库管理"],
    },
    "综合管理": {
        "财务": ["财务审批"],
        "人事": ["人事任免/考核"],
        "饭堂": [],
        "公务车辆": ["公务车辆使用管理"],
        "资产": ["资产管理"],
        "废旧物资处置": ["废旧物资处置"],
        "安全生产": ["安全生产检查"],
    },
}

# 表单举例文本中的模板占位符（用于检测单位是否直接抄模板未做替换）
TEMPLATE_PLACEHOLDERS = [
    "【A岗位】", "【B岗位】", "【C岗位】", "【D岗位】",
    "【XX岗位】", "【xx岗位】", "【xX岗位】", "【Xx岗位】",
    "XX个工作日内", "XX系统", "XX标准", "xX个工作日内",
    "XX专责", "XX经理", "XX员",
]

# 表单举例文本关键词片段（用于相似度检测）
TEMPLATE_EXAMPLE_SNIPPETS = {
    "process": [
        "窗口或线上平台接收用户申请",
        "组成勘查小组",
        "现场踏勘",
        "编制预算并上传",
        "安排施工",
        "现场监督隐蔽工程",
        "上传施工日志",
        "发起联合验收",
        "归档至档案系统",
    ],
    "risk": [
        "利用职务便利，在勘查现场索取烟酒、礼品",
        "受利益驱动，违规调整方案",
        "人为调整物料用量",
        "化整为零",
        "规避招标",
        "利益输送风险",
        "利用派工权限，指定特定施工队伍",
        "索要回扣",
        "材料验收或隐蔽工程验收中放水",
        "以次充好",
    ],
    "measures": [
        "系统自动标红提醒",
        "每月对项目进行现场复核，核对现场情况与系统的一致性",
        "两人勘查、两人验收",
        "实施岗位轮换，严禁长期负责同一片区",
        "定期开展对账，严查结余物资是否按规定退回仓库",
        "以领代耗",
    ],
}


# ============================================================================
# 最小节点清单（业务标尺）
# ============================================================================

MIN_NODE_CHECKLIST = {
    "工程建设与抢修维修（含二供）": [
        {"name": "项目立项/需求提出", "critical": True,
         "desc": "工程需求的来源、立项依据和审批流程"},
        {"name": "招投标/采购方式确定", "critical": True,
         "desc": "确定施工单位的程序合规性"},
        {"name": "合同签订", "critical": True,
         "desc": "合同条款、金额、资质审核"},
        {"name": "施工派工", "critical": False,
         "desc": "派工依据、派工记录、派工权限"},
        {"name": "材料领用", "critical": False,
         "desc": "材料申领审批、实际用量与申领量核对"},
        {"name": "现场施工/工程变更", "critical": True,
         "desc": "施工过程管理、变更签证的真实性和必要性"},
        {"name": "验收", "critical": True,
         "desc": "验收标准、验收人员、验收记录"},
        {"name": "结算", "critical": True,
         "desc": "工程量的核实、计价依据、付款审批"},
    ],
    "给水工程验收": [
        {"name": "图纸设计审查", "critical": True,
         "desc": "设计图纸的合规性、审查记录"},
        {"name": "材料进场核验", "critical": True,
         "desc": "材料规格、数量、质量证明文件核验"},
        {"name": "隐蔽工程验收", "critical": True,
         "desc": "隐蔽工程在覆盖前的验收记录和影像资料"},
        {"name": "预验收", "critical": False,
         "desc": "预验收的组织、问题记录和整改跟踪"},
        {"name": "竣工验收", "critical": True,
         "desc": "竣工验收的标准执行、签字程序"},
        {"name": "结算审核", "critical": True,
         "desc": "结算资料的完整性、审核流程"},
    ],
    "用户报装与接水业务": [
        {"name": "报装申请受理", "critical": True,
         "desc": "报装材料的接收、登记和分发"},
        {"name": "现场勘查", "critical": True,
         "desc": "勘查人员的派出、勘查记录的真实性"},
        {"name": "预算编制", "critical": False,
         "desc": "预算编制的依据和审核"},
        {"name": "审批", "critical": True,
         "desc": "审批权限、审批时限、审批依据"},
        {"name": "施工安排/派工", "critical": False,
         "desc": "施工队伍的选定、派工机制"},
        {"name": "验收通水", "critical": True,
         "desc": "验收标准、通水条件确认"},
        {"name": "归档", "critical": False,
         "desc": "档案资料的完整性和可追溯性"},
    ],
    "物资及服务采购": [
        {"name": "需求申请", "critical": True,
         "desc": "采购需求的真实性和合理性"},
        {"name": "采购方式确定", "critical": True,
         "desc": "采购方式选择的合规性（公开招标/邀请招标/询价/单一来源）"},
        {"name": "供应商选择/招标", "critical": True,
         "desc": "招标文件、评标过程、供应商资质审核"},
        {"name": "合同签订", "critical": True,
         "desc": "合同条款审核、金额与中标一致性"},
        {"name": "验收入库", "critical": True,
         "desc": "验收程序、数量规格核对、入库登记"},
        {"name": "物料申领", "critical": False,
         "desc": "申领审批、实际领用与申领一致性"},
        {"name": "退库管理", "critical": False,
         "desc": "退库物资的登记、处置和去向"},
    ],
    "综合管理": [
        {"name": "财务审批", "critical": True,
         "desc": "财务支出的审批权限、审批依据"},
        {"name": "人事任免/考核", "critical": True,
         "desc": "人事决策的程序公正性"},
        {"name": "公务车辆使用管理", "critical": False,
         "desc": "车辆使用登记、油耗管理"},
        {"name": "资产管理", "critical": True,
         "desc": "资产登记、盘点、调拨管理"},
        {"name": "废旧物资处置", "critical": True,
         "desc": "废旧物资的鉴定、处置方式和收入管理"},
        {"name": "安全生产检查", "critical": True,
         "desc": "安全检查的执行、隐患整改跟踪"},
    ],
}


# ============================================================================
# 现场核查建议模板
# ============================================================================

VERIFICATION_TEMPLATES = {
    "工程建设与抢修维修（含二供）": {
        "documents": [
            "项目立项申请及审批文件",
            "招标/采购文件（招标公告、投标文件、评标记录、中标通知书）",
            "施工合同及补充协议",
            "派工单、派工台账",
            "材料领用申请单、出库单",
            "工程变更签证单（附施工日志、监理日志对照）",
            "竣工验收报告",
            "结算书及付款凭证",
        ],
        "interview_directions": [
            "工程变更：变更的真实原因是什么？是否存在先施工后补签证？变更金额占比是否异常？",
            "招投标：是否存在为特定供应商量身定做招标条件？投标人之间是否存在关联关系？",
            "验收：验收人员是否实际到场？核对工程量是否走过场？是否存在好处费或吃请？",
            "派工：派工是否有轮流或随机机制？是否存在长期只用某支队伍的情况？",
        ],
        "questions": [
            "变更签证单上的工程量与实际施工是否一致？请提供对应的施工日志和监理日志。",
            "同一施工队近三年承接了多少项目？是否存在集中度高的情况？原因是什么？",
            "验收环节是否核对了全部必检项？验收记录是否当场签字确认？",
            "抢修项目的派工是如何决定的？有没有派工记录？谁有权派工？",
            "材料实际用量与申领量之间的差异如何核对？余料如何处理？",
        ],
    },
    "给水工程验收": {
        "documents": [
            "图纸设计审查记录",
            "材料进场验收记录、材料合格证及检测报告",
            "隐蔽工程验收记录（含影像资料）",
            "预验收报告及整改通知",
            "竣工验收报告",
            "结算审核文件",
        ],
        "interview_directions": [
            "验收标准：是否存在因人而异的验收标准？是否存在领导打招呼降低标准？",
            "材料核验：进场材料是否批批核验？不合格材料如何处理？",
            "隐蔽工程：覆盖前是否实际进行了验收？影像资料是否完整？",
            "结算：结算工程量与实际工程量是否一致？",
        ],
        "questions": [
            "验收组成员是如何确定的？是否存在固定搭配的验收组？",
            "发现的问题是否如实记录并跟踪整改？请提供近一年的整改记录。",
            "隐蔽工程验收的频率是否满足规范要求？是否有施工方通知了但没来的情况？",
            "结算审核时发现过工程量不符的情况吗？如何处理？",
        ],
    },
    "用户报装与接水业务": {
        "documents": [
            "报装申请登记台账",
            "现场勘查记录",
            "预算编制书及审批记录",
            "施工派工单",
            "验收通水记录",
            "收费凭证",
        ],
        "interview_directions": [
            "审批时限：是否存在某些工单审批特别快（插队）或特别慢（卡要）？",
            "队伍指定：施工队伍是如何选定的？用户是否有权选择？",
            "违规收费：是否存在预算外收费？收费标准是否公开透明？",
            "吃拿卡要：勘查、施工、验收人员是否接受过用户的吃请或好处？",
        ],
        "questions": [
            "报装审批的平均时限是多少？近一年是否有超时未批的工单？",
            "施工队伍的选择机制是什么？是否存在领导或内部人员指定队伍的情况？",
            "除公示的收费标准外，是否存在其他费用？如果有，是什么名目？",
            "是否做过客户满意度回访或匿名调查？回访中是否收到过廉洁方面的投诉？",
        ],
    },
    "物资及服务采购": {
        "documents": [
            "采购申请单及审批记录",
            "采购方式确定依据",
            "招标文件、投标文件、评标记录",
            "供应商资质文件",
            "采购合同",
            "验收单、入库单",
            "领料单、退库单",
        ],
        "interview_directions": [
            "围标串标：投标人之间是否存在关联？投标报价是否存在规律性差异？",
            "化整为零：是否存在将大额采购拆分为多笔小额以规避招标的情况？",
            "利益输送：供应商与内部人员是否存在亲属或利益关系？",
            "账实不符：入库物资与实际采购是否一致？领用与库存是否对得上？",
        ],
        "questions": [
            "同一供应商近三年的供货金额和频次是多少？集中度是否异常？",
            "是否存在同一品类在相近时间内多笔小额采购的情况？原因是什么？",
            "验收时是否逐项核对规格、数量、质量？验收人与采购人是否分离？",
            "库存盘点发现了哪些问题？盘盈盘亏是如何处理的？",
        ],
    },
    "综合管理": {
        "documents": [
            "财务凭证及审批记录",
            "人事任免和考核记录",
            "饭堂采购及收支台账",
            "公务车辆使用登记、油耗记录",
            "资产台账及盘点记录",
            "废旧物资处置审批及收入记录",
            "安全生产检查记录及整改台账",
        ],
        "interview_directions": [
            '制度执行：内控制度是否被绕过？是否存在「特批」和「口头同意」代替书面审批？',
            "利益输送：饭堂采购、车辆维修等是否存在固定供应商且价格偏高？",
            "失职失责：安全生产检查是否流于形式？发现的问题是否真正整改？",
        ],
        "questions": [
            "公务车辆的使用审批和登记是否规范？是否存在私用或非工作时间使用？",
            "废旧物资的处置是否按规定程序进行？处置收入是否全部入账？",
            "人事考核和任免是否有明确的书面标准和程序？是否存在因人设岗或违规提拔？",
            "饭堂物资采购是否有比价机制？是否存在固定供应商长期未更换？",
        ],
    },
}


# ============================================================================
# 流程完整性分析规则
# ============================================================================

# 合理的节点类型顺序（发起→审批→执行→核验）
AUTHORITY_ORDER = ["发起", "审批", "执行", "核验"]

# 模糊化节点名关键词（疑似隐瞒）
VAGUE_NODE_KEYWORDS = [
    "处理", "办理", "操作", "管理", "跟进", "落实", "完善", "安排",
    "相关工作", "其他事项", "日常", "常规",
]

# 不应合并的关键节点对（两个词同时出现在同一节点名中）
CRITICAL_MERGE_PAIRS = [
    ("采购", "验收"), ("审批", "结算"), ("执行", "核验"),
    ("申请", "审批"), ("招标", "评标"), ("施工", "验收"),
    ("验收", "结算"), ("入库", "领用"),
]

# 各排查领域的最低合理节点数
MIN_PROCESS_NODES = {
    "工程建设与抢修维修（含二供）": 4,
    "给水工程验收": 3,
    "用户报装与接水业务": 3,
    "物资及服务采购": 3,
    "综合管理": 2,
    "其他": 2,
}

# 各排查领域应有的闭环起点/终点关键词
CLOSURE_KEYWORDS = {
    "工程建设与抢修维修（含二供）": {
        "start": ["申请", "报修", "立项", "接单", "受理", "需求"],
        "end": ["验收", "结算", "归档", "通水", "完成"],
    },
    "给水工程验收": {
        "start": ["申请", "报验", "受理", "提交"],
        "end": ["验收", "结算", "归档", "通过"],
    },
    "用户报装与接水业务": {
        "start": ["申请", "报装", "受理", "登记"],
        "end": ["验收", "通水", "归档", "完成"],
    },
    "物资及服务采购": {
        "start": ["申请", "需求", "立项"],
        "end": ["验收", "入库", "结算", "归档"],
    },
    "综合管理": {
        "start": ["申请", "发起", "提出", "安排"],
        "end": ["审批", "归档", "完成", "处置"],
    },
}

# 关键控制节点名（这些节点必须标注岗位）
CRITICAL_POSITION_NODES = ["审批", "验收", "结算", "招标", "评标", "核验", "检查"]


# ============================================================================
# 数据结构
# ============================================================================

@dataclass
class ProcessNode:
    """单个流程节点"""
    name: str           # 节点名称（清洗后）
    raw: str            # 原始文本
    position: str       # 从节点描述中提取的岗位
    sequence: int       # 在流程中的顺序


@dataclass
class Submission:
    """一条单位自查记录"""
    unit_name: str
    unit_role: str      # "统筹" / "实施" / "双重"
    inspection_domain: str          # 排查领域（字段3，单选题）
    inspection_categories: list     # 排查分类（字段4，多选题）
    business_name_raw: str          # 业务名称（字段1，填空题）
    business_name_canonical: str    # 规范化的业务名称（经交叉校验后填入）
    domain_confidence: str          # 领域匹配置信度: "高"/"待确认"/"推断"
    raw_process: str
    nodes: list
    positions_raw: str              # "涉及关键岗位"字段原文
    positions: list                 # 解析后的岗位列表
    risk_description: str           # "廉洁风险点自查描述"字段原文
    prevention_measures: str        # "防控措施清单"字段原文
    remarks: str                    # "备注"字段原文

    def __hash__(self):
        return hash((self.unit_name, self.unit_role, self.business_name_raw))


@dataclass
class NodeDiff:
    """单个节点的应然-实然对比差异"""
    seq: int
    ought_node: Optional[str]
    is_node: Optional[str]
    ought_position: Optional[str]
    is_position: Optional[str]
    diff_type: str      # 消失/新增/偏移/一致/名称差异
    description: str


@dataclass
class BusinessRisk:
    """业务线风险评估结果"""
    business_name: str
    node_completeness: float
    position_clarity: float
    power_concentration: float
    ought_is_match: float
    self_awareness: float
    total_score: float
    level: str
    level_desc: str


# ============================================================================
# 文本解析
# ============================================================================

def parse_process_text(text: str) -> list:
    """将自由文本流程描述解析为 ProcessNode 列表"""
    if not text or not text.strip():
        return []

    text = text.strip()
    nodes_raw = _split_nodes(text)
    result = []
    for i, raw in enumerate(nodes_raw):
        raw = raw.strip().strip("。，,，、.").strip()
        if not raw or raw in ("结束", "完成", "办结", "完毕", "end", "END"):
            continue
        name, pos = _extract_position(raw)
        result.append(ProcessNode(
            name=name.strip(),
            raw=raw,
            position=pos.strip(),
            sequence=i,
        ))
    return result


def _split_nodes(text: str) -> list[str]:
    """尝试多种分隔符拆分流程文本"""
    # 1. 多连字符（中文/英文）
    if re.search(r'[-—]{2,}', text):
        return re.split(r'[-—]{2,}', text)
    # 2. 箭头
    if re.search(r'[→>]', text):
        return re.split(r'\s*[→>]\s*', text)
    # 3. 数字编号（1. 2. 3. 或 1、2、3、）
    if re.search(r'(?:^|\s)\d+[.、．)]', text):
        parts = re.split(r'(?:\d+[.、．)]\s*)', text)
        return [p for p in parts if p.strip()]
    # 4. 中文分号
    if '；' in text:
        return text.split('；')
    # 5. 换行
    if '\n' in text:
        return text.split('\n')
    # 6. 逗号/顿号（最后兜底）
    return re.split(r'[,，、]', text)


def _extract_position(node_text: str) -> tuple[str, str]:
    """从节点文本中提取岗位信息。返回 (节点名, 岗位)。"""
    # 匹配括号内容：中文括号 or 英文括号
    m = re.search(r'[（(]([^）)]+)[）)]', node_text)
    if m:
        position = m.group(1)
        name = node_text[:m.start()] + node_text[m.end():]
        # 如果提取的内容看起来像岗位
        if any(kw in position for kw in POSITION_KEYWORDS) or len(position) <= 15:
            return name.strip(), position
    return node_text, ""


def parse_positions(raw: str) -> list[str]:
    """解析'涉及岗位'字段"""
    if not raw or not raw.strip():
        return []
    # 尝试多种分隔符
    for sep in ['；', ';', '、', ',', '，', '\n', ' ']:
        if sep in raw:
            return [p.strip() for p in raw.split(sep) if p.strip()]
    return [raw.strip()]


# ============================================================================
# Excel 读取
# ============================================================================

def _fuzzy_find_column(col_names: list[str], target: str) -> Optional[str]:
    """在列名列表中模糊查找目标列名"""
    target_lower = target.lower().strip()
    for name in col_names:
        name_lower = name.lower().strip()
        if target_lower == name_lower:
            return name
        if target_lower in name_lower or name_lower in target_lower:
            # 但排除反向匹配到其他明确列
            if len(name_lower) >= 2:
                return name
    # 尝试 difflib 匹配
    best_ratio = 0
    best_match = None
    for name in col_names:
        ratio = difflib.SequenceMatcher(None, target.lower().strip(), name.lower().strip()).ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_match = name
    if best_ratio > 0.6:
        return best_match
    return None


def parse_excel(filepath: str) -> list[Submission]:
    """解析各单位提交的 Excel 表格（适配宣贯会后确定的 9 字段表单）"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Excel 文件为空")

    # 定位表头行（第一行）
    headers = [str(h).strip() if h else "" for h in rows[0]]
    print(f"[解析] 检测到表头: {headers}")

    # 模糊匹配列名 — 映射到 9 个表单字段
    col_map = {}
    required_targets = {
        "unit_name":       "单位名称",
        "unit_role":       "单位职能",
        "business_name":   "业务名称",
        "inspection_domain": "排查领域",
        "inspection_cat":  "排查分类",
        "process":         "业务流程",
        "positions":       "涉及关键岗位",
        "risk_desc":       "廉洁风险点自查描述",
        "measures":        "防控措施清单",
        "remarks":         "备注",
    }
    for key, target in required_targets.items():
        if key in col_map:
            continue
        col = _fuzzy_find_column(headers, target)
        if col is not None:
            col_map[key] = headers.index(col)
        elif key in ("unit_name", "unit_role", "inspection_domain", "process"):
            # 必需列缺失
            alternatives = {
                "unit_name": "单位名称",
                "unit_role": "单位职能",
                "inspection_domain": "排查领域",
                "process": "业务流程",
            }
            raise ValueError(
                f"找不到必需列: '{target}'。请检查表头。\n"
                f"检测到表头: {headers}\n"
                f"提示：确保 Excel 表头包含{', '.join(alternatives.values())}等列。"
            )

    print(f"[解析] 列映射: {col_map}")

    submissions = []
    missing_info = []
    for i, row in enumerate(rows[1:], start=2):
        vals = [str(v).strip() if v is not None else "" for v in row]

        def _get(key, default=""):
            idx = col_map.get(key, -1)
            return vals[idx] if idx >= 0 and idx < len(vals) else default

        unit_name = _get("unit_name")
        unit_role = _get("unit_role")
        business_name = _get("business_name")
        inspection_domain = _get("inspection_domain")
        inspection_cat_raw = _get("inspection_cat")
        process_text = _get("process")
        positions_raw = _get("positions")
        risk_desc = _get("risk_desc")
        measures = _get("measures")
        remarks = _get("remarks")

        if not unit_name:
            continue

        # 解析排查分类（多选，可能用换行/逗号/分号分隔）
        inspection_categories = _parse_multi_select(inspection_cat_raw)

        # 标准化单位职能
        role_normalized = _normalize_role(unit_role)

        # 分离节点
        nodes = parse_process_text(process_text)
        positions = parse_positions(positions_raw)

        # ---- 业务名称双源交叉校验 ----
        canonical_from_domain = _canonical_domain(inspection_domain)
        inferred_from_process = _infer_business_from_process(process_text) if process_text else ""

        if canonical_from_domain and canonical_from_domain != "其他":
            if inferred_from_process and inferred_from_process != canonical_from_domain:
                # 排查领域与流程推断矛盾 → 标记待确认
                confidence = "待确认"
                canonical = canonical_from_domain  # 以排查领域为准，但标记
                print(f"  [!] 第{i}行 {unit_name}: 排查领域='{inspection_domain}'但流程似'{inferred_from_process}'，以排查领域为准，建议人工核实")
            else:
                confidence = "高"
                canonical = canonical_from_domain
        else:
            # 排查领域为"其他"或为空 → 回退到流程推断
            if inferred_from_process:
                confidence = "推断"
                canonical = inferred_from_process
                print(f"  [i] 第{i}行 {unit_name}: 排查领域为'其他/空'，从流程推断为'{inferred_from_process}'")
            else:
                confidence = "待确认"
                canonical = ""

        s = Submission(
            unit_name=unit_name,
            unit_role=role_normalized,
            inspection_domain=inspection_domain,
            inspection_categories=inspection_categories,
            business_name_raw=business_name if business_name else (inspection_domain or inferred_from_process),
            business_name_canonical=canonical,
            domain_confidence=confidence,
            raw_process=process_text,
            nodes=nodes,
            positions_raw=positions_raw,
            positions=positions,
            risk_description=risk_desc,
            prevention_measures=measures,
            remarks=remarks,
        )

        if not s.nodes:
            missing_info.append(f"  第{i}行: {unit_name} - 流程描述无法解析出节点")
        submissions.append(s)

    if missing_info:
        print(f"\n[警告] 以下记录流程描述无法解析出节点：")
        for m in missing_info:
            print(m)

    print(f"\n[解析] 共读取 {len(submissions)} 条记录")
    role_counts = defaultdict(int)
    for s in submissions:
        role_counts[s.unit_role] += 1
    print(f"[解析] 角色分布: {dict(role_counts)}")

    # 统计领域置信度
    conf_counts = defaultdict(int)
    for s in submissions:
        conf_counts[s.domain_confidence] += 1
    if conf_counts:
        print(f"[解析] 领域置信度: {dict(conf_counts)}")

    return submissions


def _looks_like_process(text: str) -> bool:
    """判断文本是否像流程描述（而非业务名称）"""
    if not text:
        return False
    indicators = ['---', '——', '→', '->', '；', '1.', '2.', '1、', '\n']
    score = sum(1 for ind in indicators if ind in text)
    return score >= 1 or len(text) > 50


def _normalize_role(raw: str) -> str:
    """标准化单位职能（适配宣贯会后的选项：业务管理统筹/业务实施/双重职责）"""
    raw = raw.strip()
    # "双重职责"优先于"统筹"和"实施"
    if ('统筹' in raw and '实施' in raw) or '双重' in raw:
        return '双重'
    if '实施' in raw:
        return '实施'
    if '统筹' in raw:
        return '统筹'
    return raw


def _parse_multi_select(raw: str) -> list[str]:
    """解析多选字段（排查分类），支持换行/逗号/分号/顿号分隔"""
    if not raw or not raw.strip():
        return []
    # 按换行优先拆分
    if '\n' in raw:
        parts = [p.strip() for p in raw.split('\n') if p.strip()]
    elif '；' in raw or ';' in raw:
        parts = [p.strip() for p in re.split(r'[；;]', raw) if p.strip()]
    elif '，' in raw or ',' in raw:
        parts = [p.strip() for p in re.split(r'[，,]', raw) if p.strip()]
    elif '、' in raw:
        parts = [p.strip() for p in raw.split('、') if p.strip()]
    else:
        parts = [raw.strip()]
    # 清洗：去除编号前缀如"1." "1、"等
    cleaned = []
    for p in parts:
        p = re.sub(r'^\d+[.、．)]\s*', '', p).strip()
        if p:
            cleaned.append(p)
    return cleaned


def _canonical_domain(raw: str) -> str:
    """将排查领域选项映射为规范业务线名称"""
    if not raw:
        return ""
    raw = raw.strip()
    valid_domains = [
        "工程建设与抢修维修（含二供）",
        "给水工程验收",
        "用户报装与接水业务",
        "物资及服务采购",
        "综合管理",
    ]
    for d in valid_domains:
        if raw == d:
            return d
    # 模糊匹配
    raw_lower = raw.lower().replace("（", "(").replace("）", ")")
    for d in valid_domains:
        d_lower = d.lower().replace("（", "(").replace("）", ")")
        if raw_lower == d_lower:
            return d
        if raw_lower in d_lower or d_lower in raw_lower:
            return d
    if "其他" in raw:
        return "其他"
    return raw  # 保留原文，后续人工确认


def _detect_template_copy(text: str, field_type: str) -> dict:
    """
    检测文本是否疑似抄模板/举例。
    返回 {"suspicious": bool, "reasons": [str], "score_penalty": float}
    """
    if not text or not text.strip():
        return {"suspicious": False, "reasons": [], "score_penalty": 0.0}

    reasons = []
    penalty = 0.0

    # 1. 检测未替换的占位符
    placeholder_hits = [ph for ph in TEMPLATE_PLACEHOLDERS if ph in text]
    if placeholder_hits:
        reasons.append(f"含模板占位符未替换: {', '.join(placeholder_hits[:5])}")
        penalty += min(2.0, len(placeholder_hits) * 0.5)

    # 2. 与举例文本做相似度比对
    snippets = TEMPLATE_EXAMPLE_SNIPPETS.get(field_type, [])
    snippet_hits = 0
    for snippet in snippets:
        # 用最长公共子串比例来判断
        sm = difflib.SequenceMatcher(None, text, snippet)
        if sm.ratio() > 0.6:
            snippet_hits += 1
    if snippet_hits >= 2:
        reasons.append(f"与举例文本高度相似（{snippet_hits}处匹配）")
        penalty += min(1.5, snippet_hits * 0.4)

    # 3. 检测文本是否过短（敷衍）
    if field_type in ("risk", "measures"):
        text_clean = text.strip().replace(" ", "").replace("\n", "")
        if len(text_clean) < 10:
            reasons.append("文本过短，疑似敷衍填写")
            penalty += 1.0

    return {
        "suspicious": penalty >= 0.5,
        "reasons": reasons,
        "score_penalty": min(2.5, penalty),
    }


# ============================================================================
# 业务名称匹配
# ============================================================================

# 业务线关键词（用于从流程描述推断所属业务线）
BUSINESS_KEYWORDS = {
    "工程建设与抢修维修（含二供）": [
        "接单", "派工", "抢修", "维修", "工程施工", "二供", "爆漏", "管网",
        "施工员", "维修工", "调度", "领料", "仓管", "现场施工", "工程变更",
    ],
    "给水工程验收": [
        "图纸", "设计审查", "隐蔽工程", "竣工", "预验收", "材料进场",
        "结算审核", "质检", "监理",
    ],
    "用户报装与接水业务": [
        "报装", "接水", "通水", "新装", "用水", "勘查", "预算编制",
        "客服", "报装申请",
    ],
    "物资及服务采购": [
        "采购", "招标", "供应商", "入库", "领料", "退库", "评标",
        "询价", "比价", "采购员", "仓管", "合同签订",
    ],
    "综合管理": [
        "财务", "人事", "车辆", "资产", "废旧", "安全生产", "饭堂",
        "考核", "任免", "处分", "出纳",
    ],
}


def _infer_business_from_process(process_text: str) -> str:
    """从流程描述文本中推断所属业务线（基于关键词匹配）"""
    scores = {}
    for biz, keywords in BUSINESS_KEYWORDS.items():
        score = sum(1 for kw in keywords if kw in process_text)
        if score > 0:
            scores[biz] = score
    if scores:
        return max(scores, key=scores.get)
    return ""


def load_aliases() -> dict:
    """加载业务名称别名映射"""
    if ALIASES_PATH.exists():
        with open(ALIASES_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return {k: v for k, v in data.items() if not k.startswith('_')}
    return {}


def build_alias_lookup(aliases: dict) -> dict[str, str]:
    """构建别名到规范名的反向索引"""
    lookup = {}
    for canonical, alias_list in aliases.items():
        lookup[canonical] = canonical
        for alias in alias_list:
            lookup[alias] = canonical
    return lookup


def match_business_name(name: str, alias_lookup: dict[str, str]) -> tuple[Optional[str], str]:
    """
    将业务名称匹配到规范名称。
    返回 (规范名称, 匹配方式)，匹配方式: "精确"/"别名"/"模糊"/"未匹配"
    """
    if not name:
        return None, "未匹配"

    name_clean = name.strip()

    # 1. 精确匹配（规范名）
    if name_clean in alias_lookup:
        canonical = alias_lookup.get(name_clean, name_clean)
        # 如果 alias_lookup value 也是规范名，用 key
        for k, v in alias_lookup.items():
            if v == canonical and k == canonical:
                return canonical, "精确"

    # 先查自身是否就是规范名
    canonical_names = set(alias_lookup.values())
    if name_clean in canonical_names:
        return name_clean, "精确"

    # 2. 别名匹配
    if name_clean in alias_lookup:
        return alias_lookup[name_clean], "别名"

    # 3. 模糊匹配（用 difflib）
    all_names = list(alias_lookup.keys())
    best_ratio = 0
    best_match = None
    for candidate in all_names:
        ratio = difflib.SequenceMatcher(None, name_clean, candidate).ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_match = candidate
    if best_ratio >= 0.6 and best_match:
        canonical = alias_lookup.get(best_match, best_match)
        return canonical, "模糊"

    return None, "未匹配"


def pair_ought_and_is(submissions: list[Submission], alias_lookup: dict[str, str]) -> dict:
    """
    将统筹流程与实施流程配对。
    返回: {
        canonical_business_name: {
            "ought": [统筹/双重单位的Submission],
            "is": [实施/双重单位的Submission],
            "unmatched_ought": [...],
        }
    }
    """
    pairs = defaultdict(lambda: {"ought": [], "is": [], "unmatched": []})

    for s in submissions:
        canonical, match_type = match_business_name(s.business_name_raw, alias_lookup)
        s.business_name_canonical = canonical or s.business_name_raw
        if match_type == "未匹配":
            pairs[s.business_name_raw]["unmatched"].append(s)
        elif s.unit_role in ("统筹", "双重"):
            pairs[canonical]["ought"].append(s)
        elif s.unit_role == "实施":
            pairs[canonical]["is"].append(s)

    return pairs


# ============================================================================
# 孤儿业务检测
# ============================================================================

def detect_orphan_businesses(submissions: list[Submission]) -> dict:
    """
    检测同一排查领域内统筹与实施不匹配的情况。
    返回: {
        "orphan_ought": [(domain, unit_name, business_name_raw), ...],  # 有统筹无实施
        "orphan_is": [(domain, unit_name, business_name_raw), ...],      # 有实施无统筹
    }
    """
    # 按排查领域分组
    by_domain = defaultdict(lambda: {"ought": [], "is": []})
    for s in submissions:
        domain = s.inspection_domain or s.business_name_canonical or "未知领域"
        if s.unit_role in ("统筹", "双重"):
            by_domain[domain]["ought"].append(s)
        if s.unit_role in ("实施", "双重"):
            by_domain[domain]["is"].append(s)

    orphan_ought = []
    orphan_is = []

    for domain, data in by_domain.items():
        ought_units = {s.unit_name for s in data["ought"]}
        is_units = {s.unit_name for s in data["is"]}

        if data["ought"] and not data["is"]:
            for s in data["ought"]:
                orphan_ought.append((domain, s.unit_name, s.business_name_raw))
        if data["is"] and not data["ought"]:
            for s in data["is"]:
                orphan_is.append((domain, s.unit_name, s.business_name_raw))

        # 即使同一领域有统筹和实施，也检查业务名称层面的匹配
        # （同一领域可能有多条不同的业务线）
        if data["ought"] and data["is"]:
            ought_biz_names = {s.business_name_raw for s in data["ought"]}
            is_biz_names = {s.business_name_raw for s in data["is"]}
            # 业务名完全无交集 → 可能一个说的是A业务，另一个说的是B业务
            # 这种情况较弱，暂不标记为孤儿，在摘要中提示即可

    return {
        "orphan_ought": orphan_ought,
        "orphan_is": orphan_is,
    }


# ============================================================================
# 排查分类 × 最小节点清单交叉校验
# ============================================================================

def validate_inspection_categories(submission: Submission) -> dict:
    """
    校验单位勾选的排查分类是否在流程描述中有对应节点。
    返回: {
        "covered": [勾选且在流程中找到的排查分类],
        "missing": [(排查分类, 预期节点名), ...],  # 勾选了但流程中找不到
        "extra": [节点名, ...],  # 流程中有但未勾选的节点（可能是额外覆盖）
    }
    """
    domain = submission.inspection_domain or submission.business_name_canonical
    if not domain or domain == "其他":
        return {"covered": [], "missing": [], "extra": []}

    cat_map = INSPECTION_CATEGORY_NODE_MAP.get(domain, {})
    if not cat_map:
        return {"covered": [], "missing": [], "extra": []}

    process_node_names = [n.name for n in submission.nodes]
    checked_cats = submission.inspection_categories

    covered = []
    missing = []

    for cat in checked_cats:
        expected_nodes = cat_map.get(cat, [])
        if not expected_nodes:
            # 排查分类无对应清单节点（如"饭堂"），跳过校验
            covered.append(cat)
            continue
        # 检查流程中是否至少有一个节点与预期节点匹配
        found = False
        for en in expected_nodes:
            for pn in process_node_names:
                if difflib.SequenceMatcher(None, en, pn).ratio() >= 0.5:
                    found = True
                    break
            if found:
                break
        if found:
            covered.append(cat)
        else:
            missing.append((cat, expected_nodes))

    # 检测流程中多出但未勾选的节点（弱信号，仅供参考）
    all_expected = set()
    for nodes in cat_map.values():
        for n in nodes:
            all_expected.add(n)
    extra = []
    for pn in process_node_names:
        best = max(
            (difflib.SequenceMatcher(None, en, pn).ratio() for en in all_expected),
            default=0,
        )
        if best < 0.5 and pn not in extra:
            extra.append(pn)

    return {
        "covered": covered,
        "missing": missing,
        "extra": extra,
    }


# ============================================================================
# 流程完整性分析（单流程独立审查）
# ============================================================================

def analyze_process_integrity(s: Submission) -> dict:
    """
    独立审查单条流程的完整性和合理性，不依赖其他单位对比。
    检查三个维度：逻辑合理性、业务闭环、隐瞒信号。
    返回评分(0-100)和问题列表。
    """
    issues = []
    base_score = 100.0

    if not s.nodes:
        return {
            "unit": s.unit_name,
            "role": s.unit_role,
            "domain": s.inspection_domain or s.business_name_canonical,
            "business": s.business_name_raw,
            "node_count": 0,
            "score": 0,
            "grade": "严重不完整",
            "issues": [{"维度": "业务闭环", "问题": "无法解析流程节点，建议人工核查"}],
        }

    # ---- 1. 逻辑合理性 ----
    logic_issues = _check_logic_rationality(s.nodes)
    issues.extend([{"维度": "逻辑合理", "问题": li} for li in logic_issues])
    base_score -= len(logic_issues) * 8

    # ---- 2. 业务闭环 ----
    domain = s.inspection_domain or s.business_name_canonical or "其他"
    closure_issues = _check_closure(s.nodes, domain)
    issues.extend([{"维度": "业务闭环", "问题": ci} for ci in closure_issues])
    base_score -= len(closure_issues) * 10

    # ---- 3. 隐瞒信号 ----
    conceal_issues = _check_concealment(s)
    issues.extend([{"维度": "隐瞒信号", "问题": ci} for ci in conceal_issues])
    base_score -= len(conceal_issues) * 12

    score = max(0.0, min(100.0, base_score))

    return {
        "unit": s.unit_name,
        "role": s.unit_role,
        "domain": domain,
        "business": s.business_name_raw,
        "node_count": len(s.nodes),
        "score": round(score, 1),
        "grade": _integrity_grade(score),
        "issues": issues,
    }


def _check_logic_rationality(nodes: list) -> list[str]:
    """检查流程节点的逻辑合理性"""
    issues = []

    # 1.1 检测每个节点的权力类型
    node_auth_types = []
    for n in nodes:
        auths = set()
        for auth_type, patterns in AUTHORITY_NODE_PATTERNS.items():
            if any(p in n.name for p in patterns):
                auths.add(auth_type)
        node_auth_types.append(auths if auths else {"未知"})

    # 1.2 节点顺序检查：审批不应在全部执行之后、验收不应在第一个
    last_auth_index = {}
    for i, auths in enumerate(node_auth_types):
        for at in auths:
            last_auth_index[at] = i

    # 如果"核验"类节点出现在"执行"类节点之前（位置靠前很多）
    exec_indices = [i for i, auths in enumerate(node_auth_types) if "执行" in auths]
    verify_indices = [i for i, auths in enumerate(node_auth_types) if "核验" in auths]
    if exec_indices and verify_indices:
        last_exec = max(exec_indices)
        first_verify = min(verify_indices)
        # 如果所有核验都在执行之前，可疑
        if max(verify_indices) < min(exec_indices):
            issues.append(f"所有核验节点({nodes[first_verify].name}等)均在执行节点({nodes[min(exec_indices)].name}等)之前，顺序可能不合理")

    # 审批节点如果全在执行之后 -> 可疑
    approve_indices = [i for i, auths in enumerate(node_auth_types) if "审批" in auths]
    if exec_indices and approve_indices:
        if min(approve_indices) > max(exec_indices):
            issues.append(f"审批节点均在执行节点之后，可能先施工后补审批")

    # 1.3 重复/高度相似节点
    for i in range(len(nodes)):
        for j in range(i + 1, len(nodes)):
            if nodes[i].name == nodes[j].name:
                issues.append(f"节点'{nodes[i].name}'在第{i+1}和第{j+1}位重复出现")
            elif difflib.SequenceMatcher(None, nodes[i].name, nodes[j].name).ratio() >= 0.9:
                issues.append(f"高度相似节点: '{nodes[i].name}' ↔ '{nodes[j].name}'（相似度>=90%）")

    # 1.4 同一岗位跨越不相容权力类型（既执行又核验）
    pos_auths = {}
    for n in nodes:
        if n.position:
            pos = n.position
            if pos not in pos_auths:
                pos_auths[pos] = set()
            pos_auths[pos].update(node_auth_types[nodes.index(n)])

    for pos, auths in pos_auths.items():
        if "执行" in auths and "核验" in auths:
            issues.append(f"岗位'{pos}'同时涉及执行和核验，存在不相容职责未分离风险")
        if "执行" in auths and "审批" in auths:
            issues.append(f"岗位'{pos}'同时涉及执行和审批，权力过于集中")
        if "发起" in auths and "审批" in auths:
            issues.append(f"岗位'{pos}'同时涉及发起和审批，缺少独立审批")

    return issues


def _check_closure(nodes: list, domain: str) -> list[str]:
    """检查业务流程是否闭环"""
    issues = []
    closure = CLOSURE_KEYWORDS.get(domain, CLOSURE_KEYWORDS.get("综合管理", {}))
    start_kw = closure.get("start", [])
    end_kw = closure.get("end", [])

    all_names = "".join(n.name for n in nodes)

    # 2.1 起点检查
    has_start = any(kw in all_names for kw in start_kw)
    if not has_start:
        issues.append(f"流程缺少明确起点（未出现: {'/'.join(start_kw[:4])}等关键词）")

    # 2.2 终点检查
    has_end = any(kw in all_names for kw in end_kw)
    if not has_end:
        issues.append(f"流程缺少明确终点（未出现: {'/'.join(end_kw[:4])}等关键词）")

    # 2.3 最低节点数
    min_nodes = MIN_PROCESS_NODES.get(domain, 2)
    if len(nodes) < min_nodes:
        issues.append(f"流程仅有{len(nodes)}个节点，低于{domain}的最低合理节点数({min_nodes})")

    # 2.4 "断头路"检测：某节点的输出似乎没有下游承接
    # 检查是否有"审批通过"但后续没有节点
    for i, n in enumerate(nodes):
        if i == len(nodes) - 1:
            continue
        # 如果某节点名含"审批"但下一个节点不明确接受审批结果
        if "审批" in n.name and "通过" not in n.name:
            next_node = nodes[i + 1]
            # 审批后应有明确执行节点
            has_next_exec = any(p in next_node.name for p in AUTHORITY_NODE_PATTERNS["执行"])
            if not has_next_exec:
                # 不强制报错，这是弱信号
                pass

    return issues


def _check_concealment(s: Submission) -> list[str]:
    """检查是否存在隐瞒信号"""
    issues = []

    # 3.1 模糊化节点名
    for n in s.nodes:
        for kw in VAGUE_NODE_KEYWORDS:
            if kw in n.name and len(n.name) <= 6:
                issues.append(f"节点'{n.name}'过于笼统，疑似模糊化处理")
                break

    # 3.2 不应合并的关键节点
    for n in s.nodes:
        for a, b in CRITICAL_MERGE_PAIRS:
            if a in n.name and b in n.name:
                issues.append(f"节点'{n.name}'合并了关键环节'{a}'和'{b}'，建议分开描述")
                break

    # 3.3 关键控制节点未标注岗位
    for n in s.nodes:
        if not n.position:
            is_critical = any(cp in n.name for cp in CRITICAL_POSITION_NODES)
            if is_critical:
                issues.append(f"关键节点'{n.name}'未标注岗位，无法确定责任人")

    # 3.4 流程异常简短（综合考虑节点数和文本长度）
    if s.raw_process:
        text_len = len(s.raw_process.strip())
        if text_len < 30 and len(s.nodes) <= 2:
            issues.append("流程描述过于简短（<30字），可能未全面反映实际情况")

    return issues


def _integrity_grade(score: float) -> str:
    if score >= 85:
        return "完整"
    if score >= 70:
        return "基本完整"
    if score >= 50:
        return "存在明显缺陷"
    return "严重不完整"


# ============================================================================
# 流程对照
# ============================================================================

def compare_nodes(ought_nodes: list, is_nodes: list) -> list[NodeDiff]:
    """
    逐节点对比应然与实然。
    算法：对应然每个节点找实然中最佳匹配（模糊），剩余实然节点标记为新增。
    """
    diffs = []
    is_pool = list(is_nodes)

    for i, onode in enumerate(ought_nodes):
        best_match = None
        best_ratio = 0
        best_idx = -1
        for j, inode in enumerate(is_pool):
            ratio = difflib.SequenceMatcher(None, onode.name, inode.name).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_match = inode
                best_idx = j

        if best_ratio >= 0.5:
            is_pool.pop(best_idx)
            if best_ratio >= 0.85:
                diff_type = "一致"
                desc = ""
            else:
                diff_type = "名称差异"
                desc = f"应然'{onode.name}' ↔ 实然'{best_match.name}'"
            if onode.position and best_match.position and onode.position != best_match.position:
                if diff_type == "一致":
                    diff_type = "岗位偏移"
                desc = (desc + "；" if desc else "") + f"岗位: 应然'{onode.position}' vs 实然'{best_match.position}'"
            elif onode.position and not best_match.position:
                if diff_type == "一致":
                    diff_type = "岗位偏移"
                desc = (desc + "；" if desc else "") + "实然未标注岗位"
            diffs.append(NodeDiff(
                seq=i + 1,
                ought_node=onode.name,
                is_node=best_match.name,
                ought_position=onode.position or "",
                is_position=best_match.position or "",
                diff_type=diff_type,
                description=desc,
            ))
        else:
            diffs.append(NodeDiff(
                seq=i + 1,
                ought_node=onode.name,
                is_node=None,
                ought_position=onode.position or "",
                is_position="",
                diff_type="节点消失",
                description=f"应然节点'{onode.name}'在实然流程中缺失",
            ))

    # 剩余实然节点 -> 新增
    for remaining in is_pool:
        diffs.append(NodeDiff(
            seq=len(diffs) + 1,
            ought_node=None,
            is_node=remaining.name,
            ought_position="",
            is_position=remaining.position or "",
            diff_type="节点新增",
            description=f"实然流程多出节点'{remaining.name}'（应然流程中无对应）",
        ))

    return diffs


def compare_against_checklist(nodes: list, checklist: list[dict]) -> list[NodeDiff]:
    """将实施流程与最小节点清单对照（无统筹单位时使用）"""
    diffs = []
    node_names = [n.name for n in nodes]

    for i, item in enumerate(checklist):
        checklist_name = item["name"]
        # 查找最佳匹配
        best_ratio = 0
        best_node = None
        for n in nodes:
            ratio = difflib.SequenceMatcher(None, checklist_name, n.name).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_node = n

        if best_ratio >= 0.5:
            pos = best_node.position if best_node else ""
            diffs.append(NodeDiff(
                seq=i + 1,
                ought_node=checklist_name,
                is_node=best_node.name,
                ought_position="(最小节点清单)",
                is_position=pos,
                diff_type="一致" if best_ratio >= 0.85 else "名称差异",
                description="" if best_ratio >= 0.85 else f"清单'{checklist_name}' ↔ 流程'{best_node.name}'",
            ))
        else:
            diffs.append(NodeDiff(
                seq=i + 1,
                ought_node=checklist_name,
                is_node=None,
                ought_position="(最小节点清单)",
                is_position="",
                diff_type="节点消失",
                description=f"最小节点'{checklist_name}'在流程中缺失",
            ))

    return diffs


def check_power_concentration(nodes: list, positions: list[str]) -> dict:
    """
    检查权力集中度：同一个岗位是否跨越多个权力类型节点。
    返回: {岗位: [权力类型列表]} 和集中度评分(0-1)
    """
    position_authorities = defaultdict(set)

    for node in nodes:
        pos = node.position if node.position else "未标注"
        for auth_type, patterns in AUTHORITY_NODE_PATTERNS.items():
            if any(p in node.name for p in patterns):
                position_authorities[pos].add(auth_type)

    # 计算集中度：某个岗位跨越的权力类型越多，越集中
    max_span = 0
    for pos, auths in position_authorities.items():
        span = len(auths)
        if span > max_span:
            max_span = span

    # 4种权力类型全占 → 1.0, 占3种 → 0.75, 占2种 → 0.5, 占1种 → 0.25
    concentration = max_span / 4.0 if max_span > 0 else 0

    return {
        "position_authorities": {p: list(a) for p, a in position_authorities.items()},
        "max_span": max_span,
        "concentration_score": concentration,
    }


# ============================================================================
# 风险评分
# ============================================================================

def score_node_completeness(nodes: list, checklist: list[dict]) -> float:
    """节点完整度评分 (1-5, 越高风险越大)"""
    if not checklist:
        return 3.0
    node_names = [n.name for n in nodes]
    found = 0
    critical_missing = 0
    for item in checklist:
        best = max(
            (difflib.SequenceMatcher(None, item["name"], nn).ratio() for nn in node_names),
            default=0,
        )
        if best >= 0.5:
            found += 1
        elif item.get("critical"):
            critical_missing += 1

    ratio = found / len(checklist)
    # 缺失越多 + 关键节点缺失 → 高分（高风险）
    score = (1 - ratio) * 4 + critical_missing * 0.5
    return min(5.0, max(1.0, score + 1.0))


def score_position_clarity(nodes: list) -> float:
    """岗位清晰度评分 (1-5, 越高风险越大)"""
    if not nodes:
        return 5.0
    empty_positions = sum(1 for n in nodes if not n.position)
    vague_positions = sum(1 for n in nodes if n.position and len(n.position) < 2)
    ratio = (empty_positions + vague_positions * 0.5) / len(nodes)
    return 1.0 + ratio * 4.0


def score_self_awareness(risk_desc: str, prevention: str) -> tuple[float, list[str]]:
    """自查风险自觉度评分 (1-5, 越高风险越大，即自查越敷衍)。
    返回 (score, flags) 其中 flags 为质量警告列表。"""
    score = 1.0
    flags = []

    # --- 空洞检测 ---
    if not risk_desc or risk_desc in ("无", "暂无", "未发现", "无风险", "/", "无廉洁风险"):
        score += 2.0
        flags.append("风险描述为空")
    elif len(risk_desc) < 10:
        score += 1.5
        flags.append("风险描述过短(<10字)")
    elif len(risk_desc) < 30:
        score += 0.5

    empty_measures = ("无", "暂无", "加强管理", "严格执行制度", "按制度执行", "/", "加强监督", "落实责任")
    if not prevention or prevention.strip() in empty_measures:
        score += 2.0
        flags.append("防控措施为空或套话")
    elif len(prevention) < 15:
        score += 1.0
        flags.append("防控措施过短(<15字)")

    # --- 抄模板检测 ---
    risk_tc = _detect_template_copy(risk_desc, "risk")
    measures_tc = _detect_template_copy(prevention, "measures")

    if risk_tc["suspicious"]:
        score += risk_tc["score_penalty"]
        flags.extend([f"[疑似抄模板-风险] {r}" for r in risk_tc["reasons"]])
    if measures_tc["suspicious"]:
        score += measures_tc["score_penalty"]
        flags.extend([f"[疑似抄模板-措施] {r}" for r in measures_tc["reasons"]])

    return min(5.0, score), flags


def score_business_risk(
    business_name: str,
    ought_subs: list[Submission],
    is_subs: list[Submission],
) -> BusinessRisk:
    """评估业务线综合风险"""
    checklist = MIN_NODE_CHECKLIST.get(business_name, [])

    # 取实施单位的平均情况
    if is_subs:
        avg_completeness = sum(
            score_node_completeness(s.nodes, checklist) for s in is_subs
        ) / len(is_subs)
        avg_clarity = sum(
            score_position_clarity(s.nodes) for s in is_subs
        ) / len(is_subs)
        avg_awareness = sum(
            score_self_awareness(s.risk_description, s.prevention_measures)[0] for s in is_subs
        ) / len(is_subs)

        # 权力集中度
        all_conc = []
        for s in is_subs:
            pc = check_power_concentration(s.nodes, s.positions)
            all_conc.append(pc["concentration_score"])
        avg_concentration = sum(all_conc) / len(all_conc) if all_conc else 0
        power_score = 1.0 + avg_concentration * 4.0
    else:
        # 无实施单位，用统筹单位数据
        all_subs = ought_subs + is_subs
        if all_subs:
            s = all_subs[0]
            avg_completeness = score_node_completeness(s.nodes, checklist)
            avg_clarity = score_position_clarity(s.nodes)
            avg_awareness = score_self_awareness(s.risk_description, s.prevention_measures)[0]
            pc = check_power_concentration(s.nodes, s.positions)
            power_score = 1.0 + pc["concentration_score"] * 4.0
        else:
            return BusinessRisk(business_name, 3, 3, 3, 3, 3, 3.0, "C", "中等风险（数据不足）")

    # 应然-实然匹配度
    if ought_subs and is_subs:
        all_match_scores = []
        for ought in ought_subs:
            for is_sub in is_subs:
                diffs = compare_nodes(ought.nodes, is_sub.nodes)
                total = len(diffs)
                if total == 0:
                    all_match_scores.append(1.0)
                    continue
                problem_count = sum(1 for d in diffs if d.diff_type in ("节点消失", "节点新增", "岗位偏移"))
                match_ratio = problem_count / total
                all_match_scores.append(1.0 + match_ratio * 4.0)
        match_score = sum(all_match_scores) / len(all_match_scores)
    else:
        match_score = 3.0  # 无法评估

    # 加权总分
    total = (
        avg_completeness * RISK_WEIGHTS["node_completeness"]
        + avg_clarity * RISK_WEIGHTS["position_clarity"]
        + power_score * RISK_WEIGHTS["power_concentration"]
        + match_score * RISK_WEIGHTS["ought_is_match"]
        + avg_awareness * RISK_WEIGHTS["self_awareness"]
    ) / sum(RISK_WEIGHTS.values())

    # 确定等级
    level = "D"
    level_desc = "低风险"
    for threshold, lvl, desc in RISK_THRESHOLDS:
        if total >= threshold:
            level = lvl
            level_desc = desc
            break

    return BusinessRisk(
        business_name=business_name,
        node_completeness=round(avg_completeness, 2),
        position_clarity=round(avg_clarity, 2),
        power_concentration=round(power_score, 2),
        ought_is_match=round(match_score, 2),
        self_awareness=round(avg_awareness, 2),
        total_score=round(total, 2),
        level=level,
        level_desc=level_desc,
    )


def score_unit_quality(s: Submission, checklist: list[dict]) -> dict:
    """评估单个单位的自查质量（百分制），含抄模板检测扣分"""
    completeness = (1 - (score_node_completeness(s.nodes, checklist) - 1) / 4) * 25  # 节点完整
    clarity = (1 - (score_position_clarity(s.nodes) - 1) / 4) * 25  # 岗位清晰
    awareness_score, awareness_flags = score_self_awareness(s.risk_description, s.prevention_measures)
    awareness = (1 - (awareness_score - 1) / 4) * 25  # 风险自觉（含抄模板扣分）
    position_detail = min(25.0, len(s.positions) * 5.0) if s.positions else 5  # 岗位标注

    # 抄模板检测也影响节点完整度
    process_tc = _detect_template_copy(s.raw_process, "process")
    template_penalty = 0.0
    suspicious_flags = list(awareness_flags)
    if process_tc["suspicious"]:
        template_penalty += process_tc["score_penalty"] * 3  # 流程抄模板额外扣
        suspicious_flags.extend([f"[疑似抄模板-流程] {r}" for r in process_tc["reasons"]])

    total = max(0.0, completeness + clarity + awareness + position_detail - template_penalty)
    return {
        "unit": s.unit_name,
        "role": s.unit_role,
        "business": s.business_name_canonical or s.business_name_raw,
        "completeness": round(completeness, 1),
        "clarity": round(clarity, 1),
        "awareness": round(awareness, 1),
        "position_detail": round(position_detail, 1),
        "total": round(total, 1),
        "grade": _grade(total),
        "flags": suspicious_flags,
    }


def _grade(score: float) -> str:
    if score >= 85:
        return "优"
    if score >= 70:
        return "良"
    if score >= 55:
        return "中"
    return "差"


# ============================================================================
# 同业务实施单位差异对比
# ============================================================================

def compare_implementation_units(business_name: str, is_subs: list[Submission]) -> list[dict]:
    """
    比较同一业务不同实施单位的流程差异。
    产出一个差异列表，每行描述两个单位在同一节点上的不同做法。
    """
    results = []
    if len(is_subs) < 2:
        return results

    # 构建节点名字集合（所有单位）
    all_node_sets = {}
    for s in is_subs:
        all_node_sets[s.unit_name] = {n.name for n in s.nodes}

    # 两两对比
    for i in range(len(is_subs)):
        for j in range(i + 1, len(is_subs)):
            a = is_subs[i]
            b = is_subs[j]
            a_nodes = all_node_sets[a.unit_name]
            b_nodes = all_node_sets[b.unit_name]

            # A 有 B 无
            a_only = a_nodes - b_nodes
            for node_name in a_only:
                # 检查是否是名称差异（B中有相似节点）
                best_ratio = max(
                    (difflib.SequenceMatcher(None, node_name, bn).ratio() for bn in b_nodes),
                    default=0,
                )
                if best_ratio < 0.6:
                    results.append({
                        "business": business_name,
                        "unit_a": a.unit_name,
                        "unit_b": b.unit_name,
                        "diff_node": node_name,
                        "diff_type": f"{a.unit_name}有此节点，{b.unit_name}无",
                        "detail_a": _find_node_detail(a.nodes, node_name),
                        "detail_b": "（缺失）",
                        "impact": _assess_impact(node_name, "missing", business_name),
                    })

            # B 有 A 无
            b_only = b_nodes - a_nodes
            for node_name in b_only:
                best_ratio = max(
                    (difflib.SequenceMatcher(None, node_name, an).ratio() for an in a_nodes),
                    default=0,
                )
                if best_ratio < 0.6:
                    results.append({
                        "business": business_name,
                        "unit_a": a.unit_name,
                        "unit_b": b.unit_name,
                        "diff_node": node_name,
                        "diff_type": f"{b.unit_name}有此节点，{a.unit_name}无",
                        "detail_a": "（缺失）",
                        "detail_b": _find_node_detail(b.nodes, node_name),
                        "impact": _assess_impact(node_name, "missing", business_name),
                    })

            # 岗位差异（同一节点不同岗位）
            for an in a.nodes:
                for bn in b.nodes:
                    sim = difflib.SequenceMatcher(None, an.name, bn.name).ratio()
                    if sim >= 0.7 and an.position != bn.position:
                        # 两者都有岗位标注且不同
                        if an.position and bn.position:
                            results.append({
                                "business": business_name,
                                "unit_a": a.unit_name,
                                "unit_b": b.unit_name,
                                "diff_node": f"{an.name} ≈ {bn.name}",
                                "diff_type": "岗位不同",
                                "detail_a": f"岗位: {an.position}",
                                "detail_b": f"岗位: {bn.position}",
                                "impact": _assess_impact(an.name, "position_diff", business_name),
                            })
                        break

    return results


def _find_node_detail(nodes: list, name: str) -> str:
    for n in nodes:
        if n.name == name:
            return f"节点: {n.name}" + (f"，岗位: {n.position}" if n.position else "（未标注岗位）")
    return ""


def _assess_impact(node_name: str, diff_type: str, business_name: str) -> str:
    """评估差异的潜在影响"""
    if diff_type == "missing":
        if any(kw in node_name for kw in ("验收", "审批", "审核", "结算")):
            return "高风险：缺少关键控制节点，可能导致该环节失去监督"
        elif any(kw in node_name for kw in ("招标", "采购", "合同")):
            return "高风险：缺少合规审查节点，存在程序违规可能"
        elif any(kw in node_name for kw in ("核验", "检查", "盘点")):
            return "中风险：缺少核验节点，可能造成数据不实"
        else:
            return "低风险：非关键节点缺失，但仍需确认是否被合并到其他环节"
    elif diff_type == "position_diff":
        return "中风险：同一环节在不同单位由不同岗位执行，需确认是否符合职责分离原则"
    return "待评估"


# ============================================================================
# 现场核查建议生成
# ============================================================================

def generate_verification_suggestions(
    business_name: str, risk_level: str, diffs: list[NodeDiff]
) -> dict:
    """为高风险业务生成现场核查建议"""
    template = VERIFICATION_TEMPLATES.get(business_name, VERIFICATION_TEMPLATES.get("综合管理", {}))

    # 根据具体差异补充针对性建议
    extra_docs = []
    extra_questions = []

    for d in diffs:
        if d.diff_type == "节点消失" and d.ought_node:
            extra_questions.append(f"为什么'{d.ought_node}'环节在实际流程中缺失？是被合并、跳过还是另有替代流程？")
        elif d.diff_type == "节点新增" and d.is_node:
            extra_questions.append(f"新增的'{d.is_node}'环节由谁发起？目的是什么？是否有书面记录？")
        elif d.diff_type == "岗位偏移":
            extra_questions.append(
                f"'{d.ought_node}'环节的岗位从应然的'{d.ought_position}'偏移为实然的'{d.is_position}'，原因是什么？"
            )

    return {
        "documents": template.get("documents", []),
        "interview_directions": template.get("interview_directions", []),
        "general_questions": template.get("questions", []),
        "targeted_questions": extra_questions[:8],  # 最多 8 条针对性追问
    }


# ============================================================================
# Excel 报告生成
# ============================================================================

# 样式常量
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name="微软雅黑", size=10)
WRAP_ALIGN = Alignment(wrap_text=True, vertical="top")
CENTER_ALIGN = Alignment(horizontal="center", vertical="top")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
ORANGE_FILL = PatternFill(start_color="FFE0B2", end_color="FFE0B2", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
GREEN_FILL = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")


def _style_header(ws, headers: list[str]):
    """写入并样式化表头"""
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER


def _style_body(ws, start_row: int, end_row: int, ncols: int):
    """样式化表体"""
    for row in range(start_row, end_row + 1):
        for col in range(1, ncols + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = BODY_FONT
            cell.alignment = WRAP_ALIGN
            cell.border = THIN_BORDER


def _auto_width(ws, min_width: int = 8, max_width: int = 50):
    """自动调整列宽"""
    for col_cells in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            if cell.value:
                # 中文字符按2个字符宽度计算
                text = str(cell.value)
                length = sum(2 if ord(c) > 127 else 1 for c in text)
                # 取每行最大长度（考虑到换行）
                for line in text.split('\n'):
                    line_len = sum(2 if ord(c) > 127 else 1 for c in line)
                    max_len = max(max_len, line_len)
        width = min(max(max_len * 1.1, min_width), max_width)
        ws.column_dimensions[col_letter].width = width


def _apply_diff_fill(ws, row: int, diff_type: str, ncols: int):
    """根据差异类型填充行颜色"""
    fill_map = {
        "节点消失": RED_FILL,
        "岗位偏移": ORANGE_FILL,
        "节点新增": YELLOW_FILL,
    }
    fill = fill_map.get(diff_type)
    if fill:
        for col in range(1, ncols + 1):
            ws.cell(row=row, column=col).fill = fill


def generate_excel_report(
    output_path: str,
    pairs: dict,
    business_risks: list[BusinessRisk],
    unit_qualities: list[dict],
    impl_comparisons: list[dict],
    all_diffs: dict,
    integrity_results: list[dict] | None = None,
):
    """生成完整的 Excel 分析报告"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # 删除默认 sheet

    # ---- Sheet 1: 应然-实然对照矩阵 ----
    ws1 = wb.create_sheet("应然-实然对照矩阵")
    headers1 = ["业务线", "统筹单位", "实施单位", "节点序号", "应然节点", "实然节点",
                 "应然岗位", "实然岗位", "差异类型", "差异说明"]
    _style_header(ws1, headers1)
    row = 2
    for biz_name, data in pairs.items():
        ought_subs = data["ought"]
        is_subs = data["is"]
        checklist = MIN_NODE_CHECKLIST.get(biz_name, [])

        if ought_subs:
            for ought in ought_subs:
                for is_sub in is_subs:
                    diffs = compare_nodes(ought.nodes, is_sub.nodes)
                    for d in diffs:
                        ws1.cell(row=row, column=1, value=biz_name)
                        ws1.cell(row=row, column=2, value=ought.unit_name)
                        ws1.cell(row=row, column=3, value=is_sub.unit_name)
                        ws1.cell(row=row, column=4, value=d.seq)
                        ws1.cell(row=row, column=5, value=d.ought_node or "")
                        ws1.cell(row=row, column=6, value=d.is_node or "")
                        ws1.cell(row=row, column=7, value=d.ought_position or "")
                        ws1.cell(row=row, column=8, value=d.is_position or "")
                        ws1.cell(row=row, column=9, value=d.diff_type)
                        ws1.cell(row=row, column=10, value=d.description)
                        _apply_diff_fill(ws1, row, d.diff_type, 10)
                        row += 1
        else:
            # 无统筹单位，与最小节点清单对比
            for is_sub in is_subs:
                diffs = compare_against_checklist(is_sub.nodes, checklist)
                for d in diffs:
                    ws1.cell(row=row, column=1, value=biz_name)
                    ws1.cell(row=row, column=2, value="（最小节点清单）")
                    ws1.cell(row=row, column=3, value=is_sub.unit_name)
                    ws1.cell(row=row, column=4, value=d.seq)
                    ws1.cell(row=row, column=5, value=d.ought_node or "")
                    ws1.cell(row=row, column=6, value=d.is_node or "")
                    ws1.cell(row=row, column=7, value=d.ought_position or "")
                    ws1.cell(row=row, column=8, value=d.is_position or "")
                    ws1.cell(row=row, column=9, value=d.diff_type)
                    ws1.cell(row=row, column=10, value=d.description)
                    _apply_diff_fill(ws1, row, d.diff_type, 10)
                    row += 1

    _style_body(ws1, 2, row - 1, 10)
    _auto_width(ws1)
    ws1.auto_filter.ref = ws1.dimensions

    # ---- Sheet 2: 自查质量排序 ----
    ws2 = wb.create_sheet("自查质量排序")
    headers2 = ["排名", "单位名称", "单位职能", "业务线", "节点完整度(25)", "岗位清晰度(25)",
                 "风险自觉度(25)", "岗位标注度(25)", "总分(100)", "等级"]
    _style_header(ws2, headers2)
    unit_qualities.sort(key=lambda x: x["total"], reverse=True)
    for i, q in enumerate(unit_qualities):
        row = i + 2
        ws2.cell(row=row, column=1, value=i + 1)
        ws2.cell(row=row, column=2, value=q["unit"])
        ws2.cell(row=row, column=3, value=q["role"])
        ws2.cell(row=row, column=4, value=q["business"])
        ws2.cell(row=row, column=5, value=q["completeness"])
        ws2.cell(row=row, column=6, value=q["clarity"])
        ws2.cell(row=row, column=7, value=q["awareness"])
        ws2.cell(row=row, column=8, value=q["position_detail"])
        ws2.cell(row=row, column=9, value=q["total"])
        ws2.cell(row=row, column=10, value=q["grade"])
        if q["grade"] == "差":
            for c in range(1, 11):
                ws2.cell(row=row, column=c).fill = RED_FILL
        elif q["grade"] == "中":
            for c in range(1, 11):
                ws2.cell(row=row, column=c).fill = YELLOW_FILL
    _style_body(ws2, 2, len(unit_qualities) + 1, 10)
    _auto_width(ws2)
    ws2.auto_filter.ref = ws2.dimensions

    # ---- Sheet 3: 高风险靶点清单 ----
    ws3 = wb.create_sheet("高风险靶点清单")
    headers3 = ["优先级", "业务线", "单位", "差异节点", "差异类型", "风险描述", "建议追问"]
    _style_header(ws3, headers3)
    row = 2
    for biz_name, data in pairs.items():
        ought_subs = data["ought"]
        is_subs = data["is"]

        if ought_subs:
            for ought in ought_subs:
                for is_sub in is_subs:
                    diffs = compare_nodes(ought.nodes, is_sub.nodes)
                    for d in diffs:
                        if d.diff_type in ("一致", "名称差异"):
                            continue
                        priority, _ = DIFF_PRIORITY.get(d.diff_type, ("-", 5))
                        if priority in ("红", "橙"):
                            ws3.cell(row=row, column=1, value=priority)
                            ws3.cell(row=row, column=2, value=biz_name)
                            ws3.cell(row=row, column=3, value=is_sub.unit_name)
                            ws3.cell(row=row, column=4, value=d.ought_node or d.is_node or "")
                            ws3.cell(row=row, column=5, value=d.diff_type)
                            ws3.cell(row=row, column=6, value=d.description)
                            ws3.cell(row=row, column=7, value=_generate_pursuit_question(d, biz_name))
                            _apply_diff_fill(ws3, row, d.diff_type, 7)
                            row += 1
    _style_body(ws3, 2, row - 1, 7)
    _auto_width(ws3)
    ws3.auto_filter.ref = ws3.dimensions

    # ---- Sheet 4: 业务风险评估 ----
    ws4 = wb.create_sheet("业务风险评估")
    headers4 = ["业务线", "节点完整度", "岗位清晰度", "权力集中度", "应然-实然匹配度",
                 "自查风险自觉度", "加权总分", "风险等级"]
    _style_header(ws4, headers4)
    business_risks.sort(key=lambda x: x.total_score, reverse=True)
    for i, br in enumerate(business_risks):
        row = i + 2
        ws4.cell(row=row, column=1, value=br.business_name)
        ws4.cell(row=row, column=2, value=br.node_completeness)
        ws4.cell(row=row, column=3, value=br.position_clarity)
        ws4.cell(row=row, column=4, value=br.power_concentration)
        ws4.cell(row=row, column=5, value=br.ought_is_match)
        ws4.cell(row=row, column=6, value=br.self_awareness)
        ws4.cell(row=row, column=7, value=br.total_score)
        ws4.cell(row=row, column=8, value=f"{br.level}（{br.level_desc}）")
        if br.level == "A":
            for c in range(1, 9):
                ws4.cell(row=row, column=c).fill = RED_FILL
        elif br.level == "B":
            for c in range(1, 9):
                ws4.cell(row=row, column=c).fill = ORANGE_FILL
    _style_body(ws4, 2, len(business_risks) + 1, 8)
    _auto_width(ws4)
    ws4.auto_filter.ref = ws4.dimensions

    # ---- Sheet 5: 现场核查建议 ----
    ws5 = wb.create_sheet("现场核查建议")
    headers5 = ["业务线", "风险等级", "建议查阅资料", "谈话方向", "重点谈话问题"]
    _style_header(ws5, headers5)
    row = 2
    for br in business_risks:
        if br.level in ("A", "B"):
            suggestions = generate_verification_suggestions(
                br.business_name, br.level,
                all_diffs.get(br.business_name, []),
            )
            docs = "\n".join(f"• {d}" for d in suggestions["documents"])
            directions = "\n".join(f"• {d}" for d in suggestions["interview_directions"])
            questions = "\n".join(
                f"• {q}" for q in suggestions["general_questions"] + suggestions["targeted_questions"]
            )
            ws5.cell(row=row, column=1, value=br.business_name)
            ws5.cell(row=row, column=2, value=f"{br.level}（{br.level_desc}）")
            ws5.cell(row=row, column=3, value=docs)
            ws5.cell(row=row, column=4, value=directions)
            ws5.cell(row=row, column=5, value=questions)
            if br.level == "A":
                for c in range(1, 6):
                    ws5.cell(row=row, column=c).fill = RED_FILL
            else:
                for c in range(1, 6):
                    ws5.cell(row=row, column=c).fill = ORANGE_FILL
            row += 1
    _style_body(ws5, 2, row - 1, 5)
    _auto_width(ws5)
    # 核查建议列更宽
    ws5.column_dimensions['C'].width = 50
    ws5.column_dimensions['D'].width = 50
    ws5.column_dimensions['E'].width = 60

    # ---- Sheet 6: 同业务差异对比 ----
    ws6 = wb.create_sheet("同业务差异对比")
    headers6 = ["业务线", "实施单位A", "实施单位B", "差异节点", "差异类型", "A描述", "B描述", "预计影响"]
    _style_header(ws6, headers6)
    for i, comp in enumerate(impl_comparisons):
        row = i + 2
        ws6.cell(row=row, column=1, value=comp["business"])
        ws6.cell(row=row, column=2, value=comp["unit_a"])
        ws6.cell(row=row, column=3, value=comp["unit_b"])
        ws6.cell(row=row, column=4, value=comp["diff_node"])
        ws6.cell(row=row, column=5, value=comp["diff_type"])
        ws6.cell(row=row, column=6, value=comp["detail_a"])
        ws6.cell(row=row, column=7, value=comp["detail_b"])
        ws6.cell(row=row, column=8, value=comp["impact"])
        if "高风险" in comp.get("impact", ""):
            for c in range(1, 9):
                ws6.cell(row=row, column=c).fill = RED_FILL
        elif "中风险" in comp.get("impact", ""):
            for c in range(1, 9):
                ws6.cell(row=row, column=c).fill = YELLOW_FILL
    _style_body(ws6, 2, len(impl_comparisons) + 1, 8)
    _auto_width(ws6)
    ws6.auto_filter.ref = ws6.dimensions

    # ---- Sheet 7: 流程完整性分析 ----
    if integrity_results:
        ws7 = wb.create_sheet("流程完整性分析")
        headers7 = ["排名", "单位名称", "角色", "排查领域", "业务名称", "流程节点数",
                     "完整性评分", "等级", "问题维度", "具体问题"]
        _style_header(ws7, headers7)
        integrity_results.sort(key=lambda x: x["score"])
        row = 2
        for ir in integrity_results:
            if not ir.get("issues"):
                # 无问题的也输出一行（仅基本信息）
                ws7.cell(row=row, column=1, value=row - 1)
                ws7.cell(row=row, column=2, value=ir["unit"])
                ws7.cell(row=row, column=3, value=ir.get("role", ""))
                ws7.cell(row=row, column=4, value=ir["domain"])
                ws7.cell(row=row, column=5, value=ir["business"])
                ws7.cell(row=row, column=6, value=ir["node_count"])
                ws7.cell(row=row, column=7, value=ir["score"])
                ws7.cell(row=row, column=8, value=ir["grade"])
                ws7.cell(row=row, column=9, value="无")
                ws7.cell(row=row, column=10, value="流程描述基本合理")
                _style_body(ws7, row, row, 10)
                row += 1
            else:
                for issue in ir["issues"]:
                    ws7.cell(row=row, column=1, value=row - 1)
                    ws7.cell(row=row, column=2, value=ir["unit"])
                    ws7.cell(row=row, column=3, value=ir.get("role", ""))
                    ws7.cell(row=row, column=4, value=ir["domain"])
                    ws7.cell(row=row, column=5, value=ir["business"])
                    ws7.cell(row=row, column=6, value=ir["node_count"])
                    ws7.cell(row=row, column=7, value=ir["score"])
                    ws7.cell(row=row, column=8, value=ir["grade"])
                    ws7.cell(row=row, column=9, value=issue["维度"])
                    ws7.cell(row=row, column=10, value=issue["问题"])
                    if ir["grade"] in ("存在明显缺陷", "严重不完整"):
                        for c in range(1, 11):
                            ws7.cell(row=row, column=c).fill = RED_FILL
                    elif "隐瞒" in issue["维度"]:
                        for c in range(1, 11):
                            ws7.cell(row=row, column=c).fill = ORANGE_FILL
                    row += 1
        _style_body(ws7, 2, row - 1, 10)
        _auto_width(ws7)
        ws7.column_dimensions['J'].width = 60
        ws7.auto_filter.ref = ws7.dimensions

    wb.save(output_path)
    print(f"[报告] Excel 分析报告已保存至: {output_path}")


def _generate_pursuit_question(diff: NodeDiff, business_name: str) -> str:
    """根据差异类型生成追问问题"""
    if diff.diff_type == "节点消失":
        return f"'{diff.ought_node}'环节在实际中是如何执行的？是合并在其他环节中、还是被省略了？请提供相关佐证材料。"
    elif diff.diff_type == "岗位偏移":
        return f"'{diff.ought_node}'环节的审批权限为何从'{diff.ought_position}'变为'{diff.is_position}'？是否有正式授权文件？"
    elif diff.diff_type == "节点新增":
        return f"新增的'{diff.is_node}'环节由谁发起、解决什么问题？是否有制度依据？执行中是否产生书面记录？"
    return ""


# ============================================================================
# Markdown 核验手册生成
# ============================================================================

def generate_markdown_handbook(
    output_path: str,
    pairs: dict,
    business_risks: list[BusinessRisk],
    impl_comparisons: list[dict],
    all_diffs: dict,
):
    """生成可直接打印带往现场的 Markdown 核验手册"""
    lines = []
    lines.append("# 廉洁风险排查——现场核验手册")
    lines.append(f"\n> 生成日期：{datetime.now().strftime('%Y年%m月%d日')}")
    lines.append(f"> 本手册根据各单位自查材料生成，供重点排查阶段（7月-8月）现场核验使用。\n")

    # 目录
    lines.append("## 目录\n")
    for br in business_risks:
        lines.append(f"- [{br.business_name}（{br.level_desc}）](#{_anchor(br.business_name)})")
    lines.append("")

    # 按风险等级排序
    business_risks.sort(key=lambda x: x.total_score, reverse=True)

    for br in business_risks:
        data = pairs.get(br.business_name, {"ought": [], "is": []})
        lines.append(f"## {br.business_name}")
        lines.append(f"\n**风险等级：{br.level}（{br.level_desc}）** | 加权总分：{br.total_score}/5.0\n")

        # 风险维度概览
        lines.append("### 风险维度得分\n")
        lines.append(f"| 维度 | 得分 | 说明 |")
        lines.append(f"|------|------|------|")
        lines.append(f"| 节点完整度 | {br.node_completeness:.1f} | 与最小节点清单对照，缺失越多得分越高 |")
        lines.append(f"| 岗位清晰度 | {br.position_clarity:.1f} | 岗位标注越模糊得分越高 |")
        lines.append(f"| 权力集中度 | {br.power_concentration:.1f} | 同一岗位跨越权力类型越多得分越高 |")
        lines.append(f"| 应然-实然匹配度 | {br.ought_is_match:.1f} | 统筹与实施流程差异越大得分越高 |")
        lines.append(f"| 自查风险自觉度 | {br.self_awareness:.1f} | 自查越空洞得分越高 |")
        lines.append("")

        # 应然-实然对照
        if data["ought"] and data["is"]:
            lines.append("### 应然-实然差异清单\n")
            for ought in data["ought"]:
                for is_sub in data["is"]:
                    diffs = compare_nodes(ought.nodes, is_sub.nodes)
                    red_diffs = [d for d in diffs if d.diff_type in ("节点消失", "岗位偏移", "节点新增")]
                    if red_diffs:
                        lines.append(f"**统筹单位：{ought.unit_name} ↔ 实施单位：{is_sub.unit_name}**\n")
                        lines.append(f"| # | 应然节点 | 实然节点 | 差异类型 | 差异说明 |")
                        lines.append(f"|---|----------|----------|----------|----------|")
                        for d in red_diffs:
                            lines.append(
                                f"| {d.seq} | {d.ought_node or '-'} | {d.is_node or '-'} "
                                f"| {d.diff_type} | {d.description} |"
                            )
                        lines.append("")

        # 现场核查建议
        if br.level in ("A", "B"):
            suggestions = generate_verification_suggestions(
                br.business_name, br.level,
                all_diffs.get(br.business_name, []),
            )
            lines.append("### 建议查阅资料\n")
            for doc in suggestions["documents"]:
                lines.append(f"- [ ] {doc}")
            lines.append("")

            lines.append("### 谈话方向\n")
            for direction in suggestions["interview_directions"]:
                lines.append(f"- {direction}")
            lines.append("")

            lines.append("### 重点谈话问题\n")
            for q in suggestions["general_questions"]:
                lines.append(f"- {q}")
            if suggestions["targeted_questions"]:
                lines.append("\n**针对性追问（根据自查材料发现）：**\n")
                for q in suggestions["targeted_questions"]:
                    lines.append(f"- {q}")
            lines.append("")

    # 同业务实施单位差异
    if impl_comparisons:
        lines.append("## 附录：同业务不同实施单位差异对比\n")
        lines.append("以下列出同一业务在不同实施单位之间存在显著差异的情况，现场核验时应重点关注。\n")
        by_business = defaultdict(list)
        for comp in impl_comparisons:
            by_business[comp["business"]].append(comp)
        for biz_name, comps in by_business.items():
            lines.append(f"### {biz_name}\n")
            lines.append(f"| 单位A | 单位B | 差异节点 | 差异类型 | 预计影响 |")
            lines.append(f"|-------|-------|----------|----------|----------|")
            for comp in comps:
                lines.append(
                    f"| {comp['unit_a']} | {comp['unit_b']} | {comp['diff_node']} "
                    f"| {comp['diff_type']} | {comp['impact']} |"
                )
            lines.append("")

    # 页脚
    lines.append("---")
    lines.append(f"\n*本手册由 lzpc 分析工具自动生成，基于各单位提交的自查材料。*")
    lines.append(f"*如需调整分析参数，请编辑 analyzer.py 中的配置项。*")

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f"[报告] 核验手册已保存至: {output_path}")


def _anchor(text: str) -> str:
    """生成 Markdown 锚点"""
    return text.replace("（", "").replace("）", "").replace(" ", "-").lower()


# ============================================================================
# Docx 文字总结报告
# ============================================================================

def _docx_add_heading(doc, text, level=1):
    """添加带格式的标题"""
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.name = "微软雅黑"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    return h


def _docx_add_para(doc, text, bold=False, indent=False):
    """添加正文段落"""
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.74)
    run = p.add_run(text)
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    run.font.size = Pt(11)
    run.bold = bold
    return p


def _docx_add_table(doc, headers: list[str], rows: list[list[str]], col_widths=None):
    """添加带格式的表格"""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers), style="Table Grid")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 表头
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.name = "微软雅黑"
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
                run.font.size = Pt(10)
                run.bold = True
        # 表头底色
        shading = cell._element.get_or_add_tcPr()
        shd = shading.makeelement(qn("w:shd"), {
            qn("w:fill"): "4472C4",
            qn("w:val"): "clear",
        })
        shading.append(shd)
        for run in cell.paragraphs[0].runs:
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # 数据行
    for r, row_data in enumerate(rows):
        for c, val in enumerate(row_data):
            cell = table.rows[r + 1].cells[c]
            cell.text = str(val)
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.name = "微软雅黑"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
                    run.font.size = Pt(10)

    if col_widths:
        for i, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(w)

    doc.add_paragraph()  # 表后空行
    return table


def generate_docx_summary(
    output_path: str,
    submissions: list,
    pairs: dict,
    business_risks: list[BusinessRisk],
    unit_qualities: list[dict],
    integrity_results: list[dict],
    orphans: dict,
    all_diffs: dict,
    cat_issues: list[dict],
    domain_conflicts: list,
    impl_comparisons: list[dict],
):
    """生成文字版分析总结报告（.docx）"""
    if not DOCX_AVAILABLE:
        print("[警告] python-docx 未安装，跳过 docx 报告生成。请执行: pip install python-docx")
        return

    doc = Document()

    # 设置默认字体
    style = doc.styles["Normal"]
    font = style.font
    font.name = "微软雅黑"
    font.size = Pt(11)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

    # ── 封面标题 ──
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(60)
    run = title.add_run("廉洁风险排查分析总结报告")
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    run.font.size = Pt(22)
    run.bold = True

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run(f"生成日期：{datetime.now().strftime('%Y年%m月%d日')}")
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_paragraph()

    # ── 一、总体情况 ──
    _docx_add_heading(doc, "一、总体情况", level=1)
    unit_count = len(set(s.unit_name for s in submissions))
    domain_dist = defaultdict(int)
    role_dist = defaultdict(int)
    for s in submissions:
        domain_dist[s.inspection_domain or "未知"] += 1
        role_dist[s.unit_role] += 1
    _docx_add_para(doc, f"本次排查共收到 {len(submissions)} 条自查记录，涉及 {unit_count} 个单位。", indent=True)
    _docx_add_para(doc, f"排查领域分布：{', '.join(f'{k}({v}条)' for k, v in domain_dist.items())}。", indent=True)
    _docx_add_para(doc, f"角色分布：统筹 {role_dist.get('统筹', 0)} 条、实施 {role_dist.get('实施', 0)} 条、双重 {role_dist.get('双重', 0)} 条。", indent=True)

    # ── 二、高风险业务线 ──
    _docx_add_heading(doc, "二、高风险业务线评估", level=1)
    high_risk = [br for br in business_risks if br.level in ("A", "B")]
    high_risk.sort(key=lambda x: x.total_score, reverse=True)
    if high_risk:
        _docx_add_para(doc, f"经五维加权评分，以下 {len(high_risk)} 条业务线风险等级为 A 或 B，建议优先安排重点排查和现场核验：", indent=True)
        headers = ["业务线", "节点完整度", "岗位清晰度", "权力集中度", "匹配度", "自觉度", "总分", "等级"]
        rows_data = [
            [br.business_name, str(br.node_completeness), str(br.position_clarity),
             str(br.power_concentration), str(br.ought_is_match), str(br.self_awareness),
             str(br.total_score), f"{br.level}（{br.level_desc}）"]
            for br in high_risk
        ]
        _docx_add_table(doc, headers, rows_data)
    else:
        _docx_add_para(doc, "本次分析未发现 A/B 级高风险业务线。", indent=True)

    # 风险等级说明
    _docx_add_para(doc, "评分维度说明：节点完整度反映流程与最小节点清单的差距；岗位清晰度衡量各环节岗位标注的明确程度；权力集中度检测同一岗位跨越不相容权力类型的程度；匹配度反映统筹流程与实施流程的一致性；自觉度评估自查风险描述和防控措施的质量（含抄模板检测）。", indent=True)

    # ── 三、自查质量评估 ──
    _docx_add_heading(doc, "三、自查质量评估", level=1)
    _docx_add_para(doc, "从节点完整度、岗位清晰度、风险自觉度、岗位标注度四个维度对每条自查记录进行百分制评分。", indent=True)

    # 质量分布
    grade_dist = defaultdict(int)
    for q in unit_qualities:
        grade_dist[q["grade"]] += 1
    _docx_add_para(doc, f"等级分布：优 {grade_dist.get('优', 0)} 条、良 {grade_dist.get('良', 0)} 条、中 {grade_dist.get('中', 0)} 条、差 {grade_dist.get('差', 0)} 条。", indent=True)

    worst = sorted(unit_qualities, key=lambda x: x["total"])[:5]
    if worst:
        _docx_add_para(doc, "自查质量最差的 5 条记录：", indent=True)
        headers = ["单位", "业务线", "总分", "等级", "主要问题"]
        rows_data = [
            [q["unit"], q["business"], str(q["total"]), q["grade"],
             "; ".join(q.get("flags", [])[:2]) if q.get("flags") else "无明显问题"]
            for q in worst
        ]
        _docx_add_table(doc, headers, rows_data)

    # 抄模板/敷衍
    template_suspects = [q for q in unit_qualities if q.get("flags")]
    if template_suspects:
        _docx_add_para(doc, f"疑似抄模板或敷衍填写 {len(template_suspects)} 条，建议重点复核以下单位：", indent=True)
        for q in sorted(template_suspects, key=lambda x: x["total"])[:8]:
            flags_str = "；".join(q["flags"][:3])
            _docx_add_para(doc, f"• {q['unit']}（{q['business']}）— {q['total']}分 [{q['grade']}]：{flags_str}", indent=True)

    # ── 四、流程完整性 ──
    _docx_add_heading(doc, "四、流程完整性分析", level=1)
    if integrity_results:
        defect_count = sum(1 for ir in integrity_results if ir["grade"] in ("存在明显缺陷", "严重不完整"))
        severe_count = sum(1 for ir in integrity_results if ir["grade"] == "严重不完整")
        conceal_count = sum(1 for ir in integrity_results
                           if any(iss["维度"] == "隐瞒信号" for iss in ir.get("issues", [])))
        _docx_add_para(doc, f"逐条独立审查 {len(integrity_results)} 条流程描述的逻辑合理性、业务闭环和隐瞒信号。存在明显缺陷 {defect_count} 条（其中严重不完整 {severe_count} 条），存在隐瞒信号 {conceal_count} 条。", indent=True)

        worst_integrity = sorted(integrity_results, key=lambda x: x["score"])[:8]
        if worst_integrity:
            _docx_add_para(doc, "完整性最差的流程：", indent=True)
            headers = ["单位", "领域", "节点数", "评分", "等级", "主要问题"]
            rows_data = [
                [ir["unit"], ir["domain"], str(ir["node_count"]), str(ir["score"]), ir["grade"],
                 "；".join(iss["问题"][:40] for iss in ir.get("issues", [])[:3])]
                for ir in worst_integrity
            ]
            _docx_add_table(doc, headers, rows_data)

        # 隐瞒信号详情
        concealment = [ir for ir in integrity_results
                       if any(iss["维度"] == "隐瞒信号" for iss in ir.get("issues", []))]
        if concealment:
            _docx_add_para(doc, "隐瞒信号分类统计：", indent=True)
            conceal_types = defaultdict(int)
            for ir in concealment:
                for iss in ir.get("issues", []):
                    if iss["维度"] == "隐瞒信号":
                        # 简单归类
                        if "笼统" in iss["问题"]:
                            conceal_types["节点名过于笼统"] += 1
                        elif "合并" in iss["问题"]:
                            conceal_types["关键环节不当合并"] += 1
                        elif "未标注岗位" in iss["问题"]:
                            conceal_types["关键节点缺岗位"] += 1
                        elif "简短" in iss["问题"]:
                            conceal_types["流程描述过短"] += 1
                        else:
                            conceal_types[iss["问题"][:20]] += 1
            for k, v in sorted(conceal_types.items(), key=lambda x: -x[1]):
                _docx_add_para(doc, f"• {k}：{v} 处", indent=True)

    # ── 五、配对完整性 ──
    _docx_add_heading(doc, "五、统筹-实施配对完整性", level=1)
    if orphans["orphan_ought"]:
        _docx_add_para(doc, f"有统筹无实施（统筹孤儿业务）{len(orphans['orphan_ought'])} 条：", indent=True)
        for domain, unit, biz in orphans["orphan_ought"]:
            _docx_add_para(doc, f"• [{domain}] {unit}：{biz}", indent=True)
    else:
        _docx_add_para(doc, "未发现统筹孤儿业务。", indent=True)

    if orphans["orphan_is"]:
        _docx_add_para(doc, f"有实施无统筹（实施孤儿业务）{len(orphans['orphan_is'])} 条：", indent=True)
        for domain, unit, biz in orphans["orphan_is"]:
            _docx_add_para(doc, f"• [{domain}] {unit}：{biz}", indent=True)
    else:
        _docx_add_para(doc, "未发现实施孤儿业务。", indent=True)

    if not orphans["orphan_ought"] and not orphans["orphan_is"]:
        _docx_add_para(doc, "统筹与实施配对基本完整，未发现明显缺失。", indent=True)

    # ── 六、排查领域冲突 ──
    if domain_conflicts:
        _docx_add_heading(doc, "六、排查领域与流程推断冲突", level=1)
        _docx_add_para(doc, f"以下 {len(domain_conflicts)} 条记录中，单位填写的排查领域与系统从流程文本关键字推断的领域不一致，建议人工核实确认：", indent=True)
        for s in domain_conflicts:
            inferred = _infer_business_from_process(s.raw_process) if s.raw_process else "无法推断"
            _docx_add_para(doc, f"• {s.unit_name}：排查领域='{s.inspection_domain}' ↔ 流程推断='{inferred}'", indent=True)

    # ── 七、差异统计 ──
    _docx_add_heading(doc, "七、应然-实然差异统计", level=1)
    total_red = sum(1 for diffs in all_diffs.values() for d in diffs if d.diff_type == "节点消失")
    total_orange = sum(1 for diffs in all_diffs.values() for d in diffs if d.diff_type == "岗位偏移")
    total_yellow = sum(1 for diffs in all_diffs.values() for d in diffs if d.diff_type == "节点新增")
    _docx_add_para(doc, f"统筹流程（应然）与实施流程（实然）逐节点对比结果：", indent=True)
    _docx_add_para(doc, f"• 节点消失（红色，高风险）：{total_red} 处 — 应然节点在实施流程中未找到对应", indent=True)
    _docx_add_para(doc, f"• 岗位偏移（橙色，中高风险）：{total_orange} 处 — 同一节点的执行岗位发生偏移", indent=True)
    _docx_add_para(doc, f"• 节点新增（黄色，中风险）：{total_yellow} 处 — 实施流程比应然多出的节点", indent=True)

    if impl_comparisons:
        high_impact = [c for c in impl_comparisons if "高风险" in c.get("impact", "")]
        _docx_add_para(doc, f"同业务不同实施单位之间共发现 {len(impl_comparisons)} 处差异，其中高风险差异 {len(high_impact)} 处。详情见 Excel 报告。", indent=True)

    # ── 八、下一步工作建议 ──
    _docx_add_heading(doc, "八、下一步工作建议", level=1)
    suggestions = []
    suggestions.append("1. 优先对 A/B 级高风险业务线安排现场核验，重点核查红色差异节点。")
    if template_suspects:
        suggestions.append(f"2. 对 {len(template_suspects)} 条疑似抄模板或敷衍填写的记录，退回相关单位重新填报，要求按实际业务流程如实填写。")
    if orphans["orphan_ought"] or orphans["orphan_is"]:
        suggestions.append("3. 对配对缺失的统筹/实施业务，核实是否存在遗漏，补充对应的自查材料。")
    if domain_conflicts:
        suggestions.append("4. 对排查领域选择与流程描述不一致的记录，请相关单位确认正确领域后修正。")
    if cat_issues:
        suggestions.append(f"5. 对 {len(cat_issues)} 处排查分类勾选与流程节点不匹配的情况，核实是勾选错误还是流程描述遗漏。")
    suggestions.append("6. 现场核验时携带《现场核验手册》，重点追问差异节点的实际情况及佐证材料。")
    suggestions.append("7. 核验完成后汇总发现的问题，形成整改台账，明确责任人和整改时限。")

    for sug in suggestions:
        _docx_add_para(doc, sug, indent=True)

    # ── 页脚 ──
    doc.add_paragraph()
    _docx_add_para(doc, "本报告由 lzpc 分析工具自动生成，基于各单位提交的自查材料。如需调整分析参数，请编辑 analyzer.py 中的配置项。", indent=False)
    for run in doc.paragraphs[-1].runs:
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)

    doc.save(output_path)
    print(f"[报告] 文字总结报告已保存至: {output_path}")


# ============================================================================
# Docx 现场核验手册
# ============================================================================

def generate_docx_handbook(
    output_path: str,
    pairs: dict,
    business_risks: list[BusinessRisk],
    impl_comparisons: list[dict],
    all_diffs: dict,
):
    """生成 docx 格式的现场核验手册"""
    if not DOCX_AVAILABLE:
        print("[警告] python-docx 未安装，跳过 docx 核验手册生成。请执行: pip install python-docx")
        return

    doc = Document()
    style = doc.styles["Normal"]
    font = style.font
    font.name = "微软雅黑"
    font.size = Pt(11)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

    # ── 封面 ──
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(60)
    run = title.add_run("廉洁风险排查——现场核验手册")
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    run.font.size = Pt(22)
    run.bold = True

    info = doc.add_paragraph()
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = info.add_run(f"生成日期：{datetime.now().strftime('%Y年%m月%d日')}　　|　　本手册根据各单位自查材料生成，供重点排查阶段现场核验使用")
    run.font.name = "微软雅黑"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_paragraph()

    # 按风险等级排序
    business_risks.sort(key=lambda x: x.total_score, reverse=True)

    for br in business_risks:
        data = pairs.get(br.business_name, {"ought": [], "is": []})

        _docx_add_heading(doc, br.business_name, level=1)

        # 风险等级标签
        level_text = f"风险等级：{br.level}（{br.level_desc}）　　加权总分：{br.total_score}/5.0"
        _docx_add_para(doc, level_text, bold=True)

        # 风险维度概览表格
        _docx_add_para(doc, "风险维度得分：", bold=True)
        dim_headers = ["维度", "得分（1-5）", "说明"]
        dim_rows = [
            ["节点完整度", f"{br.node_completeness:.1f}", "与最小节点清单对照，缺失越多得分越高"],
            ["岗位清晰度", f"{br.position_clarity:.1f}", "岗位标注越模糊得分越高"],
            ["权力集中度", f"{br.power_concentration:.1f}", "同一岗位跨越权力类型越多得分越高"],
            ["应然-实然匹配度", f"{br.ought_is_match:.1f}", "统筹与实施流程差异越大得分越高"],
            ["自查风险自觉度", f"{br.self_awareness:.1f}", "自查越空洞得分越高"],
        ]
        _docx_add_table(doc, dim_headers, dim_rows, col_widths=[4, 2.5, 8])

        # 应然-实然差异清单
        if data["ought"] and data["is"]:
            _docx_add_para(doc, "应然-实然差异清单：", bold=True)
            for ought in data["ought"]:
                for is_sub in data["is"]:
                    diffs = compare_nodes(ought.nodes, is_sub.nodes)
                    red_diffs = [d for d in diffs if d.diff_type in ("节点消失", "岗位偏移", "节点新增")]
                    if red_diffs:
                        _docx_add_para(doc, f"统筹单位：{ought.unit_name}　→　实施单位：{is_sub.unit_name}", bold=True)
                        diff_headers = ["#", "应然节点", "实然节点", "差异类型", "差异说明"]
                        diff_rows = [
                            [str(d.seq), d.ought_node or "-", d.is_node or "-", d.diff_type, d.description]
                            for d in red_diffs
                        ]
                        _docx_add_table(doc, diff_headers, diff_rows, col_widths=[1, 4, 4, 2, 5])

        # 现场核查建议
        if br.level in ("A", "B"):
            suggestions = generate_verification_suggestions(
                br.business_name, br.level,
                all_diffs.get(br.business_name, []),
            )

            if suggestions.get("documents"):
                _docx_add_para(doc, "建议查阅资料：", bold=True)
                for d_item in suggestions["documents"]:
                    _docx_add_para(doc, f"☐ {d_item}", indent=True)

            if suggestions.get("interview_directions"):
                _docx_add_para(doc, "谈话方向：", bold=True)
                for d_item in suggestions["interview_directions"]:
                    _docx_add_para(doc, f"• {d_item}", indent=True)

            if suggestions.get("general_questions"):
                _docx_add_para(doc, "重点谈话问题：", bold=True)
                for q in suggestions["general_questions"]:
                    _docx_add_para(doc, f"• {q}", indent=True)

            if suggestions.get("targeted_questions"):
                _docx_add_para(doc, "针对性追问（根据自查材料发现）：", bold=True)
                for q in suggestions["targeted_questions"]:
                    _docx_add_para(doc, f"• {q}", indent=True)

        doc.add_page_break()

    # ── 附录：同业务实施单位差异 ──
    if impl_comparisons:
        _docx_add_heading(doc, "附录：同业务不同实施单位差异对比", level=1)
        _docx_add_para(doc, "以下列出同一业务在不同实施单位之间存在显著差异的情况，现场核验时应重点关注。", indent=True)
        by_business = defaultdict(list)
        for comp in impl_comparisons:
            by_business[comp["business"]].append(comp)
        for biz_name, comps in by_business.items():
            _docx_add_heading(doc, biz_name, level=2)
            headers = ["单位A", "单位B", "差异节点", "差异类型", "预计影响"]
            rows_data = [
                [c["unit_a"], c["unit_b"], c["diff_node"], c["diff_type"], c["impact"]]
                for c in comps[:20]  # 最多 20 行，避免过长
            ]
            _docx_add_table(doc, headers, rows_data, col_widths=[3, 3, 3, 2, 5])

    # ── 页脚 ──
    doc.add_paragraph()
    _docx_add_para(doc, "本手册由 lzpc 分析工具自动生成，基于各单位提交的自查材料。如需调整分析参数，请编辑 analyzer.py 中的配置项。", indent=False)
    for run in doc.paragraphs[-1].runs:
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)

    doc.save(output_path)
    print(f"[报告] 现场核验手册已保存至: {output_path}")


# ============================================================================
# 主流程
# ============================================================================

def main():
    if len(sys.argv) < 2:
        print("用法: python analyzer.py <输入Excel路径> [输出目录]")
        print("示例: python analyzer.py 各单位自查汇总.xlsx")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    if not input_path.exists():
        print(f"错误: 找不到文件 '{input_path}'")
        sys.exit(1)

    # 输出目录
    if len(sys.argv) >= 3:
        output_dir = Path(sys.argv[2])
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = input_path.parent / f"lzpc_output_{timestamp}"
    output_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print("  廉洁风险排查流程分析工具 (lzpc)")
    print("=" * 60)
    print(f"\n输入文件: {input_path}")
    print(f"输出目录: {output_dir}\n")

    # Step 1: 解析
    print("[1/6] 解析 Excel 表格...")
    submissions = parse_excel(str(input_path))

    # Step 2: 加载别名并匹配
    print("\n[2/6] 加载业务名称别名并匹配...")
    aliases = load_aliases()
    alias_lookup = build_alias_lookup(aliases)
    pairs = pair_ought_and_is(submissions, alias_lookup)

    matched_count = sum(1 for s in submissions if s.business_name_canonical)
    unmatched = [s for s in submissions if not s.business_name_canonical]
    print(f"[匹配] 成功匹配: {matched_count}/{len(submissions)} 条记录")
    if unmatched:
        print(f"[匹配] 未匹配 ({len(unmatched)} 条):")
        for s in unmatched:
            print(f"  - {s.unit_name}: '{s.business_name_raw}'")
        print("  提示: 请将以上业务名称的别名添加到 aliases.json 后重新运行。")

    # Step 2.5: 孤儿业务检测
    print("\n[2.5/6] 检测孤儿业务（统筹-实施配对缺失）...")
    orphans = detect_orphan_businesses(submissions)

    # Step 3: 风险评分
    print("\n[3/6] 计算风险评分...")
    business_risks = []
    all_diffs = {}

    for biz_name, data in pairs.items():
        if not data["ought"] and not data["is"]:
            continue
        risk = score_business_risk(biz_name, data["ought"], data["is"])
        business_risks.append(risk)

        # 收集所有差异备后续使用
        diffs_for_biz = []
        checklist = MIN_NODE_CHECKLIST.get(biz_name, [])
        if data["ought"]:
            for ought in data["ought"]:
                for is_sub in data["is"]:
                    d = compare_nodes(ought.nodes, is_sub.nodes)
                    diffs_for_biz.extend(d)
        elif data["is"]:
            for is_sub in data["is"]:
                d = compare_against_checklist(is_sub.nodes, checklist)
                diffs_for_biz.extend(d)
        all_diffs[biz_name] = diffs_for_biz

    # Step 4: 单位质量评分 + 排查分类交叉校验
    print("\n[4/6] 评估各单位自查质量及排查分类一致性...")
    unit_qualities = []
    cat_issues = []  # 排查分类与流程不一致的记录
    for s in submissions:
        checklist = MIN_NODE_CHECKLIST.get(s.business_name_canonical, [])
        q = score_unit_quality(s, checklist)
        unit_qualities.append(q)

        # 排查分类交叉校验
        cat_result = validate_inspection_categories(s)
        if cat_result["missing"]:
            for cat, expected in cat_result["missing"]:
                cat_issues.append({
                    "unit": s.unit_name,
                    "domain": s.inspection_domain or s.business_name_canonical,
                    "category": cat,
                    "expected_nodes": expected,
                })

    # Step 5: 同业务实施单位差异
    print("\n[5/6] 比较同业务不同实施单位差异...")
    impl_comparisons = []
    for biz_name, data in pairs.items():
        if data["is"] and len(data["is"]) >= 2:
            comps = compare_implementation_units(biz_name, data["is"])
            impl_comparisons.extend(comps)

    # Step 5.5: 流程完整性分析
    print("\n[5.5/6] 独立审查每条流程的完整性...")
    integrity_results = []
    for s in submissions:
        ir = analyze_process_integrity(s)
        integrity_results.append(ir)
    defect_count = sum(1 for ir in integrity_results if ir["grade"] in ("存在明显缺陷", "严重不完整"))
    severe_count = sum(1 for ir in integrity_results if ir["grade"] == "严重不完整")
    print(f"[完整性] 共审查 {len(integrity_results)} 条流程")
    print(f"[完整性] 存在明显缺陷: {defect_count} 条, 严重不完整: {severe_count} 条")

    # 提前收集供 docx 报告使用的变量
    domain_conflicts = [s for s in submissions if s.domain_confidence == "待确认"]

    # Step 6: 生成报告
    print("\n[6/6] 生成分析报告...")

    # Excel 数据报告
    excel_path = output_dir / "分析报告.xlsx"
    generate_excel_report(
        str(excel_path), pairs, business_risks, unit_qualities,
        impl_comparisons, all_diffs,
        integrity_results=integrity_results,
    )

    # Docx 文字总结报告
    docx_summary_path = output_dir / "分析总结报告.docx"
    generate_docx_summary(
        str(docx_summary_path), submissions, pairs, business_risks,
        unit_qualities, integrity_results, orphans, all_diffs,
        cat_issues, domain_conflicts, impl_comparisons,
    )

    # Docx 现场核验手册
    docx_handbook_path = output_dir / "现场核验手册.docx"
    generate_docx_handbook(
        str(docx_handbook_path), pairs, business_risks,
        impl_comparisons, all_diffs,
    )

    # ---- 打印摘要 ----
    print("\n" + "=" * 60)
    print("  分析摘要")
    print("=" * 60)

    # 高风险业务
    high_risk = [br for br in business_risks if br.level in ("A", "B")]
    high_risk.sort(key=lambda x: x.total_score, reverse=True)
    if high_risk:
        print(f"\n[!!] 高风险业务线 ({len(high_risk)} 条):")
        for br in high_risk:
            print(f"  [{br.level}] {br.business_name} — 总分 {br.total_score}")

    # 自查质量最差
    worst = sorted(unit_qualities, key=lambda x: x["total"])[:5]
    if worst:
        print(f"\n[!] 自查质量最差的 5 个单位/业务:")
        for q in worst:
            print(f"  [{q['grade']}] {q['unit']}（{q['business']}）— {q['total']}分")

    # 差异统计
    total_red = sum(
        1 for diffs in all_diffs.values()
        for d in diffs if d.diff_type == "节点消失"
    )
    total_orange = sum(
        1 for diffs in all_diffs.values()
        for d in diffs if d.diff_type == "岗位偏移"
    )
    total_new = sum(
        1 for diffs in all_diffs.values()
        for d in diffs if d.diff_type == "节点新增"
    )
    print(f"\n[*] 差异统计:")
    print(f"  节点消失（红）: {total_red} 处")
    print(f"  岗位偏移（橙）: {total_orange} 处")
    print(f"  节点新增（黄）: {total_new} 处")

    # 实施单位差异
    if impl_comparisons:
        print(f"\n[*] 同业务实施单位差异: {len(impl_comparisons)} 处")
        high_impact = [c for c in impl_comparisons if "高风险" in c.get("impact", "")]
        if high_impact:
            print(f"  其中高风险差异: {len(high_impact)} 处")

    # 未匹配的业务名称
    if unmatched:
        print(f"\n[!] 未匹配的业务名称 ({len(unmatched)} 条) — 请在 aliases.json 中补充别名:")
        for s in unmatched:
            print(f"  - {s.unit_name}: '{s.business_name_raw}'")

    # 孤儿业务
    if orphans["orphan_ought"]:
        print(f"\n[!] 统筹孤儿业务（有统筹无实施，{len(orphans['orphan_ought'])} 条）:")
        for domain, unit, biz in orphans["orphan_ought"]:
            print(f"  - [{domain}] {unit}: '{biz}'")

    if orphans["orphan_is"]:
        print(f"\n[!] 实施孤儿业务（有实施无统筹，{len(orphans['orphan_is'])} 条）:")
        for domain, unit, biz in orphans["orphan_is"]:
            print(f"  - [{domain}] {unit}: '{biz}'")

    # 排查分类与流程不一致
    if cat_issues:
        print(f"\n[!] 排查分类与流程不匹配 ({len(cat_issues)} 处):")
        for ci in cat_issues[:10]:
            print(f"  - {ci['unit']}: 勾选了'{ci['category']}'但流程中未找到对应节点 {ci['expected_nodes']}")

    # 疑似抄模板+敷衍单位
    template_suspects = [q for q in unit_qualities if q.get("flags")]
    if template_suspects:
        print(f"\n[!] 疑似抄模板/敷衍填写 ({len(template_suspects)} 条):")
        for q in sorted(template_suspects, key=lambda x: x["total"])[:10]:
            print(f"  [{q['grade']}] {q['unit']}（{q['business']}）— {q['total']}分")
            for flag in q["flags"][:3]:
                print(f"      {flag}")

    # 排查领域与流程推断冲突的记录
    if domain_conflicts:
        print(f"\n[!] 排查领域-流程推断冲突 ({len(domain_conflicts)} 条，建议人工确认):")
        for s in domain_conflicts:
            inferred = _infer_business_from_process(s.raw_process) if s.raw_process else "无法推断"
            print(f"  - {s.unit_name}: 排查领域='{s.inspection_domain}' vs 流程推断='{inferred}'")

    # 流程完整性分析结果
    if integrity_results:
        worst_integrity = sorted(integrity_results, key=lambda x: x["score"])[:5]
        concealment = [ir for ir in integrity_results
                       if any(iss["维度"] == "隐瞒信号" for iss in ir.get("issues", []))]
        print(f"\n[!] 流程完整性最差的 5 条:")
        for ir in worst_integrity:
            issue_summary = "; ".join(iss["问题"][:30] for iss in ir.get("issues", [])[:2])
            print(f"  [{ir['grade']}] {ir['unit']}（{ir['domain']}）— {ir['score']}分 | {issue_summary}")
        if concealment:
            print(f"\n[!] 存在隐瞒信号的流程 ({len(concealment)} 条):")
            for ir in sorted(concealment, key=lambda x: x["score"])[:5]:
                conceal_issues = [iss["问题"] for iss in ir.get("issues", []) if iss["维度"] == "隐瞒信号"]
                print(f"  [{ir['grade']}] {ir['unit']}（{ir['domain']}）:")
                for ci in conceal_issues[:3]:
                    print(f"      {ci}")

    print(f"\n详细报告:")
    print(f"  [Excel] {excel_path}")
    print(f"  [总结] {docx_summary_path}")
    print(f"  [手册] {docx_handbook_path}")
    print("\n分析完成。")


if __name__ == "__main__":
    main()
