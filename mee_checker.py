#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MEE 多国语数字校对工具 (P0)

用法:
  python mee_checker.py --base <英文指示稿.pdf> --dir <多国语PDF文件夹> [--out <输出目录>]

流程:
  1. 从英文指示稿提取红色检查点(行聚类+间隔聚合)
  2. 从各语言译文提取红色检查点
  3. 页内匹配(顺序对齐 / y行位置序列对齐) + 数值归一化比较
  4. 输出 Excel 报告(汇总/明细/矩阵/说明) + 问题项双方截图
"""
import argparse
import base64
import bisect
import os
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from statistics import median

import numpy as np
import fitz  # PyMuPDF
from scipy.optimize import linear_sum_assignment
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------- 可调参数 ----------------
RED_R_MIN = 150        # 标准红判定: R 最小值
RED_GB_MAX = 110       # 标准红判定: G/B 最大值
RED_DOMINANCE = 60     # 标准红判定: R 需比 max(G,B) 高出的量(排除棕/橙色)
LOOSE_R_MIN = 200      # 浅橙红扩展(如零件号标签上的 #F69679): R 最小值
LOOSE_RG_DIFF = 80     # 浅橙红扩展: R-G 最小差
LOOSE_GB_MAX = 60      # 浅橙红扩展: G-B 最大差(排除偏黄橙)
LINE_TOL = 3.0         # 行聚类: y 中心差容差 (pt). 收紧防跨行误聚合(小字号文档行距仅4-5pt)
JOIN_GAP = 2.0         # span 拼接: 间隔 < 此值直接连写, 否则加空格 (pt)
GAP_TOL = 15.0         # 行内聚合: 相邻 span 间隔 <= 此值视为同一检查点 (pt)
CTX_LEN = 12           # 骨架上下文: 红字紧邻文本截取字符数
ANCHOR_OFF_TOL = 60.0  # 锚点(值相同候选对)y偏移绝对值上限 (pt), 仅用于估计页级偏移
PRUNE_DY = 80.0        # 值不同的候选边: |Δy偏移| 剪枝 (pt)
PRUNE_DX = 160.0       # 值不同的候选边: |Δx偏移| 剪枝 (pt)
PRUNE_DY_SAME = 120.0  # 值相同(强信号)的候选边: 宽松剪枝
PRUNE_DX_SAME = 400.0
PRUNE_DY_LOOSE = 120.0 # 无锚点页的宽松剪枝
PRUNE_DX_LOOSE = 240.0
REJECT_COST = 70.0     # 分配代价超过此值不成立 -> 未匹配
CONF_DY_HIGH = 15.0    # 值相同且 Δy<=此值 -> high
CONF_DY_MED = 12.0     # 值不同但 Δy<=此值 -> medium(真实差异典型形态)
W_X = 0.5              # x 偏移代价权重
REWARD_SAME = -30.0    # 值相同奖励(负代价)
PENALTY_DIFF = 35.0    # 值不同惩罚(允许真实差异, 但需位置强吻合)
PENALTY_TOK = 35.0     # 数字个数不同惩罚
REWARD_CTX = 8.0       # 骨架上下文吻合奖励(每侧)
SP_WIN = 95.0          # 二级匹配: 插值期望位置窗口 (pt)
SP_DEV_TOL = 75.0      # 二级匹配: 实际位置与期望的最大偏差 (pt)
SNAP_PAD = 55          # 截图外扩边距 (pt)
SNAP_DPI = 150         # 截图分辨率

TOKEN_RE = re.compile(r'\d+(?:[.,]\d+)*')

# ---------------- 数据结构 ----------------
@dataclass
class Item:
    page: int        # 0-based
    text: str        # 聚合后文本
    bbox: tuple      # (x0, y0, x1, y1)
    yc: float        # 行中心 y
    n_spans: int
    left_ctx: str = ''   # 骨架上下文: 同行紧邻左侧非红文本尾部
    right_ctx: str = ''  # 骨架上下文: 同行紧邻右侧非红文本头部

    @property
    def xc(self) -> float:
        return (self.bbox[0] + self.bbox[2]) / 2

@dataclass
class Pair:
    cp: str                 # 检查点编号(以英文指示稿为准), 译文多出为空
    page: int               # 1-based
    en: Item | None
    xx: Item | None
    y_off: float | None
    conf: str               # high / medium / low / -
    status: str             # 一致/不一致/格式差异/待人工/译文未匹配/译文多出
    note: str
    expect_y: float | None = None   # 未匹配项在译文坐标系的插值期望 y(供截图)
    expect_dx: float | None = None  # 页级 x 偏移(供截图定位)

# ---------------- 提取 ----------------
def is_red_core(color: int) -> bool:
    """标准红(脚本标注主色)"""
    r = (color >> 16) & 255
    g = (color >> 8) & 255
    b = color & 255
    return r >= RED_R_MIN and g <= RED_GB_MAX and b <= RED_GB_MAX \
        and (r - max(g, b)) >= RED_DOMINANCE


def is_red_loose(color: int) -> bool:
    """浅橙红扩展: 覆盖彩色标签上的红系变体(如 #F69679), 排除偏黄橙"""
    r = (color >> 16) & 255
    g = (color >> 8) & 255
    b = color & 255
    return r >= LOOSE_R_MIN and (r - g) >= LOOSE_RG_DIFF \
        and 0 <= (g - b) <= LOOSE_GB_MAX


def is_red(color: int) -> bool:
    return is_red_core(color) or is_red_loose(color)


def _make_ctx(others: list, gx0: float, gx1: float):
    """从同行非红 span 中提取检查点紧邻骨架上下文"""
    left = right = ''
    lts = [o[2].strip() for o in others if o[1] <= gx0]
    rts = [o[2].strip() for o in others if o[0] >= gx1]
    if lts and lts[-1]:
        left = lts[-1][-CTX_LEN:]
    if rts and rts[0]:
        right = rts[0][:CTX_LEN]
    return left, right


def extract_items(doc: fitz.Document, color_counter: Counter | None = None) -> list[Item]:
    """提取全部红色检查点: 红色 span -> 注脚编号剔除 -> 行聚类 -> 行内按间隔聚合

    注脚剔除规则: 红色 span 为纯数字, 且同行左侧紧邻 span 文本以 '*' 结尾 ->
    判定为脚注索引(如 *5, *6, *7), 不进入检查点(非校对数据)。"""
    items = []
    for pno in range(doc.page_count):
        page = doc[pno]
        d = page.get_text("dict")
        raw = []   # [(span, is_red)]
        for blk in d.get("blocks", []):
            if blk.get("type") != 0:
                continue
            for line in blk.get("lines", []):
                for sp in line.get("spans", []):
                    t = sp["text"].strip()
                    if not t:
                        continue
                    raw.append((sp, is_red(sp["color"])))
        if not any(red for _, red in raw):
            continue
        # 全 span 行聚类(含非红, 用于注脚判断与骨架上下文)
        # 垂直同列(相邻行同 x0)的 span 即使 y 差 <= LINE_TOL 也必须拆开, 否则误聚合
        raw.sort(key=lambda s: ((s[0]["bbox"][1] + s[0]["bbox"][3]) / 2, s[0]["bbox"][0]))
        rows: list[list] = []
        for sp, red in raw:
            yc = (sp["bbox"][1] + sp["bbox"][3]) / 2
            if rows and abs(yc - sum((s[0]["bbox"][1] + s[0]["bbox"][3]) / 2
                                     for s in rows[-1]) / len(rows[-1])) <= LINE_TOL:
                # 垂直同列检查: 与同行已有 span x 几乎相同且 y 错开 -> 垂直排列, 拆行
                vert = any(abs(s[0]["bbox"][0] - sp["bbox"][0]) < 2.5
                           and abs((s[0]["bbox"][1] + s[0]["bbox"][3]) / 2 - yc) > 1.5
                           for s in rows[-1])
                if vert:
                    rows.append([(sp, red)])
                else:
                    rows[-1].append((sp, red))
            else:
                rows.append([(sp, red)])
        # 收集红色 span(剔除注脚)
        spans = []
        for row in rows:
            row.sort(key=lambda s: s[0]["bbox"][0])
            for i, (sp, red) in enumerate(row):
                if not red:
                    continue
                t = sp["text"].strip()
                prev_txt = row[i - 1][0]["text"].strip() if i > 0 else ''
                is_footnote = (prev_txt.endswith('*') and re.fullmatch(r'\d{1,3}', t) is not None)
                if is_footnote:
                    continue
                spans.append(sp)
                if color_counter is not None:
                    color_counter[f'#{sp["color"]:06x}'] += 1
        # 行聚类(红色): 按 y 中心排序, 相邻 y 中心差 <= LINE_TOL 归入同一行(垂直同列拆行)
        spans.sort(key=lambda s: ((s["bbox"][1] + s["bbox"][3]) / 2, s["bbox"][0]))
        rrows: list[list[dict]] = []
        for sp in spans:
            yc = (sp["bbox"][1] + sp["bbox"][3]) / 2
            if rrows and abs(yc - sum(
                    (s["bbox"][1] + s["bbox"][3]) / 2 for s in rrows[-1]) / len(rrows[-1])) <= LINE_TOL:
                vert = any(abs(s["bbox"][0] - sp["bbox"][0]) < 2.5
                           and abs((s["bbox"][1] + s["bbox"][3]) / 2 - yc) > 1.5
                           for s in rrows[-1])
                if vert:
                    rrows.append([sp])
                else:
                    rrows[-1].append(sp)
            else:
                rrows.append([sp])
        # 行内按 x 排序, 按间隔聚合为检查点
        others = [[o["bbox"][0], o["bbox"][2], o["text"].strip(),
                   (o["bbox"][1] + o["bbox"][3]) / 2]
                  for o, red in raw if not red]
        for row in rrows:
            row.sort(key=lambda s: s["bbox"][0])
            ry = sum((s["bbox"][1] + s["bbox"][3]) / 2 for s in row) / len(row)
            row_others = sorted([o for o in others if abs(o[3] - ry) <= LINE_TOL + 2],
                                key=lambda o: o[0])
            groups: list[list[dict]] = [[row[0]]]
            for sp in row[1:]:
                prev = groups[-1][-1]
                gap = sp["bbox"][0] - prev["bbox"][2]
                if sp["bbox"][0] < prev["bbox"][2] - 0.3:
                    # x 区间重叠 -> 不是同一行(叠放的两行文本), 必须拆开, 否则误连写(如 60/150)
                    groups.append([sp])
                elif gap <= GAP_TOL:
                    groups[-1].append(sp)
                else:
                    groups.append([sp])
            for g in groups:
                text = ""
                prev = None
                for sp in g:
                    if prev is not None:
                        gap = sp["bbox"][0] - prev["bbox"][2]
                        # 两红色 span 之间的点号/逗号(黑色): 小数/千分位, 直接插入不分隔
                        dot = ''
                        for o in row_others:
                            if prev["bbox"][2] - 0.5 <= o[0] <= sp["bbox"][0] + 0.5 \
                                    and o[1] <= sp["bbox"][0] + 0.5 \
                                    and o[2] in ('.', ',', '．', '，', '·'):
                                dot = o[2]
                                break
                        if dot:
                            text += dot
                        else:
                            text += "" if gap < JOIN_GAP else " "
                    text += sp["text"].strip()
                    prev = sp
                x0 = min(s["bbox"][0] for s in g)
                y0 = min(s["bbox"][1] for s in g)
                x1 = max(s["bbox"][2] for s in g)
                y1 = max(s["bbox"][3] for s in g)
                left, right = _make_ctx(row_others, x0, x1)
                items.append(Item(page=pno, text=text, bbox=(x0, y0, x1, y1),
                                  yc=(y0 + y1) / 2, n_spans=len(g),
                                  left_ctx=left, right_ctx=right))
    return items

# ---------------- 数值归一化与比较 ----------------
def norm_num(tok: str):
    """归一化数字 token -> float; 无法解析返回 None"""
    s = tok.replace('\u00a0', '').replace(' ', '')
    if not re.fullmatch(r'\d+(?:[.,]\d+)*', s):
        return None
    if '.' in s and ',' in s:
        if s.rfind(',') > s.rfind('.'):      # 1.234,5 欧式
            s = s.replace('.', '').replace(',', '.')
        else:                                 # 1,234.5 英式
            s = s.replace(',', '')
    elif ',' in s:
        parts = s.split(',')
        if parts[0].isdigit() and int(parts[0]) == 0:
            s = s.replace(',', '.')                       # 0,988 => 0.988 (欧式小数, 整数为0必为小数点)
        elif len(parts) == 2 and len(parts[1]) != 3:      # 2,5 -> 2.5
            s = s.replace(',', '.')
        else:                                             # 1,234 / 1,234,567
            s = s.replace(',', '')
    # 只含 '.' 的 1.234 存在欧式千分位歧义, P0 按小数处理, 双方原样不同时会落入格式差异/不一致交人工
    try:
        return float(s)
    except ValueError:
        return None


def compare(en_text: str, xx_text: str):
    """返回 (status, note)"""
    en_toks = TOKEN_RE.findall(en_text)
    xx_toks = TOKEN_RE.findall(xx_text)
    if not en_toks and not xx_toks:
        # 无数字红字: 字符串比较
        if en_text == xx_text:
            return '一致', '文本相同'
        return '待人工', f'文本不同: EN"{en_text}" vs "{xx_text}"'
    if len(en_toks) != len(xx_toks):
        # 纯数字 token 拼接相同 -> 分组写法差异, 判一致 (如 25354250 vs 25 35 42 50)
        # 含 ./, 分隔符的不走此捷径 (如 '2,5' vs '2 5' 语义不同), 保守转人工
        if all(t.isdigit() for t in en_toks) and all(t.isdigit() for t in xx_toks) \
                and ''.join(en_toks) == ''.join(xx_toks):
            return '一致', f'分组写法不同, 数字串相同: "{en_text}" vs "{xx_text}"'
        return '待人工', f'数字个数不同: EN{en_toks} vs {xx_toks}'
    fmt = []
    for a, b in zip(en_toks, xx_toks):
        na, nb = norm_num(a), norm_num(b)
        if na is None or nb is None:
            if a != b:
                return '待人工', f'无法解析: {a} vs {b}'
            continue
        if abs(na - nb) > 1e-9:
            return '不一致', f'{a} → {b}'
        if a != b:
            fmt.append((a, b))
    if fmt:
        # 仅点/逗号互换(如 0.988 vs 0,988, 13.7 vs 13,7) -> 欧式小数写法, 判一致并提示人工
        sig = lambda t: re.sub(r'[.,]', '', t)
        if all(sig(a) == sig(b) for a, b in fmt) and all('.' in (a + b) or ',' in (a + b) for a, b in fmt):
            return '一致', '疑似小数写法差异(点/逗号), 人工最好过一下; ' + ', '.join(f'{a}→{b}' for a, b in fmt)
        return '格式差异', '数值相同格式不同: ' + ', '.join(f'{a}→{b}' for a, b in fmt)
    return '一致', ''

# ---------------- 匹配 ----------------
def _expected_y(e: Item, mm: list[tuple[Item, Item]], ens_yc: list[float]):
    """已配对邻居推算 e 在译文坐标系的期望 y。
    双侧邻居: 线性插值; 首/尾单侧: 沿邻居偏移外推; 完全无邻居: None。"""
    if not mm:
        return None
    k = bisect.bisect_left(ens_yc, e.yc)
    if k == 0 or k >= len(mm):
        if k == 0:
            ne, nx = mm[0]
            return nx.yc - (ne.yc - e.yc)      # 首部: 沿下邻偏移外推
        pe, px = mm[-1]
        return px.yc + (e.yc - pe.yc)          # 尾部: 沿上邻偏移外推
    pe, px = mm[k - 1]
    ne, nx = mm[k]
    if ne.yc - pe.yc < 1e-6:
        return None
    t = (e.yc - pe.yc) / (ne.yc - pe.yc)
    return px.yc + t * (nx.yc - px.yc)


def value_same(a: str, b: str) -> bool:
    """匹配层数值等效: 归一化后相等即视为同一数字(含'一致'与'格式差异').
    注意: 格式差异(如 0.988 vs 0,988)在匹配时必须享受值相同奖励, 否则表格内
    欧式小数候选会被当成'值不同'惩罚, 导致匈牙利错位误报不一致。"""
    try:
        st, _ = compare(a, b)
    except Exception:
        return False
    return st in ('一致', '格式差异')


def token_count(a: str) -> int:
    return len(TOKEN_RE.findall(a))


def match_page(en_items: list[Item], xx_items: list[Item], page1: int) -> list[Pair]:
    """单页匹配: 页级偏移估计 -> 统一匈牙利全局最优分配 -> 邻域插值二级匹配

    设计要点:
    - 所有检查点统一参与全局分配, 不硬锁定锚点; "缺失"以 COST_BIG 表达,
      同值多点时全局最优会自动把候选让给位置最吻合的一方, 避免贪心抢占张冠李戴;
    - 值不同不阻断匹配(否则真实差异会漏配), 但要求位置强吻合;
    - 一级拒绝后进入二级匹配: 用已配对的上下邻居插值期望位置,
      唯一值相同候选 + 双侧邻居支撑 + 偏差达标三条件全过才确认, 否则保持未匹配(转人工)。
    """
    en_items = sorted(en_items, key=lambda i: (i.yc, i.bbox[0]))
    xx_items = sorted(xx_items, key=lambda i: (i.yc, i.bbox[0]))
    n, m = len(en_items), len(xx_items)

    if n == 0:
        return [Pair(cp='', page=page1, en=None, xx=x, y_off=None, conf='-',
                     status='译文多出', note='译文中多出的红字(英文侧无对应)')
                for x in xx_items]
    if m == 0:
        return [Pair(cp='', page=page1, en=e, xx=None, y_off=None, conf='-',
                     status='译文未匹配', note='译文中无红字', expect_dx=0.0)
                for e in en_items]

    # 1) 锚点统计(仅用于估计页级偏移, 不锁定配对)
    used: set[int] = set()
    ay: list[float] = []
    ax: list[float] = []
    for e in en_items:
        best = None
        for j, x in enumerate(xx_items):
            if j in used:
                continue
            if value_same(e.text, x.text):
                dy = (x.yc - e.yc)
                if abs(dy) <= ANCHOR_OFF_TOL and (best is None or abs(dy) < abs(best[0])):
                    best = (dy, x.xc - e.xc, j)
        if best is not None:
            used.add(best[2])
            ay.append(best[0])
            ax.append(best[1])
    y_med = median(ay) if ay else 0.0
    x_med = median(ax) if ax else 0.0
    prune_dy = PRUNE_DY if ay else PRUNE_DY_LOOSE
    prune_dx = PRUNE_DX if ax else PRUNE_DX_LOOSE

    # 2) 统一代价矩阵(全部检查点参与)
    #    剪枝分级: 值相同(强信号)宽松, 值不同(需位置强吻合)紧
    COST_BIG = 1e6
    cost = np.full((n, m), COST_BIG)
    for i, e in enumerate(en_items):
        for j, x in enumerate(xx_items):
            dy = abs((x.yc - e.yc) - y_med)
            dx = abs((x.xc - e.xc) - x_med)
            same = value_same(e.text, x.text)
            pdy = PRUNE_DY_SAME if same else prune_dy
            pdx = PRUNE_DX_SAME if same else prune_dx
            if dy > pdy or dx > pdx:
                continue
            c = dy + W_X * dx + (REWARD_SAME if same else
                                 (PENALTY_TOK if token_count(e.text) != token_count(x.text)
                                  else PENALTY_DIFF))
            # 骨架上下文奖励(确定性信号, 只奖励不惩罚)
            if e.left_ctx and x.left_ctx and e.left_ctx.lower() == x.left_ctx.lower():
                c -= REWARD_CTX
            if e.right_ctx and x.right_ctx and e.right_ctx.lower() == x.right_ctx.lower():
                c -= REWARD_CTX
            cost[i, j] = c

    # 3) 全局最优一对一分配
    pairs: list[Pair] = []
    matched: list[tuple[Item, Item]] = []
    ri, cj = linear_sum_assignment(cost)
    for i, j in zip(ri, cj):
        if cost[i, j] >= REJECT_COST or cost[i, j] >= COST_BIG:
            continue
        e, x = en_items[i], xx_items[j]
        dy = abs((x.yc - e.yc) - y_med)
        same = value_same(e.text, x.text)
        if same:
            conf = 'high' if dy <= CONF_DY_HIGH else 'medium'
        elif dy <= CONF_DY_MED:
            conf = 'medium'      # 值不同但位置强吻合: 真实差异典型形态
        else:
            conf = 'low'
        status, note = compare(e.text, x.text)
        if conf == 'low':
            note = (note + '; ' if note else '') + '低置信匹配, 建议核对坐标'
        pairs.append(Pair(cp='', page=page1, en=e, xx=x,
                          y_off=round(x.yc - e.yc, 1), conf=conf,
                          status=status, note=note))
        matched.append((e, x))

    # 4) 二级匹配: 邻域插值确认大位移
    #    双侧: 唯一值相同候选 + 双侧邻居插值 + 偏差达标;
    #    头/尾: 页首/页尾检查点用单侧邻居的方向性区间 + 唯一同值候选。
    #    任何条件不满足 -> 保持未匹配(转人工), 底线不破。
    mm = sorted(matched, key=lambda t: t[0].yc)
    ens_yc = [t[0].yc for t in mm]
    res_en = [e for e in en_items if id(e) not in {id(t[0]) for t in matched}]
    res_xx = [x for x in xx_items if id(x) not in {id(t[1]) for t in matched}]
    if res_en and res_xx and mm:
        taken: set[int] = set()
        for e in sorted(res_en, key=lambda t: t.yc):
            x = dev = expect = None
            side = None
            expect = _expected_y(e, mm, ens_yc)
            if expect is not None:
                cands = [x2 for x2 in res_xx if id(x2) not in taken
                         and value_same(e.text, x2.text)
                         and abs(x2.yc - expect) <= SP_WIN]
                if len(cands) == 1:
                    x = cands[0]
                    dev = abs(x.yc - expect)
                    side = '插值'
            else:
                # 头/尾单侧结构性确认
                k = bisect.bisect_left(ens_yc, e.yc)
                if k >= len(mm) and k > 0:
                    py = mm[-1][1].yc          # 尾部: 必须在上邻译文位置之后
                    cands = [x2 for x2 in res_xx if id(x2) not in taken
                             and value_same(e.text, x2.text) and x2.yc > py + 5]
                    side = '页尾'
                elif k == 0:
                    py = mm[0][1].yc           # 头部: 必须在下邻译文位置之前
                    cands = [x2 for x2 in res_xx if id(x2) not in taken
                             and value_same(e.text, x2.text) and x2.yc < py - 5]
                    side = '页首'
                else:
                    cands = []
                    py = None
                if len(cands) == 1:
                    x = cands[0]
                    dev = abs(x.yc - e.yc)     # 头尾模式: 偏差 = 实际位移量(仅展示)
                    expect = x.yc
            if x is None or dev is None:
                continue
            if side == '插值' and dev > SP_DEV_TOL:
                continue                      # 插值模式偏差过大 -> 保持未匹配(人工)
            status, _ = compare(e.text, x.text)
            pairs.append(Pair(cp='', page=page1, en=e, xx=x,
                              y_off=round(x.yc - e.yc, 1), conf='medium',
                              status=status,
                              note=(f'大位移({side}确认): 期望y≈{expect:.0f}, '
                                    f'实际y={x.yc:.0f}, 偏差{dev:.0f}pt')))
            taken.add(id(x))

    # 5) 残余 -> 未匹配(带插值期望, 供截图)/多出
    done_en = {id(p.en) for p in pairs if p.en is not None}
    done_xx = {id(p.xx) for p in pairs if p.xx is not None}
    for e in en_items:
        if id(e) in done_en:
            continue
        exp = _expected_y(e, mm, [t[0].yc for t in mm]) if mm else None
        pairs.append(Pair(cp='', page=page1, en=e, xx=None, y_off=None, conf='-',
                          status='译文未匹配',
                          note='译文中未找到对应红字(疑漏标或位置差异过大)',
                          expect_y=(round(exp, 1) if exp is not None else None),
                          expect_dx=(round(x_med, 1) if exp is not None else None)))
    for x in xx_items:
        if id(x) in done_xx:
            continue
        pairs.append(Pair(cp='', page=page1, en=None, xx=x, y_off=None, conf='-',
                          status='译文多出', note='译文中多出的红字(英文侧无对应)'))
    return pairs


def build_pairs(en_items: list[Item], xx_items: list[Item], en_pages: int, xx_pages: int) -> list[Pair]:
    pairs: list[Pair] = []
    max_page = max(en_pages, xx_pages)
    for p in range(max_page):
        e = [it for it in en_items if it.page == p]
        x = [it for it in xx_items if it.page == p]
        pairs.extend(match_page(e, x, p + 1))
    return pairs


def _num_eq(a: str, b: str) -> bool:
    """数值等效: 归一化后相等或字符串相同(点/逗号互换视为同一数字)"""
    if a == b:
        return True
    try:
        na, nb = norm_num(a.replace(',', '.')), norm_num(b.replace(',', '.'))
        if na is not None and nb is not None:
            return abs(na - nb) < 1e-9
    except Exception:
        pass
    return False


def _multiset_diff(a: list, b: list):
    """多重集合差: a 中不在 b 里的元素(保留重复, 数值等效匹配)"""
    rem = list(b)
    out = []
    for t in a:
        idx = None
        for i, r in enumerate(rem):
            if _num_eq(t, r):
                idx = i
                break
        if idx is not None:
            rem.pop(idx)
        else:
            out.append(t)
    return out


def _multiset_eq(a: list, b: list) -> bool:
    """多重集合相等(数值等效)"""
    if len(a) != len(b):
        return False
    rem = list(b)
    for t in a:
        idx = None
        for i, r in enumerate(rem):
            if _num_eq(t, r):
                idx = i
                break
        if idx is None:
            return False
        rem.pop(idx)
    return True


def _digits_compose(rest: str, pieces: list):
    """rest 数字串能否由 pieces(数字串列表)的连续子段精确组成。
    每步选择与当前剩余串前缀匹配的'最长' piece, 避免贪婪短吃长(如 60 吃掉 60245 前缀)。
    仅纯数字 piece 参与(小数/千分符不拆分连写)。"""
    clean = [p for p in pieces if p.isdigit()]
    r = rest
    used = []
    while r:
        cands = [p for p in clean if r.startswith(p)]
        if not cands:
            return None
        hit = max(cands, key=lambda p: len(p))   # 最长前缀优先
        used.append(hit)
        r = r[len(hit):]
    return used


def resolve_aggregation(pairs: list[Pair]):
    """聚合差异合并(双向): 满足四条件的「待人工+未匹配」或「待人工+多出」成对
    -> 待人工项重标为「疑聚合差异」单条, 残项标为「已并入聚合差异」不重复计数。
    方向一(译文聚合/英文拆散): 待人工项译文token>英文token, 残为未匹配邻项;
    方向二(英文聚合/译文拆散): 待人工项英文token>译文token, 残为多出邻项。
    四条件(全过才合并): 同页; 数字个数不等; 差值token与残项token完全一致(多集差);
    位置邻近(<40pt)且残项唯一。任一不满足保持原状态(转人工), 不动摇底线。"""
    by_page: dict[int, list[Pair]] = {}
    for p in pairs:
        by_page.setdefault(p.page, []).append(p)
    for ps in by_page.values():
        pend = [p for p in ps if p.status == '待人工' and p.xx is not None and p.en is not None]
        um = [p for p in ps if p.status == '译文未匹配' and p.en is not None]
        ex = [p for p in ps if p.status == '译文多出' and p.xx is not None]
        for q in pend:
            qtoks = TOKEN_RE.findall(q.xx.text)
            etoks = TOKEN_RE.findall(q.en.text)
            # 方向一: 译文聚合, 英文拆散 -> 残项为未匹配
            if len(qtoks) > len(etoks):
                matches = []
                for u in um:
                    ut = TOKEN_RE.findall(u.en.text)
                    diff = _multiset_diff(qtoks, etoks)
                    if ut and u.en and q.en and abs(u.en.yc - q.en.yc) < 40 \
                            and _multiset_eq(diff, ut):
                        matches.append(u)
                if len(matches) != 1:
                    continue            # 无匹配或不唯一 -> 保守不合并
                u = matches[0]
                q.status = '疑聚合差异'
                q.note = (f'疑聚合差异: 译文将相邻检查点 {u.cp}({u.en.text}) 的数字与 '
                          f'{q.en.text} 聚合为一个红字项, 需确认排版是否影响数值; '
                          f'原未匹配项 {u.cp} 已并入本条')
                u.status = '已并入聚合差异'
                u.note = f'已并入检查点 {q.cp}(疑聚合差异), 不再单独计为问题项; 数字疑似随排版聚合存在'
                continue
            # 方向二: 英文聚合, 译文拆散 -> 残项为多出(且已在待人工对中)
            if len(etoks) > len(qtoks):
                diff = _multiset_diff(etoks, qtoks)
                if not diff:
                    continue
                matches = []
                for o in ex:
                    ot = TOKEN_RE.findall(o.xx.text)
                    if o.xx and q.xx and abs(o.xx.yc - q.xx.yc) < 40 \
                            and _multiset_eq(ot, diff):
                        matches.append(o)
                if len(matches) != 1:
                    continue            # 无匹配或不唯一 -> 保守不合并
                o = matches[0]
                q.status = '疑聚合差异'
                q.note = (f'疑聚合差异(拆分布局): 译文将检查点 {q.cp} 的数字拆散为多个红字项'
                          f'(含多出项 "{o.xx.text}"), 需确认排版是否影响数值; 原多出项已并入本条')
                o.status = '已并入聚合差异'
                o.note = f'已并入检查点 {q.cp}(疑聚合差异-拆分), 不再单独计为问题项'
                continue
        # 方向三: 数字串级联连写。"不一致"对中, 长串一侧的数字串若由
        # 对侧当前值 + 对侧邻近检查点值完全级联组成(用到当前值) -> 连写聚合, 不是数值错误。
        for q in [p for p in ps if p.status == '不一致' and p.en is not None and p.xx is not None]:
            en_t = ''.join(TOKEN_RE.findall(q.en.text))
            xx_t = ''.join(TOKEN_RE.findall(q.xx.text))
            if not en_t or not xx_t:
                continue
            long_t, short_t, side = (xx_t, en_t, '译文') if len(xx_t) > len(en_t) else (en_t, xx_t, '英文')
            if len(long_t) < 4 or len(short_t) < 1:
                continue
            # 级联组件候选: 短侧当前值 + 与短侧检查点 y 邻近(<50)的检查点值
            if side == '译文':
                pieces = [short_t] + [t for p2 in ps if p2 is not q and p2.en is not None
                                      and abs(p2.en.yc - q.en.yc) < 50
                                      for t in TOKEN_RE.findall(p2.en.text)]
            else:
                pieces = [short_t] + [t for p2 in ps if p2 is not q and p2.xx is not None
                                      and abs(p2.xx.yc - q.xx.yc) < 50
                                      for t in TOKEN_RE.findall(p2.xx.text)]
            used = _digits_compose(long_t, pieces)
            if used is not None and short_t in used:
                q.status = '疑聚合差异'
                q.note = (f'疑聚合差异(连写): {side}值 "{long_t}" 由当前值 "{short_t}" 与相邻值 '
                          f'{" ".join(u for u in used if u != short_t)} 连续级联而成, 疑{side}侧将相邻数字连写; '
                          f'数字内容一致, 需人工确认排版; 原"{q.en.text}" vs "{q.xx.text}" 已归并')
                continue

        # 方向四: 区域成组聚合差异。同页 y 邻近(<60pt)的不一致/待人工/译文多出集合, 若
        # 其 token 全集与英文侧检查点 token 全集数值等效(含点/逗号差异, 无值变化), 判定为
        # 拆分/聚合排版差异, 全部并入(首个非多出项为主项), 不重复列示。
        # 安全: 数值等效条件保证真差异(290 vs 29)不会被合并。
        bypg: dict[int, list[Pair]] = {}
        for p in pairs:
            pg = p.en.page if p.en is not None else (p.xx.page if p.xx is not None else None)
            if pg is not None:
                bypg.setdefault(pg, []).append(p)
        for ps in bypg.values():
            region = [p for p in ps if p.status in ('不一致', '待人工', '译文多出')]
            if len(region) < 2:
                continue
            pend = [p for p in region if p.status in ('不一致', '待人工')]
            pout = [p for p in region if p.status == '译文多出']
            if not pend or not pout:
                continue
            ys = [p.en.yc if p.en is not None else p.xx.yc for p in region]
            if max(ys) - min(ys) > 60:
                continue
            # 英文侧 token 全集: 待人工/不一致项的 en 检查点
            en_tokens = []
            for p in pend:
                en_tokens.extend(TOKEN_RE.findall(p.en.text))
            # 译文侧 token 全集: 待人工/不一致已配 xx + 多出 xx
            xx_tokens = []
            for p in region:
                if p.xx is not None:
                    xx_tokens.extend(TOKEN_RE.findall(p.xx.text))
            if not en_tokens or not xx_tokens:
                continue
            if _multiset_eq(en_tokens, xx_tokens):
                main = sorted(pend, key=lambda p: p.en.yc)[0]
                rest = [p for p in region if p is not main]
                main.status = '疑聚合差异'
                main.note = (f'疑聚合差异(区域拆分): 英文侧若干检查点数字与译文拆散项数值等效'
                             f'(含 {len(pout)} 个多出项), 属拆分/聚合排版差异, 需人工确认排版影响; '
                             f'已并入 {len(rest)} 条')
                for p2 in rest:
                    p2.status = '已并入聚合差异'
                    p2.note = f'已并入同区域疑聚合差异(检查点 {main.cp}), 不再重复计为问题项'

        # 方向五: 未匹配(英文多值聚合) + 多出(译文拆分项)成组, 因排版位移大未配成一对。
        # 独立遍历(不与方向四耦合, 方向四的 continue 不影响)。
        for ps in bypg.values():
            um = [p for p in ps if p.status == '译文未匹配' and p.en is not None]
            ex = [p for p in ps if p.status == '译文多出' and p.xx is not None]
            for u in um:
                utoks = TOKEN_RE.findall(u.en.text)
                if len(utoks) < 2:
                    continue
                cands = [o for o in ex if abs(o.xx.yc - u.en.yc) < 150]
                if not cands:
                    continue
                all_toks = []
                for o in cands:
                    all_toks.extend(TOKEN_RE.findall(o.xx.text))
                if len(cands) >= 2 and _multiset_eq(all_toks, utoks):
                    u.status = '疑聚合差异'
                    u.note = (f'疑聚合差异(拆分+位移): 英文检查点 {u.cp} "{u.en.text}" 为聚合值, '
                              f'译文拆分为 {len(cands)} 项({" ".join(o.xx.text for o in cands)}), '
                              f'因排版位移大而未匹配; 数字内容一致, 已合并供人工确认排版影响')
                    for o in cands:
                        o.status = '已并入聚合差异'
                        o.note = f'已并入检查点 {u.cp}(疑聚合差异-拆分位移), 不再重复计为问题项'

        # 方向七: 排版换位确认。不一致项的英文/译文同区域(y±25pt)红字值集合数值等效
        # -> 判一致(排版换位), 消除排版错位/换位误报。
        # 安全: 集合不等(值真不同, 如 290 vs 29)不触发, 保留不一致供人工核对。
        for ps in bypg.values():
            mis = [p for p in ps if p.status == '不一致' and p.en is not None and p.xx is not None]
            R = 25.0
            for q in mis:
                E = [t for p2 in ps if p2.en is not None and abs(p2.en.yc - q.en.yc) <= R
                     for t in TOKEN_RE.findall(p2.en.text)]
                X = [t for p2 in ps if p2.xx is not None and abs(p2.xx.yc - q.xx.yc) <= R
                     for t in TOKEN_RE.findall(p2.xx.text)]
                if not E or not X:
                    continue
                if _multiset_eq(E, X):
                    q.status = '一致'
                    q.note = f'排版换位确认: 同区域红字值集合数值等效, 判一致(排版换位); ' + q.note
    return pairs

# ---------------- 截图 ----------------
def snap(doc: fitz.Document, pno: int, bbox, path: str):
    page = doc[pno]
    r = fitz.Rect(bbox)
    r = fitz.Rect(r.x0 - SNAP_PAD, r.y0 - SNAP_PAD, r.x1 + SNAP_PAD, r.y1 + SNAP_PAD) & page.rect
    pix = page.get_pixmap(clip=r, dpi=SNAP_DPI)
    pix.save(path)

# ---------------- Excel 报告 ----------------
FILL = {
    '不一致':   PatternFill('solid', fgColor='C00000'),
    '格式差异': PatternFill('solid', fgColor='ED7D31'),
    '待人工':   PatternFill('solid', fgColor='FFC000'),
    '疑聚合差异': PatternFill('solid', fgColor='DAA520'),
    '译文未匹配': PatternFill('solid', fgColor='FFC000'),
    '译文多出':  PatternFill('solid', fgColor='FFC000'),
    '已并入聚合差异': PatternFill('solid', fgColor='BFBFBF'),
    '一致':     PatternFill('solid', fgColor='70AD47'),
    'high':    PatternFill('solid', fgColor='70AD47'),
    'medium':  PatternFill('solid', fgColor='FFC000'),
    'low':     PatternFill('solid', fgColor='C00000'),
    '-':       PatternFill('solid', fgColor='BFBFBF'),
}
HDR_FILL = PatternFill('solid', fgColor='1F4E79')
HDR_FONT = Font(color='FFFFFF', bold=True)
RED_FONT = Font(color='9C0006', bold=True)


def style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal='center')
    ws.freeze_panes = 'A2'


def build_excel(path, en_file, en_items, results, snaps_dir, color_notes=None):
    wb = Workbook()

    # ---- Sheet1 汇总 ----
    ws = wb.active
    ws.title = '汇总'
    headers = ['文件名', '语言', '英文检查点数', '译文红字项', '一致', '其中大位移确认',
               '不一致', '格式差异', '待人工', '疑聚合差异', '译文未匹配', '译文多出', '结论']
    ws.append(headers)
    tot = {k: 0 for k in ['一致', '不一致', '格式差异', '待人工', '疑聚合差异',
                          '译文未匹配', '译文多出']}
    tot_sp = 0
    for r in results:
        cnt = {k: 0 for k in tot}
        for p in r['pairs']:
            if p.status in cnt:
                cnt[p.status] += 1
        for k in tot:
            tot[k] += cnt[k]
        tot_sp += r['n_sp']
        problems = cnt['不一致'] + cnt['格式差异'] + cnt['待人工'] + cnt['疑聚合差异'] \
            + cnt['译文未匹配'] + cnt['译文多出']
        concl = '✓ 通过' if problems == 0 else f'⚠ 需人工({problems}项)'
        ws.append([r['file'], r['lang'], len(en_items), r['n_xx'],
                   cnt['一致'], r['n_sp'], cnt['不一致'], cnt['格式差异'], cnt['待人工'],
                   cnt['疑聚合差异'], cnt['译文未匹配'], cnt['译文多出'], concl])
        if problems:
            for c in range(1, len(headers) + 1):
                ws.cell(row=ws.max_row, column=c).fill = PatternFill('solid', fgColor='FFF2CC')
            ws.cell(row=ws.max_row, column=len(headers)).font = RED_FONT
    ws.append(['总计', '', len(en_items) * len(results), '', tot['一致'], tot_sp,
               tot['不一致'], tot['格式差异'], tot['待人工'], tot['疑聚合差异'],
               tot['译文未匹配'], tot['译文多出'], ''])
    for c in range(1, len(headers) + 1):
        ws.cell(row=ws.max_row, column=c).font = Font(bold=True)
    style_header(ws, len(headers))
    for c, w in zip(range(1, len(headers) + 1), [34, 8, 13, 11, 7, 14, 9, 9, 8, 12, 12, 10, 15]):
        ws.column_dimensions[get_column_letter(c)].width = w

    # ---- Sheet2 明细 ----
    ws = wb.create_sheet('明细')
    headers = ['语言', '文件名', '检查点', '页', '状态', '置信度', '英文值', '译文值',
               'y偏移pt', '备注', '截图(EN)', '截图(译文)']
    ws.append(headers)
    for r in results:
        for p in r['pairs']:
            en_v = p.en.text if p.en else ''
            xx_v = p.xx.text if p.xx else ''
            has_snap = p.status != '一致' and p.status != '已并入聚合差异'
            se = os.path.join('snaps', r['lang'], f"{p.cp or 'XX'}_{r['lang']}_EN.png") if p.en and has_snap else ''
            sx = os.path.join('snaps', r['lang'], f"{p.cp or 'XX'}_{r['lang']}_XX.png") if p.xx and has_snap else ''
            sx2 = os.path.join('snaps', r['lang'], f"{p.cp or 'XX'}_{r['lang']}_XX@期望位置.png") \
                if (p.status == '译文未匹配' and p.expect_y is not None) else ''
            ws.append([r['lang'], r['file'], p.cp, p.page, p.status, p.conf,
                       en_v, xx_v, p.y_off if p.y_off is not None else '', p.note, se, sx or sx2])
            cell = ws.cell(row=ws.max_row, column=5)
            if p.status in FILL:
                cell.fill = FILL[p.status]
    ws.auto_filter.ref = f'A1:{get_column_letter(len(headers))}{ws.max_row}'
    style_header(ws, len(headers))
    for c, w in zip(range(1, len(headers) + 1), [7, 30, 9, 5, 11, 7, 16, 16, 8, 44, 26, 26]):
        ws.column_dimensions[get_column_letter(c)].width = w

    # ---- Sheet3 矩阵 ----
    ws = wb.create_sheet('语言矩阵')
    langs = [r['lang'] for r in results]
    headers = ['检查点', '英文值'] + langs
    ws.append(headers)
    # 行: 按页 + 页内序号(与明细一致), 修复全局序号导致的错位
    en_sorted = sorted(en_items, key=lambda i: (i.page, i.yc, i.bbox[0]))
    en_by_page: dict[int, list[Item]] = {}
    for it in en_sorted:
        en_by_page.setdefault(it.page, []).append(it)
    en_rows: list[Item] = []
    for pg in sorted(en_by_page):
        en_rows.extend(en_by_page[pg])
    pair_map = {}   # (lang, cp) -> pair
    xx_only = []    # (lang, pair)
    for r in results:
        for p in r['pairs']:
            if p.en is not None:
                pair_map[(r['lang'], p.cp)] = p
            else:
                xx_only.append((r['lang'], p))
    def mark(p):
        return {'一致': '✓', '不一致': '✗', '格式差异': 'F', '待人工': '?',
                '疑聚合差异': '≈', '已并入聚合差异': '·',
                '译文未匹配': '∅', '译文多出': '＋'}.get(p.status, '?')
    for it in en_rows:
        idx = en_by_page[it.page].index(it) + 1
        cp = f'P{it.page + 1}-{idx}'
        row = [cp, it.text]
        for r in results:
            p = pair_map.get((r['lang'], cp))
            row.append(mark(p) if p else '?')
        ws.append(row)
        for c in range(3, len(headers) + 1):
            v = ws.cell(row=ws.max_row, column=c).value
            for st, m in [('不一致', '✗'), ('格式差异', 'F'), ('待人工', '?'),
                          ('疑聚合差异', '≈'), ('译文未匹配', '∅')]:
                if v == m:
                    ws.cell(row=ws.max_row, column=c).fill = FILL[st]
            if v == '·':
                ws.cell(row=ws.max_row, column=c).fill = FILL['已并入聚合差异']
    if xx_only:
        ws.append([])
        ws.append(['— 以下为译文多出项 —'])
        for lang, p in xx_only:
            ws.append([f'(译文多出)', p.xx.text, *[lang if l == lang else '' for l in langs]])
    # 矩阵下方图例
    ws.append([])
    ws.append(['图例'])
    legend = [
        ('✓', '一致：匹配成功且归一化数值相等'),
        ('✗', '不一致：数值不同，必须处理'),
        ('F', '格式差异：数值相同但书写不同(如 2.5 vs 2,5)'),
        ('?', '待人工：匹配存在歧义或数字个数不同'),
        ('∅', '译文未匹配：英文有此检查点但译文未找到红字(疑漏标)'),
        ('≈', '疑聚合差异：译文将相邻检查点的数字聚为一体，已合并为单条'),
        ('·', '已并入聚合差异：上述合并项的原未匹配项，不重复计数'),
        ('＋', '译文多出：译文有红字但英文无对应'),
        ('', '底色说明：绿=一致 红=不一致 橙=格式差异 黄=待人工/未匹配 暗金=疑聚合 灰=已并入'),
    ]
    for sym, desc in legend:
        ws.append([sym, desc])
        r = ws.max_row
        cell = ws.cell(row=r, column=1)
        if sym == '✓':
            ws.cell(row=r, column=1).fill, ws.cell(row=r, column=2).fill = FILL['一致'], FILL['一致']
        elif sym == '✗':
            ws.cell(row=r, column=1).fill, ws.cell(row=r, column=2).fill = FILL['不一致'], FILL['不一致']
        elif sym == 'F':
            ws.cell(row=r, column=1).fill, ws.cell(row=r, column=2).fill = FILL['格式差异'], FILL['格式差异']
        elif sym == '?':
            ws.cell(row=r, column=1).fill, ws.cell(row=r, column=2).fill = FILL['待人工'], FILL['待人工']
        elif sym == '∅':
            ws.cell(row=r, column=1).fill, ws.cell(row=r, column=2).fill = FILL['译文未匹配'], FILL['译文未匹配']
        elif sym == '≈':
            ws.cell(row=r, column=1).fill, ws.cell(row=r, column=2).fill = FILL['疑聚合差异'], FILL['疑聚合差异']
        elif sym == '·':
            ws.cell(row=r, column=1).fill, ws.cell(row=r, column=2).fill = FILL['已并入聚合差异'], FILL['已并入聚合差异']
        elif sym == '＋':
            ws.cell(row=r, column=1).fill, ws.cell(row=r, column=2).fill = FILL['译文多出'], FILL['译文多出']
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
    style_header(ws, len(headers))
    ws.column_dimensions['A'].width = 9
    ws.column_dimensions['B'].width = 18
    for c in range(3, len(headers) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 6

    # ---- Sheet4 说明 ----
    ws = wb.create_sheet('说明')
    lines = [
        ['MEE 多国语数字校对报告说明'],
        [''],
        ['英文指示稿', en_file],
        ['检查点编号', 'P{页码}-{页内序号}, 以英文指示稿为准, 全语言统一'],
        [''],
        ['状态定义'],
        ['一致', '匹配置信度高/中, 归一化后数值完全相等且书写形式相同'],
        ['大位移确认', '计入一致; 译文排版重排导致红字大幅移动, 由上下邻居插值+唯一候选+偏差达标三条件确认, 备注含判据'],
        ['不一致', '匹配置信度高/中, 数值不同 (红色, 必须处理)'],
        ['格式差异', '数值相同但书写不同, 如 2.5 vs 2,5 (橙色, 转人工确认); 纯空格分组差异(25 35 42 50)计入一致'],
        ['待人工', '匹配存在歧义或数字个数不同 (黄色, 转人工)'],
        ['疑聚合差异', '译文将相邻两个检查点的数字聚为一个红字项, 已自动合并为单条; 需确认排版是否影响数值 (暗金, 转人工)'],
        ['已并入聚合差异', '原"译文未匹配"项因被合并, 不重复计为问题 (灰色, 供追溯)'],
        ['译文未匹配', '英文有此检查点但译文未找到红字 (黄色, 疑漏标); 截图列为译文期望位置'],
        ['译文多出', '译文有红字但英文无对应 (黄色, 疑多标)'],
        [''],
        ['自动判定底线'],
        ['原则', '不确定不判一致; 自动判一致仅有两条路径: ①一级高/中置信+归一化值相等 ②大位移三判据全过'],
        ['大位移三判据', '唯一值相同候选 + 双侧已配对邻居插值支撑 + 偏差≤阈值, 全部满足才自动判, 其余一律转人工'],
        ['聚合差异四条件', '同页; 待人工项译文token数>英文token数; 多出token与未匹配邻项token完全一致; 位置邻近且未匹配邻项唯一; 全过才合并, 否则保持待人工'],
        [''],
        ['置信度'],
        ['high', '值相同且 y偏移与页偏移基准一致'],
        ['medium', '值相同但位置略偏, 或值不同但位置强吻合(真实差异典型形态)'],
        ['low', '匹配信号存在矛盾, 备注已提示人工核对坐标'],
        [''],
        ['使用方法'],
        ['1', '先看"汇总"页, 结论为"✓ 通过"的语言无需处理'],
        ['2', '在"明细"页筛选状态列, 只看问题项; 双方截图见 snaps 文件夹'],
        ['3', '"语言矩阵"页用于快速定位: 整行绿中冒出的红/橙/黄格即为真差异'],
        ['4', '复核结论记录在明细表右侧自行加列即可'],
    ]
    if color_notes:
        lines.append([''])
        lines.append(['非标准红提示'])
        for nte in color_notes:
            lines.append(['提示', nte])
    for row in lines:
        ws.append(row)
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 80
    ws['A1'].font = Font(bold=True, size=14)
    for rw, key in [(6, '状态定义'), (14, '自动判定底线'), (17, '置信度'), (21, '使用方法')]:
        for r in range(1, ws.max_row + 1):
            if ws.cell(row=r, column=1).value == key:
                ws.cell(row=r, column=1).font = Font(bold=True)
                break

    wb.save(path)

# ---------------- 主流程 ----------------
# ---------------- HTML 报告 ----------------
LANG_NAMES = {
    'De': '德语', 'Fr': '法语', 'Nl': '荷兰语', 'Es': '西班牙语', 'It': '意大利语',
    'El': '希腊语', 'Pt': '葡萄牙语', 'Da': '丹麦语', 'Sv': '瑞典语', 'Bg': '保加利亚语',
    'Pl': '波兰语', 'No': '挪威语', 'Fi': '芬兰语', 'Cs': '捷克语', 'Sk': '斯洛伐克语',
    'Hu': '匈牙利语', 'Sl': '斯洛文尼亚语', 'Ro': '罗马尼亚语', 'Et': '爱沙尼亚语',
    'Lv': '拉脱维亚语', 'Lt': '立陶宛语', 'Hr': '克罗地亚语', 'Sr': '塞尔维亚语',
    'Uk': '乌克兰语',
}

_HTML_CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI','Microsoft YaHei',sans-serif;background:#f0f2f5;color:#24292f;font-size:14px}
header{background:#1f3a5f;color:#fff;padding:20px 32px}
header h1{font-size:20px;margin-bottom:6px}
header .meta{color:#b8c7dc;font-size:12px;line-height:1.7}
.dash{display:flex;gap:14px;flex-wrap:wrap;padding:18px 32px;background:#fff;border-bottom:1px solid #e1e4e8}
.stat{padding:10px 18px;border-radius:8px;background:#f6f8fa;border:1px solid #e1e4e8;min-width:96px}
.stat b{display:block;font-size:22px}
.stat span{font-size:12px;color:#57606a}
.stat.ok b{color:#1a7f37}.stat.err b{color:#cf222e}.stat.warn b{color:#9a6700}.stat.info b{color:#0969da}
.filters{padding:14px 32px;display:flex;gap:8px;flex-wrap:wrap;position:sticky;top:0;background:#f0f2f5;z-index:9;border-bottom:1px solid #e1e4e8}
.filters button{border:1px solid #d0d7de;background:#fff;border-radius:16px;padding:5px 14px;cursor:pointer;font-size:13px}
.filters button.active{background:#1f3a5f;color:#fff;border-color:#1f3a5f}
main{padding:20px 32px 60px}
section.lang{margin-bottom:28px;background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);overflow:hidden}
section.lang>h2{font-size:15px;padding:12px 20px;background:#fafbfc;border-bottom:1px solid #e1e4e8;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
section.lang>h2 .badge{font-size:12px;padding:2px 10px;border-radius:10px;color:#fff}
section.lang>h2 .fname{color:#57606a;font-size:12px;font-weight:400}
.allpass{padding:16px 20px;color:#1a7f37}
.item{border-left:4px solid #d0d7de;padding:14px 20px;border-bottom:1px solid #eee}
.item:last-child{border-bottom:none}
.item.st-不一致{border-left-color:#cf222e}.item.st-格式差异{border-left-color:#fb8500}
.item.st-待人工,.item.st-译文未匹配,.item.st-译文多出{border-left-color:#e3b341}
.item.st-疑聚合差异{border-left-color:#b8860b}
.item.st-大位移{border-left-color:#0969da}
.item .head{display:flex;gap:10px;align-items:baseline;flex-wrap:wrap;margin-bottom:10px}
.item .cp{font-weight:600;font-size:13px}
.badge{display:inline-block;font-size:12px;padding:2px 10px;border-radius:10px;color:#fff}
.b-不一致{background:#cf222e}.b-格式差异{background:#fb8500}.b-待人工,.b-译文未匹配,.b-译文多出{background:#bf8700}.b-大位移{background:#0969da}.b-一致{background:#1a7f37}.b-疑聚合差异{background:#b8860b}
.compare{display:flex;gap:14px;align-items:stretch;flex-wrap:wrap}
figure{background:#f6f8fa;border:1px solid #e1e4e8;border-radius:6px;padding:8px;text-align:center}
figure img{max-width:330px;max-height:180px;display:block}
figure figcaption{font-size:12px;color:#57606a;margin-top:6px}
figure img.missing{width:200px;height:110px;object-fit:contain;opacity:.35}
.vs{align-self:center;font-weight:700;color:#8c959f;font-size:13px}
.verdict{margin-top:10px;font-size:13px;line-height:1.8}
.verdict .v{display:inline-block;padding:1px 8px;border-radius:4px;margin-right:8px;font-weight:600}
.vv-en{background:#dbeafe}.vv-xx{background:#ffebe9}.v-same{background:#dafbe1}
.note{color:#57606a;font-size:12px;margin-top:4px}
"""


def _b64(path: str):
    try:
        with open(path, 'rb') as f:
            return 'data:image/png;base64,' + base64.b64encode(f.read()).decode()
    except OSError:
        return None


def build_html(path_html: str, en_file: str, en_items: list, results: list, snaps_dir: str):
    import html as _h
    from datetime import datetime

    status_key = {'不一致': '不一致', '格式差异': '格式差异', '待人工': '待人工',
                  '译文未匹配': '译文未匹配', '译文多出': '译文多出', '疑聚合差异': '疑聚合差异'}
    cnt = Counter()
    n_sp = 0
    for r in results:
        for p in r['pairs']:
            if p.status == '已并入聚合差异':
                continue                 # 合并项不重复计数
            if p.status in status_key:
                cnt[p.status] += 1
            elif p.status == '一致' and '大位移' in p.note:
                n_sp += 1
    n_items = len(en_items) * len(results)
    n_ok = n_items - sum(cnt.values())

    def snap_html(lang: str, p: Pair, kind: str) -> str:
        if p.en is None and p.xx is None:
            return ''
        tag = _snap_tag(p)
        if kind == 'EN':
            f = os.path.join(snaps_dir, lang, f'{tag}_{lang}_EN.png')
            cap = f'英文指示稿 · 值 {_h.escape(p.en.text)}' if p.en else '英文指示稿'
        elif kind == 'XX' and p.xx is not None:
            f = os.path.join(snaps_dir, lang, f'{tag}_{lang}_XX.png')
            cap = f'{LANG_NAMES.get(lang, lang)} · 值 {_h.escape(p.xx.text)}'
        else:
            f = os.path.join(snaps_dir, lang, f'{tag}_{lang}_XX@期望位置.png')
            cap = f'{LANG_NAMES.get(lang, lang)} · 期望位置(未找到红字)'
        img = _b64(f)
        # 需人工确认项可点击打开复核 PDF 定位
        need_review = p.status not in ('一致', '已并入聚合差异')
        if need_review:
            if kind == 'EN' and p.en is not None:
                href = f'复核PDF/EN.pdf#page={p.en.page + 1}'   # 英文截图 -> 英文指示稿复核版
                pg = p.en.page + 1
            elif p.xx is not None:
                href = f'复核PDF/{lang}.pdf#page={p.xx.page + 1}'  # 译文截图 -> 对应语言复核版
                pg = p.xx.page + 1
            else:
                href = f'复核PDF/{lang}.pdf#page={p.en.page + 1}'  # 期望位置 -> 对应语言复核版
                pg = p.en.page + 1
            inner = f'<figure class="linkable" style="cursor:pointer" title="点击打开复核PDF定位(第{pg}页)"><img src="{img}"><figcaption>{cap} · 点击定位</figcaption></figure>'
            return f'<a href="{href}" target="_blank">{inner}</a>'
        return f'<figure><img src="{img}"><figcaption>{cap}</figcaption></figure>'

    parts = [
        '<!DOCTYPE html><html lang="zh"><head><meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width,initial-scale=1">',
        '<title>数字校对报告</title>',
        f'<style>{_HTML_CSS}</style></head><body data-filter="all">',
        '<header><h1>多国语数字校对报告</h1>',
        f'<div class="meta">英文指示稿: {_h.escape(en_file)} &nbsp;|&nbsp; 语言数: {len(results)} '
        f'&nbsp;|&nbsp; 检查点: {len(en_items)}/语言 &nbsp;|&nbsp; 生成时间: {datetime.now():%Y-%m-%d %H:%M}</div></header>',
        '<div class="dash">',
        f'<div class="stat ok"><b>{n_ok}</b><span>一致</span></div>',
        f'<div class="stat info"><b>{n_sp}</b><span>大位移确认</span></div>',
        f'<div class="stat err"><b>{cnt["不一致"]}</b><span>不一致</span></div>',
        f'<div class="stat warn"><b>{cnt["格式差异"]}</b><span>格式差异</span></div>',
        f'<div class="stat warn"><b>{cnt["待人工"]}</b><span>待人工</span></div>',
        f'<div class="stat warn"><b>{cnt["疑聚合差异"]}</b><span>疑聚合差异</span></div>',
        f'<div class="stat warn"><b>{cnt["译文未匹配"]}</b><span>译文未匹配</span></div>',
        f'<div class="stat warn"><b>{cnt["译文多出"]}</b><span>译文多出</span></div>',
        f'<div class="stat"><b>{len(en_items)}×{len(results)}</b><span>检查点总数</span></div>',
        '</div>',
    ]

    # 筛选按钮
    btns = [('all', '全部', n_sp + sum(cnt.values()))]
    for st in ['不一致', '译文未匹配', '待人工', '疑聚合差异', '格式差异', '译文多出']:
        if cnt[st]:
            btns.append((st, st, cnt[st]))
    if n_sp:
        btns.append(('大位移', '大位移确认', n_sp))
    parts.append('<div class="filters"><button class="active" data-f="all" onclick="setFilter(this.dataset.f)">'
                 f'全部 ({btns[0][2]})</button>')
    for f, label, n in btns[1:]:
        parts.append(f'<button data-f="{f}" onclick="setFilter(this.dataset.f)">{label} ({n})</button>')
    parts.append('</div><main>')

    # 语言区块
    for r in results:
        lang = r['lang']
        lname = LANG_NAMES.get(lang, lang)
        probs = [p for p in r['pairs'] if p.status != '一致' and p.status != '已并入聚合差异']
        sps = [p for p in r['pairs'] if p.status == '一致' and '大位移' in p.note]
        problems = len(probs)
        concl = '✓ 通过' if problems == 0 else f'⚠ 需人工({problems}项)'
        bcolor = '#1a7f37' if problems == 0 else '#bf8700'
        parts.append(f'<section class="lang" id="lang-{lang}"><h2>'
                     f'{lname} <span class="fname">{_h.escape(r["file"])} · 检查点{len(en_items)} '
                     f'· 一致{len(en_items) - problems}</span>'
                     f'<span class="badge" style="background:{bcolor}">{concl}</span></h2>')
        if not probs and not sps:
            parts.append('<div class="allpass">✓ 本语言全部通过, 无需处理</div></section>')
            continue
        for p in probs + sps:
            is_sp = p.status == '一致'
            st = '大位移' if is_sp else p.status
            ev = p.en.text if p.en else '(无)'
            xv = p.xx.text if p.xx else '(缺失)'
            same = p.status == '一致'
            vv_en = 'v-same' if same else 'vv-en'
            vv_xx = 'v-same' if same else 'vv-xx'
            pg = p.en.page + 1 if p.en else (p.xx.page + 1 if p.xx else '-')
            yoff = f'{p.y_off:+.0f}pt' if p.y_off is not None else '—'
            conf = p.conf if p.conf != '-' else '—'
            parts.append(
                f'<div class="item st-{st}" data-status="{st}">'
                f'<div class="head"><span class="cp">检查点 {p.cp or "(译文多出)"}</span>'
                f'<span>第{pg}页</span><span>置信度 {conf}</span><span>y偏移 {yoff}</span>'
                f'<span class="badge b-{st}">{st}</span></div>'
                f'<div class="compare">{snap_html(lang, p, "EN")}'
                f'<div class="vs">VS</div>{snap_html(lang, p, "XX")}</div>'
                f'<div class="verdict">'
                f'<span class="v {vv_en}">英文: {_h.escape(ev)}</span>'
                f'<span class="v {vv_xx}">译文: {_h.escape(xv)}</span>'
                f'<div class="note">{_h.escape(p.note)}</div></div></div>')
        parts.append('</section>')

    parts.append('</main><script>\n' + _HTML_JS + '\n</script></body></html>')
    with open(path_html, 'w', encoding='utf-8') as f:
        f.write('\n'.join(parts))


_HTML_JS = """
function setFilter(f){
  document.body.dataset.filter=f;
  document.querySelectorAll('.item').forEach(function(el){
    el.style.display=(f==='all'||el.dataset.status===f)?'':'none';
  });
  document.querySelectorAll('section.lang').forEach(function(sec){
    var vis=[].slice.call(sec.querySelectorAll('.item')).some(function(el){return el.style.display!=='none'});
    sec.style.display=vis?'':'none';
  });
  document.querySelectorAll('.filters button').forEach(function(b){
    b.classList.toggle('active',b.dataset.f===f);
  });
}
"""


# ---------------- 复核 PDF ----------------
def _snap_tag(p: Pair) -> str:
    """截图/标注文件名: 多出项无检查点编号, 用页+y坐标唯一化, 避免互相覆盖"""
    if p.cp:
        return p.cp
    if p.xx is not None:
        return f'XX@{p.xx.page + 1}~{p.xx.yc:.0f}'
    if p.en is not None:
        return f'UM@{p.en.page + 1}~{p.en.yc:.0f}'
    return 'XX'


def build_review_pdf(src_pdf: str, out_path: str, marks: list):
    """生成译文复核 PDF: 拷自译文源文件, 问题数字区域加高亮批注(悬停显示说明),
    不画框不压版面; 密集区域多数字高亮也不会互相干扰。
    marks: [(page0, rect, label, dashed)] — dashed 保留兼容(高亮统一, 忽略)"""
    doc = fitz.open(src_pdf)
    for pno, rect, label, _dashed in marks:
        page = doc[pno]
        r = fitz.Rect(rect) & page.rect
        if r.width < 2 or r.height < 2:
            continue
        # 高亮批注(黄色半透明), 直接标注在数字区域
        ha = page.add_highlight_annot(r)
        ha.set_colors(stroke=(1, 0.92, 0.23))
        ha.update()
        # 悬停才显示的说明气泡, 不占版面
        ta = page.add_text_annot(r, label, icon='note')
        ta.update()
        ta.update()
    doc.save(out_path, garbage=3, deflate=True)
    doc.close()


def cross_validate_aggregation(results: list):
    """跨语言交叉验证聚合差异: 同一检查点在 >=3 个语言都被判为「疑聚合差异」
    (其定义已保证数字内容等价) -> 视为英文侧聚合边界的系统性问题, 非译文错误,
    自动改判「一致(排版差异)」并备注。单/双语言的聚合保留人工(如 A0400 Lv/Lt)。"""
    from collections import defaultdict as _dd
    agg: dict[str, set] = _dd(set)
    for r in results:
        for p in r['pairs']:
            if p.status == '疑聚合差异' and p.en is not None:
                agg[p.cp].add(r['lang'])
    changed = 0
    for cp, langs in agg.items():
        if len(langs) < 3:
            continue
        for r in results:
            if r['lang'] not in langs:
                continue
            for p in r['pairs']:
                if p.cp == cp and p.status == '疑聚合差异':
                    p.status = '一致'
                    p.note = '排版聚合差异(≥3语言一致出现, 内容等价), 自动判一致; ' + p.note
                    changed += 1
    return changed


def lang_code(fname: str) -> str:
    m = re.search(r'_\d+([A-Za-z]+)-', fname)
    if m:
        return m.group(1)
    return os.path.splitext(fname)[0]


def main():
    ap = argparse.ArgumentParser(description='多国语数字校对工具 (P0)')
    ap.add_argument('--base', required=True, help='英文指示稿 PDF')
    ap.add_argument('--dir', required=True, help='多国语 PDF 文件夹')
    ap.add_argument('--out', default=None, help='输出目录 (默认: <dir>/_校对结果)')
    args = ap.parse_args()

    out_dir = args.out or os.path.join(args.dir, '_校对结果')
    snaps_dir = os.path.join(out_dir, 'snaps')
    os.makedirs(snaps_dir, exist_ok=True)

    print(f'英文指示稿: {args.base}')
    en_color_counter = Counter()
    en_doc = fitz.open(args.base)
    en_items = extract_items(en_doc, en_color_counter)
    print(f'英文检查点数: {len(en_items)}')
    en_sorted = sorted(en_items, key=lambda i: (i.page, i.yc, i.bbox[0]))
    en_cp = {}
    for k, it in enumerate(en_sorted, 1):
        en_cp[(it.page, round(it.yc, 1))] = f'P{it.page + 1}-{k}'

    files = [f for f in os.listdir(args.dir)
             if f.lower().endswith('.pdf') and os.path.abspath(os.path.join(args.dir, f)) != os.path.abspath(args.base)]
    files.sort()
    print(f'待校对文件: {len(files)} 个')

    results = []
    color_notes = []   # 非标准红提示
    review_dir = os.path.join(out_dir, '复核PDF')
    en_marks: dict = {}   # 英文侧标注: (page,yc,x0) -> {bbox, langs}
    en_sorted_page = {}
    for it in en_sorted:
        en_sorted_page.setdefault(it.page, []).append(it)
    for fname in files:
        path = os.path.join(args.dir, fname)
        lang = lang_code(fname)
        doc = fitz.open(path)
        cc = Counter()
        xx_items = extract_items(doc, cc)
        pairs = build_pairs(en_items, xx_items, en_doc.page_count, doc.page_count)
        # 回填检查点编号(匹配页内的编号)
        for p in pairs:
            if p.en is not None:
                lst = en_sorted_page[p.en.page]
                p.cp = f'P{p.en.page + 1}-{lst.index(p.en) + 1}'
        resolve_aggregation(pairs)
        # 非标准红提示: 文件中出现了英文稿没有的红系颜色
        en_main = {c for c, _ in en_color_counter.most_common(3)}
        odd = {c: n for c, n in cc.items() if c not in en_main}
        if odd:
            color_notes.append(f'{lang}: 检测到非英文稿主色的红系颜色 {odd}, 已一并提取, 建议与标注方确认规范')
        results.append({'file': fname, 'lang': lang, 'pairs': pairs,
                        'n_xx': len(xx_items), 'n_sp': 0, 'doc': doc, 'path': path, 'cc': cc})

    # 跨语言交叉验证聚合差异(在截图/统计前, 影响最终状态)
    n_cross = cross_validate_aggregation(results)
    if n_cross:
        print(f'跨语言交叉验证: {n_cross} 项聚合差异自动判为一致(排版差异)')

    # 统一截图/复核 PDF/统计
    for r in results:
        doc = r['doc']
        lang = r['lang']
        fname = r['file']
        path = r['path']
        pairs = r['pairs']
        lang_snap = os.path.join(snaps_dir, lang)
        n_snaps = 0
        for p in pairs:
            # 一致且非大位移确认 -> 无需截图; 大位移确认(位移大)同样配截图供人工核对; 已并入项不再单独截图
            if p.status == '一致' and '大位移' not in p.note:
                continue
            if p.status == '已并入聚合差异':
                continue
            os.makedirs(lang_snap, exist_ok=True)
            tag = _snap_tag(p)
            if p.en is not None:
                snap(en_doc, p.en.page, p.en.bbox,
                     os.path.join(lang_snap, f'{tag}_{lang}_EN.png'))
                n_snaps += 1
            if p.xx is not None:
                snap(doc, p.xx.page, p.xx.bbox,
                     os.path.join(lang_snap, f'{tag}_{lang}_XX.png'))
                n_snaps += 1
            elif p.status == '译文未匹配' and p.expect_y is not None:
                # 译文侧期望位置截图(缺失处上下文)
                e = p.en
                cx = (e.bbox[0] + e.bbox[2]) / 2 + (p.expect_dx or 0.0)
                hh = max(e.bbox[3] - e.bbox[1], 10.0)
                page = doc[e.page]
                r2 = fitz.Rect(cx - 50, p.expect_y - hh / 2 - 14,
                                cx + 50, p.expect_y + hh / 2 + 14) & page.rect
                if r2.width > 2 and r2.height > 2:
                    page.get_pixmap(clip=r2, dpi=SNAP_DPI).save(
                        os.path.join(lang_snap, f'{tag}_{lang}_XX@期望位置.png'))
                    n_snaps += 1
        problems = sum(1 for p in pairs if p.status != '一致' and p.status != '已并入聚合差异')
        n_sp = sum(1 for p in pairs if p.status == '一致' and '大位移' in p.note)
        r['n_sp'] = n_sp
        # 复核 PDF: 仅需人工确认的问题项(大位移/排版差异已自动判一致, 不需生成)
        marks = []
        for p in pairs:
            if p.status in ('一致', '已并入聚合差异'):
                continue
            if p.xx is not None:
                marks.append((p.xx.page, p.xx.bbox, f'{p.cp or "?"} CHECK', False))
            elif p.status == '译文未匹配' and p.expect_y is not None and p.en is not None:
                cx = (p.en.bbox[0] + p.en.bbox[2]) / 2 + (p.expect_dx or 0.0)
                hh = max(p.en.bbox[3] - p.en.bbox[1], 10.0)
                marks.append((p.en.page, (cx - 45, p.expect_y - hh / 2 - 8,
                                          cx + 45, p.expect_y + hh / 2 + 8),
                              f'{p.cp} MISSING?', True))
            else:
                marks.append((p.xx.page if p.xx else 0, (40, 40, 200, 80),
                              f'{p.cp or "?"} CHECK', False))
        if marks:
            os.makedirs(review_dir, exist_ok=True)
            build_review_pdf(path, os.path.join(review_dir, f'{lang}.pdf'), marks)
        # 收集英文侧标注(同一检查点多语言问题叠加)
        for p in pairs:
            if p.status in ('一致', '已并入聚合差异') or p.en is None:
                continue
            key = (p.en.page, round(p.en.yc, 1), round(p.en.bbox[0], 1))
            en_marks.setdefault(key, {'bbox': p.en.bbox, 'langs': set()})
            en_marks[key]['langs'].add(lang)
        print(f'  [{lang}] {fname}: 检查点{len(en_items)} 译文红字{r["n_xx"]} '
              f'问题项{problems} 大位移确认{n_sp} (截图{n_snaps}张 复核PDF{len(marks)}框)')
        doc.close()

    # 英文指示稿复核 PDF: 标注所有问题检查点位置, 供点击英文截图定位
    en_rows = []
    for key, info in sorted(en_marks.items()):
        page0, yc, x0 = key
        langs = ','.join(sorted(info['langs']))
        en_rows.append([page0, info['bbox'], f'CHECK ({langs})', True])
    if en_rows:
        os.makedirs(review_dir, exist_ok=True)
        build_review_pdf(args.base, os.path.join(review_dir, 'EN.pdf'), en_rows)

    report = os.path.join(out_dir, '数字校对报告.xlsx')
    build_excel(report, os.path.basename(args.base), en_items, results, snaps_dir, color_notes)
    report_html = os.path.join(out_dir, '数字校对报告.html')
    build_html(report_html, os.path.basename(args.base), en_items, results, snaps_dir)
    print(f'\n报告已生成: {report}')
    print(f'           {report_html}')
    print(f'截图目录:   {snaps_dir}')


if __name__ == '__main__':
    if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    main()
