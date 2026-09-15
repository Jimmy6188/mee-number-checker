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
import unicodedata
import base64
import bisect
import os
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from statistics import median

import fitz  # PyMuPDF
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
W_X_SAME = 0.3         # 值相同候选对的 x 权重: 图形/文本重排常致大幅水平位移,
                       # 垂直位置才是主判据(否则同值对因 dx 超阈被拆散, 反配邻近异值项);
                       # 0.3 可救回 dx~200pt 级重排, 仍拒绝 dx>330pt 的疑似错配
REWARD_SAME = -30.0    # 值相同奖励(负代价)
PENALTY_DIFF = 35.0    # 值不同惩罚(允许真实差异, 但需位置强吻合)
PENALTY_TOK = 35.0     # 数字个数不同惩罚
REWARD_CTX = 8.0       # 骨架上下文吻合奖励(每侧)
SP_WIN = 95.0          # 二级匹配: 插值期望位置窗口 (pt)
SP_DEV_TOL = 75.0      # 二级匹配: 实际位置与期望的最大偏差 (pt)
SNAP_PAD = 55          # 截图外扩边距 (pt)
SNAP_DPI = 150         # 截图分辨率
# 锚定/高亮颜色参数(客户标注颜色, 可调; 容差容忍偏色)
ANCHOR_RED = (1.0, 0.0, 0.0)      # 红框(客户 Square 批注)基准色
ANCHOR_FILL = (0.0, 1.0, 1.0)     # 高亮文字填充基准色(青)
COLOR_TOL = 0.08                  # 颜色容差


def _color_match(c: tuple, target: tuple) -> bool:
    """颜色近似匹配(绝对值容差), 兼容 0-1 浮点"""
    if c is None:
        return False
    return all(abs(a - b) <= COLOR_TOL for a, b in zip(c, target))

TOKEN_RE = re.compile(r'\d+(?:[.,]\d+)*')
# 语言无关 Token(单位/型号/符号): 用于数字身份指纹
FP_RE = re.compile(r'(N[·•.]?m|kgf[·•.]?cm|kgf/cm|MPa|kPa|mm²|mm|Hz|°C|psi|bar|kg/cm|m³|R290|R32|R410A?|MSZ-|MUZ-|MAC-|WPA|Wi-Fi|R29\d|R3\d|ø|±|×|∅|%)', re.I)

# 不进入问题项统计/截图/报告的状态(供统一排除)
IGNORED_STATUSES = ('已并入聚合差异', '非锚定区(忽略)', '目录条目(不校对)')

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
    fp: str = ''         # 语义指纹: 数字邻接的语言无关token(单位/型号/符号)
    hl_type: str = ''    # 高亮模式内容类型: num/code/ord/symstr/phrase

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
    locate_page: int | None = None    # 该问题在译文定位 PDF 中的页码(1基, 供 HTML 跳转)
    en_locate_page: int | None = None # 英文侧定位 PDF 页码

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


def _split_number_segments(sp: dict) -> list:
    """把 span 切分为数字 token 子 span(按字符宽度比例估算位置)。
    返回含数字 token 的子 span dict 列表(bbox 按 span 内字符位置线性插值)。"""
    text = sp["text"]
    x0, x1 = sp["bbox"][0], sp["bbox"][2]
    ws = x1 - x0
    n = len(text)
    if n <= 0 or ws <= 0:
        return []
    segs = []
    u = 0
    while u < n:
        m = TOKEN_RE.search(text, u)
        if not m:
            break
        s, e = m.start(), m.end()
        sx = x0 + ws * s / n
        ex = x0 + ws * e / n
        seg = dict(sp)
        seg["text"] = m.group(0)
        seg["bbox"] = (sx, sp["bbox"][1], ex, sp["bbox"][3])
        segs.append(seg)
        u = e
    return segs


def _cluster_spans(pno: int, spans: list, others: list) -> list[Item]:
    """行聚类 + 间隔聚合 -> 检查点。spans 为待聚合的 span 列表(已筛选),
    others 为同行非聚合 span(用于点号/逗号插入与骨架上下文)。"""
    items = []
    if not spans:
        return items
    spans = sorted(spans, key=lambda s: ((s["bbox"][1] + s["bbox"][3]) / 2, s["bbox"][0]))
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
                    # 两 span 之间的点号/逗号(黑色): 小数/千分位, 直接插入不分隔
                    dot = ''
                    for o in row_others:
                        if prev["bbox"][2] - 0.5 <= o[0] <= sp["bbox"][0] + 0.5 \
                                and o[1] <= sp["bbox"][0] + 0.5 \
                                and o[2] in ('.', ',', '．', '，', '·'):
                            dot = o[2]
                            break
                    if dot:
                        text += dot
                    elif gap > 0.5 and prev["text"].rstrip()[-1:].isdigit() and sp["text"].lstrip()[:1].isdigit():
                        # 数字边界: 两个独立数字不得直接连写(如相邻竖列 390|330 误粘成 390330,
                        # 双语粘不粘不对称 -> 假"真缺失"高风险), 必须插空格分 token
                        text += " "
                    else:
                        text += "" if gap < JOIN_GAP else " "
                text += sp["text"].strip()
                prev = sp
            x0 = min(s["bbox"][0] for s in g)
            y0 = min(s["bbox"][1] for s in g)
            x1 = max(s["bbox"][2] for s in g)
            y1 = max(s["bbox"][3] for s in g)
            left, right = _make_ctx(row_others, x0, x1)
            fp = extract_fp(text, left, right)
            items.append(Item(page=pno, text=text, bbox=(x0, y0, x1, y1),
                              yc=(y0 + y1) / 2, n_spans=len(g),
                              left_ctx=left, right_ctx=right, fp=fp))
    return items


def extract_fp(text: str, left_ctx: str = '', right_ctx: str = '', line_full: str = '', header_full: str = '') -> str:
    """从检查点文本 + 紧邻上下文 + 整行 + 上方表头文本提取语义指纹(单位/型号/符号)。
    返回归一化指纹串, 无则空。"""
    for probe in (text + ' ' + right_ctx, text, line_full + ' ' + right_ctx, header_full):
        if not probe:
            continue
        m = FP_RE.search(probe)
        if m:
            return m.group(1).lower().replace(' ', '')
    return ''


def anchor_zone_rects(anchor_doc: fitz.Document) -> list:
    """返回每页的红色 Square 框矩形(客户标注的校对区域)。
    校对对象 = 红框内的高亮数字; 因此用 Square 框过滤红字检查点。"""
    per_page = []
    for pno in range(anchor_doc.page_count):
        page = anchor_doc[pno]
        squares = []
        for a in page.annots():
            if a.type[1] == 'Square':
                try:
                    c = a.colors.get('stroke') or a.colors.get('fill')
                except AttributeError:
                    c = None
                if c and _color_match(c, ANCHOR_RED):
                    squares.append(fitz.Rect(a.rect))
        per_page.append(squares)
    return per_page


def in_anchor(item: Item, zones: list) -> bool:
    """检查点是否落在锚定区域(红色 Square 框)内"""
    if item.page >= len(zones):
        return False
    bb = fitz.Rect(item.bbox)
    return any(bb.intersects(s) for s in zones[item.page])


def extract_anchor_items(doc: fitz.Document) -> list[Item]:
    """从客户指示原稿提取锚定检查点: Square红框 ∩ 红色高亮矩形 -> 覆盖的数字 span。
    红框目前为红色 Square 批注; 高亮文字为青色(#00FFFF)填充矩形覆盖的 span。
    只有框内高亮的文字是校对对象。(后续若颜色变化, 调整阈值即可)"""
    items = []
    for pno in range(doc.page_count):
        page = doc[pno]
        squares = []
        for a in page.annots():
            if a.type[1] != 'Square':
                continue
            try:
                c = a.colors.get('stroke') or a.colors.get('fill')
            except AttributeError:
                c = None
            if c and _color_match(c, ANCHOR_RED):
                squares.append(fitz.Rect(a.rect))
        if not squares:
            continue
        # 高亮矩形: 青色填充(0,1,1)或高亮批注
        highlights = [fitz.Rect(dr['rect']) for dr in page.get_drawings()
                      if dr.get('fill') and _color_match(dr['fill'], ANCHOR_FILL)]
        highlights += [fitz.Rect(a.rect) for a in page.annots()
                       if a.type[1] == 'Highlight']
        in_sq = [h for h in highlights if any(h.intersects(s) for s in squares)]
        if not in_sq:
            continue
        d = page.get_text("dict")
        spans = []
        others = []
        for blk in d.get("blocks", []):
            if blk.get("type") != 0:
                continue
            for line in blk.get("lines", []):
                for sp in line.get("spans", []):
                    t = sp["text"].strip()
                    if not t:
                        continue
                    bb = fitz.Rect(sp["bbox"])
                    if any(h.intersects(bb) for h in in_sq):
                        # 锚定对象是数字: 把高亮 span 截取为「数字 token 子 span」
                        # (span 可能含整词如 "4 mm hexagonal wrench", 只取数字段)
                        segs = _split_number_segments(sp)
                        if segs:
                            spans.extend(segs)
                        else:
                            others.append([bb[0], bb[2], t, (bb[1] + bb[3]) / 2])
                    else:
                        others.append([bb[0], bb[2], t, (bb[1] + bb[3]) / 2])
        if spans:
            items.extend(_cluster_spans(pno, spans, others))
    return items


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
                # 脚注索引必须紧邻星号词(gap≤3pt, 实测真脚注 gap≈0);
                # 同行远处的红数字(如图形尺寸 350, gap 50+pt)是检查点, 不得误杀
                near_prev = (i > 0 and sp["bbox"][0] - row[i - 1][0]["bbox"][2] <= 3.0)
                is_footnote = (near_prev and prev_txt.endswith('*')
                               and re.fullmatch(r'\d{1,3}', t) is not None)
                if is_footnote:
                    continue
                # 上标字符(²/³/¹/ⁿ): 单位的一部分(如 kgf/cm²), 非独立数字, 不作为检查点
                if t in ('²', '³', '¹', 'ⁿ', '⁴', '⁵', '⁶', '⁷', '⁸', '⁹', '⁰') or re.fullmatch(r'[²³¹ⁿ⁴⁵⁶⁷⁸⁹⁰]+', t):
                    continue
                spans.append(sp)
                if color_counter is not None:
                    color_counter[f'#{sp["color"]:06x}'] += 1
        # 行聚类(红色): 由 _cluster_spans 统一处理(行聚类+间隔聚合+点号/逗号+骨架)
        others = [[o["bbox"][0], o["bbox"][2], o["text"].strip(),
                   (o["bbox"][1] + o["bbox"][3]) / 2]
                  for o, red in raw if not red]
        items.extend(_cluster_spans(pno, spans, others))
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
    - 所有检查点统一参与全局分配, 不硬锁定锚点; "放弃分配"以 REJECT_COST 边表达(分配后丢弃),
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

    # 2) 指纹优先配对: 同指纹(单位/型号/符号) + 值相同 -> 直接锁定, 不受坐标漂移影响
    #    仅当该(指纹,值)在英文/译文两侧都唯一时锁定, 避免抢配(如两个2470同指纹)。
    COST_BIG = 1e6
    pairs: list[Pair] = []
    matched: list[tuple[Item, Item]] = []
    locked_en: set[int] = set()
    locked_xx: set[int] = set()
    # 统计(指纹, 值)与(值)出现次数
    from collections import Counter as _C
    en_val_cnt = _C(TOKEN_RE.findall(e.text)[0] for e in en_items if TOKEN_RE.findall(e.text))
    xx_val_cnt = _C(TOKEN_RE.findall(x.text)[0] for x in xx_items if TOKEN_RE.findall(x.text))
    en_fp_cnt = _C((e.fp, TOKEN_RE.findall(e.text)[0] if TOKEN_RE.findall(e.text) else '') for e in en_items if e.fp)
    xx_fp_cnt = _C((x.fp, TOKEN_RE.findall(x.text)[0] if TOKEN_RE.findall(x.text) else '') for x in xx_items if x.fp)
    for i, e in enumerate(en_items):
        if not e.fp or i in locked_en:
            continue
        etoks = TOKEN_RE.findall(e.text)
        if not etoks:
            continue
        val = etoks[0]
        key = (e.fp, val)
        if en_fp_cnt[key] > 1 or xx_fp_cnt.get(key, 0) > 1 or en_val_cnt[val] > 1 or xx_val_cnt[val] > 1:
            continue          # 不唯一 -> 留给匈牙利(避免两个2470同指纹/同值抢配)
        for j, x in enumerate(xx_items):
            if j in locked_xx:
                continue
            if x.fp and e.fp == x.fp and value_same(e.text, x.text):
                e2, x2 = en_items[i], xx_items[j]
                status, note = compare(e2.text, x2.text)
                dy = x2.yc - e2.yc
                conf = 'high' if abs(dy) <= CONF_DY_HIGH else 'medium'
                if '大位移' not in note and '指纹锁定' not in note:
                    note = (note + '; ' if note else '') + f'指纹锁定({e2.fp})'
                pairs.append(Pair(cp='', page=page1, en=e2, xx=x2,
                                  y_off=round(dy, 1), conf=conf,
                                  status=status, note=note))
                matched.append((e2, x2))
                locked_en.add(i)
                locked_xx.add(j)
                break

    # 3) 全局最优一对一分配(仅未锁定的检查点)
    free_en = [i for i in range(n) if i not in locked_en]
    free_xx = [j for j in range(m) if j not in locked_xx]
    if free_en and free_xx:
        fe, fx = len(free_en), len(free_xx)
        # 剪枝边/超阈边统一用 REJECT_COST 占位: 分配后 >=REJECT_COST 一律丢弃,
        # 等价于"放弃分配"。若用 COST_BIG 占位, 匈牙利为避开 1e6 会优先牺牲
        # 同样被丢弃的 70+ 边, 扭曲全局最优(如把 -21.5 的同值项让给垃圾分配)。
        cost_f = [[REJECT_COST] * fx for _ in range(fe)]
        for a, i in enumerate(free_en):
            for b, j in enumerate(free_xx):
                e, x = en_items[i], xx_items[j]
                dy = abs((x.yc - e.yc) - y_med)
                dx = abs((x.xc - e.xc) - x_med)
                same = value_same(e.text, x.text)
                pdy = PRUNE_DY_SAME if same else prune_dy
                pdx = PRUNE_DX_SAME if same else prune_dx
                if dy > pdy or dx > pdx:
                    continue
                c = dy + (W_X_SAME if same else W_X) * dx + (REWARD_SAME if same else
                                     (PENALTY_TOK if token_count(e.text) != token_count(x.text)
                                      else PENALTY_DIFF))
                if e.left_ctx and x.left_ctx and e.left_ctx.lower() == x.left_ctx.lower():
                    c -= REWARD_CTX
                if e.right_ctx and x.right_ctx and e.right_ctx.lower() == x.right_ctx.lower():
                    c -= REWARD_CTX
                cost_f[a][b] = min(c, REJECT_COST)
        ri, cj = _hungarian(cost_f)
        for a, b in zip(ri, cj):
            if cost_f[a][b] >= REJECT_COST or cost_f[a][b] >= COST_BIG:
                continue
            i, j = free_en[a], free_xx[b]
            e, x = en_items[i], xx_items[j]
            dy = abs((x.yc - e.yc) - y_med)
            dx = abs((x.xc - e.xc) - x_med)   # 必须重算: 否则会泄漏矩阵构建循环的残留 dx
            same = value_same(e.text, x.text)
            if same:
                conf = 'high' if dy <= CONF_DY_HIGH else 'medium'
            elif dy <= CONF_DY_MED:
                conf = 'medium'      # 值不同但位置强吻合: 真实差异典型形态
            else:
                conf = 'low'
            status, note = compare(e.text, x.text)
            if same and status == '一致' and dy + W_X * dx + REWARD_SAME >= REJECT_COST:
                # 旧代价本会被拒绝、因同值降 dx 权重才配上的对: 位移异常大,
                # 保留大位移抽查档(finalize 转低风险), 不静默绿掉
                note = (note + '; ' if note else '') + f'大位移(一级配对): y偏差{dy:.0f}pt x偏差{dx:.0f}pt, 抽查项'
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


def owner_en_of(xx_id: int, owner: dict):
    """返回占用该 xx 项的 pair 的 en (None 表示译文多出项或未占用)。"""
    p = owner.get(xx_id)
    return p.en if p is not None else None


def resolve_aggregation(pairs: list[Pair], xx_items: list | None = None):
    """聚合差异合并(双向): 满足四条件的「待人工+未匹配」或「待人工+多出」成对
    -> 待人工项重标为「疑聚合差异」单条, 残项标为「已并入聚合差异」不重复计数。
    方向一(译文聚合/英文拆散): 待人工项译文token>英文token, 残为未匹配邻项;
    方向二(英文聚合/译文拆散): 待人工项英文token>译文token, 残为多出邻项。
    四条件(全过才合并): 同页; 数字个数不等; 差值token与残项token完全一致(多集差);
    位置邻近(<40pt)且残项唯一。任一不满足保持原状态(转人工), 不动摇底线。"""
    # 第零遍: 已配对但译文粘连注释标号前缀(如竖排上标 8 + 102 -> '8 102'):
    # 译文 tokens = 1~2位纯数字前缀 + 英文值完整后缀, 且 dy≤25 -> 判一致并备注。
    # 必须跑在方向一/二之前: 否则前缀会被当成邻项数字卷入聚合合并,
    # 把同页真正的邻项检查点误标"已并入"而掩掉它的独立问题。
    for p in pairs:
        if p.en is None or p.xx is None or p.status not in ('待人工', '不一致'):
            continue
        et = TOKEN_RE.findall(p.en.text)
        xt = TOKEN_RE.findall(p.xx.text)
        if not et or len(xt) <= len(et) or len(xt) - len(et) > 2:
            continue
        if abs(p.xx.yc - p.en.yc) > 25:
            continue
        for dh in range(1, len(xt) - len(et) + 1):
            pre = xt[:dh]
            if xt[dh:] != et or not all(re.fullmatch(r'\d{1,2}', t) for t in pre):
                continue
            # 守卫: 前缀值须在同页其它译文项里另有出现 -> 才是冗余注释标号;
            # 若同页再无此值, 前缀更可能是邻项检查点的真数字(真聚合), 交方向一/二。
            redundant = any(
                it is not p.xx and it.page == p.xx.page
                and any(_num_eq(v, t) for v in pre for t in TOKEN_RE.findall(it.text))
                for it in (xx_items or []))
            if not redundant:
                continue
            p.status = '一致'
            p.note = (f"译文值含额外注释标号前缀 {' '.join(pre)}(疑注释编号粘连), "
                      f"数值 {p.en.text} 与译文一致; 建议人工过一眼")
            break
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

    # 分支B: 译文未匹配(未配对), 但期望位置附近的候选是粘连项
    # (译文 tokens = 短数字前缀 + 英文值, 同指纹) -> 从译文多出项收养, 补配判一致并备注
    owner = {id(p.xx): p for p in pairs if p.xx is not None}
    # 只排除被正常配对(en非None)占用的项; 被"译文多出"(en=None)占用的项允许收养
    used_ids = {id(xx) for xx in owner if owner_en_of(xx, owner) is not None}
    for p in pairs:
        if p.status != '译文未匹配' or p.en is None:
            continue
        et = TOKEN_RE.findall(p.en.text)
        if not et:
            continue
        for it in xx_items:
            if it.page != p.en.page or id(it) in used_ids:
                continue
            if abs(it.yc - p.en.yc) > 25:
                continue
            if p.en.fp and it.fp and it.fp != p.en.fp:
                continue
            xt = TOKEN_RE.findall(it.text)
            if len(xt) <= len(et) or len(xt) - len(et) > 2:
                continue
            done = False
            for dh in range(0, len(xt) - len(et) + 1):
                if xt[dh:dh + len(et)] == et and all(re.fullmatch(r'\d{1,2}', t) for t in xt[:dh]):
                    pre = xt[:dh]
                    old = owner.get(id(it))
                    if old is not None and old.en is None:
                        # 原"译文多出"项收养后并入本检查点, 不再单独计问题
                        old.status = '已并入聚合差异'
                        old.note = f'已并入检查点 {p.cp}(译文粘连注释标号), 不再单独计为问题项'
                        old.xx = None
                    p.xx = it
                    p.y_off = round(it.yc - p.en.yc, 1)
                    p.conf = 'medium'
                    p.status = '一致'
                    p.note = (f"译文值含额外注释标号前缀 {' '.join(pre)}(疑注释编号粘连), "
                              f"数值 {p.en.text} 与译文一致; 建议人工过一眼")
                    used_ids.add(id(it))
                    done = True
                    break
            if done:
                break
    return pairs

# ---------------- 截图 ----------------
def snap(doc: fitz.Document, pno: int, bbox, path: str):
    page = doc[pno]
    r = fitz.Rect(bbox)
    r = fitz.Rect(r.x0 - SNAP_PAD, r.y0 - SNAP_PAD, r.x1 + SNAP_PAD, r.y1 + SNAP_PAD) & page.rect
    pix = page.get_pixmap(clip=r, dpi=SNAP_DPI)
    pix.save(path)


def snap_marked(page, clip, rect, path):
    """截图并在 rect 处叠黄底红框(与复核 PDF 同色同语义), 供期望位置/同位置参考截图.
    框烘焙进 PNG: HTML 与 Excel 共用, 位置由 clip 坐标换算保证精确."""
    pix = page.get_pixmap(clip=clip, dpi=SNAP_DPI)
    try:
        import io
        from PIL import Image, ImageDraw
        img = Image.open(io.BytesIO(pix.tobytes('png'))).convert('RGBA')
        z = SNAP_DPI / 72.0   # PDF pt -> 像素缩放
        box = ((rect.x0 - clip.x0) * z, (rect.y0 - clip.y0) * z,
               (rect.x1 - clip.x0) * z, (rect.y1 - clip.y0) * z)
        ov = Image.new('RGBA', img.size, (0, 0, 0, 0))
        d = ImageDraw.Draw(ov)
        d.rectangle(box, fill=(255, 235, 59, 100), outline=(207, 33, 46, 255), width=3)
        img = Image.alpha_composite(img, ov).convert('RGB')
        img.save(path, 'PNG')
    except Exception:
        pix.save(path)   # PIL 异常降级为无框截图


# ---------------- Excel 报告 ----------------
FILL = {
    '高风险':   PatternFill('solid', fgColor='C00000'),
    '中风险':   PatternFill('solid', fgColor='FFC000'),
    '低风险':   PatternFill('solid', fgColor='4A90D9'),
    '不一致':   PatternFill('solid', fgColor='C00000'),
    '需复核':   PatternFill('solid', fgColor='FFC000'),
    '格式差异': PatternFill('solid', fgColor='ED7D31'),
    '待人工':   PatternFill('solid', fgColor='FFC000'),
    '疑聚合差异': PatternFill('solid', fgColor='DAA520'),
    '译文未匹配': PatternFill('solid', fgColor='FFC000'),
    '译文多出':  PatternFill('solid', fgColor='FFC000'),
    '已并入聚合差异': PatternFill('solid', fgColor='BFBFBF'),
    '非锚定区(忽略)': PatternFill('solid', fgColor='D9D9D9'),
    '目录条目(不校对)': PatternFill('solid', fgColor='BFBFBF'),
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
    headers = ['文件名', '语言', '英文检查点数', '译文红字项', '一致(通过)',
               '低风险(大位移确认)', '中风险(需复核)', '高风险(不一致)', '结论']
    ws.append(headers)
    tot = {k: 0 for k in ['一致', '低风险', '中风险', '高风险']}
    tot_sp = 0
    for r in results:
        cnt = {k: 0 for k in tot}
        for p in r['pairs']:
            if p.status in cnt:
                cnt[p.status] += 1
        for k in tot:
            tot[k] += cnt[k]
        tot_sp += r['n_sp']
        problems = cnt['高风险'] + cnt['中风险']
        _off = r.get('offbox', 0)
        concl = ('✓ 通过' if problems == 0 else
                 f'必办{cnt["高风险"]} · 复核{cnt["中风险"]}'
                 + (f' · 抽查{cnt["低风险"]}' if cnt['低风险'] else '')
                 + (f' · 框外差异{_off}' if _off else ''))
        ws.append([r['file'], r['lang'], len(en_items), r['n_xx'],
                   cnt['一致'], r['n_sp'], cnt['中风险'], cnt['高风险'], concl])
        if problems:
            for c in range(1, len(headers) + 1):
                ws.cell(row=ws.max_row, column=c).fill = PatternFill('solid', fgColor='FFF2CC')
            ws.cell(row=ws.max_row, column=len(headers)).font = RED_FONT
    ws.append(['总计', '', len(en_items) * len(results), '', tot['一致'], tot_sp,
               tot['中风险'], tot['高风险'], ''])
    for c in range(1, len(headers) + 1):
        ws.cell(row=ws.max_row, column=c).font = Font(bold=True)
    style_header(ws, len(headers))
    for c, w in zip(range(1, len(headers) + 1), [34, 8, 13, 11, 10, 14, 14, 14, 12]):
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
            has_snap = p.status != '一致' and p.status not in IGNORED_STATUSES
            stag = _snap_tag(p)   # 与实际存图名对齐(多出项是 XX@页~y, 非 'XX')
            se = os.path.join('snaps', r['lang'], f"{stag}_{r['lang']}_EN.png") if p.en and has_snap else ''
            if not se and p.en is None and p.xx is not None and has_snap:
                se = os.path.join('snaps', r['lang'], f"{stag}_{r['lang']}_EN@同位置参考.png")
            sx = os.path.join('snaps', r['lang'], f"{stag}_{r['lang']}_XX.png") if p.xx and has_snap else ''
            sx2 = os.path.join('snaps', r['lang'], f"{stag}_{r['lang']}_XX@期望位置.png") \
                if (has_snap and p.xx is None and p.expect_y is not None) else ''
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
        return {'一致': '✓', '高风险': '✗', '中风险': '?', '低风险': '±', '格式差异': 'F', '待人工': '?',
                '疑聚合差异': '≈', '已并入聚合差异': '·', '目录条目(不校对)': 'T',
                '译文未匹配': '∅', '译文多出': '＋', '非锚定区(忽略)': '·'}.get(p.status, '?')
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
            for st, m in [('高风险', '✗'), ('中风险', '?'), ('低风险', '±'), ('格式差异', 'F'), ('待人工', '?'),
                          ('疑聚合差异', '≈'), ('译文未匹配', '∅'), ('目录条目(不校对)', 'T')]:
                if v == m:
                    ws.cell(row=ws.max_row, column=c).fill = FILL[st]
            if v == '·':
                ws.cell(row=ws.max_row, column=c).fill = FILL['已并入聚合差异']
    if xx_only:
        ws.append([])
        ws.append(['— 以下为译文多出项 —'])
        for lang, p in xx_only:
            if p.xx is None:
                continue   # 已被分支B收养并入对应检查点, 不再单列
            ws.append([f'(译文多出)', p.xx.text, *[lang if l == lang else '' for l in langs]])
    # 矩阵下方图例
    ws.append([])
    ws.append(['图例'])
    legend = [
        ('✓', '一致：匹配成功且归一化数值相等，无风险(绿)'),
        ('✗', '高风险：数值不同或真缺失，必须处理(红)'),
        ('?', '中风险：无法确认对应(串位/聚合疑点/跨页等), 建议核对(黄)'),
        ('±', '低风险：大位移确认(值相同但位置大幅移动), 抽查项(蓝)'),
        ('F', '格式差异：数值相同但书写不同(如 2.5 vs 2,5)'),
        ('?', '待人工：匹配存在歧义或数字个数不同'),
        ('∅', '译文未匹配：英文有此检查点但译文未找到红字(疑漏标)'),
        ('≈', '疑聚合差异：译文将相邻检查点的数字聚为一体，已合并为单条'),
        ('·', '已并入聚合差异：上述合并项的原未匹配项，不重复计数'),
        ('＋', '译文多出：译文有红字但英文无对应'),
        ('T', '目录条目(不校对)：点线目录行的章节号/页码, 随各语言重排可变, 不计入问题(备注保留差异供抽查)'),
        ('', '底色说明：绿=一致 蓝=低风险 红=高风险 黄=中风险/待人工/未匹配 暗金=疑聚合 灰=已并入'),
    ]
    for sym, desc in legend:
        ws.append([sym, desc])
        r = ws.max_row
        cell = ws.cell(row=r, column=1)
        if sym == '✓':
            ws.cell(row=r, column=1).fill, ws.cell(row=r, column=2).fill = FILL['一致'], FILL['一致']
        elif sym == '✗':
            ws.cell(row=r, column=1).fill, ws.cell(row=r, column=2).fill = FILL['高风险'], FILL['高风险']
        elif sym == '±':
            ws.cell(row=r, column=1).fill, ws.cell(row=r, column=2).fill = FILL['低风险'], FILL['低风险']
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
        ['状态定义(风险分级)'],
        ['使用建议', '默认只看高风险(必办); 中/低风险仅在时间允许时浏览——中/低均为"值级可对账、仅布局/位置差异"项, 跳过不丢真错(框外差异除外, 见下)'],
        ['一致', '能确认对应且数值等价, 无风险(绿)'],
        ['低风险', '大位移确认: 值相同但译排版重排导致红字大幅移动; 抽查项(蓝)'],
        ['中风险', '无法确认对应(可能串位/聚合/跨页位移/图内), 保守转人工; 黄色, 建议核对。串位/聚合/图内降档均要求差值在页内可对账(值未丢失), 仅布局变化'],
        ['高风险', '值级不可对账必报: 确认对应但数值不同(真差异), 或值在对面完全不存在(真缺失/多余); 图内红字若值在页内不存在也判高风险; 红色, 必须处理'],
        ['已并入聚合差异', '聚合/锚定合并后的原项, 不重复计为问题 (灰色, 供追溯)'],
        ['非锚定区(忽略)', '不在客户红框内, 非校对对象 (灰, 不统计); 若框外两侧已配对且值不同, 备注标"框外值差异"留痕并在汇总/页头计数'],
        ['目录条目(不校对)', '点线目录行的章节号/页码, 页码随各语言重排可变, 不作数值校对 (灰, 不统计, 备注留差异)'],
        [''],
        ['自动判定底线'],
        ['原则', '不确定不判一致; 自动判一致仅有: ①高/中置信+值相等 ②大位移三判据 ③指纹锁定(唯一) ④编号列对齐 ⑤点逗写法'  ],
        ['大位移三判据', '唯一值相同候选 + 双侧已配对邻居插值支撑 + 偏差≤阈值, 全部满足才自动判, 其余一律转需复核'],
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
    for rw, key in [(6, '状态定义'), (13, '自动判定底线'), (16, '置信度'), (20, '使用方法')]:
        for r in range(1, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if v and v.startswith(key):
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
.filters button.btn-high{background:#cf222e;color:#fff;border-color:#cf222e}
.filters button.btn-mid{background:#e3b341;color:#fff;border-color:#e3b341}
.filters button.btn-low{background:#0969da;color:#fff;border-color:#0969da}
main{padding:20px 32px 60px}
section.lang{margin-bottom:28px;background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);overflow:hidden}
section.lang>h2{font-size:15px;padding:12px 20px;background:#fafbfc;border-bottom:1px solid #e1e4e8;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
section.lang>h2 .badge{font-size:12px;padding:2px 10px;border-radius:10px;color:#fff}
section.lang>h2 .fname{color:#57606a;font-size:12px;font-weight:400}
.allpass{padding:16px 20px;color:#1a7f37}
.item{border-left:4px solid #d0d7de;padding:14px 20px;border-bottom:1px solid #eee}
.item:last-child{border-bottom:none}
.item.st-高风险{border-left-color:#cf222e}.item.st-中风险{border-left-color:#e3b341}.item.st-低风险{border-left-color:#0969da}.item.st-不一致{border-left-color:#cf222e}.item.st-需复核{border-left-color:#e3b341}.item.st-格式差异{border-left-color:#fb8500}
.item.st-待人工,.item.st-译文未匹配,.item.st-译文多出{border-left-color:#e3b341}
.item.st-疑聚合差异{border-left-color:#b8860b}
.item.st-大位移{border-left-color:#0969da}
.item .head{display:flex;gap:10px;align-items:baseline;flex-wrap:wrap;margin-bottom:10px}
.item .cp{font-weight:600;font-size:13px}
.badge{display:inline-block;font-size:12px;padding:2px 10px;border-radius:10px;color:#fff}
.b-高风险{background:#cf222e}.b-中风险{background:#e3b341}.b-低风险{background:#0969da}.b-不一致{background:#cf222e}.b-需复核{background:#e3b341}.b-格式差异{background:#fb8500}.b-待人工,.b-译文未匹配,.b-译文多出{background:#bf8700}.b-大位移{background:#0969da}.b-一致{background:#1a7f37}.b-疑聚合差异{background:#b8860b}
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

    status_key = {'高风险': '高风险', '中风险': '中风险', '低风险': '低风险'}
    cnt = Counter()
    for r in results:
        for p in r['pairs']:
            if p.status in IGNORED_STATUSES:
                continue                 # 合并项不重复计数
            if p.status in status_key:
                cnt[p.status] += 1
    n_items = len(en_items) * len(results)
    n_ok = n_items - sum(cnt.values())

    def snap_html(lang: str, p: Pair, kind: str) -> str:
        if p.en is None and p.xx is None:
            return ''
        tag = _snap_tag(p)
        if kind == 'EN':
            if p.en is not None:
                f = os.path.join(snaps_dir, lang, f'{tag}_{lang}_EN.png')
                cap = f'英文指示稿 · 值 {_h.escape(p.en.text)}'
            else:
                f = os.path.join(snaps_dir, lang, f'{tag}_{lang}_EN@同位置参考.png')
                cap = '英文指示稿 · 同位置参考(无对应红字)'
        elif kind == 'XX' and p.xx is not None:
            f = os.path.join(snaps_dir, lang, f'{tag}_{lang}_XX.png')
            cap = f'{LANG_NAMES.get(lang, lang)} · 值 {_h.escape(p.xx.text)}'
        else:
            f = os.path.join(snaps_dir, lang, f'{tag}_{lang}_XX@期望位置.png')
            cap = f'{LANG_NAMES.get(lang, lang)} · 期望位置(未找到红字)'
        img = _b64(f)
        if img is None:
            # 缺图占位, 不再渲染 src="None" 碎图
            return (f'<figure><img class="missing" alt="截图缺失">'
                    f'<figcaption>{cap} · 截图缺失</figcaption></figure>')
        # 需人工确认项可点击打开定位 PDF(每问题一页, 打开即定位)
        need_review = p.status not in ('一致',) + IGNORED_STATUSES
        if need_review:
            href = None
            if kind == 'EN':
                if p.en is not None:
                    pg = p.en_locate_page or (p.en.page + 1)
                    href = f'复核PDF/EN.pdf#page={pg}'   # 英文截图 -> 英文定位 PDF
                elif p.en_locate_page:
                    pg = p.en_locate_page
                    href = f'复核PDF/EN.pdf#page={pg}'   # 多出项英文侧 -> 英文稿同位置参考页
            elif p.xx is not None:
                pg = p.locate_page or (p.xx.page + 1)
                href = f'复核PDF/{lang}.pdf#page={pg}'  # 译文截图 -> 对应语言定位 PDF
            else:
                pg = p.locate_page or (p.en.page + 1)
                href = f'复核PDF/{lang}.pdf#page={pg}'  # 期望位置 -> 对应语言定位 PDF
            inner = (f'<figure class="linkable" style="cursor:pointer" '
                     f'title="点击打开定位PDF(定位页{pg})"><img src="{img}">'
                     f'<figcaption>{cap} · 点击定位</figcaption></figure>')
            if href:
                return f'<a href="{href}" target="_blank">{inner}</a>'
            return inner
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
        f'<div class="stat info"><b>{cnt["低风险"]}</b><span>低风险(大位移)</span></div>',
        f'<div class="stat warn"><b>{cnt["中风险"]}</b><span>中风险(需复核)</span></div>',
        f'<div class="stat err"><b>{cnt["高风险"]}</b><span>高风险(不一致)</span></div>',
        f'<div class="stat"><b>{len(en_items)}×{len(results)}</b><span>检查点总数</span></div>',
        '</div>',
    ]
    _off_total = sum(r.get('offbox', 0) for r in results)
    if _off_total:
        parts.insert(-1, f'<div class="stat"><b>{_off_total}</b><span>框外值差异(不计入,明细留痕)</span></div>')

    # 筛选按钮(按风险颜色)
    btns = [('all', '全部', '', sum(cnt.values()))]
    for st, label, col in [('高风险', '高风险', 'btn-high'), ('中风险', '中风险', 'btn-mid'), ('低风险', '低风险', 'btn-low')]:
        if cnt[st]:
            btns.append((st, label, col, cnt[st]))
    parts.append('<div class="filters"><button class="active" data-f="all" onclick="setFilter(this.dataset.f)">'
                 f'全部 ({btns[0][3]})</button>')
    for f, label, col, n in btns[1:]:
        parts.append(f'<button class="{col}" data-f="{f}" onclick="setFilter(this.dataset.f)">{label} ({n})</button>')
    parts.append('</div><main>')

    # 语言区块
    for r in results:
        lang = r['lang']
        lname = LANG_NAMES.get(lang, lang)
        probs = [p for p in r['pairs'] if p.status not in ('一致',) + IGNORED_STATUSES]
        sps = []
        problems = len(probs)
        nh = sum(1 for p in probs if p.status == '高风险')
        nm = sum(1 for p in probs if p.status == '中风险')
        nl = problems - nh - nm
        _off = r.get('offbox', 0)
        concl = ('✓ 通过' if nh + nm == 0 else
                 f'必办{nh} · 复核{nm}'
                 + (f' · 抽查{nl}' if nl else '')
                 + (f' · 框外差异{_off}' if _off else ''))
        bcolor = '#1a7f37' if nh + nm == 0 else ('#cf222e' if nh else '#bf8700')
        parts.append(f'<section class="lang" id="lang-{lang}"><h2>'
                     f'{lname} <span class="fname">{_h.escape(r["file"])} · 检查点{len(en_items)} '
                     f'· 一致{len(en_items) - problems}</span>'
                     f'<span class="badge" style="background:{bcolor}">{concl}</span></h2>')
        if not probs:
            parts.append('<div class="allpass">✓ 本语言全部通过, 无需处理</div></section>')
            continue
        for p in probs:
            st = p.status
            ev = p.en.text if p.en else '(无)'
            xv = p.xx.text if p.xx else '(缺失)'
            same = p.status == '低风险'
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
// 默认只看高风险(存在时); 中/低/框外靠筛选按钮切换浏览
window.addEventListener('DOMContentLoaded',function(){
  if(document.querySelector('.item[data-status="高风险"]'))setFilter('高风险');
});
"""


# ---------------- 锚定与编号列 ----------------
def detect_entry_columns(items: list, x_tol: float = 15.0, min_n: int = 3) -> list:
    """识别页内科标号列(目录/列表编号): 单数字 1-20、右上下文以 '.'或')'开头、
    x 同列(±x_tol)、按 y 递增、至少 min_n 个。
    返回 [(x, [(yc, num, Item)...]), ...] 的编号列列表(按 x 一列一列)。"""
    cands = []
    for it in items:
        t = it.text.strip()
        if re.fullmatch(r'\d{1,2}', t) and 1 <= int(t) <= 20:
            rc = it.right_ctx.strip()
            if rc.startswith('.') or rc.startswith(')'):
                cands.append((it.xc, it.yc, int(t), it))
    # 按 x 列分组
    cols = {}
    for xc, yc, num, it in cands:
        key = None
        for k in cols:
            if abs(k - xc) <= x_tol:
                key = k
                break
        if key is None:
            key = xc
            cols[key] = []
        cols[key].append((yc, num, it))
    out = []
    for xc, group in cols.items():
        group.sort(key=lambda t: t[0])
        if len(group) >= min_n:
            out.append((xc, group))
    return out


def align_entry_columns_pairs(pairs: list, en_items: list, xx_items: list) -> int:
    """编号列对齐: 对每页, 英文编号列与译文编号列按编号值对齐(1↔1, 2↔2...),
    值相同的编号对 -> 该 pair 若未判一致则改判'一致(编号列对齐)'。
    返回修改数。"""
    changed = 0
    en_by_page = {}
    for it in en_items:
        en_by_page.setdefault(it.page, []).append(it)
    xx_by_page = {}
    for it in xx_items:
        xx_by_page.setdefault(it.page, []).append(it)
    # en 检查点 -> pair 映射(按对象id)
    en_pair = {id(p.en): p for p in pairs if p.en is not None}
    for pno in set(en_by_page) & set(xx_by_page):
        en_cols = detect_entry_columns(en_by_page[pno])
        xx_cols = detect_entry_columns(xx_by_page[pno])
        if not en_cols or not xx_cols:
            continue
        # 每列两两? 只取第一列(x 最小的编号列, 目录/列表通常在左侧)
        ex, eg = min(en_cols, key=lambda c: c[0])
        xx, xg = min(xx_cols, key=lambda c: c[0])
        en_nums = {n: it for _, n, it in eg}
        xx_nums = {n: it for _, n, it in xg}
        for n in sorted(set(en_nums) & set(xx_nums)):
            e_it, x_it = en_nums[n], xx_nums[n]
            p = en_pair.get(id(e_it))
            if p is None or p.xx is None:
                continue
            if p.status != '一致' and p.status not in IGNORED_STATUSES:
                p.status = '一致'
                p.note = ('编号列对齐: 编号 %d(%s) ↔ 译文编号 %d(%s), 排版位移已按序对齐; '
                          % (n, e_it.text, n, x_it.text)) + p.note
                changed += 1
    return changed


def detect_entry_columns_spans(doc: fitz.Document, pno: int) -> list:
    """span 级编号列识别: 同行首编号('1.'/'7)') x同列, 按 y 递增, 至少3个。
    返回 [(x, [(yc, num, span_bbox_x0)...])]"""
    page = doc[pno]
    d = page.get_text("dict")
    cands = []
    for blk in d.get("blocks", []):
        if blk.get("type") != 0:
            continue
        for line in blk.get("lines", []):
            spans = sorted(line.get("spans", []), key=lambda s: s["bbox"][0])
            text = "".join(sp["text"] for sp in spans)
            m = re.match(r'^\s*(\d{1,2})\s*[.)]', text)
            if not m:
                continue
            n = int(m.group(1))
            sp0 = spans[0] if spans else None
            if sp0 is None:
                continue
            yc = (sp0["bbox"][1] + sp0["bbox"][3]) / 2
            xc = sp0["bbox"][0]
            near = [sp for sp in spans if sp["bbox"][0] < xc + 4 and sp["text"].strip() == m.group(1)]
            if near:
                xc = near[0]["bbox"][0]
                yc = (near[0]["bbox"][1] + near[0]["bbox"][3]) / 2
            if 1 <= n <= 20:
                cands.append((xc, yc, n))
    # 按 x 列
    cols = {}
    for xc, yc, n in cands:
        key = None
        for k in cols:
            if abs(k - xc) <= 15:
                key = k
                break
        if key is None:
            key = xc
            cols[key] = []
        cols[key].append((yc, n))
    out = []
    for xc, grp in cols.items():
        grp.sort(key=lambda t: t[0])
        if len(grp) >= 3:
            out.append((xc, grp))
    return out


def align_entry_columns_spans(pairs: list, en_doc: fitz.Document, xx_doc: fitz.Document) -> int:
    """span级编号列对齐: 对每页, 英文编号列(编号值)与译文编号列按值对齐,
    值相同则把'不一致/待人工'且 y 在编号±8 内的检查点改判'一致(编号列对齐)'。"""
    changed = 0
    maxp = max(en_doc.page_count, xx_doc.page_count)
    for pno in range(maxp):
        en_cols = detect_entry_columns_spans(en_doc, pno)
        xx_cols = detect_entry_columns_spans(xx_doc, pno)
        if not en_cols or not xx_cols:
            continue
        ex, eg = min(en_cols, key=lambda c: c[0])
        xx, xg = min(xx_cols, key=lambda c: c[0])
        en_nums = {n: y for y, n in eg}
        xx_nums = {n: y for y, n in xg}
        for n in sorted(set(en_nums) & set(xx_nums)):
            ey, xy = en_nums[n], xx_nums[n]
            for p in pairs:
                if p.en is None or p.xx is None:
                    continue
                if p.status not in ('不一致', '待人工'):
                    continue
                # 必须: 英文检查点文本是纯编号 n, 且 x 在编号列附近(±10pt), 防止把同行的数据值误当编号
                if not re.fullmatch(r'%d' % n, p.en.text.strip()):
                    continue
                if not re.fullmatch(r'%d' % n, p.xx.text.strip()):
                    continue
                if abs(p.en.xc - ex) > 10 or abs(p.xx.xc - xx) > 10:
                    continue
                if p.en.page == pno and abs(p.en.yc - ey) <= 8 and p.xx.page == pno and abs(p.xx.yc - xy) <= 8:
                    p.status = '一致'
                    p.note = f'编号列对齐: 编号 {n} ({p.en.text}) ↔ 译文编号 {n} ({p.xx.text}), 排版位移已按序对齐'
                    changed += 1
    return changed


def attach_fp(items: list, doc: fitz.Document) -> None:
    """后处理: 为每个检查点补语义指纹(单位/型号/符号), 使用检查点所在行全文 + 页面上方表头。
    页面级: 表格单位常在表头(上方行), 因此指纹优先: 检查点行全文 -> 上方最近含单位的行。"""
    # 页级文本(按行): page -> [(yc, full_text)]
    page_lines = {}
    for pno in range(doc.page_count):
        rows = []
        for blk in doc[pno].get_text("dict").get("blocks", []):
            if blk.get("type") != 0:
                continue
            for line in blk.get("lines", []):
                ly = (line["bbox"][1] + line["bbox"][3]) / 2
                txt = "".join(sp["text"] for sp in line.get("spans", []))
                if txt.strip():
                    rows.append((ly, txt))
        rows.sort(key=lambda r: r[0])
        page_lines[pno] = rows
    for it in items:
        if it.fp:
            continue
        rows = page_lines.get(it.page, [])
        # 本行全文
        line_full = ''
        header_full = ''
        for ly, txt in rows:
            if abs(ly - it.yc) <= LINE_TOL + 4:
                line_full = txt
                break
        # 上方最近行(表头, y 小且在同一页顶部区域, 容差 60pt)
        for ly, txt in reversed(rows):
            if ly < it.yc - 8 and it.yc - ly < 60:
                header_full = txt
                break
        it.fp = extract_fp(it.text, it.left_ctx, it.right_ctx, line_full, header_full)
    return items


def _count_unpaired(items: list, vals: list, pno: int, fp: str, paired_ids: set) -> int:
    """同页同指纹同值、且未被配对的剩余候选数量。"""
    n = 0
    for it in items:
        if it.page != pno or id(it) in paired_ids:
            continue
        if fp and it.fp and it.fp != fp:
            continue
        itoks = TOKEN_RE.findall(it.text)
        if vals and any(any(_num_eq(v, t) for t in itoks) for v in vals):
            n += 1
    return n


def finalize_statuses(pairs: list, en_items: list, xx_items: list) -> int:
    """报告前状态归一化(保守):
      一致     : 配对上且值相同(含大位移/指纹/编号列/点逗写法, 备注保留)
      不一致   : 真差异/真缺失(对面无剩余候选值)
      需复核   : 无法确认(对面有未配对候选, 疑串位); 图内红字一律需复核
    """
    changed = 0
    # 已配对集合
    paired_xx = {id(p.xx) for p in pairs if p.xx is not None}
    paired_en = {id(p.en) for p in pairs if p.en is not None}
    for p in pairs:
        # 大位移确认(一致+大位移备注) -> 低风险(独立档, 抽查项)
        if p.status == '一致' and '大位移' in p.note:
            p.status = '低风险'
            continue
        if p.status in IGNORED_STATUSES or p.status == '一致':
            continue
        # 图内红字(图注/箭头/分数标记): 仅当英文值在译文同页可对账(位置失真型串位)才降需复核;
        # 值在页内根本不存在 = 真差异/真缺失, 图内也保持不一致(高风险必报)
        if p.en is not None and is_figure_red(p.en.text, p.en.left_ctx, p.en.right_ctx):
            vals = TOKEN_RE.findall(p.en.text)
            others = [it for it in xx_items if it.page == p.en.page
                      and (p.xx is None or it is not p.xx)]
            if vals and all(any(_num_eq(v, t) for it in others for t in TOKEN_RE.findall(it.text))
                            for v in vals):
                p.status = '需复核'
                p.note = '需复核(图内红字,值在页内可对账,图重排位置失真): ' + p.note
                changed += 1
                continue
        if p.status == '不一致':
            # 值确认不同 -> 保留不一致(即使低置信, 值不同是事实; 低置信仅表示位置存疑需人工复核坐标)
            continue
        if p.en is not None:
            vals = TOKEN_RE.findall(p.en.text)
            if p.xx is not None:
                p.status = '需复核'
                p.note = '需复核(有对应但无法确认): ' + p.note
                changed += 1
                continue
            # 译文未匹配: 期望位置(expect_y)±35pt 内, 同指纹同值的候选存在?
            # 存在 -> 串位(需复核); 不存在 -> 真缺失(不一致)
            def near_exists():
                # 1) 期望位置±20pt 内同指纹同值(未配对候选优先) -> 串位(需复核)
                if p.expect_y is not None and vals:
                    for it in xx_items:
                        if it.page != p.en.page or abs(it.yc - p.expect_y) > 20:
                            continue
                        if p.en.fp and it.fp and it.fp != p.en.fp:
                            continue
                        itoks = TOKEN_RE.findall(it.text)
                        if any(any(_num_eq(v, t) for t in itoks) for v in vals):
                            return True
                # 2) 同页有未配对剩余候选(值串位到别处) -> 需复核
                return _count_unpaired(xx_items, vals or [], p.en.page, p.en.fp, paired_xx) > 0
            if near_exists():
                p.status = '需复核'
                p.note = '需复核(值存在但未配到,疑串位): ' + p.note
            else:
                # 期望位置无值且无剩余候选: 真缺失(强信号) -> 不一致, 需人工最终确认
                p.status = '不一致'
                p.note = '真缺失(期望位置无对应值): ' + p.note
            changed += 1
        else:
            vals = TOKEN_RE.findall(p.xx.text)
            # 值在同页英文检查点存在(无论是否已配对) -> 疑串位(需复核)
            exists_same_page = any(
                it.page == p.xx.page and any(_num_eq(v, t) for t in TOKEN_RE.findall(it.text))
                for v in vals for it in en_items)
            if exists_same_page:
                p.status = '需复核'
                p.note = '需复核(值存在于英文,疑串位): ' + p.note
            else:
                p.status = '不一致'
                p.note = '真多余(值在英文不存在): ' + p.note
            changed += 1
    # 输出状态映射: 内部旧名 -> 四档风险
    for p in pairs:
        if p.status == '不一致':
            p.status = '高风险'
        elif p.status == '需复核':
            p.status = '中风险'
    return changed


# ---------------- 目录条目(点线条)结果级处理 ----------------
# 目录行形如 "1.......7"(章节号+点线+页码), 红字提取后被水平行聚类粘成一项。
# 点线判定只以英文侧(校对基准)为准: 译文侧点线数量/跨span不稳定, 不得据此删项过滤
# (旧方案双侧各自正则过滤 -> 英文删译文未删 -> 成批假"真缺失/真多余"高风险)。
# 对称性由"同页 + y 对齐英文目录条目行"事实兑底。
TOC_LINE_RE = re.compile(r'\.{4,}')      # 点线: ≥4 个连续点
TOC_PAGE_MIN = 3                         # 同页至少这么多个点线条目才认定为目录页
TOC_Y_TOL = 14.0                         # 译文项归入目录行的 y 对齐容差 (pt)


def is_toc_row(it) -> bool:
    """该项所在行是否点线目录行: 自身文本(整行含点线)或紧邻上下文命中点线。"""
    s = (it.left_ctx or '') + (it.right_ctx or '') + it.text
    return bool(TOC_LINE_RE.search(s))


def toc_entry_pages(en_items: list) -> dict:
    """识别目录页(仅以英文侧为准): 返回 {page: [英文目录行 item, ...]}。"""
    by_page: dict = {}
    for it in en_items:
        if is_toc_row(it):
            by_page.setdefault(it.page, []).append(it)
    return {pg: its for pg, its in by_page.items() if len(its) >= TOC_PAGE_MIN}


def in_toc_rows(it, toc_pages: dict) -> bool:
    """该项是否落在目录页且 y 对齐某条英文点线条目行(章节号/页码两列同 yc)。"""
    rows = toc_pages.get(it.page)
    if not rows:
        return False
    return any(abs(it.yc - e.yc) <= TOC_Y_TOL for e in rows)


def resolve_toc_alignments(pairs: list, toc_pages: dict) -> int:
    """目录条目对称降档(必须在 resolve_aggregation/finalize_statuses 之前):
    以英文侧为准命中目录行的 pair -> 标'目录条目(不校对)', 不判不一致/缺失/多余;
    值相同(一致)的保持"一致"不动基线统计; 译文孤儿 y 对齐英文目录行的同步降档。
    不删任何检查点、不动匹配代价/权重。"""
    if not toc_pages:
        return 0
    changed = 0
    for p in pairs:
        if p.status in ('一致',) or p.status in IGNORED_STATUSES:
            continue   # 值相同保持通过; 已忽略项不动
        if p.en is not None and in_toc_rows(p.en, toc_pages):
            p.status = '目录条目(不校对)'
            p.note = ('目录条目(点线行): 章节号/页码随各语言重排可能变化, 不作数值校对; '
                      + p.note)
            changed += 1
            continue
        # 译文孤儿: y 对齐英文目录行 -> 同为目录页码/编号变体, 不判真多余
        if p.en is None and p.xx is not None and in_toc_rows(p.xx, toc_pages):
            p.status = '目录条目(不校对)'
            p.note = ('目录条目(点线行): 译文页码/编号与英文不同或重排所致, 不作数值校对; '
                      + p.note)
            changed += 1
    return changed


# 图内/图注重排红字特征: 仅保留图注专属(词/符号), 避免误伤正文数字
FIGURE_FEATURES = [
    # 图注专属词(多国语)
    'air inlet', 'air outlet', 'airflow', 'see figure', 'see fig',
    'figure', 'fig.', 'aufnahme', 'abluft', 'anlage', 'montage', 'einbau',
    'entrée d\'air', 'sortie d\'air', 'fixation', 'puesta', 'flujo', 'montaje',
    'ingresso', 'uscita', 'fissaggio', 'inlaat', 'uitlaat', 'bevestig',
    'entrada', 'saída', 'fixação', 'prívod', 'odvod', 'pripevnenie',
    'clear', 'fix ', 'attach', 'shows', 'показано', 'ábra', 'picture', 'photo',
    'felt', 'filzband', 'tape', 'cinta', 'ruban', 'nastro', 'klebeband', 'bande',
    # 图注符号(不常见于正文数字旁)
    'ø', '※', '→', '←', '↑', '↓', '◎', '○', '●',
]


def is_figure_red(text: str, ctx_l: str = '', ctx_r: str = '') -> bool:
    """红字是否在图标注/图形注释区域(易于重排导致位置失真)。仅图注专属词/符号。"""
    s = (ctx_l + ' ' + text + ' ' + ctx_r).lower()
    return any(f.lower() in s for f in FIGURE_FEATURES)
    """同页同指纹同值、且未被配对的剩余候选数量。"""
    n = 0
    for it in items:
        if it.page != pno or id(it) in paired_ids:
            continue
        if fp and it.fp and it.fp != fp:
            continue
        itoks = TOKEN_RE.findall(it.text)
        if vals and any(any(_num_eq(v, t) for t in itoks) for v in vals):
            n += 1
    return n


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


def build_locate_pdf(src_pdf: str, out_path: str, rows: list):
    """问题定位 PDF: 每个问题独占一页, 打开即定位, 无需甄别.
    rows: [(page0, rect, title, lines)]  — page0 为源文档页(0基), rect 为问题区域(pt),
          title 为信息条主行, lines 为附加说明行列表.
    页面布局: 顶部信息条(检查点/风险/值对比) + 原文整页渲染图(180dpi)
              + 问题处黄底红边双框(页面上唯一高亮).
    """
    YELLOW = (1.0, 0.92, 0.23)
    RED = (0.81, 0.13, 0.18)
    GRAY = (0.35, 0.35, 0.35)
    out = fitz.open()
    src = fitz.open(src_pdf)
    pix_cache: dict = {}
    try:
        total = len(rows)
        for idx, (page0, rect, title, lines) in enumerate(rows, 1):
            sp = src[page0]
            page = out.new_page(width=595, height=842)   # A4 纵向
            # 信息条(中文需内置 CJK 字体, 否则丢失)
            page.insert_text((20, 24), title, fontsize=11.5, color=RED,
                             fontname='china-s')
            y = 42
            for ln in lines[:4]:
                page.insert_text((20, y), ln, fontsize=9, color=GRAY,
                                 fontname='china-s')
                y += 13
            # 原文整页渲染图
            area = fitz.Rect(15, 92, 580, 812)
            scale = min(area.width / sp.rect.width, area.height / sp.rect.height)
            w, h = sp.rect.width * scale, sp.rect.height * scale
            x0 = area.x0 + (area.width - w) / 2
            y0 = area.y0 + (area.height - h) / 2
            if page0 not in pix_cache:
                pix_cache[page0] = sp.get_pixmap(dpi=180)
            page.insert_image(fitz.Rect(x0, y0, x0 + w, y0 + h), pixmap=pix_cache[page0])
            # 问题框: 黄底(半透明) + 红边, 页面上唯一
            r = fitz.Rect(rect) & sp.rect
            if r.width >= 1 and r.height >= 1:
                fr = fitz.Rect(x0 + r.x0 * scale - 2.5, y0 + r.y0 * scale - 2.5,
                               x0 + r.x1 * scale + 2.5, y0 + r.y1 * scale + 2.5) & page.rect
                shape = page.new_shape()
                shape.draw_rect(fr)
                # 黄色半透明高亮覆盖数字区域(无框, 与 PDF 高亮习惯一致)
                shape.finish(fill=YELLOW, fill_opacity=0.5)
                shape.commit()
            # 角标: 定位页码 / 总页数
            page.insert_text((555, 830), f'{idx}/{total}', fontsize=8, color=(0.5, 0.5, 0.5))
        out.save(out_path, garbage=3, deflate=True)
    finally:
        out.close()
        src.close()


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


# ---------------- 高亮模式(客户指示稿红框∩青色高亮 = 校对对象) ----------------
# 与红字模式的区别: 检查点来自"指示稿红框内被青色高亮覆盖的内容"(数字/型号/参数/序号),
# 译文侧无红字锚点也可校; 提取只信绘图层小矩形(整页级注释是编辑器全选残留, 必须过滤)。
HL_COVER_MIN = 0.5          # span 被高亮矩形覆盖比例阈
HL_OBJ_AREA_MAX = 0.05      # 单个高亮对象面积 > 页面 5% -> 操作残留, 丢弃


def hl_norm(s: str) -> str:
    """高亮内容归一化: NFKC(全半角/上下标数字) + 连字符族 -> '-' + 空格族剔除"""
    s = unicodedata.normalize('NFKC', s)
    s = re.sub(r'[\u2010\u2011\u2012\u2013\u2014\u2212\u2043\u2212]', '-', s)
    s = re.sub(r'[\u00a0\u2000-\u200b\ufeff]', '', s)
    return s.strip()


def hl_classify(t: str) -> str:
    """高亮内容类型: num 纯数字 / ord 序号 / code 型号代码 / symstr 数字符号串 /
    phrase 短参数短语 / text 句子级(不作检查点)"""
    tt = t.strip().strip('.·*')
    if re.fullmatch(r'\d+(?:[.,]\d+)*', tt):
        return 'num'
    if re.fullmatch(r'[(（]?\d+[)）.、]', t.strip()) or re.fullmatch(r'[①-⑳]', t.strip()):
        return 'ord'
    if (re.fullmatch(r'[A-Za-z0-9/\-().,%+²³°:]{2,24}', tt) and re.search(r'[A-Za-z]', tt)
            and re.search(r'\d', tt)):
        # 含可译小写单词(3-core/4-adrig): 单词跨语言必变, 数字才是稳定锚 → phrase
        if re.search(r'[a-z]', tt):
            return 'phrase'
        return 'code'
    if re.fullmatch(r'[\d.,/\-–:()%\s≤≥±°]+', tt) and re.search(r'\d', tt):
        return 'symstr'
    if len(tt.split()) <= 3 and len(tt) <= 24:
        return 'phrase'
    return 'text'


def hl_highlight_rects(page) -> list:
    """页内真高亮矩形: 青色填充绘图矩形 + 小尺寸 Highlight 注释; 异常大对象过滤"""
    page_area = max(page.rect.get_area(), 1.0)
    out = []
    for dr in page.get_drawings():
        f = dr.get('fill')
        if f and _color_match(tuple(f), ANCHOR_FILL):
            r = fitz.Rect(dr['rect'])
            if r.width > 0.5 and r.height > 0.5 and r.get_area() < page_area * HL_OBJ_AREA_MAX:
                out.append(r)
    for a in (page.annots() or []):
        if a.type[1] == 'Highlight':
            r = fitz.Rect(a.rect)
            if r.get_area() < page_area * HL_OBJ_AREA_MAX:
                out.append(r)
    return out


def extract_highlight_items(anchor_doc, zones: list, log=print):
    """提取红框∩高亮的检查点(字符级精确提取 + 片段合并) + 跳过清单.
    必须按字符判定: 长 span 内只高亮一个词('under 40 °C' 只涂 40 / '1-2. SPECIFICATIONS'
    只涂 1-2.)时, span 级覆盖率只有 3~17% 必漏。返回 (items, skipped)。"""
    items, skipped = [], []
    for pno in range(anchor_doc.page_count):
        page = anchor_doc[pno]
        hls = hl_highlight_rects(page)
        zs = zones[pno] if pno < len(zones) else []
        if not hls:
            continue

        def char_covered(r):
            """字符被高亮: 字符垂直中心落在某矩形内(±1pt) 且水平重叠≥50%字符宽"""
            if zs and not any(r.intersects(z) for z in zs):
                return False
            cy = (r.y0 + r.y1) / 2
            for h in hls:
                if h.y0 - 1.0 <= cy <= h.y1 + 1.0:
                    ix = max(0.0, min(r.x1, h.x1) - max(r.x0, h.x0))
                    if ix >= max(r.width * 0.5, 0.4):
                        return True
            return False

        row = {}
        for b in page.get_text('rawdict')['blocks']:
            if b.get('type') != 0:
                continue
            for l in b.get('lines', []):
                for sp in l.get('spans', []):
                    for ch in sp.get('chars', []):
                        c = ch['c']
                        if not c:
                            continue
                        r = fitz.Rect(ch['bbox'])
                        key = round((r.y0 + r.y1) / 2 / 4.0)
                        row.setdefault(key, []).append((r.x0, r.x1, c, char_covered(r), (r.y0 + r.y1) / 2))
        for key, chars in sorted(row.items()):
            chars.sort()
            cov_idx = [i for i, ch in enumerate(chars) if ch[3] and not ch[2].isspace()]
            if not cov_idx:
                continue
            # 相邻覆盖字符分组(间隙<6pt 同片段)
            groups = []
            for i in cov_idx:
                if groups and chars[i][0] - chars[groups[-1][-1]][1] < 6.0:
                    groups[-1].append(i)
                else:
                    groups.append([i])
            for g in groups:
                lo, hi = g[0], g[-1]
                # 词边界补全: 客户手绘高亮常只涂型号中段('MXZ-'未涂), 向两侧扩展到完整词
                while lo > 0 and not chars[lo - 1][2].isspace() \
                        and chars[lo][0] - chars[lo - 1][1] < 1.5:
                    lo -= 1
                while hi < len(chars) - 1 and not chars[hi + 1][2].isspace() \
                        and chars[hi + 1][0] - chars[hi][1] < 1.5:
                    hi += 1
                x0, x1 = chars[lo][0], chars[hi][1]
                yc = chars[g[0]][4]
                t = ''.join(chars[k][2] for k in range(lo, hi + 1)).strip()
                if not t:
                    continue
                zt = fitz.Rect(x0, yc - 4, x1, yc + 4)
                if zs and not any(zt.intersects(z) for z in zs):
                    skipped.append((pno + 1, t, '高亮在红框外'))
                    continue
                typ = hl_classify(t)
                # TOC 点线行豁免: 整行含 ≥5 连续点 -> 目录页码/章节号, 位置随译文重排不稳
                full = page.get_textbox(fitz.Rect(max(0.0, x0 - 260), yc - 4,
                                                  min(page.rect.x1, x1 + 260), yc + 4)) or ''
                if re.search(r'\.{5,}', full):
                    typ = 'toc'
                # 脚注枚举列表(*6, *7, *8 等): 非数据内容, 同 toc 全册豁免
                if re.fullmatch(r'(?:\*\d+[,，、\s]+)+\*?\d+', t.strip()):
                    typ = 'toc'
                if typ == 'text':
                    skipped.append((pno + 1, t[:60], '句子级文字(一期不校)'))
                    continue
                items.append(Item(page=pno, text=t, bbox=(x0, yc - 4, x1, yc + 4),
                                  yc=yc, n_spans=len(t), hl_type=typ))
    log(f'高亮提取: 检查点 {len(items)} 个, 跳过 {len(skipped)} 项(句子级/框外, 见清单)')
    return items, skipped


def hl_page_lines(doc):
    """译文行级文本重建: {page0: [(yc, 归一化行文本, 原始行文本, x0, x1)]}
    先按 y 分行、行内按 x 排序拼接: 上下标(CO₂ 的 2 / mm²)yc 与正文差仅几分之一 pt,
    若按 (yc,x0) 全局排序会被排到行尾; 行内 x 排序天然落回原位。"""
    pages = {}
    for pno in range(doc.page_count):
        spans = []
        for b in doc[pno].get_text('dict')['blocks']:
            if b.get('type') != 0:
                continue
            for l in b.get('lines', []):
                for s in l.get('spans', []):
                    t = s['text'].strip()
                    if t:
                        spans.append(((s['bbox'][1] + s['bbox'][3]) / 2,
                                      s['bbox'][0], s['bbox'][2], t))
        spans.sort(key=lambda x: x[0])
        groups = []   # [center_yc, [spans...]]
        for sp in spans:
            if groups and sp[0] - groups[-1][0] <= 5.0:
                groups[-1][1].append(sp)
                groups[-1][0] = sum(q[0] for q in groups[-1][1]) / len(groups[-1][1])
            else:
                groups.append([sp[0], [sp]])
        out = []
        for cy, grp in groups:
            grp.sort(key=lambda q: q[1])
            text = ''
            prev_x1 = None
            for yc, x0, x1, t in grp:
                join = '' if prev_x1 is None or x0 - prev_x1 < 1.5 else ' '
                text += join + t
                prev_x1 = max(prev_x1 or 0, x1)
            out.append((cy, hl_norm(text), text, grp[0][1], prev_x1))
        pages[pno] = out
    return pages


def hl_num_variants(tok: str) -> list:
    """数字 token 的欧式/英式逗号变体"""
    out = [tok]
    if '.' in tok:
        out.append(tok.replace('.', ','))
    if ',' in tok:
        out.append(tok.replace(',', '.'))
    return out


def hl_keys(item) -> list:
    """检索键组: 每组=一个原始 token 的变体集(组内任一命中即算该 token 命中).
    数字类(含短语/符号串)=全部数值 token; 枚举型 code('*6, *7, *8')拆多 token 同行全含;
    其余字符串类=归一化原文一组。"""
    t = hl_norm(item.text)
    if item.hl_type in ('num', 'symstr', 'phrase', 'toc'):
        toks = re.findall(r'\d+(?:[.,]\d+)*', t)
        if not toks:
            return [frozenset({t})]
        return [frozenset(hl_num_variants(x)) for x in toks]
    if item.hl_type == 'code' and re.search(r'[,，/]', t):
        toks = [x.strip() for x in re.split(r'[,，/]', t) if x.strip()]
        if len(toks) >= 2:
            return [frozenset({x, x.strip('().,;:!?')}) for x in toks]
    if item.hl_type == 'code':
        # 去尾标点变体: 英文高亮片段 'CN750.' vs 译文 'CN750'
        stripped = t.strip('().,;:!?')
        return [frozenset({t, stripped})] if stripped and stripped != t else [frozenset({t})]
    return [frozenset({t})]


def _hl_key_all_tokens(item) -> bool:
    """需要"同一行全含所有 token"的类型: 短语/符号串/枚举 code"""
    return item.hl_type in ('phrase', 'symstr') or         (item.hl_type == 'code' and len(hl_keys(item)) > 1)


def _hl_line_hit(ntext: str, key_groups: list, loose: bool, all_tokens: bool = False) -> bool:
    """all_tokens=True: 每个 token 须同一行出现(防 '15 / 20' 单侧数字无关行误命中).
    纯数字 token 用边界断言(0.8 不命中 10.8); 字符串 token 用子串(code 边界断言会被
    行级重建的粘连 'MXZ-2G33VG0,8' 反噬, 且型号串歧义低, 子串安全)。"""
    def one(group):
        if loose:
            return any(g in ntext for g in group)
        g0 = next(iter(group))
        if re.fullmatch(r'[\d.,]+', g0):
            # 纯数字: 防 0.8 命中 10.8/0.85, 但允许句尾点号(675.)
            return any(re.search(r'(?<![\d.,])' + re.escape(g) + r'(?!\d)(?![.,]\d)', ntext)
                       for g in group)
        if any(g in ntext for g in group):
            return True
        ns = ntext.replace(' ', '')   # 去空格兜底: 行重建 'CO 2' vs 键 'CO2'
        return any(g.replace(' ', '') in ns for g in group)
    if all_tokens:
        return bool(key_groups) and all(one(g) for g in key_groups)
    return any(one(g) for g in key_groups)


def hl_find_in_lines(item, pages: dict, pno0: int, y_off: float = 0.0, page_off: int = 0):
    """在译文页行集中找高亮内容: 同页优先、邻近页次之, 多命中取 y 最近(行带按页偏移 y_off 校准).
    字符串类(code/ord)严格未中时允许子串兜底; 序号类强制行带±40(无位置约束的'(1)'无意义).
    返回 (page0, yc, 命中键, 行文本) 或 None"""
    keys = hl_keys(item)
    if not keys:
        return None
    loose = item.hl_type in ('code', 'ord')
    all_tok = _hl_key_all_tokens(item)
    order = sorted(pages.keys(), key=lambda p: abs(p - pno0 - page_off))
    # 同页(含页码偏移) strict→loose 先扫(行级粘连的型号在 loose 才命中), 再邻页兜底
    scans = ([('same', 'strict'), ('same', 'loose'), ('near', 'strict'), ('near', 'loose')]
             if loose else [('same', 'strict'), ('near', 'strict')])
    for scope, mode in scans:
        cands = []
        for p in (order[:1] if scope == 'same' else order[1:]):
            for yc, ntext, raw, x0, x1 in pages[p]:
                if item.hl_type == 'ord' and abs(yc - item.yc - y_off) > 40:
                    continue
                if _hl_line_hit(ntext, keys, mode == 'loose', all_tok):
                    cands.append((p, yc, keys[0], raw))
        if cands:
            cands.sort(key=lambda c: (abs(c[0] - item.page), abs(c[1] - item.yc - y_off)))
            return cands[0]
    if enum_code_fallback(item, keys, pages, order, y_off):
        p = order[0]
        return (p, item.yc + y_off, keys[0], '枚举型号跨行断词: 各 token 同页分别命中')
    # 断词连行兜底: 型号被行断 'MXZ-' / '2HB50VF)' → 相邻行拼接后查找
    if item.hl_type == 'code':
        for g in keys:
            for p in order:
                plines = pages.get(p, [])
                for i in range(len(plines) - 1):
                    ya, _, ra, _, _ = plines[i]
                    yb, _, rb, _, _ = plines[i + 1]
                    if 0 < yb - ya <= 12:
                        # 下行开头接上行结尾(去空格变体)
                        joined = ra.rstrip() + rb.lstrip()
                        if any(x in joined.replace(' ', '') for x in g):
                            return (p, ya, keys[0], f'型号跨行断词命中: {joined[:40]}')
    return None


def enum_code_fallback(item, keys, pages, order, y_off) -> bool:
    """逗号分隔多型号(如 'MXZ-2HB40VF,MXZ-2HB50VF)')被译文分行断开时,
    "同行全含"必败; 改为每个型号 token 在同页(含页码偏移页)任意行命中即可。"""
    if item.hl_type != 'code' or len(keys) <= 1:
        return False
    p = item.page
    for g in keys:
        if not any(_hl_line_hit(nt, [g], True, False) for _, nt, _, _, _ in pages.get(p, [])):
            return False
    return True


def hl_page_offset(items, pages, pno0_map=None):
    """校准通道(同红字模式 y_med 思路): 第一遍无视带宽收集命中位置, 得
    (每页 y 偏移中位数 {page0: off}, 全局页码偏移 page_off).
    译文行高/页码整体位移不靠拍脑袋容差吸收。"""
    from statistics import median
    from collections import Counter
    diffs = {}
    pdiff = Counter()
    for it in items:
        keys = hl_keys(it)
        if not keys:
            continue
        loose = it.hl_type in ('code', 'ord')
        all_tok = _hl_key_all_tokens(it)
        # 页码偏移: 全册最近命中页 - 英文页
        best_all = None
        for p in sorted(pages, key=lambda q: abs(q - it.page)):
            for yc, ntext, raw, x0, x1 in pages[p]:
                if _hl_line_hit(ntext, keys, loose, all_tok):
                    if best_all is None or abs(yc - it.yc) < abs(best_all[0] - it.yc):
                        best_all = (yc, p)
                    break
        if best_all is not None:
            pdiff[best_all[1] - it.page] += 1
        best = None
        for yc, ntext, raw, x0, x1 in pages.get(it.page, []):
            if _hl_line_hit(ntext, keys, loose, all_tok):
                d = yc - it.yc
                if best is None or abs(d) < abs(best):
                    best = d
        if best is not None:
            diffs.setdefault(it.page, []).append(best)
    page_off = pdiff.most_common(1)[0][0] if pdiff and pdiff.most_common(1)[0][1] >= 3 else 0
    return {p: median(v) for p, v in diffs.items() if len(v) >= 3}, page_off


def hl_compare_lang(item, pages: dict, y_off: float = 0.0, page_off: int = 0) -> tuple:
    """单语言比对一个高亮检查点 -> (status, conf, xx_repr, note)
    行带 = 英文 yc + 页级偏移 y_off ± 40(校准后真正的"同位置");
      行带内命中             -> 一致(≤20 high, 否则 medium)
      同页行带外/跨页命中    -> 中风险(疑串位/错印后在他处出现)
      全册未找到             -> 高风险(真缺失)"""
    hit = hl_find_in_lines(item, pages, item.page, y_off=y_off, page_off=page_off)
    if hit:
        p, yc, key, raw = hit
        yoff = round(yc - item.yc - y_off, 1)
        if item.hl_type == 'toc':
            # 目录页码/脚注枚举: 全册任意页命中即一致(位置随重排不稳, 页级无意义)
            return '一致', 'medium', raw, f'目录/脚注枚举命中(仅全册值校验): {item.text} (译文页{p + 1})'
        if p == item.page + page_off and abs(yoff) <= 20:
            return '一致', 'high', raw, f'高亮{item.hl_type}命中: {item.text} (译文页{p + 1}, 校准y偏移{yoff:+.0f}pt)'
        if p == item.page + page_off and abs(yoff) <= 40:
            return '一致', 'medium', raw, f'高亮{item.hl_type}命中(位移较大): {item.text} (译文页{p + 1}, 校准y偏移{yoff:+.0f}pt)'
        return ('中风险', 'low', raw,
                f'需复核(疑串位/错印): 高亮 {item.text} 同页行带内未找到, 命中文在译文页{p + 1} y={yc:.0f}: {raw[:50]}')
    return '高风险', '-', None, f'真缺失: 译文全册未找到高亮内容 {item.text!r}(疑漏印/错印, 必须人工核对)'


def run_highlight_job(instruction_pdf, data_dir, out=None, log=print):
    """高亮模式主流程: 指示稿提取 -> 各译文检索比对 -> 报告(复用四档状态)"""
    out_dir = out or os.path.join(data_dir, '_高亮校对结果')
    os.makedirs(out_dir, exist_ok=True)
    anchor_doc = fitz.open(instruction_pdf)
    zones = anchor_zone_rects(anchor_doc)
    items, skipped = extract_highlight_items(anchor_doc, zones, log)
    files = sorted(f for f in os.listdir(data_dir)
                   if f.lower().endswith('.pdf')
                   and os.path.abspath(os.path.join(data_dir, f)) != os.path.abspath(instruction_pdf))
    log(f'指示稿: {os.path.basename(instruction_pdf)} | 高亮检查点 {len(items)} | 译文 {len(files)} 个')
    results = []
    for fname in files:
        lang = lang_code(fname)
        doc = fitz.open(os.path.join(data_dir, fname))
        pages = hl_page_lines(doc)
        y_offs, page_off = hl_page_offset(items, pages)   # 页级/页码偏移校准
        pairs = []
        for i, it in enumerate(items, 1):
            st, conf, raw, note = hl_compare_lang(it, pages, y_offs.get(it.page, 0.0), page_off)
            xx = Item(page=it.page, text=raw[:60], bbox=(0, 0, 0, 0), yc=0, n_spans=0) if raw else None
            pairs.append(Pair(cp=f'HL{i:03d}', page=it.page + 1, en=it, xx=xx,
                              y_off=None, conf=conf, status=st, note=note))
        n_ok = sum(1 for p in pairs if p.status == '一致')
        log(f'  [{lang}] {fname}: 检查点{len(items)} 命中{n_ok} 未命中{len(items)-n_ok}')
        results.append({'file': fname, 'lang': lang, 'pairs': pairs, 'doc': doc,
                        'path': os.path.join(data_dir, fname), 'n_xx': 0, 'n_sp': 0, 'cc': Counter()})
    # 高亮清单(首跑人工核对)
    import csv
    with open(os.path.join(out_dir, '高亮清单.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['页', '内容', '类型', '状态'])
        for it in items:
            w.writerow([it.page + 1, it.text, it.hl_type, '检查点'])
        for pg, t, why in skipped:
            w.writerow([pg, t, '-', why])
    # 简易汇总 Excel
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '汇总'
    ws.append(['文件', '语言', '高亮检查点', '命中(一致)', '未命中(待核对)'])
    for r in results:
        ok = sum(1 for p in r['pairs'] if p.status == '一致')
        ws.append([r['file'], r['lang'], len(r['pairs']), ok, len(r['pairs']) - ok])
    ws2 = wb.create_sheet('明细')
    ws2.append(['语言', '检查点', '页', '状态', '高亮内容', '译文命中行', '备注'])
    import re as _re
    def _clean(v):
        return _re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(v)) if isinstance(v, str) else v
    for r in results:
        for p in r['pairs']:
            ws2.append([r['lang'], p.cp, p.page, p.status, _clean(p.en.text),
                        _clean(p.xx.text) if p.xx else '', _clean(p.note)])
    wb.save(os.path.join(out_dir, '高亮校对报告.xlsx'))
    xlsx = os.path.join(out_dir, '高亮校对报告.xlsx')
    log(f'报告: {xlsx}\n清单: 高亮清单.csv(首跑请人工核对提取范围)')
    for d in (anchor_doc, *[r['doc'] for r in results]):
        d.close()
    return xlsx, None, out_dir, out_dir


def run_job(base, anchor, data_dir, out=None, log=print):
    out_dir = out or os.path.join(data_dir, '_校对结果')
    snaps_dir = os.path.join(out_dir, 'snaps')
    os.makedirs(snaps_dir, exist_ok=True)

    log(f'英文指示稿: {base}')
    en_color_counter = Counter()
    en_doc = fitz.open(base)
    anchor_zones = None
    if anchor:
        anchor_doc = fitz.open(anchor)
        anchor_zones = anchor_zone_rects(anchor_doc)
        log(f'锚定模式: {anchor} (红框区域, 框内数字为校对对象)')
    en_items = extract_items(en_doc, en_color_counter)
    attach_fp(en_items, en_doc)
    toc_pages = toc_entry_pages(en_items)   # 目录页(仅英文侧判定): 后续对称降档用
    if toc_pages:
        log('目录页识别: ' + ', '.join(f'P{pg + 1}({len(its)}条点线条目)' for pg, its in sorted(toc_pages.items())))
    if anchor:
        n_in = sum(1 for it in en_items if in_anchor(it, anchor_zones))
        log(f'英文检查点(全量)={len(en_items)} 框内(锚定)={n_in}')
    else:
        log(f'英文检查点数: {len(en_items)}')
    en_sorted = sorted(en_items, key=lambda i: (i.page, i.yc, i.bbox[0]))

    files = [f for f in os.listdir(data_dir)
             if f.lower().endswith('.pdf') and os.path.abspath(os.path.join(data_dir, f)) != os.path.abspath(base)
             and (not anchor or os.path.abspath(os.path.join(data_dir, f)) != os.path.abspath(anchor))
             and not re.search(r'0\dEn|_01En\b', f, re.I)
             and not ('英文' in f and '校对' not in f) and not f.startswith('英文')]
    files.sort()
    log(f'待校对文件: {len(files)} 个')

    results = []
    color_notes = []   # 非标准红提示
    review_dir = os.path.join(out_dir, '复核PDF')
    en_marks: dict = {}   # 英文侧标注: (page,yc,x0) -> {bbox, langs}
    en_orphans: list = []  # 译文多出项: (lang, pair) -> 也入英文定位 PDF(同位置参考页)
    en_sorted_page = {}
    for it in en_sorted:
        en_sorted_page.setdefault(it.page, []).append(it)
    for fname in files:
        path = os.path.join(data_dir, fname)
        lang = lang_code(fname)
        doc = fitz.open(path)
        cc = Counter()
        xx_items = extract_items(doc, cc)
        attach_fp(xx_items, doc)
        pairs = build_pairs(en_items, xx_items, en_doc.page_count, doc.page_count)
        # 回填检查点编号(匹配页内的编号)
        for p in pairs:
            if p.en is not None:
                lst = en_sorted_page[p.en.page]
                p.cp = f'P{p.en.page + 1}-{lst.index(p.en) + 1}'
        offbox = 0   # 框外值差异计数(不计入问题, 仅留痕供参考)
        if anchor_zones:
            # 锚定模式: 英文检查点不在红框内 -> 非锚定区(忽略, 不入选统计/截图/报告)
            for p in pairs:
                if p.en is not None and not in_anchor(p.en, anchor_zones):
                    p.status = '非锚定区(忽略)'
                    p.note = '非锚定区: 英文检查点不在客户红框内'
                    if p.xx is not None:
                        # 只统计 token 数相等且逐位值不等的对(真值差异);
                        # token 数不等的是粘连/拆分粒度差异(值在页内可寻), 不算框外差异
                        et = TOKEN_RE.findall(p.en.text)
                        xt = TOKEN_RE.findall(p.xx.text)
                        if et and len(et) == len(xt) and any(not _num_eq(a, b) for a, b in zip(et, xt)):
                            p.note += f'; 框外值差异: {p.en.text} → {p.xx.text}(非校对对象, 不计入)'
                            offbox += 1
            # 译文孤儿(多出/未匹配): 其译文项不在任何红框内(x+y 全判) -> 非锚定区, 忽略
            for p in pairs:
                if p.status in ('译文多出', '译文未匹配'):
                    reffx = p.xx if p.xx is not None else p.en
                    if reffx is not None and reffx.page < len(anchor_zones):
                        bb = fitz.Rect(reffx.bbox)
                        inbox = any(bb.intersects(z) or (z.x0 - 10 <= reffx.xc <= z.x1 + 10
                                                          and z.y0 - 8 <= reffx.yc <= z.y1 + 8)
                                    for z in anchor_zones[reffx.page])
                        if not inbox:
                            p.status = '非锚定区(忽略)'
                            p.note = '非锚定区: 译文项不在客户红框范围内'
        resolve_toc_alignments(pairs, toc_pages)   # 目录条目对称降档(不删项/不动权重)
        resolve_aggregation(pairs, xx_items)
        if anchor_zones:
            # 编号列对齐(span级): 目录/列表编号按序对齐, 消除排版位移误报
            align_entry_columns_pairs(pairs, en_items, xx_items)
            align_entry_columns_spans(pairs, en_doc, doc)
        finalize_statuses(pairs, en_items, xx_items)
        # 非标准红提示: 文件中出现了英文稿没有的红系颜色
        en_main = {c for c, _ in en_color_counter.most_common(3)}
        odd = {c: n for c, n in cc.items() if c not in en_main}
        if odd:
            color_notes.append(f'{lang}: 检测到非英文稿主色的红系颜色 {odd}, 已一并提取, 建议与标注方确认规范')
        results.append({'file': fname, 'lang': lang, 'pairs': pairs,
                        'n_xx': len(xx_items), 'n_sp': 0, 'doc': doc, 'path': path, 'cc': cc,
                        'offbox': offbox})

    # 跨语言交叉验证聚合差异(在截图/统计前, 影响最终状态)
    n_cross = cross_validate_aggregation(results)
    if n_cross:
        log(f'跨语言交叉验证: {n_cross} 项聚合差异自动判为一致(排版差异)')

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
            # 一致 -> 无需截图; 低风险(大位移确认,位移大)配截图供人工抽查; 已并入项不截图
            if p.status == '一致':
                continue
            if p.status in IGNORED_STATUSES:
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
            elif p.xx is None and p.expect_y is not None and p.en is not None:
                # 译文侧期望位置截图(缺失处上下文): 黄框标出推断的期望位置, 窗口外扩留可辨认上下文
                e = p.en
                cx = (e.bbox[0] + e.bbox[2]) / 2 + (p.expect_dx or 0.0)
                hh = max(e.bbox[3] - e.bbox[1], 10.0)
                page = doc[e.page]
                mark = fitz.Rect(cx - 45, p.expect_y - hh / 2 - 8,
                                 cx + 45, p.expect_y + hh / 2 + 8)
                r2 = fitz.Rect(mark.x0 - 25, mark.y0 - 25,
                               mark.x1 + 25, mark.y1 + 25) & page.rect
                mark = mark & page.rect
                if r2.width > 2 and r2.height > 2:
                    snap_marked(page, r2, mark,
                                os.path.join(lang_snap, f'{tag}_{lang}_XX@期望位置.png'))
                    n_snaps += 1
            if p.en is None and p.xx is not None:
                # 译文多出: 英文稿同页同位置取参考截图(双语页码基本对齐),
                # 供人工看"英文稿这个位置是什么", 与译文侧"期望位置"截图对称; 黄框标出参考区域
                x = p.xx
                ep = min(x.page, en_doc.page_count - 1)
                cx = (x.bbox[0] + x.bbox[2]) / 2
                hh = max(x.bbox[3] - x.bbox[1], 10.0)
                page = en_doc[ep]
                mark = fitz.Rect(cx - 45, x.yc - hh / 2 - 8, cx + 45, x.yc + hh / 2 + 8)
                r2 = fitz.Rect(mark.x0 - 25, mark.y0 - 25,
                               mark.x1 + 25, mark.y1 + 25) & page.rect
                mark = mark & page.rect
                if r2.width > 2 and r2.height > 2:
                    snap_marked(page, r2, mark,
                                os.path.join(lang_snap, f'{tag}_{lang}_EN@同位置参考.png'))
                    n_snaps += 1
        problems = sum(1 for p in pairs if p.status != '一致' and p.status not in IGNORED_STATUSES)
        n_sp = sum(1 for p in pairs if p.status == '低风险')
        r['n_sp'] = n_sp
        # 定位 PDF: 每个问题独占一页(打开即定位, 页上唯一黄底红边框)
        loc_rows = []
        lname = LANG_NAMES.get(lang, lang)
        for p in pairs:
            if p.status in ('一致',) + IGNORED_STATUSES:
                continue
            if p.xx is not None:
                page0 = p.xx.page
                rect = p.xx.bbox
                xv = p.xx.text
            elif p.expect_y is not None and p.en is not None and p.xx is None:
                page0 = p.en.page
                cx = (p.en.bbox[0] + p.en.bbox[2]) / 2 + (p.expect_dx or 0.0)
                hh = max(p.en.bbox[3] - p.en.bbox[1], 10.0)
                rect = (cx - 45, p.expect_y - hh / 2 - 8, cx + 45, p.expect_y + hh / 2 + 8)
                xv = '(缺失, 框为期望位置)'
            else:
                page0 = p.xx.page if p.xx else 0
                rect = p.xx.bbox if p.xx else (40, 40, 200, 80)
                xv = p.xx.text if p.xx else '(缺失)'
            p.locate_page = len(loc_rows) + 1
            ev = p.en.text if p.en is not None else '(无)'
            title = f'{p.cp or "(译文多出)"} · {lname} · {p.status}'
            lines = [f'原文页码: {page0 + 1}    英文值: {ev} → 译文值: {xv}']
            if p.note:
                lines.append('备注: ' + p.note[:90])
            loc_rows.append((page0, rect, title, lines))
        if loc_rows:
            os.makedirs(review_dir, exist_ok=True)
            build_locate_pdf(path, os.path.join(review_dir, f'{lang}.pdf'), loc_rows)
        # 收集英文侧标注(同一检查点多语言问题叠加)
        for p in pairs:
            if p.status in ('一致',) + IGNORED_STATUSES or p.en is None:
                continue
            key = (p.en.page, round(p.en.yc, 1), round(p.en.bbox[0], 1))
            en_marks.setdefault(key, {'bbox': p.en.bbox, 'langs': set()})
            en_marks[key]['langs'].add(lang)
        # 译文多出项: 收集供英文定位 PDF 加页(点击英文侧不再错跳译文 PDF)
        for p in pairs:
            if p.en is None and p.xx is not None \
                    and p.status not in ('一致',) + IGNORED_STATUSES:
                en_orphans.append((lang, p))
        log(f'  [{lang}] {fname}: 检查点{len(en_items)} 译文红字{r["n_xx"]} '
              f'问题项{problems} 低风险{n_sp} 框外差异{offbox} (截图{n_snaps}张 定位PDF{len(loc_rows)}页)')
        doc.close()

    # 英文指示稿定位 PDF: 每个问题检查点独占一页
    en_rows = []
    for key, info in sorted(en_marks.items()):
        page0, yc, x0 = key
        langs = ','.join(sorted(info['langs']))
        en_rows.append([page0, info['bbox'], f'英文指示稿 · {page0 + 1}页 · 问题语言: {langs}', []])
        info['page_no'] = len(en_rows)          # 该检查点在 EN 定位 PDF 中的页码
    # 译文多出项: 英文稿同位置参考页(矩形取译文项坐标映射到英文页, 页码基本对齐)
    for lang, p in en_orphans:
        x = p.xx
        page0 = min(x.page, en_doc.page_count - 1)
        cx = (x.bbox[0] + x.bbox[2]) / 2
        hh = max(x.bbox[3] - x.bbox[1], 10.0)
        rect = (cx - 45, x.yc - hh / 2 - 8, cx + 45, x.yc + hh / 2 + 8)
        en_rows.append([page0, rect,
                        f'英文指示稿 · {page0 + 1}页 · {lang} 译文多出(英文无对应红字)',
                        [f'译文值: {x.text}    备注: {p.note[:80]}']])
        p.en_locate_page = len(en_rows)
    # 回填每条问题的英文侧定位页码
    for r in results:
        for p in r['pairs']:
            if p.en is None or p.status in ('一致',) + IGNORED_STATUSES:
                continue
            key = (p.en.page, round(p.en.yc, 1), round(p.en.bbox[0], 1))
            info = en_marks.get(key)
            if info and 'page_no' in info:
                p.en_locate_page = info['page_no']
    if en_rows:
        os.makedirs(review_dir, exist_ok=True)
        build_locate_pdf(base, os.path.join(review_dir, 'EN.pdf'), en_rows)

    report = os.path.join(out_dir, '数字校对报告.xlsx')
    build_excel(report, os.path.basename(base), en_items, results, snaps_dir, color_notes)
    report_html = os.path.join(out_dir, '数字校对报告.html')
    build_html(report_html, os.path.basename(base), en_items, results, snaps_dir)
    log(f'\n报告已生成: {report}')
    log(f'           {report_html}')
    log(f'截图目录:   {snaps_dir}')
    return report, report_html, snaps_dir, out_dir


def _hungarian(g):
    """最小成本二分配对(匈牙利算法 Kuhn-Munkres, 纯 Python), 返回 (ri, cj).
    g: list[list[float]] 成本矩阵 (n 行 x m 列). 按 e-maxx 标准写法.
    """
    n, m = len(g), len(g[0]) if g else 0
    if not n or not m:
        return [], []
    BIG = max(max(row) for row in g) + 1e12
    if n > m:
        ri, cj = _hungarian([list(row) for row in zip(*g)])
        return list(cj), list(ri)
    u = [0.0] * (n + 1)
    v = [0.0] * (m + 1)
    p = [0] * (m + 1)
    way = [0] * (m + 1)
    for i0 in range(1, n + 1):
        p[0] = i0
        j0 = 0
        minv = [BIG] * (m + 1)
        used = [False] * (m + 1)
        while True:
            used[j0] = True
            i0_ = p[j0]
            delta = BIG
            j1 = 0
            for j in range(1, m + 1):
                if used[j]:
                    continue
                cur = g[i0_ - 1][j - 1] - u[i0_] - v[j]
                if cur < minv[j]:
                    minv[j] = cur
                    way[j] = j0
                if minv[j] < delta:
                    delta = minv[j]
                    j1 = j
            for j in range(m + 1):
                if used[j]:
                    u[p[j]] += delta
                    v[j] -= delta
                else:
                    minv[j] -= delta
            j0 = j1
            if p[j0] == 0:
                break
        while True:
            j1 = way[j0]
            p[j0] = p[j1]
            j0 = j1
            if j0 == 0:
                break
    ri, cj = [], []
    for j in range(1, m + 1):
        if p[j]:
            ri.append(p[j] - 1)
            cj.append(j - 1)
    return list(ri), list(cj)


def _hungarian_test():
    """对拍验证: 与穷举最优解比较 (小规模随机矩阵)."""
    import random
    from itertools import permutations
    for _ in range(300):
        n = random.randint(1, 7)
        m = random.randint(1, 7)
        g = [[random.randint(0, 20) for _ in range(m)] for _ in range(n)]
        ri, cj = _hungarian(g)
        assert len(ri) == len(cj), (n, m)
        assert len(set(ri)) == len(ri)
        assert len(set(cj)) == len(cj)
        if n <= m:
            best = min(sum(g[i][perm[i]] for i in range(n)) for perm in permutations(range(m), n))
            total = sum(g[i][cj[ri.index(i)]] for i in range(n))
        else:
            best = min(sum(g[perm[j]][j] for j in range(m)) for perm in permutations(range(n), m))
            total = sum(g[ri[i]][cj[i]] for i in range(len(ri)))
        assert abs(total - best) < 1e-9, (n, m, total, best)
    print('_hungarian 对拍验证通过 (300 组随机矩阵)')


def main():
    ap = argparse.ArgumentParser(description='多国语数字校对工具 [P0]')
    ap.add_argument('--mode', choices=['red', 'highlight'], default='red',
                    help='red=红字模式(默认, 校全部红字数字); '
                         'highlight=高亮模式(只校指示稿红框内青色高亮内容: 数字/型号/参数/序号)')
    ap.add_argument('--base', default=None, help='英文红字指示稿(全标红版) PDF [红字模式必填]')
    ap.add_argument('--anchor', default=None, help='高亮模式下=客户指示稿(红框+青色高亮, 必填); 红字模式下=可选锚定原稿')
    ap.add_argument('--dir', required=True, help='多国语 PDF 文件夹')
    ap.add_argument('--out', default=None, help='输出目录 (默认: <dir>/_校对结果 或 _高亮校对结果)')
    args = ap.parse_args()
    if args.mode == 'highlight':
        if not args.anchor:
            ap.error('高亮模式需要 --anchor 指定客户指示稿(含红框+青色高亮)')
        run_highlight_job(args.anchor, args.dir, args.out)
    else:
        if not args.base:
            ap.error('红字模式需要 --base 指定英文红字指示稿')
        run_job(args.base, args.anchor, args.dir, args.out)


if __name__ == '__main__':
    if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    main()
