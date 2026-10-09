# -*- coding: utf-8 -*-
"""
dd_report_consistency.py — 尽调报告一致性核对器（v5.1 新增 · 客户丙0902复盘）
对生成的 docx 做三件事，任何 FAIL 不得交付：

  [1] 比率/勾稽反向复算    —— 变化率、占比、三表勾稽、净利率/毛利率/资产负债率、现金流净增减
  [2] 数字指纹漂移核对     —— 同一主体×同一指标 在全文出现多处时值必须一致（启发式，需人工复核）
  [3] 称谓/释义一致性审查  —— 简称统一、无变体、全称只首次、未释义高频词提示

用法：python dd_report_consistency.py <report.docx> [--config <json>]
退出码：0=通过 1=警告 2=FAIL(禁止交付)
"""
import re
import sys
import json
import argparse
from docx.oxml.ns import qn

KW_CHG = 0.5   # 变化率容差(百分点)
KW_PCT = 0.2   # 占比容差(百分点)
KW_RATIO = 0.5 # 比率容差(百分点)

ASSET = ('资产总计', '资产合计')
LIAB = ('负债合计', '负债总计')
EQUITY = ('所有者权益合计', '股东权益合计')
REVENUE = ('营业收入', '营业总收入')
COST = ('营业成本',)
NETPROFIT = ('净利润',)
GROSS = ('毛利润率', '毛利率')
NETR = ('净利率', '销售净利率')
CASH_NET = ('现金净增减', '现金净增加额')
CASH_END = ('期末可动用货币资金', '期末现金及现金等价物', '期末货币资金')
CASH_OP = ('经营活动现金流量净额',)
CASH_INV = ('投资活动现金流量净额',)
CASH_FIN = ('筹资活动现金流量净额',)

# ===== v5.8 新增：分项求和=合计 =====
# 容差：万元口径 0.05；可由 config.block_sum_tol 覆盖
BLOCK_SUM_TOL = 0.05
# 合计型行（块尾）
BS_TOTAL = ('资产总计', '资产合计', '负债合计', '负债总计', '所有者权益合计', '股东权益合计',
            '流动资产合计', '非流动资产合计', '流动负债合计', '非流动负债合计',
            '归属于母公司所有者权益合计', '归属于母公司股东权益合计')
# 明细子项（不计入父项求和，也非独立科目）
_SUB_SKIP = ('其中：', '其中:', '其中', '　其中')
# 减项 / 加项（计入求和但取相反/正向符号）
_SUB_MINUS = ('减：', '减:')
_SUB_PLUS = ('加：', '加:')
# 利润表结构
IS_REV = ('营业收入', '营业总收入')
IS_SUB = ('营业成本', '税金及附加', '销售费用', '管理费用', '研发费用', '财务费用')
IS_ADD = ('其他收益', '投资收益', '公允价值变动收益', '信用减值损失', '资产减值损失',
          '资产处置收益', '净敞口套期收益', '汇兑收益')


def is_bs_total(k):
    return k in BS_TOTAL or (k.endswith(('合计', '总计')) and not k.startswith(_SUB_SKIP))


def is_sub_skip(k):
    return k.startswith(_SUB_SKIP)

def num(s):
    if not isinstance(s, str):
        return None
    s = s.replace('%', '').replace(',', '')
    s = s.replace('（', '-').replace('）', '')
    s = re.sub(r'[^\d.\-]', '', s)
    if s in ('', '-', '—', '--', 'nan'):
        return None
    try:
        return float(s)
    except ValueError:
        return None

def cell_text(tc):
    return ''.join(x.text or '' for x in tc.iter(qn('w:t'))).strip()

def para_text(p):
    return ''.join(x.text or '' for x in p.iter(qn('w:t'))).strip()

def hit(k, tags):
    """精确/前缀匹配科目标签，避免'流动资产合计'误匹配'资产合计'"""
    return any(k == t or k.startswith(t) for t in tags)

class Checker:
    def __init__(self, path, config=None):
        self.path = path
        self.doc = None
        self.issues = []
        self.config = config or {}

    def add(self, level, msg):
        self.issues.append((level, msg))
        print('  [%s] %s' % (level, msg))

    def tables(self):
        return list(self.doc.element.body.iter(qn('w:tbl')))

    def parse(self, tbl):
        """返回 (header, rows), 并识别列角色"""
        trs = tbl.findall(qn('w:tr'))
        if not trs:
            return [], []
        header = [cell_text(tc) for tc in trs[0].findall(qn('w:tc'))]
        rows = []
        for tr in trs[1:]:
            tcs = tr.findall(qn('w:tc'))
            if tcs:
                rows.append([cell_text(tc) for tc in tcs])
        # 列角色: name=0; 期数值列(不含'占比/变化率/单位'的数字列); pct; chg
        val = []; pct = None; chg = None
        for i, h in enumerate(header):
            if i == 0 or not h:
                continue
            if '变化率' in h or '变动' in h:
                chg = i
            elif '占比' in h or '占营收' in h:
                pct = i
            elif '单位' in h:
                pass
            else:
                val.append(i)
        return header, rows, val, pct, chg

    def rowk(self, row, label, i):
        return row[i] if i is not None and i < len(row) else None

    # ---------- [1] 变化率/占比复算 ----------
    def check_change_pct(self, tbl):
        header, rows, val, pct, chg = self.parse(tbl)
        if not val or (chg is None and pct is None):
            return
        # 确定两期列(本期=val[0], 上期=val[1]) — 注意列序不定, 取前两个有效数值列
        cur, prev = (val[0], val[1]) if len(val) >= 2 else (None, None)
        # 若仅一列(如两期+占比), 识别: 占比列存在 → cur=val[0]
        base_rows = {}
        for i, r in enumerate(rows):
            base_rows[r[0]] = (i, r)
        aL = base_rows.get(ASSET[0]) or base_rows.get(ASSET[1])
        lL = base_rows.get(LIAB[0]) or base_rows.get(LIAB[1])
        eL = base_rows.get(EQUITY[0]) or base_rows.get(EQUITY[1])
        rev = base_rows.get(REVENUE[0]) or base_rows.get(REVENUE[1])
        def base_for(i, label):
            # 利润表: 占比基数=营业收入
            if pct is not None and '占营收' in header[pct] and rev:
                return num(rev[1][cur]) if cur is not None and cur < len(rev[1]) else None
            # 资产负债表: 按行位置划分 资产/负债/权益 区
            if aL and i <= aL[0]:
                return num(aL[1][cur]) if aL and cur < len(aL[1]) else None
            if lL and i <= lL[0]:
                return num(lL[1][cur]) if lL and cur < len(lL[1]) else None
            if eL and i <= eL[0]:
                return num(eL[1][cur]) if eL and cur < len(eL[1]) else None
            return num(eL[1][cur]) if eL and cur < len(eL[1]) else None
        for i, r in enumerate(rows):
            if cur is None or cur >= len(r):
                continue
            cv = num(r[cur]); pv = num(r[prev]) if prev is not None and prev < len(r) else None
            label = r[0]
            # 变化率(分级: >3pp FAIL / 1.5-3pp WARN / <1.5 pass; 含"约"放宽; 小数值亿取整降WARN)
            if chg is not None and cv is not None and pv is not None and abs(pv) > 1e-9:
                ccell = self.rowk(r, label, chg)
                cshow = num(ccell)
                # 比率类指标带单位后缀(pct/倍/次/天/个百分点)时,显示值为两期差值口径,复算改为差值比对
                if ccell and any(u in ccell for u in ('pct', 'PCT', '倍', '次', '天', '个百分点')):
                    calc = cv - pv
                else:
                    calc = (cv - pv) / abs(pv) * 100
                small = (abs(cv) < 1 or abs(pv) < 1)   # 亿元值较小则取整相对误差大
                if cshow is None and ccell and not any(t in ccell for t in ('扭亏', '转正', '由负', '—')):
                    self.add('WARN', '变化率未识别: %s=%r 复算=%.2f%%' % (label, ccell, calc))
                elif cshow is not None:
                    tol = 5 if ('约' in ccell) else 3
                    if abs(cshow - calc) > tol:
                        if small:
                            self.add('WARN', '变化率(小数值亿取整失真): %s 显示=%s%% 复算=%.2f%%（建议备注口径）' % (label, ccell, calc))
                        else:
                            self.add('FAIL', '变化率不符: %s 显示=%s%% 复算=%.2f%%' % (label, ccell, calc))
                    elif abs(cshow - calc) > 1.5:
                        self.add('WARN', '变化率取整疑: %s 显示=%s%% 复算=%.2f%%（若为亿元四舍五入所致请忽略）' % (label, ccell, calc))
            # 占比
            if pct is not None and cv is not None and cv != 0:
                base = base_for(i, label)
                pcell = self.rowk(r, label, pct)
                pshow = num(pcell)
                if base and pshow is not None and abs(pshow - cv / base * 100) > KW_PCT:
                    self.add('FAIL', '占比不符: %s 显示=%s%% 复算=%.2f%%(基数%s)' % (label, pcell, cv / base * 100, base))

    # ---------- [1] 三表勾稽 ----------
    def check_tieout(self, tbl):
        header, rows, val, _, _ = self.parse(tbl)
        if not val:
            return
        cur = val[0]
        d = {}
        for r in rows:
            d[r[0]] = num(r[cur]) if cur < len(r) else None
        def get(tags):
            for k, v in d.items():
                if hit(k, tags):
                    return v
            return None
        a = get(ASSET); l = get(LIAB); e = get(EQUITY)
        if a is not None and l is not None and e is not None:
            if abs(a - (l + e)) > 0.05:
                self.add('FAIL', '勾稽不平: 资产%s ≠ 负债%s+权益%s=%.2f' % (a, l, e, l + e))

    # ---------- [1] 现金流勾稽 ----------
    def check_cash(self, tbl):
        header, rows, val, _, _ = self.parse(tbl)
        if not val:
            return
        cur = val[0]
        d = {}
        for r in rows:
            d[r[0]] = num(r[cur]) if cur < len(r) else None
        def get(tags):
            for k, v in d.items():
                if hit(k, tags):
                    return v
            return None
        op = get(CASH_OP); inv = get(CASH_INV); fin = get(CASH_FIN); net = get(CASH_NET)
        if op is not None and inv is not None and fin is not None and net is not None:
            if abs(op + inv + fin - net) > 0.05:
                self.add('FAIL', '现金流勾稽不平: 经营%s+投资%s+筹资%s=%.2f ≠ 净增减%s'
                         % (op, inv, fin, op + inv + fin, net))

    # ---------- [1] 比率复算(净利率/毛利率/已获利息倍数) ----------
    def check_ratios(self, tbls):
        # 收集 各主体 利润表(营收/成本/净利) 与 比率表(毛利率/净利率)
        profits = []
        for t in tbls:
            header, rows, val, _, _ = self.parse(t)
            if not val:
                continue
            cur = val[0]
            d = {r[0]: (num(r[cur]) if cur < len(r) else None) for r in rows}
            rev = next((v for k, v in d.items() if hit(k, REVENUE)), None)
            cost = next((v for k, v in d.items() if hit(k, COST)), None)
            np_ = next((v for k, v in d.items() if hit(k, NETPROFIT)), None)
            if rev and np_:
                profits.append((rev, cost, np_, d))
        # 对每个利润表, 复核同表是否有 毛利率/净利率 行(比率表是独立表, 此处比对可能跨表, 做提示)
        for rev, cost, np_, d in profits:
            self.add('PASS', '净利率复算=%.2f%% (净利%s/营收%s)' % (np_ / rev * 100, np_, rev))
            if cost is not None:
                self.add('PASS', '毛利率(标准口径)复算=%.2f%%' % ((rev - cost) / rev * 100))
            # 已获利息倍数警示
            fin = next((v for k, v in d.items() if '财务费用' in k), None)
            if fin is not None and abs(fin) < 0.1:
                self.add('WARN', '财务费用≈0(%s)，已获利息倍数=EBIT/利息分母接近0、指标失真，参考有限，建议标注口径' % fin)

    # ---------- [1b] 分项求和=合计（v5.8 · 某锂盐客户庚0911复盘） ----------
    def _find_offenders(self, rows, col, diff):
        """反查差额来源：单项 |v|≈|diff|，或双项 a±b≈diff。返回提示串或''。"""
        vals = []
        for r in rows:
            v = num(r[col]) if col < len(r) else None
            if v is not None:
                vals.append((r[0], v))
        d = abs(diff)
        tol = BLOCK_SUM_TOL
        cands = []
        for k, v in vals:
            if abs(abs(v) - d) < tol:
                cands.append('表中"%s"(%s)' % (k, v))
        if not cands and len(vals) <= 40:
            for i in range(len(vals)):
                for j in range(i + 1, len(vals)):
                    a, b = vals[i][1], vals[j][1]
                    if abs(a + b - diff) < tol or abs(a - b - diff) < tol or abs(b - a - diff) < tol:
                        cands.append('表中"%s"(%s)±"%s"(%s) 组合' % (vals[i][0], a, vals[j][0], b))
                        if len(cands) >= 3:
                            break
                if len(cands) >= 3:
                    break
        return ('；差额恰可由 %s 解释，疑似漏科目/多计项' % '、'.join(cands[:3])) if cands else ''

    def _check_bs_block(self, rows, val):
        """资产负债类：合计行 = 分项之和。
        三种口径依次尝试，取最接近者（消除"子合计/母公司权益+少数股东"两类误报）：
          A 分项模式  —— 前向连续非合计行之和（其中：项不重复计入；减：/加：按符号）
          B 子合计模式——前面紧邻的连续合计行（跨过其间明细行）之和
          C 分项+最近前序合计 —— 如 所有者权益合计 = 归母权益合计 + 少数股东权益
        """
        def val_at(j):
            return num(rows[j][col]) if col < len(rows[j]) else None

        for col in val:
            for i, r in enumerate(rows):
                k = r[0]
                if not is_bs_total(k) or is_sub_skip(k):
                    continue
                v = val_at(i)
                if v is None:
                    continue

                # A 分项模式
                items_a, j = [], i - 1
                while j >= 0:
                    lk = rows[j][0]
                    if is_bs_total(lk) and not is_sub_skip(lk):
                        break
                    x = val_at(j)
                    if x is not None and not is_sub_skip(lk):
                        items_a.append((lk, -x if lk.startswith(_SUB_MINUS) else x))
                    j -= 1
                sum_a = sum(x for _, x in items_a)

                # B 子合计模式：向前取合计行，跨过其间明细行
                #    祖先守卫：某前序合计 > 本合计 → 它是上级（如"资产总计"之于"负债合计"），停止
                items_b, j, seen = [], i - 1, False
                while j >= 0:
                    lk = rows[j][0]
                    x = val_at(j)
                    if is_bs_total(lk) and not is_sub_skip(lk):
                        if x is not None and abs(x) > abs(v) + BLOCK_SUM_TOL:
                            break                # 祖先行，非本合计之分项
                        if x is not None:
                            items_b.append((lk, x))
                        seen = True
                    elif seen:
                        pass                     # 已被前序合计涵盖的明细行，跳过
                    else:
                        break
                    j -= 1
                sum_b = sum(x for _, x in items_b) if items_b else None

                # C 分项 + 最近前序合计
                near_tot = next((val_at(j2) for j2 in range(i - 1, -1, -1)
                                 if is_bs_total(rows[j2][0]) and not is_sub_skip(rows[j2][0])), None)
                sum_c = (sum_a + near_tot) if near_tot is not None else None

                cands = [('分项', sum_a, items_a)]
                if sum_b is not None and items_b:
                    cands.append(('子合计', sum_b, items_b))
                if sum_c is not None and items_a:
                    cands.append(('分项+前序合计', sum_c, items_a))
                mode, s, used = min(cands, key=lambda t: abs(v - t[1]))
                diff = v - s
                if abs(diff) > BLOCK_SUM_TOL:
                    self.add('FAIL', '分项求和≠合计: 【%s】显示=%.2f，%s之和=%.2f（%d项），差额=%.2f%s'
                             % (k, v, mode, s, len(used), diff, self._find_offenders(rows, col, diff)))
                else:
                    self.add('PASS', '分项求和=合计: %s %.2f（%s%d项）' % (k, s, mode, len(used)))

    def _check_is_formula(self, rows, val):
        """利润表：营业利润/利润总额/净利润 结构性复算"""
        def _norm(k):
            for p in ('加：', '加:', '减：', '减:', '其中：', '其中:'):
                if k.startswith(p):
                    return k[len(p):]
            return k

        for col in val:
            d = {}
            for r in rows:
                if len(r) > col:
                    d.setdefault(_norm(r[0]), num(r[col]))

            def g(*tags):
                for t in tags:
                    if d.get(t) is not None:
                        return d[t]
                return None

            rev, op = g(*IS_REV), g('营业利润')
            if rev is None or op is None:
                continue
            sub = [(t, g(t)) for t in IS_SUB]
            add = [(t, g(t)) for t in IS_ADD]
            miss = [t for t, v in sub + add if v is None]
            calc = (rev - sum(v for _, v in sub if v is not None)
                        + sum(v for _, v in add if v is not None))
            if not miss:
                if abs(calc - op) > BLOCK_SUM_TOL:
                    self.add('FAIL', '利润表结构不平: 营业利润 显示=%.2f，按(收入−成本费用+收益项)复算=%.2f，差额=%.2f%s'
                             % (op, calc, op - calc, self._find_offenders(rows, col, op - calc)))
                else:
                    self.add('PASS', '利润表结构复算=%.2f（营业利润）' % calc)
            else:
                gap = calc - op
                if abs(gap) <= BLOCK_SUM_TOL:
                    self.add('PASS', '利润表结构复算=%.2f（缺%s 等%d行为报表本身空白项，差额%.2f，视为通过）'
                             % (calc, '、'.join(miss[:2]), len(miss), gap))
                else:
                    self.add('WARN', '利润表结构无法完整复算：缺%s 等 %d 行；按现有行算=%.2f，与显示营业利润 %.2f 相差 %.2f'
                                     '（差额即候选漏列科目，请回原始报表核对）'
                                     % ('、'.join(miss[:3]), len(miss), calc, op, gap))
            pt = g('利润总额')
            if pt is not None:
                oi, oe = g('营业外收入'), g('营业外支出')
                if oi is not None and oe is not None and abs(op + oi - oe - pt) > BLOCK_SUM_TOL:
                    self.add('FAIL', '利润表结构不平: 利润总额 显示=%.2f，复算(营业利润+营业外收入−营业外支出)=%.2f'
                             % (pt, op + oi - oe))
            np_ = g('净利润')
            if np_ is not None and pt is not None:
                tax = g('所得税费用')
                if tax is not None and abs(pt - tax - np_) > BLOCK_SUM_TOL:
                    self.add('FAIL', '利润表结构不平: 净利润 显示=%.2f，复算(利润总额−所得税)=%.2f' % (np_, pt - tax))

    def check_block_sum(self, tbl):
        """分项求和=合计 总入口：自动判别资产负债类 / 利润表类"""
        header, rows, val, pct, chg = self.parse(tbl)
        if not val or len(rows) < 3:
            return
        labels = [r[0] for r in rows]
        has_bs = any(is_bs_total(k) and k in BS_TOTAL for k in labels)
        has_rev = any(k in IS_REV for k in labels)
        has_is = has_rev and any(k in ('营业利润', '利润总额', '净利润') for k in labels)
        if has_is:
            self._check_is_formula(rows, val)
        if has_bs and not has_rev:
            self._check_bs_block(rows, val)

    # ---------- [3] 称谓/释义 ----------
    def check_terms(self, texts):
        term_tbl = None
        for t in self.tables():
            h = [cell_text(tc) for tc in t.findall(qn('w:tr'))[0].findall(qn('w:tc'))] if t.findall(qn('w:tr')) else []
            if len(h) >= 3 and h[0] == '简称' and '全称' in h[2]:
                term_tbl = t
                break
        if term_tbl is None:
            self.add('WARN', '未找到释义表，跳过称谓核查')
            return
        terms = []
        for tr in term_tbl.findall(qn('w:tr'))[1:]:
            tcs = [cell_text(tc) for tc in tr.findall(qn('w:tc'))]
            if len(tcs) >= 3 and tcs[0]:
                terms.append((tcs[0], tcs[2]))
        full = '\n'.join(texts)
        for abbr, fname in terms:
            # 变体: 简称+"循环/系/体系" 后缀（排除全称"XX客户丙之关联企业科技有限公司"）
            for v in ('循环', '系', '体系'):
                if abbr == '客户丙':
                    c = len(re.findall(r'客户丙' + v + r'(?!科技)', full))
                    if c:
                        self.add('FAIL', '称谓变体未统一: 释义简称%s, 正文仍有"%s" %d次（应用"%s"）' % (abbr, abbr + v, c, abbr))
        freq_terms = ['白名单', '梯次利用', '再生锂盐', '危废', '口径', '回收率', '磷酸铁锂']
        defined = set()
        for _, f in terms:
            for x in freq_terms:
                if x in f:
                    defined.add(x)
        for x in freq_terms:
            c = len(re.findall(x, full))
            if c >= 4 and x not in defined:
                self.add('INFO', '高频未释义术语: "%s" %d次，建议确认是否入释义表' % (x, c))
        if '示例A股份' in full and '示例A新材' in full:
            self.add('INFO', '简称混用(示例A股份/示例A新材)，请统一为规范证券简称（按贵司常见混用对扩展此检查）')

    # ---------- [2] 数字指纹(启发式) ----------
    def check_fingerprint(self, texts):
        self.add('WARN', '数字指纹：请人工核对以下高价值指标全文是否一致——市场占有率、产能(已投产/规划)、毛利率、净利润同比增幅、白名单数量、磷酸铁锂产能（脚本不替代人工判断）')

    def run(self):
        import docx as _dx
        self.doc = _dx.Document(self.path)
        texts = [para_text(p) for p in self.doc.element.body.iter(qn('w:p'))]
        texts = [t for t in texts if t]
        tbls = self.tables()
        global BLOCK_SUM_TOL
        if self.config.get('block_sum_tol') is not None:
            BLOCK_SUM_TOL = float(self.config['block_sum_tol'])
        print('===== [1] 比率/勾稽复算 =====')
        for t in tbls:
            self.check_change_pct(t)
            self.check_tieout(t)
            self.check_cash(t)
        self.check_ratios(tbls)
        print('===== [1b] 分项求和=合计（容差%s） =====' % BLOCK_SUM_TOL)
        if self.config.get('skip_block_sum'):
            print('  SKIP 配置关闭该检查')
        else:
            for t in tbls:
                self.check_block_sum(t)
        print('===== [2] 数字指纹核对 =====')
        self.check_fingerprint(texts)
        print('===== [3] 称谓/释义审查 =====')
        self.check_terms(texts)
        fails = [i for i in self.issues if i[0] == 'FAIL']
        warns = [i for i in self.issues if i[0] == 'WARN']
        print('\n===== 汇总: FAIL=%d WARN=%d PASS=%d =====' % (len(fails), len(warns), len([i for i in self.issues if i[0] == 'PASS'])))
        for lv, m in self.issues:
            if lv in ('FAIL', 'WARN'):
                print('  %s %s' % ('❌' if lv == 'FAIL' else '⚠', m))
        return 2 if fails else (1 if warns else 0)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('--config', default=None)
    a = ap.parse_args()
    cfg = json.load(open(a.config, encoding='utf-8')) if a.config else None
    sys.exit(Checker(a.docx, cfg).run())

if __name__ == '__main__':
    main()

