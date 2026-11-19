# -*- coding: utf-8 -*-
"""建筑市场行为检查主流程：花名册 → 项目经理/安全员 → 平台核查 → 结果 JSON

用法:
    python market_check.py <花名册目录> [输出JSON路径]

花名册为「合同用工备案花名册」，常见形态是**网页导出的 HTML 表**（扩展名可能是 .xls）。
本脚本同时兼容两种情形：真 HTML 表 / 真 Excel（后者需自行先用 pandas 转成行列表）。
"""
import sys, io, os, re, html as H, json
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from jzsc_query import (query, query_all_pages, exact_hits, save_cache,
                        PID_REG, PID_SAFE, PID_TECH)

CTYPE = {'ZB': '施工总承包', 'ZY': '专业分包', 'LW': '劳务分包', 'JL': '监理/咨询'}

# 岗位工种里含有这些关键词 → 视为安全管理人员
SAFE_KEYS = ('安全',)
# 视为项目经理/项目负责人的岗位
PM_KEYS = ('项目负责人', '项目经理', '项目副经理')


# ---------------- 花名册解析 ----------------
def parse_roster_html(path):
    txt = open(path, 'rb').read().decode('utf-8', 'replace')
    txt = re.sub(r'<style[\s\S]*?</style>', '', txt, flags=re.I)
    txt = re.sub(r'<script[\s\S]*?</script>', '', txt, flags=re.I)
    rows = []
    for r in re.findall(r'<tr[\s\S]*?</tr>', txt, flags=re.I):
        cells = []
        for _a, c in re.findall(r'<t[dh]\b([^>]*)>([\s\S]*?)</t[dh]>', r, flags=re.I):
            c = re.sub(r'<br\s*/?>', ' ', c, flags=re.I)
            c = re.sub(r'<[^>]+>', '', c)
            c = H.unescape(c)
            c = re.sub(r'[\r\n\t\u00a0]+', ' ', c)
            cells.append(re.sub(r' {2,}', ' ', c).strip())
        if cells:
            rows.append(cells)
    return rows


def is_xls_binary(path):
    with open(path, 'rb') as f:
        return f.read(8) == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'


def load_units(folder):
    """返回 [{proj, ent, cn, no, types, ctype, decl, ppl}]"""
    units = []
    for fn in sorted(os.listdir(folder)):
        if not fn.lower().endswith(('.xls', '.xlsx', '.html', '.htm')):
            continue
        p = os.path.join(folder, fn)
        if fn.lower().endswith(('.xlsx',)) or (fn.lower().endswith('.xls') and is_xls_binary(p)):
            print(f'!! 跳过真 Excel 文件（需先转 HTML 或改用 pandas 读取）: {fn}')
            continue
        proj = os.path.splitext(fn)[0]
        rows = parse_roster_html(p)
        blocks, cur = [], None
        for r in rows:
            for c in [x for x in r if x]:
                m = re.fullmatch(r'企业名称[：:]\s*(.+)', c)
                if m:
                    cur = {'proj': proj, 'ent': m.group(1).strip(), 'cn': [], 'no': [],
                           'decl': 0, 'ppl': []}
                    blocks.append(cur)
                    break
                if cur is None:
                    break
                m = re.fullmatch(r'合同名称[：:]\s*(.+)', c)
                if m:
                    cur['cn'] = [x.strip() for x in m.group(1).split(',') if x.strip()]
                m = re.fullmatch(r'合同编号[：:]\s*(.+)', c)
                if m:
                    cur['no'] = [x.strip() for x in m.group(1).split(',') if x.strip()]
                m = re.fullmatch(r'备案人数[：:]\s*(\d+)', c)
                if m:
                    cur['decl'] = int(m.group(1))
        cur = None
        for r in rows:
            cells = [x for x in r if x]
            if cells and cells[0].startswith('企业名称'):
                e = re.split(r'[：:]', cells[0], maxsplit=1)[-1].strip()
                cur = next((b for b in blocks if b['ent'] == e), None)
                continue
            if len(r) >= 14 and r[0].strip().isdigit() and cur is not None:
                cur['ppl'].append(r)
        for b in blocks:
            types = []
            for no in b['no']:
                m = re.search(r'(ZB|ZY|LW|JL)\d*$', no)
                if m:
                    types.append(m.group(1))
            b['types'] = sorted(set(types))
            b['ctype'] = [CTYPE[t] for t in b['types']] or ['未知']
            units.append(b)
    return units


def is_pm(p):
    return any(k in (p[6] or '') for k in PM_KEYS)


def is_safe(p):
    return any(k in (p[6] or '') for k in SAFE_KEYS)


# ---------------- 核查 ----------------
def check_person(name, ent, max_pages=3):
    """返回 dict: reg/safe = 本单位名下命中；other = 同名但挂在其他单位的记录

    max_pages: 全库同名检索只翻前几页。重名率高的姓名可能返回几十条，
               结果会有遗漏，但足以证明「该企业名下无记录」并找出他单位挂靠线索。
    """
    _, reg_hits = exact_hits(PID_REG, name, ent)
    _, safe_hits = exact_hits(PID_SAFE, name, ent)

    other = []
    for tag, pid, nf in (('注册', PID_REG, 'xm'), ('安管', PID_SAFE, 'workername')):
        _, rows = query_all_pages(pid, max_pages=max_pages, **{nf: name})
        for r in rows:
            if len(r) < 4 or r[0] == 'ERROR' or r[1].strip() != name:
                continue
            e = r[2].strip()
            if not (ent in e or e in ent):
                other.append(f'[{tag}] {r[3]} @ {e}')
    return {'reg': reg_hits, 'safe': safe_hits, 'other': sorted(set(other))}


def run(folder):
    units = load_units(folder)
    out = []
    for b in units:
        ent = b['ent']
        rec = {'proj': b['proj'], 'ent': ent, 'ctype': '、'.join(b['ctype']),
               'codes': b['no'], 'decl': b['decl'], 'actual': len(b['ppl']),
               'pm': [], 'safe': []}
        for p in [x for x in b['ppl'] if is_pm(x)]:
            r = check_person(p[2], ent)
            rec['pm'].append({'name': p[2], 'id': p[8], 'role': p[6], **r})
        for p in [x for x in b['ppl'] if is_safe(x)]:
            r = check_person(p[2], ent)
            rec['safe'].append({'name': p[2], 'id': p[8], 'role': p[6],
                                'ptype': p[5], **r})
        out.append(rec)
        print(f"✓ {ent}  (项目经理{len(rec['pm'])} / 安全员{len(rec['safe'])})", flush=True)
        save_cache()          # 每个单位存一次，长任务中断后可复用缓存续跑
    save_cache()
    return out


if __name__ == '__main__':
    # write_through=True 很关键：否则重包 stdout 会变成块缓冲，长时间跑看不到任何进度
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                                  errors='replace', write_through=True)
    folder = sys.argv[1] if len(sys.argv) > 1 else '.'
    dest = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, 'check_result.json')
    res = run(folder)
    json.dump(res, open(dest, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('\n完成 →', dest)
