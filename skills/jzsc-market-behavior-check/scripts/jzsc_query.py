# -*- coding: utf-8 -*-
"""北京市建设行业职业人员信息查询 —— 查询客户端

入口页面:
  http://bjjs.zjw.beijing.gov.cn/bjjs/sy/zwfw66/cxzx-sjzt/gcjsl/jsxyzyryxxcx58/index.shtml

接口（GET 或 POST 均可，无需登录/验证码/Cookie）:
  /eportal/ui?pageId=<pid>&currentPage=<n>&filter_LIKE_<字段>=<值>

页签 pageId:
  53614300  注册人员          （一级/二级建造师、监理工程师、造价工程师）
  53614302  安管人员          （安全生产考核合格证 A/B/C1/C2/C3）
  53614304  建筑施工特种作业人员
  53614306  施工现场专业人员    （施工员、质检员、资料员等）
  53614308  施工现场技术工人
  53615694  （新版）技能人员
  53614392  其他
  53614310  按人员
  53614329  按企业

各页签的 filter 字段名（见结果页 <input id="filter_LIKE_xxx">）:
  注册人员 : xm(姓名) / zjhm(身份证件号) / qymc(企业名称) / zclx(人员类型) / zcdj(注册等级)
  安管人员 : workername(姓名) / workercertificatecode(证件号) / unitname(企业名称) / postname(人员类型)
             certificatecode(证书编号)
  专业人员 : 同「安管人员」

⚠️ 解析要点：结果表 <table class="xx"> 内数据行**缺少 </tr> 闭合标签**，且「序号」列由 JS 回填（服务端为空）。
   因此不能按 <tr> 切分，必须取表内全部 <td>，再按表头 <th> 数量分块。

⚠️ 匹配语义：所有 filter 均为 **LIKE 模糊匹配**。查「张伟」会返回「张伟明/张伟华/王张伟」。
   必须再用「姓名 + 企业名称」双条件收紧，并在客户端做精确姓名相等 + 企业名互相包含的过滤。
"""
import re, html as H, urllib.parse, urllib.request, time, os, json, sys

BASE = 'http://bjjs.zjw.beijing.gov.cn/eportal/ui?pageId='
REF = ('http://bjjs.zjw.beijing.gov.cn/bjjs/sy/zwfw66/cxzx-sjzt/gcjsl/'
       'jsxyzyryxxcx58/index.shtml')
UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
      '(KHTML, like Gecko) Chrome/120.0 Safari/537.36')

PID_REG = 53614300
PID_SAFE = 53614302
PID_SPECIAL = 53614304
PID_TECH = 53614306
PID_WORKER = 53614308

CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'cache.json')
_cache = json.load(open(CACHE, encoding='utf-8')) if os.path.exists(CACHE) else {}


def _parse(h):
    hs = re.sub(r'<script[\s\S]*?</script>', '', h, flags=re.I)
    tot = re.search(r'总记录数[：:]\s*(\d+)', re.sub(r'<[^>]+>', ' ', hs))
    total = int(tot.group(1)) if tot else -1
    m = re.search(r'<table class="xx"[\s\S]*?</table>', hs, re.I)
    rows = []
    if not m:
        return total, rows
    body = m.group(0)
    ncol = len(re.findall(r'<th\b', body, re.I))
    if ncol == 0:
        return total, rows
    cells = [re.sub(r'\s+', ' ', H.unescape(re.sub(r'<[^>]+>', '', c))).strip()
             for c in re.findall(r'<td[^>]*>([\s\S]*?)</td>', body, re.I)]
    for i in range(0, len(cells) - ncol + 1, ncol):
        grp = cells[i:i + ncol]
        if any(grp):
            rows.append(grp)
    return total, rows


def query(pid, page=1, sleep=0.35, retries=3, **filters):
    """按 filter_LIKE_<key>=<value> 查询。key 用页签对应的字段名。返回 (总记录数, [行])"""
    kw = {('filter_LIKE_' + k): v for k, v in filters.items() if v}
    key = f'{pid}|{page}|' + urllib.parse.urlencode(sorted(kw.items()))
    if key in _cache:
        return _cache[key][0], [list(r) for r in _cache[key][1]]
    url = BASE + str(pid) + '&currentPage=' + str(page) + '&' + urllib.parse.urlencode(kw)
    req = urllib.request.Request(url, headers={'User-Agent': UA, 'Referer': REF})
    h = None
    for a in range(retries):
        try:
            h = urllib.request.urlopen(req, timeout=45).read().decode('utf-8', 'replace')
            break
        except Exception as e:
            if a == retries - 1:
                return -1, [['ERROR', str(e)]]
            time.sleep(1.5)
    total, rows = _parse(h)
    _cache[key] = [total, rows]
    time.sleep(sleep)
    return total, rows


def query_all_pages(pid, max_pages=20, **filters):
    """翻页取回全部结果"""
    out, total = [], None
    for p in range(1, max_pages + 1):
        t, rows = query(pid, page=p, **filters)
        if total is None:
            total = t
        if not rows:
            break
        out.extend(rows)
        if t <= len(out):
            break
    return total, out


def exact_hits(pid, name, ent, namecol=1, entcol=2, **extra):
    """精确匹配：姓名完全相等 + 企业名称互相包含。返回命中行。"""
    nf = 'workername' if pid in (PID_SAFE, PID_SPECIAL, PID_TECH, PID_WORKER) else 'xm'
    ef = 'unitname' if pid in (PID_SAFE, PID_SPECIAL, PID_TECH, PID_WORKER) else 'qymc'
    kw = {nf: name, ef: ent}
    kw.update(extra)
    t, rows = query(pid, **kw)
    hits = []
    for r in rows:
        if len(r) < 3 or r[0] == 'ERROR':
            continue
        if r[namecol].strip() != name:
            continue
        e = r[entcol].strip()
        if ent in e or e in ent:
            hits.append(r)
    return t, hits


def save_cache():
    json.dump(_cache, open(CACHE, 'w', encoding='utf-8'), ensure_ascii=False)


if __name__ == '__main__':
    import io as _io
    sys.stdout = _io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                                    errors='replace', write_through=True)
    # 自检：python jzsc_query.py [姓名] [企业名称]
    name = sys.argv[1] if len(sys.argv) > 1 else '张伟'
    ent = sys.argv[2] if len(sys.argv) > 2 else ''

    t, r = query(PID_REG, xm=name)
    print(f'### 注册人员 xm={name}  总记录数={t}')
    for x in r[:6]:
        print('   ', ' | '.join(x))

    if ent:
        t, r = query(PID_SAFE, workername=name, unitname=ent)
        print(f'### 安管人员 workername={name} + unitname={ent}  总记录数={t}')
        for x in r[:6]:
            print('   ', ' | '.join(x))
    else:
        print('（传入第 2 个参数「企业名称」可同时验证安管人员双条件查询）')
    save_cache()
