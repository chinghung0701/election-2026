import pandas as pd, re, json, unicodedata

UP = '/mnt/user-data/uploads/'
# 中選會檔案以私用區碼點存 Big5 缺字；對應表由 village-2018.js 的正確字名比對推導，
# 0xe01d 由三個名稱的交集確認為「廍」，0xe01f 由 Big5 缺字性質確認為「獇」。
PUA = {chr(int(k,16)): v for k, v in {"0xe001": "𥕢", "0xe008": "磘", "0xe435": "坔", "0xe6ff": "峯", "0xe8d7": "脚", "0xee8e": "濓", "0xeebe": "舘", "0xe808": "堃", "0xe01d": "廍", "0xe01f": "獇"}.items()}

def clean(s):
    if s is None: return ''
    if isinstance(s, float) and pd.isna(s): return ''
    s = str(s)
    for k, v in PUA.items(): s = s.replace(k, v)
    return s.replace('\u3000', ' ').strip()

def parse_cand_header(cell):
    """'1\\n游錫堃\\n民主進步黨' or '(1)\\n朱立倫\\n王如玄' -> (name, party|None)"""
    parts = [clean(p) for p in clean(cell).split('\n')]
    parts = [p for p in parts if p and not re.fullmatch(r'\(?\d+\)?', p)]
    return parts

# ---------------- 2016 總統（行政區層級） ----------------
def parse_pres(path):
    d = pd.read_excel(path, engine='xlrd', header=None)
    hdr = [parse_cand_header(d.iat[2, c]) for c in (1, 2, 3)]
    cands = [h[0] for h in hdr]                 # 正／副中的總統候選人
    rows, total = [], None
    for i in range(3, len(d)):
        name = clean(d.iat[i, 0])
        if not name: continue
        if pd.isna(d.iat[i, 1]): continue
        votes = [int(d.iat[i, c]) for c in (1, 2, 3)]
        if name.replace(' ', '') == '總計':
            total = votes; continue
        rows.append((name, votes))
    return cands, rows, total

# ---------------- 2014 縣市長（行政區 + 村里） ----------------
def parse_mayor(path):
    d = pd.read_excel(path, engine='xlrd', header=None)
    ncand = 0
    for c in range(3, 14):
        if parse_cand_header(d.iat[2, c]): ncand += 1
        else: break
    cols = list(range(3, 3 + ncand))
    hdr = [parse_cand_header(d.iat[2, c]) for c in cols]
    cands = [h[0] for h in hdr]
    parties = {h[0]: (h[1] if len(h) > 1 and h[1] else '無黨籍及未經政黨推薦') for h in hdr}

    dist, vill, cur, lastv, total = [], [], None, None, None
    for i in range(3, len(d)):
        a, b = clean(d.iat[i, 0]), clean(d.iat[i, 1])
        votes = None
        if not pd.isna(d.iat[i, cols[0]]):
            votes = [int(d.iat[i, c]) for c in cols]
        if a and votes is not None:                       # 行政區小計（或總計）
            if a.replace(' ', '') == '總計': total = votes
            else:
                cur = a; lastv = None; dist.append((a, votes))
        elif votes is not None and cur:                   # 投開票所 -> 併入村里
            if b: lastv = b                               # 空白＝垂直合併儲存格的延續列
            if lastv: vill.append((cur, lastv, votes))
    # 村里逐所加總
    agg = {}
    order = []
    for dn, vn, v in vill:
        k = (dn, vn)
        if k not in agg:
            agg[k] = [0]*ncand; order.append(k)
        for j, x in enumerate(v): agg[k][j] += x
    vill = [(k[0], k[1], agg[k]) for k in order]
    return cands, parties, dist, vill, total

def pack(cands, votes):
    tot = sum(votes)
    return [{"name": cands[j], "votes": votes[j],
             "rate": round(votes[j]/tot*100, 2) if tot else 0.0} for j in range(len(cands))]

def js(o): return json.dumps(o, ensure_ascii=False)

# ================= 2016 =================
P16 = {}
for code, cty, f in [('NTC','新北市','總統-A05-2-_新北市_.xls'),
                     ('KHH','高雄市','總統-A05-2-_高雄市_.xls')]:
    cands, rows, total = parse_pres(UP+f)
    P16[code] = dict(county=cty, cands=cands, rows=rows, total=total)
    s = [sum(r[1][j] for r in rows) for j in range(3)]
    print(f'2016 {cty}: {len(rows)} 區　小計合計 {s}　檔案總計 {total}　{"✓" if s==total else "✗ 不符"}')

# ================= 2014 =================
P14 = {}
for code, cty, f in [('NTC','新北市','縣表3-1-200_新北市_-候選人得票數一覽表.xls'),
                     ('KHH','高雄市','縣表3-1-600_高雄市_-候選人得票數一覽表.xls')]:
    cands, parties, dist, vill, total = parse_mayor(UP+f)
    P14[code] = dict(county=cty, cands=cands, parties=parties, dist=dist, vill=vill, total=total)
    n = len(cands)
    ds = [sum(r[1][j] for r in dist) for j in range(n)]
    vs = [sum(r[2][j] for r in vill) for j in range(n)]
    print(f'2014 {cty}: {len(dist)} 區 / {len(vill)} 村里　候選人 {cands}')
    print(f'     政黨 {parties}')
    print(f'     行政區合計 {ds}')
    print(f'     村里合計　 {vs}　{"✓ 一致" if ds==vs else "✗ 不符"}')
    if total: print(f'     檔案總計   {total}　{"✓" if ds==total else "✗ 不符"}')

# ---------------- 輸出 ----------------
def write_pres(path, var, year, data):
    L = [f'// {year} 總統選舉 鄉鎮市區得票資料（完整候選人，未套用顯示規則）',
         f'// 目前僅含新北市、高雄市；其餘 20 縣市待補。',
         f'var {var} = {{']
    for code, d in data.items():
        L += [f'  {code}: {{', f'    county: "{d["county"]}",',
              f'    cands: {js(d["cands"])},', '    districts: [']
        for name, v in d['rows']:
            L.append(f'      {{d:{js(name)},c:{js(pack(d["cands"], v))}}},')
        L += ['    ]', '  },']
    L += ['};', '']
    open(path, 'w', encoding='utf-8').write('\n'.join(L))

def write_mayor(path, var, vpath, vvar, year, data):
    L = [f'// {year} 縣市長選舉 行政區得票資料（完整候選人，未套用顯示規則）',
         f'// 注意：村里層級資料已拆分至 {vpath.split("/")[-1]}（var {vvar}），本檔僅含行政區層級。',
         f'//       county.hasVil = true 表示該縣市在 {vpath.split("/")[-1]} 中有村里資料。',
         f'// 目前僅含新北市、高雄市；其餘 20 縣市待補。',
         f'var {var} = {{']
    V = [f'// {year} 縣市長選舉 村里層級得票資料（自 {path.split("/")[-1]} 拆出）',
         f'// 對應 {path.split("/")[-1]}；僅在使用者切換到「村里」層級時才需載入。',
         '// 結構：{ <縣市代碼>: [ {d:行政區, v:村里, c:[{name,votes,rate}]} ] }',
         f'var {vvar} = {{']
    for code, d in data.items():
        L += [f'  {code}: {{', f'    county: "{d["county"]}",',
              f'    cands: {js(d["cands"])},',
              f'    parties: {js(d["parties"])},', '    hasVil: true,', '    districts: [']
        for name, v in d['dist']:
            L.append(f'      {{"d":{js(name)},"c":{js(pack(d["cands"], v))}}},')
        L += ['    ]', '  },']
        V.append(f'  {code}: [')
        for dn, vn, v in d['vill']:
            V.append(f'    {{"d":{js(dn)},"v":{js(vn)},"c":{js(pack(d["cands"], v))}}},')
        V.append('  ],')
    L += ['};', '']
    V += ['};', '']
    open(path, 'w', encoding='utf-8').write('\n'.join(L))
    open(vpath, 'w', encoding='utf-8').write('\n'.join(V))

import os
os.makedirs('/home/claude/out', exist_ok=True)
write_pres('/home/claude/out/district-2016.js', 'DIST2016', 2016, P16)
write_mayor('/home/claude/out/district-2014.js', 'DIST2014',
            '/home/claude/out/village-2014.js', 'VILL2014', 2014, P14)
print('\n輸出完成')
for f in sorted(os.listdir('/home/claude/out')):
    print(' ', f, os.path.getsize('/home/claude/out/'+f)//1024, 'KB')
