"""Compose explicit NSO changes with the 01/07/2025 mapping.

Never equate two footprints merely because they retain a code. Partial notes,
missing links and conflicting paths propagate a review requirement.
"""
from pathlib import Path
from collections import defaultdict, Counter
import json, re, unicodedata

ROOT=Path(__file__).parent
def read(name): return json.loads((ROOT/name).read_text(encoding='utf-8-sig'))
def norm(s):
    s=unicodedata.normalize('NFD',s.lower().replace('đ','d'))
    return re.sub(r'\s+',' ',''.join(c for c in s if not unicodedata.combining(c))).strip()
def bare(s):
    return re.sub(r'^(?:thanh pho|thi xa|thi tran|phuong|huyen|quan|tinh|xa) ','',norm(s)).replace('nt moc chau','nong truong moc chau')
def node(x):
    return (x['MaPhuongXa'],x['TenPhuongXa'],x['MaQuanHuyen'],x['TenQuanHuyen'],x['MaTinh'],x['TenTinh'])
def compare_node(row,new):
    if new: return (row[4],row[5],row[2],row[3],row[0],row[1])
    return (row[9],row[8],row[11],row[12],row[15],row[14])
def identity(n): return n[0],bare(n[1]),n[2],n[4]
def refs(text): return ';'.join(dict.fromkeys(re.findall(r'\d+/(?:NQ-UBTVQH\d+|\d+/QH\d+)',text)))

def sources_in_note(note, candidates):
    text=norm(note)
    text=re.sub(r'^.*?tren co so\s+','',text,count=1)
    text=re.split(r'\b(?:vao (?:xa|phuong|thi tran)|thanh (?:xa|phuong|thi tran)|co ten goi la)\b',text,maxsplit=1)[0]
    # Adjacent units and the location of the new office are not merger inputs.
    text=re.split(r'\b(?:sau khi (?:thanh lap|nhap|sap xep)|tru so|giap)\b',text,maxsplit=1)[0]
    if not any(w in text for w in ['nhap','hop nhat','toan bo','mot phan','doi ten','dieu chinh','con lai']): return []
    matched=[]; occupied=[]
    for n in sorted(candidates,key=lambda n:len(bare(n[1])),reverse=True):
        name=bare(n[1]); full=norm(n[1])
        pattern=full if name.isdigit() else name
        for match in re.finditer(r'(?<!\w)'+re.escape(pattern)+r'(?!\w)',text):
            if any(match.start()<b and match.end()>a for a,b in occupied): continue
            occupied.append(match.span()); matched.append(n); break
    return matched

def make_stage(old_nodes,new_nodes,rows):
    graph=defaultdict(list)
    old_by_id={identity(n):n for n in old_nodes}
    new_by_code={n[0]:n for n in new_nodes}
    changed_old=set()
    changed_targets=set()
    for row in rows:
        if row[9]: changed_old.add(identity(compare_node(row,False)))
        if row[4]: changed_targets.add(row[4])
    for row in rows:
        if not row[4]: continue
        target=compare_node(row,True)
        if target[0] not in new_by_code: raise ValueError('Missing target snapshot '+str(target))
        scope='partial' if re.search(r'mot phan|1 phan|con lai|sau khi (?:dieu chinh|sap xep)',norm(row[16])) else 'full'
        if scope=='full' and not re.search(r'toan bo|sap nhap|hop nhat|doi ten',norm(row[16])):
            scope='unknown'
        legal=refs(row[6]); direct=compare_node(row,False) if row[9] else None
        # Notes/decree fields are not always versioned on the NSO website.
        # A July-2025 decree attached to a January-2025 row cannot prove a
        # pre-July merger. Keep only a conservative direct candidate for review.
        future_note=bool(re.search(r'\b16[5-8]\d/NQ-UBTVQH15',row[6]))
        if future_note:
            scope='unknown'
        if direct and identity(direct) in old_by_id:
            graph[identity(direct)].append((target,scope,legal,'explicit comparison'))
        if future_note: continue
        # Restrict names to the old district(s) in this actual comparison row.
        parents={row[2],row[11]}- {''}
        candidates=[n for n in old_nodes if n[4] in {row[0],row[15]} and n[2] in parents]
        for source in sources_in_note(row[16],candidates):
            graph[identity(source)].append((target,scope,legal,'NSO operative note'))
        # Some mergers keep the destination name but adopt the incoming code.
        # Only link the former destination when the operative note names it.
        for source in candidates:
            if source[0] not in new_by_code and bare(source[1])==bare(target[1]) and norm(source[1]) in norm(row[16]) and 'vao ' in norm(row[16]):
                graph[identity(source)].append((target,scope,legal,'NSO named retained destination'))
    for old in old_nodes:
        key=identity(old)
        if key in graph: continue
        new=new_by_code.get(old[0])
        if new and identity(new)==key and key not in changed_old and new[0] not in changed_targets:
            graph[key].append((new,'full','','unchanged between snapshots'))
    return graph

def bac_giang_nodes():
    # Both current services reuse province codes 24/27 across epochs. Resolve
    # this supplement by the ten historical district codes, confirmed in NQ1191.
    district_codes={'213','215','216','217','218','219','220','221','222','223'}
    return list(dict.fromkeys((x['code'],x['name'],x['districtCode'],x['districtName'],'24','Tỉnh Bắc Giang')
            for x in read('casso-2024.json')['communes'] if x['districtCode'] in district_codes))

def bac_giang_stage(old_nodes,new_nodes):
    # NQ 1191, Articles 1 and 3: complete mergers, effective 01/01/2025.
    groups=[('Thành phố Bắc Giang','Ngô Quyền',['Trần Nguyên Hãn']),
            ('Thành phố Bắc Giang','Trần Phú',['Lê Lợi']),
            ('Huyện Yên Dũng','Tân An',['Lão Hộ','Thị trấn Tân An']),
            ('Huyện Hiệp Hòa','Sơn Thịnh',['Quang Minh','Hòa Sơn','Đại Thành']),
            ('Huyện Hiệp Hòa','Hùng Thái',['Hùng Sơn','Thái Sơn']),
            ('Huyện Hiệp Hòa','Toàn Thắng',['Hoàng Lương','Hoàng Thanh']),
            ('Huyện Hiệp Hòa','Hoàng Vân',['Hoàng An']),
            ('Huyện Hiệp Hòa','Đồng Tiến',['Đồng Tân','Thanh Vân']),
            ('Huyện Lạng Giang','Tiên Lục',['Mỹ Hà']),
            ('Huyện Lạng Giang','Hương Lạc',['Yên Mỹ']),
            ('Huyện Lục Nam','Cẩm Lý',['Vũ Xá']),
            ('Huyện Tân Yên','Quang Trung',['Đại Hóa','Lan Giới','Quang Tiến']),
            ('Huyện Tân Yên','Lam Sơn',['Phúc Sơn','Lam Cốt']),
            ('Huyện Yên Thế','Tân Hiệp',['Tam Hiệp']),
            ('Huyện Yên Thế','Đồng Kỳ',['Hồng Kỳ'])]
    graph=defaultdict(list); new_by_code={n[0]:n for n in new_nodes}
    for district,target_name,names in groups:
        targets=[n for n in new_nodes if bare(n[1])==norm(target_name) and
                 (bare(n[3])==bare(district) or bare(district)=='yen dung' and bare(n[3])=='bac giang')]
        assert len(targets)==1,(district,target_name,targets)
        for name in names:
            source=[n for n in old_nodes if (norm(n[1])==norm(name) if norm(name).startswith('thi tran ') else bare(n[1])==norm(name)) and bare(n[3])==bare(district)]
            assert len(source)==1,(district,name,source)
            graph[identity(source[0])].append((targets[0],'full','1191/NQ-UBTVQH15','NQ1191 complete merger'))
        if target_name!='Tân An':
            for source in old_nodes:
                if bare(source[1])==norm(target_name) and bare(source[3])==bare(district):
                    graph[identity(source)].append((targets[0],'full','1191/NQ-UBTVQH15','NQ1191 retained destination'))
    for old in old_nodes:
        if identity(old) in graph: continue
        new=new_by_code.get(old[0])
        if new:
            # Article 2: Thanh Hai loses a portion to Bien Son; never mark as full.
            scope='partial' if bare(old[1]) in {'thanh hai','bien son'} and bare(old[3])=='luc ngan' else 'full'
            graph[identity(old)].append((new,scope,'1191/NQ-UBTVQH15','NQ1191 retained unit / district change'))
            if bare(old[1])=='thanh hai' and bare(old[3])=='luc ngan':
                extra=[n for n in new_nodes if bare(n[1])=='bien son']
                assert len(extra)==1
                graph[identity(old)].append((extra[0],'partial','1191/NQ-UBTVQH15','NQ1191 Article2 partial transfer'))
    return graph

def build(base_rows):
    # NSO SOAP currently omits historical province 24; supplement against NQ1191.
    op={x['code']:x for x in json.loads((ROOT.parent/'provinces-legacy.json').read_text(encoding='utf-8'))['data']}
    od={x['code']:x for x in json.loads((ROOT.parent/'districts-legacy.json').read_text(encoding='utf-8'))['data']}
    def canonical(n):
        p=od[n[2]]['province_code'] if n[2] in od else n[4]
        province=op[p]
        return n[:4]+(p,province['type']+' '+province['name'])
    snapshots=[[canonical(node(x)) for x in read('snapshot-'+date+'.json')] for date in ['01-01-2024','01-01-2025','30-06-2025']]
    old_bg=bac_giang_nodes()
    comparisons=[read('compare-01-01-2024-01-01-2025.json'),read('compare-01-01-2025-30-06-2025.json')]
    for epoch in [0,1]:
        old_parents={}
        for rows in comparisons[epoch:][::-1]:
            for row in rows:
                if row[11]: old_parents[row[11]]=row[12]
        snapshots[epoch]=[(n[0],n[1],n[2],old_parents.get(n[2],n[3]),n[4],
                           'Tỉnh Thừa Thiên Huế' if epoch==0 and n[4]=='46' else n[5]) for n in snapshots[epoch]]
    # Comparison rows carry historical parent metadata; prefer it over the SOAP
    # join to today's province table, which reuses codes across epochs.
    for i,rows in enumerate(comparisons):
        for side,epoch in [(False,i),(True,i+1)]:
            replacements={}
            for row in rows:
                n=compare_node(row,side)
                if n[0]: replacements[(n[0],bare(n[1]))]=canonical(n)
            snapshots[epoch]=[replacements.get((n[0],bare(n[1])),n) for n in snapshots[epoch]]
    bg=[]
    for x in json.loads((ROOT.parent/'wards-legacy.json').read_text(encoding='utf-8'))['data']:
        if x['province_code']=='24':
            district=od[x['district_code']]; province=op['24']
            bg.append((x['code'],x['type']+' '+x['name'],district['code'],district['type']+' '+district['name'],'24',province['type']+' '+province['name']))
    # Normalize comparison province labels before matching them to snapshots.
    for rows in comparisons:
        for row in rows:
            for side in [False,True]:
                n=compare_node(row,side)
                if not n[0]: continue
                c=canonical(n)
                if side: row[0],row[1]=c[4],c[5]
                else: row[15],row[14]=c[4],c[5]
    stages=[make_stage(snapshots[0],snapshots[1],comparisons[0]),make_stage(snapshots[1],snapshots[2],comparisons[1])]
    bg_graph=bac_giang_stage(old_bg,bg)
    old_bg=[n for n in old_bg if identity(n) in bg_graph]
    print('Verified Bac Giang source nodes:',len(old_bg))
    stages[0].update(bg_graph)
    for n in bg: stages[1][identity(n)]=[(n,'full','','unchanged in first half of 2025')]
    snapshots[0]=list(dict.fromkeys(snapshots[0]+old_bg))
    snapshots[1]=list(dict.fromkeys(snapshots[1]+bg))
    snapshots[2]=list(dict.fromkeys(snapshots[2]+bg))
    base_by_code=defaultdict(list)
    for row in base_rows: base_by_code[row[0]].append(row)
    existing={(r[0],bare(r[1]),bare(r[2]),bare(r[3]),r[7],r[6]) for r in base_rows}
    additions=[]; audit=[]; unknown=[]
    for epoch in [0,1]:
        for source in snapshots[epoch]:
            sourcekey=(source[0],bare(source[1]),bare(source[3]),bare(source[5]))
            source_refs=[]
            for rows in comparisons[epoch:]:
                for row in rows:
                    if row[9]==source[0] and row[10]:
                        source_refs.extend(refs(row[10]).split(';'))
            paths=[(source,'full',list(dict.fromkeys(x for x in source_refs if x)),[])]
            for stage in stages[epoch:]:
                following=[]
                for current,scope,legal,trace in paths:
                    if current is None:
                        following.append((None,'unknown',legal,trace)); continue
                    links=stage.get(identity(current),[])
                    if not links:
                        unknown.append({'source':source,'epoch':epoch,'missing_at':current})
                        following.append((None,'unknown',legal,trace)); continue
                    for target,edge_scope,ref,reason in links:
                        next_scope='unknown' if 'unknown' in {scope,edge_scope} else ('full' if scope==edge_scope=='full' else 'partial')
                        following.append((target,next_scope,legal+[ref] if ref else legal,trace+[(current,target,edge_scope,ref,reason)]))
                paths=following
            for final,scope,legal,trace in paths:
                matches=base_by_code.get(final[0],[]) if final else []
                # Code plus name and parents must match the pre-July base unit.
                matches=[r for r in matches if bare(r[1])==bare(final[1]) and bare(r[2])==bare(final[3]) and bare(r[3])==bare(final[5])]
                if not matches:
                    matches=[[source[0],source[1],source[3],source[5],'','','unknown','','','']]
                    scope='unknown'
                for target in matches:
                    finalscope='unknown' if scope=='unknown' else ('full' if scope==target[6]=='full' else 'partial')
                    signature=sourcekey+(target[7],finalscope)
                    if signature in existing: continue
                    existing.add(signature)
                    evidence=';'.join(dict.fromkeys(legal+[target[8]]))
                    row=[source[0],source[1],source[3],source[5],target[4],target[5],finalscope,target[7],evidence,'NSO 2024-2025 + 2026.10.8']
                    additions.append(row)
                    audit.append({'row':row,'source_date':['01/01/2024','01/01/2025'][epoch],'path':trace})
    (ROOT/'historical-additions.json').write_text(json.dumps(additions,ensure_ascii=False),encoding='utf-8')
    (ROOT/'composition-audit.json').write_text(json.dumps(audit,ensure_ascii=False,indent=2),encoding='utf-8')
    (ROOT/'unresolved-paths.json').write_text(json.dumps(unknown,ensure_ascii=False,indent=2),encoding='utf-8')
    print('Historical additions',len(additions),dict(Counter(r[6] for r in additions)), 'unique sources',len({tuple(r[:4]) for r in additions}))
    return additions
