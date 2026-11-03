"""Read-only export of the public NSO administrative comparison form."""
from pathlib import Path
from html.parser import HTMLParser
import urllib.request, urllib.parse, datetime, json, re, sys, http.cookiejar, ast

ROOT=Path(__file__).parent
BASE='https://danhmuchanhchinh.nso.gov.vn/'
class Inputs(HTMLParser):
    def __init__(self):
        super().__init__(); self.fields={}
    def handle_starttag(self,tag,attrs):
        a=dict(attrs)
        if tag=='input' and a.get('name') and a.get('type') not in ('submit','button','checkbox'):
            self.fields[a['name']]=a.get('value','')

class GridRows(HTMLParser):
    def __init__(self):
        super().__init__(); self.rows=[]; self.row=None; self.cell=None; self.depth=0
    def handle_starttag(self,tag,attrs):
        a=dict(attrs)
        if tag=='tr' and re.fullmatch(r'ctl00_PlaceHolderMain_gridXa_DXDataRow\d+',a.get('id','')):
            self.row=[]
        if self.row is not None and tag=='td': self.cell=[]
    def handle_data(self,data):
        if self.cell is not None: self.cell.append(data)
    def handle_endtag(self,tag):
        if tag=='td' and self.cell is not None:
            self.row.append(re.sub(r'\s+',' ',''.join(self.cell)).strip()); self.cell=None
        if tag=='tr' and self.row is not None:
            self.rows.append(self.row); self.row=None

def decode_callback(content):
    match=re.search(r"'result':('(?:\\.|[^'\\])*')",content)
    if not match: raise ValueError('No callback result: '+content[:300])
    return ast.literal_eval(match[1])

def post_compare(old='01/01/2024',new='30/06/2025'):
    url=BASE+'Doi_Chieu_Moi.aspx'
    opener=urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
    with opener.open(url,timeout=40) as response: initial=response.read().decode('utf-8-sig')
    parser=Inputs(); parser.feed(initial); fields=parser.fields
    fields['__EVENTTARGET']='ctl00$PlaceHolderMain$cmdThucHien'
    fields['__EVENTARGUMENT']='Click'
    fields['ctl00_PlaceHolderMain_cmbCap_VI']='3'
    fields['ctl00$PlaceHolderMain$cmbCap']='Xã'
    for name,date in [('txtNgay',new),('txtNgayDC',old)]:
        fields['ctl00$PlaceHolderMain$'+name]=date
        fields['ctl00_PlaceHolderMain_'+name+'_Raw']=str(int(datetime.datetime.strptime(date,'%d/%m/%Y').replace(tzinfo=datetime.timezone.utc).timestamp()*1000))
    # Keep all provinces selected; this is a read-only comparison query.
    fields['ctl00$PlaceHolderMain$chkCheck']='C'
    fields['ctl00$PlaceHolderMain$txtBoQH']='0'
    fields['ctl00$PlaceHolderMain$txtBoQH_DC']='0'
    fields['__EVENTTARGET']=''
    fields['__EVENTARGUMENT']=''
    fields['__CALLBACKID']='ctl00$PlaceHolderMain$gridXa'
    command='14|CUSTOMCALLBACK0|'
    fields['__CALLBACKPARAM']='GB|'+str(len(command))+';'+command+';'
    def query(command):
        fields['__CALLBACKPARAM']='GB|'+str(len(command))+';'+command+';'
        request=urllib.request.Request(url,data=urllib.parse.urlencode(fields).encode(),headers={'Content-Type':'application/x-www-form-urlencoded'})
        with opener.open(request,timeout=60) as response: content=response.read().decode('utf-8-sig')
        result=decode_callback(content)
        state=Inputs(); state.feed(result); fields.update(state.fields)
        return result
    result=query(command)
    stem='compare-'+old.replace('/','-')+'-'+new.replace('/','-')
    (ROOT/(stem+'-first.html')).write_text(result,encoding='utf-8')
    count_match=re.search(r'\((\d+) items\)',result)
    first_grid=GridRows(); first_grid.feed(result)
    total=int(count_match[1]) if count_match else len(first_grid.rows)
    pages_match=re.search(r'Page 1 of (\d+)',result)
    pages=int(pages_match[1]) if pages_match else 1
    allrows=[]
    for page in range(pages):
        if page:
            value=str(page); command='8|GOTOPAGE'+str(len(value))+'|'+value
            result=query(command)
        (ROOT/(stem+'-page'+str(page+1)+'.html')).write_text(result,encoding='utf-8')
        grid=GridRows(); grid.feed(result); allrows.extend(grid.rows)
        print('page',page+1,'of',pages,'rows',len(grid.rows),flush=True)
    assert len(allrows)==total,(len(allrows),total)
    (ROOT/(stem+'.json')).write_text(json.dumps(allrows,ensure_ascii=False,indent=2),encoding='utf-8')
    print('total',total,flush=True)

if __name__=='__main__': post_compare(*sys.argv[1:])
