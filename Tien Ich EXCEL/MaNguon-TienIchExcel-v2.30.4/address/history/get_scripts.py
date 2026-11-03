import re,html,urllib.request
from pathlib import Path
root=Path(__file__).parent
s=(root/'compare.html').read_text(encoding='utf-8')
for i,path in enumerate(re.findall(r'<script[^>]* src="([^"]+)',s)):
    url=urllib.parse.urljoin('https://danhmuchanhchinh.nso.gov.vn/',html.unescape(path))
    with urllib.request.urlopen(url,timeout=30) as response: content=response.read()
    (root/('script-'+str(i)+'.js')).write_bytes(content)
    print(i,len(content))
