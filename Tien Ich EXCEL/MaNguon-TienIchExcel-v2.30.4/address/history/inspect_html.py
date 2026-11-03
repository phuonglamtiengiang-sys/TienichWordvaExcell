import re,html,sys
from pathlib import Path
s=Path(sys.argv[1]).read_text(encoding='utf-8-sig')
if len(sys.argv)>2 and sys.argv[2]=='scripts':
    print('\n'.join(html.unescape(x) for x in re.findall(r'<script[^>]* src="([^"]+)',s)))
else:
    s=re.sub(r'<script.*?</script>','',s,flags=re.S)
    print(re.sub(r'\s+',' ',html.unescape(re.sub('<[^>]+>',' ',s)))[:6000])
