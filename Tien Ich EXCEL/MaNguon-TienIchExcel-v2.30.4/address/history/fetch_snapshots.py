from pathlib import Path
import urllib.request, xml.etree.ElementTree as ET, json

ROOT=Path(__file__).parent
URL='https://danhmuchanhchinh.nso.gov.vn/DMDVHC.asmx'
def fetch(date):
    path=ROOT/('snapshot-'+date.replace('/','-')+'.xml')
    if not path.exists():
        body=('<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'
              '<soap:Body><DanhMucPhuongXa xmlns="http://tempuri.org/"><DenNgay>'+date+'</DenNgay><Tinh></Tinh><TenTinh></TenTinh>'
              '<QuanHuyen></QuanHuyen><TenQuanHuyen></TenQuanHuyen></DanhMucPhuongXa></soap:Body></soap:Envelope>').encode()
        request=urllib.request.Request(URL,data=body,headers={'Content-Type':'text/xml; charset=utf-8','SOAPAction':'http://tempuri.org/DanhMucPhuongXa'})
        with urllib.request.urlopen(request,timeout=50) as response: path.write_bytes(response.read())
    nodes=[{c.tag:c.text or '' for c in x} for x in ET.parse(path).iter('TABLE')]
    assert len(nodes)>9000
    path.with_suffix('.json').write_text(json.dumps(nodes,ensure_ascii=False),encoding='utf-8')
    print(date,len(nodes),flush=True)

if __name__=='__main__':
    for date in ['01/01/2024','01/01/2025','30/06/2025']: fetch(date)
