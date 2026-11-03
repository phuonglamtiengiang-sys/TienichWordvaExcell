from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import math

ROOT=Path(__file__).resolve().parent.parent
OUT=ROOT/'output/branding'; OUT.mkdir(parents=True,exist_ok=True)
S=3
BLUE='#185ABD'; BLUE_DARK='#12458F'; GREEN='#107C41'; GREEN_DARK='#0A6032'; INK='#17324D'; GRAY='#526579'; LIGHT='#F5F9FD'
FONT_BOLD='C:/Windows/Fonts/seguisb.ttf'; FONT_REG='C:/Windows/Fonts/segoeui.ttf'

def px(v): return int(round(v*S))
def icon_base():
    W=600; H=600
    canvas=Image.new('RGBA',(px(W),px(H)),(0,0,0,0))
    # Draw two overlapping office documents on separate transparent layers.
    def card(kind,angle):
        layer=Image.new('RGBA',(px(290),px(370)),(0,0,0,0)); d=ImageDraw.Draw(layer)
        color=BLUE if kind=='word' else GREEN
        dark=BLUE_DARK if kind=='word' else GREEN_DARK
        d.rounded_rectangle((px(6),px(6),px(284),px(364)),radius=px(34),fill=color)
        d.rounded_rectangle((px(35),px(34),px(255),px(337)),radius=px(19),fill='#FFFFFF')
        # Folded page corner
        d.polygon([(px(202),px(34)),(px(255),px(87)),(px(202),px(87))],fill='#E6F0FF' if kind=='word' else '#E2F4E9')
        d.line([(px(202),px(35)),(px(202),px(87)),(px(254),px(87))],fill='#9AB9E5' if kind=='word' else '#87C69E',width=px(2))
        if kind=='word':
            font=ImageFont.truetype(FONT_BOLD,px(136))
            d.text((px(145),px(211)),'W',font=font,fill=color,anchor='mm',stroke_width=px(1),stroke_fill=color)
            d.rounded_rectangle((px(79),px(276),px(211),px(287)),radius=px(5),fill='#DCE9FA')
        else:
            # A compact spreadsheet grid with a strong X tile.
            d.rounded_rectangle((px(67),px(118),px(223),px(278)),radius=px(12),fill='#E9F6ED')
            for xx in (px(119),px(171)):
                d.line((xx,px(123),xx,px(273)),fill='#87C69E',width=px(4))
            for yy in (px(169),px(220)):
                d.line((px(72),yy,px(218),yy),fill='#87C69E',width=px(4))
            d.rounded_rectangle((px(74),px(125),px(118),px(166)),radius=px(7),fill=dark)
            f=ImageFont.truetype(FONT_BOLD,px(31));d.text((px(96),px(145)),'X',font=f,fill='white',anchor='mm')
        return layer.rotate(angle,resample=Image.Resampling.BICUBIC,expand=True)
    # Soft grounded shadow beneath the pair.
    shadow=Image.new('RGBA',canvas.size,(0,0,0,0)); sd=ImageDraw.Draw(shadow)
    sd.ellipse((px(82),px(478),px(520),px(545)),fill=(22,47,72,42))
    shadow=shadow.filter(ImageFilter.GaussianBlur(px(16)));canvas.alpha_composite(shadow)
    left=card('word',8);right=card('excel',-7)
    canvas.alpha_composite(left,(px(25),px(100)))
    canvas.alpha_composite(right,(px(250),px(94)))
    return canvas.resize((W,H),Image.Resampling.LANCZOS)

mark=icon_base()
mark.save(OUT/'TienIchOffice-VietNam-Logo.png',optimize=True)
mark.resize((256,256),Image.Resampling.LANCZOS).save(OUT/'TienIchOffice-VietNam-Icon.png',optimize=True)

# Horizontal brand lockup.
W,H=1680,520
lock=Image.new('RGBA',(px(W),px(H)),(255,255,255,0))
mark_small=mark.resize((px(450),px(450)),Image.Resampling.LANCZOS)
lock.alpha_composite(mark_small,(px(20),px(35)))
d=ImageDraw.Draw(lock)
font_title=ImageFont.truetype(FONT_BOLD,px(82));font_sub=ImageFont.truetype(FONT_BOLD,px(36));font_product=ImageFont.truetype(FONT_REG,px(29))
d.text((px(505),px(201)),'TIỆN ÍCH OFFICE',font=font_title,fill=INK,anchor='lm')
d.rounded_rectangle((px(510),px(265),px(699),px(273)),radius=px(4),fill=BLUE)
d.rounded_rectangle((px(707),px(265),px(896),px(273)),radius=px(4),fill=GREEN)
d.text((px(510),px(326)),'VIỆT NAM',font=font_sub,fill=GREEN,anchor='lm')
d.text((px(510),px(385)),'WORD  •  EXCEL',font=font_product,fill=GRAY,anchor='lm')
lock.resize((W,H),Image.Resampling.LANCZOS).save(OUT/'TienIchOffice-VietNam-Logo-Ngang.png',optimize=True)
white=Image.new('RGBA',lock.size,'white');white.alpha_composite(lock)
white.resize((W,H),Image.Resampling.LANCZOS).convert('RGB').save(OUT/'TienIchOffice-VietNam-Logo-Ngang-Trang.png',optimize=True)

# Editable vector master matching the mark and wordmark.
svg='''<svg xmlns="http://www.w3.org/2000/svg" width="1680" height="520" viewBox="0 0 1680 520">
<defs><filter id="shadow" x="-20%" y="-20%" width="140%" height="150%"><feGaussianBlur stdDeviation="10"/></filter></defs>
<ellipse cx="300" cy="535" rx="205" ry="27" fill="#162F48" opacity=".14" filter="url(#shadow)"/>
<g transform="translate(25 100) rotate(-8 145 185)">
 <rect x="6" y="6" width="278" height="358" rx="34" fill="#185ABD"/><rect x="35" y="34" width="220" height="303" rx="19" fill="#fff"/>
 <path d="M202 34h1l52 53h-53z" fill="#E6F0FF" stroke="#9AB9E5" stroke-width="2"/><text x="145" y="226" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif" font-size="136" font-weight="700" fill="#185ABD">W</text><rect x="79" y="276" width="132" height="11" rx="5" fill="#DCE9FA"/>
</g>
<g transform="translate(250 94) rotate(7 145 185)">
 <rect x="6" y="6" width="278" height="358" rx="34" fill="#107C41"/><rect x="35" y="34" width="220" height="303" rx="19" fill="#fff"/>
 <path d="M202 34h1l52 53h-53z" fill="#E2F4E9" stroke="#87C69E" stroke-width="2"/><rect x="67" y="118" width="156" height="160" rx="12" fill="#E9F6ED"/>
 <path d="M119 123v150m52-150v150M72 169h146m-146 51h146" stroke="#87C69E" stroke-width="4"/><rect x="74" y="125" width="44" height="41" rx="7" fill="#0A6032"/><text x="96" y="155" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif" font-size="31" font-weight="700" fill="#fff">X</text>
</g>
<text x="525" y="214" font-family="Segoe UI,Arial,sans-serif" font-size="82" font-weight="700" letter-spacing="1" fill="#17324D">TIỆN ÍCH OFFICE</text>
<rect x="510" y="265" width="189" height="8" rx="4" fill="#185ABD"/><rect x="707" y="265" width="189" height="8" rx="4" fill="#107C41"/>
<text x="510" y="338" font-family="Segoe UI,Arial,sans-serif" font-size="36" font-weight="700" letter-spacing="2" fill="#107C41">VIỆT NAM</text>
<text x="510" y="410" font-family="Segoe UI,Arial,sans-serif" font-size="29" fill="#526579">WORD  •  EXCEL</text>
</svg>'''
(OUT/'TienIchOffice-VietNam-Logo.svg').write_text(svg,encoding='utf-8')

def product_logo(kind):
    label='WORD' if kind=='word' else 'EXCEL'
    color=BLUE if kind=='word' else GREEN
    dark=BLUE_DARK if kind=='word' else GREEN_DARK
    pale='#E6F0FF' if kind=='word' else '#E2F4E9'
    W=600;S2=3
    mark2=Image.new('RGBA',(W*S2,W*S2),(0,0,0,0));d2=ImageDraw.Draw(mark2)
    d2.ellipse((px(78),px(500),px(522),px(548)),fill=(22,47,72,34))
    d2.rounded_rectangle((px(82),px(50),px(518),px(510)),radius=px(54),fill=color)
    d2.rounded_rectangle((px(140),px(108),px(460),px(454)),radius=px(28),fill='white')
    d2.polygon([(px(365),px(108)),(px(460),px(203)),(px(365),px(203))],fill=pale)
    d2.line([(px(365),px(109)),(px(365),px(203)),(px(459),px(203))],fill='#87C69E' if kind=='excel' else '#9AB9E5',width=px(3))
    if kind=='word':
        f=ImageFont.truetype(FONT_BOLD,px(205));d2.text((px(300),px(294)),'W',font=f,fill=color,anchor='mm',stroke_width=px(1),stroke_fill=color)
        d2.rounded_rectangle((px(205),px(375),px(395),px(390)),radius=px(7),fill='#DCE9FA')
    else:
        d2.rounded_rectangle((px(202),px(216),px(398),px(402)),radius=px(16),fill='#E9F6ED')
        for xx in (px(267),px(333)):d2.line((xx,px(224),xx,px(394)),fill='#87C69E',width=px(5))
        for yy in (px(278),px(340)):d2.line((px(210),yy,px(390),yy),fill='#87C69E',width=px(5))
        d2.rounded_rectangle((px(211),px(224),px(266),px(276)),radius=px(8),fill=dark)
        ff=ImageFont.truetype(FONT_BOLD,px(40));d2.text((px(238),px(250)),'X',font=ff,fill='white',anchor='mm')
    mark2=mark2.resize((600,600),Image.Resampling.LANCZOS)
    horizontal=Image.new('RGBA',(px(1680),px(520)),(0,0,0,0));horizontal.alpha_composite(mark2.resize((px(450),px(450)),Image.Resampling.LANCZOS),(px(20),px(35)))
    dd=ImageDraw.Draw(horizontal);ft=ImageFont.truetype(FONT_BOLD,px(82));fs=ImageFont.truetype(FONT_BOLD,px(36));fp=ImageFont.truetype(FONT_REG,px(29))
    dd.text((px(505),px(201)),f'TIỆN ÍCH {label}',font=ft,fill=INK,anchor='lm')
    dd.rounded_rectangle((px(510),px(265),px(699),px(273)),radius=px(4),fill=color)
    dd.text((px(510),px(326)),'VIỆT NAM',font=fs,fill=color,anchor='lm')
    dd.text((px(510),px(385)),f'MICROSOFT {label}',font=fp,fill=GRAY,anchor='lm')
    stem=f'TienIch{label.title()}-Logo'
    horizontal.resize((1680,520),Image.Resampling.LANCZOS).save(OUT/f'{stem}-Ngang.png',optimize=True)
    white=Image.new('RGBA',horizontal.size,'white');white.alpha_composite(horizontal)
    white=white.resize((1680,520),Image.Resampling.LANCZOS).convert('RGB')
    white.save(OUT/f'{stem}-Ngang-Trang.png',optimize=True)
    mark2.resize((256,256),Image.Resampling.LANCZOS).save(OUT/f'TienIch{label.title()}-Icon.png',optimize=True)
    white.resize((840,260),Image.Resampling.LANCZOS).save(OUT/f'{stem}-Ngang.bmp')

    header=Image.new('RGBA',(px(1680),px(220)),(255,255,255,0))
    header.alpha_composite(mark2.resize((px(160),px(160)),Image.Resampling.LANCZOS),(px(20),px(30)))
    hd=ImageDraw.Draw(header);fh=ImageFont.truetype(FONT_BOLD,px(70));fsub=ImageFont.truetype(FONT_BOLD,px(27))
    hd.text((px(205),px(88)),f'TIỆN ÍCH {label}',font=fh,fill=INK,anchor='lm')
    hd.rounded_rectangle((px(208),px(137),px(375),px(144)),radius=px(3),fill=color)
    hd.text((px(208),px(181)),'VIỆT NAM',font=fsub,fill=color,anchor='lm')
    header_png=header.resize((1680,220),Image.Resampling.LANCZOS)
    header_png.save(OUT/f'TienIch{label.title()}-AboutHeader.png',optimize=True)
    header_white=Image.new('RGBA',header.size,'white');header_white.alpha_composite(header)
    header_white.resize((840,110),Image.Resampling.LANCZOS).convert('RGB').save(OUT/f'TienIch{label.title()}-AboutHeader.bmp')

    svg_mark=f'''<svg xmlns="http://www.w3.org/2000/svg" width="1680" height="520" viewBox="0 0 1680 520"><rect x="82" y="50" width="436" height="460" rx="54" fill="{color}"/><rect x="140" y="108" width="320" height="346" rx="28" fill="white"/><path d="M365 108h1l94 95h-95z" fill="{pale}" stroke="{color}" stroke-opacity=".35" stroke-width="3"/>'''
    if kind=='word':
        svg_mark+=f'''<text x="300" y="340" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif" font-size="205" font-weight="700" fill="{color}">W</text><rect x="205" y="375" width="190" height="15" rx="7" fill="#DCE9FA"/>'''
    else:
        svg_mark+=f'''<rect x="202" y="216" width="196" height="186" rx="16" fill="#E9F6ED"/><path d="M267 224v170m66-170v170M210 278h180m-180 62h180" stroke="#87C69E" stroke-width="5"/><rect x="211" y="224" width="55" height="52" rx="8" fill="{dark}"/><text x="238" y="262" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif" font-size="40" font-weight="700" fill="white">X</text>'''
    svg_mark+=f'''<text x="505" y="214" font-family="Segoe UI,Arial,sans-serif" font-size="82" font-weight="700" fill="{INK}">TIỆN ÍCH {label}</text><rect x="510" y="265" width="189" height="8" rx="4" fill="{color}"/><text x="510" y="338" font-family="Segoe UI,Arial,sans-serif" font-size="36" font-weight="700" letter-spacing="2" fill="{color}">VIỆT NAM</text><text x="510" y="410" font-family="Segoe UI,Arial,sans-serif" font-size="29" fill="{GRAY}">MICROSOFT {label}</text></svg>'''
    (OUT/f'{stem}.svg').write_text(svg_mark,encoding='utf-8')

product_logo('word')
product_logo('excel')
print('Created logo files in',OUT)
