"""Generate crisp, size-adaptive Word installer icons."""
from pathlib import Path
import struct
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
TARGET = ROOT.parent / 'output/Word-2.30.4/TienIchWord.ico'
TARGET.parent.mkdir(parents=True, exist_ok=True)
FONT = Path('C:/Windows/Fonts/seguisb.ttf')
SIZES = [16, 24, 32, 48, 64, 128, 256]

def frame(size):
    scale = 4
    n = size * scale
    im = Image.new('RGBA', (n, n), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    blue = '#185ABD'
    d.rounded_rectangle((scale, scale, n-scale-1, n-scale-1), radius=round(size*.20*scale), fill=blue)
    if size <= 24:
        # At toolbar/taskbar sizes, prioritize a clear W over document details.
        font = ImageFont.truetype(str(FONT), round(size*.78*scale))
        d.text((n/2, n/2-size*.015*scale), 'W', font=font, fill='white', anchor='mm', stroke_width=max(0,round(size*.018*scale)), stroke_fill='white')
    else:
        # Simplified page silhouette with a large, high-contrast Word mark.
        x0,y0,x1,y1 = round(size*.22*scale),round(size*.13*scale),round(size*.78*scale),round(size*.87*scale)
        d.rounded_rectangle((x0,y0,x1,y1), radius=round(size*.035*scale), fill='white')
        fold=round(size*.16*scale)
        d.polygon([(x1-fold,y0),(x1,y0+fold),(x1-fold,y0+fold)], fill='#DCEBFF')
        d.line([(x1-fold,y0+1),(x1-fold,y0+fold),(x1-1,y0+fold)], fill='#78A7E8', width=max(scale,round(size*.012*scale)))
        font = ImageFont.truetype(str(FONT), round(size*.49*scale))
        d.text((n/2, n*.54), 'W', font=font, fill=blue, anchor='mm', stroke_width=max(0,round(size*.012*scale)), stroke_fill=blue)
    return im.resize((size,size), Image.Resampling.LANCZOS)

images=[frame(n) for n in SIZES]
# Use classic Windows DIB frames: the .NET compiler and older shell icon readers
# handle these consistently, unlike PNG-compressed frames in an ICO.
def dib(im):
    w,h=im.size; rgba=im.convert('RGBA'); pixels=rgba.load()
    xor=bytearray()
    for y in range(h-1,-1,-1):
        for x in range(w):
            r,g,b,a=pixels[x,y]; xor.extend((b,g,r,a))
    stride=((w+31)//32)*4; mask=bytearray(stride*h)
    for y in range(h):
        row=(h-1-y)*stride
        for x in range(w):
            if pixels[x,y][3] == 0: mask[row+x//8] |= 1 << (7-x%8)
    info=struct.pack('<IiiHHIIiiII',40,w,h*2,1,32,0,len(xor),0,0,0,0)
    return info+xor+mask
frames=[dib(im) for im in images]
header=struct.pack('<HHH',0,1,len(SIZES)); offset=6+16*len(SIZES); entries=[]
for size,data in zip(SIZES,frames):
    entries.append(struct.pack('<BBBBHHII',0 if size==256 else size,0 if size==256 else size,0,0,1,32,len(data),offset));offset+=len(data)
TARGET.write_bytes(header+b''.join(entries)+b''.join(frames))
images[-1].save(TARGET.with_suffix('.png'))
print(TARGET)
