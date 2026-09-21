from PIL import Image, ImageDraw

for size in (192, 512):
    scale = size / 512
    im = Image.new('RGB', (size, size), '#142a31')
    d = ImageDraw.Draw(im)
    def box(coords): return tuple(int(v * scale) for v in coords)
    d.rounded_rectangle(box((104, 134, 408, 390)), radius=int(36*scale), fill='#f5f6f2')
    for y, right in ((196, 362), (247, 362), (298, 283)):
        d.line(box((150, y, right, y)), fill='#142a31', width=max(2,int(18*scale)))
    d.ellipse(box((287,273,405,391)), fill='#d3a261')
    d.line([box((319,332))[:2],box((338,351))[:2],box((376,310))[:2]],fill='#142a31',width=max(2,int(17*scale)),joint='curve')
    im.save(f'icon-{size}.png')
