#!/usr/bin/env python3
"""生成 Outlook Add-in 所需的各种尺寸图标"""
from pathlib import Path

try:
    from PIL import Image, ImageDraw
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'Pillow', '-q'])
    from PIL import Image, ImageDraw

ASSETS = Path(__file__).parent / 'src' / 'assets'
ASSETS.mkdir(parents=True, exist_ok=True)

# 颜色配置
BG_COLOR = '#1a1a1a'
ACCENT_COLOR = '#f4a261'

sizes = [16, 32, 64, 80, 128]

for size in sizes:
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # 圆角矩形背景
    pad = size // 16
    draw.rounded_rectangle(
        [pad, pad, size - pad, size - pad],
        radius=size // 4,
        fill=BG_COLOR
    )
    
    # 中心圆点（模拟 Kimi 月亮/星星的感觉）
    center = size // 2
    r = size // 5
    draw.ellipse(
        [center - r, center - r, center + r, center + r],
        fill=ACCENT_COLOR
    )
    
    # 小星星点缀
    if size >= 32:
        sr = max(1, size // 20)
        draw.ellipse([center + r, center - r - sr*2, center + r + sr*2, center - r], fill='white')
        draw.ellipse([center - r - sr*3, center + r, center - r - sr, center + r + sr*2], fill='white')
    
    img.save(ASSETS / f'icon-{size}.png')
    print(f'Generated icon-{size}.png ({size}x{size})')

print('Done!')
