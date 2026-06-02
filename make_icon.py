# -*- coding: utf-8 -*-
"""
公文自动排版工具 - 纯代码图标生成脚本
使用 Pillow 矢量参数渲染多分辨率帧并打包为 icon.ico
避免使用 cairosvg 和外部 SVG 依赖，保证 100% 环境无关兼容性
"""
import os
import sys
from PIL import Image, ImageDraw

def draw_icon_at_size(size):
    scale = size / 80.0
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # 1. 背景圆角矩形
    bg_fill = (21, 101, 192, 255)  # #1565C0
    rx = 18 * scale
    draw.rounded_rectangle([0, 0, size, size], radius=rx, fill=bg_fill)
    
    # 2. 白色文档轮廓（带透明度）
    doc_left = 12 * scale
    doc_top = 10 * scale
    doc_right = 54 * scale  # 12 + 42
    doc_bottom = 62 * scale  # 10 + 52
    doc_rx = 5 * scale
    draw.rounded_rectangle([doc_left, doc_top, doc_right, doc_bottom], radius=doc_rx, 
                           fill=(255, 255, 255, 46))  # opacity 0.18
    draw.rounded_rectangle([doc_left, doc_top, doc_right, doc_bottom], radius=doc_rx, 
                           outline=(255, 255, 255, 204), width=max(1, int(2 * scale)))  # opacity 0.8

    # 3. 文档页眉
    draw.rounded_rectangle([doc_left, doc_top, doc_right, doc_top + 10 * scale], radius=doc_rx,
                           fill=(255, 255, 255, 64))  # opacity 0.25

    # 4. 正文线条
    lines_info = [
        (20, 26, 26, 3.5, 242),  # opacity 0.95
        (20, 34, 26, 3, 178),    # opacity 0.7
        (20, 42, 18, 3, 178),    # opacity 0.7
        (20, 50, 22, 3, 128)     # opacity 0.5
    ]
    for lx, ly, lw, lh, l_alpha in lines_info:
        draw.rounded_rectangle([lx * scale, ly * scale, (lx + lw) * scale, (ly + lh) * scale],
                               radius=1.5 * scale, fill=(255, 255, 255, l_alpha))

    # 5. 右下角圆形徽章
    cx = 57 * scale
    cy = 57 * scale
    r = 15 * scale
    draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(13, 71, 161, 255),  # #0D47A1
                 outline=(100, 181, 246, 255), width=max(1, int(2 * scale)))  # #64B5F6

    # 6. 徽章内的“A”形图案
    # 对角线端点：51,61 -> 56,50 -> 61,61
    p1 = (51 * scale, 61 * scale)
    p2 = (56 * scale, 50 * scale)
    p3 = (61 * scale, 61 * scale)
    a_width = max(1, int(2.2 * scale))
    draw.line([p1, p2, p3], fill=(255, 255, 255, 255), width=a_width, joint='round')
    
    # 横档：53,57 -> 59,57
    draw.line([(53 * scale, 57 * scale), (59 * scale, 57 * scale)], 
              fill=(255, 255, 255, 255), width=max(1, int(2 * scale)))
              
    # 底部装饰短立线：57,63 -> 57,65
    draw.line([(57 * scale, 63 * scale), (57 * scale, 65 * scale)], 
              fill=(255, 255, 255, 255), width=max(1, int(2 * scale)))

    return img

def main():
    print("正在用 Pillow 矢量参数绘制多尺寸图标...")
    
    # 绘制高清晰度原版 (512x512)
    base_img = draw_icon_at_size(512)
    
    # 保存为标准的包含多分辨率帧的 icon.ico
    sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    ico_path = "icon.ico"
    # PIL 会自动在保存为 ICO 时使用高清晰度下采样生成对应的帧
    base_img.save(ico_path, format="ICO", sizes=sizes)
    print(f"\n[OK] 图标生成成功，已写入：{ico_path} （包含全部 {len(sizes)} 个常用尺寸帧）")

if __name__ == "__main__":
    main()
