"""
生成XLSX转Canvas工具的ICO图标
使用PIL库创建多尺寸图标
"""

from PIL import Image, ImageDraw, ImageFont
import os


def create_icon():
    """
    创建ICO图标
    设计概念：Excel表格(绿色) + 双向箭头 + Canvas卡片(紫色/蓝色)
    """
    # 创建多尺寸图标
    sizes = [256, 128, 64, 48, 32, 16]
    images = []
    
    for size in sizes:
        # 创建图像
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        # 计算比例
        scale = size / 256
        margin = int(10 * scale)
        
        # 绘制背景圆角矩形
        bg_color = (59, 130, 246)  # 蓝色 #3b82f6
        draw.rounded_rectangle(
            [0, 0, size-1, size-1],
            radius=int(40 * scale),
            fill=bg_color
        )
        
        # 绘制Excel表格部分（左侧）
        excel_x = int(40 * scale)
        excel_y = int(70 * scale)
        excel_w = int(70 * scale)
        excel_h = int(90 * scale)
        
        # Excel绿色背景
        excel_color = (34, 197, 94)  # 绿色 #22c55e
        draw.rounded_rectangle(
            [excel_x, excel_y, excel_x + excel_w, excel_y + excel_h],
            radius=int(8 * scale),
            fill=excel_color
        )
        
        # 表格线条
        line_color = (255, 255, 255, 200)
        # 横线
        for i in range(1, 4):
            y = excel_y + int(excel_h * i / 4)
            draw.line([(excel_x + 5, y), (excel_x + excel_w - 5, y)],
                     fill=line_color, width=max(1, int(2 * scale)))
        # 竖线
        draw.line([(excel_x + int(excel_w/2), excel_y + 5),
                  (excel_x + int(excel_w/2), excel_y + excel_h - 5)],
                 fill=line_color, width=max(1, int(2 * scale)))
        
        # 绘制Canvas卡片部分（右侧）
        canvas_x = int(145 * scale)
        canvas_y = int(70 * scale)
        canvas_w = int(70 * scale)
        canvas_h = int(90 * scale)
        
        # Canvas卡片白色背景
        canvas_color = (255, 255, 255)
        draw.rounded_rectangle(
            [canvas_x, canvas_y, canvas_x + canvas_w, canvas_y + canvas_h],
            radius=int(8 * scale),
            fill=canvas_color
        )
        
        # 卡片内的文字线条
        text_color = (107, 114, 128)  # 灰色
        line_height = int(12 * scale)
        for i in range(3):
            y = canvas_y + int(20 * scale) + i * line_height * 2
            line_w = int(40 * scale) if i < 2 else int(25 * scale)
            draw.rounded_rectangle(
                [canvas_x + int(10 * scale), y,
                 canvas_x + int(10 * scale) + line_w, y + line_height],
                radius=int(2 * scale),
                fill=text_color
            )
        
        # 绘制双向箭头
        arrow_y = int(115 * scale)
        arrow_color = (255, 255, 255)
        
        # 左箭头（从Excel指向Canvas）
        arrow_x1 = int(115 * scale)
        arrow_x2 = int(140 * scale)
        # 箭头线
        draw.line([(arrow_x1, arrow_y), (arrow_x2, arrow_y)],
                 fill=arrow_color, width=max(2, int(4 * scale)))
        # 箭头头部
        head_size = int(8 * scale)
        draw.polygon([(arrow_x2, arrow_y),
                     (arrow_x2 - head_size, arrow_y - head_size//2),
                     (arrow_x2 - head_size, arrow_y + head_size//2)],
                    fill=arrow_color)
        
        # 右箭头（从Canvas指向Excel）
        arrow_y2 = int(135 * scale)
        draw.line([(arrow_x2, arrow_y2), (arrow_x1, arrow_y2)],
                 fill=arrow_color, width=max(2, int(4 * scale)))
        draw.polygon([(arrow_x1, arrow_y2),
                     (arrow_x1 + head_size, arrow_y2 - head_size//2),
                     (arrow_x1 + head_size, arrow_y2 + head_size//2)],
                    fill=arrow_color)
        
        images.append(img)
    
    # 保存ICO文件
    output_path = 'd:\\Users\\Jliu Pureey\\Downloads\\gitmine\\cpp23\\xlsx_to_canvas.ico'
    images[0].save(
        output_path,
        format='ICO',
        sizes=[(s, s) for s in sizes],
        append_images=images[1:]
    )
    
    print(f"图标已生成：{output_path}")
    return output_path


if __name__ == '__main__':
    create_icon()
