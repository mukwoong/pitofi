from openpyxl import load_workbook
from PIL import Image
import os

def download_all_images_from_sheet(filename, save_dir):
    wb = load_workbook(filename)
    # 获取第一个工作表
    sheet = wb.worksheets[0]

    for i, image1 in enumerate(sheet._images, start=1):
        image_anchor = image1.anchor
        image_from = image_anchor._from

        col, colOff, row, rowOff = image_from.col, image_from.colOff, image_from.row, image_from.rowOff

        # 获取图像的数据
        img = Image.open(image1.ref).convert("RGB")

        # 将图像数据保存到文件
        save_path = f"{save_dir}/image{row}_{col}.png"  # 使用行列信息来保存图片，确保顺序
        img.save(save_path)

        print(f"已下载图像: {save_path}")

    wb.close()

# 指定要读取的 xlsx 文件路径和保存目录
xlsx_file = 'akira-XS2024388.xlsx'
save_dir = 'img'  # 修改为当前工作目录下的 'img' 文件夹

# 调用函数读取工作表中的所有图像并下载
download_all_images_from_sheet(xlsx_file, save_dir)
