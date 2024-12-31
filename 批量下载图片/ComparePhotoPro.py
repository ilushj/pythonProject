import cv2
import numpy as np
import os
import xlsxwriter
import logging
import shutil

# 设置日志记录
logging.basicConfig(filename='image_quality.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 黑边检测
def detect_dark_edges(img, threshold=50, edge_width_ratio=0.05):
    """
    检测图片是否有深色边缘（类似黑边但不完全纯黑）。

    参数:
        img: 输入的图像 (BGR 格式)
        threshold: 边缘区域亮度阈值，小于该值判断为深色边缘
        edge_width_ratio: 边缘检测的宽度比例 (相对于图像宽度或高度)

    返回:
        是否有深色边缘 (True/False)
    """
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    _, _, v = cv2.split(hsv)  # 提取亮度 (V 通道)
    height, width = v.shape

    edge_width = int(min(height, width) * edge_width_ratio)

    # 提取上下左右边缘
    top_edge = v[:edge_width, :]
    bottom_edge = v[-edge_width:, :]
    left_edge = v[:, :edge_width]
    right_edge = v[:, -edge_width:]

    # 计算边缘区域的平均亮度
    top_mean = np.mean(top_edge)
    bottom_mean = np.mean(bottom_edge)
    left_mean = np.mean(left_edge)
    right_mean = np.mean(right_edge)

    # 判断是否存在深色边缘
    if (
        top_mean < threshold or
        bottom_mean < threshold or
        left_mean < threshold or
        right_mean < threshold
    ):
        return True
    return False

# 对比度检测
def calculate_contrast(img):
    """
    计算图像对比度，使用亮度（灰度）直方图计算对比度值。

    参数:
        img: 输入的图像 (BGR 格式)

    返回:
        对比度值 (float)
    """
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)  # 转换为灰度图
    return np.std(gray)  # 计算灰度值的标准差作为对比度

# 纹理失真检测
def calculate_texture_distortion(img):
    """
    计算纹理失真，通过 Sobel 算子检测纹理变化。

    参数:
        img: 输入的图像 (BGR 格式)

    返回:
        纹理失真值 (float)
    """
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    sobel_x = cv2.Sobel(gray, cv2.CV_64F, 1, 0, ksize=3)
    sobel_y = cv2.Sobel(gray, cv2.CV_64F, 0, 1, ksize=3)
    texture = np.sqrt(sobel_x ** 2 + sobel_y ** 2)
    return np.mean(texture)

# 主处理逻辑
def process_images(dir_path):
    """
    处理输入目录中的图片，检测质量问题并生成报告。

    参数:
        dir_path: 图片目录路径
    """
    # 创建 Excel 文件
    workbook = xlsxwriter.Workbook('image_quality.xlsx')
    worksheet = workbook.add_worksheet()

    # 设置标题行
    worksheet.write(0, 0, '文件名')
    worksheet.write(0, 1, '黑边检测')
    worksheet.write(0, 2, '偏色度')
    worksheet.write(0, 3, '清晰度')
    worksheet.write(0, 4, '纹理失真')
    worksheet.write(0, 5, '对比度')

    # 初始化行号
    row = 1

    # 创建“疑似”目录
    suspicious_dir = os.path.join(dir_path, "疑似")
    if not os.path.exists(suspicious_dir):
        os.makedirs(suspicious_dir)

    # 遍历目录中的 JPG 文件
    for filename in os.listdir(dir_path):
        if filename.endswith('.jpg'):
            img_path = os.path.join(dir_path, filename)
            logging.info(f"正在处理文件：{img_path}")

            try:
                # 读取图片
                img = cv2.imdecode(np.fromfile(img_path, dtype=np.uint8), cv2.IMREAD_COLOR)

                # 检查图像是否为空
                if img is None:
                    logging.error(f"无法读取图像文件：{img_path}")
                    continue

                # 黑边检测
                black_edge = detect_dark_edges(img, threshold=50, edge_width_ratio=0.05)

                # 偏色检测
                hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
                h, s, _ = cv2.split(hsv)
                color_deviation = int(np.std(h) + np.std(s))

                # 清晰度检测
                clarity = int(cv2.Laplacian(img, cv2.CV_64F).var())

                # 纹理失真检测
                texture_distortion = calculate_texture_distortion(img)

                # 对比度检测
                contrast = calculate_contrast(img)

                # 写入 Excel 文件
                worksheet.write(row, 0, filename)
                worksheet.write(row, 1, '是' if black_edge else '否')
                worksheet.write(row, 2, color_deviation)
                worksheet.write(row, 3, clarity)
                worksheet.write(row, 4, texture_distortion)
                worksheet.write(row, 5, contrast)

                # 复制文件到“疑似”目录
                if black_edge or (clarity < 100 and texture_distortion < 20) :
                    shutil.copy(img_path, suspicious_dir)

                row += 1
            except Exception as e:
                logging.error(f"处理文件 {filename} 时发生错误：{e}")

    # 关闭 Excel 文件
    workbook.close()

if __name__ == "__main__":
    dir_path = input("请输入图片目录路径：")
    if os.path.exists(dir_path):
        process_images(dir_path)
        print("处理完成，结果已保存到 image_quality.xlsx")
    else:
        print("输入的目录路径不存在，请重新输入！")
