import cv2
import numpy as np
import os
import xlsxwriter
import logging
import shutil

# 设置日志记录
logging.basicConfig(filename='image_quality.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# 输入目录路径
dir_path = input("请输入图片目录路径：")

# 创建 Excel 文件
workbook = xlsxwriter.Workbook('image_quality.xlsx')
worksheet = workbook.add_worksheet()

# 设置标题行
worksheet.write(0, 0, '文件名')
worksheet.write(0, 1, '黑边检测')
worksheet.write(0, 2, '偏色度')
worksheet.write(0, 3, '清晰度 (拉普拉斯方差)')
worksheet.write(0, 4, '熵值')
worksheet.write(0, 5, '边缘像素比例')

# 初始化行号
row = 1

# 创建“疑似”目录
suspicious_dir = os.path.join(dir_path, "疑似")
if not os.path.exists(suspicious_dir):
    os.makedirs(suspicious_dir)


# 黑边检测函数
def check_black_edge(img, edge_width=6, brightness_threshold=60):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    top_edge = gray[:edge_width, :]
    bottom_edge = gray[-edge_width:, :]
    if np.all(top_edge < brightness_threshold) or np.all(bottom_edge < brightness_threshold):
        return True
    return False


# 偏色度检测函数
def check_color_deviation(img):
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    h, s, v = cv2.split(hsv)
    return int(np.std(h) + np.std(s) + np.std(v))


# 拉普拉斯方差（清晰度）检测函数
def calculate_laplacian_variance(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    laplacian = cv2.Laplacian(gray, cv2.CV_64F)
    return laplacian.var()


# 熵值检测函数
def calculate_entropy(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    hist = cv2.calcHist([gray], [0], None, [256], [0, 256])
    hist = hist / hist.sum()  # 归一化
    entropy = -np.sum(hist * np.log2(hist + np.finfo(float).eps))  # 防止log(0)导致错误
    return entropy


# 边缘像素比例检测函数
def calculate_edge_pixel_ratio(img, low_threshold=100, high_threshold=200):
    # 将图像转为灰度图
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 使用 Canny 边缘检测
    edges = cv2.Canny(gray, low_threshold, high_threshold)

    # 计算边缘像素数量
    edge_pixels = np.sum(edges > 0)

    # 计算图像总像素数量
    total_pixels = img.shape[0] * img.shape[1]

    # 计算边缘像素比例
    edge_pixel_ratio = edge_pixels / total_pixels
    return edge_pixel_ratio


# 遍历目录中的图像文件（支持 .jpg 和 .png）
for filename in os.listdir(dir_path):
    if filename.lower().endswith(('.jpg', '.jpeg', '.png')):
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
            black_edge = check_black_edge(img)

            # 偏色度检测
            color_deviation = check_color_deviation(img)

            # 拉普拉斯方差（清晰度）检测
            clarity = calculate_laplacian_variance(img)

            # 熵值检测
            entropy = calculate_entropy(img)

            # 边缘像素比例检测
            edge_pixel_ratio = calculate_edge_pixel_ratio(img)

            # 写入 Excel 文件
            worksheet.write(row, 0, filename)
            worksheet.write(row, 1, '是' if black_edge else '否')
            worksheet.write(row, 2, color_deviation)
            worksheet.write(row, 3, clarity)
            worksheet.write(row, 4, entropy)
            worksheet.write(row, 5, edge_pixel_ratio)

            # 如果检测到黑边、低清晰度、低熵值或低边缘像素比例，将文件复制到“疑似”目录
            if black_edge or (clarity < 100 and edge_pixel_ratio < 0.01):  # 边缘像素比例阈值可以根据需要调整
                shutil.copy(img_path, suspicious_dir)

            row += 1
        except Exception as e:
            logging.error(f"处理文件 {filename} 时发生错误：{e}")

# 关闭 Excel 文件
workbook.close()

print("处理完成，结果已保存至 'image_quality.xlsx'，疑似图片已复制到 '疑似' 目录。")
