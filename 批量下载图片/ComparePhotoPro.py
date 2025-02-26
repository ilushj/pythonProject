import cv2
import numpy as np
from PIL import Image
import matplotlib.pyplot as plt
import matplotlib
import pytesseract

# 强制使用无 GUI 后端
matplotlib.use('Agg')


# 加载图片
def load_image(image_path):
    img = cv2.imread(image_path)
    if img is None:
        raise FileNotFoundError(f"Image not found: {image_path}")
    return img


# 预处理：转换为灰度图并平滑
def preprocess_image(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    return gray, blurred


# 边缘检测
def edge_detection(blurred):
    edges = cv2.Canny(blurred, 50, 150)
    return edges


# 定位包含冒号的文字区域
def locate_colon_text(img):
    """
    使用OCR定位包含冒号的文字区域，返回边界框和文字内容。

    参数：
        img (ndarray): 输入图像

    返回：
        boxes (list): [(x, y, w, h, text, colon_idx), ...]
    """
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    data = pytesseract.image_to_data(gray, lang='chi_sim', output_type=pytesseract.Output.DICT)
    boxes = []
    for i in range(len(data['text'])):
        text = data['text'][i].strip()
        if ':' in text or '：' in text:  # 支持英文和中文冒号
            x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
            colon_idx = text.index(':' if ':' in text else '：')
            boxes.append((x, y, w, h, text, colon_idx))
    return boxes


# 噪声分析（按冒号分割区域）
def noise_analysis_colon_split(gray, box, colon_idx, char_width_approx):
    """
    按冒号分割区域，分别分析前后的噪声水平。

    参数：
        gray (ndarray): 灰度图像
        box (tuple): (x, y, w, h) 文字区域边界框
        colon_idx (int): 冒号在文字中的索引
        char_width_approx (int): 每个字符的近似宽度

    返回：
        before_noise (float): 冒号前噪声水平
        after_noise (float): 冒号后噪声水平
    """
    x, y, w, h = box
    colon_x = x + colon_idx * char_width_approx  # 估算冒号的x坐标
    colon_end_x = colon_x + char_width_approx

    # 前部分区域
    before_region = gray[y:y + h, x:colon_x]
    if before_region.size == 0:
        before_noise = 0
    else:
        before_laplacian = cv2.Laplacian(before_region, cv2.CV_64F)
        before_noise = np.var(before_laplacian)

    # 后部分区域
    after_region = gray[y:y + h, colon_end_x:x + w]
    if after_region.size == 0:
        after_noise = 0
    else:
        after_laplacian = cv2.Laplacian(after_region, cv2.CV_64F)
        after_noise = np.var(after_laplacian)

    return before_noise, after_noise


# ELA分析
def ela_analysis(image_path, quality=90):
    img = Image.open(image_path)
    temp_path = "temp.jpg"
    img.save(temp_path, "JPEG", quality=quality)
    temp_img = cv2.imread(temp_path)
    orig_img = cv2.imread(image_path)
    if temp_img.shape != orig_img.shape:
        raise ValueError("Original and temporary images have different shapes.")
    diff = cv2.absdiff(orig_img, temp_img)
    exaggerated = cv2.convertScaleAbs(diff, alpha=10)
    return exaggerated


# 检查篡改痕迹（比较冒号前后）
def detect_tampering_colon(image_path, noise_diff_threshold=500):
    img = load_image(image_path)
    gray, _ = preprocess_image(img)

    # 定位包含冒号的文字
    text_boxes = locate_colon_text(img)
    if not text_boxes:
        return "未找到包含冒号的文字区域", []

    # 估算每个字符的宽度
    results = []
    for box in text_boxes:
        x, y, w, h, text, colon_idx = box
        char_width_approx = w // len(text)  # 近似每个字符宽度

        # 计算前后噪声
        before_noise, after_noise = noise_analysis_colon_split(gray, (x, y, w, h), colon_idx, char_width_approx)
        noise_diff = after_noise - before_noise
        is_tampered = noise_diff > noise_diff_threshold

        results.append((x, y, w, h, text, before_noise, after_noise, is_tampered))
        print(
            f"文字: '{text}', 前噪声: {before_noise:.2f}, 后噪声: {after_noise:.2f}, 差异: {noise_diff:.2f}, 是否篡改: {is_tampered}")

    tampered_count = sum(1 for _, _, _, _, _, _, _, is_tampered in results if is_tampered)
    return f"检测到 {len(text_boxes)} 个含冒号文字区域，其中 {tampered_count} 个可能被篡改", results


# 可视化结果
def show_results(img, edges, ela, text_results, output_path="output.png"):
    """
    可视化结果，标注冒号前后区域及其篡改状态。

    参数：
        img (ndarray): 原始图像
        edges (ndarray): 边缘检测结果
        ela (ndarray): ELA分析结果
        text_results (list): [(x, y, w, h, text, before_noise, after_noise, is_tampered), ...]
        output_path (str): 输出图像保存路径
    """
    plt.figure(figsize=(12, 4))
    img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

    for x, y, w, h, text, before_noise, after_noise, is_tampered in text_results:
        char_width = w // len(text)
        colon_x = x + text.index(':' if ':' in text else '：') * char_width
        colon_end_x = colon_x + char_width

        # 前部分（绿色）
        cv2.rectangle(img_rgb, (x, y), (colon_x, y + h), (0, 255, 0), 1)
        # 后部分（根据篡改状态）
        color = (255, 0, 0) if is_tampered else (0, 255, 0)
        cv2.rectangle(img_rgb, (colon_end_x, y), (x + w, y + h), color, 2)

        # 标注文字和状态
        label = f"{text[:10]} {'(T)' if is_tampered else ''}"
        cv2.putText(img_rgb, label, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.6, color, 2)

    plt.subplot(131), plt.imshow(img_rgb), plt.title('Original + Text Detection')
    plt.subplot(132), plt.imshow(edges, cmap='gray'), plt.title('Edges')
    plt.subplot(133), plt.imshow(ela, cmap='gray'), plt.title('ELA')

    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"结果已保存至: {output_path}")


# 主程序
if __name__ == "__main__":
    image_path = "test_image.jpg"  # 替换为你的图片路径

    # 检测篡改
    result, text_results = detect_tampering_colon(image_path, noise_diff_threshold=500)
    print(f"检测结果: {result}")

    # 可视化
    img = load_image(image_path)
    _, blurred = preprocess_image(img)
    edges = edge_detection(blurred)
    ela = ela_analysis(image_path)
    show_results(img, edges, ela, text_results, output_path="tamper_detection_results.png")