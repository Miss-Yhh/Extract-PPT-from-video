import cv2
import os
import numpy as np

def extract_ppt_from_image(image, output_path):
    # 将图像转换为灰度图
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # 应用高斯模糊以去除噪声
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)

    # 使用Canny边缘检测
    edged = cv2.Canny(blurred, 50, 150)

    # 查找轮廓
    contours, _ = cv2.findContours(edged, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # 初始化最大矩形区域
    max_rect = None
    max_area = 0

    for contour in contours:
        # 计算轮廓的边界矩形
        (x, y, w, h) = cv2.boundingRect(contour)
        area = w * h

        # 找到面积最大的矩形
        if area > max_area:
            max_area = area
            max_rect = (x, y, w, h)

    if max_rect is not None:
        x, y, w, h = max_rect
        ppt_region = image[y:y+h, x:x+w]
        cv2.imwrite(output_path, ppt_region)
        print(f"PPT区域已保存到: {output_path}")
    else:
        print("未找到PPT区域")

def process_images_in_folder(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for filename in os.listdir(input_folder):
        if filename.endswith(".jpg") or filename.endswith(".png"):
            image_path = os.path.join(input_folder, filename)
            image = cv2.imread(image_path)

            if image is not None:
                output_path = os.path.join(output_folder, filename)
                extract_ppt_from_image(image, output_path)
            else:
                print(f"无法读取图片: {filename}")

# 使用示例
input_folder = "frames"  # 替换成你的图片文件夹路径
output_folder = "largest-images"  # 替换成你希望保存图片的文件夹路径

process_images_in_folder(input_folder, output_folder)

def find_largest_image(input_folder, output_path):
    max_area = 0
    largest_image = None
    largest_image_path = None

    for filename in os.listdir(input_folder):
        if filename.endswith(".jpg") or filename.endswith(".png"):
            image_path = os.path.join(input_folder, filename)
            image = cv2.imread(image_path)

            if image is not None:
                height, width, _ = image.shape
                area = width * height

                if area > max_area:
                    max_area = area
                    largest_image = image
                    largest_image_path = image_path

    if largest_image is not None:
        cv2.imwrite(output_path, largest_image)
        print(f"最大尺寸的图片已保存到: {output_path}")
        print(f"最大尺寸的图片路径是: {largest_image_path}")
    else:
        print("未找到任何图片")

# 使用示例
input_folder = "largest-images" # 替换成你的图片文件夹路径
output_path = "output_largest_image.jpg"  # 替换成你希望保存最大图片的路径

find_largest_image(input_folder, output_path)
def find_template_in_image(image, template):
    result = cv2.matchTemplate(image, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
    top_left = max_loc
    height, width, _ = template.shape
    bottom_right = (top_left[0] + width, top_left[1] + height)
    return top_left, bottom_right

def crop_images_to_template(input_folder, template_path, output_folder):
    # 读取模板图像
    template = cv2.imread(template_path)
    if template is None:
        print("无法读取模板图片")
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for filename in os.listdir(input_folder):
        if filename.endswith(".jpg") or filename.endswith(".png"):
            image_path = os.path.join(input_folder, filename)
            image = cv2.imread(image_path)

            if image is not None:
                top_left, bottom_right = find_template_in_image(image, template)
                cropped_image = image[top_left[1]:bottom_right[1], top_left[0]:bottom_right[0]]
                output_path = os.path.join(output_folder, filename)
                cv2.imwrite(output_path, cropped_image)
                print(f"{filename} 已裁剪并保存到 {output_path}")
            else:
                print(f"无法读取图片: {filename}")

# 使用示例
input_folder = "frames"  # 替换成你的图片文件夹路径
template_path = "output_largest_image.jpg"  # 替换成你的模板图片路径
output_folder = "Cropped-frames"  # 替换成你希望保存裁剪图片的文件夹路径

crop_images_to_template(input_folder, template_path, output_folder)
