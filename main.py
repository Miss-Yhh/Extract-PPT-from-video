import cv2
import os
from PIL import Image
import imagehash

# 创建一个文件夹来保存提取的帧
output_folder = "frames"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 打开视频文件
video_path = "test.mp4"  # 替换成你的视频路径
cap = cv2.VideoCapture(video_path)

# 获取视频帧率
fps = cap.get(cv2.CAP_PROP_FPS)
interval = int(fps * 1.5)  # 每隔1.5秒提取一帧

frame_count = 0
saved_count = 0

while cap.isOpened():
    ret, frame = cap.read()
    # frame = frame[121:678, 142:1136]
    if not ret:
        break

    if frame_count % interval == 0:
        frame_filename = os.path.join(output_folder, f"frame_{saved_count}.jpg")
        cv2.imwrite(frame_filename, frame)
        saved_count += 1

    frame_count += 1

cap.release()

# 删除相似图片并重命名剩余图片
def find_similar_images_and_rename(folder, threshold=5):
    image_hashes = []
    files_to_keep = []

    # 计算每个图片的哈希值
    for filename in os.listdir(folder):
        if filename.endswith(".jpg"):
            filepath = os.path.join(folder, filename)
            image = Image.open(filepath)
            image_hash = imagehash.average_hash(image)

            is_similar = False
            for existing_hash in image_hashes:
                if image_hash - existing_hash <= threshold:
                    is_similar = True
                    break

            if not is_similar:
                image_hashes.append(image_hash)
                files_to_keep.append(filepath)

    # 删除不保留的图片
    for filename in os.listdir(folder):
        if filename.endswith(".jpg"):
            filepath = os.path.join(folder, filename)
            if filepath not in files_to_keep:
                os.remove(filepath)

    # 重新命名图片文件
    remaining_files = sorted([f for f in os.listdir(folder) if f.endswith(".jpg")], key=lambda x: int(x.split('_')[1].split('.')[0]))
    for i, filename in enumerate(remaining_files, start=1):
        file_extension = os.path.splitext(filename)[1]
        new_filename = f"{i}{file_extension}"
        old_filepath = os.path.join(folder, filename)
        new_filepath = os.path.join(folder, new_filename)
        os.rename(old_filepath, new_filepath)

find_similar_images_and_rename(output_folder, threshold=5)

print("完成！")
