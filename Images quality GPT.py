import os
import shutil

# 文件夹路径（原始图片所在）
folder_path = r"D:\Codex\命名测试"  # 替换成你的文件夹路径

# 清晰度和倍率列表
resolution_list = ["new", "old"]
multiplier_list = ["0.5x", "1x", "2x", "3x"]

# 遍历文件夹
for filename in os.listdir(folder_path):
    file_path = os.path.join(folder_path, filename)
    if not os.path.isfile(file_path):
        continue

    # 找到文件名中的倍率
    multiplier_found = None
    for m in multiplier_list:
        if m in filename:
            multiplier_found = m
            break
    if not multiplier_found:
        print(f"{filename} 未匹配到倍率，保留在原文件夹")
        continue

    # 找到文件名中的清晰度
    resolution_found = None
    for r in resolution_list:
        if r in filename:
            resolution_found = r
            break
    if not resolution_found:
        print(f"{filename} 未匹配到清晰度，保留在原文件夹")
        continue

    # 创建目标文件夹：倍率/清晰度
    target_folder = os.path.join(folder_path, multiplier_found, resolution_found)
    os.makedirs(target_folder, exist_ok=True)

    # 移动文件
    new_path = os.path.join(target_folder, filename)
    shutil.move(file_path, new_path)
    print(f"{filename} → {multiplier_found}/{resolution_found}/{filename}")