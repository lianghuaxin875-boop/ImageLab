import os
from PIL import Image, ExifTags

# ----------------------------
# 配置文件夹路径（原始照片所在）
folder_path = r"D:\Codex\命名测试"  # 替换为你的文件夹路径

# ----------------------------
# 自定义倍率映射表（根据镜头型号）
lens_multiplier = {
    "Luna 18-55mm": "1x",
    "Luna 50-135mm": "2x",
    "GoPro Hero10": "1x",
    "GoPro Hero11": "2x",
    # 可以继续添加其他镜头
}

# ----------------------------
# 分辨率判断清晰度
def get_resolution_label(width, height):
    total_pixels = width * height
    if total_pixels < 1280*720:
        return "标清"
    elif total_pixels <= 1920*1080:
        return "高清"
    else:
        return "超清"

# ----------------------------
# 根据 EXIF 获取设备类型和镜头型号
def get_device_and_lens(exif_data):
    make = exif_data.get(271, "")   # Make
    model = exif_data.get(272, "")  # Model
    lens = exif_data.get(42036, "") # LensModel
    make = str(make).lower()
    model = str(model).lower()
    lens = str(lens)

    if "luna" in model or "luna" in make:
        device_type = "Luna"
    elif "gopro" in model or "osmo" in model:
        device_type = "运动相机"
    elif any(x in make for x in ["iphone","huawei","xiaomi","samsung"]):
        device_type = "手机"
    else:
        device_type = "普通相机"

    return device_type, lens

# ----------------------------
# 遍历文件夹处理图片
for filename in os.listdir(folder_path):
    file_path = os.path.join(folder_path, filename)
    if not os.path.isfile(file_path):
        continue

    try:
        img = Image.open(file_path)
        width, height = img.size
        exif_data = img._getexif() or {}
        img.close()  # 关闭文件，避免锁定

        resolution_label = get_resolution_label(width, height)
        device_type, lens_model = get_device_and_lens(exif_data)

        # 查表获取倍率
        multiplier_label = lens_multiplier.get(lens_model, "未知倍率")

        # 生成前缀 + 原文件名
        prefix = f"{resolution_label}_{multiplier_label}_{device_type}_"
        new_name = prefix + filename
        new_path = os.path.join(folder_path, new_name)

        os.rename(file_path, new_path)
        print(f"{filename} → {new_name}")

    except Exception as e:
        print(f"无法处理 {filename}: {e}")