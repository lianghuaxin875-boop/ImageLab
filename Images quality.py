#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
图片智能整理工具 - 针对 iPhone 14 Pro Max 及通用设备优化
功能：
- 根据分辨率判断清晰度（标清/高清/超清）
- 根据 EXIF 焦段信息判断倍率（1x/2x/3x/倍率未知）
- 自动重命名：格式 [清晰度]_[倍率]_原文件名
- 支持复制或移动模式
"""

import os
import sys
import shutil
import json
import csv
import re
from pathlib import Path
from typing import Dict, Optional, List, Tuple

from PIL import Image, ImageOps
from PIL.ExifTags import TAGS

# ===================== 用户配置区域 =====================
# 输入文件夹（存放原始图片）
INPUT_DIR = r"D:\Codex\命名测试"  # 请修改为你的路径

# 输出文件夹（处理后的图片存放位置）
OUTPUT_DIR = r"D:\Codex\命名测试\新建文件夹"  # 请修改为你的路径

# 模式选择：True=复制（保留原文件），False=移动/重命名原文件（危险）
COPY_MODE = True

# 是否仅预览（不实际修改任何文件）
DRY_RUN = False

# 是否打印详细信息（调试用）
VERBOSE = True

# 清晰度阈值 (宽, 高) - 可根据需要调整
SD_THRESHOLD = (1280, 720)  # < 1280x720 为标清
HD_THRESHOLD = (1920, 1080)  # 1280x720~1920x1080 为高清
# > 1920x1080 为超清

# ===================== 焦段识别配置 =====================
# 主摄基准物理焦距（单位：mm），用于计算光学变焦倍数。
# iPhone 14 Pro Max 主摄实际焦距 = 6.86mm
MAIN_CAMERA_FOCAL_MM = 6.86  # 如不确定可设为 None，脚本会尝试自动推断

# 镜头关键词映射表：关键词（支持部分匹配）-> 倍率
# 优先级最高，请根据实际打印的 LensModel 添加关键词
LENS_KEYWORD_TO_ZOOM = {
    # ---------- iPhone 14 Pro Max 专用 ----------
    "back camera 6.86mm f/1.78": 1,  # 1x 主摄
    "back dual camera 9mm f/2.8": 2,  # 2x 裁切模式
    "back triple camera 13mm f/2.4": 3,  # 3x 长焦
    # 简化关键词（慎用，可能误判）
    "6.86mm": 1,
    "9mm": 2,
    "13mm": 3,

    # ---------- 其他品牌示例（可按需添加）----------
    "main": 1, "wide": 1, "26mm": 1, "28mm": 1,
    "2x": 2, "52mm": 2, "telephoto 2x": 2,
    "3x": 3, "77mm": 3, "85mm": 3, "periscope": 3,
}

# 等效35mm焦距区间映射（当镜头关键词匹配失败时使用）
FOCAL_35MM_RANGES = [
    (0, 35, 1),  # 0-35mm -> 1x
    (35, 65, 2),  # 35-65mm -> 2x
    (65, 200, 3),  # 65-200mm -> 3x
    (200, 999, 4),  # 200mm以上 -> 4x（可根据需要调整）
]

# 实际物理焦距直接映射（后备方案，优先级低于比例计算）
ACTUAL_FOCAL_MAP = {
    6.86: 1,  # iPhone 主摄
    9.0: 2,  # iPhone 2x 模式
    13.0: 3,  # iPhone 3x 长焦
}

# 手工规则持久化文件（自动生成，用于保存用户手动标注的规则）
UNKNOWN_RULES_FILE = "unknown_focal_rules.json"

# 支持的图片扩展名
IMAGE_EXTENSIONS = [".jpg", ".jpeg", ".png", ".tiff", ".bmp", ".webp", ".heic", ".HEIC"]


# ===================== 配置结束 =====================


# ---------- 底层工具函数 ----------
def to_float(value) -> Optional[float]:
    """将有理数、字符串等转换为浮点数"""
    if value is None:
        return None
    if isinstance(value, tuple) and len(value) == 2:
        return value[0] / value[1] if value[1] != 0 else None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            num = re.search(r'[\d\.]+', value)
            if num:
                return float(num.group())
        except:
            pass
    return None


def get_exif_dict(img_path: Path) -> Dict:
    """读取图片 EXIF 信息，返回标签名 -> 值的字典"""
    exif = {}
    try:
        with Image.open(img_path) as img:
            # 先纠正方向，确保尺寸正确
            img = ImageOps.exif_transpose(img)
            raw = img._getexif()
            if raw:
                for tag_id, value in raw.items():
                    tag_name = TAGS.get(tag_id, tag_id)
                    exif[tag_name] = value
    except Exception as e:
        if VERBOSE:
            print(f"  EXIF读取警告: {e}")
    return exif


def get_image_size(img_path: Path) -> Tuple[int, int]:
    """返回 (width, height)"""
    with Image.open(img_path) as img:
        img = ImageOps.exif_transpose(img)
        return img.width, img.height


def get_clarity_label(width: int, height: int) -> str:
    """根据分辨率返回清晰度标签"""
    if width < SD_THRESHOLD[0] or height < SD_THRESHOLD[1]:
        return "标清"
    elif width <= HD_THRESHOLD[0] and height <= HD_THRESHOLD[1]:
        return "高清"
    else:
        return "超清"


def extract_focal_info(exif: Dict) -> Dict:
    """
    从 EXIF 中提取关键焦段信息
    返回字典包含: lens_model, focal_35mm, focal_actual, make, model
    """
    # 镜头型号
    lens_model = exif.get("LensModel", "")
    if not lens_model:
        lens_model = exif.get("LensMake", "")
    if not lens_model:
        lens_model = exif.get("LensSpecification", "")

    # 等效35mm焦距（多种可能的标签名）
    focal_35mm = None
    for tag in ["FocalLengthIn35mmFilm", "FocalLength35mmEquiv", "FocalLengthIn35mmFormat"]:
        if tag in exif:
            focal_35mm = to_float(exif[tag])
            if focal_35mm is not None:
                break

    # 实际物理焦距
    focal_actual = None
    if "FocalLength" in exif:
        focal_actual = to_float(exif["FocalLength"])

    # 相机信息
    make = exif.get("Make", "")
    model = exif.get("Model", "")

    return {
        "lens_model": str(lens_model),
        "focal_35mm": focal_35mm,
        "focal_actual": focal_actual,
        "make": str(make),
        "model": str(model),
    }


def get_zoom_by_lens_model(lens_str: str) -> Optional[int]:
    """通过镜头型号关键词匹配倍率"""
    if not lens_str:
        return None
    lower = lens_str.lower()
    for keyword, zoom in LENS_KEYWORD_TO_ZOOM.items():
        if keyword.lower() in lower:
            return zoom
    # 尝试提取 "数字x" 模式
    match = re.search(r'(\d+)x', lower)
    if match:
        return int(match.group(1))
    return None


def get_zoom_by_35mm_focal(f35: float) -> Optional[int]:
    """通过等效35mm焦距区间匹配倍率"""
    for lo, hi, zoom in FOCAL_35MM_RANGES:
        if lo <= f35 < hi:
            return zoom
    return None


def get_zoom_by_actual_focal(f_actual: float, main_focal: Optional[float]) -> Optional[int]:
    """
    通过实际物理焦距计算倍率：倍率 = 实际焦距 / 主摄基准焦距
    如果提供了 main_focal 且有效，则计算；否则使用预定义的映射表 ACTUAL_FOCAL_MAP
    """
    if f_actual is None or f_actual <= 0:
        return None

    # 优先使用比例计算
    if main_focal is not None and main_focal > 0:
        ratio = f_actual / main_focal
        # 四舍五入取整，限制在 1~10 之间
        zoom = int(round(ratio))
        if zoom < 1:
            zoom = 1
        if zoom > 10:
            zoom = None
        if zoom is not None:
            return zoom

    # 后备：直接映射表（允许模糊匹配）
    for known_focal, zoom in ACTUAL_FOCAL_MAP.items():
        if abs(f_actual - known_focal) < 0.2:  # 允许 0.2mm 误差
            return zoom
    return None


def determine_zoom(focal_info: Dict, main_focal: Optional[float]) -> Optional[int]:
    """
    综合判断倍率，优先级：
    1. 镜头型号关键词匹配
    2. 等效35mm焦距区间
    3. 实际物理焦距比例计算/映射
    """
    # 1. 镜头型号
    zoom = get_zoom_by_lens_model(focal_info["lens_model"])
    if zoom is not None:
        return zoom

    # 2. 等效35mm焦距
    if focal_info["focal_35mm"] is not None:
        zoom = get_zoom_by_35mm_focal(focal_info["focal_35mm"])
        if zoom is not None:
            return zoom

    # 3. 实际物理焦距
    if focal_info["focal_actual"] is not None:
        zoom = get_zoom_by_actual_focal(focal_info["focal_actual"], main_focal)
        if zoom is not None:
            return zoom

    return None


def load_manual_rules(rules_file: str) -> Dict[str, int]:
    """加载手工标注的规则（文件路径 -> 倍率）"""
    if os.path.exists(rules_file):
        try:
            with open(rules_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}


def save_manual_rules(rules: Dict[str, int], rules_file: str):
    """保存手工规则"""
    with open(rules_file, 'w', encoding='utf-8') as f:
        json.dump(rules, f, indent=2, ensure_ascii=False)


def process_one_image(img_path: Path, output_dir: Path, main_focal: Optional[float],
                      manual_rules: Dict[str, int], dry_run: bool, copy_mode: bool) -> Dict:
    """处理单张图片，返回处理结果字典"""
    result = {
        "original": str(img_path),
        "new_name": None,
        "clarity": None,
        "zoom": None,
        "status": "skip",
        "message": "",
    }
    try:
        # 1. 尺寸与清晰度
        width, height = get_image_size(img_path)
        clarity = get_clarity_label(width, height)
        result["clarity"] = clarity

        # 2. 焦段识别
        exif = get_exif_dict(img_path)
        focal_info = extract_focal_info(exif)

        if VERBOSE:
            print(f"  相机: {focal_info['make']} {focal_info['model']}")
            print(f"  镜头: {focal_info['lens_model']}")
            print(f"  等效35mm焦距: {focal_info['focal_35mm']}")
            print(f"  实际焦距: {focal_info['focal_actual']} mm")

        # 优先使用手工规则（基于文件完整路径）
        file_key = str(img_path.absolute())
        zoom = manual_rules.get(file_key)
        if zoom is None:
            zoom = determine_zoom(focal_info, main_focal)

        if zoom is None:
            zoom_label = "倍率未知"
            zoom_str = "?"
        else:
            zoom_label = f"{zoom}x"
            zoom_str = str(zoom)
        result["zoom"] = zoom_str

        # 3. 生成新文件名
        prefix = f"{clarity}_{zoom_label}_"
        new_name = prefix + img_path.name
        result["new_name"] = new_name

        # 4. 执行操作
        if dry_run:
            result["status"] = "dry_run"
            result["message"] = f"[预览] 将重命名为: {new_name}"
        else:
            out_path = output_dir / new_name
            if copy_mode:
                shutil.copy2(img_path, out_path)
                result["status"] = "copied"
                result["message"] = f"已复制到 {out_path}"
            else:
                os.rename(img_path, out_path)
                result["status"] = "renamed"
                result["message"] = f"已重命名为 {out_path}"
    except Exception as e:
        result["status"] = "error"
        result["message"] = str(e)
    return result


def export_unknown_to_csv(unknown_list: List[Dict], csv_path: str):
    """将无法识别的图片信息导出为 CSV，供手动标注"""
    if not unknown_list:
        return
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['file_path', 'make', 'model', 'lens_model', 'focal_35mm', 'focal_actual',
                                               'suggested_zoom'])
        writer.writeheader()
        for item in unknown_list:
            writer.writerow(item)
    print(f"\n未识别倍率的图片信息已导出到: {csv_path}")
    print("请在 'suggested_zoom' 列填入倍率（数字1,2,3...），然后运行 --import-csv 导入。")


def import_csv_rules(csv_path: str, rules_file: str):
    """从手动标注的 CSV 中读取倍率，存入规则文件"""
    if not os.path.exists(csv_path):
        print(f"错误: CSV 文件不存在: {csv_path}")
        return
    rules = load_manual_rules(rules_file)
    with open(csv_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            zoom_str = row.get('suggested_zoom', '').strip()
            if zoom_str and zoom_str.isdigit():
                file_path = row['file_path']
                rules[file_path] = int(zoom_str)
    save_manual_rules(rules, rules_file)
    print(f"已导入 {len(rules)} 条规则，保存到 {rules_file}")


def main():
    import argparse
    parser = argparse.ArgumentParser(description="图片智能整理工具（iPhone/通用）")
    parser.add_argument("-i", "--input", default=INPUT_DIR, help="输入文件夹")
    parser.add_argument("-o", "--output", default=OUTPUT_DIR, help="输出文件夹")
    parser.add_argument("--copy", action="store_true", default=COPY_MODE, help="复制模式（默认）")
    parser.add_argument("--move", action="store_true", help="移动/重命名模式（危险）")
    parser.add_argument("--dry-run", action="store_true", default=DRY_RUN, help="仅预览")
    parser.add_argument("--verbose", action="store_true", default=VERBOSE, help="打印详细信息")
    parser.add_argument("--main-focal", type=float, default=MAIN_CAMERA_FOCAL_MM, help="主摄实际焦距(mm)")
    parser.add_argument("--export-csv", action="store_true", help="导出未识别图片信息到CSV")
    parser.add_argument("--import-csv", type=str, help="从CSV导入人工标注的倍率")
    args = parser.parse_args()

    # 处理导入CSV命令
    if args.import_csv:
        import_csv_rules(args.import_csv, UNKNOWN_RULES_FILE)
        return

    # 检查输入文件夹
    input_path = Path(args.input)
    if not input_path.is_dir():
        print(f"错误: 输入文件夹不存在: {input_path}")
        print("请修改脚本顶部的 INPUT_DIR 或使用 -i 参数指定正确路径")
        sys.exit(1)

    output_path = Path(args.output)
    copy_mode = not args.move if args.move else args.copy

    # 创建输出目录
    if not args.dry_run and not output_path.exists():
        output_path.mkdir(parents=True)
        print(f"创建输出目录: {output_path}")

    if args.move and not args.dry_run:
        confirm = input("警告: 移动模式会直接重命名原文件，不可恢复！继续？(y/N): ")
        if confirm.lower() != 'y':
            print("取消操作")
            sys.exit(0)

    # 收集图片
    images = []
    for ext in IMAGE_EXTENSIONS:
        images.extend(input_path.glob(f"*{ext}"))
        # 也支持大写扩展名
        images.extend(input_path.glob(f"*{ext.upper()}"))
    images = list(set(images))
    if not images:
        print(f"未找到任何图片，支持的扩展名: {IMAGE_EXTENSIONS}")
        return

    print(f"找到 {len(images)} 张图片")
    if args.dry_run:
        print("【预览模式】不会实际修改文件")

    # 加载已有手工规则
    manual_rules = load_manual_rules(UNKNOWN_RULES_FILE)

    results = []
    unknown_list = []

    for idx, img in enumerate(sorted(images), 1):
        print(f"\n[{idx}/{len(images)}] 处理: {img.name}")
        res = process_one_image(img, output_path, args.main_focal, manual_rules,
                                args.dry_run, copy_mode)
        results.append(res)
        print(f"  => {res['message']}")

        # 收集未识别的图片信息
        if res["zoom"] == "?":
            # 重新提取一次完整 EXIF 信息用于 CSV
            exif = get_exif_dict(img)
            info = extract_focal_info(exif)
            unknown_list.append({
                "file_path": str(img.absolute()),
                "make": info["make"],
                "model": info["model"],
                "lens_model": info["lens_model"],
                "focal_35mm": info["focal_35mm"],
                "focal_actual": info["focal_actual"],
                "suggested_zoom": "",
            })

    # 导出 CSV（如果需要或者有未识别图片）
    if args.export_csv or unknown_list:
        csv_path = output_path / "unknown_focal_list.csv" if not args.dry_run else "unknown_focal_list.csv"
        export_unknown_to_csv(unknown_list, str(csv_path))

    # 统计结果
    print("\n========== 处理完成 ==========")
    stats = {}
    for r in results:
        stats[r["status"]] = stats.get(r["status"], 0) + 1
    zoom_stats = {}
    for r in results:
        z = r["zoom"]
        zoom_stats[z] = zoom_stats.get(z, 0) + 1

    print(f"总计: {len(results)} 张")
    print(f"  复制: {stats.get('copied', 0)}")
    print(f"  重命名: {stats.get('renamed', 0)}")
    print(f"  预览: {stats.get('dry_run', 0)}")
    print(f"  错误: {stats.get('error', 0)}")
    print("倍率分布:")
    for z, cnt in sorted(zoom_stats.items()):
        print(f"  {z}: {cnt} 张")

    # 保存处理日志
    if not args.dry_run and (stats.get('copied', 0) + stats.get('renamed', 0)):
        log_file = output_path / "processing_log.txt"
        with open(log_file, 'w', encoding='utf-8') as f:
            for r in results:
                f.write(f"{r['original']} -> {r['new_name']} [{r['clarity']} {r['zoom']}] {r['status']}\n")
        print(f"日志已保存至: {log_file}")


if __name__ == "__main__":
    main()