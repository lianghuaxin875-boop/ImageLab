# -*- coding: utf-8 -*-
"""
图片质量检测工具 GUI 版

功能：
- 选择输入图片目录、OK/NG 输出目录、报告目录
- 支持整图检测、九宫格/中心/四角区域检测
- 支持预设权重和自定义权重
- 后台线程执行检测，生成多 Sheet Excel 报告
"""

import os
import queue
import threading
from datetime import datetime
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:  # Web/容器环境可能没有 Tk；检测核心仍可独立运行。
    class _TkUnavailable:
        def __init__(self, *args, **kwargs):
            raise RuntimeError("当前运行环境没有 Tkinter，网页模式不需要桌面 GUI 依赖")

    class _TkModule:
        Tk = _TkUnavailable

    class _TkWidgetsUnavailable:
        def __getattr__(self, name):
            raise RuntimeError("当前运行环境没有 Tkinter，网页模式不需要桌面 GUI 依赖")

    tk = _TkModule()
    ttk = _TkWidgetsUnavailable()
    filedialog = messagebox = _TkWidgetsUnavailable()

import cv2
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage, ImageDraw, ImageFont


IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".bmp")
DETECTOR_VERSION = "2026-05-12-safe-image-output-v2"
WEIGHT_MODES = {
    "默认综合模式": (0.40, 0.20, 0.20, 0.10, 0.10),
    "Camera对焦测试模式": (0.60, 0.10, 0.10, 0.10, 0.10),
    "低光噪声测试模式": (0.25, 0.20, 0.15, 0.30, 0.10),
    "文档扫描检测模式": (0.45, 0.20, 0.25, 0.05, 0.05),
}


def score_laplacian(lap):
    if lap >= 300:
        return 100
    if lap >= 150:
        return 75 + (lap - 150) / 150 * 25
    if lap >= 80:
        return 50 + (lap - 80) / 70 * 25
    if lap >= 0:
        return max(0, lap / 80 * 50)
    return 0


def score_tenengrad(ten):
    if ten >= 2000:
        return 100
    if ten >= 800:
        return 75 + (ten - 800) / 1200 * 25
    if ten >= 300:
        return 50 + (ten - 300) / 500 * 25
    if ten >= 0:
        return max(0, ten / 300 * 50)
    return 0


def score_brightness(bright):
    dist = abs(bright - 128)
    if dist <= 18:
        return 100 - dist / 18 * 15
    if dist <= 38:
        return 85 - (dist - 18) / 20 * 15
    if dist <= 58:
        return 70 - (dist - 38) / 20 * 20
    if dist <= 128:
        return max(0, 50 - (dist - 58) / 70 * 50)
    return 0


def score_overexposure(ratio):
    if ratio <= 0.01:
        return 100
    if ratio <= 0.05:
        return 85 + (0.05 - ratio) / 0.04 * 15
    if ratio <= 0.15:
        return 60 + (0.15 - ratio) / 0.10 * 25
    if ratio <= 1.0:
        return max(0, 60 * (1 - ratio) / 0.85)
    return 0


def score_std(std):
    if 50 <= std <= 80:
        return 100
    if 35 <= std < 50:
        return 75 + (std - 35) / 15 * 25
    if 80 < std <= 100:
        return 100 - (std - 80) / 20 * 25
    if std < 35:
        return max(0, std / 35 * 75)
    if std > 100:
        return max(0, 75 - (std - 100) / 20 * 25)
    return 0


def score_dynamic_range(dr):
    if dr >= 180:
        return 100
    if dr >= 100:
        return 75 + (dr - 100) / 80 * 25
    if dr >= 60:
        return 50 + (dr - 60) / 40 * 25
    if dr >= 0:
        return max(0, dr / 60 * 50)
    return 0


def score_noise(noise):
    if noise <= 5:
        return 95 + (5 - noise) / 5 * 5
    if noise <= 10:
        return 75 + (10 - noise) / 5 * 20
    if noise <= 20:
        return 50 + (20 - noise) / 10 * 25
    if noise <= 35:
        return 30 + (35 - noise) / 15 * 20
    if noise <= 50:
        return max(0, 30 - (noise - 35) / 15 * 30)
    return 0


def level_from_score(score):
    if score >= 85:
        return "EXCELLENT"
    if score >= 70:
        return "OK"
    if score >= 50:
        return "WARNING"
    return "NG"


def clarity_desc_sug(lap_score, ten_score, total_score):
    descs, sugs = [], []
    if lap_score < 50:
        descs.append("图像边缘变化弱，疑似模糊、失焦或拍摄抖动")
        sugs.append("建议检查对焦、镜头清洁度或重新拍摄")
    elif lap_score < 70:
        descs.append("边缘强度偏低，清晰度一般")
        sugs.append("建议检查对焦或避免轻微抖动")
    if ten_score < 50:
        descs.append("纹理细节明显不足，图像不够锐利")
        sugs.append("建议提高快门速度或使用三脚架")
    elif ten_score < 70:
        descs.append("细节表现力略有不足")
        sugs.append("可适当提高锐度或检查镜头解析力")
    if not descs:
        descs.append("清晰度优秀，边缘锐利" if total_score >= 85 else "清晰度良好")
    return "；".join(descs), "；".join(sugs)


def contrast_desc_sug(std_score, range_score, total_score):
    descs, sugs = [], []
    if std_score < 50:
        descs.append("明暗变化不足，画面可能发灰")
        sugs.append("建议调整光照层次或进行对比度增强")
    elif std_score < 70:
        descs.append("对比度略低，层次感一般")
        sugs.append("可适当增强对比度")
    if range_score < 50:
        descs.append("动态范围不足，亮暗细节丢失")
        sugs.append("建议改善布光，避免过强光比")
    elif range_score < 70:
        descs.append("动态范围偏低，细节表现欠佳")
        sugs.append("可尝试HDR或调整曝光")
    if not descs:
        descs.append("对比度优秀，层次分明" if total_score >= 85 else "对比度正常")
    return "；".join(descs), "；".join(sugs)


def exposure_desc_sug(mean_bright, bright_score, over_score, under_score, total_score):
    descs, sugs = [], []
    if bright_score < 50:
        if mean_bright < 90:
            descs.append("图像平均亮度偏低，疑似欠曝")
            sugs.append("建议增加光照或提高曝光补偿")
        elif mean_bright > 170:
            descs.append("图像平均亮度偏高，疑似过曝")
            sugs.append("建议降低曝光或避免强光反射")
    elif bright_score < 70:
        descs.append("亮度偏暗，建议适当提亮" if mean_bright < 110 else "亮度偏亮，建议适当降低曝光")
    if over_score < 50:
        descs.append("高亮区域溢出，亮部细节可能丢失")
        sugs.append("控制高光光源或降低曝光")
    elif over_score < 70:
        descs.append("高光部分偏多，亮部细节略有损失")
        sugs.append("可适当减少高光区域")
    if under_score < 50:
        descs.append("暗部区域过多，暗部细节可能丢失")
        sugs.append("增加暗部照明或提高曝光补偿")
    elif under_score < 70:
        descs.append("暗部区域偏多，部分细节丢失")
        sugs.append("可适当提亮暗部")
    if not descs:
        descs.append("曝光精准，亮度适中" if total_score >= 85 else "曝光正常")
    return "；".join(descs), "；".join(sugs)


def noise_desc_sug(noise_score):
    if noise_score < 50:
        return "高频残差较大，疑似噪声明显", "建议降低ISO、增加环境光或使用降噪处理"
    if noise_score < 70:
        return "存在轻微噪声，可能影响细节", "可尝试降噪或使用更低ISO"
    if noise_score >= 85:
        return "画面纯净，噪声极低", ""
    return "噪声水平正常", ""


def artifact_desc_sug(artifact_score):
    if artifact_score < 50:
        return "8x8块边界差异明显，疑似JPEG块效应", "建议提高JPEG保存质量，避免重复压缩"
    if artifact_score < 70:
        return "可见轻微块效应，压缩痕迹略明显", "建议使用更高质量保存或使用PNG格式"
    if artifact_score >= 85:
        return "无可见压缩伪影", ""
    return "压缩伪影轻微", ""


def calc_clarity_metrics(gray):
    lap = cv2.Laplacian(gray, cv2.CV_64F).var()
    gx = cv2.Sobel(gray, cv2.CV_64F, 1, 0, ksize=3)
    gy = cv2.Sobel(gray, cv2.CV_64F, 0, 1, ksize=3)
    ten = np.sqrt(gx**2 + gy**2).mean()
    return lap, ten


def calc_noise(gray, ksize=5):
    blur = cv2.GaussianBlur(gray, (ksize, ksize), 0)
    return np.std(gray.astype(np.float32) - blur.astype(np.float32))


def calc_artifact(gray, width, height):
    v = [np.mean(np.abs(gray[:, x].astype(np.float32) - gray[:, x + 1].astype(np.float32))) for x in range(7, width - 1, 8)]
    h = [np.mean(np.abs(gray[y, :].astype(np.float32) - gray[y + 1, :].astype(np.float32))) for y in range(7, height - 1, 8)]
    if v and h:
        return np.mean(v + h)
    return np.mean(v) if v else (np.mean(h) if h else 0.0)


def _load_chinese_font(size=22):
    font_candidates = [
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/simsun.ttc",
        "C:/Windows/Fonts/arial.ttf",
    ]
    for font_path in font_candidates:
        if os.path.exists(font_path):
            return ImageFont.truetype(font_path, size=size)
    return ImageFont.load_default()


def _fmt_score(value):
    return f"{float(value):.1f}".rstrip("0").rstrip(".")


def primary_issue(metric_items):
    weak = [(name, score, desc) for name, score, desc in metric_items if score < 70 and desc]
    if weak:
        weak.sort(key=lambda item: item[1])
        return weak[0][2].split("；")[0]

    mild_candidates = [(name, score) for name, score, _desc in metric_items if score < 85]
    if not mild_candidates:
        return "无明显问题"

    name, _score = min(mild_candidates, key=lambda item: item[1])
    mild_map = {
        "清晰度": "清晰度略有不足",
        "曝光": "曝光略有偏差",
        "对比度": "对比度略有不足",
        "噪声": "轻微噪声",
        "压缩": "轻微压缩痕迹",
    }
    return mild_map.get(name, "存在轻微质量波动")


def calc_color_cast(image):
    b_mean, g_mean, r_mean = [float(v) for v in cv2.mean(image)[:3]]
    max_channel = max(("R", r_mean), ("G", g_mean), ("B", b_mean), key=lambda item: item[1])
    min_channel = min(("R", r_mean), ("G", g_mean), ("B", b_mean), key=lambda item: item[1])
    diff = max_channel[1] - min_channel[1]

    result, desc = "正常", "RGB通道均值接近，未见明显色偏"
    if diff <= 15:
        pass
    elif r_mean > b_mean + 15 and g_mean > b_mean + 15:
        result = "偏黄"
        desc = "B 通道明显偏低，画面偏黄 / 偏暖"
    elif g_mean > r_mean + 15 and b_mean > r_mean + 15:
        result = "偏青"
        desc = "R 通道明显偏低，画面偏青"
    elif r_mean > g_mean + 15 and b_mean > g_mean + 15:
        result = "偏紫/洋红"
        desc = "G 通道明显偏低，画面偏紫 / 偏洋红"
    elif diff > 15:
        if max_channel[0] == "R":
            result = "偏红"
            desc = "R 通道偏高，画面偏红 / 偏暖"
        elif max_channel[0] == "G":
            result = "偏绿"
            desc = "G 通道偏高，画面偏绿"
        else:
            result = "偏蓝"
            desc = "B 通道偏高，画面偏蓝 / 偏冷"

    return {
        "色偏结果": result,
        "色偏说明": desc,
        "R通道均值": round(r_mean, 2),
        "G通道均值": round(g_mean, 2),
        "B通道均值": round(b_mean, 2),
        "通道最大差": round(diff, 2),
    }


def score_white_balance_diff(max_diff):
    if max_diff <= 10:
        return round(100 - max_diff / 10 * 5, 1)
    if max_diff <= 15:
        return round(94 - (max_diff - 11) / 4 * 9, 1)
    if max_diff <= 25:
        return round(84 - (max_diff - 16) / 9 * 14, 1)
    if max_diff <= 35:
        return round(69 - (max_diff - 26) / 9 * 9, 1)
    return round(max(0, 59 - (max_diff - 35) / 65 * 59), 1)


def white_balance_level_from_score(score):
    if score >= 95:
        return "优秀"
    if score >= 85:
        return "OK"
    if score >= 70:
        return "轻微偏色"
    if score >= 60:
        return "偏色明显"
    return "NG"


def calc_white_balance(color_metrics):
    cast_result = color_metrics["色偏结果"]
    max_diff = color_metrics["通道最大差"]
    wb_score = score_white_balance_diff(max_diff)
    level = white_balance_level_from_score(wb_score)

    direction_desc = {
        "正常": "RGB通道差异在可接受范围内",
        "偏红": "R 通道偏高，画面偏红 / 偏暖",
        "偏蓝": "B 通道偏高，画面偏蓝 / 偏冷",
        "偏绿": "G 通道偏高，画面偏绿",
        "偏黄": "B 通道明显偏低，画面偏黄 / 偏暖",
        "偏青": "R 通道明显偏低，画面偏青",
        "偏紫/洋红": "G 通道明显偏低，画面偏紫 / 偏洋红",
    }.get(cast_result, "RGB通道均衡性不足，白平衡可能存在偏差")
    return wb_score, level, f"{level}；{direction_desc}；RGB通道最大差 {max_diff:.2f}"


def calc_luminance_histogram(gray):
    total = gray.size
    dark_ratio = float(np.sum(gray < 85) / total)
    mid_ratio = float(np.sum((gray >= 85) & (gray < 170)) / total)
    highlight_ratio = float(np.sum(gray >= 170) / total)
    dead_black_ratio = float(np.sum(gray <= 5) / total)
    dead_white_ratio = float(np.sum(gray >= 250) / total)

    if dead_black_ratio >= 0.02:
        judgement = "存在死黑"
    elif dead_white_ratio >= 0.02:
        judgement = "存在死白"
    elif dark_ratio >= 0.40 and highlight_ratio < 0.12:
        judgement = "画面偏暗"
    elif highlight_ratio >= 0.40 and dark_ratio < 0.12:
        judgement = "画面偏亮"
    else:
        judgement = "亮度分布正常"

    return {
        "暗部占比": round(dark_ratio, 4),
        "中间调占比": round(mid_ratio, 4),
        "高光占比": round(highlight_ratio, 4),
        "死黑占比": round(dead_black_ratio, 4),
        "死白占比": round(dead_white_ratio, 4),
        "直方图判断": judgement,
    }


def annotate_image(image, score, status, metrics):
    annotated_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    pil_img = PILImage.fromarray(annotated_rgb).convert("RGBA")
    overlay = PILImage.new("RGBA", pil_img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    font = _load_chinese_font(22)
    title_font = _load_chinese_font(24)
    padding = 14
    line_height = 30
    lines = [
        ("综合评分", _fmt_score(score)),
        ("检测结果", status),
        ("清晰度", _fmt_score(metrics["清晰度"]["score"])),
        ("曝光", _fmt_score(metrics["曝光"]["score"])),
        ("对比度", _fmt_score(metrics["对比度"]["score"])),
        ("噪声", _fmt_score(metrics["噪声"]["score"])),
        ("压缩", _fmt_score(metrics["压缩"]["score"])),
        ("问题", metrics["问题"]),
    ]
    text_lines = [f"{key}：{value}" for key, value in lines]
    text_width = max(draw.textbbox((0, 0), text, font=title_font if idx < 2 else font)[2] for idx, text in enumerate(text_lines))
    panel_width = min(max(text_width + padding * 2, 260), pil_img.width - 20)
    panel_height = padding * 2 + line_height * len(text_lines)
    panel_x, panel_y = 10, 10

    draw.rounded_rectangle(
        (panel_x, panel_y, panel_x + panel_width, panel_y + panel_height),
        radius=12,
        fill=(8, 18, 34, 215),
        outline=(56, 189, 248, 235),
        width=2,
    )

    level_color = {
        "EXCELLENT": (134, 239, 172, 255),
        "OK": (134, 239, 172, 255),
        "WARNING": (253, 230, 138, 255),
        "NG": (253, 164, 175, 255),
        "ERROR": (253, 164, 175, 255),
    }.get(status, (226, 232, 240, 255))

    y = panel_y + padding
    for idx, text in enumerate(text_lines):
        active_font = title_font if idx < 2 else font
        fill = level_color if idx == 1 else (240, 249, 255, 255)
        draw.text((panel_x + padding, y), text, font=active_font, fill=fill)
        y += line_height

    combined = PILImage.alpha_composite(pil_img, overlay).convert("RGB")
    return cv2.cvtColor(np.array(combined), cv2.COLOR_RGB2BGR)


def get_region_metrics(gray, region_names):
    h, w = gray.shape
    hb, wb = h // 3, w // 3
    pos_map = {"上": 0, "中": 1, "下": 2}
    side_map = {"左": 0, "中": 1, "右": 2}
    metrics = {}
    for name in region_names:
        i, j = pos_map[name[0]], side_map[name[1]]
        y0, y1 = i * hb, (i + 1) * hb if i < 2 else h
        x0, x1 = j * wb, (j + 1) * wb if j < 2 else w
        block = gray[y0:y1, x0:x1]
        lap, ten = calc_clarity_metrics(block)
        metrics[name] = (lap, ten, np.std(block))
    return metrics


def region_names_for(label):
    all_regions = ["上左", "上中", "上右", "中左", "中中", "中右", "下左", "下中", "下右"]
    if label == "九宫格全部":
        return all_regions, "九宫格"
    if label == "中心区域":
        return ["中中"], "中心"
    if label == "四角区域":
        return ["上左", "上右", "下左", "下右"], "四角"
    return [], "无"


class ImageQualityDetector:
    def __init__(self, config, log_callback=None, progress_callback=None):
        self.config = config
        self.log = log_callback or (lambda msg: None)
        self.progress = progress_callback or (lambda done, total: None)

    def run(self):
        input_dir = Path(self.config["input_dir"])
        ok_dir = Path(self.config["ok_dir"])
        ng_dir = Path(self.config["ng_dir"])
        report_dir = Path(self.config["report_dir"])
        for directory in (input_dir, ok_dir, ng_dir, report_dir):
            directory.mkdir(parents=True, exist_ok=True)
        self.log(f"检测器版本：{DETECTOR_VERSION}")
        self.log(f"输入目录：{input_dir}")
        self.log(f"OK输出目录：{ok_dir}")
        self.log(f"NG输出目录：{ng_dir}")

        files = [p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_EXTENSIONS]
        results = []
        region_names, sub_label = region_names_for(self.config["region_mode"])

        if not files:
            raise RuntimeError("输入目录中没有找到 jpg/jpeg/png/bmp 图片。")

        for index, image_path in enumerate(files, 1):
            self.log(f"正在检测：{image_path.name}")
            results.append(self._process_one(image_path, ok_dir, ng_dir, region_names))
            self.progress(index, len(files))

        if not self.config.get("generate_report", True):
            self.log("检测完成，报告暂未生成")
            if self.config.get("return_results"):
                return {"results": results, "report_path": ""}
            return results

        report_path = self._write_report(results, report_dir, sub_label)
        self.log(f"检测完成，报告已生成：{report_path}")
        if self.config.get("return_results"):
            return {"results": results, "report_path": str(report_path)}
        return report_path

    @staticmethod
    def _safe_output_path(source_path, target_dir):
        source_path = Path(source_path).resolve()
        target_dir = Path(target_dir).resolve()
        output_path = target_dir / source_path.name
        if output_path.resolve() == source_path:
            output_path = target_dir / f"{source_path.stem}_result{source_path.suffix}"
        return output_path

    @staticmethod
    def _save_annotated_image(source_path, target_dir, annotated_image):
        output_path = ImageQualityDetector._safe_output_path(source_path, target_dir)
        ok, encoded = cv2.imencode(output_path.suffix, annotated_image)
        if not ok:
            raise RuntimeError(f"图片编码失败：{source_path.name}")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_bytes(encoded.tobytes())
        return output_path

    def _process_one(self, image_path, ok_dir, ng_dir, region_names):
        check_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        image = cv2.imdecode(np.fromfile(str(image_path), dtype=np.uint8), cv2.IMREAD_COLOR)
        if image is None:
            return {
                "文件名": image_path.name,
                "检测时间": check_time,
                "综合评分": 0,
                "综合等级": "ERROR",
                "综合说明": "图片无法读取",
            }

        height, width = image.shape[:2]
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        weights = self.config["weights"]

        lap_val, ten_val = calc_clarity_metrics(gray)
        lap_score = round(score_laplacian(lap_val), 1)
        ten_score = round(score_tenengrad(ten_val), 1)
        clarity_score = round(lap_score * 0.5 + ten_score * 0.5, 1)
        clarity_level = level_from_score(clarity_score)
        clarity_desc, clarity_sug = clarity_desc_sug(lap_score, ten_score, clarity_score)

        gray_min, gray_max = int(gray.min()), int(gray.max())
        gray_range = gray_max - gray_min
        std_val = np.std(gray)
        std_score = round(score_std(std_val), 1)
        range_score = round(score_dynamic_range(gray_range), 1)
        contrast_score = round(std_score * 0.7 + range_score * 0.3, 1)
        contrast_level = level_from_score(contrast_score)
        contrast_desc, contrast_sug = contrast_desc_sug(std_score, range_score, contrast_score)

        mean_bright = gray.mean()
        over_ratio = np.sum(gray > 240) / gray.size
        under_ratio = np.sum(gray < 30) / gray.size
        bright_score = round(score_brightness(mean_bright), 1)
        over_score = round(score_overexposure(over_ratio), 1)
        under_score = round(score_overexposure(under_ratio), 1)
        exposure_score = round(bright_score * 0.6 + over_score * 0.2 + under_score * 0.2, 1)
        exposure_level = level_from_score(exposure_score)
        exposure_desc, exposure_sug = exposure_desc_sug(mean_bright, bright_score, over_score, under_score, exposure_score)

        noise_val = calc_noise(gray)
        noise_score = round(score_noise(noise_val), 1)
        noise_level = level_from_score(noise_score)
        noise_desc, noise_sug = noise_desc_sug(noise_score)

        artifact_val = calc_artifact(gray, width, height)
        artifact_score = round(score_noise(artifact_val), 1)
        artifact_level = level_from_score(artifact_score)
        artifact_desc, artifact_sug = artifact_desc_sug(artifact_score)

        color_metrics = calc_color_cast(image)
        wb_score, wb_status, wb_desc = calc_white_balance(color_metrics)
        hist_metrics = calc_luminance_histogram(gray)

        total_score = round(
            clarity_score * weights[0]
            + exposure_score * weights[1]
            + contrast_score * weights[2]
            + noise_score * weights[3]
            + artifact_score * weights[4],
            1,
        )
        total_level = level_from_score(total_score)
        target_dir = ng_dir if total_score < 50 else ok_dir
        location = target_dir.name
        issue = primary_issue([
            ("清晰度", clarity_score, clarity_desc),
            ("曝光", exposure_score, exposure_desc),
            ("对比度", contrast_score, contrast_desc),
            ("噪声", noise_score, noise_desc),
            ("压缩", artifact_score, artifact_desc),
        ])
        annotation_metrics = {
            "清晰度": {"score": clarity_score},
            "曝光": {"score": exposure_score},
            "对比度": {"score": contrast_score},
            "噪声": {"score": noise_score},
            "压缩": {"score": artifact_score},
            "问题": issue,
        }
        try:
            output_path = self._save_annotated_image(
                image_path,
                target_dir,
                annotate_image(image, total_score, total_level, annotation_metrics),
            )
            output_status = "已保存"
        except Exception as exc:
            output_path = ""
            output_status = f"保存失败：{exc}"
            self.log(f"标注图片保存失败，已保留评分结果：{image_path.name}，原因：{exc}")

        record = {
            "文件名": image_path.name,
            "检测时间": check_time,
            "图片宽度": width,
            "图片高度": height,
            "图片格式": image_path.suffix.lstrip(".").lower(),
            "图片路径": str(image_path),
            "存放位置": location,
            "输出图片路径": str(output_path),
            "标注图状态": output_status,
            "Laplacian值": round(lap_val, 2),
            "Laplacian得分": lap_score,
            "Tenengrad值": round(ten_val, 2),
            "Tenengrad得分": ten_score,
            "清晰度得分": clarity_score,
            "清晰度等级": clarity_level,
            "清晰度说明": clarity_desc,
            "清晰度建议": clarity_sug,
            "灰度标准差": round(std_val, 2),
            "标准差得分": std_score,
            "动态范围": gray_range,
            "动态范围得分": range_score,
            "对比度得分": contrast_score,
            "对比度等级": contrast_level,
            "对比度说明": contrast_desc,
            "对比度建议": contrast_sug,
            "平均亮度": round(mean_bright, 2),
            "亮度得分": bright_score,
            "过曝比例": round(over_ratio, 4),
            "过曝得分": over_score,
            "欠曝比例": round(under_ratio, 4),
            "欠曝得分": under_score,
            "曝光得分": exposure_score,
            "曝光等级": exposure_level,
            "曝光说明": exposure_desc,
            "曝光建议": exposure_sug,
            "噪声估计值": round(noise_val, 2),
            "噪声得分": noise_score,
            "噪声等级": noise_level,
            "噪声说明": noise_desc,
            "噪声建议": noise_sug,
            "块效应分数": round(artifact_val, 2),
            "伪影得分": artifact_score,
            "伪影等级": artifact_level,
            "伪影说明": artifact_desc,
            "伪影建议": artifact_sug,
            "色偏结果": color_metrics["色偏结果"],
            "色偏说明": color_metrics["色偏说明"],
            "R通道均值": color_metrics["R通道均值"],
            "G通道均值": color_metrics["G通道均值"],
            "B通道均值": color_metrics["B通道均值"],
            "通道最大差": color_metrics["通道最大差"],
            "白平衡评分": wb_score,
            "白平衡状态": wb_status,
            "白平衡说明": wb_desc,
            "暗部占比": hist_metrics["暗部占比"],
            "中间调占比": hist_metrics["中间调占比"],
            "高光占比": hist_metrics["高光占比"],
            "死黑占比": hist_metrics["死黑占比"],
            "死白占比": hist_metrics["死白占比"],
            "亮度直方图判断": hist_metrics["直方图判断"],
            "综合评分": total_score,
            "综合等级": total_level,
            "主要问题": issue,
            "综合说明": f"清晰度{clarity_level},对比度{contrast_level},曝光{exposure_level},噪声{noise_level},伪影{artifact_level}",
        }

        if self.config["analysis_mode"] == "区域检测" and region_names:
            for region_name, values in get_region_metrics(gray, region_names).items():
                rlap, rten, rstd = values
                record[f"{region_name}_Laplacian"] = round(rlap, 2)
                record[f"{region_name}_Tenengrad"] = round(rten, 2)
                record[f"{region_name}_灰度标准差"] = round(rstd, 2)

        return record

    def _write_report(self, results, report_dir, sub_label):
        report_dir = Path(report_dir)
        report_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if self.config["analysis_mode"] == "区域检测":
            report_name = f"图像质量检测报告({sub_label})_{timestamp}.xlsx"
        else:
            report_name = f"图像质量检测报告_{timestamp}.xlsx"
        report_path = report_dir / report_name

        wb = Workbook()
        thin_border = Border(
            left=Side("thin", color="D9E2F3"),
            right=Side("thin", color="D9E2F3"),
            top=Side("thin", color="D9E2F3"),
            bottom=Side("thin", color="D9E2F3"),
        )
        fill_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        fill_yellow = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        fill_red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")

        df = pd.DataFrame(results)
        ws1 = wb.active
        ws1.title = "检测明细"
        self._write_dataframe(ws1, df, thin_border, header_font, "DetailTable")
        self._color_score_columns(ws1, list(df.columns), fill_green, fill_yellow, fill_red)
        ws1.freeze_panes = "A2"

        ws2 = wb.create_sheet("统计汇总")
        self._write_summary(ws2, results, thin_border)

        ws3 = wb.create_sheet("阈值配置")
        self._write_config(ws3, thin_border, header_font)

        ws4 = wb.create_sheet("颜色与直方图分析")
        self._write_color_histogram_sheet(ws4, results, thin_border, header_font, fill_green, fill_yellow, fill_red)

        ws5 = wb.create_sheet("问题图片列表")
        self._write_problem_sheet(ws5, results, thin_border, header_font, fill_green, fill_yellow, fill_red)

        wb.save(report_path)
        return report_path

    @staticmethod
    def _display_width(value):
        text = "" if value is None else str(value)
        return sum(2 if ord(char) > 255 else 1 for char in text)

    @staticmethod
    def _is_long_text_column(name):
        keywords = ("说明", "建议", "路径", "规则", "综合说明")
        return any(keyword in str(name) for keyword in keywords)

    @classmethod
    def _fit_columns_and_rows(cls, ws, columns, long_text_max=48):
        widths = {}
        for c_idx, column_name in enumerate(columns, 1):
            letter = get_column_letter(c_idx)
            max_width = cls._display_width(column_name)
            for row in range(2, ws.max_row + 1):
                max_width = max(max_width, cls._display_width(ws.cell(row, c_idx).value))

            if cls._is_long_text_column(column_name):
                width = min(max(max_width + 2, 22), long_text_max)
            elif "文件名" in str(column_name):
                width = min(max(max_width + 2, 18), 34)
            elif "等级" in str(column_name):
                width = 14
            elif "得分" in str(column_name) or "评分" in str(column_name) or "比例" in str(column_name):
                width = 13
            else:
                width = min(max(max_width + 2, 12), 22)
            ws.column_dimensions[letter].width = width
            widths[c_idx] = width

        for row_idx in range(2, ws.max_row + 1):
            estimated_lines = 1
            for c_idx, column_name in enumerate(columns, 1):
                value = ws.cell(row_idx, c_idx).value
                if isinstance(value, str) and (cls._is_long_text_column(column_name) or cls._display_width(value) > widths[c_idx]):
                    estimated_lines = max(estimated_lines, int(cls._display_width(value) / max(widths[c_idx] - 2, 8)) + 1)
            ws.row_dimensions[row_idx].height = min(max(22, estimated_lines * 18), 92)

    @staticmethod
    def _add_excel_table(ws, table_name):
        # Keep the report as plain styled ranges. Excel table XML is fragile when
        # generated from dynamic, multi-sheet reports with images and optional
        # columns, and some Excel versions repair/delete those table parts.
        return

    @classmethod
    def _write_dataframe(cls, ws, df, border, header_font, table_name=None):
        for c_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=c_idx, value=col_name)
            cell.font = header_font
            cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r_idx, row in enumerate(df.values, 2):
            for c_idx, value in enumerate(row, 1):
                col_name = str(df.columns[c_idx - 1])
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                cell.border = border
                horizontal = "center" if isinstance(value, (int, float)) or "等级" in col_name else "left"
                cell.alignment = Alignment(horizontal=horizontal, vertical="top", wrap_text=True)

        ws.sheet_view.showGridLines = False
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 32
        cls._fit_columns_and_rows(ws, list(df.columns))
        if table_name:
            cls._add_excel_table(ws, table_name)

    @staticmethod
    def _color_score_columns(ws, columns, fill_green, fill_yellow, fill_red):
        level_cols = [i for i, name in enumerate(columns, 1) if "等级" in name]
        score_cols = [i for i, name in enumerate(columns, 1) if "得分" in name or "评分" in name]
        for row in range(2, ws.max_row + 1):
            for col in level_cols:
                val = ws.cell(row, col).value
                if val in ("EXCELLENT", "OK"):
                    ws.cell(row, col).fill = fill_green
                elif val == "WARNING":
                    ws.cell(row, col).fill = fill_yellow
                elif val in ("NG", "ERROR"):
                    ws.cell(row, col).fill = fill_red
            for col in score_cols:
                val = ws.cell(row, col).value
                if isinstance(val, (int, float)):
                    ws.cell(row, col).fill = fill_green if val >= 85 else fill_yellow if val >= 50 else fill_red

    @staticmethod
    def _metric_average(results, key):
        values = [r[key] for r in results if isinstance(r.get(key), (int, float)) and r.get("综合等级") != "ERROR"]
        return round(np.mean(values), 1) if values else 0

    @staticmethod
    def _summarize_issues(issue_values):
        counts = {}
        for issue in issue_values:
            if not issue or issue == "无明显问题":
                continue
            counts[issue] = counts.get(issue, 0) + 1
        if not counts:
            return "未发现集中性问题"
        top_items = sorted(counts.items(), key=lambda item: (-item[1], item[0]))[:3]
        return "、".join(issue for issue, _count in top_items)

    def _build_report_summary(self, results):
        total = len(results)
        if not total:
            return "本次未检测到有效文件。"

        levels = [r.get("综合等级") for r in results]
        ok_count = levels.count("EXCELLENT") + levels.count("OK")
        ng_count = levels.count("WARNING") + levels.count("NG") + levels.count("ERROR")
        ng_ratio = ng_count / total * 100 if total else 0
        main_issues = self._summarize_issues(r.get("主要问题") or r.get("综合说明") for r in results)
        return (
            f"本次共检测 {total} 个文件，其中图片 {total} 张、视频 0 个。\n"
            f"检测结果 OK {ok_count} 个，NG/需关注 {ng_count} 个，NG/需关注占比 {ng_ratio:.1f}%。\n"
            f"主要问题为：{main_issues}。"
        )

    def _write_summary(self, ws, results, border):
        total = len(results)
        levels = [r.get("综合等级") for r in results]
        rows = [
            ["检测总数", total],
            ["EXCELLENT 数量", levels.count("EXCELLENT"), f"占比 {levels.count('EXCELLENT') / total * 100:.1f}%" if total else "0%"],
            ["OK 数量", levels.count("OK"), f"占比 {levels.count('OK') / total * 100:.1f}%" if total else "0%"],
            ["WARNING 数量", levels.count("WARNING"), f"占比 {levels.count('WARNING') / total * 100:.1f}%" if total else "0%"],
            ["NG 数量", levels.count("NG"), f"占比 {levels.count('NG') / total * 100:.1f}%" if total else "0%"],
            ["ERROR 数量", levels.count("ERROR")],
            ["平均综合评分", self._metric_average(results, "综合评分")],
            ["平均清晰度得分", self._metric_average(results, "清晰度得分")],
            ["平均对比度得分", self._metric_average(results, "对比度得分")],
            ["平均曝光得分", self._metric_average(results, "曝光得分")],
            ["平均噪声得分", self._metric_average(results, "噪声得分")],
            ["平均伪影得分", self._metric_average(results, "伪影得分")],
        ]
        ws.merge_cells("A1:C3")
        summary_cell = ws["A1"]
        summary_cell.value = self._build_report_summary(results)
        summary_cell.font = Font(bold=True, size=12, color="1F4E78")
        summary_cell.fill = PatternFill(start_color="EAF4FF", end_color="EAF4FF", fill_type="solid")
        summary_cell.border = border
        summary_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 30
        ws.row_dimensions[2].height = 30
        ws.row_dimensions[3].height = 30

        start_row = 5
        for r, row_data in enumerate(rows, start_row):
            for c, value in enumerate(row_data, 1):
                cell = ws.cell(row=r, column=c, value=value)
                cell.border = border
                cell.font = Font(bold=True) if c == 1 else Font()
                cell.alignment = Alignment(horizontal="center" if c > 1 else "left", vertical="center", wrap_text=True)
                if c == 1:
                    cell.fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 24
        ws.column_dimensions["B"].width = 16
        ws.column_dimensions["C"].width = 20
        for row_idx in range(1, ws.max_row + 1):
            if row_idx > 3:
                ws.row_dimensions[row_idx].height = 24

    def _write_config(self, ws, border, header_font):
        weights = self.config["weights"]
        rows = [
            ["参数类别", "参数名称", "设定值/规则"],
            ["分析模式", "当前模式", self.config["analysis_mode"]],
            ["区域模式", "当前区域", self.config["region_mode"]],
            ["权重模式", "当前模式", self.config["weight_mode"]],
            ["清晰度权重", "WEIGHT_CLARITY", weights[0]],
            ["曝光权重", "WEIGHT_EXPOSURE", weights[1]],
            ["对比度权重", "WEIGHT_CONTRAST", weights[2]],
            ["噪声权重", "WEIGHT_NOISE", weights[3]],
            ["伪影权重", "WEIGHT_ARTIFACT", weights[4]],
            ["清晰度指标", "Laplacian 阈值", "EXCELLENT>=300, OK>=150, WARNING>=80"],
            ["", "Tenengrad 阈值", "EXCELLENT>=2000, OK>=800, WARNING>=300"],
            ["曝光指标", "平均亮度理想值", "128，越近得分越高"],
            ["", "过曝/欠曝比例阈值", "EXCELLENT<=1%, OK<=5%, WARNING<=15%"],
            ["对比度指标", "灰度标准差理想范围", "50-80"],
            ["", "动态范围阈值", "EXCELLENT>=180, OK>=100, WARNING>=60"],
            ["噪声指标", "噪声分段阈值", "<=5, <=10, <=20, <=35, >35"],
            ["压缩伪影", "块效应分段阈值", "同噪声分段"],
            ["色偏检测", "偏红", "R 明显高于 G 和 B，判断为偏红 / 偏暖"],
            ["", "偏蓝", "B 明显高于 R 和 G，判断为偏蓝 / 偏冷"],
            ["", "偏绿", "G 明显高于 R 和 B，判断为偏绿"],
            ["", "偏黄", "R 和 G 偏高且 B 偏低，判断为偏黄 / 偏暖"],
            ["", "偏青", "G 和 B 偏高且 R 偏低，判断为偏青"],
            ["", "偏紫/洋红", "R 和 B 偏高且 G 偏低，判断为偏紫 / 偏洋红"],
            ["白平衡评分", "RGB通道最大差 0-10", "95-100 分，优秀"],
            ["", "RGB通道最大差 11-15", "85-94 分，OK"],
            ["", "RGB通道最大差 16-25", "70-84 分，轻微偏色"],
            ["", "RGB通道最大差 26-35", "60-69 分，偏色明显"],
            ["", "RGB通道最大差 >35", "<60 分，NG"],
            ["", "适用范围说明", "白平衡检测基于画面整体 RGB 通道差异进行估算，适用于普通自然光、室内光、商品图、测试样张等场景；对于夕阳、舞台灯光、强色彩场景可能存在误判"],
            ["亮度直方图", "暗部/中间调/高光", "灰度<85 为暗部，85-169 为中间调，>=170 为高光"],
            ["", "死黑/死白", "灰度<=5 占比>=2% 判断死黑；灰度>=250 占比>=2% 判断死白"],
            ["综合评分", "等级划分", ">=85 EXCELLENT, 70-84 OK, 50-69 WARNING, <50 NG"],
        ]
        for r, row_data in enumerate(rows, 1):
            for c, value in enumerate(row_data, 1):
                cell = ws.cell(row=r, column=c, value=value)
                cell.border = border
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                if r == 1:
                    cell.font = header_font
                    cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                elif c in (1, 2):
                    cell.font = Font(bold=True) if value else Font()
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A2"
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 28
        ws.column_dimensions["C"].width = 62
        for row_idx in range(1, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 34 if row_idx > 1 else 30

    def _write_color_histogram_sheet(self, ws, results, border, header_font, fill_green, fill_yellow, fill_red):
        cols = [
            "文件名", "色偏结果", "色偏说明", "R通道均值", "G通道均值", "B通道均值", "通道最大差",
            "白平衡评分", "白平衡状态", "白平衡说明", "暗部占比", "中间调占比", "高光占比", "死黑占比", "死白占比", "亮度直方图判断",
        ]
        rows = []
        for record in results:
            rows.append({col: record.get(col, "") for col in cols})
        df = pd.DataFrame(rows, columns=cols)
        self._write_dataframe(ws, df, border, header_font, "ColorHistogramTable")

        color_col = cols.index("色偏结果") + 1
        wb_col = cols.index("白平衡状态") + 1
        hist_col = cols.index("亮度直方图判断") + 1
        for row in range(2, ws.max_row + 1):
            if ws.cell(row, color_col).value == "正常":
                ws.cell(row, color_col).fill = fill_green
            else:
                ws.cell(row, color_col).fill = fill_yellow

            wb_value = ws.cell(row, wb_col).value
            if wb_value in ("优秀", "OK"):
                ws.cell(row, wb_col).fill = fill_green
            elif wb_value in ("轻微偏色", "偏色明显"):
                ws.cell(row, wb_col).fill = fill_yellow
            else:
                ws.cell(row, wb_col).fill = fill_red

            hist_value = ws.cell(row, hist_col).value
            if hist_value == "亮度分布正常":
                ws.cell(row, hist_col).fill = fill_green
            elif hist_value in ("画面偏暗", "画面偏亮"):
                ws.cell(row, hist_col).fill = fill_yellow
            else:
                ws.cell(row, hist_col).fill = fill_red

        ws.freeze_panes = "A2"

    def _write_problem_sheet(self, ws, results, border, header_font, fill_green, fill_yellow, fill_red):
        problem_records = [r for r in results if r.get("综合等级") in ("WARNING", "NG", "ERROR")]
        if not problem_records:
            ws.cell(1, 1, "无问题图片").font = header_font
            ws.cell(1, 1).fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            ws.cell(1, 1).alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions["A"].width = 24
            ws.sheet_view.showGridLines = False
            return

        key_cols = [
            "文件名", "综合评分", "综合等级", "清晰度得分", "清晰度等级", "对比度得分", "对比度等级",
            "曝光得分", "曝光等级", "噪声得分", "噪声等级", "伪影得分", "伪影等级", "综合说明",
        ]
        prob_df = pd.DataFrame(problem_records)
        prob_df = prob_df[[col for col in key_cols if col in prob_df.columns]]
        self._write_dataframe(ws, prob_df, border, header_font, "ProblemTable")
        img_col = len(prob_df.columns) + 1
        img_letter = get_column_letter(img_col)
        ws.cell(1, img_col, "图片").font = header_font
        ws.cell(1, img_col).fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        ws.cell(1, img_col).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(1, img_col).border = border
        ws.column_dimensions[img_letter].width = 20

        for r_idx, rec in enumerate(problem_records, 2):
            if rec.get("综合等级") == "ERROR":
                continue
            img_path = Path(rec.get("输出图片路径") or "") if rec.get("输出图片路径") else Path(self.config["ng_dir"] if rec.get("综合评分", 0) < 50 else self.config["ok_dir"]) / rec["文件名"]
            if img_path.exists():
                try:
                    img = OpxlImage(str(img_path))
                    img.width, img.height = 100, 75
                    ws.add_image(img, f"{img_letter}{r_idx}")
                    ws.row_dimensions[r_idx].height = 80
                except Exception:
                    pass
            ws.cell(r_idx, img_col).border = border

        self._color_score_columns(ws, list(prob_df.columns), fill_green, fill_yellow, fill_red)
        ws.freeze_panes = "A2"


class ImageQualityApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("图片质量检测工具 Pro")
        self.geometry("1180x760")
        self.minsize(1040, 680)
        self.configure(bg="#090d18")
        self.log_queue = queue.Queue()
        self.worker = None
        self.weight_entries = []

        base = Path.cwd()
        self.input_dir = tk.StringVar(value=str(base / "input_images"))
        self.ok_dir = tk.StringVar(value=str(base / "OK_images"))
        self.ng_dir = tk.StringVar(value=str(base / "NG_images"))
        self.report_dir = tk.StringVar(value=str(base / "report"))
        self.analysis_mode = tk.StringVar(value="整图检测")
        self.region_mode = tk.StringVar(value="九宫格全部")
        self.weight_mode = tk.StringVar(value="默认综合模式")
        self.custom_weight_enabled = tk.BooleanVar(value=False)
        self.weight_vars = [tk.DoubleVar(value=v) for v in WEIGHT_MODES["默认综合模式"]]

        self._build_ui()
        self._poll_log_queue()

    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Root.TFrame", background="#090d18")
        style.configure("Card.TFrame", background="#111827")
        style.configure("Panel.TFrame", background="#0f172a")
        style.configure("Title.TLabel", background="#090d18", foreground="#f8fafc", font=("Microsoft YaHei UI", 24, "bold"))
        style.configure("SubTitle.TLabel", background="#090d18", foreground="#94a3b8", font=("Microsoft YaHei UI", 10))
        style.configure("CardTitle.TLabel", background="#111827", foreground="#e2e8f0", font=("Microsoft YaHei UI", 12, "bold"))
        style.configure("CardText.TLabel", background="#111827", foreground="#94a3b8", font=("Microsoft YaHei UI", 9))
        style.configure("Metric.TLabel", background="#0f172a", foreground="#38bdf8", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("MetricName.TLabel", background="#0f172a", foreground="#94a3b8", font=("Microsoft YaHei UI", 9))
        style.configure("Dark.TLabel", background="#111827", foreground="#cbd5e1", font=("Microsoft YaHei UI", 10))
        style.configure("Dark.TRadiobutton", background="#111827", foreground="#cbd5e1", font=("Microsoft YaHei UI", 10))
        style.map("Dark.TRadiobutton", background=[("active", "#111827")], foreground=[("active", "#f8fafc")])
        style.configure("Dark.TEntry", fieldbackground="#020617", foreground="#e2e8f0", bordercolor="#334155", insertcolor="#e2e8f0")
        style.configure("Dark.TCombobox", fieldbackground="#020617", background="#020617", foreground="#e2e8f0", arrowcolor="#38bdf8")
        style.map("Dark.TCombobox", fieldbackground=[("readonly", "#020617")], foreground=[("readonly", "#e2e8f0")])
        style.configure("Accent.TButton", background="#0ea5e9", foreground="#ffffff", borderwidth=0, focusthickness=0, font=("Microsoft YaHei UI", 11, "bold"), padding=(18, 10))
        style.map("Accent.TButton", background=[("active", "#0284c7"), ("disabled", "#334155")], foreground=[("disabled", "#94a3b8")])
        style.configure("Ghost.TButton", background="#1e293b", foreground="#dbeafe", borderwidth=0, padding=(12, 6))
        style.map("Ghost.TButton", background=[("active", "#334155")])
        style.configure("Cyber.Horizontal.TProgressbar", troughcolor="#020617", background="#22d3ee", bordercolor="#020617", lightcolor="#22d3ee", darkcolor="#06b6d4")

        root = ttk.Frame(self, padding=20, style="Root.TFrame")
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        header = ttk.Frame(root, style="Root.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="图片质量检测工具 Pro", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="清晰度 / 曝光 / 对比度 / 噪声 / 压缩伪影 自动检测与多 Sheet 报告生成", style="SubTitle.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 0))

        status_strip = ttk.Frame(header, style="Root.TFrame")
        status_strip.grid(row=0, column=1, rowspan=2, sticky="e")
        self._metric_card(status_strip, "状态", "待命", 0)
        self._metric_card(status_strip, "图片", "0/0", 1)
        self._metric_card(status_strip, "进度", "0%", 2)

        main = ttk.Frame(root, style="Root.TFrame")
        main.grid(row=1, column=0, sticky="nsew")
        main.columnconfigure(0, weight=3)
        main.columnconfigure(1, weight=2)
        main.rowconfigure(0, weight=1)

        left = ttk.Frame(main, style="Root.TFrame")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(2, weight=1)
        right = ttk.Frame(main, style="Root.TFrame")
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        path_frame = self._card(left, "目录设置", "选择待检测图片、分类输出和报告保存位置")
        path_frame.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        self._path_row(path_frame, "输入图片目录", self.input_dir, 0)
        self._path_row(path_frame, "OK输出目录", self.ok_dir, 1)
        self._path_row(path_frame, "NG输出目录", self.ng_dir, 2)
        self._path_row(path_frame, "报告保存目录", self.report_dir, 3)
        path_frame.columnconfigure(1, weight=1)

        options = ttk.Frame(left, style="Root.TFrame")
        options.grid(row=1, column=0, sticky="ew", pady=(0, 14))
        options.columnconfigure(0, weight=1)
        options.columnconfigure(1, weight=1)
        detect_frame = self._card(options, "检测模式", "整图或指定区域采样分析")
        detect_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 7))
        ttk.Radiobutton(detect_frame, text="整图检测", variable=self.analysis_mode, value="整图检测", command=self._sync_region_state, style="Dark.TRadiobutton").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Radiobutton(detect_frame, text="区域检测", variable=self.analysis_mode, value="区域检测", command=self._sync_region_state, style="Dark.TRadiobutton").grid(row=2, column=1, sticky="w", padx=16, pady=(8, 0))
        ttk.Label(detect_frame, text="区域类型", style="Dark.TLabel").grid(row=3, column=0, sticky="w", pady=(16, 0))
        self.region_combo = ttk.Combobox(detect_frame, textvariable=self.region_mode, values=["九宫格全部", "中心区域", "四角区域"], state="disabled", width=18, style="Dark.TCombobox")
        self.region_combo.grid(row=3, column=1, sticky="ew", pady=(16, 0))

        weight_frame = self._card(options, "评分权重", "选择业务场景或手动调整指标权重")
        weight_frame.grid(row=0, column=1, sticky="nsew", padx=(7, 0))
        self.weight_combo = ttk.Combobox(
            weight_frame,
            textvariable=self.weight_mode,
            values=list(WEIGHT_MODES.keys()) + ["自定义模式"],
            state="readonly",
            width=24,
            style="Dark.TCombobox",
        )
        self.weight_combo.grid(row=2, column=0, columnspan=5, sticky="ew", pady=(8, 0))
        self.weight_combo.bind("<<ComboboxSelected>>", lambda _event: self._on_weight_mode_changed())
        labels = ["清晰度", "曝光", "对比度", "噪声", "伪影"]
        for idx, label in enumerate(labels):
            ttk.Label(weight_frame, text=label, style="Dark.TLabel").grid(row=3, column=idx, padx=4, pady=(14, 2))
            entry = ttk.Entry(weight_frame, textvariable=self.weight_vars[idx], width=8, style="Dark.TEntry", justify="center")
            entry.grid(row=4, column=idx, padx=4, sticky="ew")
            self.weight_entries.append(entry)
            weight_frame.columnconfigure(idx, weight=1)

        action_frame = self._card(left, "控制台", "启动检测后会自动生成标注图片和 Excel 报告")
        action_frame.grid(row=2, column=0, sticky="nsew")
        action_frame.columnconfigure(1, weight=1)
        self.run_button = ttk.Button(action_frame, text="开始检测", style="Accent.TButton", command=self.start_detection)
        self.run_button.grid(row=2, column=0, sticky="w", pady=(8, 12))
        self.progress_bar = ttk.Progressbar(action_frame, mode="determinate", style="Cyber.Horizontal.TProgressbar")
        self.progress_bar.grid(row=2, column=1, sticky="ew", padx=16, pady=(8, 12))
        self.progress_label = ttk.Label(action_frame, text="等待开始", style="Dark.TLabel")
        self.progress_label.grid(row=2, column=2, sticky="e", pady=(8, 12))
        tips = tk.Text(action_frame, height=7, wrap="word", font=("Microsoft YaHei UI", 10), relief="flat", bd=0, bg="#0b1220", fg="#cbd5e1", insertbackground="#cbd5e1")
        tips.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(4, 0))
        tips.insert("end", "工作流\n\n1. 放入待检测图片，或在上方选择图片目录。\n2. 选择检测模式和权重场景。\n3. 点击开始检测，结果会写入 OK/NG 输出目录，并生成多 Sheet Excel 报告。")
        tips.configure(state="disabled")

        log_frame = self._card(right, "运行日志", "实时显示当前检测文件和报告生成状态")
        log_frame.grid(row=0, column=0, rowspan=2, sticky="nsew")
        log_frame.rowconfigure(2, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, height=18, wrap="word", font=("Consolas", 10), relief="flat", bd=0, bg="#020617", fg="#b6e3ff", insertbackground="#e2e8f0", selectbackground="#164e63")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.grid(row=2, column=0, sticky="nsew", pady=(8, 0))
        scrollbar.grid(row=2, column=1, sticky="ns", pady=(8, 0))

        self._on_weight_mode_changed()

    def _card(self, parent, title, subtitle):
        outer = ttk.Frame(parent, style="Card.TFrame", padding=16)
        ttk.Label(outer, text=title, style="CardTitle.TLabel").grid(row=0, column=0, columnspan=8, sticky="w")
        ttk.Label(outer, text=subtitle, style="CardText.TLabel").grid(row=1, column=0, columnspan=8, sticky="w", pady=(4, 8))
        return outer

    def _metric_card(self, parent, name, value, column):
        card = ttk.Frame(parent, style="Panel.TFrame", padding=(16, 10))
        card.grid(row=0, column=column, padx=(10, 0), sticky="n")
        value_label = ttk.Label(card, text=value, style="Metric.TLabel")
        value_label.pack(anchor="center")
        ttk.Label(card, text=name, style="MetricName.TLabel").pack(anchor="center")
        if name == "状态":
            self.status_value = value_label
        elif name == "图片":
            self.count_value = value_label
        elif name == "进度":
            self.percent_value = value_label

    def _path_row(self, parent, label, variable, row):
        actual_row = row + 2
        ttk.Label(parent, text=label, style="Dark.TLabel").grid(row=actual_row, column=0, sticky="w", padx=(0, 10), pady=6)
        ttk.Entry(parent, textvariable=variable, style="Dark.TEntry").grid(row=actual_row, column=1, sticky="ew", pady=6)
        ttk.Button(parent, text="浏览", style="Ghost.TButton", command=lambda: self._choose_dir(variable)).grid(row=actual_row, column=2, sticky="e", padx=(10, 0), pady=6)

    def _choose_dir(self, variable):
        chosen = filedialog.askdirectory(initialdir=variable.get() or str(Path.cwd()))
        if chosen:
            variable.set(chosen)

    def _sync_region_state(self):
        self.region_combo.configure(state="readonly" if self.analysis_mode.get() == "区域检测" else "disabled")

    def _on_weight_mode_changed(self):
        mode = self.weight_mode.get()
        editable = mode == "自定义模式"
        if mode in WEIGHT_MODES:
            for var, value in zip(self.weight_vars, WEIGHT_MODES[mode]):
                var.set(value)
        for entry in self.weight_entries:
            entry.configure(state="normal" if editable else "disabled")

    def _collect_config(self):
        weights = tuple(float(var.get()) for var in self.weight_vars)
        if any(value < 0 for value in weights) or abs(sum(weights) - 1.0) > 0.001:
            raise ValueError("五个权重必须为非负数，且总和必须为 1.0。")
        return {
            "input_dir": self.input_dir.get().strip(),
            "ok_dir": self.ok_dir.get().strip(),
            "ng_dir": self.ng_dir.get().strip(),
            "report_dir": self.report_dir.get().strip(),
            "analysis_mode": self.analysis_mode.get(),
            "region_mode": self.region_mode.get(),
            "weight_mode": self.weight_mode.get(),
            "weights": weights,
        }

    def start_detection(self):
        if self.worker and self.worker.is_alive():
            return
        try:
            config = self._collect_config()
        except Exception as exc:
            messagebox.showerror("配置错误", str(exc))
            return

        self.run_button.configure(state="disabled")
        self.progress_bar.configure(value=0, maximum=100)
        self.progress_label.configure(text="准备中")
        self.status_value.configure(text="运行中")
        self.count_value.configure(text="0/0")
        self.percent_value.configure(text="0%")
        self.log_text.delete("1.0", "end")
        self._log("开始检测任务")

        def task():
            try:
                detector = ImageQualityDetector(config, self._thread_log, self._thread_progress)
                report_path = detector.run()
                self.log_queue.put(("done", str(report_path)))
            except Exception as exc:
                self.log_queue.put(("error", str(exc)))

        self.worker = threading.Thread(target=task, daemon=True)
        self.worker.start()

    def _thread_log(self, message):
        self.log_queue.put(("log", message))

    def _thread_progress(self, done, total):
        self.log_queue.put(("progress", done, total))

    def _poll_log_queue(self):
        try:
            while True:
                item = self.log_queue.get_nowait()
                kind = item[0]
                if kind == "log":
                    self._log(item[1])
                elif kind == "progress":
                    done, total = item[1], item[2]
                    self.progress_bar.configure(maximum=total, value=done)
                    self.progress_label.configure(text=f"{done}/{total}")
                    self.count_value.configure(text=f"{done}/{total}")
                    percent = round(done / total * 100) if total else 0
                    self.percent_value.configure(text=f"{percent}%")
                elif kind == "done":
                    self.run_button.configure(state="normal")
                    self.progress_label.configure(text="完成")
                    self.status_value.configure(text="完成")
                    self.percent_value.configure(text="100%")
                    self._log(f"报告位置：{item[1]}")
                    messagebox.showinfo("检测完成", f"报告已生成：\n{item[1]}")
                elif kind == "error":
                    self.run_button.configure(state="normal")
                    self.progress_label.configure(text="失败")
                    self.status_value.configure(text="失败")
                    self._log(f"错误：{item[1]}")
                    messagebox.showerror("检测失败", item[1])
        except queue.Empty:
            pass
        self.after(100, self._poll_log_queue)

    def _log(self, message):
        now = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{now}] {message}\n")
        self.log_text.see("end")


if __name__ == "__main__":
    app = ImageQualityApp()
    app.mainloop()
