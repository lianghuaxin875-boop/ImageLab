# -*- coding: utf-8 -*-
"""Quality evaluation and rule selection for ImageLab image enhancement."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import cv2
import numpy as np

from image_quality_gui import (
    WEIGHT_MODES,
    artifact_desc_sug,
    calc_artifact,
    calc_clarity_metrics,
    calc_noise,
    clarity_desc_sug,
    contrast_desc_sug,
    exposure_desc_sug,
    level_from_score,
    noise_desc_sug,
    primary_issue,
    score_brightness,
    score_dynamic_range,
    score_laplacian,
    score_noise,
    score_overexposure,
    score_std,
    score_tenengrad,
)


@dataclass
class EnhanceStep:
    name: str
    label: str
    params: dict[str, Any] = field(default_factory=dict)
    reason: str = ""


def normalize_strength(strength: str | None = "standard") -> str:
    value = str(strength or "standard").strip().lower()
    aliases = {
        "light": "light",
        "soft": "light",
        "mild": "light",
        "轻度": "light",
        "standard": "standard",
        "normal": "standard",
        "标准": "standard",
        "strong": "strong",
        "obvious": "strong",
        "明显": "strong",
    }
    return aliases.get(value, "standard")


def strength_params(strength: str | None = "standard") -> dict[str, float]:
    """Parameter presets. Strong is more visible but still avoids extreme processing."""
    presets = {
        "light": {
            "gamma_dark": 0.90,
            "gamma_very_dark": 0.75,
            "clahe_clip": 1.6,
            "denoise_h": 6.0,
            "bilateral_sigma": 30.0,
            "sharpen_noise": 0.18,
            "sharpen": 0.25,
            "darken_beta": -8,
            "highlight_strength": 0.30,
            "ocr_clip": 2.0,
        },
        "standard": {
            "gamma_dark": 0.82,
            "gamma_very_dark": 0.68,
            "clahe_clip": 2.0,
            "denoise_h": 7.0,
            "bilateral_sigma": 35.0,
            "sharpen_noise": 0.22,
            "sharpen": 0.32,
            "darken_beta": -10,
            "highlight_strength": 0.38,
            "ocr_clip": 2.2,
        },
        "strong": {
            "gamma_dark": 0.74,
            "gamma_very_dark": 0.60,
            "clahe_clip": 2.4,
            "denoise_h": 8.0,
            "bilateral_sigma": 42.0,
            "sharpen_noise": 0.24,
            "sharpen": 0.40,
            "darken_beta": -14,
            "highlight_strength": 0.48,
            "ocr_clip": 2.5,
        },
    }
    return presets[normalize_strength(strength)]


def evaluate_quality(image: np.ndarray, image_path: str | Path = "") -> dict[str, Any]:
    """Evaluate one image with the same scoring logic as the existing quality module."""
    if image is None:
        return {
            "文件名": Path(image_path).name if image_path else "",
            "综合评分": 0.0,
            "综合等级": "ERROR",
            "主要问题": "图片无法读取",
            "综合说明": "图片无法读取",
        }

    height, width = image.shape[:2]
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    weights = WEIGHT_MODES.get("默认综合模式", (0.4, 0.2, 0.2, 0.1, 0.1))

    lap_val, ten_val = calc_clarity_metrics(gray)
    lap_score = round(score_laplacian(lap_val), 1)
    ten_score = round(score_tenengrad(ten_val), 1)
    clarity_score = round(lap_score * 0.5 + ten_score * 0.5, 1)
    clarity_level = level_from_score(clarity_score)
    clarity_desc, clarity_sug = clarity_desc_sug(lap_score, ten_score, clarity_score)

    gray_min, gray_max = int(gray.min()), int(gray.max())
    gray_range = gray_max - gray_min
    std_val = float(np.std(gray))
    std_score = round(score_std(std_val), 1)
    range_score = round(score_dynamic_range(gray_range), 1)
    contrast_score = round(std_score * 0.7 + range_score * 0.3, 1)
    contrast_level = level_from_score(contrast_score)
    contrast_desc, contrast_sug = contrast_desc_sug(std_score, range_score, contrast_score)

    mean_bright = float(gray.mean())
    over_ratio = float(np.sum(gray > 240) / gray.size)
    under_ratio = float(np.sum(gray < 30) / gray.size)
    bright_score = round(score_brightness(mean_bright), 1)
    over_score = round(score_overexposure(over_ratio), 1)
    under_score = round(score_overexposure(under_ratio), 1)
    exposure_score = round(bright_score * 0.6 + over_score * 0.2 + under_score * 0.2, 1)
    exposure_level = level_from_score(exposure_score)
    exposure_desc, exposure_sug = exposure_desc_sug(mean_bright, bright_score, over_score, under_score, exposure_score)

    noise_val = float(calc_noise(gray))
    noise_score = round(score_noise(noise_val), 1)
    noise_level = level_from_score(noise_score)
    noise_desc, noise_sug = noise_desc_sug(noise_score)

    artifact_val = float(calc_artifact(gray, width, height))
    artifact_score = round(score_noise(artifact_val), 1)
    artifact_level = level_from_score(artifact_score)
    artifact_desc, artifact_sug = artifact_desc_sug(artifact_score)

    total_score = round(
        clarity_score * weights[0]
        + exposure_score * weights[1]
        + contrast_score * weights[2]
        + noise_score * weights[3]
        + artifact_score * weights[4],
        1,
    )
    total_level = level_from_score(total_score)
    issue = primary_issue([
        ("清晰度", clarity_score, clarity_desc),
        ("曝光", exposure_score, exposure_desc),
        ("对比度", contrast_score, contrast_desc),
        ("噪声", noise_score, noise_desc),
        ("压缩", artifact_score, artifact_desc),
    ])

    return {
        "文件名": Path(image_path).name if image_path else "",
        "图片路径": str(image_path) if image_path else "",
        "图片宽度": width,
        "图片高度": height,
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
        "综合评分": total_score,
        "综合等级": total_level,
        "主要问题": issue,
        "综合说明": f"清晰度{clarity_level},对比度{contrast_level},曝光{exposure_level},噪声{noise_level},伪影{artifact_level}",
    }


def choose_rule_steps(metrics: dict[str, Any], ocr_mode: bool = False, strength: str = "standard") -> list[EnhanceStep]:
    """Choose conservative enhancement steps based on measured quality issues."""
    steps: list[EnhanceStep] = []

    mean_bright = float(metrics.get("平均亮度", 0) or 0)
    over_ratio = float(metrics.get("过曝比例", 0) or 0)
    under_ratio = float(metrics.get("欠曝比例", 0) or 0)
    std_val = float(metrics.get("灰度标准差", 0) or 0)
    dynamic_range = float(metrics.get("动态范围", 0) or 0)
    noise_val = float(metrics.get("噪声估计值", 0) or 0)
    noise_level = str(metrics.get("噪声等级", ""))
    lap_val = float(metrics.get("Laplacian值", 0) or 0)
    ten_val = float(metrics.get("Tenengrad值", 0) or 0)
    artifact_val = float(metrics.get("块效应分数", 0) or 0)
    artifact_score = float(metrics.get("伪影得分", 100) or 100)

    noise_high = noise_val > 20 or noise_level in {"WARNING", "NG", "ERROR"}
    blur_low = lap_val < 150 or ten_val < 800
    artifact_obvious = artifact_val > 20 or artifact_score < 70

    if noise_high:
        steps.append(EnhanceStep("denoise_nlm", "去噪", {"h": 7.0, "h_color": 7.0}, "噪声估计值较高"))
    elif artifact_obvious:
        steps.append(EnhanceStep("denoise_bilateral", "轻度平滑", {"diameter": 5, "sigma_color": 30, "sigma_space": 30}, "压缩伪影或块效应明显"))

    if mean_bright < 70 or under_ratio > 0.12:
        steps.append(EnhanceStep("gamma", "Gamma增亮", {"gamma": 0.7}, "严重偏暗或欠曝比例高"))
    elif mean_bright < 90 or under_ratio > 0.05:
        steps.append(EnhanceStep("gamma", "Gamma增亮", {"gamma": 0.85}, "图片偏暗或存在欠曝"))
    elif mean_bright > 190 or over_ratio > 0.12:
        steps.append(EnhanceStep("highlight", "高光压制", {"threshold": 205, "strength": 0.35}, "高光或过曝比例较高"))
        steps.append(EnhanceStep("darken", "亮度降低", {"beta": -10, "alpha": 0.96}, "整体亮度偏高"))
    elif mean_bright > 170 or over_ratio > 0.05:
        steps.append(EnhanceStep("darken", "亮度降低", {"beta": -8, "alpha": 0.97}, "图片偏亮或轻微过曝"))

    if std_val < 35 or dynamic_range < 100:
        steps.append(EnhanceStep("clahe", "CLAHE对比度增强", {"clip_limit": 1.6, "tile_grid_size": 8}, "对比度或动态范围偏低"))

    if blur_low and not artifact_obvious:
        amount = 0.2 if noise_high else 0.28
        steps.append(EnhanceStep("sharpen", "轻度锐化", {"amount": amount, "sigma": 1.0}, "清晰度或边缘梯度偏低"))

    if ocr_mode:
        steps.append(EnhanceStep("ocr", "OCR二值化预处理", {"clip_limit": 2.0, "tile_grid_size": 8, "denoise_h": 6.0}, "OCR模式需要提升文字可读性"))

    if not steps:
        steps.append(EnhanceStep("none", "未执行明显增强", {}, "图片质量正常"))

    return tune_steps_for_strength(steps, strength)


def tune_steps_for_strength(steps: list[EnhanceStep], strength: str = "standard") -> list[EnhanceStep]:
    params = strength_params(strength)
    tuned: list[EnhanceStep] = []
    for step in steps:
        step_params = dict(step.params)
        if step.name == "gamma":
            current = float(step_params.get("gamma", params["gamma_dark"]))
            step_params["gamma"] = params["gamma_very_dark"] if current <= 0.75 else params["gamma_dark"]
        elif step.name == "clahe":
            step_params["clip_limit"] = params["clahe_clip"]
        elif step.name == "denoise_nlm":
            step_params["h"] = params["denoise_h"]
            step_params["h_color"] = params["denoise_h"]
        elif step.name == "denoise_bilateral":
            step_params["sigma_color"] = params["bilateral_sigma"]
            step_params["sigma_space"] = params["bilateral_sigma"]
        elif step.name == "sharpen":
            current = float(step_params.get("amount", params["sharpen"]))
            step_params["amount"] = params["sharpen_noise"] if current <= 0.22 else params["sharpen"]
        elif step.name == "highlight":
            step_params["strength"] = params["highlight_strength"]
        elif step.name == "darken":
            step_params["beta"] = params["darken_beta"]
        elif step.name == "ocr":
            step_params["clip_limit"] = params["ocr_clip"]
            step_params["denoise_h"] = params["denoise_h"]
        tuned.append(EnhanceStep(step.name, step.label, step_params, step.reason))
    return tuned


def fixed_steps() -> list[EnhanceStep]:
    return [
        EnhanceStep("gamma", "Gamma增亮", {"gamma": 0.9}, "固定自动流程"),
        EnhanceStep("clahe", "CLAHE对比度增强", {"clip_limit": 1.5, "tile_grid_size": 8}, "固定自动流程"),
        EnhanceStep("denoise_nlm", "轻度去噪", {"h": 5.0, "h_color": 5.0}, "固定自动流程"),
        EnhanceStep("sharpen", "轻度锐化", {"amount": 0.28, "sigma": 1.0}, "固定自动流程"),
    ]
