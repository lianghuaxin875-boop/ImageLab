# -*- coding: utf-8 -*-
"""Rule-based automatic image enhancement test script for ImageLab.

Examples:
    python -m image_enhance.enhance_main --input input_images --mode rule
    python -m image_enhance.enhance_main --input "a.jpg" --mode fixed
    python -m image_enhance.enhance_main --input input_images --mode manual --steps gamma,clahe,denoise,sharpen
    python -m image_enhance.enhance_main --input input_images --mode rule --ocr-mode
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any

import cv2
import numpy as np

if __package__ in {None, ""}:
    sys.path.append(str(Path(__file__).resolve().parents[1]))

from image_quality_gui import IMAGE_EXTENSIONS
from image_enhance.enhance_algorithms import (
    clahe_contrast,
    compress_highlights,
    denoise_bilateral,
    denoise_median,
    denoise_nlm,
    gamma_adjust,
    ocr_preprocess,
    reduce_brightness,
    unsharp_mask,
)
from image_enhance.enhance_compare import imread_unicode, imwrite_unicode, make_compare_image
from image_enhance.enhance_report import write_reports
from image_enhance.enhance_rules import EnhanceStep, choose_rule_steps, evaluate_quality, fixed_steps, tune_steps_for_strength


def ensure_output_dirs(output_dir: str | Path) -> dict[str, Path]:
    root = Path(output_dir)
    dirs = {
        "root": root,
        "enhanced": root / "enhanced_images",
        "compare": root / "compare_images",
        "report": root / "report",
        "logs": root / "logs",
    }
    for directory in dirs.values():
        directory.mkdir(parents=True, exist_ok=True)
    return dirs


def iter_images(input_path: str | Path) -> list[Path]:
    path = Path(input_path)
    if path.is_file():
        if path.suffix.lower() not in IMAGE_EXTENSIONS:
            raise ValueError(f"不支持的图片格式：{path.suffix}")
        return [path]
    if not path.exists():
        raise FileNotFoundError(f"输入路径不存在：{path}")
    files = [item for item in path.iterdir() if item.is_file() and item.suffix.lower() in IMAGE_EXTENSIONS]
    if not files:
        raise RuntimeError(f"输入目录中没有找到图片：{path}")
    return sorted(files, key=lambda item: item.name.lower())


def safe_output_path(source_path: Path, target_dir: Path, suffix: str = "") -> Path:
    suffix_text = suffix or "_enhanced"
    output = target_dir / f"{source_path.stem}{suffix_text}{source_path.suffix}"
    index = 2
    while output.exists():
        output = target_dir / f"{source_path.stem}{suffix_text}_{index}{source_path.suffix}"
        index += 1
    return output


def manual_steps_from_args(args: argparse.Namespace) -> list[EnhanceStep]:
    mapping = {
        "gamma": EnhanceStep("gamma", "Gamma增亮", {"gamma": args.gamma}, "手动选择"),
        "brighten": EnhanceStep("gamma", "Gamma增亮", {"gamma": args.gamma}, "手动选择"),
        "darken": EnhanceStep("darken", "亮度降低", {"beta": args.darken_beta, "alpha": args.darken_alpha}, "手动选择"),
        "highlight": EnhanceStep("highlight", "高光压制", {"threshold": args.highlight_threshold, "strength": args.highlight_strength}, "手动选择"),
        "clahe": EnhanceStep("clahe", "CLAHE对比度增强", {"clip_limit": args.clahe_clip, "tile_grid_size": args.clahe_grid}, "手动选择"),
        "denoise": EnhanceStep("denoise_nlm", "去噪", {"h": args.denoise_h, "h_color": args.denoise_h_color}, "手动选择"),
        "median": EnhanceStep("denoise_median", "中值滤波", {"ksize": args.median_ksize}, "手动选择"),
        "bilateral": EnhanceStep("denoise_bilateral", "双边滤波", {"diameter": args.bilateral_diameter, "sigma_color": args.bilateral_sigma_color, "sigma_space": args.bilateral_sigma_space}, "手动选择"),
        "sharpen": EnhanceStep("sharpen", "轻度锐化", {"amount": args.sharpen_amount, "sigma": args.sharpen_sigma}, "手动选择"),
        "ocr": EnhanceStep("ocr", "OCR二值化预处理", {"clip_limit": args.ocr_clip, "tile_grid_size": args.ocr_grid, "denoise_h": args.ocr_denoise_h}, "手动选择"),
        "ocr_binary": EnhanceStep("ocr", "OCR二值化预处理", {"clip_limit": args.ocr_clip, "tile_grid_size": args.ocr_grid, "denoise_h": args.ocr_denoise_h}, "手动选择"),
    }
    names = [item.strip().lower() for item in args.steps.split(",") if item.strip()]
    steps = []
    for name in names:
        if name not in mapping:
            raise ValueError(f"未知手动增强步骤：{name}，可选：{', '.join(mapping)}")
        steps.append(mapping[name])
    return steps or [EnhanceStep("none", "未执行明显增强", {}, "未选择手动步骤")]


def select_steps(mode: str, before_metrics: dict[str, Any], args: argparse.Namespace | None = None, ocr_mode: bool = False) -> list[EnhanceStep]:
    strength = getattr(args, "enhance_strength", "standard") if args is not None else "standard"
    if mode == "rule":
        return choose_rule_steps(before_metrics, ocr_mode=ocr_mode, strength=strength)
    if mode == "fixed":
        steps = tune_steps_for_strength(fixed_steps(), strength)
        if ocr_mode:
            steps.append(EnhanceStep("ocr", "OCR二值化预处理", {"clip_limit": 2.0, "tile_grid_size": 8, "denoise_h": 6.0}, "OCR模式"))
        return steps
    if mode == "manual":
        if args is None:
            raise ValueError("手动模式需要传入参数")
        return tune_steps_for_strength(manual_steps_from_args(args), strength)
    raise ValueError(f"未知增强模式：{mode}")


def apply_steps(image: np.ndarray, steps: list[EnhanceStep]) -> np.ndarray:
    result = image.copy()
    for step in steps:
        if step.name == "none":
            continue
        if step.name == "gamma":
            result = gamma_adjust(result, **step.params)
        elif step.name == "darken":
            result = reduce_brightness(result, **step.params)
        elif step.name == "highlight":
            result = compress_highlights(result, **step.params)
        elif step.name == "clahe":
            result = clahe_contrast(result, **step.params)
        elif step.name == "denoise_nlm":
            result = denoise_nlm(result, **step.params)
        elif step.name == "denoise_median":
            result = denoise_median(result, **step.params)
        elif step.name == "denoise_bilateral":
            result = denoise_bilateral(result, **step.params)
        elif step.name == "sharpen":
            result = unsharp_mask(result, **step.params)
        elif step.name == "ocr":
            result = ocr_preprocess(result, **step.params)
        else:
            raise ValueError(f"未知增强步骤：{step.name}")
    return result


def judge_effect(before: dict[str, Any], after: dict[str, Any]) -> tuple[str, str]:
    before_score = float(before.get("综合评分", 0) or 0)
    after_score = float(after.get("综合评分", 0) or 0)
    delta = after_score - before_score
    advice = "可采用增强后图片"
    if delta >= 5:
        effect = "增强有效"
    elif delta >= -5:
        effect = "效果不明显"
        advice = "建议结合视觉效果人工复核，必要时微调参数"
    else:
        effect = "增强可能过度或无效"
        advice = "建议降低增强强度或保留原图"

    clarity_delta = float(after.get("清晰度得分", 0) or 0) - float(before.get("清晰度得分", 0) or 0)
    noise_delta = float(after.get("噪声得分", 0) or 0) - float(before.get("噪声得分", 0) or 0)
    if clarity_delta >= 5 and noise_delta <= -8:
        advice = "锐化可能过强，建议降低锐化强度"
    elif noise_delta >= 5 and clarity_delta <= -8:
        advice = "去噪可能过强，建议降低去噪强度"
    return effect, advice


def build_report_row(
    image_path: Path,
    enhanced_path: Path,
    compare_path: Path,
    before: dict[str, Any],
    after: dict[str, Any],
    steps: list[EnhanceStep],
) -> dict[str, Any]:
    before_score = float(before.get("综合评分", 0) or 0)
    after_score = float(after.get("综合评分", 0) or 0)
    effect, advice = judge_effect(before, after)
    used = " + ".join(step.label for step in steps) if steps else "未执行明显增强"
    return {
        "检测时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "文件名": image_path.name,
        "原图路径": str(image_path),
        "增强图路径": str(enhanced_path),
        "对比图路径": str(compare_path),
        "原图综合分": before_score,
        "增强后综合分": after_score,
        "分数变化": round(after_score - before_score, 1),
        "增强效果判断": effect,
        "原图质量等级": before.get("综合等级", ""),
        "增强后质量等级": after.get("综合等级", ""),
        "原图问题原因": before.get("主要问题", ""),
        "执行增强方法": used,
        "增强后问题原因": after.get("主要问题", ""),
        "处理建议": advice,
        "原图清晰度分": before.get("清晰度得分", ""),
        "增强后清晰度分": after.get("清晰度得分", ""),
        "原图曝光分": before.get("曝光得分", ""),
        "增强后曝光分": after.get("曝光得分", ""),
        "原图对比度分": before.get("对比度得分", ""),
        "增强后对比度分": after.get("对比度得分", ""),
        "原图噪声分": before.get("噪声得分", ""),
        "增强后噪声分": after.get("噪声得分", ""),
        "原图压缩伪影分": before.get("伪影得分", ""),
        "增强后压缩伪影分": after.get("伪影得分", ""),
    }


def process_one_image(
    image_path: str | Path,
    output_dirs: dict[str, Path],
    mode: str = "rule",
    args: argparse.Namespace | None = None,
    ocr_mode: bool = False,
) -> dict[str, Any]:
    image_path = Path(image_path)
    image = imread_unicode(image_path)
    if image is None:
        raise RuntimeError(f"图片无法读取：{image_path}")

    before = evaluate_quality(image, image_path)
    steps = select_steps(mode, before, args=args, ocr_mode=ocr_mode)
    enhanced = apply_steps(image, steps)
    after = evaluate_quality(enhanced, image_path)

    enhanced_path = safe_output_path(image_path, output_dirs["enhanced"], "_enhanced")
    compare_path = safe_output_path(image_path, output_dirs["compare"], "_compare")
    imwrite_unicode(enhanced_path, enhanced)
    make_compare_image(
        image,
        enhanced,
        float(before.get("综合评分", 0) or 0),
        float(after.get("综合评分", 0) or 0),
        " + ".join(step.label for step in steps),
        compare_path,
    )
    row = build_report_row(image_path, enhanced_path, compare_path, before, after, steps)
    row["_steps"] = [asdict(step) for step in steps]
    return row


def run_enhancement(
    input_path: str | Path,
    output_dir: str | Path = "enhance_output",
    mode: str = "rule",
    args: argparse.Namespace | None = None,
    ocr_mode: bool = False,
) -> dict[str, Any]:
    output_dirs = ensure_output_dirs(output_dir)
    files = iter_images(input_path)
    rows: list[dict[str, Any]] = []
    log_lines = [
        f"开始图像增强：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"输入：{input_path}",
        f"输出：{output_dirs['root']}",
        f"模式：{mode}",
        f"OCR模式：{ocr_mode}",
    ]

    for index, image_path in enumerate(files, 1):
        try:
            row = process_one_image(image_path, output_dirs, mode=mode, args=args, ocr_mode=ocr_mode)
            rows.append(row)
            log_lines.append(
                f"[{index}/{len(files)}] {image_path.name} 原分 {row['原图综合分']} -> 增强后 {row['增强后综合分']}，{row['增强效果判断']}，步骤：{row['执行增强方法']}"
            )
        except Exception as exc:
            log_lines.append(f"[{index}/{len(files)}] {image_path.name} 处理失败：{exc}")

    xlsx_path, csv_path = write_reports(rows, output_dirs["report"])
    log_lines.append(f"Excel报告：{xlsx_path}")
    log_lines.append(f"CSV报告：{csv_path}")
    log_path = output_dirs["logs"] / f"enhance_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    log_path.write_text("\n".join(log_lines), encoding="utf-8")

    return {
        "ok": True,
        "total": len(files),
        "success": len(rows),
        "failed": len(files) - len(rows),
        "output_dir": str(output_dirs["root"]),
        "enhanced_dir": str(output_dirs["enhanced"]),
        "compare_dir": str(output_dirs["compare"]),
        "excel_report": str(xlsx_path),
        "csv_report": str(csv_path),
        "log_path": str(log_path),
        "rows": rows,
    }


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="ImageLab 规则版自动图像增强测试脚本")
    parser.add_argument("--input", required=True, help="单张图片路径或图片文件夹路径")
    parser.add_argument("--output", default="enhance_output", help="输出目录，默认 enhance_output")
    parser.add_argument("--mode", choices=["manual", "fixed", "rule"], default="rule", help="增强模式")
    parser.add_argument("--enhance-strength", choices=["light", "standard", "strong"], default="standard", help="增强强度：light/standard/strong")
    parser.add_argument("--ocr-mode", action="store_true", help="启用 OCR 二值化预处理，仅适合文字/标签/文档图片")
    parser.add_argument("--steps", default="", help="手动模式步骤，逗号分隔：gamma,darken,highlight,clahe,denoise,median,bilateral,sharpen,ocr")

    parser.add_argument("--gamma", type=float, default=0.85)
    parser.add_argument("--darken-alpha", type=float, default=0.96)
    parser.add_argument("--darken-beta", type=int, default=-10)
    parser.add_argument("--highlight-threshold", type=int, default=205)
    parser.add_argument("--highlight-strength", type=float, default=0.35)
    parser.add_argument("--clahe-clip", type=float, default=1.8)
    parser.add_argument("--clahe-grid", type=int, default=8)
    parser.add_argument("--denoise-h", type=float, default=6.0)
    parser.add_argument("--denoise-h-color", type=float, default=6.0)
    parser.add_argument("--median-ksize", type=int, default=3)
    parser.add_argument("--bilateral-diameter", type=int, default=5)
    parser.add_argument("--bilateral-sigma-color", type=float, default=35)
    parser.add_argument("--bilateral-sigma-space", type=float, default=35)
    parser.add_argument("--sharpen-amount", type=float, default=0.25)
    parser.add_argument("--sharpen-sigma", type=float, default=1.0)
    parser.add_argument("--ocr-clip", type=float, default=2.0)
    parser.add_argument("--ocr-grid", type=int, default=8)
    parser.add_argument("--ocr-denoise-h", type=float, default=6.0)
    return parser


def main() -> None:
    parser = build_arg_parser()
    args = parser.parse_args()
    result = run_enhancement(
        input_path=args.input,
        output_dir=args.output,
        mode=args.mode,
        args=args,
        ocr_mode=args.ocr_mode,
    )
    printable = {key: value for key, value in result.items() if key != "rows"}
    print(json.dumps(printable, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
