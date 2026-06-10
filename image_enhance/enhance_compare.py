# -*- coding: utf-8 -*-
"""Before/after comparison image generation."""

from __future__ import annotations

from pathlib import Path

import cv2
import numpy as np


def imread_unicode(path: str | Path) -> np.ndarray | None:
    data = np.fromfile(str(path), dtype=np.uint8)
    if data.size == 0:
        return None
    return cv2.imdecode(data, cv2.IMREAD_COLOR)


def imwrite_unicode(path: str | Path, image: np.ndarray) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    suffix = path.suffix if path.suffix else ".jpg"
    ok, encoded = cv2.imencode(suffix, image)
    if not ok:
        raise RuntimeError(f"图片编码失败：{path}")
    path.write_bytes(encoded.tobytes())


def make_compare_image(
    before: np.ndarray,
    after: np.ndarray,
    before_score: float,
    after_score: float,
    used_steps: str,
    output_path: str | Path,
    max_side_width: int = 720,
) -> Path:
    """Create a side-by-side comparison image with simple English labels."""
    output_path = Path(output_path)
    height = max(before.shape[0], after.shape[0])

    def resize_to_height(image: np.ndarray) -> np.ndarray:
        scale = height / image.shape[0]
        width = int(image.shape[1] * scale)
        resized = cv2.resize(image, (width, height), interpolation=cv2.INTER_AREA)
        if resized.shape[1] > max_side_width:
            scale2 = max_side_width / resized.shape[1]
            resized = cv2.resize(
                resized,
                (max_side_width, int(resized.shape[0] * scale2)),
                interpolation=cv2.INTER_AREA,
            )
        return resized

    left = resize_to_height(before)
    right = resize_to_height(after)
    final_height = max(left.shape[0], right.shape[0])

    def pad_height(image: np.ndarray) -> np.ndarray:
        if image.shape[0] == final_height:
            return image
        pad = np.full((final_height - image.shape[0], image.shape[1], 3), 255, dtype=np.uint8)
        return np.vstack([image, pad])

    left = pad_height(left)
    right = pad_height(right)
    gap = np.full((final_height, 12, 3), 245, dtype=np.uint8)
    canvas = np.hstack([left, gap, right])
    header = np.full((86, canvas.shape[1], 3), 255, dtype=np.uint8)
    canvas = np.vstack([header, canvas])

    change = after_score - before_score
    cv2.putText(canvas, f"Before Score: {before_score:.1f}", (18, 28), cv2.FONT_HERSHEY_SIMPLEX, 0.75, (30, 30, 30), 2)
    cv2.putText(canvas, f"After Score: {after_score:.1f}", (18, 58), cv2.FONT_HERSHEY_SIMPLEX, 0.75, (30, 30, 30), 2)
    cv2.putText(canvas, f"Change: {change:+.1f}", (310, 28), cv2.FONT_HERSHEY_SIMPLEX, 0.75, (20, 110, 30) if change >= 0 else (30, 30, 180), 2)
    cv2.putText(canvas, f"Used: {used_steps[:70]}", (310, 58), cv2.FONT_HERSHEY_SIMPLEX, 0.62, (70, 70, 70), 2)

    imwrite_unicode(output_path, canvas)
    return output_path

