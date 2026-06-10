# -*- coding: utf-8 -*-
"""Conservative OpenCV enhancement algorithms used by ImageLab."""

from __future__ import annotations

import cv2
import numpy as np


def gamma_adjust(image: np.ndarray, gamma: float = 0.9) -> np.ndarray:
    """Apply gamma correction. gamma < 1 brightens, gamma > 1 darkens."""
    gamma = max(0.1, float(gamma))
    table = np.array([((i / 255.0) ** gamma) * 255 for i in range(256)], dtype=np.uint8)
    return cv2.LUT(image, table)


def reduce_brightness(image: np.ndarray, beta: int = -12, alpha: float = 0.96) -> np.ndarray:
    """Gently reduce overall brightness for bright or overexposed images."""
    return cv2.convertScaleAbs(image, alpha=float(alpha), beta=int(beta))


def compress_highlights(image: np.ndarray, threshold: int = 205, strength: float = 0.35) -> np.ndarray:
    """Softly compress highlights without changing midtones too much."""
    image_f = image.astype(np.float32)
    mask = image_f > threshold
    image_f[mask] = threshold + (image_f[mask] - threshold) * (1.0 - float(strength))
    return np.clip(image_f, 0, 255).astype(np.uint8)


def clahe_contrast(image: np.ndarray, clip_limit: float = 1.8, tile_grid_size: int = 8) -> np.ndarray:
    """Apply CLAHE on the L channel in LAB color space."""
    lab = cv2.cvtColor(image, cv2.COLOR_BGR2LAB)
    l_channel, a_channel, b_channel = cv2.split(lab)
    clahe = cv2.createCLAHE(
        clipLimit=float(clip_limit),
        tileGridSize=(int(tile_grid_size), int(tile_grid_size)),
    )
    enhanced_l = clahe.apply(l_channel)
    merged = cv2.merge((enhanced_l, a_channel, b_channel))
    return cv2.cvtColor(merged, cv2.COLOR_LAB2BGR)


def denoise_nlm(image: np.ndarray, h: float = 6.0, h_color: float = 6.0) -> np.ndarray:
    """Conservative colored non-local-means denoising."""
    return cv2.fastNlMeansDenoisingColored(
        image,
        None,
        float(h),
        float(h_color),
        7,
        21,
    )


def denoise_median(image: np.ndarray, ksize: int = 3) -> np.ndarray:
    """Median filtering, useful for salt-and-pepper noise."""
    ksize = max(3, int(ksize))
    if ksize % 2 == 0:
        ksize += 1
    return cv2.medianBlur(image, ksize)


def denoise_bilateral(image: np.ndarray, diameter: int = 5, sigma_color: float = 35, sigma_space: float = 35) -> np.ndarray:
    """Bilateral filtering for edge-preserving smoothing."""
    return cv2.bilateralFilter(image, int(diameter), float(sigma_color), float(sigma_space))


def unsharp_mask(image: np.ndarray, amount: float = 0.25, sigma: float = 1.0) -> np.ndarray:
    """Light unsharp masking. Keep amount conservative to avoid halos/noise."""
    blurred = cv2.GaussianBlur(image, (0, 0), float(sigma))
    return cv2.addWeighted(image, 1.0 + float(amount), blurred, -float(amount), 0)


def ocr_preprocess(image: np.ndarray, clip_limit: float = 2.0, tile_grid_size: int = 8, denoise_h: float = 6.0) -> np.ndarray:
    """OCR-oriented preprocessing: gray -> CLAHE -> denoise -> Otsu binarization."""
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(
        clipLimit=float(clip_limit),
        tileGridSize=(int(tile_grid_size), int(tile_grid_size)),
    )
    gray = clahe.apply(gray)
    gray = cv2.fastNlMeansDenoising(gray, None, float(denoise_h), 7, 21)
    _threshold, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return cv2.cvtColor(binary, cv2.COLOR_GRAY2BGR)

