#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OCR 图像识别一体化脚本 - PaddleOCR 3.x 完整可运行版

功能：
1. 支持单张图片识别
2. 支持文件夹批量识别
3. 支持视频抽帧识别
4. 支持中文路径、特殊符号路径读取
5. 支持 PaddleOCR 3.x predict 结果解析
6. 支持识别文本、置信度、文字框输出
7. 支持 OK / NG 判定
8. 支持 TXT 文本输出
9. 支持 Excel 报告输出
10. 支持标注图片输出
11. 支持 NG 图片单独保存

推荐版本：
paddleocr==3.1.1
paddlepaddle==3.1.1
paddlex==3.1.3

运行示例：
python "D:\\Codex\\OCR Image Recognition.py"

或者：
python "D:\\Codex\\OCR Image Recognition.py" --source "D:\\Codex\\OCR测试" --mode batch
"""

import os

# 必须放在 paddleocr 导入前
os.environ["FLAGS_use_mkldnn"] = "0"
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"
os.environ["PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"] = "True"

import re
import cv2
import shutil
import argparse
import datetime
import numpy as np

from typing import List, Dict, Tuple, Generator, Optional
from openpyxl import Workbook


try:
    from paddleocr import PaddleOCR
except ImportError:
    print("❌ 缺少 paddleocr，请先安装：")
    print("pip install paddleocr")
    raise SystemExit(1)


IMAGE_EXTS = [".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"]
VIDEO_EXTS = [".mp4", ".avi", ".mov", ".mkv", ".wmv", ".flv", ".mpeg", ".mpg"]


# =========================
# 工具函数
# =========================

def read_image_chinese_path(path: str) -> Optional[np.ndarray]:
    """
    支持中文路径、特殊符号路径的图片读取
    """
    if not os.path.exists(path):
        print(f"❌ 文件不存在：{path}")
        return None

    try:
        data = np.fromfile(path, dtype=np.uint8)
        image = cv2.imdecode(data, cv2.IMREAD_COLOR)

        if image is None:
            print(f"❌ 图片解码失败，可能文件损坏或格式不支持：{path}")
            return None

        return image

    except Exception as e:
        print(f"❌ 读取图片失败：{path}")
        print(e)
        return None


def write_image_chinese_path(path: str, image: np.ndarray) -> bool:
    """
    支持中文路径、特殊符号路径的图片保存
    """
    try:
        ext = os.path.splitext(path)[1]

        if ext == "":
            ext = ".jpg"
            path += ext

        ok, buffer = cv2.imencode(ext, image)

        if ok:
            buffer.tofile(path)
            return True

        return False

    except Exception as e:
        print(f"❌ 保存图片失败：{path}")
        print(e)
        return False


def safe_filename(name: str) -> str:
    """
    清理 Windows 文件名非法字符
    """
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def normalize_confidence(value: float) -> float:
    """
    支持 0.8 或 80 两种输入
    """
    if value > 1:
        return value / 100
    return value


def infer_mode(source: str) -> str:
    """
    自动判断 single / batch / video
    """
    if os.path.isdir(source):
        return "batch"

    ext = os.path.splitext(source)[1].lower()

    if ext in VIDEO_EXTS:
        return "video"

    return "single"


def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


# =========================
# 模块1：输入模块
# =========================

class InputModule:
    """
    负责读取单张图片、批量图片、视频抽帧
    """

    def __init__(self, source: str, mode: str = "single", frame_interval: int = 30):
        self.source = source
        self.mode = mode
        self.frame_interval = max(1, frame_interval)

    def get_data(self) -> Generator[Tuple[str, np.ndarray, Optional[str]], None, None]:
        if self.mode == "single":
            yield from self._read_single()

        elif self.mode == "batch":
            yield from self._read_batch()

        elif self.mode == "video":
            yield from self._read_video()

        else:
            raise ValueError(f"不支持的模式：{self.mode}")

    def _read_single(self):
        image = read_image_chinese_path(self.source)

        if image is not None:
            name = os.path.basename(self.source)
            yield name, image, self.source

    def _read_batch(self):
        if not os.path.isdir(self.source):
            raise ValueError(f"批量模式要求 source 是文件夹：{self.source}")

        image_paths = []

        skip_dirs = {
            "ocr_output",
            "reports",
            "texts",
            "annotated_images",
            "ng_images"
        }

        for root, dirs, files in os.walk(self.source):
            dirs[:] = [d for d in dirs if d not in skip_dirs]

            for file in files:
                ext = os.path.splitext(file)[1].lower()

                if ext in IMAGE_EXTS:
                    image_paths.append(os.path.join(root, file))

        image_paths = sorted(image_paths)

        if not image_paths:
            print("⚠️ 当前文件夹没有找到图片。")
            return

        for path in image_paths:
            image = read_image_chinese_path(path)

            if image is not None:
                name = os.path.basename(path)
                yield name, image, path

    def _read_video(self):
        if not os.path.exists(self.source):
            raise FileNotFoundError(f"视频文件不存在：{self.source}")

        cap = cv2.VideoCapture(self.source)

        if not cap.isOpened():
            raise IOError(f"无法打开视频：{self.source}")

        frame_index = 0
        saved_index = 0

        while True:
            ret, frame = cap.read()

            if not ret:
                break

            if frame_index % self.frame_interval == 0:
                name = f"frame_{saved_index:06d}.jpg"
                yield name, frame, None
                saved_index += 1

            frame_index += 1

        cap.release()


# =========================
# 模块2：OCR 引擎
# =========================

class OCREngine:
    """
    PaddleOCR 3.x OCR 引擎
    """

    def __init__(self, lang: str = "ch", use_gpu: bool = False):
        device = "gpu" if use_gpu else "cpu"

        print("正在初始化 PaddleOCR，请稍等...")

        self.ocr = PaddleOCR(
            lang=lang,
            device=device,
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=True
        )

        print("PaddleOCR 初始化完成。\n")

    def recognize(self, image: np.ndarray) -> List[Dict]:
        """
        识别图片，输入必须是 numpy 图像数组
        """
        if image is None:
            return []

        if len(image.shape) == 2:
            image = cv2.cvtColor(image, cv2.COLOR_GRAY2BGR)

        try:
            raw_result = self.ocr.predict(image)
        except Exception as e:
            print("⚠️ OCR predict 调用失败：")
            print(e)
            return []

        return self._parse_predict_result(raw_result)

    def _ocr_result_to_dict(self, item):
        """
        将 PaddleOCR 3.x OCRResult 转为 dict
        """
        if isinstance(item, dict):
            return item

        if hasattr(item, "to_dict"):
            try:
                return item.to_dict()
            except Exception:
                pass

        if hasattr(item, "json"):
            try:
                js = item.json
                return js() if callable(js) else js
            except Exception:
                pass

        if hasattr(item, "res"):
            try:
                return {"res": item.res}
            except Exception:
                pass

        if hasattr(item, "__dict__"):
            try:
                return item.__dict__
            except Exception:
                pass

        return None

    def _parse_predict_result(self, raw_result) -> List[Dict]:
        """
        正确解析 PaddleOCR 3.x predict 返回格式

        常见字段：
        rec_texts
        rec_scores
        rec_polys
        dt_polys
        """
        results = []

        if not raw_result:
            return results

        for item in raw_result:
            data = self._ocr_result_to_dict(item)

            if not data:
                continue

            if "res" in data and isinstance(data["res"], dict):
                data = data["res"]

            texts = data.get("rec_texts", [])
            scores = data.get("rec_scores", [])

            boxes = data.get("rec_polys", None)

            if boxes is None:
                boxes = data.get("dt_polys", [])

            if boxes is None:
                boxes = []

            for i, text in enumerate(texts):
                if text is None or str(text).strip() == "":
                    continue

                if i < len(scores):
                    try:
                        confidence = float(scores[i])
                    except Exception:
                        confidence = 0.0
                else:
                    confidence = 0.0

                if i < len(boxes):
                    bbox = boxes[i]

                    if hasattr(bbox, "tolist"):
                        bbox = bbox.tolist()
                else:
                    bbox = [[0, 0], [0, 0], [0, 0], [0, 0]]

                results.append({
                    "bbox": bbox,
                    "text": str(text),
                    "confidence": confidence
                })

        return results


# =========================
# 模块3：结果分析
# =========================

class ResultAnalyzer:
    """
    分析识别文本、置信度、关键字、乱码比例
    """

    def __init__(self, keywords: List[str] = None):
        self.keywords = keywords or []

    def analyze(self, ocr_results: List[Dict]) -> Dict:
        texts = [item["text"] for item in ocr_results if item.get("text")]
        full_text = " ".join(texts)

        count = len(ocr_results)

        if count > 0:
            avg_confidence = sum(item.get("confidence", 0.0) for item in ocr_results) / count
        else:
            avg_confidence = 0.0

        keyword_hits = {}

        for keyword in self.keywords:
            keyword_hits[keyword] = keyword in full_text

        all_keywords_found = all(keyword_hits.values()) if keyword_hits else True
        garbled_ratio = self._calc_garbled_ratio(full_text)

        return {
            "full_text": full_text,
            "avg_confidence": round(avg_confidence, 4),
            "keyword_hits": keyword_hits,
            "all_keywords_found": all_keywords_found,
            "garbled_ratio": round(garbled_ratio, 4),
            "is_complete": count > 0 and garbled_ratio < 0.3,
            "ocr_count": count
        }

    def _calc_garbled_ratio(self, text: str) -> float:
        if not text:
            return 1.0

        normal_pattern = re.compile(
            r'[\u4e00-\u9fffA-Za-z0-9\s\.\,\-\_\:\/\(\)\[\]\{\}，。！？、；：“”￥%℃#]+'
        )

        normal_chunks = normal_pattern.findall(text)
        normal_len = sum(len(x) for x in normal_chunks)

        return 1 - normal_len / max(len(text), 1)


# =========================
# 模块4：质量判定
# =========================

class QualityJudge:
    """
    OK / NG 判断
    """

    def __init__(self, min_confidence: float = 0.8, max_garbled: float = 0.2):
        self.min_confidence = normalize_confidence(min_confidence)
        self.max_garbled = max_garbled

    def judge(self, analysis: Dict) -> Dict:
        reasons = []

        if analysis["ocr_count"] == 0:
            reasons.append("未识别到文字")

        if analysis["avg_confidence"] < self.min_confidence:
            reasons.append(
                f"平均置信度过低：{analysis['avg_confidence']:.2f} < {self.min_confidence:.2f}"
            )

        if not analysis["all_keywords_found"]:
            missing = [
                kw for kw, hit in analysis["keyword_hits"].items()
                if not hit
            ]
            reasons.append(f"缺失关键字：{missing}")

        if analysis["garbled_ratio"] > self.max_garbled:
            reasons.append(
                f"乱码比例过高：{analysis['garbled_ratio']:.2f} > {self.max_garbled:.2f}"
            )

        status = "OK" if not reasons else "NG"

        return {
            "status": status,
            "reasons": reasons
        }


# =========================
# 模块5：报告输出
# =========================

class ReportWriter:
    """
    保存 TXT、标注图、NG 图片、Excel 报告
    """

    def __init__(self, output_dir: str):
        self.output_dir = output_dir
        self.report_dir = os.path.join(output_dir, "reports")
        self.text_dir = os.path.join(output_dir, "texts")
        self.annotated_dir = os.path.join(output_dir, "annotated_images")
        self.ng_dir = os.path.join(output_dir, "ng_images")

        for folder in [
            self.output_dir,
            self.report_dir,
            self.text_dir,
            self.annotated_dir,
            self.ng_dir
        ]:
            ensure_dir(folder)

    def save_text(self, image_name: str, text: str):
        base = os.path.splitext(safe_filename(image_name))[0]
        path = os.path.join(self.text_dir, f"{base}.txt")

        with open(path, "w", encoding="utf-8") as f:
            f.write(text)

    def save_annotated_image(
        self,
        image_name: str,
        image: np.ndarray,
        ocr_results: List[Dict],
        status: str,
        original_path: Optional[str] = None
    ):
        annotated = image.copy()

        for item in ocr_results:
            bbox = item.get("bbox", None)
            confidence = item.get("confidence", 0.0)

            if not bbox:
                continue

            try:
                points = np.array(bbox, dtype=np.int32).reshape((-1, 1, 2))

                cv2.polylines(
                    annotated,
                    [points],
                    True,
                    (0, 255, 0),
                    2
                )

                x = int(points[0][0][0])
                y = int(points[0][0][1])

                cv2.putText(
                    annotated,
                    f"{confidence:.2f}",
                    (x, max(y - 5, 18)),
                    cv2.FONT_HERSHEY_SIMPLEX,
                    0.55,
                    (0, 0, 255),
                    1
                )

            except Exception:
                continue

        safe_name = safe_filename(image_name)
        out_path = os.path.join(self.annotated_dir, safe_name)

        write_image_chinese_path(out_path, annotated)

        if status == "NG" and original_path and os.path.exists(original_path):
            try:
                ng_path = os.path.join(self.ng_dir, safe_name)
                image_copy = read_image_chinese_path(original_path)

                if image_copy is not None:
                    write_image_chinese_path(ng_path, image_copy)
                else:
                    shutil.copy2(original_path, ng_path)

            except Exception as e:
                print(f"⚠️ NG 图片保存失败：{e}")

    def generate_excel(self, records: List[Dict]) -> str:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "OCR检测报告"

        headers = [
            "序号",
            "图片名称",
            "识别文字",
            "平均置信度",
            "文字块数量",
            "关键字匹配结果",
            "OK/NG",
            "异常原因"
        ]

        sheet.append(headers)

        for index, record in enumerate(records, start=1):
            sheet.append([
                index,
                record["image_name"],
                record["full_text"],
                record["avg_confidence"],
                record["ocr_count"],
                str(record["keyword_hits"]),
                record["status"],
                "；".join(record["reasons"])
            ])

        widths = [8, 35, 90, 15, 12, 35, 12, 60]

        for i, width in enumerate(widths, start=1):
            sheet.column_dimensions[chr(64 + i)].width = width

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = os.path.join(
            self.report_dir,
            f"OCR检测报告_{timestamp}.xlsx"
        )

        workbook.save(report_path)
        return report_path


# =========================
# 交互输入
# =========================

def interactive_input():
    print("\n🔍 OCR 图像识别工具")
    print("=" * 45)
    print("请选择识别模式：")
    print("1. 单张图片")
    print("2. 批量文件夹")
    print("3. 视频抽帧")

    choice = input("请输入数字 1/2/3，默认自动判断：").strip()

    source = input("请输入图片 / 文件夹 / 视频路径：").strip().strip('"').strip("'")

    while not source:
        source = input("路径不能为空，请重新输入：").strip().strip('"').strip("'")

    source = os.path.abspath(source)

    if choice == "1":
        mode = "single"
    elif choice == "2":
        mode = "batch"
    elif choice == "3":
        mode = "video"
    else:
        mode = infer_mode(source)

    frame_interval = 30

    if mode == "video":
        value = input("请输入抽帧间隔，单位：帧，默认 30：").strip()

        if value.isdigit():
            frame_interval = int(value)

    return source, mode, frame_interval


# =========================
# 参数解析
# =========================

def parse_args():
    parser = argparse.ArgumentParser(description="OCR 图像识别一体化工具")

    parser.add_argument(
        "--source",
        type=str,
        default=None,
        help="图片 / 文件夹 / 视频路径"
    )

    parser.add_argument(
        "--mode",
        type=str,
        default=None,
        choices=["single", "batch", "video"],
        help="识别模式"
    )

    parser.add_argument(
        "--frame_interval",
        type=int,
        default=30,
        help="视频抽帧间隔，单位：帧"
    )

    parser.add_argument(
        "--lang",
        type=str,
        default="ch",
        help="OCR 语言，中文默认 ch"
    )

    parser.add_argument(
        "--use_gpu",
        action="store_true",
        help="是否使用 GPU"
    )

    parser.add_argument(
        "--keywords",
        nargs="+",
        default=[],
        help="关键字检测，例如：--keywords 生产日期 规格 品牌"
    )

    parser.add_argument(
        "--min_confidence",
        type=float,
        default=0.8,
        help="最低平均置信度，支持 0.8 或 80"
    )

    parser.add_argument(
        "--max_garbled",
        type=float,
        default=0.2,
        help="最大乱码比例"
    )

    parser.add_argument(
        "--output_dir",
        type=str,
        default=None,
        help="输出目录"
    )

    return parser.parse_args()


# =========================
# 主流程
# =========================

def main():
    args = parse_args()

    if args.source:
        source = os.path.abspath(args.source.strip().strip('"').strip("'"))
        mode = args.mode if args.mode else infer_mode(source)
        frame_interval = args.frame_interval
    else:
        source, mode, frame_interval = interactive_input()

    if not os.path.exists(source):
        print(f"❌ 输入路径不存在：{source}")
        return

    if args.output_dir:
        output_dir = args.output_dir
    else:
        if mode == "batch":
            output_dir = os.path.join(source, "ocr_output")
        else:
            parent = os.path.dirname(source) or "."
            output_dir = os.path.join(parent, "ocr_output")

    min_confidence = normalize_confidence(args.min_confidence)

    print("\n========== OCR 检测任务开始 ==========")
    print(f"输入路径：{source}")
    print(f"识别模式：{mode}")
    print(f"输出目录：{output_dir}")
    print(f"OCR语言：{args.lang}")
    print(f"是否使用GPU：{args.use_gpu}")
    print(f"关键字：{args.keywords}")
    print(f"最低置信度：{min_confidence}")
    print("=====================================\n")

    input_module = InputModule(
        source=source,
        mode=mode,
        frame_interval=frame_interval
    )

    ocr_engine = OCREngine(
        lang=args.lang,
        use_gpu=args.use_gpu
    )

    analyzer = ResultAnalyzer(
        keywords=args.keywords
    )

    judge = QualityJudge(
        min_confidence=min_confidence,
        max_garbled=args.max_garbled
    )

    writer = ReportWriter(
        output_dir=output_dir
    )

    records = []
    total = 0
    ok_count = 0
    ng_count = 0

    for image_name, image, original_path in input_module.get_data():
        total += 1

        print(f"[{total}] 正在处理：{image_name}")

        results = ocr_engine.recognize(image)
        analysis = analyzer.analyze(results)
        judgement = judge.judge(analysis)

        if judgement["status"] == "OK":
            ok_count += 1
        else:
            ng_count += 1

        record = {
            "image_name": image_name,
            "full_text": analysis["full_text"],
            "avg_confidence": analysis["avg_confidence"],
            "ocr_count": analysis["ocr_count"],
            "keyword_hits": analysis["keyword_hits"],
            "status": judgement["status"],
            "reasons": judgement["reasons"]
        }

        records.append(record)

        writer.save_text(
            image_name=image_name,
            text=analysis["full_text"]
        )

        writer.save_annotated_image(
            image_name=image_name,
            image=image,
            ocr_results=results,
            status=judgement["status"],
            original_path=original_path
        )

        print(f"    状态：{judgement['status']}")
        print(f"    平均置信度：{analysis['avg_confidence']:.2f}")
        print(f"    识别文字块数量：{analysis['ocr_count']}")

        if analysis["full_text"]:
            print(f"    识别文本：{analysis['full_text'][:100]}...")
        else:
            print("    识别文本：无")

        if judgement["reasons"]:
            print(f"    异常原因：{'；'.join(judgement['reasons'])}")

        print()

    if total == 0:
        print("⚠️ 未找到任何可处理图像。")
        return

    report_path = writer.generate_excel(records)

    print("========== OCR 检测完成 ==========")
    print(f"共处理：{total}")
    print(f"OK 数量：{ok_count}")
    print(f"NG 数量：{ng_count}")
    print(f"Excel 报告：{report_path}")
    print(f"识别文本目录：{writer.text_dir}")
    print(f"标注图片目录：{writer.annotated_dir}")
    print(f"NG 图片目录：{writer.ng_dir}")
    print("=================================")


if __name__ == "__main__":
    main()