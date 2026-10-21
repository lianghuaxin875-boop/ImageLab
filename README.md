# ImageLab 图像检测平台

ImageLab 是一个本地图像/视频检测平台，集成图片质量检测、图像增强、视频码率分析、视频综合检测和 OCR 文字识别。

## 功能模块

- 图片质量检测：清晰度、曝光、对比度、噪声、压缩伪影与综合评分
- 图像增强：规则版自动增强、固定增强、手动增强、增强前后对比图和报告
- 视频码率分析：码率、帧率、分辨率、白平衡与达标结果
- 视频综合检测：抽帧分析清晰度、曝光、稳定性、掉帧、黑白屏、冻结和音频等指标
- OCR 文字识别：PaddleOCR 2.0/3.x 依赖环境下的图片文字识别与报告导出

## 环境安装

```bash
pip install -r requirements.txt
```

OCR 模块依赖 `paddleocr`、`paddlepaddle` 和 `paddlex[ocr]`，安装耗时较长，建议使用匹配本机环境的 Python 虚拟环境。

## 启动网页

```bash
set WEB_PORT=5001
python image_quality_web.py
```

打开：

```text
http://127.0.0.1:5001/
```

## 部署到公网

仓库已经包含 `Dockerfile` 和 `render.yaml`，可以直接连接 Render、Railway 等支持 Docker 的平台部署。部署平台会把 `PORT` 注入容器，程序监听 `0.0.0.0`，生成的地址可以从公网打开。

公网部署默认安装图片、视频和增强功能的核心依赖。OCR 的 Paddle 依赖体积较大，按需执行：

```bash
pip install -r requirements-ocr.txt
```

上传文件和检测结果保存在运行实例的临时磁盘中，适合演示和个人使用；如果要长期保存文件，需要再接对象存储。

## 目录说明

运行时生成的上传缓存、检测结果、报告、临时文件和测试素材已通过 `.gitignore` 排除，不建议提交到 GitHub。
