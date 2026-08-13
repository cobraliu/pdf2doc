# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置 —— 产出单文件可执行程序 (Windows: pdfrec.exe)

构建:  pyinstaller pdfrec.spec --noconfirm
要点:
  1. RapidOCR 的 3 个 ONNX 模型随包分发, 运行时零联网 (convert.py 的 model_root() 会找 models/)
  2. rapidocr 的 yaml 配置必须一起打包, 否则启动即崩
  3. 排除 torch/paddle 等可选后端, 否则体积翻数倍
"""
import os, glob
from PyInstaller.utils.hooks import collect_submodules

import rapidocr
ROD = os.path.dirname(rapidocr.__file__)

datas = [
    # 模型: 放到解包根目录的 models/ 下, 与 convert.py 的 model_root() 对应
    *[(f, 'models') for f in glob.glob(os.path.join(ROD, 'models', '*.onnx'))],
    # rapidocr 自身的配置文件(缺一不可)
    (os.path.join(ROD, 'config.yaml'), 'rapidocr'),
    (os.path.join(ROD, 'default_models.yaml'), 'rapidocr'),
    (os.path.join(ROD, 'inference_engine', 'pytorch', 'networks', 'arch_config.yaml'),
     os.path.join('rapidocr', 'inference_engine', 'pytorch', 'networks')),
]

hiddenimports = [
    *collect_submodules('rapidocr'),
    'onnxruntime', 'onnxruntime.capi', 'onnxruntime.capi._pybind_state',
    'PIL.Image', 'docx', 'openpyxl', 'pymupdf', 'yaml', 'cv2', 'shapely',
]

excludes = [
    'torch', 'torchvision', 'paddle', 'paddlepaddle', 'tensorflow',
    'matplotlib', 'scipy', 'pandas', 'IPython', 'jupyter', 'notebook',
    'tkinter', 'PyQt5', 'PySide2', 'pytest', 'sympy', 'numba',
]

a = Analysis(
    ['convert.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name='pdfrec',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,             # UPX 会让部分杀软误报, 且对 onnx 模型无效
    runtime_tmpdir=None,
    console=True,          # 需要控制台显示进度
    disable_windowed_traceback=False,
    argv_emulation=False,  # macOS 拖拽支持; Windows 无影响
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
