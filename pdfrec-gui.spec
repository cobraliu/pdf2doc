# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置 —— 图形界面版, 三平台通用

构建 (在目标平台各跑一次, PyInstaller 不支持交叉编译):
    Windows :  pyinstaller pdfrec-gui.spec --noconfirm   ->  dist/扫描件转Word.exe
    macOS   :  pyinstaller pdfrec-gui.spec --noconfirm   ->  dist/扫描件转Word.app
    Linux   :  pyinstaller pdfrec-gui.spec --noconfirm   ->  dist/扫描件转Word

要点:
  1. 3 个 ONNX 模型随包分发, 运行时零联网
  2. PySide6 默认会拖进 QtWebEngine 等巨型模块, 必须显式排除, 否则体积翻 3 倍
  3. Windows 下 console=False, 避免弹出黑色命令行窗口
"""
import glob, os, sys

import rapidocr
ROD = os.path.dirname(rapidocr.__file__)

APP_NAME = '扫描件转Word'

datas = [
    *[(f, 'models') for f in glob.glob(os.path.join(ROD, 'models', '*.onnx'))],
    (os.path.join(ROD, 'config.yaml'), 'rapidocr'),
    (os.path.join(ROD, 'default_models.yaml'), 'rapidocr'),
    (os.path.join(ROD, 'inference_engine', 'pytorch', 'networks', 'arch_config.yaml'),
     os.path.join('rapidocr', 'inference_engine', 'pytorch', 'networks')),
]

from PyInstaller.utils.hooks import collect_submodules
hiddenimports = [
    *collect_submodules('rapidocr'),
    'onnxruntime', 'onnxruntime.capi', 'onnxruntime.capi._pybind_state',
    'PIL.Image', 'docx', 'openpyxl', 'pymupdf', 'yaml', 'cv2', 'shapely', 'convert',
]

# PySide6 里没用到的一律排除 —— 这是体积能否控制住的关键
excludes = [
    'torch', 'torchvision', 'paddle', 'paddlepaddle', 'tensorflow',
    'matplotlib', 'scipy', 'pandas', 'IPython', 'jupyter', 'notebook',
    'tkinter', 'PyQt5', 'PyQt6', 'pytest', 'sympy', 'numba',
    # Qt 巨型模块(只用 QtCore/QtGui/QtWidgets)
    'PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets', 'PySide6.QtWebEngineQuick',
    'PySide6.QtQuick', 'PySide6.QtQuick3D', 'PySide6.QtQml', 'PySide6.QtQuickWidgets',
    'PySide6.Qt3DCore', 'PySide6.Qt3DRender', 'PySide6.Qt3DExtras', 'PySide6.Qt3DAnimation',
    'PySide6.QtCharts', 'PySide6.QtDataVisualization', 'PySide6.QtMultimedia',
    'PySide6.QtMultimediaWidgets', 'PySide6.QtBluetooth', 'PySide6.QtNfc',
    'PySide6.QtPositioning', 'PySide6.QtLocation', 'PySide6.QtSerialPort',
    'PySide6.QtSql', 'PySide6.QtTest', 'PySide6.QtDesigner', 'PySide6.QtHelp',
    'PySide6.QtPdf', 'PySide6.QtPdfWidgets', 'PySide6.QtSpatialAudio',
    'PySide6.QtRemoteObjects', 'PySide6.QtScxml', 'PySide6.QtSensors',
    'PySide6.QtStateMachine', 'PySide6.QtTextToSpeech', 'PySide6.QtWebChannel',
    'PySide6.QtWebSockets', 'PySide6.QtOpenGL', 'PySide6.QtOpenGLWidgets',
    'shiboken6.support', 'PySide6.scripts',
]

a = Analysis(
    ['gui.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
)
pyz = PYZ(a.pure)

if sys.platform == 'darwin':
    # macOS: 必须 onedir + BUNDLE (onefile 与 .app 冲突, PyInstaller 7 起会报错)
    exe = EXE(
        pyz, a.scripts, [], exclude_binaries=True,
        name=APP_NAME, debug=False, strip=False, upx=False, console=False,
        argv_emulation=True,        # 支持把 PDF 拖到 Dock 图标上
    )
    coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name=APP_NAME)
    app = BUNDLE(
        coll,
        name=f'{APP_NAME}.app',
        icon=None,
        bundle_identifier='com.local.pdfrec',
        info_plist={
            'CFBundleDisplayName': APP_NAME,
            'NSHighResolutionCapable': True,
            'LSMinimumSystemVersion': '11.0',
            'CFBundleDocumentTypes': [{
                'CFBundleTypeName': 'PDF',
                'CFBundleTypeExtensions': ['pdf'],
                'CFBundleTypeRole': 'Viewer',
            }],
        },
    )
else:
    # Windows / Linux: 单文件, 双击即用
    exe = EXE(
        pyz, a.scripts, a.binaries, a.datas, [],
        name=APP_NAME, debug=False, bootloader_ignore_signals=False,
        strip=False, upx=False, runtime_tmpdir=None,
        console=False,              # 不弹命令行黑框
        disable_windowed_traceback=False,
        target_arch=None, codesign_identity=None, entitlements_file=None,
    )
