# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：单文件、无控制台、内嵌 logo 图标。

构建：  pyinstaller --noconfirm md2word.spec
产物：  dist/Markdown转Word批量转换器.exe
"""

# 运行期完全用不到的大体积第三方库，剔除后能显著压缩产物体积。
EXCLUDES = [
    # 科学计算 / 机器学习
    'numpy', 'scipy', 'pandas', 'matplotlib', 'geopandas',
    'sklearn', 'scikit-learn', 'skimage', 'sympy', 'numba',
    'torch', 'torchvision', 'torchmetrics', 'pytorch_lightning', 'lightning',
    'tensorflow', 'tensorflow_intel', 'tensorflow_estimator', 'keras',
    'cv2', 'nltk', 'h5py', 'pyarrow',
    # 其它 GUI 工具包（本项目只用 wxPython）
    'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'tkinter', '_tkinter',
    # 交互式计算环境
    'IPython', 'ipython', 'jupyter', 'notebook', 'jupyterlab',
    'nbformat', 'nbconvert', 'qtconsole', 'traitlets', 'zmq',
    # 图像处理（仅打包阶段用于生成 .ico，运行时不需要）
    'PIL', 'Pillow',
    # 测试 / 文档工具
    'pytest', 'nose', 'sphinx', 'docutils',
]

a = Analysis(
    ['md2word.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('logo/logo.ico', 'logo'),
        ('logo/logo.png', 'logo'),
    ],
    hiddenimports=['wx.dataview'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Markdown转Word批量转换器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='logo/logo.ico',
)
