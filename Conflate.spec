# Conflate v1.1 spec - sklearn pinned to 1.5.2, build from plain CMD not Anaconda Prompt
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

datas  = []
datas += collect_data_files("customtkinter")
datas += collect_data_files("sklearn")
datas += [
    ("conflate_icon.ico", "."),
    ("conflate_icon.png", "."),
]

hiddenimports = [
    "sklearn.feature_extraction.text",
    "sklearn.feature_extraction._hashing_fast",
    "sklearn.utils._cython_blas",
    "sklearn.utils._typedefs",
    "sklearn.utils._heap",
    "sklearn.utils._sorting",
    "sklearn.utils._vector_sentinel",
    "sklearn.neighbors._partition_nodes",
    "scipy.sparse",
    "scipy.sparse.csr",
    "scipy.sparse.coo",
    "scipy.sparse.csgraph._validation",
    "scipy.sparse.linalg",
    "scipy.sparse._compressed",
    "scipy.sparse._index",
    "numpy",
    "openpyxl",
    "PIL", "PIL.Image", "PIL.ImageDraw",
    "rapidfuzz", "rapidfuzz.fuzz", "rapidfuzz.process",
    "packaging", "packaging.version",
    "packaging.specifiers", "packaging.requirements",
]

a = Analysis(
    ["Conflate.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "PyQt5", "PyQt6", "PySide2", "PySide6",
        "matplotlib",
        "torch", "torchvision", "torchaudio",
        "tensorflow", "keras",
        "cv2",
        "bokeh", "plotly", "altair", "seaborn",
        "statsmodels", "sympy",
        "pyarrow", "fastparquet",
        "dask", "distributed",
        "numba", "llvmlite",
        "skimage", "astropy", "h5py", "tables",
        "xarray", "panel", "intake",
        "streamlit", "sphinx", "sphinxcontrib",
        "notebook", "ipython", "IPython",
        "ipykernel", "ipywidgets", "jupyter",
        "nbconvert", "nbformat",
        "pytest",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="Conflate",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="conflate_icon.ico",
    version="version_info.txt",
)
