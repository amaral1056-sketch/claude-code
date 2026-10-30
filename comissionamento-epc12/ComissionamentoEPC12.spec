# PyInstaller: python -m PyInstaller --noconfirm --clean ComissionamentoEPC12.spec
# Gera um executável único, sem console, que não precisa de instalação (dist/ComissionamentoEPC12[.exe]).
from PyInstaller.utils.hooks import collect_data_files

a = Analysis(['main.py'], pathex=['.'],
             datas=collect_data_files('customtkinter') + [('comissionamento/recursos', 'comissionamento/recursos')],
             hiddenimports=['PIL._tkinter_finder', 'PIL.ImageTk'], excludes=['pytest', 'numpy', 'pandas'])
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name='ComissionamentoEPC12', console=False, upx=False,
          runtime_tmpdir=None, icon='comissionamento/recursos/icone.ico')
