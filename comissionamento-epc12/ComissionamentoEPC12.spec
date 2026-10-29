# PyInstaller: python -m PyInstaller --noconfirm --clean ComissionamentoEPC12.spec
# Gera um executável único, sem console, que não precisa de instalação (dist/ComissionamentoEPC12[.exe]).
a = Analysis(['main.py'], pathex=['.'], hiddenimports=[], excludes=['pytest', 'numpy', 'pandas', 'PIL'])
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name='ComissionamentoEPC12', console=False, upx=False,
          runtime_tmpdir=None)
