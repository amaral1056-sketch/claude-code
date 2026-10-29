"""Teste do executável no Windows (CI): roda o .exe com planilhas fictícias e confere o relatório gerado.
Uso: python ci/teste_exe.py dist/ComissionamentoEPC12.exe PASTA_TEMPORARIA"""
import subprocess
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from testes import fabrica as f  # noqa: E402

exe, pasta = Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve()
pasta.mkdir(parents=True, exist_ok=True)
li = f.li(pasta / 'li.xlsx')
cad = f.cadastro(pasta / 'cad.xlsx', f.TAGS_SINTETICAS)
rel = pasta / 'relatorio.xlsx'
r = subprocess.run([str(exe), '--banco', str(pasta / 'teste.db'), '--li', str(li), '--importar', 'FVI', str(cad),
                    '--exportar', str(rel)], timeout=300)
assert r.returncode == 0, f'executável terminou com código {r.returncode}'
vg = openpyxl.load_workbook(rel)['VISAO_GERENCIAL']
total = [vg.cell(7, c).value for c in (1, 4, 5)]
assert total[0] == 'TOTAL' and total[1] == 7 and abs(total[2] - 4 / 6) < 1e-9, total
print('executável OK:', exe.name, f'{exe.stat().st_size / 1e6:.1f} MB', total)
