"""Teste do executável no Windows (CI): roda o .exe com planilhas fictícias e confere o relatório gerado.
Uso: python ci/teste_exe.py dist/ComissionamentoEPC12.exe PASTA_TEMPORARIA"""
import os
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
wb = openpyxl.load_workbook(rel)
assert len(wb['Visão Gerencial']._charts) == 5        # pizzas FVI, FVM, CCM, TAP e TTAS
dg = wb['Dados dos gráficos']
total = [dg.cell(1, 1).value, dg.cell(3, 1).value, dg.cell(3, 2).value]   # status FVI mais frequente das TAGs fictícias
assert total == ['FVI – Status FVI por TAG', 'APROVADO', 3], total
# interface: abre a janela em modo de teste (todas as páginas) e confere que ela gravou o sinal de "ok"
marca = pasta / 'interface_ok.txt'
r = subprocess.run([str(exe), '--banco', str(pasta / 'teste.db')], timeout=300,
                   env={**os.environ, 'EPC12_TESTE_INICIO': str(marca)})
assert r.returncode == 0 and marca.exists() and marca.read_text() == 'ok', f'interface não abriu (código {r.returncode})'
print('executável OK:', exe.name, f'{exe.stat().st_size / 1e6:.1f} MB', total, '· interface abriu')
