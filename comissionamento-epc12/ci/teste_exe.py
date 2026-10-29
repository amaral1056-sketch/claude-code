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
li = f.cadastros_basicos(pasta / 'sop_ssop.xlsx', f.BASICOS)          # layouts da FIC (Lista de Cadastros Básicos)
cad = f.cadastro(pasta / 'cad.xlsx', f.TAGS_SINTETICAS)
fvm_base, fvm = f.fvm(pasta, [('M1', f.P1, f.S1, 'APROVADO')])       # FVM BASE + Lista de FVM
rel = pasta / 'relatorio.xlsx'
r = subprocess.run([str(exe), '--banco', str(pasta / 'teste.db'), '--li', str(li), '--importar', 'FVI', str(cad),
                    '--fvm-base', str(fvm_base), '--importar', 'FVM', str(fvm), '--exportar', str(rel)], timeout=300)
assert r.returncode == 0, f'executável terminou com código {r.returncode}'
wb = openpyxl.load_workbook(rel)
assert wb.sheetnames == ['Visão Gerencial', 'FVI', 'FVM', 'CCM', 'TAP', 'TTAS'], wb.sheetnames
vg = wb['Visão Gerencial']
assert len(vg._charts) == 5 and len(vg._images) == 1   # pizzas FVI, FVM, CCM, TAP, TTAS e a logo
total = [vg['A60'].value, vg['A62'].value, vg['B62'].value]   # status FVI mais frequente das TAGs fictícias
assert total == ['FVI – Status FVI por TAG', 'APROVADO', 3], total
assert [c.value for c in wb['FVM'][5]][3:7] == ['M1', 'T-M1-1', '100', 'APROVADO']   # malha da base com a FVM
pdf = rel.with_suffix('.pdf')                                  # PDF gerado junto, com as fontes embutidas
assert pdf.exists() and pdf.read_bytes().startswith(b'%PDF') and b'EPC12Display' in pdf.read_bytes()
# interface: abre a janela em modo de teste (todas as páginas) e confere que ela gravou o sinal de "ok"
marca = pasta / 'interface_ok.txt'
r = subprocess.run([str(exe), '--banco', str(pasta / 'teste.db')], timeout=300,
                   env={**os.environ, 'EPC12_TESTE_INICIO': str(marca)})
assert r.returncode == 0 and marca.exists() and marca.read_text() == 'ok', f'interface não abriu (código {r.returncode})'
print('executável OK:', exe.name, f'{exe.stat().st_size / 1e6:.1f} MB', total, '· interface abriu')
