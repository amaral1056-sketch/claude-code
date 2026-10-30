"""Validação com os dados reais (rodar localmente): o programa precisa chegar aos MESMOS números da planilha
CONTROLE_GERENCIAL_EPC12.xlsx entregue em 28/10/2026, calculada pelas fórmulas do Excel/LibreOffice.

Os arquivos ficam em comissionamento-epc12/dados/ (fora do git – informação classificada NP-1):
  LI-*.xlsx (aba LISTA) · FIC_Cadastro_Geral.xlsx · CONTROLE_GERENCIAL_EPC12.xlsx · status_fic/STATUS_*.xlsx
Sem eles, os testes são pulados."""
from pathlib import Path

import openpyxl
import pytest

from comissionamento import servico
from comissionamento.banco import Banco

DADOS = Path(__file__).resolve().parents[1] / 'dados'
LI = next(iter(sorted(DADOS.glob('LI-*.xlsx'))), DADOS / 'LI.xlsx')
CAD = DADOS / 'FIC_Cadastro_Geral.xlsx'
PLANILHA = DADOS / 'CONTROLE_GERENCIAL_EPC12.xlsx'
pytestmark = pytest.mark.skipif(not (LI.exists() and CAD.exists() and PLANILHA.exists()),
                                reason='dados reais ausentes em comissionamento-epc12/dados/')


def igual(a, b):
    if a in (None, '') and b in (None, ''):
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) < 1e-9
    return a == b


@pytest.fixture(scope='module')
def ref():
    return openpyxl.load_workbook(PLANILHA, data_only=True, read_only=True)


@pytest.fixture(scope='module')
def calculado(tmp_path_factory, ref):
    assert ref['CRITERIOS']['F2'].value == 'SIM'   # parâmetro com que a planilha de referência foi calculada
    b = Banco(tmp_path_factory.mktemp('db') / 'teste.db')
    servico.importar_li(b, LI)
    for f in sorted((DADOS / 'status_fic').glob('STATUS_*.xlsx')):
        servico.importar_status(b, f)
    _, data, _ = servico.importar_fic(b, 'FVI', CAD)
    R = servico.calcular(b)
    yield b, R, data
    b.fechar()


def test_parametro_nao_bate_com_historico(tmp_path):
    # docs/HISTORICO da planilha: Campo II 63,0% com SIM e 79,0% com NÃO (28/10/2026)
    b = Banco(tmp_path / 'nao.db')
    servico.importar_li(b, LI)
    servico.importar_fic(b, 'FVI', CAD)
    assert round(servico.calcular(b).total['pct_cii'] * 100, 1) == 63.0
    b.definir_incluir_sem_fvi(False)
    assert round(servico.calcular(b).total['pct_cii'] * 100, 1) == 79.0
    b.fechar()


def test_data_do_rodape(calculado):
    assert calculado[2] == '2026-10-28'


def test_fvi_por_ssop_igual_a_planilha(calculado, ref):
    _, R, _ = calculado
    chaves = ['tags', 'cii_aplic', 'cii_aprov', 'pct_cii', 'ciii_aplic', 'ciii_aprov', 'pct_ciii',
              'cv_aplic', 'cv_aprov', 'pct_cv', 'tags_sem_fvi']
    planilha = {r[2]: r for r in ref['FVI'].iter_rows(min_row=2, max_col=15, values_only=True) if r[2]}
    assert len(planilha) == len(R.ssop)
    dif = [(s['ssop'], k) for s in R.ssop for j, k in enumerate(chaves, 4) if not igual(s[k], planilha[s['ssop']][j])]
    assert dif == []


def test_fvi_por_sop_igual_a_planilha(calculado, ref):
    _, R, _ = calculado
    chaves = ['qtd_ssop', 'tags', 'cii_aplic', 'cii_aprov', 'pct_cii', 'ciii_aplic', 'ciii_aprov', 'pct_ciii',
              'cv_aplic', 'cv_aprov', 'pct_cv', 'tags_sem_fvi']
    planilha = {r[1]: r for r in ref['FVI'].iter_rows(min_row=2, min_col=17, max_col=31, values_only=True) if r[1]}
    assert len(planilha) == len(R.sop)
    dif = [(s['sop'], k) for s in R.sop for j, k in enumerate(chaves, 3) if not igual(s[k], planilha[s['sop']][j])]
    assert dif == []


def test_visao_gerencial_igual_a_planilha(calculado, ref):
    _, R, _ = calculado
    chaves = ['unidade', 'sop', 'ssop', 'tags', 'pct_cii', 'pct_ciii', 'pct_cv', 'malhas', 'pct_fvm_ciii', 'pct_fvm_cv',
              'ccm', 'pct_ccm', 'tap', 'pct_tap', 'ttas1', 'pct_ttas1', 'ttas2', 'pct_ttas2', 'tags_sem_fvi']
    linhas = [r for r in ref['VISAO_GERENCIAL'].iter_rows(min_row=5, max_row=17, max_col=19, values_only=True)]
    nossas = R.visao + [R.total]
    assert [r[0] for r in linhas] == [v['unidade'] for v in nossas]
    dif = [(v['unidade'], k, v[k], r[j]) for v, r in zip(nossas, linhas) for j, k in enumerate(chaves) if not igual(v[k], r[j])]
    assert dif == []


def test_verificacao_igual_a_planilha(calculado, ref):
    _, R, _ = calculado
    ver = {r[0]: r[2] for r in ref['VERIFICACAO'].iter_rows(values_only=True) if r[0] and r[3] in ('OK', 'VERIFICAR', 'INFO')}
    nossa = {t: a for _, t, _, a, _ in R.verificacao}
    pares = {'IMP_FVI – linhas com SSOP fora da base': 'Cadastro FVI – TAGs com SSOP fora da base',
             'IMP_FVI – status com código desconhecido': 'Cadastro FVI – status com código desconhecido',
             'IMP_FVI – TAGs com Campo IV diferente de N/A (campo não acompanhado)':
                 'Cadastro FVI – TAGs com Campo IV diferente de N/A (campo não acompanhado)',
             'IMP_FVI – TAGs lidas': 'Cadastro FVI – TAGs lidas',
             'IMP_FVI – TAGs sem SSOP (não alocadas – não contam avanço)':
                 'Cadastro FVI – TAGs sem SSOP (não alocadas – não contam avanço)',
             'IMP_FVI – TAGs sem FVI gerada': 'Cadastro FVI – TAGs sem FVI gerada'}
    assert {k: ver[k] for k in pares} == {k: nossa[v] for k, v in pares.items()}
