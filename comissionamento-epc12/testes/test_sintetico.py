"""Casos de borda com respostas calculadas à mão – os mesmos 12 do teste_sintetico.py da planilha original,
agora passando pelo caminho completo do programa: planilha Excel -> importador -> banco SQLite -> cálculo."""
import pytest

from comissionamento import servico
from comissionamento.banco import Banco
from comissionamento.importacao import ErroImportacao
from testes import fabrica as f


@pytest.fixture
def banco(tmp_path):
    b = Banco(tmp_path / 'teste.db')
    servico.importar_li(b, f.li(tmp_path / 'li.xlsx'))
    yield b
    b.fechar()


def carregar(banco, tmp_path):
    # FVI: data do rodapé (28/10/2026). Os demais não têm rodapé: a data é informada, como o usuário faz na tela.
    servico.importar_fic(banco, 'FVI', f.cadastro(tmp_path / 'cad.xlsx', f.TAGS_SINTETICAS))
    d = '28/10/2026'
    base, lista = f.fvm(tmp_path, f.FVM_SINTETICO)      # FVM BASE (Hierarquia de Malhas) + Lista de FVM
    servico.importar_fvm_base(banco, base)
    servico.importar_fic(banco, 'FVM', lista, d)
    servico.importar_fic(banco, 'CCM', f.provisorio(tmp_path / 'ccm.xlsx', ['Nº CCM', 'SSOP', 'STATUS'], f.CCM_SINTETICO), d)
    servico.importar_fic(banco, 'TAP', f.provisorio(tmp_path / 'tap.xlsx', f.CABECALHO_TAP, f.TAP_SINTETICO), d)
    servico.importar_fic(banco, 'TTAS', f.provisorio(tmp_path / 'ttas.xlsx', ['Nº TTAS', 'SOP', 'STATUS'], f.TTAS_SINTETICO), d)
    return servico.calcular(banco)


def linha(lista, campo, valor):
    return next(r for r in lista if r[campo] == valor)


def valores(r, chaves):
    return tuple(r[k] for k in chaves)


FVI = ('tags', 'cii_aplic', 'cii_aprov', 'pct_cii', 'ciii_aplic', 'ciii_aprov', 'pct_ciii',
       'cv_aplic', 'cv_aprov', 'pct_cv', 'tags_sem_fvi')


def test_fvi_s1(banco, tmp_path):
    # válidas T1,T2,T3,T4,T6 = 5 | CII aplic 4 (T3 N/A) ok 2 | CIII 5/3 | CV aplic 4 ok 1 | sem FVI 1
    R = carregar(banco, tmp_path)
    assert valores(linha(R.ssop, 'ssop', f.S1), FVI) == (5, 4, 2, 0.5, 5, 3, 0.6, 4, 1, 0.25, 1)


def test_fvi_s2_espacos_e_codigo_desconhecido(banco, tmp_path):
    R = carregar(banco, tmp_path)
    assert valores(linha(R.ssop, 'ssop', f.S2), FVI) == (1, 1, 1, 1.0, 1, 1, 1.0, 1, 0, 0.0, 0)


def test_fvi_sop_p2(banco, tmp_path):
    R = carregar(banco, tmp_path)
    assert valores(linha(R.sop, 'sop', f.P2), FVI) == (2, 2, 2, 1.0, 2, 2, 1.0, 2, 1, 0.5, 0)


def test_fvm_s1_libtfunc_conta_no_ciii(banco, tmp_path):
    R = carregar(banco, tmp_path)
    assert valores(linha(R.ssop, 'ssop', f.S1), ('malhas', 'fvm_ciii', 'pct_fvm_ciii', 'fvm_cv', 'pct_fvm_cv')) == (4, 3, 0.75, 1, 0.25)


def test_ccm_cancelado_mais_revisao_concluida(banco, tmp_path):
    R = carregar(banco, tmp_path)
    assert valores(linha(R.ssop, 'ssop', f.S1), ('ccm_reg', 'ccm_aprovado', 'ccm_situacao')) == (1, 1, 'APROVADO')
    assert valores(linha(R.ssop, 'ssop', f.S2), ('ccm_reg', 'ccm_aprovado', 'ccm_situacao')) == (1, 0, 'EM ANDAMENTO')
    assert linha(R.ssop, 'ssop', f.S3)['ccm_situacao'] == 'SEM REGISTRO'


def test_tap_por_sop_pelo_numero(banco, tmp_path):
    R = carregar(banco, tmp_path)
    ch = ('tap_reg', 'tap_concluido', 'tap_situacao', 'tap_status')
    assert valores(linha(R.sop, 'sop', f.P1), ch) == (1, 1, 'CONCLUÍDO', 'TAP CONCLUIDO')
    assert valores(linha(R.sop, 'sop', f.P2), ch) == (2, 0, 'EM ANDAMENTO', 'XYZ')   # último válido do arquivo
    assert linha(R.sop, 'sop', f.P3)['tap_situacao'] == 'SEM REGISTRO'
    assert (R.total['tap'], R.total['pct_tap']) == (1, pytest.approx(1 / 3))           # SOP com TAP ÷ SOP


def test_ttas_registro_unico_que_avanca(banco, tmp_path):
    R = carregar(banco, tmp_path)
    ch = ('ttas_reg', 'ttas1_aprovado', 'ttas2_aprovado', 'ttas_situacao')
    assert valores(linha(R.sop, 'sop', f.P1), ch) == (1, 1, 1, 'TTAS2 APROVADO')
    assert valores(linha(R.sop, 'sop', f.P2), ch) == (1, 1, 0, 'TTAS1 APROVADO')


def test_verificacao(banco, tmp_path):
    R = carregar(banco, tmp_path)
    v = {t: a for _, t, _, a, _ in R.verificacao}
    assert (v['Cadastro FVI – TAGs com SSOP fora da base'], v['Cadastro FVI – status com código desconhecido'],
            v['TAP – status com código desconhecido'],
            v['Cadastro FVI – TAGs com Campo IV diferente de N/A (campo não acompanhado)'],
            v['Cadastro FVI – TAGs lidas'], v['Cadastro FVI – TAGs sem SSOP (não alocadas – não contam avanço)'],
            v['Cadastro FVI – TAGs sem FVI gerada']) == (1, 1, 1, 1, 10, 1, 1)


def test_parametro_nao_tira_tag_sem_fvi_dos_denominadores(banco, tmp_path):
    carregar(banco, tmp_path)
    banco.definir_incluir_sem_fvi(False)
    R = servico.calcular(banco)
    assert valores(linha(R.ssop, 'ssop', f.S1), FVI) == pytest.approx((5, 3, 2, 2 / 3, 4, 3, 0.75, 3, 1, 1 / 3, 1))


def test_visao_gerencial_e_total_ponderado(banco, tmp_path):
    R = carregar(banco, tmp_path)
    u1 = linha(R.visao, 'unidade', f.UN1)
    # U-9001: S1 (CII 2/4) + S2 (1/1) + S3 (1/1) = 4/6; U-9002 sem TAGs -> % vazio
    assert (u1['sop'], u1['ssop'], u1['tags'], u1['pct_cii']) == (2, 3, 7, 4 / 6)
    assert linha(R.visao, 'unidade', f.UN2)['pct_cii'] is None
    assert R.total['pct_cii'] == 4 / 6                      # ponderado, não média de %
    assert (R.total['ccm'], R.total['pct_ccm']) == (1, 1 / 4)  # 1 SSOP com CCM aprovado de 4
    assert (R.total['ttas1'], R.total['pct_ttas1']) == (2, 2 / 3)


def test_comparacao_nao_diferencia_maiusculas(banco, tmp_path):
    servico.importar_fic(banco, 'FVI', f.cadastro(tmp_path / 'c.xlsx', [('X1', f.S4.lower(), '1', 'aprovado', 'Aprovado', 'n/a', 'N/A', 'APROVADO')]))
    R = servico.calcular(banco)
    assert valores(linha(R.ssop, 'ssop', f.S4), ('tags', 'cii_aprov', 'ciii_aplic', 'cv_aprov')) == (1, 1, 0, 1)


def test_historico_gravado_na_data_do_rodape(banco, tmp_path):
    carregar(banco, tmp_path)
    h = servico.historico_com_percentuais(banco)
    assert {x['data_referencia'] for x in h} == {'2026-10-28'}
    total = linha(h, 'unidade', 'TOTAL')
    assert (total['tags'], total['pct_cii']) == (7, 4 / 6)
    # nova importação com outra data -> nova linha no histórico, a anterior fica
    servico.importar_fic(banco, 'FVI', f.cadastro(tmp_path / 'c2.xlsx', f.TAGS_SINTETICAS[:3], emissao='04/11/2026'))
    assert {x['data_referencia'] for x in servico.historico_com_percentuais(banco)} == {'2026-10-28', '2026-11-04'}


def test_criterio_editado_muda_o_resultado(banco, tmp_path):
    carregar(banco, tmp_path)
    id_ = next(i for i, e, c, _ in banco.criterios() if e == 'FVI_CII' and c == 'APROVADO')
    banco.remover_criterio(id_)
    assert linha(servico.calcular(banco).ssop, 'ssop', f.S1)['cii_aprov'] == 0
    banco.adicionar_criterio('FVI_CII', 'AAPROV', 'teste')
    assert linha(servico.calcular(banco).ssop, 'ssop', f.S1)['cii_aprov'] == 1   # T4
    banco.restaurar_criterios()
    assert linha(servico.calcular(banco).ssop, 'ssop', f.S1)['cii_aprov'] == 2


def test_fic_sem_li_recusada(tmp_path):
    b = Banco(tmp_path / 'vazio.db')
    with pytest.raises(ErroImportacao, match='SOP / SSOP'):
        servico.importar_fic(b, 'FVI', f.cadastro(tmp_path / 'c.xlsx', f.TAGS_SINTETICAS))
    assert b.dados('FVI') == []
    b.fechar()
