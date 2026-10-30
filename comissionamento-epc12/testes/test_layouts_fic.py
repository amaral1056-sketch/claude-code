"""Layouts das exportações da FIC usados como modelo (29/10/2026) – reproduzidos com dados fictícios:
Lista de Cadastros Básicos (SOP / SSOP), Relatório de Hierarquia de Malhas (FVM), listas de CCM, TAP e TTAS."""
import sqlite3

import pytest

from comissionamento import importacao, servico
from comissionamento.banco import Banco
from comissionamento.importacao import ErroImportacao
from testes import fabrica as f


@pytest.fixture
def banco(tmp_path):
    b = Banco(tmp_path / 'teste.db')
    yield b
    b.fechar()


def test_cadastros_basicos_unidade_e_sop_pelo_codigo(banco, tmp_path):
    msg = servico.importar_li(banco, f.cadastros_basicos(tmp_path / 'b.xlsx', f.BASICOS))
    assert msg == 'SOP / SSOP importada: 2 unidade(s), 3 SOP, 4 SSOP.'
    un, sops, ssops = banco.base()
    assert un == [f.UN1, f.UN2]
    assert [(s['ssop'], s['sop'], s['unidade'], s['descricao']) for s in ssops][:2] == [
        (f.S1, f.P1, f.UN1, 'SUB 1'), (f.S2, f.P2, f.UN1, 'SUB 2A')]
    assert [s['descricao'] for s in sops] == ['UTILIDADE 1', 'AUX 2', 'ELÉTRICA 3']
    assert banco.avisos_base() == []


def test_cadastros_basicos_avisos_vao_para_a_verificacao(banco, tmp_path):
    linhas = f.BASICOS + [('SSOP-U-9001-EL-7009-0001', 'SEM SOP NA LISTA', 'SSOP'),
                          ('SOP-U-9001-UT-1001-001', 'CÓDIGO DE SOP COMO SSOP', 'SSOP')]
    msg = servico.importar_li(banco, f.cadastros_basicos(tmp_path / 'b.xlsx', linhas))
    assert '2 aviso(s)' in msg
    _, sops, ssops = banco.base()
    assert sops[-1] == {'sop': 'SOP-U-9001-EL-7009', 'unidade': f.UN1, 'descricao': '', 'grupo': 'EL'}
    assert ssops[-1]['sop'] == f.P1                                   # SOP tirado do próprio código
    R = servico.calcular(banco)
    v = {t: (a, r) for _, t, _, a, r in R.verificacao}
    assert v['SOP / SSOP – avisos da importação (SOP fora da lista, código fora do padrão)'] == (2, 'VERIFICAR')
    assert [p[2] for p in R.pendencias if p[0] == 'SOP / SSOP'] == ['SSOP-U-9001-EL-7009-0001', 'SOP-U-9001-UT-1001-001']


def test_cadastros_basicos_com_codigo_duplicado_nao_grava(banco, tmp_path):
    with pytest.raises(ErroImportacao, match='código duplicado'):
        servico.importar_li(banco, f.cadastros_basicos(tmp_path / 'b.xlsx', f.BASICOS + [(f.S1, 'DE NOVO', 'SSOP')]))
    assert not banco.tem_base()


def test_hierarquia_de_malhas(banco, tmp_path):
    servico.importar_li(banco, f.cadastros_basicos(tmp_path / 'b.xlsx', f.BASICOS))
    p = f.hierarquia_malhas(tmp_path / 'fvm.xlsx', [
        ('M1', 'OPERAÇÃO', [('TAG', 'T1', f.S1, f.P1), ('LINHA', 'L1', f.S1, f.P1)]),
        ('M2', 'PRONTO', [('TAG', 'T2', f.S1, f.P1), ('TAG', 'T3', f.S2, f.P2)]),     # pode, mas não deve
        ('M3', 'OPERAÇÃO', []),                                                       # sem TAG/linha: sem SSOP
    ])
    with pytest.raises(ErroImportacao, match='FVM BASE'):    # a hierarquia é base, não atualização
        servico.importar_fic(banco, 'FVM', p, '29/10/2026')
    assert servico.importar_fvm_base(banco, p) == 'FVM BASE importada: 3 malhas, 5 TAG(s)/linha(s), 2 SSOP.'
    regs = banco.dados('FVM_BASE')
    assert [(r['malha'], r['tag'], r['ssop']) for r in regs] == [       # uma linha por TAG/linha da malha
        ('M1', 'T1', f.S1), ('M1', 'L1', f.S1), ('M2', 'T2', f.S1), ('M2', 'T3', f.S2), ('M3', '', '')]
    # atualização pela Lista de FVM (status em código); M1 com duas FVMs: vale a aprovada
    lista = f.lista_fvm(tmp_path / 'lista.xlsx', [('10', 'M1', 'APROVADO'), ('11', 'M2', 'CANCEL'), ('12', 'M9', 'APROVADO'),
                                                   ('13', 'M1', 'LIBEXEC')])
    _, data, _ = servico.importar_fic(banco, 'FVM', lista, '29/10/2026')
    assert data == '2026-10-29' and len(banco.dados('FVM_BASE')) == 5     # a base não foi substituída
    R = servico.calcular(banco)
    s1 = next(s for s in R.ssop if s['ssop'] == f.S1)
    assert (s1['malhas'], s1['fvm_ciii'], s1['fvm_cv']) == (1, 1, 1)   # M1 conta uma vez; M2 cancelada fica fora
    m1 = next(m for m in R.malhas if m['malha'] == 'M1')
    assert (m1['n_fvm'], m1['status']) == ('10', 'APROVADO')
    assert 'M9' not in {m['malha'] for m in R.malhas}                  # fora da FVM BASE: não vira linha
    v = {t: a for _, t, _, a, _ in R.verificacao}
    assert v['FVM – malhas em mais de uma SSOP (pode, mas não deve)'] == 1
    assert v['FVM – status com código desconhecido'] == 0
    assert v['FVM – FVMs de malhas que não estão na FVM BASE'] == 1
    assert v['FVM – malhas com mais de uma FVM na atualização'] == 1
    assert v['FVM – malhas da FVM BASE sem FVM na atualização'] == 1   # M3
    assert v['FVM – registros sem SSOP'] == 1                         # M3 (sem TAG)


def test_ccm_sem_numero_conta_pelo_ssop(banco, tmp_path):
    servico.importar_li(banco, f.cadastros_basicos(tmp_path / 'b.xlsx', f.BASICOS))
    p = f.provisorio(tmp_path / 'ccm.xlsx', ['ID Interno da FIC', 'Número CCM', 'Revisão', 'SSOP', 'SOP', 'Status'],
                     [('1', 'CCM-' + f.S1, 0, f.S1, f.P1, 'CCMCONCLUIDO'), ('2', '', 0, f.S2, '', 'CCMOCULTO'),
                      ('3', '', 0, '', '', '')])
    regs, _, _ = servico.importar_fic(banco, 'CCM', p, '29/10/2026')
    assert [(r['numero'], r['ssop'], r['status']) for r in regs] == [('CCM-' + f.S1, f.S1, 'CCMCONCLUIDO'),
                                                                     ('', f.S2, 'CCMOCULTO')]
    R = servico.calcular(banco)
    assert [s['ccm_status'] for s in R.ssop] == ['CCMCONCLUIDO', 'CCMOCULTO', '', '']


def test_tap_e_ttas_por_sop(banco, tmp_path):
    servico.importar_li(banco, f.cadastros_basicos(tmp_path / 'b.xlsx', f.BASICOS))
    tap = f.provisorio(tmp_path / 'tap.xlsx', ['Id FIC', 'Número do TAP', 'Descrição', 'Documento de Referência',
                                               'TAP em duas fases', 'Status', 'Data Últ. Status', 'Empreendimento'],
                       [('1', 'TAP-' + f.P1, 'UTILIDADE 1', 'PR-1', 'N', 'TAP CONCLUIDO', '', 'X')])
    regs, _, _ = servico.importar_fic(banco, 'TAP', tap, '29/10/2026')
    assert regs[0]['sop'] == f.P1                                     # "TAP-" + SOP
    ttas = f.provisorio(tmp_path / 'ttas.xlsx', ['Id FIC', 'Número TTAS', 'SOP/SSOP', 'Status', 'Data Últ. Status',
                                                 'TTAS por', 'Empreendimento'],
                        [('1', 'TTAS-' + f.P1, f.P1, 'TTASCONCLUIDO', '', 'SOP', 'X')])
    servico.importar_fic(banco, 'TTAS', ttas, '29/10/2026')
    R = servico.calcular(banco)
    p1 = next(s for s in R.sop if s['sop'] == f.P1)
    assert (p1['tap_status'], p1['ttas2_status']) == ('TAP CONCLUIDO', 'TTASCONCLUIDO')
    assert (R.total['pct_tap'], R.total['pct_ttas2']) == (pytest.approx(1 / 3), pytest.approx(1 / 3))


def test_banco_antigo_e_migrado(tmp_path):
    caminho = tmp_path / 'antigo.db'
    b = Banco(caminho)
    b.fechar()
    con = sqlite3.connect(caminho)
    con.execute('DROP TABLE tap')
    con.execute('CREATE TABLE tap (linha INTEGER PRIMARY KEY, numero TEXT, ssop TEXT, status TEXT)')
    con.execute('CREATE TABLE li_item (linha INTEGER PRIMARY KEY, tag TEXT)')
    con.execute('DROP TABLE fvm')
    con.execute('CREATE TABLE fvm (linha INTEGER PRIMARY KEY, malha TEXT, sop TEXT, ssop TEXT, status TEXT)')
    con.commit()
    con.close()
    b = Banco(caminho)
    assert [r[1] for r in b.con.execute('PRAGMA table_info(tap)')] == ['linha', 'numero', 'sop', 'status']
    assert b.con.execute("SELECT COUNT(*) FROM sqlite_master WHERE name = 'li_item'").fetchone()[0] == 0
    assert [r[1] for r in b.con.execute('PRAGMA table_info(fvm)')] == ['linha', 'malha', 'tag', 'sop', 'ssop', 'status_item']
    assert b.meta('versao_esquema') == '3'
    b.fechar()


def test_li_antiga_continua_aceita(tmp_path):
    linhas, sops, avisos = importacao.ler_sop_ssop(f.li(tmp_path / 'li.xlsx'))
    assert (len(linhas), sops, avisos) == (4, None, [])


def test_templates_de_fvm_do_pdf(banco, tmp_path):
    pdf = f.templates_fvm_pdf(tmp_path / 'tpl.pdf', [
        ('FVM-INS001', '2.0', 'INSTRUMENTAÇÃO', [('III.1', 'LINHAS: Malha montada.'), ('III.2', 'TAGs: Instrumento instalado.'),
                                                  ('V.7', 'TAGs: Teste de malha aprovado.'), ('V.8', 'TAGs: Alarme testado.')]),
        ('FVM-ELE001', '3.0', 'ELÉTRICA', [('III.1', 'LINHAS: Cabo lançado.'), ('V.9', 'TAGs: Blank test aprovado.')])])
    doc, tpls = importacao.ler_templates(pdf)
    assert doc == 'FVM' and [t['template'] for t in tpls] == ['FVM-INS001', 'FVM-ELE001']
    assert (tpls[0]['revisao'], tpls[0]['disciplina']) == ('2.0', 'INSTRUMENTAÇÃO')
    assert tpls[0]['perguntas'] == [('III/IV', 'III.1', 'LINHAS: Malha montada.'), ('III/IV', 'III.2', 'TAGs: Instrumento instalado.'),
                                    ('V', 'V.7', 'TAGs: Teste de malha aprovado.'), ('V', 'V.8', 'TAGs: Alarme testado.')]  # V.7 colado
    assert servico.importar_templates(banco, pdf) == 'Templates de FVM importados: 2 template(s), 6 pergunta(s).'
    assert banco.templates('FVM')['FVM-ELE001']['perguntas'] == [('III/IV', 'III.1', 'LINHAS: Cabo lançado.'),
                                                                 ('V', 'V.9', 'TAGs: Blank test aprovado.')]
    # a malha leva o template da Lista de FVM; template fora do PDF vai para a Verificação
    servico.importar_li(banco, f.cadastros_basicos(tmp_path / 'b.xlsx', f.BASICOS))
    base, _ = f.fvm(tmp_path, [('M1', f.P1, f.S1, 'APROVADO'), ('M2', f.P1, f.S1, 'APROVADO')])
    servico.importar_fvm_base(banco, base)
    lista = f.provisorio(tmp_path / 'l.xlsx', ['Código FVM', 'Malha', 'Código do Template', 'Status'],
                         [('1', 'M1', 'FVM-INS001', 'APROVADO'), ('2', 'M2', 'FVM-XYZ001', 'APROVADO')])
    servico.importar_fic(banco, 'FVM', lista, '30/10/2026')
    R = servico.calcular(banco)
    assert {m['malha']: m['template'] for m in R.malhas} == {'M1': 'FVM-INS001', 'M2': 'FVM-XYZ001'}
    v = {t: a for _, t, _, a, _ in R.verificacao}
    assert v['FVM – templates usados nas FVMs que não estão no PDF de templates importado'] == 1


def test_pdf_que_nao_e_template_de_fvm_e_recusado(tmp_path):
    from fpdf import FPDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font('Helvetica', size=9)
    pdf.cell(0, 5, 'Outro documento qualquer')
    pdf.output(str(tmp_path / 'x.pdf'))
    with pytest.raises(ErroImportacao, match='Template de FVM'):
        importacao.ler_templates(tmp_path / 'x.pdf')
