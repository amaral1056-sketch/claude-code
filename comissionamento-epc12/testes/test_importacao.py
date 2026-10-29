"""Importadores: layouts aceitos, recusas (nada é gravado) e exportações."""
import datetime

import openpyxl
import pytest

from comissionamento import exportacao, importacao, servico, tema
from comissionamento.banco import Banco
from comissionamento.importacao import ErroImportacao
from testes import fabrica as f


@pytest.fixture
def banco(tmp_path):
    b = Banco(tmp_path / 'teste.db')
    servico.importar_li(b, f.li(tmp_path / 'li.xlsx'))
    yield b
    b.fechar()


def test_li_base_e_documentos(banco):
    un, sops, ssops = banco.base()
    assert un == [f.UN1, f.UN2]
    assert [s['sop'] for s in sops] == [f.P1, f.P2, f.P3]
    assert ssops[0]['grupo'] == 'UT'
    assert ssops[0]['documentos'] == 'DE-0000.00-0000-000-XXX-001 | DE-0000.00-0000-000-XXX-002'


@pytest.mark.parametrize('linhas, msg', [
    ([(f.UN1, f.P1, 'A', f.S1, 'X'), (f.UN1, f.P1, 'A', f.S1, 'Y')], 'SSOP duplicado'),
    ([(f.UN1, f.P1, 'A', f.S2, 'X')], 'não começa com'),
    ([(f.UN2, f.P1, 'A', f.S1, 'X')], 'não confere'),
])
def test_li_com_problema_nao_grava(banco, tmp_path, linhas, msg):
    with pytest.raises(ErroImportacao, match=msg):
        servico.importar_li(banco, f.li(tmp_path / 'ruim.xlsx', linhas))
    assert len(banco.base()[2]) == 4   # base anterior preservada


def test_li_sem_aba_lista(tmp_path):
    p = tmp_path / 'x.xlsx'
    openpyxl.Workbook().save(p)
    with pytest.raises(ErroImportacao, match='LISTA'):
        importacao.ler_li(p)


def test_cadastro_layout_mudou_para(banco, tmp_path):
    p = f.cadastro(tmp_path / 'c.xlsx', f.TAGS_SINTETICAS)
    wb = openpyxl.load_workbook(p)
    wb.active.cell(5, 37, 'CAMPO II')          # coluna AK renomeada
    wb.save(p)
    with pytest.raises(ErroImportacao, match='AK'):
        servico.importar_fic(banco, 'FVI', p)
    assert banco.dados('FVI') == []


def test_cadastro_data_e_regra_de_linha(tmp_path):
    regs, data, info = importacao.ler_cadastro(f.cadastro(tmp_path / 'c.xlsx', f.TAGS_SINTETICAS))
    assert data == datetime.date(2026, 10, 28)
    assert [r['tag'] for r in regs] == [f'T{i}' for i in range(1, 11)]   # sem cabeçalho repetido e sem rodapé
    assert regs[0]['dados']['STATUS CAMPO IV'] == 'ACAMPO3E4'


def test_cadastro_sem_rodape_usa_data_informada(banco, tmp_path):
    p = f.cadastro(tmp_path / 'c.xlsx', f.TAGS_SINTETICAS, emissao=None)
    _, data, _ = servico.importar_fic(banco, 'FVI', p, datetime.date(2026, 11, 1))
    assert data == '2026-11-01'


def test_provisorio_cabecalho_por_nome_em_outra_ordem_e_com_titulo(banco, tmp_path):
    p = f.provisorio(tmp_path / 'fvm.xlsx', ['STATUS DA FVM', 'Nº MALHA', 'SSOP', 'OBS'],
                     [('APROVADO', 'M1', f.S1, ''), ('ACAMPO5', 'M2', f.S1, ''), ('', '', '', ''),
                      ('Emissão: ', '', '', ''), ('30/10/2026', '', '', '')], titulo='Relatório de FVM')
    regs, data, info = servico.importar_fic(banco, 'FVM', p)
    assert [r['malha'] for r in regs] == ['M1', 'M2']
    assert data == '2026-10-30'


def test_provisorio_sem_cabecalho_reconhecido(banco, tmp_path):
    p = f.provisorio(tmp_path / 'x.xlsx', ['A', 'B', 'C'], [('1', '2', '3')])
    with pytest.raises(ErroImportacao, match='Cabeçalho de CCM'):
        servico.importar_fic(banco, 'CCM', p)


def test_modelos_exportados_sao_reconhecidos(banco, tmp_path):
    p = tmp_path / 'modelos.xlsx'
    exportacao.exportar_modelos(p)
    wb = openpyxl.load_workbook(p)
    wb['CCM'].append(['C1', f.S1, 'APROVADO'])
    wb['TTAS'].append(['T1', f.P1, 'TTAS1APROVADO'])
    wb.save(p)
    servico.importar_fic(banco, 'CCM', p)
    servico.importar_fic(banco, 'TTAS', p)
    R = servico.calcular(banco)
    assert R.total['ccm'] == 1 and R.total['ttas1'] == 1


def test_status_oficial_pelo_nome_do_arquivo(banco, tmp_path):
    p = f.provisorio(tmp_path / 'STATUS_TAP.xlsx', ['Valor', 'Descrição'], [('NOVO', 'Novo código'), ('TAP CONCLUIDO', 'x')])
    assert servico.importar_status(banco, p) == ('TAP', 2)
    assert [c for d, c, _ in banco.status_fic() if d == 'TAP'] == ['NOVO', 'TAP CONCLUIDO']


def test_relatorio_excel_e_backup(banco, tmp_path):
    servico.importar_fic(banco, 'FVI', f.cadastro(tmp_path / 'c.xlsx', f.TAGS_SINTETICAS))
    R = servico.calcular(banco)
    destino = tmp_path / 'rel.xlsx'
    exportacao.exportar_relatorio(R, destino, '28/10/2026', servico.historico_com_percentuais(banco), banco.criterios(),
                                  True, banco.status_fic())
    wb = openpyxl.load_workbook(destino)
    assert wb.sheetnames == ['Visão Gerencial', 'FVI', 'FVM', 'CCM', 'TAP', 'TTAS', 'Verificação', 'Pendências',
                             'Histórico', 'Critérios', 'Dados dos gráficos']
    vg = wb['Visão Gerencial']
    assert len(vg._charts) == 5                                  # pizzas FVI, FVM, CCM, TAP e TTAS
    assert all(type(g).__name__ == 'PieChart' for g in vg._charts)
    # mesma cor da tela para o mesmo status (APROVADO é a fatia maior do FVI)
    ordem_fvi = [c for d, c, _ in banco.status_fic() if d == 'FVI']
    cor_tela = tema.cores_pizza(['APROVADO'], ordem_fvi)[0].lstrip('#').upper()
    assert vg._charts[0].series[0].dPt[0].graphicalProperties.solidFill.srgbClr == cor_tela
    assert vg.max_row <= 2                                       # só título e gráficos
    dg = wb['Dados dos gráficos']
    assert [dg.cell(1, 1).value, dg.cell(2, 1).value, dg.cell(3, 1).value, dg.cell(3, 2).value] == [
        'FVI – Status FVI por TAG', 'Status (FIC)', 'APROVADO', 3]
    blocos = [c.value for c in dg['A'] if isinstance(c.value, str) and c.value.split(' – ')[0] in
              ('FVI', 'FVM', 'CCM', 'TAP', 'TTAS')]
    assert blocos == ['FVI – Status FVI por TAG', 'FVM – Status da FVM por malha', 'CCM – Status do CCM por SSOP',
                      'TAP – Status do TAP por SSOP', 'TTAS – Status do TTAS por SOP']
    fvi = list(wb['FVI'].iter_rows(values_only=True))
    assert fvi[0] == ('Unidade', 'SOP', 'SSOP', 'TAG', 'Descrição da TAG', 'Nº da FVI', 'Status FVI', 'FVI Campo II',
                      'FVI Campo III', 'FVI Campo V')
    assert len(fvi) == 11
    t1 = next(r for r in fvi if r[3] == 'T1')
    assert t1[6:] == ('ACAMPO5', 'APROVADO', 'ACAMPO5', 'AAPROV')
    assert [c.value for c in wb['CCM'][1]] == ['Unidade', 'SOP', 'SSOP', 'Descrição da SSOP', 'Status do CCM']
    assert [c.value for c in wb['TTAS'][1]] == ['Unidade', 'SOP', 'Descrição do SOP', 'Situação TTAS 1', 'Situação TTAS 2']
    copia = tmp_path / 'copia.db'
    banco.copiar_para(copia)
    b2 = Banco(copia)
    assert len(b2.dados('FVI')) == 10
    b2.fechar()


def test_li_de_itens_e_malhas_conferidas_com_a_fic(banco, tmp_path):
    servico.importar_fic(banco, 'FVI', f.cadastro(tmp_path / 'c.xlsx', f.TAGS_SINTETICAS))
    servico.importar_fic(banco, 'FVM', f.provisorio(tmp_path / 'fvm.xlsx', ['MALHA', 'SOP', 'SSOP', 'STATUS'],
                                                    f.FVM_SINTETICO), '28/10/2026')
    itens = f.provisorio(tmp_path / 'itens.xlsx', ['TAG', 'SOP', 'SSOP', 'DESCRIÇÃO'],
                         [('T1', f.P1, f.S1, 'A'), ('T2', f.P2, f.S2, 'B'), ('TX', f.P1, f.S1, 'C'),
                          ('TY', '', '', 'não alocado'), ('Emissão: ', '', '', '')])
    assert servico.importar_base(banco, 'LI_ITENS', itens) == 'LI de Itens importada: 4 registro(s).'
    malhas = f.provisorio(tmp_path / 'malhas.xlsx', ['MALHA', 'SOP', 'SSOP', 'DESCRIÇÃO'],
                          [('M1', f.P1, f.S1, ''), ('M1', f.P2, f.S2, ''), ('M2', f.P1, f.S1, ''), ('MZ', 'X', 'SSOP-ZZ', '')])
    servico.importar_base(banco, 'LI_MALHAS', malhas)
    R = servico.calcular(banco)
    v = {t: (a, r) for _, t, _, a, r in R.verificacao}
    assert v['LI de Itens – registros lidos'] == (4, 'INFO')
    assert v['LI de Itens – registros sem SSOP'] == (1, 'INFO')
    assert v['LI de Itens – TAGs do cadastro FVI que não estão na LI de Itens'] == (8, 'INFO')   # T3..T10
    assert v['LI de Itens – TAGs da LI de Itens que não estão no cadastro FVI'] == (2, 'INFO')   # TX, TY
    assert v['LI de Itens – TAGs com SSOP diferente entre a LI de Itens e o cadastro FVI'] == (1, 'INFO')  # T2 (S2 x S1)
    assert v['LI de Malhas – malhas em mais de uma SSOP (pode, mas não deve)'] == (1, 'VERIFICAR')
    assert v['LI de Malhas – registros com SSOP fora da base'] == (1, 'VERIFICAR')
    assert v['LI de Malhas – MALHAs com SSOP diferente entre a LI de Malhas e o FVM'] == (1, 'INFO')     # M1
    # as bases não entram na data de referência nem no histórico, e não mudam o avanço
    assert banco.data_referencia_atual() == '2026-10-28'
    assert R.total['pct_cii'] == pytest.approx(4 / 6)
