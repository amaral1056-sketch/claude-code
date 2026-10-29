"""Importadores: layouts aceitos, recusas (nada é gravado) e exportações."""
import datetime

import openpyxl
import pytest

from comissionamento import exportacao, importacao, servico
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
    exportacao.exportar_relatorio(R, destino, '28/10/2026', servico.historico_com_percentuais(banco), banco.criterios())
    wb = openpyxl.load_workbook(destino)
    assert wb.sheetnames == ['Visão Gerencial', 'FVI', 'FVM', 'CCM', 'TAP', 'TTAS', 'Verificação', 'Pendências',
                             'Histórico', 'Critérios']
    vg = wb['Visão Gerencial']
    # indicadores gerais com nomes por extenso
    assert [vg.cell(5, j).value for j in range(1, 10)] == ['FVI Campo II', 'FVI Campo III', 'FVI Campo V', 'FVM Campo III',
                                                          'FVM Campo V', 'CCM', 'TAP', 'TTAS 1', 'TTAS 2']
    assert vg.cell(6, 1).value == pytest.approx(4 / 6)
    assert [vg.cell(9, j).value for j in (1, 5, 9, 15)] == ['Unidade', 'FVI Campo II (%)', 'FVM Campo III (%)',
                                                           'SOP com TTAS 1 aprovado']
    assert [vg.cell(r, 1).value for r in (10, 11, 12)] == [f.UN1, f.UN2, 'TOTAL']
    assert vg.cell(12, 5).value == pytest.approx(4 / 6)
    assert len(vg._charts) == 4                                  # FVI, FVM, CCM e TAP, TTAS por unidade
    fvi = list(wb['FVI'].iter_rows(values_only=True))
    assert fvi[0] == ('Unidade', 'SOP', 'SSOP', 'TAG', 'Descrição da TAG', 'Nº da FVI', 'FVI Campo II', 'FVI Campo III',
                      'FVI Campo V', 'Situação')
    assert len(fvi) == 11
    t1 = next(r for r in fvi if r[3] == 'T1')
    assert t1[6:] == ('Aprovado', 'Aprovado (ACAMPO5)', 'Pendente (AAPROV)', 'Válida')
    copia = tmp_path / 'copia.db'
    banco.copiar_para(copia)
    b2 = Banco(copia)
    assert len(b2.dados('FVI')) == 10
    b2.fechar()
