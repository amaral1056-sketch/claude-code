"""Exportação para Excel: relatório gerencial (mesma organização e nomes da tela) e planilhas-modelo de importação."""
import datetime

from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

from . import rotulos, tema
from .importacao import CABECALHO_MODELO
from .rotulos import COL_VISAO, NOTAS_VISAO

F = 'Segoe UI'
HF = Font(name=F, bold=True, color='FFFFFF', size=10)
HFILL = PatternFill('solid', fgColor='2A5FA8')
NF = Font(name=F, size=10)
BF = Font(name=F, size=10, bold=True)
TOTAL_FILL = PatternFill('solid', fgColor='EAF2FD')
AMARELO = PatternFill('solid', fgColor='FFF4D6')
VERMELHO = PatternFill('solid', fgColor='FBE9E5')
LINHA = Border(bottom=Side(style='thin', color='E1E5EA'))
PCT = '0.0%'


def _cabecalho(ws, colunas, linha=1, col0=1):
    for j, (titulo, _, larg, _) in enumerate(colunas, col0):
        c = ws.cell(linha, j, titulo)
        c.font, c.fill = HF, HFILL
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[L(j)].width = max(larg + 2, 9)
    ws.row_dimensions[linha].height = 34


def _escala(ws, rng):
    ws.conditional_formatting.add(rng, ColorScaleRule(start_type='num', start_value=0, start_color='F6C7BC',
                                                      mid_type='num', mid_value=0.5, mid_color='FCEBB6',
                                                      end_type='num', end_value=1, end_color='BFE3C8'))


def _tabela(ws, colunas, linhas, linha0=1, col0=1, filtro=True):
    _cabecalho(ws, colunas, linha0, col0)
    for i, r in enumerate(linhas, linha0 + 1):
        for j, (_, k, _, p) in enumerate(colunas, col0):
            c = ws.cell(i, j, r.get(k))
            c.font, c.border = NF, LINHA
            if p == 1:
                c.number_format = PCT
            elif p == 2:
                c.alignment = Alignment(horizontal='center')
    if linhas:
        for j, (_, _, _, p) in enumerate(colunas, col0):
            if p == 1:
                _escala(ws, f'{L(j)}{linha0 + 1}:{L(j)}{linha0 + len(linhas)}')
        if filtro:
            ws.auto_filter.ref = f'{L(col0)}{linha0}:{L(col0 + len(colunas) - 1)}{linha0 + len(linhas)}'
    ws.freeze_panes = ws.cell(linha0 + 1, col0 + 1)
    return linha0 + len(linhas)


def _realcar(ws, linhas, coluna, cores, linha0=1, ncols=None):
    """Pinta a linha inteira conforme o texto da coluna de situação."""
    for i, r in enumerate(linhas, linha0 + 1):
        fill = cores.get(r.get(coluna))
        if fill:
            for j in range(1, (ncols or ws.max_column) + 1):
                ws.cell(i, j).fill = fill


def _pizza(ws, dados, titulo, a, b, ancora, cores):
    """Pizza com rótulos de %, mesmo tamanho para todos os documentos."""
    g = PieChart()
    g.title = titulo
    g.add_data(Reference(dados, min_col=2, min_row=a, max_row=b), titles_from_data=False)
    g.set_categories(Reference(dados, min_col=1, min_row=a, max_row=b))
    g.height, g.width = 8.5, 15
    s = g.series[0]
    for i, cor in enumerate(cores):
        cor = cor.lstrip('#').upper()
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = cor
        pt.graphicalProperties.line.solidFill = 'FFFFFF'
        s.dPt.append(pt)
    s.dLbls = DataLabelList()
    s.dLbls.showPercent, s.dLbls.showVal, s.dLbls.showCatName, s.dLbls.showSerName = True, False, False, False
    s.dLbls.showLeaderLines = False
    s.dLbls.showLegendKey = False
    g.legend.position = 'r'
    ws.add_chart(g, ancora)


def exportar_relatorio(R, destino, data_referencia='', historico=(), criterios=(), incluir_sem_fvi=True, status_fic=()):
    rotulos.preparar(R)
    wb = Workbook()

    # ---------------------------------------------------------------- Visão Gerencial (só gráficos)
    ws = wb.active
    ws.title = 'Visão Gerencial'
    ws.sheet_view.showGridLines = False
    ws.cell(1, 1, 'Visão Gerencial – Acompanhamento do Comissionamento – EPC 12').font = Font(name=F, bold=True, size=15)
    ws.cell(2, 1, f'Data de referência: {data_referencia or "–"}   ·   Gerado em {datetime.datetime.now():%d/%m/%Y %H:%M}'
                  f'   ·   TAGs sem FVI gerada entram no cálculo do FVI: {"Sim" if incluir_sem_fvi else "Não"}').font = Font(
        name=F, size=10, color='5F6670')
    # distribuição dos status da FIC (geral) em aba própria; as pizzas leem dela
    dados = wb.create_sheet('Dados dos gráficos')
    fontes = (('FVI', 'Status FVI por TAG', [t['st_fvi'] for t in R.tags]),
              ('FVM', 'Status da FVM por malha', [m['status'] for m in R.malhas]),
              ('CCM', 'Status do CCM por SSOP', [x['txt_ccm_status'] for x in R.ssop]),
              ('TAP', 'Status do TAP por SSOP', [x['txt_tap_status'] for x in R.ssop]),
              ('TTAS', 'Status do TTAS por SOP', [x['txt_ttas_status'] for x in R.sop]))
    lin = 1
    for k, (doc, titulo, valores) in enumerate(fontes):
        itens = rotulos.distribuicao(valores) or [('Sem dados importados', 0)]   # as 5 caixas sempre presentes
        dados.cell(lin, 1, f'{doc} – {titulo}').font = BF
        _cabecalho(dados, [('Status (FIC)', 's', 22, 0), ('Quantidade', 'n', 12, 0)], linha=lin + 1)
        for i_, (rot, n) in enumerate(itens, lin + 2):
            dados.cell(i_, 1, rot).font = NF
            dados.cell(i_, 2, n).font = NF
        cores = tema.cores_pizza([r for r, _ in itens], [c for d_, c, _ in status_fic if d_ == doc])
        _pizza(ws, dados, doc, lin + 2, lin + 1 + len(itens), ('A4', 'J4', 'A22', 'J22', 'A40')[k], cores)
        lin += len(itens) + 4
    # avanço por unidade (números do acompanhamento) logo abaixo
    dados.cell(lin, 1, 'Avanço por unidade').font = BF
    linhas = [dict(v) for v in R.visao] + [{**R.total, 'unidade': 'Total'}]
    fim = _tabela(dados, COL_VISAO, linhas, linha0=lin + 1, filtro=False)
    for j_ in range(1, len(COL_VISAO) + 1):
        dados.cell(fim, j_).font, dados.cell(fim, j_).fill = BF, TOTAL_FILL
    for i_, t in enumerate(NOTAS_VISAO, fim + 2):
        dados.cell(i_, 1, t).font = Font(name=F, italic=True, size=9, color='5F6670')
    dados.freeze_panes = None

    # ---------------------------------------------------------------- FVI e FVM (item a item)
    cores = {'Sem FVI gerada': AMARELO, 'Sem SSOP': VERMELHO, 'SSOP fora da base': VERMELHO}
    ws = wb.create_sheet('FVI')
    _tabela(ws, rotulos.COL_TAGS_XLS, R.tags)
    _realcar(ws, R.tags, 'txt_situacao', cores, ncols=len(rotulos.COL_TAGS_XLS))
    ws = wb.create_sheet('FVM')
    _tabela(ws, rotulos.COL_MALHAS_XLS, R.malhas)
    _realcar(ws, R.malhas, 'txt_situacao', cores, ncols=len(rotulos.COL_MALHAS_XLS))

    # ---------------------------------------------------------------- documentos
    for nome, cols, linhas in (('CCM', rotulos.COL_CCM, R.ssop), ('TAP', rotulos.COL_TAP, R.ssop),
                               ('TTAS', rotulos.COL_TTAS, R.sop)):
        _tabela(wb.create_sheet(nome), cols, linhas)

    # ---------------------------------------------------------------- controle
    ws = wb.create_sheet('Verificação')
    ver = [{'g': g, 't': t, 'e': '-' if e is None else e, 'a': a, 'r': {'OK': 'OK', 'VERIFICAR': 'Verificar',
                                                                     'INFO': 'Informação'}[r]}
           for g, t, e, a, r in R.verificacao]
    _tabela(ws, rotulos.COL_VER, ver)
    _realcar(ws, ver, 'r', {'Verificar': VERMELHO}, ncols=len(rotulos.COL_VER))
    _tabela(wb.create_sheet('Pendências'), rotulos.COL_PEND, [dict(zip('dlisp', p)) for p in R.pendencias])
    if historico:
        _tabela(wb.create_sheet('Histórico'), rotulos.COL_HIST, list(historico))
    if criterios:
        ws = wb.create_sheet('Critérios')
        crit = [{'etapa': e, 'codigo': c, 'origem': o} for _, e, c, o in criterios]
        _tabela(ws, rotulos.COL_CRIT, crit)
        for i, r in enumerate(crit, 2):
            if str(r['origem']).startswith('A CONFIRMAR'):
                for j in (1, 2, 3):
                    ws.cell(i, j).fill = AMARELO
    wb.move_sheet('Dados dos gráficos', offset=len(wb.sheetnames))
    wb.save(destino)


def exportar_modelos(destino):
    """Uma aba por documento com o cabeçalho que o importador reconhece (layout provisório)."""
    wb = Workbook()
    wb.remove(wb.active)
    for tipo, cab in CABECALHO_MODELO.items():
        ws = wb.create_sheet(tipo)
        for j, t in enumerate(cab, 1):
            c = ws.cell(1, j, t)
            c.font, c.fill = HF, HFILL
            ws.column_dimensions[L(j)].width = 26
        ws.freeze_panes = 'A2'
    wb.save(destino)
