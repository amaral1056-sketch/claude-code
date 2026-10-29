"""Exportação para Excel: relatório gerencial (valores formatados) e planilhas-modelo de importação."""
import datetime

from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

from .importacao import CABECALHO_MODELO

F = 'Arial'
HF = Font(name=F, bold=True, color='FFFFFF')
HFILL = PatternFill('solid', fgColor='1F4E78')
NF = Font(name=F, size=10)
BF = Font(name=F, size=10, bold=True)
AMARELO = PatternFill('solid', fgColor='FFF2CC')
PCT = '0.0%'

# (título, chave, largura, é percentual)
COL_VISAO = [('UNIDADE', 'unidade', 11, 0), ('SOP', 'sop', 7, 0), ('SSOP', 'ssop', 7, 0), ('TAGS (FVI)', 'tags', 9, 0),
             ('% FVI CII', 'pct_cii', 9, 1), ('% FVI CIII', 'pct_ciii', 9, 1), ('% FVI CV', 'pct_cv', 9, 1),
             ('MALHAS (FVM)', 'malhas', 9, 0), ('% FVM CIII', 'pct_fvm_ciii', 9, 1), ('% FVM CV', 'pct_fvm_cv', 9, 1),
             ('SSOP c/ CCM APROV.', 'ccm', 10, 0), ('% CCM', 'pct_ccm', 8, 1), ('SSOP c/ TAP CONCL.', 'tap', 10, 0),
             ('% TAP', 'pct_tap', 8, 1), ('SOP c/ TTAS1', 'ttas1', 9, 0), ('% TTAS1', 'pct_ttas1', 8, 1),
             ('SOP c/ TTAS2', 'ttas2', 9, 0), ('% TTAS2', 'pct_ttas2', 8, 1), ('TAGS SEM FVI', 'tags_sem_fvi', 9, 0)]
_FVI = [('TAGS', 'tags', 8, 0), ('CII APLIC.', 'cii_aplic', 9, 0), ('CII APROV.', 'cii_aprov', 9, 0), ('% CII', 'pct_cii', 8, 1),
        ('CIII APLIC.', 'ciii_aplic', 9, 0), ('CIII APROV.', 'ciii_aprov', 9, 0), ('% CIII', 'pct_ciii', 8, 1),
        ('CV APLIC.', 'cv_aplic', 9, 0), ('CV APROV.', 'cv_aprov', 9, 0), ('% CV', 'pct_cv', 8, 1),
        ('TAGS SEM FVI', 'tags_sem_fvi', 9, 0)]
_FVM = [('MALHAS', 'malhas', 9, 0), ('CIII APROV.', 'fvm_ciii', 9, 0), ('% CIII', 'pct_fvm_ciii', 8, 1),
        ('CV APROV.', 'fvm_cv', 9, 0), ('% CV', 'pct_fvm_cv', 8, 1)]
_SSOP = [('UNIDADE', 'unidade', 10, 0), ('SOP', 'sop', 22, 0), ('SSOP', 'ssop', 26, 0), ('DESCRIÇÃO SSOP', 'descricao', 40, 0)]
_SOP = [('UNIDADE', 'unidade', 10, 0), ('SOP', 'sop', 22, 0), ('DESCRIÇÃO SOP', 'descricao', 40, 0), ('QTD SSOP', 'qtd_ssop', 8, 0)]
COL_FVI_SSOP = _SSOP + _FVI
COL_FVI_SOP = _SOP + _FVI
COL_FVM_SSOP = _SSOP + _FVM
COL_FVM_SOP = _SOP + _FVM
COL_CCM = _SSOP + [('CCM REGISTRADOS', 'ccm_reg', 12, 0), ('APROVADO', 'ccm_aprovado', 10, 0), ('SITUAÇÃO', 'ccm_situacao', 16, 0)]
COL_TAP = _SSOP + [('TAP REGISTRADOS', 'tap_reg', 12, 0), ('CONCLUÍDO', 'tap_concluido', 10, 0), ('SITUAÇÃO', 'tap_situacao', 16, 0)]
COL_TTAS = _SOP + [('TTAS REGISTRADOS', 'ttas_reg', 12, 0), ('TTAS1 APROVADO', 'ttas1_aprovado', 12, 0),
                   ('TTAS2 APROVADO', 'ttas2_aprovado', 12, 0), ('SITUAÇÃO', 'ttas_situacao', 18, 0)]

COL_TAGS_XLS = [('TAG', 'tag', 20, 0), ('DESCRIÇÃO', 'descricao', 34, 0), ('UNIDADE', 'unidade', 10, 0),
                ('SOP', 'sop', 22, 0), ('SSOP', 'ssop', 28, 0), ('CLASSE', 'classe', 30, 0), ('Nº FVI', 'n_fvi', 10, 0),
                ('STATUS FVI', 'st_fvi', 13, 0), ('STATUS CAMPO II', 'st_cii', 14, 0), ('CII CONTA?', 'x_cii', 9, 0),
                ('STATUS CAMPO III', 'st_ciii', 14, 0), ('CIII CONTA?', 'x_ciii', 9, 0), ('STATUS CAMPO IV', 'st_civ', 13, 0),
                ('STATUS CAMPO V', 'st_cv', 14, 0), ('CV CONTA?', 'x_cv', 9, 0), ('SITUAÇÃO', 'situacao', 14, 0),
                ('SEM FVI', 'x_sem_fvi', 8, 0)]
COL_MALHAS_XLS = [('MALHA', 'malha', 22, 0), ('UNIDADE', 'unidade', 10, 0), ('SOP', 'sop', 22, 0), ('SSOP', 'ssop', 28, 0),
                  ('STATUS', 'status', 16, 0), ('CAMPO III ATINGIDO?', 'x_ciii', 12, 0),
                  ('CAMPO V ATINGIDO?', 'x_cv', 12, 0), ('SITUAÇÃO', 'situacao', 14, 0)]


def _sim_nao(v):
    return 'N/A' if v is None else ('SIM' if v else 'NÃO')


NOTAS_VISAO = [
    '% FVI = aprovados ÷ aplicáveis (N/A e cancelados fora do denominador). % FVM = aprovadas ÷ malhas válidas.',
    '% CCM e % TAP = SSOP com documento aprovado/concluído ÷ total de SSOP da unidade. % TTAS = SOP com TTAS1/TTAS2 aprovado ÷ total de SOP.',
    'Totais ponderados (não é média de %). Itens com SSOP/SOP fora da base e TAGs sem SSOP não entram – ver VERIFICACAO.',
]


def _cabecalho(ws, colunas, linha=1, col0=1):
    for j, (titulo, _, larg, _) in enumerate(colunas, col0):
        c = ws.cell(linha, j, titulo)
        c.font, c.fill = HF, HFILL
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[L(j)].width = larg
    ws.row_dimensions[linha].height = 32


def _escala(ws, rng):
    ws.conditional_formatting.add(rng, ColorScaleRule(start_type='num', start_value=0, start_color='F8696B',
                                                      mid_type='num', mid_value=0.5, mid_color='FFEB84',
                                                      end_type='num', end_value=1, end_color='63BE7B'))


def _tabela(ws, colunas, linhas, linha0=1, col0=1, fonte=NF):
    _cabecalho(ws, colunas, linha0, col0)
    for i, r in enumerate(linhas, linha0 + 1):
        for j, (_, k, _, p) in enumerate(colunas, col0):
            c = ws.cell(i, j, r.get(k))
            c.font = fonte
            if p:
                c.number_format = PCT
    if linhas:
        for j, (_, _, _, p) in enumerate(colunas, col0):
            if p:
                _escala(ws, f'{L(j)}{linha0 + 1}:{L(j)}{linha0 + len(linhas)}')
    return linha0 + len(linhas)


def exportar_relatorio(R, destino, data_referencia='', historico=(), criterios=(), incluir_sem_fvi=True):
    wb = Workbook()
    ws = wb.active
    ws.title = 'VISAO_GERENCIAL'
    ws.cell(1, 1, 'VISÃO GERENCIAL – AVANÇO DO COMISSIONAMENTO POR UNIDADE – EPC 12').font = Font(name=F, bold=True, size=14)
    ws.cell(2, 1, f'Data de referência: {data_referencia or "-"}   ·   Gerado em {datetime.datetime.now():%d/%m/%Y %H:%M}'
                  f'   ·   TAGs sem FVI no denominador: {"SIM" if incluir_sem_fvi else "NÃO"}').font = BF
    fim = _tabela(ws, COL_VISAO, R.visao + [R.total], linha0=4)
    for j in range(1, len(COL_VISAO) + 1):
        ws.cell(fim, j).font = BF
    for i, t in enumerate(NOTAS_VISAO, fim + 2):
        ws.cell(i, 1, t).font = Font(name=F, italic=True, size=9)
    ws.freeze_panes = 'B5'

    for nome, partes in (('FVI', [(COL_FVI_SSOP, R.ssop), (COL_FVI_SOP, R.sop)]),
                         ('FVM', [(COL_FVM_SSOP, R.ssop), (COL_FVM_SOP, R.sop)]),
                         ('CCM', [(COL_CCM, R.ssop)]), ('TAP', [(COL_TAP, R.ssop)]), ('TTAS', [(COL_TTAS, R.sop)])):
        ws = wb.create_sheet(nome)
        col0 = 1
        for cols, linhas in partes:
            _tabela(ws, cols, linhas, col0=col0)
            col0 += len(cols) + 1
        ws.freeze_panes = 'E2'

    ws = wb.create_sheet('TAGS_FVI')
    _tabela(ws, COL_TAGS_XLS, [{**t, 'x_cii': _sim_nao(t['ok_cii']), 'x_ciii': _sim_nao(t['ok_ciii']),
                                 'x_cv': _sim_nao(t['ok_cv']), 'x_sem_fvi': 'SIM' if t['sem_fvi'] else ''} for t in R.tags])
    ws.freeze_panes = 'B2'
    ws.auto_filter.ref = f'A1:{L(len(COL_TAGS_XLS))}{len(R.tags) + 1}'
    ws = wb.create_sheet('MALHAS_FVM')
    _tabela(ws, COL_MALHAS_XLS, [{**m, 'x_ciii': _sim_nao(m['ok_ciii']), 'x_cv': _sim_nao(m['ok_cv'])} for m in R.malhas])
    ws.freeze_panes = 'B2'
    ws.auto_filter.ref = f'A1:{L(len(COL_MALHAS_XLS))}{len(R.malhas) + 1}'

    ws = wb.create_sheet('VERIFICACAO')
    cols = [('GRUPO', 'g', 12, 0), ('TESTE', 't', 70, 0), ('ESPERADO', 'e', 10, 0), ('APURADO', 'a', 10, 0), ('RESULTADO', 'r', 12, 0)]
    _tabela(ws, cols, [{'g': g, 't': t, 'e': '-' if e is None else e, 'a': a, 'r': r} for g, t, e, a, r in R.verificacao])
    for i in range(2, len(R.verificacao) + 2):
        if ws.cell(i, 5).value == 'VERIFICAR':
            ws.cell(i, 5).fill = AMARELO
    ws2 = wb.create_sheet('PENDENCIAS')
    cols = [('DOCUMENTO', 'd', 11, 0), ('LINHA NO ARQUIVO', 'l', 10, 0), ('IDENTIFICAÇÃO', 'i', 26, 0),
            ('SSOP / SOP', 's', 28, 0), ('PROBLEMA', 'p', 70, 0)]
    _tabela(ws2, cols, [dict(zip('dlisp', p)) for p in R.pendencias])
    ws2.freeze_panes = 'A2'

    if historico:
        ws = wb.create_sheet('HISTORICO')
        cols = [('DATA', 'data', 11, 0), ('UNIDADE', 'unidade', 11, 0), ('TAGS', 'tags', 8, 0), ('% FVI CII', 'pct_cii', 9, 1),
                ('% FVI CIII', 'pct_ciii', 9, 1), ('% FVI CV', 'pct_cv', 9, 1), ('MALHAS', 'malhas', 8, 0),
                ('% FVM CIII', 'pct_fvm_ciii', 9, 1), ('% FVM CV', 'pct_fvm_cv', 9, 1), ('% CCM', 'pct_ccm', 8, 1),
                ('% TAP', 'pct_tap', 8, 1), ('% TTAS1', 'pct_ttas1', 8, 1), ('% TTAS2', 'pct_ttas2', 8, 1)]
        _tabela(ws, cols, list(historico))
        ws.freeze_panes = 'C2'

    if criterios:
        ws = wb.create_sheet('CRITERIOS')
        cols = [('ETAPA', 'e', 14, 0), ('CÓDIGO DE STATUS', 'c', 22, 0), ('ORIGEM', 'o', 80, 0)]
        _tabela(ws, cols, [{'e': e, 'c': c, 'o': o} for _, e, c, o in criterios])
        for i in range(2, len(criterios) + 2):
            if str(ws.cell(i, 3).value or '').startswith('A CONFIRMAR'):
                for j in (1, 2, 3):
                    ws.cell(i, j).fill = AMARELO
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
