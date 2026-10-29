"""Exportação do relatório gerencial: Excel e PDF com as páginas Visão Gerencial, FVI, FVM, CCM, TAP e TTAS
(mesmos nomes, colunas e cores da tela) e a logo em cada aba / página."""
import datetime
import io
from pathlib import Path

from fpdf import FPDF
from fpdf.fonts import FontFace
from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.drawing.colors import ColorChoice
from openpyxl.drawing.fill import PatternFillProperties
from openpyxl.drawing.image import Image as ImagemXls
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

from . import rotulos, tema
from .graficos import imagem_pizza
from .icones import arquivo_logo, recurso

# no Excel, fontes presentes em qualquer Windows: Bahnschrift (títulos) e Consolas (dados)
F_TIT, F = 'Bahnschrift', 'Consolas'
HF = Font(name=F_TIT, bold=True, color='FFFFFF', size=10)
HFILL = PatternFill('solid', fgColor='0A0A0A')
NF = Font(name=F, size=10, color='0A0A0A')
BF = Font(name=F, size=10, bold=True, color='0A0A0A')
AMARELO = PatternFill('solid', fgColor='EBEBED')      # aviso: cinza (sunken)
VERMELHO = PatternFill('solid', fgColor='D4D4D8')     # alerta: cinza mais forte
LINHA = Border(bottom=Side(style='thin', color='D4D4D8'))
TEXTURA_XLS = {'hachura': 'wdDnDiag', 'pontos': 'pct25', 'cruzado': 'diagCross', 'linhas': 'narHorz', 'grade': 'smGrid'}
PCT = '0.0%'
TITULO = 'CONTROLE GERENCIAL DE COMISSIONAMENTO // EPC-12'
SUBTITULOS = {'FVI': 'Acompanhamento TAG a TAG – status da FIC dos Campos II, III e V',
              'FVM': 'Acompanhamento malha a malha – TAGs da malha e status da FVM',
              'CCM': 'Status do CCM por SSOP', 'TAP': 'Status do TAP por SOP',
              'TTAS': 'Situação do TTAS 1 e do TTAS 2 por SOP'}


def _paginas(R):
    """página -> (colunas da tela, usadas no PDF; colunas do Excel; linhas)."""
    return {'FVI': (rotulos.COL_TAGS, rotulos.COL_TAGS_XLS, R.tags),
            'FVM': (rotulos.COL_MALHAS, rotulos.COL_MALHAS_XLS, R.malhas),
            'CCM': (rotulos.COL_CCM, rotulos.COL_CCM, R.ssop),
            'TAP': (rotulos.COL_TAP, rotulos.COL_TAP, R.sop),
            'TTAS': (rotulos.COL_TTAS, rotulos.COL_TTAS, R.sop)}


def distribuicoes(R, status_fic):
    """As 5 pizzas da Visão Gerencial (geral): [(documento, título, unidade, [(status, n)], [estilo])]."""
    fontes = (('FVI', 'Status FVI por TAG', 'TAGs', [t['st_fvi'] for t in R.tags]),
              ('FVM', 'Status da FVM por malha', 'malhas', list({m['malha']: m['status'] for m in R.malhas}.values())),
              ('CCM', 'Status do CCM por SSOP', 'SSOP', [x['txt_ccm_status'] for x in R.ssop]),
              ('TAP', 'Status do TAP por SOP', 'SOP', [x['txt_tap_status'] for x in R.sop]),
              ('TTAS', 'Status do TTAS por SOP', 'SOP', [x['txt_ttas_status'] for x in R.sop]))
    saida = []
    for doc, titulo, unidade, valores in fontes:
        itens = rotulos.distribuicao(valores) or [('Sem dados importados', 0)]   # as 5 caixas sempre presentes
        cores = tema.cores_pizza([r for r, _ in itens], [c for d_, c, _ in status_fic if d_ == doc])
        saida.append((doc, titulo, unidade, itens, cores))
    return saida


def exportar_relatorio(R, destino, data_referencia='', incluir_sem_fvi=True, status_fic=(), pastas_logo=()):
    """Grava destino (.xlsx) e o PDF de mesmo nome. Devolve (caminho do Excel, caminho do PDF)."""
    rotulos.preparar(R)
    destino = Path(destino)
    pdf = destino.with_suffix('.pdf')
    logo = arquivo_logo(pastas_logo)
    gerado = f'{datetime.datetime.now():%d/%m/%Y %H:%M}'
    pizzas = distribuicoes(R, status_fic)
    _excel(R, destino, data_referencia, gerado, incluir_sem_fvi, pizzas, logo)
    _pdf(R, pdf, data_referencia, gerado, incluir_sem_fvi, pizzas, logo)
    return destino, pdf


# ====================================================================== Excel
def _cabecalho(ws, colunas, linha=1, col0=1):
    for j, (titulo, _, larg, _) in enumerate(colunas, col0):
        c = ws.cell(linha, j, titulo)
        c.font, c.fill = HF, HFILL
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[L(j)].width = max(larg + 2, 9)
    ws.row_dimensions[linha].height = 34


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
    if linhas and filtro:
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


def _topo_xls(ws, logo, titulo, subtitulo, col_texto):
    """Logo no canto e título/subtítulo ao lado (linhas 1 e 2)."""
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 18
    if logo:
        img = ImagemXls(str(logo))
        img.height, img.width = 54, round(img.width * 54 / img.height)
        ws.add_image(img, 'A1')
    ws.cell(1, col_texto, titulo).font = Font(name=F_TIT, bold=True, size=15, color='0A0A0A')
    ws.cell(2, col_texto, subtitulo).font = Font(name=F, size=10, color='71717A')


def _pizza_xls(ws, a, b, titulo, ancora, cores):
    """Pizza nativa com rótulos de %, mesmo tamanho para todos os documentos; dados nas colunas A:B da própria aba."""
    g = PieChart()
    g.title = titulo
    g.add_data(Reference(ws, min_col=2, min_row=a, max_row=b), titles_from_data=False)
    g.set_categories(Reference(ws, min_col=1, min_row=a, max_row=b))
    g.height, g.width = 8.5, 15
    s = g.series[0]
    for i, (tipo, cor) in enumerate(cores):
        cor = cor.lstrip('#').upper()
        pt = DataPoint(idx=i)
        if tipo == 'solido':
            pt.graphicalProperties.solidFill = cor
        elif tipo == 'vazio':
            pt.graphicalProperties.solidFill = 'FFFFFF'
        else:   # textura nativa do Excel, mesma ideia da tela
            pt.graphicalProperties.pattFill = PatternFillProperties(
                prst=TEXTURA_XLS[tipo], fgClr=ColorChoice(srgbClr=cor), bgClr=ColorChoice(srgbClr='FFFFFF'))
        # separador branco entre fatias; a fatia vazia (sem status/registro) ganha contorno cinza para aparecer
        pt.graphicalProperties.line.solidFill = 'A1A1AA' if tipo == 'vazio' else 'FFFFFF'
        s.dPt.append(pt)
    s.dLbls = DataLabelList()
    s.dLbls.showPercent, s.dLbls.showVal, s.dLbls.showCatName, s.dLbls.showSerName = True, False, False, False
    s.dLbls.showLeaderLines = False
    s.dLbls.showLegendKey = False
    g.legend.position = 'r'
    ws.add_chart(g, ancora)


def _excel(R, destino, data_referencia, gerado, incluir_sem_fvi, pizzas, logo):
    wb = Workbook()
    ref = f'Data de referência: {data_referencia or "–"}   ·   Gerado em {gerado}'
    # ---------------------------------------------------------------- Visão Gerencial: 5 pizzas + os números delas
    ws = wb.active
    ws.title = 'Visão Gerencial'
    ws.sheet_view.showGridLines = False
    _topo_xls(ws, logo, 'VISÃO GERENCIAL // ' + TITULO, ref + f'   ·   TAGs sem FVI gerada entram no cálculo do FVI: '
              f'{"Sim" if incluir_sem_fvi else "Não"}', 3)
    lin = 60                                         # abaixo das pizzas (A4, J4, A22, J22, A40)
    ws.cell(lin - 2, 1, 'DADOS DOS GRÁFICOS – quantidade por status da FIC').font = Font(name=F_TIT, bold=True, size=12)
    for k, (doc, titulo, _, itens, cores) in enumerate(pizzas):
        ws.cell(lin, 1, f'{doc} – {titulo}').font = BF
        _cabecalho(ws, [('Status (FIC)', 's', 24, 0), ('Quantidade', 'n', 12, 0)], linha=lin + 1)
        for i_, (rot, n) in enumerate(itens, lin + 2):
            ws.cell(i_, 1, rot).font = NF
            ws.cell(i_, 2, n).font = NF
        _pizza_xls(ws, lin + 2, lin + 1 + len(itens), doc, ('A4', 'J4', 'A22', 'J22', 'A40')[k], cores)
        lin += len(itens) + 4
    ws.freeze_panes = None
    # ---------------------------------------------------------------- FVI, FVM, CCM, TAP, TTAS
    destaque = {'Sem FVI gerada': AMARELO, 'Sem SSOP': VERMELHO, 'SSOP fora da base': VERMELHO}
    for nome, (_, cols, linhas) in _paginas(R).items():
        ws = wb.create_sheet(nome)
        _topo_xls(ws, logo, f'{nome} // {TITULO}', f'{SUBTITULOS[nome]}   ·   {ref}', 2)
        _tabela(ws, cols, linhas, linha0=4)
        if nome in ('FVI', 'FVM'):
            _realcar(ws, linhas, 'txt_situacao', destaque, linha0=4, ncols=len(cols))
    wb.save(destino)


# ====================================================================== PDF
class _Relatorio(FPDF):
    """A4 paisagem com cabeçalho (logo, título, data) e rodapé (página) em todas as páginas."""

    def __init__(self, logo, ref):
        super().__init__(orientation='L', unit='mm', format='A4')
        self.logo, self.ref, self.secao = logo, ref, ''
        pasta = recurso('fontes')
        self.add_font('display', 'B', str(pasta / 'EPC12Display-Bold.ttf'))
        self.add_font('titulo', '', str(pasta / 'Rajdhani-Regular.ttf'))
        self.add_font('titulo', 'B', str(pasta / 'Rajdhani-Bold.ttf'))
        self.add_font('corpo', '', str(pasta / 'IBMPlexMono-Regular.ttf'))
        self.add_font('corpo', 'B', str(pasta / 'IBMPlexMono-Bold.ttf'))
        self.set_margins(10, 24, 10)
        self.set_auto_page_break(True, 12)
        self.set_title('Controle Gerencial de Comissionamento – EPC 12')
        self.set_creator('Controle Gerencial de Comissionamento – EPC 12')

    def header(self):
        if self.logo:
            self.image(str(self.logo), x=10, y=7, h=12)
        self.set_xy(30, 7)
        self.set_font('display', 'B', 13)
        self.set_text_color(10, 10, 10)
        self.cell(0, 6, self.secao.upper())
        self.set_xy(30, 13.5)
        self.set_font('titulo', 'B', 9)
        self.set_text_color(113, 113, 122)
        self.cell(0, 4, TITULO)
        self.set_xy(-110, 13.5)
        self.set_font('corpo', '', 7.5)
        self.cell(100, 4, self.ref, align='R')
        self.set_draw_color(10, 10, 10)
        self.set_line_width(0.5)
        self.line(10, 20.5, self.w - 10, 20.5)
        self.set_y(24)

    def footer(self):
        self.set_y(-9)
        self.set_font('corpo', '', 7)
        self.set_text_color(113, 113, 122)
        self.cell(0, 4, f'PÁGINA {self.page_no()}/{{nb}}', align='R')


def _rgb(cor):
    cor = cor.lstrip('#')
    return tuple(int(cor[i:i + 2], 16) for i in (0, 2, 4))


def _cartao_pizza(pdf, x, y, w, h, doc, titulo, unidade, itens, cores):
    pdf.set_draw_color(*_rgb(tema.LINHA_SUTIL))
    pdf.set_fill_color(*_rgb(tema.SUPERFICIE))
    pdf.set_line_width(0.2)
    pdf.rect(x, y, w, h, style='DF')
    pdf.set_xy(x + 4, y + 3)
    pdf.set_font('display', 'B', 12)
    pdf.set_text_color(10, 10, 10)
    pdf.cell(w - 8, 6, doc)
    total = sum(n for _, n in itens)
    pdf.set_xy(x + 4, y + 9.5)
    pdf.set_font('corpo', '', 7)
    pdf.set_text_color(113, 113, 122)
    pdf.cell(w - 8, 4, f'{titulo} · {total:,} {unidade}'.replace(',', '.'))
    d = h - 19
    if not total:
        pdf.set_xy(x + 4, y + h / 2)
        pdf.cell(w - 8, 5, 'SEM DADOS IMPORTADOS', align='C')
        return
    fatias = [(n, e) for (_, n), e in zip(itens, cores) if n > 0]
    img, _ = imagem_pizza(fatias, 600)
    buf = io.BytesIO()
    img.save(buf, 'PNG')
    pdf.image(buf, x=x + 4, y=y + 15, w=d, h=d)
    # legenda: amostra da cor, status, quantidade e %
    lx, ly = x + 8 + d, y + 15 + max(0, (d - 5.2 * len(itens)) / 2)
    for (rot, n), estilo in zip(itens, cores):
        if estilo[0] == 'vazio':
            pdf.set_fill_color(255, 255, 255)
            pdf.set_draw_color(*_rgb(tema.DESABILITADO))
            pdf.rect(lx, ly + 0.9, 3, 3, style='DF')
        else:
            pdf.set_fill_color(*_rgb(estilo[1]))
            pdf.rect(lx, ly + 0.9, 3, 3, style='F')
        pdf.set_xy(lx + 4.5, ly)
        pdf.set_font('titulo', 'B', 8.5)
        pdf.set_text_color(10, 10, 10)
        larg = x + w - 4 - 30 - (lx + 4.5)             # 30 mm para quantidade e %
        texto = rot
        while pdf.get_string_width(texto) > larg - 1 and len(texto) > 1:
            texto = texto[:-2] + '…'
        pdf.cell(larg, 4.8, texto)
        pdf.set_font('corpo', '', 7.5)
        pdf.cell(14, 4.8, f'{n:,}'.replace(',', '.'), align='R')
        pdf.cell(16, 4.8, f'{n / total:.1%}'.replace('.', ','), align='R')
        ly += 5.2


def _tabela_pdf(pdf, colunas, linhas):
    if not linhas:
        pdf.set_font('corpo', '', 8)
        pdf.cell(0, 6, 'Sem dados importados.')
        return
    larguras = [max(larg, 8) for _, _, larg, _ in colunas]
    cab = FontFace(family='titulo', emphasis='B', color=(255, 255, 255), fill_color=_rgb(tema.INVERSO), size_pt=8.5)
    pdf.set_font('corpo', '', 7)
    pdf.set_text_color(10, 10, 10)
    pdf.set_draw_color(*_rgb(tema.LINHA_FRACA))
    with pdf.table(col_widths=larguras, headings_style=cab, line_height=4.2, borders_layout='HORIZONTAL_LINES',
                   cell_fill_color=_rgb('#FAFAFA'), cell_fill_mode='ROWS', repeat_headings=1,
                   text_align=['CENTER' if p == 2 else 'LEFT' for _, _, _, p in colunas]) as t:
        cabecalho = t.row()
        for titulo, _, _, _ in colunas:
            cabecalho.cell(titulo.upper())
        for r in linhas:
            linha = t.row()
            for _, k, _, p in colunas:
                v = r.get(k)
                linha.cell('' if v is None else (f'{v:.1%}'.replace('.', ',') if p == 1 else str(v)))


def _pdf(R, destino, data_referencia, gerado, incluir_sem_fvi, pizzas, logo):
    pdf = _Relatorio(logo, f'REF {data_referencia or "–"}  ·  GERADO EM {gerado}')
    pdf.secao = 'Visão Gerencial'
    pdf.add_page()
    pdf.set_font('corpo', '', 7.5)
    pdf.set_text_color(113, 113, 122)
    pdf.cell(0, 4, 'Status da FIC em cada documento – geral. TAGs sem FVI gerada entram no cálculo do FVI: '
                   f'{"Sim" if incluir_sem_fvi else "Não"}.')
    w, h, gap, y0 = (pdf.w - 20 - 4) / 2, 52, 4, 30        # 2 cartões por linha: legenda com espaço
    for i, (doc, titulo, unidade, itens, cores) in enumerate(pizzas):
        lin, col = divmod(i, 2)
        _cartao_pizza(pdf, 10 + col * (w + gap), y0 + lin * (h + gap), w, h, doc, titulo, unidade, itens, cores)
    for nome, (cols, _, linhas) in _paginas(R).items():
        pdf.secao = nome
        pdf.add_page()
        pdf.set_font('corpo', '', 7.5)
        pdf.set_text_color(113, 113, 122)
        pdf.cell(0, 5, f'{SUBTITULOS[nome]} · {len(linhas):,} linha(s)'.replace(',', '.'), new_x='LMARGIN', new_y='NEXT')
        pdf.ln(1)
        _tabela_pdf(pdf, cols, linhas)
    pdf.output(str(destino))

