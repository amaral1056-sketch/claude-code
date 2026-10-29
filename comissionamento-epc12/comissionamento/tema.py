"""Identidade visual: HUD monocromático (tokens hud-mono-light) – base branca, escalas de preto e cinza,
painéis retos com cantoneiras, alta densidade (grade de 4 px) e tipografia técnica em caixa-alta.

Fontes (licença OFL, embutidas em recursos/fontes e carregadas só para o programa):
display EPC12 Display (Orbitron com o til corrigido – ferramentas/fonte_display.py) · títulos Rajdhani · corpo IBM Plex Mono · dados Share Tech Mono."""
import tkinter as tk
from pathlib import Path
from tkinter import font as tkfont, ttk

import customtkinter as ctk

# ---------------------------------------------------------------- cores (tokens color.*)
BASE = '#FFFFFF'           # background.base
SUPERFICIE = '#F5F5F6'     # background.surface – painéis e janelas
AFUNDADO = '#EBEBED'       # background.sunken – campos, áreas internas
INVERSO = '#0A0A0A'        # background.inverse – blocos de destaque (texto branco)
TEXTO = '#0A0A0A'          # foreground.primary
TEXTO_2 = '#3F3F46'        # foreground.secondary
TEXTO_MUDO = '#71717A'     # foreground.muted
DESABILITADO = '#A1A1AA'   # foreground.disabled
SOBRE_INVERSO = '#FFFFFF'  # foreground.onInverse
ACENTO = '#000000'         # accent.solid
ACENTO_SUAVE = '#52525B'   # accent.soft
LINHA_FORTE = '#0A0A0A'    # line.strong
LINHA = '#52525B'          # line.default
LINHA_SUTIL = '#D4D4D8'    # line.subtle – grade, divisores
LINHA_FRACA = '#E4E4E7'    # line.faint – pontilhados
DADOS = ('#0A0A0A', '#52525B', '#A1A1AA', '#D4D4D8')   # data.series1..4 (tons da interface)
# gráficos: paleta categórica validada (daltonismo e visão normal, ordem fixa – a ordem faz parte da validação;
# 4 cores ficam abaixo de 3:1 no fundo, por isso as pizzas sempre têm legenda com quantidade e %)
CATEGORICAS = ('#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#6250d6', '#e34948')

# papéis usados pelo restante do código
FUNDO = BASE
BORDA = LINHA_SUTIL
BORDA_FORTE = LINHA
GRADE = LINHA_FRACA
EIXO = LINHA
DESTAQUE = INVERSO
DESTAQUE_ESCURO = TEXTO_2          # hover do bloco preto
DESTAQUE_CLARO = AFUNDADO
PASSAR = AFUNDADO
LISTRA = '#FAFAFA'
SELECAO = LINHA_SUTIL
SERIE = CATEGORICAS[:3]            # linhas: as 3 primeiras valem em todos os pares
NEUTRO = LINHA_SUTIL
# estados: o destaque máximo é o preenchimento sólido (sem vermelho/amarelo)
OK = TEXTO
ERRO = SOBRE_INVERSO               # texto da linha em alerta
ERRO_FUNDO = INVERSO               # linha em alerta = bloco preto
ALERTA_FUNDO = AFUNDADO            # linha de aviso = cinza
AMARELO_CRITERIO = LINHA_FRACA     # critério "a confirmar"

RAIO = 0          # border.radius.none – painéis
RAIO_BOTAO = 2    # border.radius.xs – botões e chips
CANTONEIRA = 8    # border.cornerBracketSize
TRACO_FORTE = 2   # border.width.strong

# ---------------------------------------------------------------- pizzas: cores categóricas; cinzas para o que não é status
# (tipo, cor): 'solido' | 'hachura' | 'pontos' | 'cruzado' | 'linhas' | 'grade' | 'vazio'
PIZZA = tuple(('solido', c) for c in CATEGORICAS)
PIZZA_NA = ('solido', DESABILITADO)
PIZZA_SEM = ('vazio', DESABILITADO)
PIZZA_OUTROS = ('hachura', TEXTO_MUDO)


def cores_pizza(rotulos_fatias, ordem_oficial):
    """Estilo de cada fatia: segue o código (posição na lista oficial do documento), não o tamanho da fatia.
    N/A, sem status/registro e 'Outros' têm estilos fixos. Usada na tela e no Excel."""
    fixas = {'N/A': PIZZA_NA, 'Sem status': PIZZA_SEM, 'Sem registro': PIZZA_SEM, 'Sem dados importados': PIZZA_SEM,
             'Outros': PIZZA_OUTROS}
    pos = {c.upper(): i for i, c in enumerate(ordem_oficial)}
    usadas, estilos = set(), []
    for rot in rotulos_fatias:
        if rot in fixas:
            estilos.append(fixas[rot])
            continue
        i = pos.get(rot.upper(), len(ordem_oficial) + len(usadas)) % len(PIZZA)
        while i in usadas and len(usadas) < len(PIZZA):
            i = (i + 1) % len(PIZZA)
        usadas.add(i)
        estilos.append(PIZZA[i])
    return estilos


# ---------------------------------------------------------------- tipografia (tokens font.*)
PREFERIDAS = {
    'display': ('EPC12 Display', 'Orbitron', 'Michroma', 'Bahnschrift', 'Segoe UI'),
    'titulo': ('Rajdhani', 'Chakra Petch', 'Bahnschrift', 'Segoe UI'),
    'corpo': ('IBM Plex Mono', 'JetBrains Mono', 'Consolas', 'DejaVu Sans Mono', 'Courier New'),
    'dado': ('Share Tech Mono', 'Space Mono', 'Consolas', 'DejaVu Sans Mono', 'Courier New'),
}
FAMILIAS = {k: v[-1] for k, v in PREFERIDAS.items()}
FAMILIA = FAMILIAS['corpo']
escala = 1.0
_fontes_carregadas = False


def pasta_fontes():
    return Path(__file__).resolve().parent / 'recursos' / 'fontes'


def carregar_fontes():
    """Registra as fontes embutidas só para este processo (Windows: AddFontResourceEx privado). Antes da janela."""
    global _fontes_carregadas
    if _fontes_carregadas:
        return
    for arq in sorted(pasta_fontes().glob('*.ttf')):
        try:
            ctk.FontManager.load_font(str(arq))
        except Exception:
            pass
    _fontes_carregadas = True


def fonte(tamanho=9, peso='normal', tipo='corpo'):
    """Fonte Tk (tamanho em pontos) – tabelas e gráficos."""
    return (FAMILIAS[tipo], tamanho, peso)


def cfonte(tamanho=12, peso='normal', tipo='corpo'):
    """Fonte CustomTkinter (tamanho em px dos tokens: micro 9 · caption 10 · label 11 · body 12 · title 14 ·
    metric 28 · display 40)."""
    return ctk.CTkFont(family=FAMILIAS[tipo], size=tamanho, weight=peso)


def caixa_alta(texto):
    """textTransform: uppercase (títulos, rótulos e botões)."""
    return texto.upper()


def aplicar(raiz):
    global FAMILIA, escala
    disponiveis = set(tkfont.families(raiz))
    for tipo, opcoes in PREFERIDAS.items():
        FAMILIAS[tipo] = next((f for f in opcoes if f in disponiveis), 'TkFixedFont' if tipo in ('corpo', 'dado')
                              else 'TkDefaultFont')
    FAMILIA = FAMILIAS['corpo']
    try:
        escala = ctk.ScalingTracker.get_window_scaling(raiz)
    except Exception:
        escala = 1.0
    for nome in ('TkDefaultFont', 'TkTextFont', 'TkMenuFont'):
        try:
            tkfont.nametofont(nome).configure(family=FAMILIAS['corpo'], size=9)
        except tk.TclError:
            pass
    s = ttk.Style(raiz)
    s.theme_use('clam')
    s.configure('Treeview', background=BASE, fieldbackground=BASE, foreground=TEXTO, rowheight=int(24 * escala),
                borderwidth=0, relief='flat', font=fonte(9))
    s.map('Treeview', background=[('selected', SELECAO)], foreground=[('selected', TEXTO)])
    # cabeçalho = bloco preto com texto branco em caixa-alta
    s.configure('Treeview.Heading', background=INVERSO, foreground=SOBRE_INVERSO, font=fonte(11, 'bold', 'titulo'),
                relief='flat', borderwidth=0, padding=(6, 5))
    s.map('Treeview.Heading', background=[('active', TEXTO_2)], foreground=[('active', SOBRE_INVERSO)])
    s.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
    raiz.option_add('*Menu.background', BASE)
    raiz.option_add('*Menu.foreground', TEXTO)
    raiz.option_add('*Menu.activeBackground', INVERSO)
    raiz.option_add('*Menu.activeForeground', SOBRE_INVERSO)
    raiz.option_add('*Menu.relief', 'solid')
    raiz.option_add('*Menu.borderWidth', 1)
    raiz.option_add('*Menu.font', fonte(9))
    return s
