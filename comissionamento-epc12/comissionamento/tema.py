"""Tema visual: paleta clara e suave, tipografia e estilos ttk.

Cores de série (gráficos) = 3 primeiras cores da paleta categórica de referência, validadas para daltonismo em
qualquer par (azul, laranja, água). Cores de estado são reservadas e sempre acompanham ícone + texto."""
import tkinter as tk
from tkinter import font as tkfont, ttk

# superfícies e texto
FUNDO = '#F3F5F8'          # plano da janela
SUPERFICIE = '#FFFFFF'     # cartões, tabelas, gráficos
LATERAL = '#EAEEF3'        # barra de navegação
LATERAL_ATIVO = '#D9E4F2'
BORDA = '#E1E5EA'
GRADE = '#ECEEF1'          # linhas de grade dos gráficos
EIXO = '#C9CED6'
TEXTO = '#2B2F36'
TEXTO_2 = '#5F6670'
TEXTO_MUDO = '#8A9099'
DESTAQUE = '#2A78D6'
DESTAQUE_ESCURO = '#1C5CAB'
LISTRA = '#F8FAFC'
SELECAO = '#DCE6F2'
TRILHA = '#E6ECF4'         # fundo das barras de progresso

# séries dos gráficos (ordem fixa)
SERIE = ('#2A78D6', '#EB6834', '#1BAF7A')
NEUTRO = '#C9CED6'         # categoria "N/A"/sem aplicação

# estados (reservados)
OK = '#2E7D32'
ALERTA_FUNDO = '#FFF4D6'
ERRO_FUNDO = '#FBE9E5'
ERRO = '#C0392B'
AMARELO_CRITERIO = '#FFF6DA'

FAMILIA = 'Segoe UI'


def fonte(tamanho=10, peso='normal'):
    return (FAMILIA, tamanho, peso)


def aplicar(raiz):
    """Configura fontes e estilos ttk (tema 'clam', que aceita cores personalizadas)."""
    global FAMILIA
    familias = set(tkfont.families(raiz))
    FAMILIA = next((f for f in ('Segoe UI', 'Inter', 'Noto Sans', 'DejaVu Sans', 'Helvetica') if f in familias), 'TkDefaultFont')
    for nome in ('TkDefaultFont', 'TkTextFont', 'TkMenuFont', 'TkHeadingFont'):
        try:
            tkfont.nametofont(nome).configure(family=FAMILIA, size=10)
        except tk.TclError:
            pass
    raiz.configure(background=FUNDO)
    s = ttk.Style(raiz)
    s.theme_use('clam')
    s.configure('.', background=FUNDO, foreground=TEXTO, font=fonte(10), bordercolor=BORDA,
                lightcolor=BORDA, darkcolor=BORDA, troughcolor=TRILHA, focuscolor=DESTAQUE)
    s.configure('TFrame', background=FUNDO)
    s.configure('Card.TFrame', background=SUPERFICIE, relief='flat')
    s.configure('Lateral.TFrame', background=LATERAL)
    s.configure('Topo.TFrame', background=SUPERFICIE)
    s.configure('TLabel', background=FUNDO, foreground=TEXTO)
    s.configure('Card.TLabel', background=SUPERFICIE)
    s.configure('Titulo.TLabel', font=fonte(16, 'bold'), foreground=TEXTO)
    s.configure('Pagina.TLabel', font=fonte(14, 'bold'), foreground=TEXTO)
    s.configure('Sub.TLabel', font=fonte(9), foreground=TEXTO_2)
    s.configure('Nota.TLabel', font=fonte(8), foreground=TEXTO_MUDO)
    s.configure('Topo.TLabel', background=SUPERFICIE)
    s.configure('TopoSub.TLabel', background=SUPERFICIE, font=fonte(9), foreground=TEXTO_2)
    s.configure('TopoTitulo.TLabel', background=SUPERFICIE, font=fonte(15, 'bold'), foreground=TEXTO)
    s.configure('Marca.TLabel', background=LATERAL, font=fonte(11, 'bold'), foreground=DESTAQUE_ESCURO)
    s.configure('LateralNota.TLabel', background=LATERAL, font=fonte(8), foreground=TEXTO_MUDO)

    s.configure('TButton', background=SUPERFICIE, foreground=TEXTO, padding=(12, 6), relief='flat', borderwidth=1)
    s.map('TButton', background=[('active', SELECAO), ('pressed', SELECAO)])
    s.configure('Destaque.TButton', background=DESTAQUE, foreground='#FFFFFF', bordercolor=DESTAQUE)
    s.map('Destaque.TButton', background=[('active', DESTAQUE_ESCURO), ('pressed', DESTAQUE_ESCURO)])
    s.configure('Nav.TButton', background=LATERAL, foreground=TEXTO_2, anchor='w', padding=(16, 9),
                borderwidth=0, font=fonte(10))
    s.map('Nav.TButton', background=[('active', LATERAL_ATIVO)])
    s.configure('NavAtivo.TButton', background=LATERAL_ATIVO, foreground=DESTAQUE_ESCURO, anchor='w', padding=(16, 9),
                borderwidth=0, font=fonte(10, 'bold'))
    s.map('NavAtivo.TButton', background=[('active', LATERAL_ATIVO)])
    s.configure('Seg.TRadiobutton', background=SUPERFICIE, foreground=TEXTO_2, padding=(12, 5), indicatoron=False,
                borderwidth=1, relief='flat')
    s.map('Seg.TRadiobutton', background=[('selected', SELECAO), ('active', LISTRA)],
          foreground=[('selected', DESTAQUE_ESCURO)])
    s.layout('Seg.TRadiobutton', [('Radiobutton.padding', {'sticky': 'nswe', 'children': [
        ('Radiobutton.label', {'sticky': 'nswe'})]})])
    s.configure('TRadiobutton', background=FUNDO)
    s.configure('Card.TRadiobutton', background=SUPERFICIE)
    s.configure('TMenubutton', background=SUPERFICIE, padding=(12, 6), relief='flat', borderwidth=1)
    s.configure('Destaque.TMenubutton', background=DESTAQUE, foreground='#FFFFFF', arrowcolor='#FFFFFF')
    s.map('Destaque.TMenubutton', background=[('active', DESTAQUE_ESCURO)])
    s.configure('TEntry', fieldbackground=SUPERFICIE, padding=5)
    s.configure('TCombobox', fieldbackground=SUPERFICIE, background=SUPERFICIE, padding=4, arrowcolor=TEXTO_2)
    s.map('TCombobox', fieldbackground=[('readonly', SUPERFICIE)], selectbackground=[('readonly', SUPERFICIE)],
          selectforeground=[('readonly', TEXTO)])
    s.configure('TLabelframe', background=SUPERFICIE, bordercolor=BORDA)
    s.configure('TLabelframe.Label', background=SUPERFICIE, foreground=TEXTO_2, font=fonte(9, 'bold'))

    s.configure('Treeview', background=SUPERFICIE, fieldbackground=SUPERFICIE, foreground=TEXTO, rowheight=26,
                borderwidth=0, font=fonte(9))
    s.map('Treeview', background=[('selected', SELECAO)], foreground=[('selected', TEXTO)])
    s.configure('Treeview.Heading', background='#F1F4F8', foreground=TEXTO_2, font=fonte(9, 'bold'),
                relief='flat', borderwidth=0, padding=(6, 6))
    s.map('Treeview.Heading', background=[('active', SELECAO)])
    s.configure('Vertical.TScrollbar', background=FUNDO, troughcolor=SUPERFICIE, arrowcolor=TEXTO_MUDO, borderwidth=0)
    s.configure('Horizontal.TScrollbar', background=FUNDO, troughcolor=SUPERFICIE, arrowcolor=TEXTO_MUDO, borderwidth=0)
    raiz.option_add('*Menu.background', SUPERFICIE)
    raiz.option_add('*Menu.foreground', TEXTO)
    raiz.option_add('*Menu.activeBackground', SELECAO)
    raiz.option_add('*Menu.activeForeground', TEXTO)
    raiz.option_add('*Menu.relief', 'flat')
    raiz.option_add('*TCombobox*Listbox.background', SUPERFICIE)
    raiz.option_add('*TCombobox*Listbox.selectBackground', SELECAO)
    raiz.option_add('*TCombobox*Listbox.selectForeground', TEXTO)
    return s
