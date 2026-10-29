"""Tema visual: claro e suave, cartões arredondados (CustomTkinter) e tabelas ttk no mesmo estilo.

Cores de série (gráficos) = 3 primeiras cores da paleta categórica de referência, validadas para daltonismo em
qualquer par (azul, laranja, água). Cores de estado são reservadas e sempre acompanham texto."""
import tkinter as tk
from tkinter import font as tkfont, ttk

import customtkinter as ctk

# superfícies e texto
FUNDO = '#F4F6FA'          # plano da janela
SUPERFICIE = '#FFFFFF'     # cartões, tabelas, gráficos
BORDA = '#E6EAF0'
BORDA_FORTE = '#D5DBE5'
GRADE = '#EEF1F5'
EIXO = '#CDD3DC'
TEXTO = '#1F2633'
TEXTO_2 = '#566070'
TEXTO_MUDO = '#8A93A1'
DESTAQUE = '#2A78D6'
DESTAQUE_ESCURO = '#1C5CAB'
DESTAQUE_CLARO = '#EAF2FD'
PASSAR = '#F1F4F9'         # hover
LISTRA = '#FAFBFD'
SELECAO = '#E3EDFB'
TRILHA = '#E8EDF5'

SERIE = ('#2A78D6', '#EB6834', '#1BAF7A')
NEUTRO = '#CDD3DC'

OK = '#1E7B34'
OK_FUNDO = '#E7F5EA'
ALERTA = '#9A6700'
ALERTA_FUNDO = '#FFF4D6'
ERRO = '#B4372B'
ERRO_FUNDO = '#FBE9E5'
AMARELO_CRITERIO = '#FFF6DA'

FAMILIA = 'Segoe UI'
RAIO = 14
escala = 1.0


def fonte(tamanho=10, peso='normal'):
    return (FAMILIA, tamanho, peso)


def cfonte(tamanho=13, peso='normal'):
    """Fonte para widgets CustomTkinter (tamanho em px lógicos)."""
    return ctk.CTkFont(family=FAMILIA, size=tamanho, weight=peso)


def aplicar(raiz):
    global FAMILIA, escala
    familias = set(tkfont.families(raiz))
    FAMILIA = next((f for f in ('Segoe UI Variable Text', 'Segoe UI', 'Inter', 'Noto Sans', 'DejaVu Sans')
                    if f in familias), 'TkDefaultFont')
    try:
        escala = ctk.ScalingTracker.get_window_scaling(raiz)
    except Exception:
        escala = 1.0
    for nome in ('TkDefaultFont', 'TkTextFont', 'TkMenuFont', 'TkHeadingFont'):
        try:
            tkfont.nametofont(nome).configure(family=FAMILIA, size=10)
        except tk.TclError:
            pass
    s = ttk.Style(raiz)
    s.theme_use('clam')
    s.configure('Treeview', background=SUPERFICIE, fieldbackground=SUPERFICIE, foreground=TEXTO,
                rowheight=int(30 * escala), borderwidth=0, relief='flat', font=fonte(10))
    s.map('Treeview', background=[('selected', SELECAO)], foreground=[('selected', TEXTO)])
    s.configure('Treeview.Heading', background=SUPERFICIE, foreground=TEXTO_2, font=fonte(9, 'bold'),
                relief='flat', borderwidth=0, padding=(8, 9))
    s.map('Treeview.Heading', background=[('active', PASSAR)])
    s.layout('Treeview', [('Treeview.treearea', {'sticky': 'nswe'})])
    raiz.option_add('*Menu.background', SUPERFICIE)
    raiz.option_add('*Menu.foreground', TEXTO)
    raiz.option_add('*Menu.activeBackground', SELECAO)
    raiz.option_add('*Menu.activeForeground', TEXTO)
    raiz.option_add('*Menu.relief', 'flat')
    raiz.option_add('*Menu.font', fonte(10))
    return s
