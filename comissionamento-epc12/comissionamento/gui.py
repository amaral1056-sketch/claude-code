"""Interface gráfica (CustomTkinter) do Controle Gerencial de Comissionamento – EPC 12."""
import datetime
import os
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

import customtkinter as ctk
from tkinter import font as tkfont

from . import __version__, configuracao, exportacao, rotulos, servico, tema
from .banco import Banco
from .graficos import Grafico, pct_txt
from .icones import icone, logo, recurso
from .importacao import ErroImportacao
from .padroes import ETAPAS

TITULO = 'Controle Gerencial de Comissionamento – EPC 12'
XLSX = [('Planilhas Excel', '*.xlsx *.xlsm'), ('Todos os arquivos', '*.*')]

# página -> (ícone, subtítulo)
PAGINAS = {
    'Visão Gerencial': ('visao', 'Avanço do comissionamento por indicador e por unidade'),
    'FVI': ('fvi', 'Acompanhamento TAG a TAG dos Campos II, III e V da FVI'),
    'FVM': ('fvm', 'Acompanhamento malha a malha dos Campos III e V da FVM'),
    'CCM': ('ccm', 'Acompanhamento do CCM por SSOP'),
    'TAP': ('tap', 'Acompanhamento do TAP por SSOP'),
    'TTAS': ('ttas', 'Acompanhamento do TTAS 1 e do TTAS 2 por SOP'),
    'Verificação': ('verificacao', 'Testes das importações e pendências encontradas linha a linha'),
    'Histórico': ('historico', 'Evolução do avanço em cada data de referência'),
    'Importações': ('importacoes', 'Registro de todas as importações gravadas neste banco de dados'),
    'Critérios': ('criterios', 'Quais códigos de status da FIC contam como etapa aprovada'),
}
SECOES = [('ACOMPANHAMENTO', ('Visão Gerencial', 'FVI', 'FVM', 'CCM', 'TAP', 'TTAS')),
          ('CONTROLE', ('Verificação', 'Histórico', 'Importações')),
          ('CONFIGURAÇÃO', ('Critérios',))]
TODAS_UNIDADES = 'Todas as unidades'
TODOS_SOP = 'Todos os SOP'
TODAS_SSOP = 'Todas as SSOP'
TODAS_SITUACOES = 'Todas as situações'
R_SSOP, R_SOP = 'ssop', 'sop'


def fmt(v, percentual):
    if v is None:
        return ''
    if percentual:
        return pct_txt(v)
    if isinstance(v, float):
        return f'{v:.2f}'.replace('.', ',')
    return str(v)


def numero(n):
    return f'{n:,}'.replace(',', '.')


# ====================================================================== componentes
class Cartao(ctk.CTkFrame):
    def __init__(self, master, **kw):
        kw.setdefault('fg_color', tema.SUPERFICIE)
        kw.setdefault('corner_radius', tema.RAIO)
        kw.setdefault('border_width', 1)
        kw.setdefault('border_color', tema.BORDA)
        super().__init__(master, **kw)


def titulo_secao(master, texto, sub=''):
    f = ctk.CTkFrame(master, fg_color='transparent')
    ctk.CTkLabel(f, text=texto, font=tema.cfonte(16, 'bold'), text_color=tema.TEXTO).pack(anchor='w')
    if sub:
        ctk.CTkLabel(f, text=sub, font=tema.cfonte(12), text_color=tema.TEXTO_MUDO).pack(anchor='w')
    return f


def nota(master, texto, largura=1100):
    return ctk.CTkLabel(master, text=texto, font=tema.cfonte(11), text_color=tema.TEXTO_MUDO, justify='left',
                        wraplength=largura, anchor='w')


def estilo_opcoes():
    return dict(height=36, corner_radius=10, fg_color=tema.FUNDO, button_color=tema.FUNDO,
                button_hover_color=tema.BORDA, text_color=tema.TEXTO, font=tema.cfonte(12),
                dropdown_font=tema.cfonte(12), dynamic_resizing=False)


def botao_sec(master, texto, ic, comando, **kw):
    """Botão secundário (transparente) com ícone padronizado."""
    return ctk.CTkButton(master, text=texto, image=icone(ic, tema.TEXTO_2), height=kw.pop('height', 36),
                         corner_radius=10, fg_color='transparent', hover_color=tema.PASSAR, text_color=tema.TEXTO_2,
                         font=tema.cfonte(12), command=comando, **kw)


def botao_pri(master, texto, ic, comando, **kw):
    """Botão principal (azul) com ícone padronizado."""
    return ctk.CTkButton(master, text=texto, image=icone(ic, '#FFFFFF'), height=kw.pop('height', 36),
                         corner_radius=10, fg_color=tema.DESTAQUE, hover_color=tema.DESTAQUE_ESCURO,
                         text_color='#FFFFFF', font=tema.cfonte(12, 'bold'), command=comando, **kw)


class Tabela(Cartao):
    """Tabela (ttk.Treeview) dentro de um cartão: linhas listradas, cabeçalho clicável para ordenar."""

    def __init__(self, master, colunas, altura=18, ao_duplo_clique=None):
        super().__init__(master)
        self.colunas = colunas
        ids = [k for _, k, _, _ in colunas]
        self.tree = ttk.Treeview(self, columns=ids, show='headings', height=altura, selectmode='extended')
        self._f_corpo = tkfont.Font(family=tema.FAMILIA, size=10)
        self._f_cab = tkfont.Font(family=tema.FAMILIA, size=9, weight='bold')
        for titulo, k, larg, p in colunas:
            self.tree.heading(k, text=titulo, anchor='center' if p == 2 else 'w', command=lambda k=k: self.ordenar(k))
            self.tree.column(k, width=self._f_cab.measure(titulo) + 30, stretch=True)
        sy = ctk.CTkScrollbar(self, orientation='vertical', command=self.tree.yview, button_color=tema.BORDA_FORTE,
                              button_hover_color=tema.EIXO, width=12)
        sx = ctk.CTkScrollbar(self, orientation='horizontal', command=self.tree.xview, button_color=tema.BORDA_FORTE,
                              button_hover_color=tema.EIXO, height=12)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.grid(row=0, column=0, sticky='nsew', padx=(10, 0), pady=(8, 0))
        sy.grid(row=0, column=1, sticky='ns', padx=(2, 6), pady=(10, 0))
        sx.grid(row=1, column=0, sticky='ew', padx=(10, 0), pady=(2, 8))
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        for tag, kw in (('par', {'background': tema.LISTRA}), ('amarelo', {'background': tema.AMARELO_CRITERIO}),
                        ('total', {'background': tema.DESTAQUE_CLARO, 'font': tema.fonte(10, 'bold')}),
                        ('alerta', {'background': tema.ERRO_FUNDO, 'foreground': tema.ERRO}),
                        ('aviso', {'background': tema.ALERTA_FUNDO}), ('mudo', {'foreground': tema.TEXTO_MUDO}),
                        ('ok', {'foreground': tema.OK})):
            self.tree.tag_configure(tag, **kw)
        if ao_duplo_clique:
            self.tree.bind('<Double-1>', lambda ev: self.selecionadas() and ao_duplo_clique(self.selecionadas()[0]))
        self.linhas, self._tags, self._ordem = [], None, (None, False)

    def preencher(self, linhas, tags=None):
        self.linhas = list(linhas)
        self._tags = tags
        self._mostrar()

    def _mostrar(self):
        amostra = self.linhas[:250]
        for titulo, k, larg, p in self.colunas:
            vals = [r.get(k) for r in amostra if r.get(k) not in (None, '')]
            # coluna só de números fica à direita, como no Excel; Sim/Não centralizado
            numerica = p == 1 or (vals and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in vals))
            # largura pelo conteúdo real (medido com a fonte da tabela), limitada a 'larg' caracteres
            conteudo = max([self._f_corpo.measure(fmt(v, p == 1)) for v in vals] or [0]) + 20
            limite = self._f_corpo.measure('n' * larg) + 20
            largura = max(self._f_cab.measure(titulo) + 26, min(conteudo, limite), 52)
            self.tree.column(k, anchor='center' if p == 2 else ('e' if numerica else 'w'), width=largura)
        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(self.linhas):
            tags = tuple(self._tags(r)) if self._tags else ()
            self.tree.insert('', 'end', iid=str(i), values=[fmt(r.get(k), p == 1) for _, k, _, p in self.colunas],
                             tags=tags or (('par',) if i % 2 else ()))

    def ordenar(self, k):
        atual, inv = self._ordem
        inv = not inv if atual == k else False
        self._ordem = (k, inv)
        fixas = [r for r in self.linhas if r.get('unidade') == 'TOTAL']
        resto = [r for r in self.linhas if r.get('unidade') != 'TOTAL']
        resto.sort(key=lambda r: (r.get(k) is None, r.get(k).upper() if isinstance(r.get(k), str) else r.get(k)),
                   reverse=inv)
        self.linhas = resto + fixas
        self._mostrar()

    def selecionadas(self):
        return [self.linhas[int(i)] for i in self.tree.selection()]


class CartaoGrafico(Cartao):
    def __init__(self, master, titulo, subtitulo, altura=290):
        super().__init__(master)
        self.g = Grafico(self, titulo, subtitulo, altura)
        self.g.pack(fill='both', expand=True, padx=18, pady=16)


# ====================================================================== aplicação
class App(ctk.CTk):
    def __init__(self, caminho_banco=None):
        ctk.set_appearance_mode('light')
        super().__init__(fg_color=tema.FUNDO)
        self.title(TITULO)
        self.geometry('1440x900')
        self.minsize(1180, 700)
        tema.aplicar(self)
        self._icone_janela()
        self.banco = None
        self.R = None
        self._agendado = None
        self._montar()
        self.protocol('WM_DELETE_WINDOW', self.sair)
        self.abrir_banco(caminho_banco or configuracao.banco_padrao())

    def _icone_janela(self):
        try:
            if os.name == 'nt':
                self.iconbitmap(str(recurso('icone.ico')))
            self._img_icone = tk.PhotoImage(file=str(recurso('icone.png')))
            self.iconphoto(True, self._img_icone)
        except tk.TclError:
            pass

    # ------------------------------------------------------------------ estrutura
    def _menu(self, itens):
        m = tk.Menu(self, tearoff=0, bd=0)
        for item in itens:
            if item is None:
                m.add_separator()
            else:
                m.add_command(label=item[0], command=item[1])
        return m

    def _abrir_menu(self, botao, menu):
        menu.tk_popup(botao.winfo_rootx(), botao.winfo_rooty() + botao.winfo_height() + 4)

    def menu_importar(self):
        return self._menu([('SSOP', self.importar_li)]
                          + [(t, lambda t=t: self.importar_fic(t)) for t in servico.TIPOS_FIC])

    def menu_banco(self):
        return self._menu([('Abrir ou criar banco de dados…', self.escolher_banco),
                           ('Salvar cópia de segurança do banco…', self.backup), None,
                           ('Exportar planilhas-modelo (FVM, CCM, TAP, TTAS)…', self.exportar_modelos)])

    def _montar(self):
        lateral = ctk.CTkFrame(self, width=252, fg_color=tema.SUPERFICIE, corner_radius=0)
        lateral.pack(side='left', fill='y')
        lateral.pack_propagate(False)
        ctk.CTkFrame(self, width=1, fg_color=tema.BORDA, corner_radius=0).pack(side='left', fill='y')

        marca = ctk.CTkFrame(lateral, fg_color='transparent')
        marca.pack(fill='x', padx=22, pady=(22, 16))
        img = logo([configuracao.pasta_programa(), configuracao.pasta_programa() / 'dados'], altura=64)
        self.logo = ctk.CTkLabel(marca, text='', image=img) if img else ctk.CTkLabel(
            marca, text='LOGO', width=64, height=64, corner_radius=12, fg_color=tema.FUNDO,
            text_color=tema.TEXTO_MUDO, font=tema.cfonte(9, 'bold'))
        self.logo.pack(anchor='w')
        ctk.CTkLabel(marca, text='Controle Gerencial de Comissionamento', font=tema.cfonte(12, 'bold'),
                     text_color=tema.TEXTO, wraplength=200, justify='left').pack(anchor='w', pady=(10, 0))
        ctk.CTkLabel(marca, text='EPC 12', font=tema.cfonte(11), text_color=tema.TEXTO_MUDO).pack(anchor='w')

        self.nav = {}
        for secao, paginas in SECOES:
            ctk.CTkLabel(lateral, text=secao, font=tema.cfonte(10, 'bold'), text_color=tema.TEXTO_MUDO).pack(
                anchor='w', padx=24, pady=(10, 4))
            for nome in paginas:
                b = ctk.CTkButton(lateral, text=nome, image=icone(PAGINAS[nome][0], tema.TEXTO_2), anchor='w',
                                  compound='left', height=40, corner_radius=10, fg_color='transparent',
                                  hover_color=tema.PASSAR, text_color=tema.TEXTO_2, font=tema.cfonte(13),
                                  command=lambda n=nome: self.mostrar(n))
                b.pack(fill='x', padx=14, pady=1)
                self.nav[nome] = b
        self.selo_ver = ctk.CTkLabel(self.nav['Verificação'], text='', height=22, corner_radius=11,
                                     fg_color=tema.ERRO_FUNDO, text_color=tema.ERRO, font=tema.cfonte(11, 'bold'))

        rod = ctk.CTkFrame(lateral, fg_color='transparent')
        rod.pack(side='bottom', fill='x', padx=14, pady=(0, 16))
        self.btn_banco = botao_sec(rod, 'Banco de dados', 'banco', None, anchor='w', height=38)
        self.btn_banco.configure(command=lambda: self._abrir_menu(self.btn_banco, self.menu_banco()))
        self.btn_banco.pack(fill='x')
        self.btn_ajuda = botao_sec(rod, 'Ajuda', 'ajuda', None, anchor='w', height=38)
        self.btn_ajuda.configure(command=lambda: self._abrir_menu(self.btn_ajuda, self._menu([
            ('Como usar', self.ajuda), ('Sobre', self.sobre)])))
        self.btn_ajuda.pack(fill='x')
        self.lbl_banco = ctk.CTkLabel(rod, text='', font=tema.cfonte(10), text_color=tema.TEXTO_MUDO, anchor='w',
                                      justify='left', wraplength=210)
        self.lbl_banco.pack(fill='x', padx=10, pady=(6, 0))

        # área principal
        principal = ctk.CTkFrame(self, fg_color='transparent')
        principal.pack(side='left', fill='both', expand=True)
        topo = ctk.CTkFrame(principal, fg_color='transparent')
        topo.pack(fill='x', padx=26, pady=(24, 12))
        txt = ctk.CTkFrame(topo, fg_color='transparent')
        txt.pack(side='left')
        self.lbl_pagina = ctk.CTkLabel(txt, text='', font=tema.cfonte(26, 'bold'), text_color=tema.TEXTO)
        self.lbl_pagina.pack(anchor='w')
        self.lbl_sub = ctk.CTkLabel(txt, text='', font=tema.cfonte(13), text_color=tema.TEXTO_MUDO)
        self.lbl_sub.pack(anchor='w')
        self.btn_exportar = ctk.CTkButton(topo, text='Exportar Excel', image=icone('exportar', tema.TEXTO_2), height=40,
                                          corner_radius=10, fg_color=tema.SUPERFICIE, hover_color=tema.PASSAR,
                                          text_color=tema.TEXTO, border_width=1, border_color=tema.BORDA_FORTE,
                                          font=tema.cfonte(13), command=self.exportar)
        self.btn_exportar.pack(side='right')
        self.btn_importar = botao_pri(topo, 'Importar planilha', 'importar', None, height=40)
        self.btn_importar.configure(command=lambda: self._abrir_menu(self.btn_importar, self.menu_importar()))
        self.btn_importar.pack(side='right', padx=10)
        self.chip_data = ctk.CTkLabel(topo, text='', image=icone('calendario', tema.TEXTO_2), compound='left',
                                      height=40, corner_radius=10, fg_color=tema.SUPERFICIE, text_color=tema.TEXTO_2,
                                      font=tema.cfonte(12), padx=14)
        self.chip_data.pack(side='right')

        self.corpo = ctk.CTkFrame(principal, fg_color='transparent')
        self.corpo.pack(fill='both', expand=True, padx=26, pady=(4, 0))
        self.status = ctk.CTkLabel(principal, text='', font=tema.cfonte(11), text_color=tema.TEXTO_MUDO, anchor='w')
        self.status.pack(fill='x', padx=34, pady=(4, 10))

        self.paginas = {}
        for nome, montar in (('Visão Gerencial', self._pg_visao), ('FVI', self._pg_fvi), ('FVM', self._pg_fvm),
                             ('CCM', lambda f: self._pg_doc(f, 'CCM', rotulos.COL_CCM, R_SSOP, 'txt_ccm_situacao',
                                                            rotulos.SITUACOES_CCM)),
                             ('TAP', lambda f: self._pg_doc(f, 'TAP', rotulos.COL_TAP, R_SSOP, 'txt_tap_situacao',
                                                            rotulos.SITUACOES_TAP)),
                             ('TTAS', lambda f: self._pg_doc(f, 'TTAS', rotulos.COL_TTAS, R_SOP, 'txt_ttas_situacao',
                                                             rotulos.SITUACOES_TTAS)),
                             ('Verificação', self._pg_verificacao), ('Histórico', self._pg_historico),
                             ('Importações', self._pg_importacoes), ('Critérios', self._pg_criterios)):
            f = ctk.CTkFrame(self.corpo, fg_color='transparent')
            montar(f)
            self.paginas[nome] = f
        self.pagina_atual = None
        self.mostrar('Visão Gerencial')

    def mostrar(self, nome):
        if self.pagina_atual:
            self.paginas[self.pagina_atual].pack_forget()
            self.nav[self.pagina_atual].configure(fg_color='transparent', text_color=tema.TEXTO_2, font=tema.cfonte(13),
                                                  image=icone(PAGINAS[self.pagina_atual][0], tema.TEXTO_2))
        self.paginas[nome].pack(fill='both', expand=True)
        self.nav[nome].configure(fg_color=tema.DESTAQUE_CLARO, text_color=tema.DESTAQUE_ESCURO,
                                 image=icone(PAGINAS[nome][0], tema.DESTAQUE), font=tema.cfonte(13, 'bold'))
        self.lbl_pagina.configure(text=nome)
        self.lbl_sub.configure(text=PAGINAS[nome][1])
        self.pagina_atual = nome

    # ------------------------------------------------------------------ páginas
    def _pg_visao(self, f):
        sc = self.visao_rolagem = ctk.CTkScrollableFrame(f, fg_color='transparent',
                                                         scrollbar_button_color=tema.BORDA_FORTE)
        sc.pack(fill='both', expand=True)
        graf = ctk.CTkFrame(sc, fg_color='transparent')
        graf.pack(fill='x')
        self.g = {}
        defs = [('fvi', 'FVI', 'TAGs aprovadas ÷ TAGs aplicáveis, por campo, em cada unidade e no total', 0, 2),
                ('fvm', 'FVM', 'Malhas aprovadas ÷ malhas válidas, por campo, em cada unidade e no total', 1, 0),
                ('ttas', 'TTAS', 'SOP com TTAS 1 e TTAS 2 aprovados ÷ SOP, em cada unidade e no total', 1, 1),
                ('ccm', 'CCM', 'SSOP com CCM aprovado ÷ SSOP, em cada unidade e no total', 2, 0),
                ('tap', 'TAP', 'SSOP com TAP concluído ÷ SSOP, em cada unidade e no total', 2, 1)]
        for k, t, sub, lin, col in defs:
            c = CartaoGrafico(graf, t, sub, altura=340)
            larga = col == 2
            c.grid(row=lin, column=0 if larga else col, columnspan=2 if larga else 1, sticky='nsew',
                   padx=(0 if larga or col == 0 else 14, 0), pady=(0, 14))
            self.g[k] = c.g
        graf.columnconfigure(0, weight=1, uniform='g')
        graf.columnconfigure(1, weight=1, uniform='g')

    def _filtros(self, f, d, situacoes=None):
        """Unidade → SOP → SSOP em cascata; busca por TAG/malha (FVI/FVM) ou situação (CCM/TAP/TTAS)."""
        barra = Cartao(f)
        barra.pack(fill='x', pady=(0, 12))
        dentro = ctk.CTkFrame(barra, fg_color='transparent')
        dentro.pack(fill='x', padx=14, pady=12)
        d['contagem'] = ctk.CTkLabel(dentro, text='', height=30, corner_radius=15, fg_color=tema.DESTAQUE_CLARO,
                                     text_color=tema.DESTAQUE_ESCURO, font=tema.cfonte(12, 'bold'), padx=12)
        d['contagem'].pack(side='right')
        d['unidade'] = ctk.CTkOptionMenu(dentro, values=[TODAS_UNIDADES], width=170,
                                         command=lambda v: self._mudou_unidade(d), **estilo_opcoes())
        d['unidade'].pack(side='left')
        estilo_cb = dict(height=36, corner_radius=10, border_width=1, border_color=tema.BORDA, fg_color=tema.FUNDO,
                         button_color=tema.FUNDO, button_hover_color=tema.BORDA, text_color=tema.TEXTO,
                         font=tema.cfonte(12), dropdown_font=tema.cfonte(12))
        d['sop'] = ctk.CTkComboBox(dentro, values=[TODOS_SOP], width=210, command=lambda v: self._mudou_sop(d), **estilo_cb)
        d['sop'].set(TODOS_SOP)
        d['sop'].pack(side='left', padx=(10, 0))
        d['ssop'] = ctk.CTkComboBox(dentro, values=[TODAS_SSOP], width=240, command=lambda v: self._preencher(d),
                                    **estilo_cb)
        d['ssop'].set(TODAS_SSOP)
        d['ssop'].pack(side='left', padx=(10, 0))
        for w in (d['sop'], d['ssop']):
            w.bind('<KeyRelease>', lambda ev: self._agendar(d))
        if situacoes:
            d['situacao'] = ctk.CTkOptionMenu(dentro, values=[TODAS_SITUACOES] + list(situacoes), width=190,
                                              command=lambda v: self._preencher(d), **estilo_opcoes())
            d['situacao'].pack(side='left', padx=(10, 0))
        if d.get('ident'):
            d['busca'] = ctk.CTkEntry(dentro, placeholder_text=d['dica_busca'], width=190, height=36, corner_radius=10,
                                      border_width=1, border_color=tema.BORDA, fg_color=tema.FUNDO, font=tema.cfonte(12))
            d['busca'].pack(side='left', padx=(10, 0))
            d['busca'].bind('<KeyRelease>', lambda ev: self._agendar(d))
        botao_sec(dentro, 'Limpar filtros', 'limpar', lambda: self._limpar(d)).pack(side='left', padx=(6, 0))

    def _pg_fvi(self, f):
        d = self.fvi = {'nome': 'FVI', 'ident': 'tag', 'dica_busca': 'Buscar TAG'}
        self._filtros(f, d)
        nota(f, 'Cada campo mostra se conta como aprovado pelos critérios e, entre parênteses, o status da FIC: '
                '"Aprovado", "Pendente" ou "Não se aplica" (fora da conta). Linhas em amarelo: TAG sem FVI gerada · '
                'em vermelho: sem SSOP ou SSOP fora da base · em cinza: TAG cancelada.').pack(side='bottom', fill='x', pady=(8, 0))
        d['tabela'] = Tabela(f, rotulos.COL_TAGS)
        d['tabela'].pack(fill='both', expand=True)

    def _pg_fvm(self, f):
        d = self.fvm = {'nome': 'FVM', 'ident': 'malha', 'dica_busca': 'Buscar malha'}
        self._filtros(f, d)
        nota(f, 'Cada campo mostra se conta como aprovado pelos critérios da FVM (página Critérios) e, entre parênteses, '
                'o status da malha. Layout da planilha de FVM ainda provisório: MALHA | SOP | SSOP | STATUS.').pack(
            side='bottom', fill='x', pady=(8, 0))
        d['tabela'] = Tabela(f, rotulos.COL_MALHAS)
        d['tabela'].pack(fill='both', expand=True)

    def _pg_doc(self, f, nome, colunas, fonte, campo_situacao, situacoes):
        d = {'nome': nome, 'fonte': fonte, 'campo_situacao': campo_situacao}
        self._filtros(f, d, situacoes)
        d['tabela'] = Tabela(f, colunas)
        d['tabela'].pack(fill='both', expand=True)
        setattr(self, 'doc_' + nome.lower(), d)

    def _pg_verificacao(self, f):
        titulo_secao(f, 'Testes das importações', 'Os testes devem resultar OK antes de apresentar os números').pack(
            fill='x', pady=(0, 8))
        self.t_ver = Tabela(f, rotulos.COL_VER, altura=10)
        self.t_ver.pack(fill='both', expand=True)
        titulo_secao(f, 'Pendências encontradas', 'Linha a linha do arquivo importado').pack(fill='x', pady=(14, 8))
        self.t_pend = Tabela(f, rotulos.COL_PEND, altura=8)
        self.t_pend.pack(fill='both', expand=True)

    def _pg_historico(self, f):
        barra = Cartao(f)
        barra.pack(fill='x', pady=(0, 12))
        dentro = ctk.CTkFrame(barra, fg_color='transparent')
        dentro.pack(fill='x', padx=14, pady=12)
        self.cb_hist = ctk.CTkOptionMenu(dentro, values=['TOTAL'], width=200,
                                         command=lambda v: self._preencher_historico(), **estilo_opcoes())
        self.cb_hist.pack(side='left')
        botao_pri(dentro, 'Registrar situação atual em uma data', 'mais', self.registrar_historico).pack(side='left', padx=10)
        botao_sec(dentro, 'Remover data selecionada', 'lixeira', self.remover_historico).pack(side='left')
        c = CartaoGrafico(f, 'Evolução', 'FVI Campo II, FVI Campo III e FVI Campo V', altura=250)
        c.pack(fill='x', pady=(0, 12))
        self.g_hist = c.g
        self.t_hist = Tabela(f, rotulos.COL_HIST, altura=8)
        self.t_hist.pack(fill='both', expand=True)
        nota(f, 'Cada importação grava o avanço na data de referência informada; importar de novo na mesma data substitui '
                'o registro daquela data.').pack(fill='x', pady=(8, 0))

    def _pg_importacoes(self, f):
        self.t_imp = Tabela(f, rotulos.COL_IMP)
        self.t_imp.pack(fill='both', expand=True)

    def _pg_criterios(self, f):
        par = Cartao(f)
        par.pack(fill='x', pady=(0, 12))
        dentro = ctk.CTkFrame(par, fg_color='transparent')
        dentro.pack(fill='x', padx=18, pady=14)
        ctk.CTkLabel(dentro, text='TAGs sem FVI gerada entram no cálculo do FVI?', font=tema.cfonte(13, 'bold'),
                     text_color=tema.TEXTO).pack(side='left')
        self.seg_sem_fvi = ctk.CTkSegmentedButton(dentro, values=['Sim', 'Não'], command=lambda v: self.mudar_parametro(),
                                                  selected_color='#CFE0F8', selected_hover_color='#BBD3F5',
                                                  unselected_color=tema.FUNDO, unselected_hover_color=tema.BORDA,
                                                  fg_color=tema.FUNDO, text_color=tema.TEXTO, font=tema.cfonte(12, 'bold'),
                                                  height=34, width=150, corner_radius=10)
        self.seg_sem_fvi.pack(side='left', padx=14)
        ctk.CTkLabel(dentro, text='Sim = TAG sem FVI conta como pendente (visão conservadora). '
                                  'Não = só TAGs com FVI entram nos percentuais.',
                     font=tema.cfonte(11), text_color=tema.TEXTO_MUDO).pack(side='left')
        corpo = ctk.CTkFrame(f, fg_color='transparent')
        corpo.pack(fill='both', expand=True)
        dir_ = ctk.CTkFrame(corpo, fg_color='transparent')
        dir_.pack(side='right', fill='y', padx=(14, 0))
        cab = ctk.CTkFrame(dir_, fg_color='transparent')
        cab.pack(fill='x', pady=(0, 8))
        titulo_secao(cab, 'Listas oficiais de status da FIC').pack(side='left')
        botao_sec(cab, 'Importar listas', 'importar', self.importar_status).pack(side='right')
        self.t_status = Tabela(dir_, rotulos.COL_STATUS)
        self.t_status.pack(fill='both', expand=True)
        esq = ctk.CTkFrame(corpo, fg_color='transparent')
        esq.pack(side='left', fill='both', expand=True)
        titulo_secao(esq, 'Códigos que contam como etapa aprovada',
                     'Linhas em amarelo: incluídas pela lógica do fluxo da FIC, ainda a confirmar').pack(fill='x', pady=(0, 8))
        form = Cartao(esq)
        form.pack(side='bottom', fill='x', pady=(12, 0))
        fi = ctk.CTkFrame(form, fg_color='transparent')
        fi.pack(fill='x', padx=14, pady=12)
        self.cb_etapa = ctk.CTkOptionMenu(fi, values=list(ETAPAS), width=150, **estilo_opcoes())
        self.cb_etapa.set('Etapa')
        self.cb_etapa.pack(side='left')
        self.cb_codigo = ctk.CTkComboBox(fi, values=[], width=190, height=36, corner_radius=10, border_width=1,
                                         border_color=tema.BORDA, fg_color=tema.FUNDO, button_color=tema.FUNDO,
                                         button_hover_color=tema.BORDA, font=tema.cfonte(12), dropdown_font=tema.cfonte(12))
        self.cb_codigo.set('Código de status')
        self.cb_codigo.pack(side='left', padx=8)
        self.en_origem = ctk.CTkEntry(fi, placeholder_text='Origem / justificativa', height=36, corner_radius=10,
                                      border_width=1, border_color=tema.BORDA, fg_color=tema.FUNDO, font=tema.cfonte(12))
        self.en_origem.pack(side='left', fill='x', expand=True)
        botao_pri(fi, 'Adicionar', 'mais', self.adicionar_criterio, width=120).pack(side='left', padx=(8, 0))
        bt = ctk.CTkFrame(form, fg_color='transparent')
        bt.pack(fill='x', padx=14, pady=(0, 12))
        botao_sec(bt, 'Remover selecionados', 'lixeira', self.remover_criterio).pack(side='left', padx=(0, 6))
        botao_sec(bt, 'Restaurar critérios padrão', 'restaurar', self.restaurar_criterios).pack(side='left')
        self.t_crit = Tabela(esq, rotulos.COL_CRIT, altura=10)
        self.t_crit.pack(fill='both', expand=True)

    # ------------------------------------------------------------------ banco
    def abrir_banco(self, caminho):
        try:
            novo = Banco(caminho)
        except Exception as e:
            messagebox.showerror('Banco de dados', f'Não foi possível abrir {caminho}:\n{e}')
            return
        if self.banco:
            self.banco.fechar()
        self.banco = novo
        configuracao.lembrar_banco(caminho)
        self.atualizar()
        if not self.banco.tem_base():
            self.status_msg('Banco novo: comece por Importar planilha ▸ LI – base de SOP e SSOP.')

    def escolher_banco(self):
        p = filedialog.asksaveasfilename(title='Abrir ou criar banco de dados', defaultextension='.db',
                                         filetypes=[('Banco de dados', '*.db'), ('Todos os arquivos', '*.*')],
                                         initialdir=str(self.banco.caminho.parent) if self.banco else None,
                                         confirmoverwrite=False)
        if p:
            self.abrir_banco(p)

    def backup(self):
        nome = f'{self.banco.caminho.stem}_copia_{datetime.datetime.now():%Y%m%d_%H%M}.db'
        p = filedialog.asksaveasfilename(title='Salvar cópia de segurança', defaultextension='.db', initialfile=nome,
                                         filetypes=[('Banco de dados', '*.db')])
        if p:
            self._executar(lambda: self.banco.copiar_para(p), f'Cópia salva em {p}')

    # ------------------------------------------------------------------ importações
    def _executar(self, acao, msg_ok=None):
        self.configure(cursor='watch')
        self.update_idletasks()
        try:
            r = acao()
        except ErroImportacao as e:
            messagebox.showerror('Importação não realizada', str(e))
            return None
        except Exception as e:
            configuracao.registrar_erro(traceback.format_exc())
            messagebox.showerror('Erro', f'{type(e).__name__}: {e}\n\nDetalhes em {configuracao.arquivo_log()}')
            return None
        finally:
            self.configure(cursor='')
        if msg_ok:
            self.status_msg(msg_ok)
        return r if r is not None else True

    def importar_li(self):
        p = filedialog.askopenfilename(title='Selecione a LI (lista de SOP e SSOP)', filetypes=XLSX)
        if not p:
            return
        if self.banco.tem_base() and not messagebox.askyesno(
                'Substituir base', 'A base de SOP e SSOP atual será substituída pela desta LI. Continuar?'):
            return
        msg = self._executar(lambda: servico.importar_li(self.banco, p))
        if msg:
            self.atualizar()
            self.status_msg(msg)
            messagebox.showinfo('LI importada', msg)

    def importar_fic(self, tipo):
        if not self.banco.tem_base():
            messagebox.showwarning('Base ausente', 'Importe primeiro a LI (Importar planilha ▸ LI – base de SOP e SSOP).')
            return
        p = filedialog.askopenfilename(title=f'Selecione o arquivo de {servico.NOMES[tipo]}', filetypes=XLSX)
        if not p:
            return
        lido = self._executar(lambda: servico.ler_fic(tipo, p))
        if not lido:
            return
        regs, data_rodape, info = lido
        origem = 'data de emissão no rodapé do arquivo' if data_rodape else 'arquivo sem data de emissão – confira'
        sugestao = (data_rodape or datetime.date.today()).strftime('%d/%m/%Y')
        data = simpledialog.askstring(
            'Data de referência',
            f'{len(regs)} registro(s) de {tipo} lidos de {Path(p).name}.\n\n'
            f'Data de referência ({origem}).\nO avanço será registrado no histórico nesta data (dd/mm/aaaa):',
            initialvalue=sugestao, parent=self)
        if not data:
            return
        anterior = self.banco.ultima_importacao(tipo)
        R = self._executar(lambda: servico.gravar_fic(self.banco, tipo, p, regs, data, info))
        if R:
            self.atualizar()
            extra = f'\nSubstituiu a importação de {anterior["importado_em"]}.' if anterior else ''
            ign = info.get('linhas_ignoradas') or []
            if ign:
                extra += f'\n{len(ign)} linha(s) ignorada(s) (cabeçalho repetido ou rodapé): {", ".join(map(str, ign[:10]))}'
            alerta = [t for _, t, _, _, r in self.R.verificacao if r == 'VERIFICAR']
            if alerta:
                extra += '\n\nAtenção – testes para verificar (página Verificação):\n- ' + '\n- '.join(alerta)
            messagebox.showinfo('Importação concluída', f'{len(regs)} registro(s) de {tipo} gravados.{extra}')
            self.status_msg(f'{tipo}: {len(regs)} registros importados de {Path(p).name}')

    def importar_status(self):
        arquivos = filedialog.askopenfilenames(title='Selecione STATUS_FVI.xlsx, STATUS_FVM.xlsx…', filetypes=XLSX)
        feitos = []
        for p in arquivos:
            r = self._executar(lambda p=p: servico.importar_status(self.banco, p))
            if r:
                feitos.append(f'{r[0]}: {r[1]} códigos')
        if feitos:
            self.atualizar()
            messagebox.showinfo('Listas de status', 'Importadas:\n' + '\n'.join(feitos))

    # ------------------------------------------------------------------ exportações
    def exportar(self):
        if not self.R:
            return
        data = servico.data_br(self.banco.data_referencia_atual())
        nome = f'CONTROLE_GERENCIAL_EPC12_{(data or "sem_data").replace("/", "-")}.xlsx'
        p = filedialog.asksaveasfilename(title='Exportar relatório Excel', defaultextension='.xlsx', initialfile=nome,
                                         filetypes=[('Planilha Excel', '*.xlsx')])
        if not p:
            return
        ok = self._executar(lambda: exportacao.exportar_relatorio(
            self.R, p, data, servico.historico_com_percentuais(self.banco), self.banco.criterios(),
            self.banco.incluir_sem_fvi), f'Relatório exportado: {p}')
        if ok and messagebox.askyesno('Relatório exportado', f'Arquivo salvo:\n{p}\n\nAbrir agora?'):
            abrir_arquivo(p)

    def exportar_modelos(self):
        p = filedialog.asksaveasfilename(title='Exportar planilhas-modelo', defaultextension='.xlsx',
                                         initialfile='MODELOS_IMPORTACAO_EPC12.xlsx', filetypes=[('Planilha Excel', '*.xlsx')])
        if p:
            self._executar(lambda: exportacao.exportar_modelos(p), f'Modelos salvos em {p}')

    # ------------------------------------------------------------------ histórico e critérios
    def registrar_historico(self):
        data = simpledialog.askstring('Registrar no histórico', 'Data de referência (dd/mm/aaaa):',
                                      initialvalue=datetime.date.today().strftime('%d/%m/%Y'), parent=self)
        if data and self._executar(lambda: servico.registrar_historico(self.banco, data), f'Histórico registrado em {data}'):
            self.atualizar()

    def remover_historico(self):
        datas = sorted({r['data_referencia'] for r in self.t_hist.selecionadas()})
        if datas and messagebox.askyesno('Remover', 'Remover do histórico as datas: '
                                         + ', '.join(servico.data_br(d) for d in datas) + '?'):
            for d in datas:
                self.banco.remover_historico(d)
            self.atualizar()

    def mudar_parametro(self):
        self.banco.definir_incluir_sem_fvi(self.seg_sem_fvi.get() == 'Sim')
        self.atualizar()
        self.status_msg('Parâmetro alterado. O histórico já gravado não muda; use "Registrar situação atual" se quiser.')

    def adicionar_criterio(self):
        e, c, o = self.cb_etapa.get(), self.cb_codigo.get().strip(), self.en_origem.get().strip()
        if e not in ETAPAS or not c:
            messagebox.showwarning('Critério', 'Escolha a etapa e informe o código de status.')
            return
        if self._executar(lambda: self.banco.adicionar_criterio(e, c, o or 'incluído pelo usuário')):
            self.atualizar()

    def remover_criterio(self):
        sel = self.t_crit.selecionadas()
        if sel and messagebox.askyesno('Remover', f'Remover {len(sel)} critério(s)?'):
            for r in sel:
                self.banco.remover_criterio(r['id'])
            self.atualizar()

    def restaurar_criterios(self):
        if messagebox.askyesno('Restaurar', 'Voltar todos os critérios ao padrão (as alterações feitas serão perdidas)?'):
            self.banco.restaurar_criterios()
            self.atualizar()

    # ------------------------------------------------------------------ atualização
    def atualizar(self):
        b = self.banco
        try:
            self.R = rotulos.preparar(servico.calcular(b))
        except Exception:
            configuracao.registrar_erro(traceback.format_exc())
            raise
        R = self.R
        data = servico.data_br(b.data_referencia_atual())
        self.chip_data.configure(text=f'Data de referência: {data or "sem importação"}')
        self.lbl_banco.configure(text=f'{b.caminho.name}\nversão {__version__}')
        self.title(f'{TITULO} – {b.caminho.name}')
        unidades = [TODAS_UNIDADES] + R.unidades
        self._sop_de_ssop = {s['ssop']: s['sop'] for s in R.ssop}
        for d in self._detalhes():
            d['unidade'].configure(values=unidades)
            if d['unidade'].get() not in unidades:
                d['unidade'].set(TODAS_UNIDADES)
            self._opcoes_sop(d)
            self._preencher(d)
        res = {'OK': 'OK', 'VERIFICAR': 'Verificar', 'INFO': 'Informação'}
        self.t_ver.preencher([{'g': g, 't': t, 'e': '-' if e is None else e, 'a': a, 'r': res[r]}
                              for g, t, e, a, r in R.verificacao],
                             lambda r: ('alerta',) if r['r'] == 'Verificar' else (('ok',) if r['r'] == 'OK' else ()))
        self.t_pend.preencher([dict(zip('dlisp', p)) for p in R.pendencias])
        self.cb_hist.configure(values=['TOTAL'] + R.unidades)
        if self.cb_hist.get() not in ['TOTAL'] + R.unidades:
            self.cb_hist.set('TOTAL')
        self._preencher_historico()
        self.seg_sem_fvi.set('Sim' if b.incluir_sem_fvi else 'Não')
        self.t_crit.preencher([{'id': i, 'etapa': e, 'codigo': c, 'origem': o} for i, e, c, o in b.criterios()],
                              lambda r: ('amarelo',) if str(r['origem']).startswith('A CONFIRMAR') else ())
        st = b.status_fic()
        self.t_status.preencher([{'doc': d, 'codigo': c, 'descricao': x} for d, c, x in st])
        self.cb_codigo.configure(values=sorted({c for _, c, _ in st}))
        self.t_imp.preencher([{**r, 'data_br': servico.data_br(r['data_referencia'])} for r in b.importacoes()])
        n_alerta = sum(1 for *_, r in R.verificacao if r == 'VERIFICAR')
        if n_alerta:
            self.selo_ver.configure(text=f'  {n_alerta}  ')
            self.selo_ver.place(relx=1.0, x=-10, rely=0.5, anchor='e')
        else:
            self.selo_ver.place_forget()
        self._atualizar_visao()

    def _detalhes(self):
        return [self.fvi, self.fvm, self.doc_ccm, self.doc_tap, self.doc_ttas]

    def _atualizar_visao(self):
        R = self.R
        linhas = R.visao + [R.total]
        cat = [v['unidade'] for v in R.visao] + ['Total']

        def serie(nome, k, i):
            return (nome, tema.SERIE[i], [v[k] for v in linhas])

        def rodape(campos):
            return lambda i, j: f'{numero(linhas[i][campos[j][0]])} de {numero(linhas[i][campos[j][1]])}'
        graf = (('fvi', [serie('FVI Campo II', 'pct_cii', 0), serie('FVI Campo III', 'pct_ciii', 1),
                         serie('FVI Campo V', 'pct_cv', 2)],
                 [('cii_aprov', 'cii_aplic'), ('ciii_aprov', 'ciii_aplic'), ('cv_aprov', 'cv_aplic')]),
                ('fvm', [serie('FVM Campo III', 'pct_fvm_ciii', 0), serie('FVM Campo V', 'pct_fvm_cv', 1)],
                 [('fvm_ciii', 'malhas'), ('fvm_cv', 'malhas')]),
                ('ttas', [serie('TTAS 1', 'pct_ttas1', 0), serie('TTAS 2', 'pct_ttas2', 1)],
                 [('ttas1', 'sop'), ('ttas2', 'sop')]),
                ('ccm', [serie('CCM', 'pct_ccm', 0)], [('ccm', 'ssop')]),
                ('tap', [serie('TAP', 'pct_tap', 0)], [('tap', 'ssop')]))
        for k, series, campos in graf:
            self.g[k].colunas_agrupadas(cat, series, rodape(campos), destaque_ultimo=True)

    # ------------------------------------------------------------------ páginas de detalhe
    def _opcoes_sop(self, d):
        u = d['unidade'].get()
        sops = [s['sop'] for s in self.R.sop if u == TODAS_UNIDADES or s['unidade'] == u]
        d['sop'].configure(values=[TODOS_SOP] + sops)
        if d['sop'].get() not in [TODOS_SOP] + sops:
            d['sop'].set(TODOS_SOP)
        self._opcoes_ssop(d)

    def _opcoes_ssop(self, d):
        u, sop = d['unidade'].get(), d['sop'].get()
        ssops = [s['ssop'] for s in self.R.ssop if (u == TODAS_UNIDADES or s['unidade'] == u)
                 and (sop in ('', TODOS_SOP) or s['sop'] == sop)]
        d['ssop'].configure(values=[TODAS_SSOP] + ssops)
        if d['ssop'].get() not in [TODAS_SSOP] + ssops:
            d['ssop'].set(TODAS_SSOP)

    def _mudou_unidade(self, d):
        d['sop'].set(TODOS_SOP)
        d['ssop'].set(TODAS_SSOP)
        self._opcoes_sop(d)
        self._preencher(d)

    def _mudou_sop(self, d):
        d['ssop'].set(TODAS_SSOP)
        self._opcoes_ssop(d)
        self._preencher(d)

    def _limpar(self, d):
        d['unidade'].set(TODAS_UNIDADES)
        d['sop'].set(TODOS_SOP)
        d['ssop'].set(TODAS_SSOP)
        if 'busca' in d:
            d['busca'].delete(0, 'end')
        if 'situacao' in d:
            d['situacao'].set(TODAS_SITUACOES)
        self._opcoes_sop(d)
        self._preencher(d)

    def _agendar(self, d):
        if self._agendado:
            self.after_cancel(self._agendado)
        self._agendado = self.after(250, lambda: self._preencher(d))

    @staticmethod
    def _texto_filtro(widget, padrao):
        v = widget.get().strip().upper()
        return '' if v in ('', padrao.upper()) else v

    def _preencher(self, d):
        R = self.R
        if d['nome'] == 'FVI':
            linhas, unidade = R.tags, 'TAGs'
        elif d['nome'] == 'FVM':
            linhas, unidade = R.malhas, 'malhas'
        elif d['fonte'] == R_SSOP:
            linhas, unidade = R.ssop, 'SSOP'
        else:
            linhas, unidade = R.sop, 'SOP'
        u = d['unidade'].get()
        if u != TODAS_UNIDADES:
            linhas = [r for r in linhas if r.get('unidade') == u]
        sop = self._texto_filtro(d['sop'], TODOS_SOP)
        if sop:
            linhas = [r for r in linhas if sop in str(r.get('sop') or '').upper()]
        ssop = self._texto_filtro(d['ssop'], TODAS_SSOP)
        if ssop:
            if unidade == 'SOP':   # TTAS é por SOP: mantém o SOP que contém a SSOP buscada
                sops = {self._sop_de_ssop[s] for s in self._sop_de_ssop if ssop in s.upper()}
                linhas = [r for r in linhas if r['sop'] in sops]
            else:
                linhas = [r for r in linhas if ssop in str(r.get('ssop') or '').upper()]
        if 'situacao' in d and d['situacao'].get() != TODAS_SITUACOES:
            linhas = [r for r in linhas if r[d['campo_situacao']] == d['situacao'].get()]
        if 'busca' in d:
            busca = d['busca'].get().strip().upper()
            if busca:
                linhas = [r for r in linhas if busca in str(r.get(d['ident']) or '').upper()]
        d['tabela'].preencher(linhas, self._cor_linha if d['nome'] in ('FVI', 'FVM') else None)
        d['contagem'].configure(text=f'{numero(len(linhas))} {unidade}')

    @staticmethod
    def _cor_linha(r):
        if r['situacao'] == 'CANCELADA':
            return ('mudo',)
        if r['situacao'] in ('SEM SSOP', 'FORA DA BASE'):
            return ('alerta',)
        if r.get('sem_fvi'):
            return ('aviso',)
        return ()

    def _preencher_historico(self):
        u = self.cb_hist.get() or 'TOTAL'
        linhas = [r for r in servico.historico_com_percentuais(self.banco) if r['unidade'] == u]
        self.t_hist.preencher(linhas, lambda r: ('total',) if r['unidade'] == 'TOTAL' else ())
        self.g_hist.titulo = f'Evolução – {"total do empreendimento" if u == "TOTAL" else u}'
        self.g_hist.linhas([r['data'] for r in linhas],
                           [(nome, tema.SERIE[i], [r[k] for r in linhas]) for i, (k, nome) in enumerate(rotulos.INDICADORES[:3])])

    def status_msg(self, msg):
        self.status.configure(text=msg)

    def ajuda(self):
        messagebox.showinfo('Como usar', (
            '1. Importar planilha ▸ SSOP: a LI com a base de SOP e SSOP (uma vez, ou quando a LI for revisada).\n'
            '2. Importar planilha ▸ FVI: o Relatório de Cadastro Geral exportado da FIC, sem editar.\n'
            '   FVM, CCM, TAP e TTAS: planilhas com cabeçalho reconhecido pelo nome das colunas\n'
            '   (Banco de dados ▸ Exportar planilhas-modelo mostra o formato).\n'
            '3. Confirme a data de referência: o avanço fica registrado no Histórico nessa data.\n'
            '4. Confira a página Verificação (os testes devem resultar OK) antes de apresentar.\n'
            '5. Exportar Excel gera a planilha gerencial para envio, com as mesmas páginas e gráficos.\n\n'
            'Visão Gerencial: gráficos de FVI, FVM, CCM, TAP e TTAS por unidade e o total.\n'
            'FVI e FVM: filtre por unidade, SOP, SSOP e TAG (ou malha).\n'
            'CCM, TAP e TTAS: filtre por unidade, SOP, SSOP e situação.\n'
            'As listas oficiais de status da FIC são importadas na página Critérios.'))

    def sobre(self):
        messagebox.showinfo('Sobre', f'{TITULO}\nVersão {__version__}\n\nBanco de dados: {self.banco.caminho if self.banco else "-"}')

    def sair(self):
        if self.banco:
            self.banco.fechar()
        self.destroy()


def abrir_arquivo(p):
    try:
        if hasattr(os, 'startfile'):
            os.startfile(p)  # Windows
        else:
            import subprocess
            subprocess.Popen(['xdg-open', str(p)])
    except Exception:
        pass


def executar(caminho_banco=None):
    app = App(caminho_banco)
    marca = os.environ.get('EPC12_TESTE_INICIO')
    if marca:   # usado pelo teste do executável: abre todas as páginas, grava "ok" e fecha
        def teste():
            for nome in PAGINAS:
                app.mostrar(nome)
                app.update()
            Path(marca).write_text('ok', encoding='utf-8')
            app.sair()
        app.after(1500, teste)
    app.mainloop()
