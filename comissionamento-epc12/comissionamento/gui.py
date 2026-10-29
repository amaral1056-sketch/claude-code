"""Interface gráfica (Tkinter) do Controle Gerencial de Comissionamento – EPC 12."""
import datetime
import os
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from . import __version__, configuracao, exportacao, servico, tema
from .banco import Banco
from .graficos import Grafico, Medidor
from .importacao import ErroImportacao
from .padroes import ETAPAS

TITULO = 'Controle Gerencial de Comissionamento – EPC 12'
XLSX = [('Planilhas Excel', '*.xlsx *.xlsm'), ('Todos os arquivos', '*.*')]

# (título, chave, largura em caracteres, é percentual)
COL_TAGS = [('TAG', 'tag', 17, 0), ('DESCRIÇÃO', 'descricao', 18, 0), ('UNIDADE', 'unidade', 8, 0),
            ('SSOP', 'ssop', 28, 0), ('Nº FVI', 'n_fvi', 7, 0), ('STATUS FVI', 'st_fvi', 11, 0),
            ('CAMPO II', 'd_cii', 13, 0), ('CAMPO III', 'd_ciii', 14, 0), ('CAMPO IV', 'st_civ', 9, 0),
            ('CAMPO V', 'd_cv', 14, 0), ('SITUAÇÃO', 'd_situacao', 11, 0)]
COL_MALHAS = [('MALHA', 'malha', 20, 0), ('UNIDADE', 'unidade', 9, 0), ('SOP', 'sop', 22, 0), ('SSOP', 'ssop', 26, 0),
              ('STATUS', 'status', 16, 0), ('CAMPO III', 'd_ciii', 14, 0), ('CAMPO V', 'd_cv', 14, 0),
              ('SITUAÇÃO', 'situacao', 14, 0)]
COL_HIST = [('DATA', 'data', 11, 0), ('UNIDADE', 'unidade', 11, 0), ('TAGS', 'tags', 8, 0), ('% FVI CII', 'pct_cii', 9, 1),
            ('% FVI CIII', 'pct_ciii', 9, 1), ('% FVI CV', 'pct_cv', 9, 1), ('MALHAS', 'malhas', 8, 0),
            ('% FVM CIII', 'pct_fvm_ciii', 9, 1), ('% FVM CV', 'pct_fvm_cv', 9, 1), ('% CCM', 'pct_ccm', 8, 1),
            ('% TAP', 'pct_tap', 8, 1), ('% TTAS1', 'pct_ttas1', 8, 1), ('% TTAS2', 'pct_ttas2', 8, 1),
            ('TAGS SEM FVI', 'tags_sem_fvi', 9, 0)]
COL_VER = [('GRUPO', 'g', 10, 0), ('TESTE', 't', 70, 0), ('ESPERADO', 'e', 9, 0), ('APURADO', 'a', 9, 0), ('RESULTADO', 'r', 12, 0)]
COL_PEND = [('DOCUMENTO', 'd', 10, 0), ('LINHA', 'l', 7, 0), ('IDENTIFICAÇÃO', 'i', 22, 0), ('SSOP / SOP', 's', 28, 0),
            ('PROBLEMA', 'p', 70, 0)]
COL_IMP = [('Nº', 'id', 5, 0), ('TIPO', 'tipo', 8, 0), ('IMPORTADO EM', 'importado_em', 18, 0),
           ('DATA REF.', 'data_br', 10, 0), ('LINHAS', 'linhas', 8, 0), ('ARQUIVO', 'arquivo', 60, 0),
           ('OBSERVAÇÃO', 'observacao', 50, 0)]
COL_CRIT = [('ETAPA', 'etapa', 11, 0), ('CÓDIGO', 'codigo', 18, 0), ('ORIGEM', 'origem', 50, 0)]
COL_STATUS = [('DOC.', 'doc', 6, 0), ('CÓDIGO', 'codigo', 16, 0), ('DESCRIÇÃO', 'descricao', 21, 0)]

SITUACOES_TAG = ('Todas', 'Válidas', 'Com pendência (II, III ou V)', 'Campo II aprovado', 'Sem FVI gerada',
                 'Canceladas', 'Sem SSOP', 'Fora da base')
SITUACOES_MALHA = ('Todas', 'Válidas', 'Campo III pendente', 'Campo III atingido', 'Canceladas', 'Sem SSOP', 'Fora da base')
PAGINAS = ('Painel', 'Visão Gerencial', 'FVI', 'FVM', 'CCM', 'TAP', 'TTAS', 'Verificação', 'Histórico', 'Critérios',
           'Importações')
ICONES = {'Painel': '◧', 'Visão Gerencial': '▤', 'FVI': '◉', 'FVM': '◎', 'CCM': '◇', 'TAP': '◆', 'TTAS': '▲',
          'Verificação': '✓', 'Histórico': '↗', 'Critérios': '⚙', 'Importações': '⇣'}


def fmt(v, percentual):
    if v is None:
        return ''
    if percentual:
        return f'{v * 100:.1f}%'.replace('.', ',')
    if isinstance(v, float):
        return f'{v:.2f}'.replace('.', ',')
    if isinstance(v, bool):
        return 'Sim' if v else 'Não'
    return str(v)


def status_campo(texto, ok):
    """Status de um campo com marca de leitura rápida: ✔ conta como aprovado · ○ pendente · — N/A/cancelada."""
    t = texto.strip() if texto else ''
    if ok is None:
        return f'— {t}' if t else '—'
    return f'✔ {t}' if ok else f'○ {t or "(vazio)"}'


def etapa_malha(ok):
    return '—' if ok is None else ('✔ atingido' if ok else '○ pendente')


class Tabela(ttk.Frame):
    """Treeview com rolagem, linhas listradas, cabeçalhos clicáveis (ordenar) e % formatado."""

    def __init__(self, master, colunas, altura=20, ao_duplo_clique=None):
        super().__init__(master, style='Card.TFrame')
        self.colunas = colunas
        ids = [k for _, k, _, _ in colunas]
        self.tree = ttk.Treeview(self, columns=ids, show='headings', height=altura, selectmode='extended')
        for titulo, k, larg, p in colunas:
            self.tree.heading(k, text=titulo, command=lambda k=k: self.ordenar(k))
            self.tree.column(k, width=max(48, larg * 7, len(titulo) * 8 + 16), anchor='e' if p else 'w', stretch=larg >= 22)
        sy = ttk.Scrollbar(self, orient='vertical', command=self.tree.yview)
        sx = ttk.Scrollbar(self, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        sy.grid(row=0, column=1, sticky='ns')
        sx.grid(row=1, column=0, sticky='ew')
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.tree.tag_configure('par', background=tema.LISTRA)
        self.tree.tag_configure('amarelo', background=tema.AMARELO_CRITERIO)
        self.tree.tag_configure('total', font=tema.fonte(9, 'bold'), background='#EEF3F9')
        self.tree.tag_configure('alerta', background=tema.ERRO_FUNDO, foreground=tema.ERRO)
        self.tree.tag_configure('aviso', background=tema.ALERTA_FUNDO)
        self.tree.tag_configure('mudo', foreground=tema.TEXTO_MUDO)
        self.tree.tag_configure('ok', foreground=tema.OK)
        if ao_duplo_clique:
            self.tree.bind('<Double-1>', lambda e: self._duplo(ao_duplo_clique))
        self.linhas = []
        self._tags = None
        self._ordem = (None, False)

    def _duplo(self, funcao):
        sel = self.selecionadas()
        if sel:
            funcao(sel[0])

    def preencher(self, linhas, tags=None):
        self.linhas = list(linhas)
        self._tags = tags
        self._mostrar()

    def _mostrar(self):
        # coluna só com números (contagens) fica alinhada à direita, como no Excel
        for _, k, _, p in self.colunas:
            vals = [r.get(k) for r in self.linhas[:300] if r.get(k) not in (None, '')]
            numerica = p or (vals and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in vals))
            self.tree.column(k, anchor='e' if numerica else 'w')
        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(self.linhas):
            vals = [fmt(r.get(k), p) for _, k, _, p in self.colunas]
            tags = tuple(self._tags(r)) if self._tags else ()
            if not tags and i % 2:
                tags = ('par',)
            self.tree.insert('', 'end', iid=str(i), values=vals, tags=tags)

    def ordenar(self, k):
        atual, inv = self._ordem
        inv = not inv if atual == k else False
        self._ordem = (k, inv)
        fixas = [r for r in self.linhas if r.get('unidade') == 'TOTAL']
        resto = [r for r in self.linhas if r.get('unidade') != 'TOTAL']
        resto.sort(key=lambda r: (r.get(k) is None, r.get(k) if not isinstance(r.get(k), str) else r.get(k).upper()),
                   reverse=inv)
        self.linhas = resto + fixas
        self._mostrar()

    def selecionadas(self):
        return [self.linhas[int(i)] for i in self.tree.selection()]


class Indicador(ttk.Frame):
    """Cartão de indicador: título, valor grande, medidor e linha de apoio."""

    def __init__(self, master, titulo, cor=tema.DESTAQUE, medidor=True):
        super().__init__(master, style='Card.TFrame', padding=(16, 12))
        ttk.Label(self, text=titulo.upper(), style='Card.TLabel', font=tema.fonte(8, 'bold'),
                  foreground=tema.TEXTO_MUDO).pack(anchor='w')
        self.valor = ttk.Label(self, text='–', style='Card.TLabel', font=tema.fonte(20, 'bold'), foreground=tema.TEXTO)
        self.valor.pack(anchor='w', pady=(2, 4))
        self.medidor = Medidor(self, cor) if medidor else None
        if self.medidor:
            self.medidor.pack(fill='x')
        self.apoio = ttk.Label(self, text='', style='Card.TLabel', font=tema.fonte(8), foreground=tema.TEXTO_2)
        self.apoio.pack(anchor='w', pady=(6, 0))

    def definir(self, valor, apoio='', pct=None):
        self.valor.config(text=valor)
        self.apoio.config(text=apoio)
        if self.medidor:
            self.medidor.definir(pct)


class App(tk.Tk):
    def __init__(self, caminho_banco=None):
        super().__init__()
        self.title(TITULO)
        self.geometry('1360x840')
        self.minsize(1040, 640)
        tema.aplicar(self)
        self.banco = None
        self.R = None
        self._busca_pendente = None
        self._montar()
        self.protocol('WM_DELETE_WINDOW', self.sair)
        self.abrir_banco(caminho_banco or configuracao.banco_padrao())

    # ================================================================== montagem
    def _menu_importar(self, master):
        m = tk.Menu(master, tearoff=0)
        m.add_command(label='LI – base SOP/SSOP…', command=self.importar_li)
        m.add_separator()
        for t in servico.TIPOS_FIC:
            m.add_command(label=servico.NOMES[t] + '…', command=lambda t=t: self.importar_fic(t))
        m.add_separator()
        m.add_command(label='Listas oficiais de status (STATUS_*.xlsx)…', command=self.importar_status)
        return m

    def _montar(self):
        menu = tk.Menu(self)
        arq = tk.Menu(menu, tearoff=0)
        arq.add_command(label='Abrir ou criar banco de dados…', command=self.escolher_banco)
        arq.add_command(label='Salvar cópia de segurança do banco…', command=self.backup)
        arq.add_separator()
        arq.add_command(label='Exportar relatório Excel…', command=self.exportar)
        arq.add_command(label='Exportar planilhas-modelo (FVM/CCM/TAP/TTAS)…', command=self.exportar_modelos)
        arq.add_separator()
        arq.add_command(label='Sair', command=self.sair)
        menu.add_cascade(label='Arquivo', menu=arq)
        menu.add_cascade(label='Importar', menu=self._menu_importar(menu))
        aj = tk.Menu(menu, tearoff=0)
        aj.add_command(label='Como usar', command=self.ajuda)
        aj.add_command(label='Sobre', command=lambda: messagebox.showinfo(
            'Sobre', f'{TITULO}\nVersão {__version__}\n\nBanco de dados: {self.banco.caminho if self.banco else "-"}'))
        menu.add_cascade(label='Ajuda', menu=aj)
        self.config(menu=menu)

        # barra lateral
        lateral = ttk.Frame(self, style='Lateral.TFrame', width=210)
        lateral.pack(side='left', fill='y')
        lateral.pack_propagate(False)
        ttk.Label(lateral, text='◆ Comissionamento', style='Marca.TLabel').pack(anchor='w', padx=18, pady=(20, 0))
        ttk.Label(lateral, text='EPC 12 · Controle gerencial', style='LateralNota.TLabel').pack(anchor='w', padx=18, pady=(0, 18))
        self.nav = {}
        for nome in PAGINAS:
            b = ttk.Button(lateral, text=f'{ICONES[nome]}   {nome}', style='Nav.TButton', command=lambda n=nome: self.mostrar(n))
            b.pack(fill='x')
            self.nav[nome] = b
        ttk.Label(lateral, text=f'versão {__version__}', style='LateralNota.TLabel').pack(side='bottom', anchor='w', padx=18, pady=12)

        direita = ttk.Frame(self)
        direita.pack(side='left', fill='both', expand=True)

        # cabeçalho
        topo = ttk.Frame(direita, style='Topo.TFrame', padding=(24, 14))
        topo.pack(fill='x')
        textos = ttk.Frame(topo, style='Topo.TFrame')
        textos.pack(side='left')
        self.lbl_pagina = ttk.Label(textos, text='Painel', style='TopoTitulo.TLabel')
        self.lbl_pagina.pack(anchor='w')
        self.lbl_info = ttk.Label(textos, text='', style='TopoSub.TLabel')
        self.lbl_info.pack(anchor='w')
        ttk.Button(topo, text='Exportar Excel', command=self.exportar).pack(side='right', padx=(8, 0))
        mb = ttk.Menubutton(topo, text='Importar planilha', style='Destaque.TMenubutton')
        mb['menu'] = self._menu_importar(mb)
        mb.pack(side='right')
        tk.Frame(direita, height=1, background=tema.BORDA).pack(fill='x')

        self.corpo = ttk.Frame(direita, padding=(20, 16))
        self.corpo.pack(fill='both', expand=True)
        self.status = ttk.Label(direita, text='', style='Nota.TLabel', padding=(22, 4))
        self.status.pack(fill='x', side='bottom')

        self.paginas = {}
        for nome, montar in (('Painel', self._pg_painel), ('Visão Gerencial', self._pg_visao), ('FVI', self._pg_fvi),
                             ('FVM', self._pg_fvm), ('CCM', lambda f: self._pg_doc(f, 'CCM', {'SSOP': exportacao.COL_CCM})),
                             ('TAP', lambda f: self._pg_doc(f, 'TAP', {'SSOP': exportacao.COL_TAP})),
                             ('TTAS', lambda f: self._pg_doc(f, 'TTAS', {'SOP': exportacao.COL_TTAS})),
                             ('Verificação', self._pg_verificacao), ('Histórico', self._pg_historico),
                             ('Critérios', self._pg_criterios), ('Importações', self._pg_importacoes)):
            f = ttk.Frame(self.corpo)
            montar(f)
            self.paginas[nome] = f
        self.pagina_atual = None
        self.mostrar('Painel')

    def mostrar(self, nome):
        if self.pagina_atual:
            self.paginas[self.pagina_atual].pack_forget()
            self.nav[self.pagina_atual].configure(style='Nav.TButton')
        self.paginas[nome].pack(fill='both', expand=True)
        self.nav[nome].configure(style='NavAtivo.TButton')
        self.lbl_pagina.config(text=nome)
        self.pagina_atual = nome

    # ------------------------------------------------------------------ páginas
    def _pg_painel(self, f):
        cards = ttk.Frame(f)
        cards.pack(fill='x')
        self.ind = {}
        defs = [('tags', 'TAGs válidas', tema.DESTAQUE, False), ('cii', 'FVI · Campo II', tema.SERIE[0], True),
                ('ciii', 'FVI · Campo III', tema.SERIE[1], True), ('cv', 'FVI · Campo V', tema.SERIE[2], True),
                ('fvm', 'FVM · Campo III', tema.SERIE[0], True), ('ccm', 'CCM aprovados', tema.SERIE[0], True),
                ('tap', 'TAP concluídos', tema.SERIE[1], True), ('ttas', 'TTAS1 aprovados', tema.SERIE[2], True)]
        for i, (k, t, cor, med) in enumerate(defs):
            c = Indicador(cards, t, cor, med)
            c.grid(row=i // 4, column=i % 4, sticky='nsew', padx=(0 if i % 4 == 0 else 12, 0), pady=(0, 12))
            self.ind[k] = c
        for j in range(4):
            cards.columnconfigure(j, weight=1, uniform='c')
        graf = ttk.Frame(f)
        graf.pack(fill='both', expand=True)
        self.g_unid = Grafico(graf, 'Avanço do FVI por unidade', 'aprovados ÷ aplicáveis · passe o mouse nas colunas')
        self.g_sit = Grafico(graf, 'Situação das TAGs por campo', 'TAGs válidas alocadas em SSOP da base')
        self.g_hist = Grafico(graf, 'Evolução do FVI (total)', 'um ponto por data de referência registrada no histórico')
        self.g_doc = Grafico(graf, 'Documentos por unidade', '% de SSOP com CCM aprovado / TAP concluído · % de SOP com TTAS1')
        for i, g in enumerate((self.g_unid, self.g_sit, self.g_hist, self.g_doc)):
            g.grid(row=i // 2, column=i % 2, sticky='nsew', padx=(0 if i % 2 == 0 else 12, 0), pady=(0 if i < 2 else 12, 0))
        graf.columnconfigure(0, weight=3, uniform='g')
        graf.columnconfigure(1, weight=2, uniform='g')
        graf.rowconfigure(0, weight=1)
        graf.rowconfigure(1, weight=1)

    def _pg_visao(self, f):
        self.lbl_visao = ttk.Label(f, text='', style='Sub.TLabel')
        self.lbl_visao.pack(anchor='w', pady=(0, 8))
        self.t_visao = Tabela(f, exportacao.COL_VISAO, altura=16)
        self.t_visao.pack(fill='both', expand=True)
        for t in exportacao.NOTAS_VISAO:
            ttk.Label(f, text=t, style='Nota.TLabel').pack(anchor='w')

    def _barra_filtros(self, f, d, niveis, situacoes):
        barra = ttk.Frame(f)
        barra.pack(fill='x', pady=(0, 10))
        d['contagem'] = ttk.Label(barra, text='', style='Sub.TLabel')
        d['contagem'].pack(side='right')
        if len(niveis) > 1:
            seg = ttk.Frame(barra, style='Card.TFrame')
            seg.pack(side='left', padx=(0, 16))
            for k in niveis:
                ttk.Radiobutton(seg, text=k, value=k, variable=d['nivel'], style='Seg.TRadiobutton',
                                command=lambda: self._trocar_nivel(d)).pack(side='left')
        ttk.Label(barra, text='Unidade').pack(side='left')
        d['unidade'] = ttk.Combobox(barra, state='readonly', width=10)
        d['unidade'].pack(side='left', padx=(6, 14))
        d['unidade'].bind('<<ComboboxSelected>>', lambda e: self._preencher(d))
        if situacoes:
            d['lbl_sit'] = ttk.Label(barra, text='Situação')
            d['lbl_sit'].pack(side='left')
            d['situacao'] = ttk.Combobox(barra, state='readonly', width=21, values=situacoes)
            d['situacao'].set(situacoes[0])
            d['situacao'].pack(side='left', padx=(6, 14))
            d['situacao'].bind('<<ComboboxSelected>>', lambda e: self._preencher(d))
        ttk.Label(barra, text='Buscar').pack(side='left')
        d['busca'] = tk.StringVar()
        en = ttk.Entry(barra, textvariable=d['busca'], width=16)
        en.pack(side='left', padx=6)
        en.bind('<KeyRelease>', lambda e: self._agendar(d))
        ttk.Button(barra, text='Limpar', width=8, command=lambda: self._limpar(d)).pack(side='left')

    def _pg_fvi(self, f):
        d = self.fvi = {'nome': 'FVI', 'nivel': tk.StringVar(value='TAGs'), 'tabelas': {}}
        self._barra_filtros(f, d, ('TAGs', 'SSOP', 'SOP'), SITUACOES_TAG)
        d['tabelas']['TAGs'] = Tabela(f, COL_TAGS, ao_duplo_clique=None)
        d['tabelas']['SSOP'] = Tabela(f, exportacao.COL_FVI_SSOP, ao_duplo_clique=lambda r: self._abrir_tags(d, r['ssop']))
        d['tabelas']['SOP'] = Tabela(f, exportacao.COL_FVI_SOP, ao_duplo_clique=lambda r: self._abrir_nivel(d, 'SSOP', r['sop']))
        d['nota'] = ttk.Label(f, style='Nota.TLabel', wraplength=1050, text=(
            '✔ conta como aprovado pelos critérios · ○ pendente · — N/A ou TAG cancelada (fora do denominador).  '
            'Linhas em amarelo: TAG sem FVI gerada · em vermelho: sem SSOP ou SSOP fora da base · cinza: cancelada.  '
            'Duplo clique numa SSOP mostra as TAGs dela.'))
        d['nota'].pack(side='bottom', anchor='w', pady=(6, 0))
        d['tabelas']['TAGs'].pack(fill='both', expand=True)

    def _pg_fvm(self, f):
        d = self.fvm = {'nome': 'FVM', 'nivel': tk.StringVar(value='Malhas'), 'tabelas': {}}
        self._barra_filtros(f, d, ('Malhas', 'SSOP', 'SOP'), SITUACOES_MALHA)
        d['tabelas']['Malhas'] = Tabela(f, COL_MALHAS)
        d['tabelas']['SSOP'] = Tabela(f, exportacao.COL_FVM_SSOP, ao_duplo_clique=lambda r: self._abrir_tags(d, r['ssop']))
        d['tabelas']['SOP'] = Tabela(f, exportacao.COL_FVM_SOP, ao_duplo_clique=lambda r: self._abrir_nivel(d, 'SSOP', r['sop']))
        ttk.Label(f, style='Nota.TLabel', wraplength=1050, text=(
            'Etapa por malha conforme os critérios FVM_CIII e FVM_CV (aba Critérios): ✔ atingido · ○ pendente · — cancelada.  '
            'Layout da FVM ainda provisório (MALHA | SOP | SSOP | STATUS).  Duplo clique numa SSOP mostra as malhas dela.')
                  ).pack(side='bottom', anchor='w', pady=(6, 0))
        d['tabelas']['Malhas'].pack(fill='both', expand=True)

    def _pg_doc(self, f, nome, opcoes):
        d = {'nome': nome, 'nivel': tk.StringVar(value=next(iter(opcoes))), 'tabelas': {}}
        self._barra_filtros(f, d, tuple(opcoes), None)
        for k, cols in opcoes.items():
            d['tabelas'][k] = Tabela(f, cols)
        d['tabelas'][d['nivel'].get()].pack(fill='both', expand=True)
        setattr(self, 'doc_' + nome.lower(), d)

    def _pg_verificacao(self, f):
        ttk.Label(f, text='Testes das importações (devem dar OK) e informações', style='Sub.TLabel').pack(anchor='w')
        self.t_ver = Tabela(f, COL_VER, altura=12)
        self.t_ver.pack(fill='both', expand=True, pady=(4, 12))
        ttk.Label(f, text='Pendências encontradas, linha a linha do arquivo importado', style='Sub.TLabel').pack(anchor='w')
        self.t_pend = Tabela(f, COL_PEND, altura=10)
        self.t_pend.pack(fill='both', expand=True, pady=(4, 0))

    def _pg_historico(self, f):
        barra = ttk.Frame(f)
        barra.pack(fill='x', pady=(0, 10))
        ttk.Label(barra, text='Unidade').pack(side='left')
        self.cb_hist = ttk.Combobox(barra, state='readonly', width=14)
        self.cb_hist.pack(side='left', padx=(6, 14))
        self.cb_hist.bind('<<ComboboxSelected>>', lambda e: self._preencher_historico())
        ttk.Button(barra, text='Registrar situação atual em uma data…', command=self.registrar_historico).pack(side='left')
        ttk.Button(barra, text='Remover data selecionada', command=self.remover_historico).pack(side='left', padx=8)
        self.g_hist2 = Grafico(f, 'Evolução', 'FVI Campo II, III e V da seleção', altura=230)
        self.g_hist2.pack(fill='x', pady=(0, 12))
        ttk.Label(f, text='Cada importação grava o avanço na data de referência informada (mesma data = substitui).',
                  style='Nota.TLabel').pack(anchor='w')
        self.t_hist = Tabela(f, COL_HIST)
        self.t_hist.pack(fill='both', expand=True)

    def _pg_criterios(self, f):
        par = ttk.Frame(f, style='Card.TFrame', padding=(16, 10))
        par.pack(fill='x', pady=(0, 12))
        self.var_sem_fvi = tk.StringVar(value='SIM')
        ttk.Label(par, text='Incluir TAGs sem FVI gerada no denominador do FVI?', style='Card.TLabel',
                  font=tema.fonte(10, 'bold')).pack(side='left')
        seg = ttk.Frame(par, style='Card.TFrame')
        seg.pack(side='left', padx=12)
        for v in ('SIM', 'NÃO'):
            ttk.Radiobutton(seg, text=v, value=v, variable=self.var_sem_fvi, style='Seg.TRadiobutton',
                            command=self.mudar_parametro).pack(side='left')
        ttk.Label(par, text='SIM = TAG sem FVI conta como pendente (visão conservadora) · NÃO = só TAGs com FVI entram nos %.',
                  style='Card.TLabel', font=tema.fonte(8), foreground=tema.TEXTO_MUDO).pack(side='left', padx=8)
        corpo = ttk.Frame(f)
        corpo.pack(fill='both', expand=True)
        dir_ = ttk.Frame(corpo)
        dir_.pack(side='right', fill='y', padx=(16, 0))
        ttk.Label(dir_, text='Listas oficiais de status da FIC', style='Sub.TLabel').pack(anchor='w')
        self.t_status = Tabela(dir_, COL_STATUS)
        self.t_status.pack(fill='both', expand=True, pady=(4, 0))
        esq = ttk.Frame(corpo)
        esq.pack(side='left', fill='both', expand=True)
        ttk.Label(esq, text='Códigos de status que contam como "etapa atingida" (amarelo = a confirmar)',
                  style='Sub.TLabel').pack(anchor='w')
        botoes = ttk.Frame(esq)
        botoes.pack(side='bottom', fill='x', pady=(8, 0))
        form = ttk.Frame(esq)
        form.pack(side='bottom', fill='x')
        self.t_crit = Tabela(esq, COL_CRIT, altura=10)
        self.t_crit.pack(fill='both', expand=True, pady=(4, 8))
        for j, t in enumerate(('Etapa', 'Código', 'Origem / justificativa')):
            ttk.Label(form, text=t, style='Sub.TLabel').grid(row=0, column=j, sticky='w')
        self.cb_etapa = ttk.Combobox(form, values=ETAPAS, state='readonly', width=12)
        self.cb_etapa.grid(row=1, column=0, padx=(0, 6))
        self.cb_codigo = ttk.Combobox(form, width=20)
        self.cb_codigo.grid(row=1, column=1, padx=(0, 6))
        self.en_origem = ttk.Entry(form, width=30)
        self.en_origem.grid(row=1, column=2, padx=(0, 6), sticky='ew')
        form.columnconfigure(2, weight=1)
        ttk.Button(form, text='Adicionar', style='Destaque.TButton', command=self.adicionar_criterio).grid(row=1, column=3)
        ttk.Button(botoes, text='Remover selecionados', command=self.remover_criterio).pack(side='left')
        ttk.Button(botoes, text='Restaurar critérios padrão', command=self.restaurar_criterios).pack(side='left', padx=8)

    def _pg_importacoes(self, f):
        ttk.Label(f, text='Registro de todas as importações gravadas neste banco', style='Sub.TLabel').pack(anchor='w')
        self.t_imp = Tabela(f, COL_IMP)
        self.t_imp.pack(fill='both', expand=True, pady=(4, 0))

    # ================================================================== banco
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
            self.status_msg('Banco novo: comece por Importar planilha ▸ LI – base SOP/SSOP.')

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

    # ================================================================== importações
    def _executar(self, acao, msg_ok=None):
        self.config(cursor='watch')
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
            self.config(cursor='')
        if msg_ok:
            self.status_msg(msg_ok)
        return r if r is not None else True

    def importar_li(self):
        p = filedialog.askopenfilename(title='Selecione a LI (lista SOP/SSOP)', filetypes=XLSX)
        if not p:
            return
        if self.banco.tem_base() and not messagebox.askyesno(
                'Substituir base', 'A base SOP/SSOP atual será substituída pela desta LI. Continuar?'):
            return
        msg = self._executar(lambda: servico.importar_li(self.banco, p))
        if msg:
            self.atualizar()
            self.status_msg(msg)
            messagebox.showinfo('LI importada', msg)

    def importar_fic(self, tipo):
        if not self.banco.tem_base():
            messagebox.showwarning('Base ausente', 'Importe primeiro a LI (Importar planilha ▸ LI – base SOP/SSOP).')
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
                extra += f'\n{len(ign)} linha(s) ignorada(s) (cabeçalho repetido/rodapé): {", ".join(map(str, ign[:10]))}'
            alerta = [t for _, t, _, _, r in self.R.verificacao if r == 'VERIFICAR']
            if alerta:
                extra += '\n\nAtenção – testes com VERIFICAR (ver Verificação):\n- ' + '\n- '.join(alerta)
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

    # ================================================================== exportações
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

    # ================================================================== histórico e critérios
    def registrar_historico(self):
        data = simpledialog.askstring('Registrar no histórico', 'Data de referência (dd/mm/aaaa):',
                                      initialvalue=datetime.date.today().strftime('%d/%m/%Y'), parent=self)
        if data and self._executar(lambda: servico.registrar_historico(self.banco, data), f'Histórico registrado em {data}'):
            self.atualizar()

    def remover_historico(self):
        sel = self.t_hist.selecionadas()
        datas = sorted({r['data_referencia'] for r in sel})
        if datas and messagebox.askyesno('Remover', 'Remover do histórico as datas: '
                                         + ', '.join(servico.data_br(d) for d in datas) + '?'):
            for d in datas:
                self.banco.remover_historico(d)
            self.atualizar()

    def mudar_parametro(self):
        self.banco.definir_incluir_sem_fvi(self.var_sem_fvi.get() == 'SIM')
        self.atualizar()
        self.status_msg('Parâmetro alterado. O histórico já gravado não muda; use "Registrar situação atual" se quiser.')

    def adicionar_criterio(self):
        e, c, o = self.cb_etapa.get(), self.cb_codigo.get().strip(), self.en_origem.get().strip()
        if not e or not c:
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

    # ================================================================== atualização das telas
    def atualizar(self):
        b = self.banco
        try:
            self.R = servico.calcular(b)
        except Exception:
            configuracao.registrar_erro(traceback.format_exc())
            raise
        R = self.R
        data = servico.data_br(b.data_referencia_atual())
        self.lbl_info.config(text=f'Data de referência: {data or "–"}   ·   Banco: {b.caminho}')
        self.title(f'{TITULO} – {b.caminho.name}')
        self.lbl_visao.config(text=f'Avanço por unidade – data de referência {data or "–"}   ·   '
                                   f'TAGs sem FVI no denominador: {"SIM" if b.incluir_sem_fvi else "NÃO"}')
        self.t_visao.preencher(R.visao + [R.total], lambda r: ('total',) if r['unidade'] == 'TOTAL' else ())
        self._preparar_linhas()
        unidades = ['(todas)'] + R.unidades
        for d in self._detalhes():
            if d['unidade'].get() not in unidades:
                d['unidade'].set('(todas)')
            d['unidade']['values'] = unidades
            self._preencher(d)
        self.t_ver.preencher([{'g': g, 't': t, 'e': '-' if e is None else e, 'a': a, 'r': r} for g, t, e, a, r in R.verificacao],
                             lambda r: ('alerta',) if r['r'] == 'VERIFICAR' else (('ok',) if r['r'] == 'OK' else ()))
        self.t_pend.preencher([dict(zip('dlisp', p)) for p in R.pendencias])
        self.cb_hist['values'] = ['TOTAL', '(todas)'] + R.unidades
        if not self.cb_hist.get():
            self.cb_hist.set('TOTAL')
        self._preencher_historico()
        self.var_sem_fvi.set('SIM' if b.incluir_sem_fvi else 'NÃO')
        self.t_crit.preencher([{'id': i, 'etapa': e, 'codigo': c, 'origem': o} for i, e, c, o in b.criterios()],
                              lambda r: ('amarelo',) if str(r['origem']).startswith('A CONFIRMAR') else ())
        st = b.status_fic()
        self.t_status.preencher([{'doc': d, 'codigo': c, 'descricao': x} for d, c, x in st])
        self.cb_codigo['values'] = sorted({c for _, c, _ in st})
        self.t_imp.preencher([{**r, 'data_br': servico.data_br(r['data_referencia'])} for r in b.importacoes()])
        n_alerta = sum(1 for *_, r in R.verificacao if r == 'VERIFICAR')
        self.nav['Verificação'].configure(text=f'{ICONES["Verificação"]}   Verificação' + (f'   ⚠ {n_alerta}' if n_alerta else ''))
        self._atualizar_painel()

    def _detalhes(self):
        return [self.fvi, self.fvm, self.doc_ccm, self.doc_tap, self.doc_ttas]

    def _preparar_linhas(self):
        for t in self.R.tags:
            for c in ('cii', 'ciii', 'cv'):
                t['d_' + c] = status_campo(t['st_' + c], t['ok_' + c])
            t['d_situacao'] = 'SEM FVI' if t['situacao'] == 'VÁLIDA' and t['sem_fvi'] else t['situacao']
        for m in self.R.malhas:
            m['d_ciii'], m['d_cv'] = etapa_malha(m['ok_ciii']), etapa_malha(m['ok_cv'])

    # ------------------------------------------------------------------ painel
    def _atualizar_painel(self):
        R, t = self.R, self.R.total
        p = lambda v: '–' if v is None else f'{v * 100:.1f}%'.replace('.', ',')
        n_fvi = R.contagens.get('FVI', {})
        self.ind['tags'].definir(f'{t["tags"]:,}'.replace(',', '.'),
                                 f'{n_fvi.get("lidas", 0)} lidas · {t["tags_sem_fvi"]} sem FVI · {n_fvi.get("sem_ssop", 0)} sem SSOP')
        for k in ('cii', 'ciii', 'cv'):
            self.ind[k].definir(p(t['pct_' + k]), f'{t[k + "_aprov"]} de {t[k + "_aplic"]} aplicáveis', t['pct_' + k])
        self.ind['fvm'].definir(p(t['pct_fvm_ciii']), f'{t["fvm_ciii"]} de {t["malhas"]} malhas · CV {p(t["pct_fvm_cv"])}',
                                t['pct_fvm_ciii'])
        self.ind['ccm'].definir(p(t['pct_ccm']), f'{t["ccm"]} de {t["ssop"]} SSOP', t['pct_ccm'])
        self.ind['tap'].definir(p(t['pct_tap']), f'{t["tap"]} de {t["ssop"]} SSOP', t['pct_tap'])
        self.ind['ttas'].definir(p(t['pct_ttas1']), f'{t["ttas1"]} de {t["sop"]} SOP · TTAS2 {p(t["pct_ttas2"])}', t['pct_ttas1'])

        un = [v['unidade'] for v in R.visao]
        nomes = (('Campo II', 'cii'), ('Campo III', 'ciii'), ('Campo V', 'cv'))
        self.g_unid.colunas_agrupadas(
            un, [(n, tema.SERIE[i], [v['pct_' + k] for v in R.visao]) for i, (n, k) in enumerate(nomes)],
            rodape=lambda i, j: f'{R.visao[i][nomes[j][1] + "_aprov"]} de {R.visao[i][nomes[j][1] + "_aplic"]} aplicáveis')
        validas = [x for x in R.tags if x['situacao'] == 'VÁLIDA']
        cont = {k: [sum(1 for x in validas if x['ok_' + k] is v) for v in (True, False, None)] for _, k in nomes}
        self.g_sit.barras_empilhadas(
            [n for n, _ in nomes],
            [('Aprovadas', tema.SERIE[0], [cont[k][0] for _, k in nomes]),
             ('Pendentes', tema.SERIE[1], [cont[k][1] for _, k in nomes]),
             ('N/A', tema.NEUTRO, [cont[k][2] for _, k in nomes])])
        hist = [h for h in servico.historico_com_percentuais(self.banco) if h['unidade'] == 'TOTAL']
        self.g_hist.linhas([h['data'][:5] for h in hist],
                           [(n, tema.SERIE[i], [h['pct_' + k] for h in hist]) for i, (n, k) in enumerate(nomes)])
        self.g_doc.colunas_agrupadas(un, [('CCM', tema.SERIE[0], [v['pct_ccm'] for v in R.visao]),
                                          ('TAP', tema.SERIE[1], [v['pct_tap'] for v in R.visao]),
                                          ('TTAS1', tema.SERIE[2], [v['pct_ttas1'] for v in R.visao])])

    # ------------------------------------------------------------------ detalhes (FVI, FVM, CCM, TAP, TTAS)
    def _trocar_nivel(self, d):
        for t in d['tabelas'].values():
            t.pack_forget()
        d['tabelas'][d['nivel'].get()].pack(fill='both', expand=True)
        tem_sit = d['nivel'].get() in ('TAGs', 'Malhas')
        if 'situacao' in d:
            d['situacao'].configure(state='readonly' if tem_sit else 'disabled')
        self._preencher(d)

    def _limpar(self, d):
        d['busca'].set('')
        d['unidade'].set('(todas)')
        if 'situacao' in d:
            d['situacao'].set(d['situacao']['values'][0])
        self._preencher(d)

    def _agendar(self, d):
        if self._busca_pendente:
            self.after_cancel(self._busca_pendente)
        self._busca_pendente = self.after(250, lambda: self._preencher(d))

    def _abrir_tags(self, d, ssop):
        d['nivel'].set('TAGs' if d['nome'] == 'FVI' else 'Malhas')
        if 'situacao' in d:
            d['situacao'].set(d['situacao']['values'][0])
        d['busca'].set(ssop)
        self._trocar_nivel(d)

    def _abrir_nivel(self, d, nivel, filtro):
        d['nivel'].set(nivel)
        d['busca'].set(filtro)
        self._trocar_nivel(d)

    def _preencher(self, d):
        nivel = d['nivel'].get()
        R = self.R
        if nivel == 'TAGs':
            linhas, campos = R.tags, ('tag', 'ssop', 'descricao', 'classe', 'n_fvi')
        elif nivel == 'Malhas':
            linhas, campos = R.malhas, ('malha', 'ssop', 'sop', 'status')
        else:
            linhas, campos = (R.ssop if nivel == 'SSOP' else R.sop), ('sop', 'ssop', 'descricao')
        u = d['unidade'].get()
        if u and u != '(todas)':
            linhas = [r for r in linhas if r.get('unidade') == u]
        if nivel in ('TAGs', 'Malhas') and 'situacao' in d:
            linhas = [r for r in linhas if self._filtro_situacao(nivel, d['situacao'].get(), r)]
        busca = d['busca'].get().strip().upper()
        if busca:
            linhas = [r for r in linhas if any(busca in str(r.get(c) or '').upper() for c in campos)]
        tags = self._cor_tag if nivel == 'TAGs' else (self._cor_malha if nivel == 'Malhas' else None)
        d['tabelas'][nivel].preencher(linhas, tags)
        d['contagem'].config(text=f'{len(linhas):,} {nivel if nivel in ("TAGs", "Malhas") else nivel}'.replace(',', '.'))

    @staticmethod
    def _filtro_situacao(nivel, sit, r):
        if sit in ('', 'Todas'):
            return True
        valida = r['situacao'] == 'VÁLIDA'
        if sit == 'Válidas':
            return valida
        if sit == 'Canceladas':
            return r['situacao'] == 'CANCELADA'
        if sit == 'Sem SSOP':
            return r['situacao'] == 'SEM SSOP'
        if sit == 'Fora da base':
            return r['situacao'] == 'FORA DA BASE'
        if nivel == 'TAGs':
            if sit == 'Sem FVI gerada':
                return r['sem_fvi']
            if sit == 'Campo II aprovado':
                return valida and r['ok_cii'] is True
            if sit.startswith('Com pendência'):
                return valida and any(r['ok_' + c] is False for c in ('cii', 'ciii', 'cv'))
        else:
            if sit == 'Campo III pendente':
                return valida and r['ok_ciii'] is False
            if sit == 'Campo III atingido':
                return valida and r['ok_ciii'] is True
        return True

    @staticmethod
    def _cor_tag(r):
        if r['situacao'] == 'CANCELADA':
            return ('mudo',)
        if r['situacao'] in ('SEM SSOP', 'FORA DA BASE'):
            return ('alerta',)
        if r['sem_fvi']:
            return ('aviso',)
        return ()

    @staticmethod
    def _cor_malha(r):
        if r['situacao'] == 'CANCELADA':
            return ('mudo',)
        if r['situacao'] in ('SEM SSOP', 'FORA DA BASE'):
            return ('alerta',)
        return ()

    def _preencher_historico(self):
        u = self.cb_hist.get()
        linhas = servico.historico_com_percentuais(self.banco)
        if u and u != '(todas)':
            linhas = [r for r in linhas if r['unidade'] == u]
        self.t_hist.preencher(linhas, lambda r: ('total',) if r['unidade'] == 'TOTAL' else ())
        serie = [r for r in linhas if r['unidade'] == (u if u and u != '(todas)' else 'TOTAL')]
        self.g_hist2.titulo = f'Evolução – {u if u and u != "(todas)" else "TOTAL"}'
        self.g_hist2.linhas([r['data'] for r in serie],
                            [(n, tema.SERIE[i], [r['pct_' + k] for r in serie])
                             for i, (n, k) in enumerate((('Campo II', 'cii'), ('Campo III', 'ciii'), ('Campo V', 'cv')))])

    def status_msg(self, msg):
        self.status.config(text=msg)

    def ajuda(self):
        messagebox.showinfo('Como usar', (
            '1. Importar planilha ▸ LI – base SOP/SSOP (uma vez, ou quando a LI for revisada).\n'
            '2. Importar planilha ▸ Relatório de Cadastro Geral (FVI): o arquivo exportado da FIC, sem editar.\n'
            '   FVM, CCM, TAP e TTAS: planilhas com cabeçalho reconhecido pelo nome das colunas\n'
            '   (Arquivo ▸ Exportar planilhas-modelo mostra o formato).\n'
            '3. Confirme a data de referência: o avanço fica registrado no Histórico nessa data.\n'
            '4. Confira a página Verificação (testes devem dar OK) antes de apresentar.\n'
            '5. Exportar Excel gera a planilha gerencial para envio.\n\n'
            'FVI: veja TAG a TAG (status de cada campo), por SSOP ou por SOP; duplo clique numa SSOP abre as TAGs dela.\n'
            'Cada importação substitui a anterior do mesmo tipo (a FIC exporta sempre a lista completa).\n'
            'Critérios de "aprovado" e o parâmetro de TAGs sem FVI ficam em Critérios.\n'
            'Faça cópias de segurança do banco em Arquivo ▸ Salvar cópia de segurança.'))

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
    App(caminho_banco).mainloop()
