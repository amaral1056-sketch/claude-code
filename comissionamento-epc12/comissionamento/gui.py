"""Interface gráfica (Tkinter) do Controle Gerencial de Comissionamento – EPC 12."""
import datetime
import os
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from . import __version__, configuracao, exportacao, servico
from .banco import Banco
from .importacao import ErroImportacao
from .padroes import ETAPAS

TITULO = 'Controle Gerencial de Comissionamento – EPC 12'
XLSX = [('Planilhas Excel', '*.xlsx *.xlsm'), ('Todos os arquivos', '*.*')]
COL_HIST = [('DATA', 'data', 11, 0), ('UNIDADE', 'unidade', 11, 0), ('TAGS', 'tags', 8, 0), ('% FVI CII', 'pct_cii', 9, 1),
            ('% FVI CIII', 'pct_ciii', 9, 1), ('% FVI CV', 'pct_cv', 9, 1), ('MALHAS', 'malhas', 8, 0),
            ('% FVM CIII', 'pct_fvm_ciii', 9, 1), ('% FVM CV', 'pct_fvm_cv', 9, 1), ('% CCM', 'pct_ccm', 8, 1),
            ('% TAP', 'pct_tap', 8, 1), ('% TTAS1', 'pct_ttas1', 8, 1), ('% TTAS2', 'pct_ttas2', 8, 1),
            ('TAGS SEM FVI', 'tags_sem_fvi', 9, 0)]
COL_VER = [('GRUPO', 'g', 10, 0), ('TESTE', 't', 70, 0), ('ESPERADO', 'e', 9, 0), ('APURADO', 'a', 9, 0), ('RESULTADO', 'r', 11, 0)]
COL_PEND = [('DOCUMENTO', 'd', 10, 0), ('LINHA', 'l', 7, 0), ('IDENTIFICAÇÃO', 'i', 22, 0), ('SSOP / SOP', 's', 28, 0),
            ('PROBLEMA', 'p', 70, 0)]
COL_IMP = [('Nº', 'id', 5, 0), ('TIPO', 'tipo', 8, 0), ('IMPORTADO EM', 'importado_em', 18, 0),
           ('DATA REF.', 'data_br', 10, 0), ('LINHAS', 'linhas', 8, 0), ('ARQUIVO', 'arquivo', 60, 0),
           ('OBSERVAÇÃO', 'observacao', 50, 0)]
COL_CRIT = [('ETAPA', 'etapa', 11, 0), ('CÓDIGO', 'codigo', 18, 0), ('ORIGEM', 'origem', 70, 0)]
COL_STATUS = [('DOC.', 'doc', 6, 0), ('CÓDIGO', 'codigo', 18, 0), ('DESCRIÇÃO', 'descricao', 40, 0)]


def fmt(v, percentual):
    if v is None:
        return ''
    if percentual:
        return f'{v * 100:.1f}%'.replace('.', ',')
    if isinstance(v, float):
        return f'{v:.2f}'.replace('.', ',')
    return str(v)


class Tabela(ttk.Frame):
    """Treeview com rolagem, cabeçalhos clicáveis para ordenar e % formatado. Colunas no formato de exportacao.COL_*."""

    def __init__(self, master, colunas, altura=20):
        super().__init__(master)
        self.colunas = colunas
        ids = [k for _, k, _, _ in colunas]
        self.tree = ttk.Treeview(self, columns=ids, show='headings', height=altura, selectmode='extended')
        for titulo, k, larg, p in colunas:
            self.tree.heading(k, text=titulo, command=lambda k=k: self.ordenar(k))
            self.tree.column(k, width=max(50, larg * 8, len(titulo) * 8 + 14), anchor='e' if p else 'w', stretch=larg >= 40)
        sy = ttk.Scrollbar(self, orient='vertical', command=self.tree.yview)
        sx = ttk.Scrollbar(self, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        sy.grid(row=0, column=1, sticky='ns')
        sx.grid(row=1, column=0, sticky='ew')
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.tree.tag_configure('amarelo', background='#FFF2CC')
        self.tree.tag_configure('total', font=('Arial', 9, 'bold'))
        self.tree.tag_configure('alerta', background='#F8CBAD')
        self.linhas = []
        self._ordem = (None, False)

    def preencher(self, linhas, tags=None):
        self.linhas = list(linhas)
        self._tags = tags
        self._mostrar()

    def _mostrar(self):
        # coluna só com números (contagens) fica alinhada à direita, como no Excel
        for _, k, _, p in self.colunas:
            vals = [r.get(k) for r in self.linhas if r.get(k) not in (None, '')]
            numerica = p or (vals and all(isinstance(v, (int, float)) for v in vals))
            self.tree.column(k, anchor='e' if numerica else 'w')
        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(self.linhas):
            vals = [fmt(r.get(k), p) for _, k, _, p in self.colunas]
            self.tree.insert('', 'end', iid=str(i), values=vals, tags=self._tags(r) if self._tags else ())

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


class App(tk.Tk):
    def __init__(self, caminho_banco=None):
        super().__init__()
        self.title(TITULO)
        self.geometry('1280x760')
        self.minsize(900, 560)
        estilo = ttk.Style(self)
        if 'vista' in estilo.theme_names():
            estilo.theme_use('vista')
        estilo.configure('Treeview', rowheight=22)
        self.banco = None
        self.R = None
        self._montar()
        self.protocol('WM_DELETE_WINDOW', self.sair)
        self.abrir_banco(caminho_banco or configuracao.banco_padrao())

    # ================================================================== montagem
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
        imp = tk.Menu(menu, tearoff=0)
        imp.add_command(label='LI – base SOP/SSOP…', command=self.importar_li)
        imp.add_separator()
        for t in servico.TIPOS_FIC:
            imp.add_command(label=servico.NOMES[t] + '…', command=lambda t=t: self.importar_fic(t))
        imp.add_separator()
        imp.add_command(label='Listas oficiais de status (STATUS_*.xlsx)…', command=self.importar_status)
        menu.add_cascade(label='Importar', menu=imp)
        aj = tk.Menu(menu, tearoff=0)
        aj.add_command(label='Como usar', command=self.ajuda)
        aj.add_command(label='Sobre', command=lambda: messagebox.showinfo(
            'Sobre', f'{TITULO}\nVersão {__version__}\n\nBanco de dados: {self.banco.caminho if self.banco else "-"}'))
        menu.add_cascade(label='Ajuda', menu=aj)
        self.config(menu=menu)

        topo = ttk.Frame(self, padding=(8, 6))
        topo.pack(fill='x')
        ttk.Button(topo, text='Importar Cadastro Geral (FVI)', command=lambda: self.importar_fic('FVI')).pack(side='left')
        ttk.Button(topo, text='Importar FVM', command=lambda: self.importar_fic('FVM')).pack(side='left', padx=(4, 0))
        ttk.Button(topo, text='Importar CCM', command=lambda: self.importar_fic('CCM')).pack(side='left', padx=(4, 0))
        ttk.Button(topo, text='Importar TAP', command=lambda: self.importar_fic('TAP')).pack(side='left', padx=(4, 0))
        ttk.Button(topo, text='Importar TTAS', command=lambda: self.importar_fic('TTAS')).pack(side='left', padx=(4, 0))
        ttk.Button(topo, text='Exportar Excel', command=self.exportar).pack(side='left', padx=(12, 0))
        self.lbl_info = ttk.Label(topo, text='', foreground='#1F4E78')
        self.lbl_info.pack(side='right')

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill='both', expand=True, padx=8, pady=(0, 4))
        self.status = ttk.Label(self, text='', anchor='w', relief='sunken', padding=(6, 2))
        self.status.pack(fill='x', side='bottom')

        # Visão gerencial
        f = ttk.Frame(self.nb, padding=6)
        self.nb.add(f, text='Visão Gerencial')
        self.lbl_visao = ttk.Label(f, text='', font=('Arial', 10, 'bold'))
        self.lbl_visao.pack(anchor='w', pady=(0, 4))
        self.t_visao = Tabela(f, exportacao.COL_VISAO, altura=16)
        self.t_visao.pack(fill='both', expand=True)
        for t in exportacao.NOTAS_VISAO:
            ttk.Label(f, text=t, font=('Arial', 8, 'italic')).pack(anchor='w')

        # Detalhes por documento
        self.detalhes = {}
        for nome, opcoes in (('FVI', {'SSOP': exportacao.COL_FVI_SSOP, 'SOP': exportacao.COL_FVI_SOP}),
                             ('FVM', {'SSOP': exportacao.COL_FVM_SSOP, 'SOP': exportacao.COL_FVM_SOP}),
                             ('CCM', {'SSOP': exportacao.COL_CCM}), ('TAP', {'SSOP': exportacao.COL_TAP}),
                             ('TTAS', {'SOP': exportacao.COL_TTAS})):
            self.detalhes[nome] = self._aba_detalhe(nome, opcoes)

        # Verificação
        f = self.f_ver = ttk.Frame(self.nb, padding=6)
        self.nb.add(f, text='Verificação')
        ttk.Label(f, text='Testes das importações (devem dar OK) e informações', font=('Arial', 10, 'bold')).pack(anchor='w')
        self.t_ver = Tabela(f, COL_VER, altura=12)
        self.t_ver.pack(fill='both', expand=True)
        ttk.Label(f, text='Pendências encontradas (linha a linha)', font=('Arial', 10, 'bold')).pack(anchor='w', pady=(6, 0))
        self.t_pend = Tabela(f, COL_PEND, altura=10)
        self.t_pend.pack(fill='both', expand=True)

        # Histórico
        f = ttk.Frame(self.nb, padding=6)
        self.nb.add(f, text='Histórico')
        barra = ttk.Frame(f)
        barra.pack(fill='x', pady=(0, 4))
        ttk.Label(barra, text='Unidade:').pack(side='left')
        self.cb_hist = ttk.Combobox(barra, state='readonly', width=14)
        self.cb_hist.pack(side='left', padx=4)
        self.cb_hist.bind('<<ComboboxSelected>>', lambda e: self._preencher_historico())
        ttk.Button(barra, text='Registrar situação atual em uma data…', command=self.registrar_historico).pack(side='left', padx=8)
        ttk.Button(barra, text='Remover data selecionada', command=self.remover_historico).pack(side='left')
        ttk.Label(f, text='Um registro por data de referência: cada importação grava o avanço na data informada '
                          '(mesma data = substitui).', font=('Arial', 8, 'italic')).pack(anchor='w')
        self.t_hist = Tabela(f, COL_HIST)
        self.t_hist.pack(fill='both', expand=True)

        # Critérios
        f = ttk.Frame(self.nb, padding=6)
        self.nb.add(f, text='Critérios')
        par = ttk.LabelFrame(f, text='Parâmetro', padding=6)
        par.pack(fill='x')
        self.var_sem_fvi = tk.StringVar(value='SIM')
        ttk.Label(par, text='Incluir TAGs sem FVI gerada no denominador do FVI?').pack(side='left')
        for v in ('SIM', 'NÃO'):
            ttk.Radiobutton(par, text=v, value=v, variable=self.var_sem_fvi, command=self.mudar_parametro).pack(side='left', padx=4)
        ttk.Label(par, text='SIM = TAG sem FVI conta como pendente (visão conservadora). NÃO = só TAGs com FVI entram nos %.',
                  font=('Arial', 8, 'italic')).pack(side='left', padx=8)
        corpo = ttk.Frame(f)
        corpo.pack(fill='both', expand=True, pady=(6, 0))
        esq = ttk.Frame(corpo)
        esq.pack(side='left', fill='both', expand=True)
        ttk.Label(esq, text='Códigos de status que contam como "etapa atingida" (amarelo = a confirmar)',
                  font=('Arial', 10, 'bold')).pack(anchor='w')
        self.t_crit = Tabela(esq, COL_CRIT)
        self.t_crit.pack(fill='both', expand=True)
        form = ttk.Frame(esq)
        form.pack(fill='x', pady=4)
        ttk.Label(form, text='Etapa').grid(row=0, column=0, sticky='w')
        ttk.Label(form, text='Código').grid(row=0, column=1, sticky='w')
        ttk.Label(form, text='Origem / justificativa').grid(row=0, column=2, sticky='w')
        self.cb_etapa = ttk.Combobox(form, values=ETAPAS, state='readonly', width=12)
        self.cb_etapa.grid(row=1, column=0, padx=(0, 4))
        self.cb_codigo = ttk.Combobox(form, width=20)
        self.cb_codigo.grid(row=1, column=1, padx=(0, 4))
        self.en_origem = ttk.Entry(form, width=50)
        self.en_origem.grid(row=1, column=2, padx=(0, 4), sticky='ew')
        form.columnconfigure(2, weight=1)
        ttk.Button(form, text='Adicionar', command=self.adicionar_criterio).grid(row=1, column=3)
        botoes = ttk.Frame(esq)
        botoes.pack(fill='x')
        ttk.Button(botoes, text='Remover selecionados', command=self.remover_criterio).pack(side='left')
        ttk.Button(botoes, text='Restaurar critérios padrão', command=self.restaurar_criterios).pack(side='left', padx=4)
        dir_ = ttk.Frame(corpo)
        dir_.pack(side='left', fill='both', padx=(8, 0))
        ttk.Label(dir_, text='Listas oficiais de status da FIC', font=('Arial', 10, 'bold')).pack(anchor='w')
        self.t_status = Tabela(dir_, COL_STATUS)
        self.t_status.pack(fill='both', expand=True)

        # Importações
        f = ttk.Frame(self.nb, padding=6)
        self.nb.add(f, text='Importações')
        ttk.Label(f, text='Registro de todas as importações gravadas neste banco', font=('Arial', 10, 'bold')).pack(anchor='w')
        self.t_imp = Tabela(f, COL_IMP)
        self.t_imp.pack(fill='both', expand=True)

    def _aba_detalhe(self, nome, opcoes):
        f = ttk.Frame(self.nb, padding=6)
        self.nb.add(f, text=nome)
        barra = ttk.Frame(f)
        barra.pack(fill='x', pady=(0, 4))
        d = {'opcoes': opcoes, 'nivel': tk.StringVar(value=next(iter(opcoes))), 'tabelas': {}}
        if len(opcoes) > 1:
            ttk.Label(barra, text='Agrupar por:').pack(side='left')
            for k in opcoes:
                ttk.Radiobutton(barra, text=k, value=k, variable=d['nivel'],
                                command=lambda n=nome: self._trocar_nivel(n)).pack(side='left', padx=2)
        ttk.Label(barra, text='Unidade:').pack(side='left', padx=(12, 0))
        d['unidade'] = ttk.Combobox(barra, state='readonly', width=14)
        d['unidade'].pack(side='left', padx=4)
        d['unidade'].bind('<<ComboboxSelected>>', lambda e, n=nome: self._preencher_detalhe(n))
        ttk.Label(barra, text='Buscar:').pack(side='left', padx=(12, 0))
        d['busca'] = tk.StringVar()
        en = ttk.Entry(barra, textvariable=d['busca'], width=30)
        en.pack(side='left', padx=4)
        en.bind('<KeyRelease>', lambda e, n=nome: self._preencher_detalhe(n))
        d['contagem'] = ttk.Label(barra, text='')
        d['contagem'].pack(side='right')
        for k, cols in opcoes.items():
            d['tabelas'][k] = Tabela(f, cols)
        d['tabelas'][d['nivel'].get()].pack(fill='both', expand=True)
        return d

    def _trocar_nivel(self, nome):
        d = self.detalhes[nome]
        for t in d['tabelas'].values():
            t.pack_forget()
        d['tabelas'][d['nivel'].get()].pack(fill='both', expand=True)
        self._preencher_detalhe(nome)

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
            self.status_msg('Banco novo: comece por Importar > LI – base SOP/SSOP.')

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
            messagebox.showwarning('Base ausente', 'Importe primeiro a LI (Importar > LI – base SOP/SSOP).')
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
                extra += '\n\nAtenção – testes com VERIFICAR (ver aba Verificação):\n- ' + '\n- '.join(alerta)
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
        data = servico.data_br(b.data_referencia_atual())
        self.lbl_info.config(text=f'Banco: {b.caminho.name}   ·   Data de referência: {data or "-"}')
        self.title(f'{TITULO} – {b.caminho}')
        self.lbl_visao.config(text=f'Avanço por unidade – data de referência {data or "-"}   ·   '
                                   f'TAGs sem FVI no denominador: {"SIM" if b.incluir_sem_fvi else "NÃO"}')
        self.t_visao.preencher(self.R.visao + [self.R.total], lambda r: ('total',) if r['unidade'] == 'TOTAL' else ())
        unidades = ['(todas)'] + self.R.unidades
        for nome in self.detalhes:
            cb = self.detalhes[nome]['unidade']
            cb['values'] = unidades
            if cb.get() not in unidades:
                cb.set('(todas)')
            self._preencher_detalhe(nome)
        self.t_ver.preencher([{'g': g, 't': t, 'e': '-' if e is None else e, 'a': a, 'r': r} for g, t, e, a, r in self.R.verificacao],
                             lambda r: ('alerta',) if r['r'] == 'VERIFICAR' else ())
        self.t_pend.preencher([dict(zip('dlisp', p)) for p in self.R.pendencias])
        self.cb_hist['values'] = ['TOTAL', '(todas)'] + self.R.unidades
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
        n_alerta = sum(1 for *_, r in self.R.verificacao if r == 'VERIFICAR')
        self.nb.tab(self.f_ver, text=f'Verificação ({n_alerta} ⚠)' if n_alerta else 'Verificação')

    def _preencher_detalhe(self, nome):
        d = self.detalhes[nome]
        nivel = d['nivel'].get()
        linhas = self.R.ssop if nivel == 'SSOP' else self.R.sop
        u = d['unidade'].get()
        if u and u != '(todas)':
            linhas = [r for r in linhas if r['unidade'] == u]
        busca = d['busca'].get().strip().upper()
        if busca:
            linhas = [r for r in linhas if busca in f'{r.get("sop", "")} {r.get("ssop", "")} {r.get("descricao", "")}'.upper()]
        d['tabelas'][nivel].preencher(linhas)
        d['contagem'].config(text=f'{len(linhas)} {nivel}')

    def _preencher_historico(self):
        u = self.cb_hist.get()
        linhas = servico.historico_com_percentuais(self.banco)
        if u and u != '(todas)':
            linhas = [r for r in linhas if r['unidade'] == u]
        self.t_hist.preencher(linhas, lambda r: ('total',) if r['unidade'] == 'TOTAL' else ())

    def status_msg(self, msg):
        self.status.config(text=msg)

    def ajuda(self):
        messagebox.showinfo('Como usar', (
            '1. Importar > LI – base SOP/SSOP (uma vez, ou quando a LI for revisada).\n'
            '2. Importar > Relatório de Cadastro Geral (FVI): o arquivo exportado da FIC, sem editar.\n'
            '   FVM, CCM, TAP e TTAS: planilhas com cabeçalho reconhecido pelo nome das colunas\n'
            '   (Arquivo > Exportar planilhas-modelo mostra o formato).\n'
            '3. Confirme a data de referência: o avanço fica registrado no Histórico nessa data.\n'
            '4. Confira a aba Verificação (testes devem dar OK) antes de apresentar.\n'
            '5. Arquivo > Exportar relatório Excel gera a planilha gerencial para envio.\n\n'
            'Cada importação substitui a anterior do mesmo tipo (a FIC exporta sempre a lista completa).\n'
            'Critérios de "aprovado" e o parâmetro de TAGs sem FVI ficam na aba Critérios.\n'
            'Faça cópias de segurança do banco em Arquivo > Salvar cópia de segurança.'))

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
