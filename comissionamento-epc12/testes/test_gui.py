"""Fluxos da interface com os diálogos simulados (arquivo escolhido, data confirmada, mensagens capturadas)."""
import pytest

import time

tk = pytest.importorskip('tkinter')


def criar_tk(fabrica, tentativas=3):
    """Cria a janela Tk. Nos runners Windows do GitHub (Python do hostedtoolcache) o Tcl às vezes falha ao ler o próprio
    init.tcl ("couldn't read file ... init.tcl": No error) – falha conhecida do ambiente, não do programa
    (microsoft/azure-pipelines-tasks#16426). Só esse erro é tentado de novo; qualquer outro sobe na hora."""
    for i in range(tentativas):
        try:
            return fabrica()
        except tk.TclError as e:
            if 'init.tcl' not in str(e) or i == tentativas - 1:
                raise
            time.sleep(0.5)


try:
    criar_tk(tk.Tk).destroy()
except tk.TclError as e:
    if 'init.tcl' in str(e):
        raise                                   # Tcl quebrado não é "sem display": não pula os testes em silêncio
    pytest.skip('sem display para o Tk', allow_module_level=True)

from comissionamento import gui  # noqa: E402
from testes import fabrica as f  # noqa: E402


@pytest.fixture
def app(tmp_path, monkeypatch):
    msgs = []
    monkeypatch.setattr(gui.configuracao, 'pasta_usuario', lambda: tmp_path)
    for nome in ('showinfo', 'showerror', 'showwarning'):
        monkeypatch.setattr(gui.messagebox, nome, lambda t, m, nome=nome, **k: msgs.append((nome, t, m)))
    monkeypatch.setattr(gui.messagebox, 'askyesno', lambda *a, **k: False)
    a = criar_tk(lambda: gui.App(tmp_path / 'gui.db'))
    a.msgs = msgs
    yield a
    a.sair()


def escolher(monkeypatch, arquivo, data=None):
    monkeypatch.setattr(gui.filedialog, 'askopenfilename', lambda **k: str(arquivo))
    monkeypatch.setattr(gui.simpledialog, 'askstring', lambda *a, **k: data or k.get('initialvalue'))


def carregar_fvi(app, tmp_path, monkeypatch):
    escolher(monkeypatch, f.li(tmp_path / 'li.xlsx'))
    app.importar_li()
    escolher(monkeypatch, f.cadastro(tmp_path / 'cad.xlsx', f.TAGS_SINTETICAS))
    app.importar_fic('FVI')


def tags_na_tela(app):
    return [r['tag'] for r in app.fvi['tabela'].linhas]


def test_fluxo_completo(app, tmp_path, monkeypatch):
    app.importar_fic('FVI')                                   # sem LI: avisa e não pede arquivo
    assert app.msgs[-1][0] == 'showwarning'

    escolher(monkeypatch, f.li(tmp_path / 'li.xlsx'))
    app.importar_li()
    assert app.msgs[-1][1] == 'SOP / SSOP importada'

    escolher(monkeypatch, f.cadastro(tmp_path / 'cad.xlsx', f.TAGS_SINTETICAS))
    app.importar_fic('FVI')
    tipo, titulo, msg = app.msgs[-1]
    assert titulo == 'Importação concluída' and '10 registro(s) de FVI' in msg and 'Verificação' in msg
    assert app.selo_ver.cget('text').strip() == '3'           # fora da base, código desconhecido, Campo IV

    # FVI: TAG, SSOP e o status de cada campo exatamente como vem da FIC
    assert [c[0] for c in gui.rotulos.COL_TAGS] == ['TAG', 'SSOP', 'FVI Campo II', 'FVI Campo III', 'FVI Campo V']
    t1 = next(r for r in app.fvi['tabela'].linhas if r['tag'] == 'T1')
    assert [t1[k] for _, k, _, _ in gui.rotulos.COL_TAGS] == ['T1', f.S1, 'APROVADO', 'ACAMPO5', 'AAPROV']
    t3 = next(r for r in app.fvi['tabela'].linhas if r['tag'] == 'T3')
    assert (t3['st_cii'], t3['st_cv']) == ('N/A', 'N/A')
    assert 'situacao' not in app.fvi                          # filtros: unidade, SOP, SSOP e TAG
    # cascata: unidade limita os SOP, SOP limita as SSOP
    app.fvi['unidade'].set(f.UN1)
    app._mudou_unidade(app.fvi)
    assert app.fvi['sop'].cget('values') == [gui.TODOS_SOP, f.P1, f.P2]
    assert sorted(tags_na_tela(app)) == ['T1', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'T8']
    app.fvi['sop'].set(f.P2)
    app._mudou_sop(app.fvi)
    assert app.fvi['ssop'].cget('values') == [gui.TODAS_SSOP, f.S2, f.S3]
    assert sorted(tags_na_tela(app)) == ['T7', 'T8']
    app.fvi['ssop'].set(f.S2)
    app._preencher(app.fvi)
    assert tags_na_tela(app) == ['T7']
    assert app.fvi['contagem'].cget('text') == '1 TAGs'
    app._limpar(app.fvi)
    app.fvi['busca'].insert(0, 't1')                          # busca por TAG (T1 e T10)
    app._preencher(app.fvi)
    assert sorted(tags_na_tela(app)) == ['T1', 'T10']
    app._limpar(app.fvi)
    assert len(tags_na_tela(app)) == 10

    escolher(monkeypatch, f.provisorio(tmp_path / 'ccm.xlsx', ['Nº CCM', 'SSOP', 'STATUS'], f.CCM_SINTETICO), '28/10/2026')
    app.importar_fic('CCM')
    assert app.R.total['ccm'] == 1
    assert [c[0] for c in gui.rotulos.COL_CCM] == ['Unidade', 'SOP', 'SSOP', 'Descrição da SSOP', 'Status do CCM']
    d = app.doc_ccm
    # status da FIC por SSOP: S1 tem CANCELADO + CCMCONCLUIDO (vale o aprovado); S2 AGAPROVA; sem documento = Sem registro
    assert {r['ssop']: r['txt_ccm_status'] for r in d['tabela'].linhas} == {
        f.S1: 'CCMCONCLUIDO', f.S2: 'AGAPROVA', f.S3: 'Sem registro', f.S4: 'Sem registro'}
    assert d['situacao'].cget('values') == [gui.TODAS_SITUACOES, 'AGAPROVA', 'CCMCONCLUIDO', 'Sem registro']
    d['situacao'].set('CCMCONCLUIDO')
    app._preencher(d)
    assert [r['ssop'] for r in d['tabela'].linhas] == [f.S1]
    d['situacao'].set('AGAPROVA')
    app._preencher(d)
    assert [r['ssop'] for r in d['tabela'].linhas] == [f.S2]
    app._limpar(d)
    assert len(d['tabela'].linhas) == 4
    # FVM: M6 aparece em duas SSOP (pode, mas não deve) – entra nas duas e é sinalizada
    fvm = f.FVM_SINTETICO + [('M6', f.P1, f.S1, 'ACAMPO5'), ('M6', f.P2, f.S2, 'ACAMPO5')]
    base, lista = f.fvm(tmp_path, fvm)
    escolher(monkeypatch, lista, '28/10/2026')
    app.importar_fic('FVM')                                   # sem FVM BASE: recusa e não grava
    assert app.msgs[-1][:2] == ('showerror', 'Importação não realizada') and app.banco.dados('FVM') == []
    escolher(monkeypatch, base)
    app.importar_fvm_base()
    assert app.msgs[-1][1] == 'FVM BASE importada'
    escolher(monkeypatch, lista, '28/10/2026')
    app.importar_fic('FVM')
    assert [c[0] for c in gui.rotulos.COL_MALHAS] == ['Malha', 'TAG', 'SSOP', 'Status da FVM']
    malhas = {(m['malha'], m['ssop']): m for m in app.fvm['tabela'].linhas}
    assert malhas[('M6', f.S2)]['tag'] == 'T-M6-2'
    assert (malhas[('M1', f.S1)]['status'], malhas[('M3', f.S1)]['status'], malhas[('M4', f.S1)]['status']) == (
        'ACAMPO5', 'AAPRV3E4FIS', 'CANCEL')
    assert malhas[('M6', f.S1)]['multi_ssop'] and malhas[('M6', f.S2)]['multi_ssop'] and not malhas[('M1', f.S1)]['multi_ssop']
    assert app._cor_linha(malhas[('M6', f.S2)]) == ('aviso',)
    assert (next(r for r in app.R.ssop if r['ssop'] == f.S1)['malhas'], next(r for r in app.R.ssop if r['ssop'] == f.S2)['malhas']) == (5, 1)
    ver = {t: (a, r) for _, t, _, a, r in app.R.verificacao}
    assert ver['FVM – malhas em mais de uma SSOP (pode, mas não deve)'] == (1, 'VERIFICAR')
    app.fvm['busca'].insert(0, 'M5')
    app._preencher(app.fvm)
    assert [m['malha'] for m in app.fvm['tabela'].linhas] == ['M5']
    app.fvm['busca'].delete(0, 'end')
    app.fvm['busca'].insert(0, 'T-M6-2')                      # a busca também acha pela TAG
    app._preencher(app.fvm)
    assert [(m['malha'], m['ssop']) for m in app.fvm['tabela'].linhas] == [('M6', f.S2)]
    app.fvm['busca'].delete(0, 'end')
    # TTAS por SOP: colunas renomeadas e filtro por SSOP encontra o SOP dela
    escolher(monkeypatch, f.provisorio(tmp_path / 'ttas.xlsx', ['Nº TTAS', 'SOP', 'STATUS'], f.TTAS_SINTETICO), '28/10/2026')
    app.importar_fic('TTAS')
    assert [c[0] for c in gui.rotulos.COL_TTAS] == ['Unidade', 'SOP', 'Descrição do SOP', 'Situação TTAS 1',
                                                     'Situação TTAS 2']
    tt = {r['sop']: (r['ttas1_status'], r['ttas2_status']) for r in app.doc_ttas['tabela'].linhas}
    # registro único que avança: em etapa de TTAS 2, a coluna TTAS 1 mostra TTAS1APROVADO
    assert tt == {f.P1: ('TTAS1APROVADO', 'TTASCONCLUIDO'), f.P2: ('TTAS1APROVADO', '—'),
                  f.P3: ('Sem registro', 'Sem registro')}
    app.doc_ttas['ssop'].set(f.S3)
    app._preencher(app.doc_ttas)
    assert [r['sop'] for r in app.doc_ttas['tabela'].linhas] == [f.P2]
    app.doc_ttas['ssop'].set(gui.TODAS_SSOP)
    app.doc_ttas['situacao'].set('TTASCONCLUIDO')
    app._preencher(app.doc_ttas)
    assert [r['sop'] for r in app.doc_ttas['tabela'].linhas] == [f.P1]
    for nome in gui.PAGINAS:                                  # todas as páginas abrem
        app.mostrar(nome)
        app.update()
    assert [r['data_referencia'] for r in app.t_hist.linhas] == ['2026-10-28']   # TOTAL de 28/10
    assert [r['tipo'] for r in app.t_imp.linhas] == ['TTAS', 'FVM', 'FVM_BASE', 'CCM', 'FVI', 'LI']


def test_arquivo_errado_nao_grava(app, tmp_path, monkeypatch):
    escolher(monkeypatch, f.li(tmp_path / 'li.xlsx'))
    app.importar_li()
    escolher(monkeypatch, f.provisorio(tmp_path / 'x.xlsx', ['A', 'B'], [('1', '2')]))
    app.importar_fic('TAP')
    assert app.msgs[-1][:2] == ('showerror', 'Importação não realizada')
    assert app.banco.dados('TAP') == []


def test_parametro_e_criterio(app, tmp_path, monkeypatch):
    carregar_fvi(app, tmp_path, monkeypatch)
    app.seg_sem_fvi.set('Não')
    app.mudar_parametro()
    s1 = next(r for r in app.R.ssop if r['ssop'] == f.S1)
    assert s1['cii_aplic'] == 3
    app.cb_etapa.set('FVI_CII')
    app.cb_codigo.set('AAPROV')
    app.adicionar_criterio()
    assert next(r for r in app.R.ssop if r['ssop'] == f.S1)['cii_aprov'] == 3


def test_visao_gerencial_pizzas(app, tmp_path, monkeypatch):
    carregar_fvi(app, tmp_path, monkeypatch)
    app.mostrar('Visão Gerencial')
    app.update()
    assert list(app.pizzas) == ['FVI', 'FVM', 'CCM', 'TAP', 'TTAS']        # ordem pedida
    alturas = {c.winfo_height() for c in (p.master for p in app.pizzas.values())}
    assert len(alturas) == 1                                                  # caixas do mesmo tamanho
    fvi = app.pizzas['FVI']
    # Status FVI das 10 TAGs: APROVADO 3, ACAMPO3E4 2, demais 1 (vazio = "Sem status")
    assert [(r, n) for r, n, _, _ in fvi.fatias] == [('APROVADO', 3), ('ACAMPO3E4', 2), ('ACAMPO5', 1),
                                                     ('ALIBTFUNC', 1), ('CANCELADO', 1), ('Sem status', 1), ('ARECEB', 1)]
    assert fvi.fatias[0][3] == 'Aprovado'                                     # descrição da lista oficial
    assert len({c for _, _, c, _ in fvi.fatias}) == len(fvi.fatias)           # uma cor por status
    cx, cy, r, cortes = fvi._geo
    import math
    a = math.radians(cortes[0][0] + 5)
    assert fvi.fatia_em(cx + r / 2 * math.cos(a), cy + r / 2 * math.sin(a)) == 0
    assert fvi.fatia_em(cx + r * 2, cy) is None
    app.seg_fvi.set('Campo II')
    app._atualizar_visao()
    assert fvi.fatias[0][:2] == ('APROVADO', 5)
    app.cb_visao.set(f.UN1)
    app._atualizar_visao()
    assert '8 TAGs' in app.sub_pizza['FVI'].cget('text')
    assert [(r, n) for r, n, _, _ in app.pizzas['CCM'].fatias] == [('Sem registro', 3)]


def test_menus_de_importacao_e_logo(app):
    assert not hasattr(app, 'btn_bases')                                            # só o botão de atualização
    m = app.menu_importar()
    rotulos_menu = [m.entrycget(i, 'label') for i in range(m.index('end') + 1) if m.type(i) == 'command']
    assert rotulos_menu == ['SOP / SSOP', 'FVM BASE', 'Templates de FVI / FVM (PDF)', 'FVI', 'FVM', 'CCM', 'TAP', 'TTAS']
    assert app.logo.cget('image') is not None and app.logo.cget('text') == ''       # logo embutida, não o quadrado


def test_modelos_da_fic_pela_tela(app, tmp_path, monkeypatch):
    escolher(monkeypatch, f.cadastros_basicos(tmp_path / 'b.xlsx', f.BASICOS))
    app.importar_li()
    assert app.msgs[-1][1] == 'SOP / SSOP importada'
    escolher(monkeypatch, f.hierarquia_malhas(tmp_path / 'fvm.xlsx', [
        ('M1', 'OPERAÇÃO', [('TAG', 'T1', f.S1, f.P1)]),
        ('M2', 'PRONTO', [('TAG', 'T2', f.S1, f.P1), ('TAG', 'T3', f.S2, f.P2)])]))
    app.importar_fvm_base()
    assert len(app.fvm['tabela'].linhas) == 3                  # tabela: uma linha por TAG da malha
    assert '2 malhas' in app.sub_pizza['FVM'].cget('text')     # pizza: cada malha uma vez
    assert [(r, n) for r, n, _, _ in app.pizzas['FVM'].fatias] == [('Sem status', 2)]   # status vem da atualização
    escolher(monkeypatch, f.lista_fvm(tmp_path / 'lista.xlsx', [('1', 'M1', 'APROVADO'), ('2', 'M9', 'APROVADO')]))
    app.importar_fic('FVM')
    assert sorted((r, n) for r, n, _, _ in app.pizzas['FVM'].fatias) == [('APROVADO', 1), ('Sem status', 1)]
    assert {m['malha'] for m in app.R.malhas} == {'M1', 'M2'}  # M9 não está na FVM BASE: só na Verificação


def test_clique_na_malha_mostra_as_perguntas(app, tmp_path, monkeypatch):
    escolher(monkeypatch, f.cadastros_basicos(tmp_path / 'b.xlsx', f.BASICOS))
    app.importar_li()
    base, lista = f.fvm(tmp_path, [('M1', f.P1, f.S1, 'APROVADO')])     # Lista de FVM com template FVM-INS001
    escolher(monkeypatch, base)
    app.importar_fvm_base()
    escolher(monkeypatch, lista, '30/10/2026')
    app.importar_fic('FVM')
    app.fvm['tabela'].tree.selection_set('0')
    app.folha_fvm(app.fvm['tabela'].selecionadas()[0])          # sem templates importados: explica o motivo
    assert 'não está no PDF importado' in app._janela_folha.corpo.winfo_children()[0].cget('text')
    pdf = f.templates_fvm_pdf(tmp_path / 'tpl.pdf', [('FVM-INS001', '2.0', 'INSTRUMENTAÇÃO',
                                                       [('III.1', 'LINHAS: Malha montada.'), ('V.7', 'TAGs: Teste aprovado.')])])
    monkeypatch.setattr(gui.filedialog, 'askopenfilename', lambda **k: str(pdf))
    app.importar_templates()
    assert app.msgs[-1][1] == 'Templates importados'
    app.folha_fvm(app.fvm['tabela'].linhas[0])
    textos = [w.cget('text') for lin in app._janela_folha.corpo.winfo_children() for w in lin.winfo_children()
              if isinstance(w, gui.ctk.CTkLabel)]
    assert 'LINHAS: Malha montada.' in textos and 'TAGs: Teste aprovado.' in textos
    assert 'FVM-INS001' in app._janela_folha.lbl_dados.cget('text')


def test_clique_na_tag_mostra_as_perguntas_da_classe(app, tmp_path, monkeypatch):
    escolher(monkeypatch, f.cadastros_basicos(tmp_path / 'b.xlsx', f.BASICOS))
    app.importar_li()
    escolher(monkeypatch, f.cadastro(tmp_path / 'c.xlsx', [
        ('T1', f.S1, '1', 'APROVADO', 'APROVADO', 'ACAMPO3E4', 'N/A', 'N/A', '(007-C) TRANSMISSOR DE PRESSÃO')]))
    app.importar_fic('FVI')
    linha = app.fvi['tabela'].linhas[0]
    app.folha_fvi(linha)                                      # sem templates de FVI: explica o motivo
    assert 'nenhum PDF de templates de FVI' in app._janela_folha.corpo.winfo_children()[0].cget('text')
    pdf = f.templates_fvi_pdf(tmp_path / 'fvi.pdf', [
        ('FVI-007', '(007) TRANSMISSOR DE PROCESSO', {'II': ['Recebido.'], 'III': ['Instalado.'], 'V': ['Testado.']})])
    monkeypatch.setattr(gui.filedialog, 'askopenfilename', lambda **k: str(pdf))
    app.importar_templates()
    app.folha_fvi(linha)
    filhos = app._janela_folha.corpo.winfo_children()
    assert 'A CONFIRMAR' in filhos[0].cget('text')            # família da classe: inferência sinalizada
    textos = [w.cget('text') for lin in filhos for w in lin.winfo_children() if isinstance(w, gui.ctk.CTkLabel)]
    assert {'Recebido.', 'Instalado.', 'Testado.'} <= set(textos)
    assert any('status na FIC: ACAMPO3E4' in t for t in textos)   # status do Campo III ao lado do título
