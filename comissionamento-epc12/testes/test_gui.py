"""Fluxos da interface com os diálogos simulados (arquivo escolhido, data confirmada, mensagens capturadas)."""
import pytest

tk = pytest.importorskip('tkinter')
try:
    _raiz = tk.Tk()
    _raiz.destroy()
except tk.TclError:
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
    a = gui.App(tmp_path / 'gui.db')
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
    assert app.msgs[-1][1] == 'LI importada'

    escolher(monkeypatch, f.cadastro(tmp_path / 'cad.xlsx', f.TAGS_SINTETICAS))
    app.importar_fic('FVI')
    tipo, titulo, msg = app.msgs[-1]
    assert titulo == 'Importação concluída' and '10 registro(s) de FVI' in msg and 'Verificação' in msg
    assert app.selo_ver.cget('text').strip() == '3'           # fora da base, código desconhecido, Campo IV

    # FVI: TAG, SSOP e a situação de cada campo por extenso (sem coluna Situação)
    t1 = next(r for r in app.fvi['tabela'].linhas if r['tag'] == 'T1')
    assert (t1['ssop'], t1['txt_cii'], t1['txt_ciii'], t1['txt_cv']) == (
        f.S1, 'Aprovado', 'Aprovado (ACAMPO5)', 'Pendente (AAPROV)')
    t = {r['tag']: r for r in app.R.tags}
    assert (t['T3']['txt_cii'], t['T4']['txt_cv'], t['T5']['txt_cii'], t['T6']['txt_situacao']) == (
        'Não se aplica', 'Pendente (sem status)', 'Cancelada', 'Sem FVI gerada')
    assert (t['T9']['txt_situacao'], t['T10']['txt_situacao']) == ('SSOP fora da base', 'Sem SSOP')
    assert [c[0] for c in gui.rotulos.COL_TAGS] == ['TAG', 'SSOP', 'FVI Campo II', 'FVI Campo III', 'FVI Campo V']
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
    assert [c[0] for c in gui.rotulos.COL_CCM] == ['Unidade', 'SOP', 'SSOP', 'Descrição da SSOP', 'Situação do CCM']
    d = app.doc_ccm
    d['situacao'].set('Aprovado')
    app._preencher(d)
    assert [r['ssop'] for r in d['tabela'].linhas] == [f.S1]
    d['situacao'].set('Em andamento')
    app._preencher(d)
    assert [r['ssop'] for r in d['tabela'].linhas] == [f.S2]
    app._limpar(d)
    assert len(d['tabela'].linhas) == 4
    escolher(monkeypatch, f.provisorio(tmp_path / 'fvm.xlsx', ['MALHA', 'SOP', 'SSOP', 'STATUS'], f.FVM_SINTETICO), '28/10/2026')
    app.importar_fic('FVM')
    malhas = {m['malha']: m for m in app.fvm['tabela'].linhas}
    assert (malhas['M1']['txt_ciii'], malhas['M2']['txt_cv'], malhas['M3']['txt_ciii'], malhas['M4']['txt_ciii']) == (
        'Aprovado (ACAMPO5)', 'Aprovado', 'Pendente (AAPRV3E4FIS)', 'Cancelada')
    assert [c[0] for c in gui.rotulos.COL_MALHAS] == ['Malha', 'SSOP', 'FVM Campo III', 'FVM Campo V']
    app.fvm['busca'].insert(0, 'M5')
    app._preencher(app.fvm)
    assert [m['malha'] for m in app.fvm['tabela'].linhas] == ['M5']
    # TTAS por SOP: colunas renomeadas e filtro por SSOP encontra o SOP dela
    escolher(monkeypatch, f.provisorio(tmp_path / 'ttas.xlsx', ['Nº TTAS', 'SOP', 'STATUS'], f.TTAS_SINTETICO), '28/10/2026')
    app.importar_fic('TTAS')
    assert [c[0] for c in gui.rotulos.COL_TTAS] == ['Unidade', 'SOP', 'Descrição do SOP', 'Situação TTAS 1',
                                                     'Situação TTAS 2']
    tt = {r['sop']: r for r in app.doc_ttas['tabela'].linhas}
    assert (tt[f.P1]['txt_ttas1'], tt[f.P1]['txt_ttas2'], tt[f.P2]['txt_ttas2'], tt[f.P3]['txt_ttas1']) == (
        'Aprovado', 'Aprovado', 'Em andamento', 'Sem registro')
    app.doc_ttas['ssop'].set(f.S3)
    app._preencher(app.doc_ttas)
    assert [r['sop'] for r in app.doc_ttas['tabela'].linhas] == [f.P2]
    app.doc_ttas['ssop'].set(gui.TODAS_SSOP)
    app.doc_ttas['situacao'].set('TTAS 2 aprovado')
    app._preencher(app.doc_ttas)
    assert [r['sop'] for r in app.doc_ttas['tabela'].linhas] == [f.P1]
    for nome in gui.PAGINAS:                                  # todas as páginas abrem
        app.mostrar(nome)
        app.update()
    assert [r['data_referencia'] for r in app.t_hist.linhas] == ['2026-10-28']   # TOTAL de 28/10
    assert [r['tipo'] for r in app.t_imp.linhas] == ['TTAS', 'FVM', 'CCM', 'FVI', 'LI']


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


def test_visao_gerencial_so_graficos(app, tmp_path, monkeypatch):
    carregar_fvi(app, tmp_path, monkeypatch)
    app.mostrar('Visão Gerencial')
    app.update()
    assert set(app.g) == {'fvi', 'fvm', 'ccm', 'tap', 'ttas'}
    assert not hasattr(app, 'ind') and not hasattr(app, 't_visao')
    g = app.g['fvi']
    assert g._dicas, 'colunas do gráfico sem dica'
    dicas = list(g._dicas.values())
    assert any(t.startswith('U-9001\nFVI Campo II: 66,7%\n4 de 6') for t in dicas)
    assert any(t.startswith('Total\nFVI Campo II: 66,7%') for t in dicas)          # total no mesmo gráfico
    item = next(iter(g._dicas))
    x0, y0, x1, y1 = g.coords(item)

    class Ev:
        x, y = int((x0 + x1) / 2), int((y0 + y1) / 2)
    g._mover(Ev)
    assert g.find_withtag('dica'), 'dica não apareceu'


def test_menu_importar_e_logo(app):
    m = app.menu_importar()
    assert [m.entrycget(i, 'label') for i in range(m.index('end') + 1)] == ['SSOP', 'FVI', 'FVM', 'CCM', 'TAP', 'TTAS']
    assert app.logo.cget('image') is not None and app.logo.cget('text') == ''       # logo embutida, não o quadrado
