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
    assert app.t_visao.linhas[-1]['unidade'] == 'TOTAL'
    assert app.t_visao.linhas[-1]['pct_cii'] == 4 / 6
    assert app.selo_ver.cget('text').strip() == '3'           # fora da base, código desconhecido, Campo IV

    # FVI: uma visão só – TAG, SSOP e a situação de cada campo por extenso
    t1 = next(r for r in app.fvi['tabela'].linhas if r['tag'] == 'T1')
    assert (t1['ssop'], t1['txt_cii'], t1['txt_ciii'], t1['txt_cv']) == (
        f.S1, 'Aprovado', 'Aprovado (ACAMPO5)', 'Pendente (AAPROV)')
    t = {r['tag']: r for r in app.R.tags}
    assert (t['T3']['txt_cii'], t['T4']['txt_cv'], t['T5']['txt_cii'], t['T6']['txt_situacao']) == (
        'Não se aplica', 'Pendente (sem status)', 'Cancelada', 'Sem FVI gerada')
    assert (t['T9']['txt_situacao'], t['T10']['txt_situacao']) == ('SSOP fora da base', 'Sem SSOP')
    assert [c[0] for c in gui.rotulos.COL_TAGS] == ['TAG', 'SSOP', 'FVI Campo II', 'FVI Campo III',
                                                     'FVI Campo V', 'Situação']
    app.fvi['situacao'].set('Canceladas')
    app._preencher(app.fvi)
    assert tags_na_tela(app) == ['T5']
    app.fvi['situacao'].set('Todas as situações')
    app.fvi['busca'].insert(0, f.S2)
    app._preencher(app.fvi)
    assert tags_na_tela(app) == ['T7']
    assert app.fvi['contagem'].cget('text') == '1 TAGs'
    app._limpar(app.fvi)
    assert len(tags_na_tela(app)) == 10

    escolher(monkeypatch, f.provisorio(tmp_path / 'ccm.xlsx', ['Nº CCM', 'SSOP', 'STATUS'], f.CCM_SINTETICO), '28/10/2026')
    app.importar_fic('CCM')
    assert app.R.total['ccm'] == 1
    s1 = next(r for r in app.doc_ccm['tabela'].linhas if r['ssop'] == f.S1)
    assert (s1['txt_ccm_aprovado'], s1['txt_ccm_situacao']) == ('Sim', 'Aprovado')
    escolher(monkeypatch, f.provisorio(tmp_path / 'fvm.xlsx', ['MALHA', 'SOP', 'SSOP', 'STATUS'], f.FVM_SINTETICO), '28/10/2026')
    app.importar_fic('FVM')
    malhas = {m['malha']: m for m in app.fvm['tabela'].linhas}
    assert (malhas['M1']['txt_ciii'], malhas['M2']['txt_cv'], malhas['M3']['txt_ciii'], malhas['M4']['txt_situacao']) == (
        'Aprovado (ACAMPO5)', 'Aprovado', 'Pendente (AAPRV3E4FIS)', 'Cancelada')
    for nome in gui.PAGINAS:                                  # todas as páginas abrem
        app.mostrar(nome)
        app.update()
    assert [r['data_referencia'] for r in app.t_hist.linhas] == ['2026-10-28']   # TOTAL de 28/10
    assert [r['tipo'] for r in app.t_imp.linhas] == ['FVM', 'CCM', 'FVI', 'LI']


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


def test_visao_gerencial_indicadores_e_graficos(app, tmp_path, monkeypatch):
    carregar_fvi(app, tmp_path, monkeypatch)
    app.mostrar('Visão Gerencial')
    app.update()
    assert set(app.ind) == {'cii', 'ciii', 'cv', 'fvm_ciii', 'fvm_cv', 'ccm', 'tap', 'ttas1', 'ttas2'}
    assert app.ind['cii'].anel.cget('text') == '66,7%'
    assert app.ind['cii'].detalhe.cget('text') == '4 de 6 TAGs aplicáveis aprovadas'
    assert app.resumo['tags'].cget('text') == '7'
    assert set(app.g) == {'geral', 'fvi', 'fvm', 'ccm', 'ttas', 'evol'}
    g = app.g['fvi']
    assert g._dicas, 'colunas do gráfico sem dica'
    assert any('FVI Campo II: 66,7%' in t for t in g._dicas.values())        # U-9001: 4 de 6
    item = next(iter(g._dicas))
    x0, y0, x1, y1 = g.coords(item)

    class Ev:
        x, y = int((x0 + x1) / 2), int((y0 + y1) / 2)
    g._mover(Ev)
    assert g.find_withtag('dica'), 'dica não apareceu'
    geral = ' '.join(app.g['geral']._dicas.values())
    for nome in ('FVI Campo II', 'FVI Campo V', 'FVM Campo III', 'FVM Campo V', 'CCM', 'TAP', 'TTAS 1', 'TTAS 2'):
        assert nome in geral
