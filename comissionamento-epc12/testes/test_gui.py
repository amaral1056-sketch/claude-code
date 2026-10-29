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
    a.withdraw()
    a.msgs = msgs
    yield a
    a.sair()


def escolher(monkeypatch, arquivo, data=None):
    monkeypatch.setattr(gui.filedialog, 'askopenfilename', lambda **k: str(arquivo))
    monkeypatch.setattr(gui.simpledialog, 'askstring', lambda *a, **k: data or k.get('initialvalue'))


def test_fluxo_completo(app, tmp_path, monkeypatch):
    app.importar_fic('FVI')                                   # sem LI: avisa e não pede arquivo
    assert app.msgs[-1][0] == 'showwarning'

    escolher(monkeypatch, f.li(tmp_path / 'li.xlsx'))
    app.importar_li()
    assert app.msgs[-1][1] == 'LI importada'

    escolher(monkeypatch, f.cadastro(tmp_path / 'cad.xlsx', f.TAGS_SINTETICAS))
    app.importar_fic('FVI')
    tipo, titulo, msg = app.msgs[-1]
    assert titulo == 'Importação concluída' and '10 registro(s) de FVI' in msg and 'VERIFICAR' in msg
    assert app.t_visao.linhas[-1]['unidade'] == 'TOTAL'
    assert app.t_visao.linhas[-1]['pct_cii'] == 4 / 6
    assert '⚠ 3' in app.nav['Verificação'].cget('text')       # fora da base, código desconhecido, Campo IV
    # visão TAG a TAG: status de cada campo e marca de aprovado/pendente/N/A
    t1 = next(r for r in app.fvi['tabelas']['TAGs'].linhas if r['tag'] == 'T1')
    assert (t1['d_cii'], t1['d_ciii'], t1['d_cv'], t1['ssop']) == ('✔ APROVADO', '✔ ACAMPO5', '○ AAPROV', f.S1)
    t3 = next(r for r in app.R.tags if r['tag'] == 'T3')
    assert t3['d_cii'] == '— N/A'
    app.fvi['situacao'].set('Canceladas')
    app._preencher(app.fvi)
    assert [r['tag'] for r in app.fvi['tabelas']['TAGs'].linhas] == ['T5']
    app._abrir_tags(app.fvi, f.S2)                            # duplo clique numa SSOP
    assert [r['tag'] for r in app.fvi['tabelas']['TAGs'].linhas] == ['T7']

    escolher(monkeypatch, f.provisorio(tmp_path / 'ccm.xlsx', ['Nº CCM', 'SSOP', 'STATUS'], f.CCM_SINTETICO), '28/10/2026')
    app.importar_fic('CCM')
    assert app.R.total['ccm'] == 1
    escolher(monkeypatch, f.provisorio(tmp_path / 'fvm.xlsx', ['MALHA', 'SOP', 'SSOP', 'STATUS'], f.FVM_SINTETICO), '28/10/2026')
    app.importar_fic('FVM')
    malhas = {m['malha']: m for m in app.fvm['tabelas']['Malhas'].linhas}
    assert (malhas['M1']['d_ciii'], malhas['M2']['d_cv'], malhas['M4']['situacao']) == ('✔ atingido', '✔ atingido', 'CANCELADA')
    for nome in gui.PAGINAS:                                  # todas as páginas abrem
        app.mostrar(nome)
        app.update()
    assert [r['data_referencia'] for r in app.t_hist.linhas] == ['2026-10-28']   # só a linha TOTAL de 28/10
    assert [r['tipo'] for r in app.t_imp.linhas] == ['FVM', 'CCM', 'FVI', 'LI']


def test_arquivo_errado_nao_grava(app, tmp_path, monkeypatch):
    escolher(monkeypatch, f.li(tmp_path / 'li.xlsx'))
    app.importar_li()
    escolher(monkeypatch, f.provisorio(tmp_path / 'x.xlsx', ['A', 'B'], [('1', '2')]))
    app.importar_fic('TAP')
    assert app.msgs[-1][:2] == ('showerror', 'Importação não realizada')
    assert app.banco.dados('TAP') == []


def test_parametro_e_criterio(app, tmp_path, monkeypatch):
    escolher(monkeypatch, f.li(tmp_path / 'li.xlsx'))
    app.importar_li()
    escolher(monkeypatch, f.cadastro(tmp_path / 'cad.xlsx', f.TAGS_SINTETICAS))
    app.importar_fic('FVI')
    app.var_sem_fvi.set('NÃO')
    app.mudar_parametro()
    s1 = next(r for r in app.R.ssop if r['ssop'] == f.S1)
    assert s1['cii_aplic'] == 3
    app.cb_etapa.set('FVI_CII')
    app.cb_codigo.set('AAPROV')
    app.adicionar_criterio()
    assert next(r for r in app.R.ssop if r['ssop'] == f.S1)['cii_aprov'] == 3


def test_graficos_do_painel_e_dica(app, tmp_path, monkeypatch):
    escolher(monkeypatch, f.li(tmp_path / 'li.xlsx'))
    app.importar_li()
    escolher(monkeypatch, f.cadastro(tmp_path / 'cad.xlsx', f.TAGS_SINTETICAS))
    app.importar_fic('FVI')
    app.deiconify()
    app.mostrar('Painel')
    app.update()
    g = app.g_unid
    assert g._dicas, 'colunas do gráfico sem dica'
    item, texto = next(iter(g._dicas.items()))
    x0, y0, x1, y1 = g.coords(item)

    class Ev:
        x, y = int((x0 + x1) / 2), int((y0 + y1) / 2)
    g._mover(Ev)
    assert g.find_withtag('dica'), 'dica não apareceu'
    assert any('Campo II: 66,7%' in t for t in g._dicas.values())           # U-9001: 4 de 6
    assert app.ind['cii'].valor.cget('text') == '66,7%'
