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
    assert app.nb.tab(app.f_ver, 'text').startswith('Verificação (')

    escolher(monkeypatch, f.provisorio(tmp_path / 'ccm.xlsx', ['Nº CCM', 'SSOP', 'STATUS'], f.CCM_SINTETICO), '28/10/2026')
    app.importar_fic('CCM')
    assert app.R.total['ccm'] == 1
    assert [r['data_referencia'] for r in app.t_hist.linhas] == ['2026-10-28']   # só a linha TOTAL de 28/10
    assert [r['tipo'] for r in app.t_imp.linhas] == ['CCM', 'FVI', 'LI']


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
