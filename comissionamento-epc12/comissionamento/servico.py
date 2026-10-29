"""Operações do programa usadas pela interface e pela linha de comando."""
import datetime

from . import calculo, importacao
from .importacao import ErroImportacao

TIPOS_FIC = ('FVI', 'FVM', 'CCM', 'TAP', 'TTAS')
NOMES = {'FVI': 'Relatório de Cadastro Geral (FVI)', 'FVM': 'FVM', 'CCM': 'CCM', 'TAP': 'TAP', 'TTAS': 'TTAS'}


def importar_li(banco, arquivo):
    linhas = importacao.ler_li(arquivo)
    banco.substituir_base(linhas, arquivo)
    un = len({r['unidade'] for r in linhas})
    sop = len({r['sop'] for r in linhas})
    return f'LI importada: {un} unidades, {sop} SOP, {len(linhas)} SSOP.'


def ler_fic(tipo, arquivo):
    """Lê e valida o arquivo sem gravar. Devolve (registros, data_de_emissão|None, info)."""
    if tipo == 'FVI':
        return importacao.ler_cadastro(arquivo)
    return importacao.ler_provisorio(tipo, arquivo)


def gravar_fic(banco, tipo, arquivo, registros, data_referencia, info=None):
    """Grava os registros (substituindo a importação anterior do tipo), recalcula e registra o histórico na data."""
    if not banco.tem_base():
        raise ErroImportacao('Importe a LI (base SOP/SSOP) antes das exportações da FIC.')
    data_iso = _iso(data_referencia)
    obs = ''
    if info and info.get('linhas_ignoradas'):
        obs = f'{len(info["linhas_ignoradas"])} linha(s) ignorada(s) (cabeçalho repetido ou rodapé)'
    banco.substituir_dados(tipo, registros, arquivo, data_iso, obs)
    R = calcular(banco)
    banco.gravar_historico(data_iso, R.historico_por_unidade())
    return R


def importar_fic(banco, tipo, arquivo, data_referencia=None):
    """Lê + grava. Sem data informada, usa a data de emissão do rodapé; sem rodapé, a data de hoje."""
    regs, data_rodape, info = ler_fic(tipo, arquivo)
    data = data_referencia or data_rodape or datetime.date.today()
    gravar_fic(banco, tipo, arquivo, regs, data, info)
    return regs, _iso(data), info


def importar_status(banco, arquivo, tipo=None):
    tipo = tipo or importacao.tipo_pelo_nome(arquivo)
    if tipo not in TIPOS_FIC:
        raise ErroImportacao(f'Não foi possível identificar o documento (FVI/FVM/CCM/TAP/TTAS) pelo nome do arquivo {arquivo}.')
    lista = importacao.ler_status(arquivo)
    banco.substituir_status_fic(tipo, lista)
    return tipo, len(lista)


def importar_base(banco, tipo, arquivo):
    """LI de Itens / LI de Malhas: substitui a base anterior do tipo (sem data de referência nem histórico)."""
    regs, _, info = importacao.ler_provisorio(tipo, arquivo)
    obs = f'{len(info["linhas_ignoradas"])} linha(s) ignorada(s) (cabeçalho repetido ou rodapé)' if info['linhas_ignoradas'] else ''
    banco.substituir_dados(tipo, regs, arquivo, None, obs)
    return f'{importacao.NOMES_BASE[tipo]} importada: {len(regs)} registro(s).'


def calcular(banco):
    dados = {t: banco.dados(t) for t in TIPOS_FIC + ('LI_ITENS', 'LI_MALHAS')}
    return calculo.calcular(banco.base(), banco.criterios(), banco.status_fic(), banco.incluir_sem_fvi, dados)


def registrar_historico(banco, data_referencia):
    data_iso = _iso(data_referencia)
    banco.gravar_historico(data_iso, calcular(banco).historico_por_unidade())
    return data_iso


def _iso(d):
    if isinstance(d, datetime.datetime):
        d = d.date()
    if isinstance(d, datetime.date):
        return d.isoformat()
    s = str(d).strip()
    for fmt in ('%d/%m/%Y', '%Y-%m-%d'):
        try:
            return datetime.datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    raise ErroImportacao(f'Data inválida: {d!r} (use dd/mm/aaaa).')


def data_br(iso):
    return datetime.date.fromisoformat(iso).strftime('%d/%m/%Y') if iso else ''


def historico_com_percentuais(banco):
    """Linhas do histórico com os % calculados, mais uma linha TOTAL por data (ponderada)."""
    por_data = {}
    for h in banco.historico():
        por_data.setdefault(h['data_referencia'], []).append(h)
    saida = []
    for data, linhas in por_data.items():
        total = {'unidade': 'TOTAL', 'data_referencia': data}
        for c in linhas[0]:
            if c not in ('data_referencia', 'unidade', 'registrado_em'):
                total[c] = sum(x[c] or 0 for x in linhas)
        for h in linhas + [total]:
            p = calculo.pct
            saida.append({**h, 'data': data_br(data),
                          'pct_cii': p(h['cii_aprov'], h['cii_aplic']), 'pct_ciii': p(h['ciii_aprov'], h['ciii_aplic']),
                          'pct_cv': p(h['cv_aprov'], h['cv_aplic']), 'pct_fvm_ciii': p(h['fvm_ciii'], h['malhas']),
                          'pct_fvm_cv': p(h['fvm_cv'], h['malhas']), 'pct_ccm': p(h['ccm'], h['ssop']),
                          'pct_tap': p(h['tap'], h['ssop']), 'pct_ttas1': p(h['ttas1'], h['sop']),
                          'pct_ttas2': p(h['ttas2'], h['sop'])})
    return saida
