"""Operações do programa usadas pela interface e pela linha de comando."""
import datetime

from . import calculo, importacao
from .importacao import ErroImportacao

TIPOS_FIC = ('FVI', 'FVM', 'CCM', 'TAP', 'TTAS')
NOMES = {'FVI': 'Relatório de Cadastro Geral (FVI)', 'FVM': 'FVM', 'CCM': 'CCM', 'TAP': 'TAP', 'TTAS': 'TTAS'}


def importar_li(banco, arquivo):
    """Base SOP / SSOP (Lista de Cadastros Básicos da FIC ou LI antiga)."""
    linhas, sops, avisos = importacao.ler_sop_ssop(arquivo)
    banco.substituir_base(linhas, arquivo, sops, avisos)
    un, _, _ = banco.base()
    n_sop = len(sops) if sops else len({r['sop'] for r in linhas})
    msg = f'SOP / SSOP importada: {len(un)} unidade(s), {n_sop} SOP, {len(linhas)} SSOP.'
    if avisos:
        msg += f' {len(avisos)} aviso(s) na página Verificação.'
    return msg


def ler_fic(tipo, arquivo):
    """Lê e valida o arquivo sem gravar. Devolve (registros, data_de_emissão|None, info)."""
    if tipo == 'FVI':
        return importacao.ler_cadastro(arquivo)
    if tipo == 'FVM':
        return importacao.ler_fvm(arquivo)
    return importacao.ler_provisorio(tipo, arquivo)


def gravar_fic(banco, tipo, arquivo, registros, data_referencia, info=None):
    """Grava os registros (substituindo a importação anterior do tipo), recalcula e registra o histórico na data."""
    if not banco.tem_base():
        raise ErroImportacao('Importe primeiro a lista de SOP / SSOP (Importar atualização ▸ SOP / SSOP).')
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


def calcular(banco):
    dados = {t: banco.dados(t) for t in TIPOS_FIC}
    dados['AVISOS_BASE'] = banco.avisos_base()
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
                          'pct_tap': p(h['tap'], h['sop']), 'pct_ttas1': p(h['ttas1'], h['sop']),
                          'pct_ttas2': p(h['ttas2'], h['sop'])})
    return saida
