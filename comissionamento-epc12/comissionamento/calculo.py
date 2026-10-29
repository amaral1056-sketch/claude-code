"""Cálculo do avanço – reprodução das fórmulas da planilha CONTROLE_GERENCIAL_EPC12.xlsx.

Referência de cada regra (planilha original, scripts/02_estrutura.py):
- IMP_FVI AP:BD (por TAG): VALIDO, CII/CIII/CV_APLIC e _OK, FORA_DA_BASE, COD_DESCONHECIDO, SEM_FVI, SEM_SSOP, CIV_ATIVO;
- IMP_FVM/IMP_CCM/IMP_TAP/IMP_TTAS: VALIDO, *_OK, FORA_DA_BASE, COD_DESCONHECIDO;
- abas FVI/FVM/CCM/TAP (por SSOP e SOP), TTAS (por SOP), VISAO_GERENCIAL (por unidade, totais ponderados), VERIFICACAO.
Comparações seguem o Excel: TRIM e sem diferença de maiúsculas/minúsculas (ver texto.py)."""
import json
from collections import defaultdict

from .texto import chave

CAMPOS_FVI = (('cii', 'st_cii', 'FVI_CII'), ('ciii', 'st_ciii', 'FVI_CIII'), ('cv', 'st_cv', 'FVI_CV'))


def pct(num, den):
    return num / den if den else None


class Resultado:
    """Tudo o que as telas e a exportação mostram."""

    def __init__(self):
        self.unidades = []          # ordem da LI
        self.ssop = []              # dicts por SSOP (ordem da LI)
        self.sop = []               # dicts por SOP
        self.visao = []             # dicts por unidade
        self.total = {}             # linha TOTAL da visão gerencial
        self.verificacao = []       # (grupo, teste, esperado, apurado, resultado)
        self.pendencias = []        # (documento, linha do arquivo, identificação, SSOP/SOP, problema)
        self.tags = []              # uma linha por TAG lida do cadastro, com o status de cada campo
        self.malhas = []            # uma linha por malha lida da FVM
        self.contagens = {}         # números brutos (TAGs lidas etc.)

    def historico_por_unidade(self):
        cols = {'sop': 'sop', 'ssop': 'ssop', 'tags': 'tags', 'cii_aplic': 'cii_aplic', 'cii_aprov': 'cii_aprov',
                'ciii_aplic': 'ciii_aplic', 'ciii_aprov': 'ciii_aprov', 'cv_aplic': 'cv_aplic', 'cv_aprov': 'cv_aprov',
                'malhas': 'malhas', 'fvm_ciii': 'fvm_ciii', 'fvm_cv': 'fvm_cv', 'ccm': 'ccm', 'tap': 'tap',
                'ttas1': 'ttas1', 'ttas2': 'ttas2', 'tags_sem_fvi': 'tags_sem_fvi'}
        return {v['unidade']: {k: v[c] for k, c in cols.items()} for v in self.visao}


def _descricao(dados):
    """Descrição da TAG guardada no JSON da linha do cadastro."""
    if isinstance(dados, str):
        try:
            dados = json.loads(dados)
        except ValueError:
            return ''
    return (dados or {}).get('DESCRIÇÃO DO TAG', '')


def _conjuntos(criterios, status_fic):
    crit = defaultdict(set)
    for _, etapa, codigo, _ in criterios:
        crit[chave(etapa)].add(chave(codigo))
    cod = defaultdict(set)
    for doc, codigo, _ in status_fic:
        cod[chave(doc)].add(chave(codigo))
    return crit, cod


def calcular(base, criterios, status_fic, incluir_sem_fvi, dados):
    """base = (unidades, sops, ssops) de Banco.base(); dados = {'FVI': [...], 'FVM': [...], ...} de Banco.dados()."""
    unidades, sops, ssops = base
    crit, cod = _conjuntos(criterios, status_fic)
    canc = crit['CANCELADO']
    R = Resultado()
    R.unidades = list(unidades)

    zeros = ('tags', 'cii_aplic', 'cii_aprov', 'ciii_aplic', 'ciii_aprov', 'cv_aplic', 'cv_aprov', 'tags_sem_fvi',
             'malhas', 'fvm_ciii', 'fvm_cv', 'ccm_reg', 'ccm_ok', 'tap_reg', 'tap_ok')
    acc = {chave(s['ssop']): dict.fromkeys(zeros, 0) for s in ssops}
    base_ssop = {chave(s['ssop']): s for s in ssops}
    pend = R.pendencias
    cont = R.contagens

    def desconhecido(doc, st, ignora_na=True):
        k = chave(st)
        return k != '' and not (ignora_na and k == 'N/A') and k not in cod[doc]

    # ------------------------------------------------------------------ FVI (por TAG)
    n = dict.fromkeys(('lidas', 'fora', 'desconhecido', 'sem_fvi', 'sem_ssop', 'civ', 'canceladas'), 0)
    for t in dados.get('FVI', []):
        n['lidas'] += 1
        k = chave(t['ssop'])
        valido = not any(chave(t[c]) in canc for c in ('st_fvi', 'st_cii', 'st_ciii', 'st_cv'))
        sem_fvi = chave(t['n_fvi']) == ''
        n['canceladas'] += not valido
        n['sem_fvi'] += sem_fvi
        ref = base_ssop.get(k)
        rec = {'linha': t['linha'], 'tag': t['tag'], 'descricao': _descricao(t.get('dados')),
               'unidade': ref['unidade'] if ref else t['unidade'], 'sop': ref['sop'] if ref else t['sop'],
               'ssop': t['ssop'], 'classe': t['classe'], 'n_fvi': t['n_fvi'], 'sem_fvi': sem_fvi,
               'st_fvi': t['st_fvi'], 'st_civ': t['st_civ'],
               'situacao': 'CANCELADA' if not valido else ('SEM SSOP' if k == '' else ('FORA DA BASE' if ref is None else 'VÁLIDA'))}
        for campo, col, etapa in CAMPOS_FVI:
            st = chave(t[col])
            rec['st_' + campo] = t[col]
            # True = conta como aprovado · False = pendente · None = N/A ou TAG cancelada
            rec['ok_' + campo] = None if (not valido or st == 'N/A') else (st != '' and st in crit[etapa])
        R.tags.append(rec)
        unk = sum(desconhecido('FVI', t[c]) for c in ('st_fvi', 'st_cii', 'st_ciii', 'st_civ', 'st_cv'))
        if unk:
            n['desconhecido'] += unk
            pend.append(('FVI', t['linha'], t['tag'], t['ssop'], 'status com código desconhecido: ' + ', '.join(
                t[c] for c in ('st_fvi', 'st_cii', 'st_ciii', 'st_civ', 'st_cv') if desconhecido('FVI', t[c]))))
        if chave(t['st_civ']) not in ('', 'N/A'):
            n['civ'] += 1
            pend.append(('FVI', t['linha'], t['tag'], t['ssop'], f'Campo IV preenchido ({t["st_civ"]}) – campo não acompanhado'))
        if k == '':
            n['sem_ssop'] += 1
            pend.append(('FVI', t['linha'], t['tag'], '', 'TAG sem SSOP (não alocada – não conta avanço)'))
            continue
        if k not in acc:
            n['fora'] += 1
            pend.append(('FVI', t['linha'], t['tag'], t['ssop'], 'SSOP fora da base (LI)'))
            continue
        a = acc[k]
        a['tags_sem_fvi'] += sem_fvi   # coluna O da aba FVI: soma SEM_FVI de todas as linhas do SSOP
        if not valido:
            continue
        a['tags'] += 1
        for campo, col, etapa in CAMPOS_FVI:
            st = chave(t[col])
            if st == 'N/A' or (sem_fvi and not incluir_sem_fvi):
                continue
            a[campo + '_aplic'] += 1
            a[campo + '_aprov'] += st != '' and st in crit[etapa]
    cont['FVI'] = n

    # ------------------------------------------------------------------ FVM, CCM, TAP (por SSOP)
    def por_ssop(doc, regs, ident, cod_doc, ignora_na):
        m = dict.fromkeys(('lidas', 'fora', 'desconhecido', 'sem_ssop', 'canceladas'), 0)
        for r in regs:
            m['lidas'] += 1
            k = chave(r['ssop'])
            st = chave(r['status'])
            valido = st not in canc or st == ''
            m['canceladas'] += not valido
            if cod_doc and desconhecido(cod_doc, r['status'], ignora_na):
                m['desconhecido'] += 1
                pend.append((doc, r['linha'], r[ident], r['ssop'], f'status com código desconhecido: {r["status"]}'))
            if k == '':
                m['sem_ssop'] += 1
                pend.append((doc, r['linha'], r[ident], '', 'registro sem SSOP'))
                continue
            if k not in acc:
                m['fora'] += 1
                pend.append((doc, r['linha'], r[ident], r['ssop'], 'SSOP fora da base (LI)'))
                continue
            if valido:
                yield acc[k], st
        cont[doc] = m

    for r in dados.get('FVM', []):
        k, st = chave(r['ssop']), chave(r['status'])
        ref = base_ssop.get(k)
        valido = st not in canc or st == ''
        R.malhas.append({
            'linha': r['linha'], 'malha': r['malha'], 'unidade': ref['unidade'] if ref else '',
            'sop': ref['sop'] if ref else r['sop'], 'ssop': r['ssop'], 'status': r['status'],
            'ok_ciii': (st != '' and st in crit['FVM_CIII']) if valido else None,
            'ok_cv': (st != '' and st in crit['FVM_CV']) if valido else None,
            'situacao': 'CANCELADA' if not valido else ('SEM SSOP' if k == '' else ('FORA DA BASE' if ref is None else 'VÁLIDA'))})
    for a, st in por_ssop('FVM', dados.get('FVM', []), 'malha', 'FVM', True):
        a['malhas'] += 1
        a['fvm_ciii'] += st != '' and st in crit['FVM_CIII']
        a['fvm_cv'] += st != '' and st in crit['FVM_CV']
    for a, st in por_ssop('CCM', dados.get('CCM', []), 'numero', 'CCM', True):
        a['ccm_reg'] += 1
        a['ccm_ok'] += st != '' and st in crit['CCM']
    for a, st in por_ssop('TAP', dados.get('TAP', []), 'numero', 'TAP', False):
        a['tap_reg'] += 1
        a['tap_ok'] += st != '' and st in crit['TAP']

    # ------------------------------------------------------------------ TTAS (por SOP)
    acc_sop = {chave(s['sop']): {'ttas_reg': 0, 'ttas1': 0, 'ttas2': 0} for s in sops}
    m = dict.fromkeys(('lidas', 'fora', 'desconhecido', 'sem_ssop', 'canceladas'), 0)
    for r in dados.get('TTAS', []):
        m['lidas'] += 1
        k = chave(r['sop'])
        st = chave(r['status'])
        valido = st not in canc or st == ''
        m['canceladas'] += not valido
        if desconhecido('TTAS', r['status']):
            m['desconhecido'] += 1
            pend.append(('TTAS', r['linha'], r['numero'], r['sop'], f'status com código desconhecido: {r["status"]}'))
        if k == '':
            m['sem_ssop'] += 1
            pend.append(('TTAS', r['linha'], r['numero'], '', 'registro sem SOP'))
            continue
        if k not in acc_sop:
            m['fora'] += 1
            pend.append(('TTAS', r['linha'], r['numero'], r['sop'], 'SOP fora da base (LI)'))
            continue
        if valido:
            b = acc_sop[k]
            b['ttas_reg'] += 1
            b['ttas1'] += st != '' and st in crit['TTAS1']
            b['ttas2'] += st != '' and st in crit['TTAS2']
    cont['TTAS'] = m

    # ------------------------------------------------------------------ linhas por SSOP
    for s in ssops:
        a = acc[chave(s['ssop'])]
        R.ssop.append({
            'unidade': s['unidade'], 'sop': s['sop'], 'ssop': s['ssop'], 'descricao': s['descricao'], **a,
            'pct_cii': pct(a['cii_aprov'], a['cii_aplic']), 'pct_ciii': pct(a['ciii_aprov'], a['ciii_aplic']),
            'pct_cv': pct(a['cv_aprov'], a['cv_aplic']),
            'pct_fvm_ciii': pct(a['fvm_ciii'], a['malhas']), 'pct_fvm_cv': pct(a['fvm_cv'], a['malhas']),
            'ccm_aprovado': 1 if a['ccm_ok'] else 0,
            'ccm_situacao': 'APROVADO' if a['ccm_ok'] else ('SEM REGISTRO' if a['ccm_reg'] == 0 else 'EM ANDAMENTO'),
            'tap_concluido': 1 if a['tap_ok'] else 0,
            'tap_situacao': 'CONCLUÍDO' if a['tap_ok'] else ('SEM REGISTRO' if a['tap_reg'] == 0 else 'EM ANDAMENTO'),
        })

    # ------------------------------------------------------------------ linhas por SOP (somas das SSOP)
    somas = ('tags', 'cii_aplic', 'cii_aprov', 'ciii_aplic', 'ciii_aprov', 'cv_aplic', 'cv_aprov', 'tags_sem_fvi',
             'malhas', 'fvm_ciii', 'fvm_cv')
    por_sop = defaultdict(lambda: dict.fromkeys(somas + ('qtd_ssop',), 0))
    for linha in R.ssop:
        p = por_sop[chave(linha['sop'])]
        p['qtd_ssop'] += 1
        for c in somas:
            p[c] += linha[c]
    for s in sops:
        k = chave(s['sop'])
        p, b = por_sop[k], acc_sop[k]
        R.sop.append({
            'unidade': s['unidade'], 'sop': s['sop'], 'descricao': s['descricao'], **p, **b,
            'pct_cii': pct(p['cii_aprov'], p['cii_aplic']), 'pct_ciii': pct(p['ciii_aprov'], p['ciii_aplic']),
            'pct_cv': pct(p['cv_aprov'], p['cv_aplic']),
            'pct_fvm_ciii': pct(p['fvm_ciii'], p['malhas']), 'pct_fvm_cv': pct(p['fvm_cv'], p['malhas']),
            'ttas1_aprovado': 1 if b['ttas1'] else 0, 'ttas2_aprovado': 1 if b['ttas2'] else 0,
            'ttas_situacao': 'TTAS2 APROVADO' if b['ttas2'] else ('TTAS1 APROVADO' if b['ttas1'] else
                                                                  ('SEM REGISTRO' if b['ttas_reg'] == 0 else 'EM ANDAMENTO')),
        })

    # ------------------------------------------------------------------ visão gerencial
    def linha_visao(nome, ssop_rows, sop_rows):
        v = {'unidade': nome, 'sop': len(sop_rows), 'ssop': len(ssop_rows)}
        for c in ('tags', 'cii_aplic', 'cii_aprov', 'ciii_aplic', 'ciii_aprov', 'cv_aplic', 'cv_aprov', 'malhas',
                  'fvm_ciii', 'fvm_cv', 'tags_sem_fvi'):
            v[c] = sum(r[c] for r in ssop_rows)
        v['ccm'] = sum(r['ccm_aprovado'] for r in ssop_rows)
        v['tap'] = sum(r['tap_concluido'] for r in ssop_rows)
        v['ttas1'] = sum(r['ttas1_aprovado'] for r in sop_rows)
        v['ttas2'] = sum(r['ttas2_aprovado'] for r in sop_rows)
        v.update({
            'pct_cii': pct(v['cii_aprov'], v['cii_aplic']), 'pct_ciii': pct(v['ciii_aprov'], v['ciii_aplic']),
            'pct_cv': pct(v['cv_aprov'], v['cv_aplic']),
            'pct_fvm_ciii': pct(v['fvm_ciii'], v['malhas']), 'pct_fvm_cv': pct(v['fvm_cv'], v['malhas']),
            'pct_ccm': pct(v['ccm'], v['ssop']), 'pct_tap': pct(v['tap'], v['ssop']),
            'pct_ttas1': pct(v['ttas1'], v['sop']), 'pct_ttas2': pct(v['ttas2'], v['sop']),
        })
        return v

    for u in unidades:
        R.visao.append(linha_visao(u, [r for r in R.ssop if r['unidade'] == u], [r for r in R.sop if r['unidade'] == u]))
    R.total = linha_visao('TOTAL', R.ssop, R.sop)

    # ------------------------------------------------------------------ verificação
    def teste(grupo, nome, apurado, esperado=0):
        R.verificacao.append((grupo, nome, esperado, apurado,
                              'INFO' if esperado is None else ('OK' if apurado == esperado else 'VERIFICAR')))

    f, c = cont['FVI'], cont
    teste('FVI', 'Cadastro FVI – TAGs com SSOP fora da base', f['fora'])
    teste('FVI', 'Cadastro FVI – status com código desconhecido', f['desconhecido'])
    teste('FVI', 'Cadastro FVI – TAGs com Campo IV diferente de N/A (campo não acompanhado)', f['civ'])
    for doc in ('FVM', 'CCM', 'TAP'):
        teste(doc, f'{doc} – linhas com SSOP fora da base', c[doc]['fora'])
        teste(doc, f'{doc} – status com código desconhecido', c[doc]['desconhecido'])
    teste('TTAS', 'TTAS – linhas com SOP fora da base', c['TTAS']['fora'])
    teste('TTAS', 'TTAS – status com código desconhecido', c['TTAS']['desconhecido'])
    teste('FVI', 'Cadastro FVI – TAGs lidas', f['lidas'], None)
    teste('FVI', 'Cadastro FVI – TAGs sem SSOP (não alocadas – não contam avanço)', f['sem_ssop'], None)
    teste('FVI', 'Cadastro FVI – TAGs sem FVI gerada', f['sem_fvi'], None)
    teste('FVI', 'Cadastro FVI – TAGs canceladas (fora das contas)', f['canceladas'], None)
    for doc in ('FVM', 'CCM', 'TAP', 'TTAS'):
        teste(doc, f'{doc} – registros lidos', c[doc]['lidas'], None)
        teste(doc, f'{doc} – registros sem {"SOP" if doc == "TTAS" else "SSOP"}', c[doc]['sem_ssop'], None)
        teste(doc, f'{doc} – registros cancelados (fora das contas)', c[doc]['canceladas'], None)
    for doc in ('FVI', 'FVM', 'CCM', 'TAP', 'TTAS'):
        teste('CRITÉRIOS', f'Lista oficial de status {doc} carregada (códigos)', len(cod[doc]), None)
    return R
