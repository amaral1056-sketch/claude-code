"""Leitura das planilhas Excel que alimentam o banco.

- SOP / SSOP: "Lista de Cadastros Básicos" da FIC (Código, Descrição, Tipo SOP/SSOP). Unidade e SOP de cada SSOP vêm
  do código; SSOP cujo SOP não está na lista e código fora do padrão vão para a Verificação (não param a importação).
  A LI antiga (aba LISTA) continua aceita, com as checagens da etapa 01 da planilha original.
- Relatório de Cadastro Geral da FIC (FVI): o arquivo exportado, sem edição (título, cabeçalho, dados, rodapé).
  Mesma regra de leitura da IMP_FVI: linha de dado = TAG e CONTRATO preenchidos e TAG diferente de "TAG".
- FVM BASE: "Relatório de Hierarquia de Malhas" – linhas MALHA seguidas das TAG/LINHA da malha, com SSOP/SOP
  (uma linha por TAG). Define as malhas acompanhadas.
- FVM (atualização): "Lista de FVM" da FIC (Código FVM, Malha, Status – código) ou "Relatório de Status de Malhas"
  (Malha, FVM, Status – descrição).
- CCM (por SSOP), TAP e TTAS (por SOP): listas da FIC; colunas localizadas pelo NOME do cabeçalho
  (COLUNAS_PROVISORIAS). No TAP sem coluna SOP, o SOP vem do número ("TAP-" + SOP).
- STATUS_*.xlsx: listas oficiais de status (código | descrição).
Nada do texto das fontes é corrigido: só é lido e, quando há inconsistência, reportado."""
import datetime
import re
import unicodedata
from pathlib import Path

import openpyxl

from .texto import chave, texto

DATA_RODAPE = re.compile(r'\s*(\d\d/\d\d/\d{4})')


class ErroImportacao(Exception):
    """Arquivo fora do layout esperado: a importação é cancelada e nada é gravado."""


def _abrir(arquivo):
    try:
        return openpyxl.load_workbook(arquivo, read_only=True, data_only=True)
    except Exception as e:  # arquivo aberto/corrompido/formato errado
        raise ErroImportacao(f'Não foi possível abrir {Path(arquivo).name}: {e}') from e


def _sem_acento(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')


def _nome_coluna(v):
    """Nome de cabeçalho comparável: TRIM, maiúsculas, sem acento, 'N°' (grau) -> 'Nº'."""
    s = _sem_acento(chave(v)).replace('N°', 'Nº')
    return re.sub(r'\s+', ' ', s)


def data_do_rodape(linhas):
    """Data de emissão no rodapé do relatório (1ª coluna, de baixo para cima)."""
    for row in reversed(linhas):
        if row:
            m = DATA_RODAPE.match(texto(row[0]))
            if m:
                return datetime.datetime.strptime(m.group(1), '%d/%m/%Y').date()
    return None


# ====================================================================== LI
COLUNAS_LI = ('ITEM', 'UNIDADE', 'SOP', 'DESCRIÇÃO SOP', 'SSOP', 'DESCRIÇÃO SSOP', 'DOCUMENTOS DE REFERÊNCIA')


def ler_sop_ssop(arquivo):
    """Base SOP / SSOP. Devolve (linhas por SSOP, SOPs em ordem | None, avisos [(código, problema)])."""
    wb = _abrir(arquivo)
    eh_li = 'LISTA' in wb.sheetnames
    wb.close()
    if eh_li:
        return ler_li(arquivo), None, []
    return ler_cadastros_basicos(arquivo)


PADRAO_CODIGO = re.compile(r'^S?SOP-(.+?)-([A-Z]{2})-\d{4}')       # unidade e grupo (disciplina do SOP)
PADRAO_SSOP = re.compile(r'^S?(SOP-.+?-[A-Z]{2}-\d{4})-[^-]+$')     # SOP a que a SSOP pertence


def ler_cadastros_basicos(arquivo):
    """Lista de Cadastros Básicos (Código | Descrição | Tipo | Status | ...)."""
    spec = [('codigo', True, ('CODIGO',)), ('descricao', True, ('DESCRICAO',)), ('tipo', True, ('TIPO',))]
    wb = _abrir(arquivo)
    linhas = [tuple(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    wb.close()
    i, pos = _localizar_cabecalho(linhas, spec)
    if i is None:
        raise ErroImportacao('Planilha de SOP / SSOP fora do formato: esperado cabeçalho com Código, Descrição e Tipo '
                             '(Lista de Cadastros Básicos da FIC) ou a aba LISTA da LI.')
    regs = []
    for n, r in enumerate(linhas[i + 1:], i + 2):
        v = {c: (texto(r[p]).strip() if p < len(r) else '') for c, p in pos.items()}
        if v['codigo']:
            regs.append((n, v))
    problemas, vistos = [], set()
    for n, v in regs:
        if v['codigo'] in vistos:
            problemas.append(f'linha {n}: código duplicado {v["codigo"]}')
        vistos.add(v['codigo'])
        if chave(v['tipo']) not in ('SOP', 'SSOP'):
            problemas.append(f'linha {n}: tipo "{v["tipo"]}" (esperado SOP ou SSOP) em {v["codigo"]}')
        elif not PADRAO_CODIGO.match(v['codigo']) or (chave(v['tipo']) == 'SSOP' and not PADRAO_SSOP.match(v['codigo'])):
            problemas.append(f'linha {n}: código {v["codigo"]} fora do padrão (unidade/SOP não identificáveis)')
    if not any(chave(v['tipo']) == 'SSOP' for _, v in regs):
        problemas.append('nenhuma linha do tipo SSOP')
    if problemas:
        raise ErroImportacao('A lista de SOP / SSOP não passou nas checagens (nada foi gravado):\n- '
                             + '\n- '.join(problemas[:30]) + ('\n…' if len(problemas) > 30 else ''))
    sops = {}
    for _, v in regs:
        if chave(v['tipo']) == 'SOP':
            m = PADRAO_CODIGO.match(v['codigo'])
            sops[v['codigo']] = {'sop': v['codigo'], 'unidade': m.group(1), 'descricao': v['descricao'], 'grupo': m.group(2)}
    avisos, out = [], []
    for _, v in regs:
        if chave(v['tipo']) != 'SSOP':
            continue
        sop = PADRAO_SSOP.match(v['codigo']).group(1)
        m = PADRAO_CODIGO.match(v['codigo'])
        if not v['codigo'].startswith('SSOP-'):
            avisos.append((v['codigo'], f'código de SSOP não começa com "SSOP-" (SOP considerado: {sop})'))
        if sop not in sops:
            avisos.append((v['codigo'], f'SOP {sop} não está na lista (incluído sem descrição – A CONFIRMAR)'))
            sops[sop] = {'sop': sop, 'unidade': m.group(1), 'descricao': '', 'grupo': m.group(2)}
        out.append({'item': len(out) + 1, 'unidade': m.group(1), 'sop': sop, 'descricao_sop': sops[sop]['descricao'],
                    'ssop': v['codigo'], 'descricao_ssop': v['descricao'], 'documentos': '', 'grupo': m.group(2)})
    return out, list(sops.values()), avisos


def ler_li(arquivo):
    wb = _abrir(arquivo)
    if 'LISTA' not in wb.sheetnames:
        raise ErroImportacao(f'A LI precisa ter a aba LISTA. Abas encontradas: {", ".join(wb.sheetnames)}')
    hdr, idx, linhas = None, None, []
    for r in wb['LISTA'].iter_rows(values_only=True):
        if hdr is None:
            cab = [str(x).strip() if x is not None else '' for x in r]
            if 'SSOP' in cab:
                hdr = cab
                idx = {h: i for i, h in enumerate(hdr) if h}
                faltam = [c for c in COLUNAS_LI if c not in idx]
                if faltam:
                    raise ErroImportacao(f'Colunas não encontradas na aba LISTA: {", ".join(faltam)}')
            continue
        if idx['ITEM'] >= len(r) or r[idx['ITEM']] is None:
            continue
        linhas.append({h: (str(r[i]).strip() if i < len(r) and r[i] is not None else '') for h, i in idx.items()})
    wb.close()
    if hdr is None:
        raise ErroImportacao('Cabeçalho com a coluna SSOP não encontrado na aba LISTA.')
    if not linhas:
        raise ErroImportacao('A aba LISTA não tem linhas de SSOP.')

    problemas = []
    vistos = set()
    for r in linhas:
        if r['SSOP'] in vistos:
            problemas.append(f'SSOP duplicado: {r["SSOP"]} (ITEM {r["ITEM"]})')
        vistos.add(r['SSOP'])
        if not r['SSOP'].startswith('S' + r['SOP'] + '-'):
            problemas.append(f'ITEM {r["ITEM"]}: SSOP {r["SSOP"]} não começa com "S" + SOP {r["SOP"]}')
        m = re.match(r'SOP-(.+?)-[A-Z]{2}-\d{4}$', r['SOP'])
        if not m or m.group(1) != r['UNIDADE']:
            problemas.append(f'ITEM {r["ITEM"]}: unidade {r["UNIDADE"]} não confere com o código do SOP {r["SOP"]}')
    try:
        itens = [int(float(r['ITEM'])) for r in linhas]
    except ValueError:
        itens = None
    if itens != list(range(1, len(linhas) + 1)):
        problemas.append('A coluna ITEM não é a sequência 1, 2, 3… sem saltos.')
    if problemas:
        raise ErroImportacao('A LI não passou nas checagens de integridade (nada foi gravado):\n- '
                             + '\n- '.join(problemas[:30]) + ('\n…' if len(problemas) > 30 else ''))
    out = []
    for r, item in zip(linhas, itens):
        grupo = re.match(r'SOP-.+-([A-Z]{2})-\d{4}$', r['SOP']).group(1)
        out.append({'item': item, 'unidade': r['UNIDADE'], 'sop': r['SOP'], 'descricao_sop': r['DESCRIÇÃO SOP'],
                    'ssop': r['SSOP'], 'descricao_ssop': r['DESCRIÇÃO SSOP'],
                    'documentos': re.sub(r'\s{2,}', ' | ', r['DOCUMENTOS DE REFERÊNCIA']), 'grupo': grupo})
    return out


# ====================================================================== Cadastro Geral (FVI)
N_COLS_CADASTRO = 41  # A:AO
# coluna (1 = A) -> nome esperado no cabeçalho; se o layout mudar, a importação para (premissa das regras)
CADASTRO_ESPERADO = {1: 'TAG', 4: 'SOP', 5: 'SSOP', 6: 'CLASSE', 29: 'CONTRATO', 31: 'Nº FVI', 36: 'STATUS FVI',
                     37: 'STATUS CAMPO II', 38: 'STATUS CAMPO III', 39: 'STATUS CAMPO IV', 40: 'STATUS CAMPO V'}
CAMPOS_CADASTRO = {'tag': 1, 'unidade': 3, 'sop': 4, 'ssop': 5, 'classe': 6, 'disciplina': 7, 'malha': 16,
                   'contrato': 29, 'n_fvi': 31, 'st_fvi': 36, 'st_cii': 37, 'st_ciii': 38, 'st_civ': 39, 'st_cv': 40}


def ler_cadastro(arquivo):
    """Devolve (registros, data_de_emissão|None, informações)."""
    wb = _abrir(arquivo)
    linhas = [tuple(r[:N_COLS_CADASTRO]) + (None,) * (N_COLS_CADASTRO - len(r[:N_COLS_CADASTRO]))
              for r in wb.worksheets[0].iter_rows(values_only=True)]
    wb.close()
    cab_i = next((i for i, r in enumerate(linhas[:30]) if chave(r[0]) == 'TAG'), None)
    if cab_i is None:
        raise ErroImportacao('Cabeçalho do Relatório de Cadastro Geral não encontrado (coluna A = "TAG" nas 30 primeiras linhas).')
    cab = [texto(v).strip() for v in linhas[cab_i]]
    for col, nome in CADASTRO_ESPERADO.items():
        if cab[col - 1] != nome:
            raise ErroImportacao(f'O layout do Relatório de Cadastro Geral mudou: coluna {openpyxl.utils.get_column_letter(col)} '
                                 f'= "{cab[col - 1]}", esperado "{nome}". Nada foi gravado.')
    regs = []
    for n, r in enumerate(linhas, 1):
        # regra da IMP_FVI: A<>"" E AC (CONTRATO)<>"" E TRIM(A)<>"TAG"
        if texto(r[0]) == '' or texto(r[28]) == '' or chave(r[0]) == 'TAG':
            continue
        reg = {k: texto(r[c - 1]) for k, c in CAMPOS_CADASTRO.items()}
        reg['linha'] = n
        reg['dados'] = {cab[j] or f'COL{j + 1}': texto(v) for j, v in enumerate(r) if texto(v) != ''}
        regs.append(reg)
    if not regs:
        raise ErroImportacao('Nenhuma TAG encontrada (linhas com TAG e CONTRATO preenchidos).')
    return regs, data_do_rodape(linhas), {'linhas_arquivo': len(linhas), 'tags': len(regs)}


# ====================================================================== FVM: relatórios de malhas da FIC
SPEC_HIERARQUIA = [('tipo', True, ('TIPO',)), ('malha', True, ('MALHA',)), ('tag', True, ('TAG/LINHAS', 'TAG/LINHA', 'TAG')),
                   ('status', True, ('STATUS',)), ('ssop', True, ('SSOP',)), ('sop', False, ('SOP',))]


def _linhas_planilha(arquivo):
    wb = _abrir(arquivo)
    linhas = [tuple(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    aba = wb.worksheets[0].title
    wb.close()
    return linhas, aba


def ler_fvm_base(arquivo):
    """FVM BASE = Relatório de Hierarquia de Malhas. Devolve (registros, data_de_emissão|None, informações)."""
    linhas, aba = _linhas_planilha(arquivo)
    i, pos = _localizar_cabecalho(linhas, SPEC_HIERARQUIA)
    if i is None:
        raise ErroImportacao('A FVM BASE deve ser o "Relatório de Hierarquia de Malhas" da FIC (colunas TIPO, MALHA, '
                             'TAG/LINHAS, STATUS, SSOP). Nada foi gravado.')
    return _hierarquia(linhas, i, pos, aba)


def ler_fvm(arquivo):
    """Atualização de FVM: Lista de FVM (Código FVM, Malha, Status) ou Relatório de Status de Malhas.
    Devolve (registros malha/n_fvm/status, data_de_emissão|None, informações)."""
    linhas, aba = _linhas_planilha(arquivo)
    if _localizar_cabecalho(linhas, SPEC_HIERARQUIA)[0] is not None:
        raise ErroImportacao('Este é o Relatório de Hierarquia de Malhas: importe-o em Importar atualização ▸ FVM BASE. '
                             'Em FVM vai a Lista de FVM da FIC. Nada foi gravado.')
    spec_l = [('n_fvm', True, ('CODIGO FVM', 'CODIGO DA FVM')), ('malha', True, ('MALHA',)), ('status', True, ('STATUS',)),
              ('template', False, ('CODIGO DO TEMPLATE', 'TEMPLATE'))]
    i, pos = _localizar_cabecalho(linhas, spec_l)
    if i is not None:
        return _lista_fvm(linhas, i, pos, aba)
    spec_s = [('malha', True, ('MALHA',)), ('n_fvm', True, ('FVM', 'Nº FVM', 'NUMERO FVM')), ('status', True, ('STATUS',))]
    i, pos = _localizar_cabecalho(linhas, spec_s)
    if i is not None:
        return _status_malhas(linhas, i, pos, aba)
    raise ErroImportacao('Arquivo de FVM não reconhecido. Use a "Lista de FVM" da FIC (colunas Código FVM, Malha, Status).')


def _lista_fvm(linhas, i, pos, aba):
    """Lista de FVM: uma linha por FVM, status como código (ex.: APROVADO)."""
    regs, ignoradas = [], []
    for n, r in enumerate(linhas[i + 1:], i + 2):
        if _nome_coluna(r[0] if r else '').startswith('EMISSAO'):
            ignoradas.extend(range(n, len(linhas) + 1))
            break
        malha = _valor(r, pos, 'malha')
        if malha:
            regs.append({'linha': n, 'malha': malha, 'n_fvm': _valor(r, pos, 'n_fvm'), 'status': _valor(r, pos, 'status'),
                         'template': _valor(r, pos, 'template')})
        elif any(texto(x) for x in r):
            ignoradas.append(n)
    if not regs:
        raise ErroImportacao(f'Nenhuma FVM encontrada na Lista de FVM (aba {aba}).')
    return regs, data_do_rodape(linhas), {'aba': aba, 'registros': len(regs), 'linhas_ignoradas': ignoradas}


def _valor(r, pos, c):
    p = pos.get(c)
    v = texto(r[p]).strip() if p is not None and p < len(r) else ''
    return '' if v == '-' else v          # "-" = sem valor no relatório


def _hierarquia(linhas, i, pos, aba):
    """Uma linha por TAG/LINHA da malha (malha sem TAG: uma linha sem TAG); status_item = status da malha (item)."""
    regs, status_malha, ignoradas = [], {}, []
    for n, r in enumerate(linhas[i + 1:], i + 2):
        tipo, malha = chave(_valor(r, pos, 'tipo')), _valor(r, pos, 'malha')
        if tipo == 'MALHA' and malha:
            status_malha[malha] = (n, _valor(r, pos, 'status'))
        elif tipo in ('TAG', 'LINHA') and malha:
            regs.append({'linha': n, 'malha': malha, 'tag': _valor(r, pos, 'tag'), 'sop': _valor(r, pos, 'sop'),
                         'ssop': _valor(r, pos, 'ssop'), 'status_item': _valor(r, pos, 'status')})
        elif any(texto(x) for x in r):
            ignoradas.append(n)           # rodapé (emissão, classificação, folha)
    com_tag = {r['malha'] for r in regs}
    for malha, (n, st) in status_malha.items():
        if malha not in com_tag:
            regs.append({'linha': n, 'malha': malha, 'tag': '', 'sop': '', 'ssop': '', 'status_item': st})
    if not regs:
        raise ErroImportacao(f'Nenhuma malha encontrada no Relatório de Hierarquia de Malhas (aba {aba}).')
    regs.sort(key=lambda r: r['linha'])
    return regs, data_do_rodape(linhas), {'aba': aba, 'registros': len(regs),
                                          'malhas': len({r['malha'] for r in regs}), 'linhas_ignoradas': ignoradas}


def _status_malhas(linhas, i, pos, aba):
    """Uma linha por malha: Nº da FVM e Status da FVM como vêm. Guarda o filtro de status do cabeçalho, se houver."""
    filtro = next((texto(c).strip() for r in linhas[:i] for c in r if texto(c).strip().upper().startswith('STATUS FVM:')), '')
    regs, ignoradas = [], []
    for n, r in enumerate(linhas[i + 1:], i + 2):
        if _nome_coluna(r[0] if r else '').startswith('EMISSAO'):
            ignoradas.extend(range(n, len(linhas) + 1))   # rodapé: emissão, emissor, classificação, folha, data
            break
        malha = _valor(r, pos, 'malha')
        if malha and chave(malha) != 'MALHA':
            regs.append({'linha': n, 'malha': malha, 'n_fvm': _valor(r, pos, 'n_fvm'), 'status': _valor(r, pos, 'status'),
                         'template': ''})
        elif any(texto(x) for x in r):
            ignoradas.append(n)
    if not regs:
        raise ErroImportacao(f'Nenhuma malha encontrada no Relatório de Status de Malhas (aba {aba}).')
    return regs, data_do_rodape(linhas), {'aba': aba, 'registros': len(regs),
                                          'filtro_status': filtro.rstrip(', '), 'linhas_ignoradas': ignoradas}


# ====================================================================== CCM, TAP, TTAS
# tipo -> [(campo no banco, obrigatório, nomes aceitos no cabeçalho)]
COLUNAS_PROVISORIAS = {
    'CCM': [('numero', True, ('Nº CCM', 'CCM', 'NUMERO CCM', 'NUMERO DO CCM', 'NUMERO', 'Nº')),
            ('ssop', True, ('SSOP',)),
            ('status', True, ('STATUS', 'STATUS CCM', 'STATUS DO CCM', 'SITUACAO'))],
    'TAP': [('numero', True, ('Nº TAP', 'TAP', 'NUMERO TAP', 'NUMERO DO TAP', 'NUMERO', 'Nº')),
            ('sop', False, ('SOP', 'SOP/SSOP')),
            ('status', True, ('STATUS', 'STATUS TAP', 'STATUS DO TAP', 'SITUACAO'))],
    'TTAS': [('numero', True, ('Nº TTAS', 'TTAS', 'NUMERO TTAS', 'NUMERO DO TTAS', 'NUMERO', 'Nº')),
             ('sop', True, ('SOP', 'SOP/SSOP')),
             ('status', True, ('STATUS', 'STATUS TTAS', 'STATUS DO TTAS', 'SITUACAO'))],
}


def _localizar_cabecalho(linhas, spec):
    for i, r in enumerate(linhas[:30]):
        nomes = [_nome_coluna(v) for v in r]
        pos = {}
        for campo, _, aceitos in spec:
            for nome in aceitos:  # ordem de preferência
                if nome in nomes and nomes.index(nome) not in pos.values():
                    pos[campo] = nomes.index(nome)
                    break
        if all(campo in pos for campo, obrig, _ in spec if obrig):
            return i, pos
    return None, None


def ler_provisorio(tipo, arquivo):
    """Devolve (registros, data_de_emissão|None, informações) para CCM/TAP/TTAS."""
    spec = COLUNAS_PROVISORIAS[tipo]
    wb = _abrir(arquivo)
    achado = None
    for ws in wb.worksheets:
        linhas = [tuple(r) for r in ws.iter_rows(values_only=True)]
        i, pos = _localizar_cabecalho(linhas, spec)
        if i is not None:
            achado = (ws.title, linhas, i, pos)
            break
    wb.close()
    if achado is None:
        esperado = '; '.join(f'{campo.upper()}: ' + ' / '.join(aceitos) + ('' if obrig else ' (opcional)')
                             for campo, obrig, aceitos in spec)
        raise ErroImportacao(f'Cabeçalho de {tipo} não encontrado. Colunas esperadas (por nome) – {esperado}. '
                             f'Use o arquivo exportado da FIC, sem editar.')
    aba, linhas, i, pos = achado
    ident = spec[0][0]
    chave_doc = 'sop' if tipo in ('TAP', 'TTAS') else 'ssop'
    regs, ignoradas, nao_dados = [], [], []
    for n, r in enumerate(linhas[i + 1:], i + 2):
        val = {campo: (texto(r[p]) if p < len(r) else '') for campo, p in pos.items()}
        if all(v == '' for v in val.values()):
            nao_dados.append(r)
            continue
        if _nome_coluna(val[ident]) in spec[0][2]:
            ignoradas.append(n)  # cabeçalho repetido
            nao_dados.append(r)
            continue
        if tipo == 'TAP' and val.get('sop', '') == '':
            m = re.match(r'TAP-(S?SOP-.+)$', val[ident].strip())
            val['sop'] = m.group(1) if m else ''   # o número do TAP é "TAP-" + SOP
        # registro sem número (ex.: CCMOCULTO) só vale com SSOP/SOP E status; rodapé não tem os dois
        sem_doc, sem_status = chave(val.get(chave_doc, '')) == '', chave(val['status']) == ''
        if (sem_doc and sem_status) or (val[ident] == '' and (sem_doc or sem_status)):
            ignoradas.append(n)  # rodapé/linha sem SSOP/SOP e sem status
            nao_dados.append(r)
            continue
        reg = {'linha': n}
        reg.update({campo: val.get(campo, '') for campo, _, _ in spec})
        regs.append(reg)
    if not regs:
        raise ErroImportacao(f'Nenhum registro de {tipo} encontrado abaixo do cabeçalho (aba {aba}).')
    return regs, data_do_rodape(nao_dados), {'aba': aba, 'registros': len(regs), 'linhas_ignoradas': ignoradas}


# ====================================================================== listas de status
def tipo_pelo_nome(arquivo):
    m = re.search(r'STATUS[_ -]?(FVI|FVM|CCM|TAP|TTAS)', Path(arquivo).stem.upper())
    return m.group(1) if m else None


def ler_status(arquivo):
    wb = _abrir(arquivo)
    linhas = list(wb.worksheets[0].iter_rows(values_only=True))
    wb.close()
    lista = [(texto(r[0]).strip(), texto(r[1]).strip() if len(r) > 1 else '') for r in linhas[1:] if r and texto(r[0]).strip()]
    if not lista:
        raise ErroImportacao(f'{Path(arquivo).name}: nenhuma linha de código de status (coluna A, a partir da linha 2).')
    return lista


# ====================================================================== templates (PDF da FIC)
RE_TEMPLATE = re.compile(r'Template:\s*(F[A-Z]{2}-[A-Z0-9]+?)\s*(?=Revis|$)', re.M)
RE_PERGUNTA = re.compile(r'(?<![A-Z0-9.])((?:III|IV|V)\.\d+)\s+(.+?)(?=\s*(?:(?:III|IV|V)\.\d+\s|ACEITA[ÇC][ÃA]O|V\)\s*TESTES|$))',
                         re.S)


def ler_templates(arquivo):
    """PDF de templates exportado da FIC – "Template de FVM - Folha de Verificação de Malhas" ou "Template de FVI -
    Folha de Verificação de Item" (este com a CLASSIFICAÇÃO = classe do item). Uma ou mais páginas por template.
    Devolve (documento, [{'template', 'revisao', 'disciplina', 'classe', 'perguntas': [(campo, item, texto)]}])."""
    from pypdf import PdfReader
    try:
        paginas = [p.extract_text() or '' for p in PdfReader(str(arquivo)).pages]
    except Exception as e:
        raise ErroImportacao(f'Não foi possível ler o PDF {Path(arquivo).name}: {e}') from e
    texto_todo = '\n'.join(paginas)
    if 'Template de FVI' in texto_todo:
        return 'FVI', _templates_fvi(paginas)
    if 'Template de FVM' not in texto_todo:
        raise ErroImportacao('PDF não reconhecido: esperado o "Template de FVM" ou o "Template de FVI" exportado da FIC.')
    blocos = {}
    for pg in paginas:                   # um template pode ocupar mais de uma página: junta pelo código
        m = RE_TEMPLATE.search(pg)
        if m:
            blocos.setdefault(m.group(1), []).append(pg)
    templates = []
    for codigo, pgs in blocos.items():
        t = '\n'.join(pgs)
        rev = re.search(r'Revis[ãa]o:\s*([\d.]+)', t)
        disc = re.search(r'DISCIPLINA:\s*(.+?)\s*DOCS REF', t, re.S)
        corpo = t[t.find('INSPEÇÃO MECÂNICA'):] if 'INSPEÇÃO MECÂNICA' in t else ''
        # o PDF cola o título da seção na 1ª pergunta ("TESTES FUNCIONAISV.7 ..."): separa antes de ler os itens
        corpo = re.sub(r'(INSPEÇÃO MECÂNICA|TESTES FUNCIONAIS)', r'\1\n', corpo)
        perguntas, vistos = [], set()
        for item, txt in RE_PERGUNTA.findall(corpo):
            if item in vistos:
                continue
            vistos.add(item)
            perguntas.append(('V' if item.startswith('V.') else 'III/IV', item, re.sub(r'\s+', ' ', txt).strip()))
        templates.append({'template': codigo, 'revisao': rev.group(1) if rev else '', 'classe': '',
                          'disciplina': re.sub(r'\s+', ' ', disc.group(1)).strip() if disc else '', 'perguntas': perguntas})
    if not templates:
        raise ErroImportacao('Nenhum template encontrado no PDF (linha "Template: FVM-...").')
    return 'FVM', templates


# seções do Template de FVI -> campo acompanhado (não há Campo IV no modelo)
SECOES_FVI = (('II) INSPEÇÃO TÉCNICA', 'II'), ('III) INSPEÇÃO MECÂNICA', 'III'), ('V) TESTES FUNCIONAIS', 'V'))
_CABECALHO_PAGINA = re.compile(r'^(Template de FVI\b.*|.+ / \d{4}\.\d+\.\d+\.\d+|Template:\s*\S+|Revisão:.*|Status:.*|'
                               r'Emissão:.*|\d\d/\d\d/\d{4}\s+\d\d:\d\d:\d\d.*|LEGENDA:.*)$')


def _templates_fvi(paginas):
    blocos = {}
    for pg in paginas:
        m = re.search(r'Template:\s*(FVI-[A-Z0-9]+)', pg)
        if m:
            blocos.setdefault(m.group(1), []).append(pg)
    templates = []
    for codigo, pgs in blocos.items():
        cab = pgs[0]
        rev = re.search(r'Revisão:\s*([\d.]+)', cab)
        disc = re.search(r'DISCIPLINA:\s*(.*?)\s*RM:', cab)
        classe = re.search(r'CLASSIFICAÇÃO:\s*(.*?)\s*MÓDULO:', cab)
        # tira cabeçalho/rodapé repetidos de cada página: uma seção pode continuar na página seguinte
        corpo = '\n'.join(ln for pg in pgs for ln in pg.splitlines() if not _CABECALHO_PAGINA.match(ln.strip()))
        perguntas = []
        for titulo, campo in SECOES_FVI:
            i = corpo.find(titulo)
            if i < 0:
                continue
            trecho = corpo[i + len(titulo):]
            fim = trecho.find('ACEITAÇÃO')
            trecho = trecho[:fim] if fim >= 0 else trecho
            for num, txt in re.findall(r'(?:^|\n)\s*(\d+)\)\s+(.+?)(?=\n\s*\d+\)\s|\Z)', trecho, re.S):
                perguntas.append((campo, f'{campo}.{num}', re.sub(r'\s+', ' ', txt).strip()))
        templates.append({'template': codigo, 'revisao': rev.group(1) if rev else '',
                          'disciplina': disc.group(1).strip() if disc else '',
                          'classe': re.sub(r'\s+', ' ', classe.group(1)).strip() if classe else '', 'perguntas': perguntas})
    if not templates:
        raise ErroImportacao('Nenhum template encontrado no PDF (linha "Template: FVI-...").')
    return templates
