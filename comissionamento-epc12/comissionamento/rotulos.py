"""Textos exibidos ao usuário (tela e Excel), por extenso e sem abreviações.

Colunas no formato (título, chave, largura em caracteres, é percentual)."""

SITUACAO_DOC = {'APROVADO': 'Aprovado', 'CONCLUÍDO': 'Concluído', 'EM ANDAMENTO': 'Em andamento',
                'SEM REGISTRO': 'Sem registro', 'TTAS2 APROVADO': 'TTAS 2 aprovado', 'TTAS1 APROVADO': 'TTAS 1 aprovado'}
SITUACAO_ITEM = {'VÁLIDA': 'Válida', 'CANCELADA': 'Cancelada', 'SEM SSOP': 'Sem SSOP', 'FORA DA BASE': 'SSOP fora da base'}

COL_VISAO = [('Unidade', 'unidade', 11, 0), ('SOP', 'sop', 6, 0), ('SSOP', 'ssop', 6, 0), ('TAGs (FVI)', 'tags', 9, 0),
             ('FVI Campo II (%)', 'pct_cii', 14, 1), ('FVI Campo III (%)', 'pct_ciii', 15, 1),
             ('FVI Campo V (%)', 'pct_cv', 14, 1), ('Malhas (FVM)', 'malhas', 11, 0),
             ('FVM Campo III (%)', 'pct_fvm_ciii', 16, 1), ('FVM Campo V (%)', 'pct_fvm_cv', 15, 1),
             ('SSOP com CCM aprovado', 'ccm', 20, 0), ('CCM (%)', 'pct_ccm', 9, 1),
             ('SOP com TAP concluído', 'tap', 20, 0), ('TAP (%)', 'pct_tap', 9, 1),
             ('SOP com TTAS 1 aprovado', 'ttas1', 21, 0), ('TTAS 1 (%)', 'pct_ttas1', 10, 1),
             ('SOP com TTAS 2 aprovado', 'ttas2', 21, 0), ('TTAS 2 (%)', 'pct_ttas2', 10, 1),
             ('TAGs sem FVI gerada', 'tags_sem_fvi', 17, 0)]
# acompanhamento: o status exatamente como vem da FIC
COL_TAGS = [('TAG', 'tag', 18, 0), ('SSOP', 'ssop', 27, 0),
            ('FVI Campo II', 'st_cii', 16, 0), ('FVI Campo III', 'st_ciii', 16, 0), ('FVI Campo V', 'st_cv', 16, 0)]
COL_TAGS_XLS = [('Unidade', 'unidade', 10, 0), ('SOP', 'sop', 22, 0), ('SSOP', 'ssop', 28, 0), ('TAG', 'tag', 20, 0),
                ('Descrição da TAG', 'descricao', 36, 0), ('Nº da FVI', 'n_fvi', 10, 0), ('Status FVI', 'st_fvi', 16, 0),
                ('FVI Campo II', 'st_cii', 16, 0), ('FVI Campo III', 'st_ciii', 16, 0), ('FVI Campo V', 'st_cv', 16, 0)]
COL_MALHAS = [('Malha', 'malha', 22, 0), ('SSOP', 'ssop', 27, 0), ('Status da malha', 'status', 18, 0)]
COL_MALHAS_XLS = [('Unidade', 'unidade', 10, 0), ('SOP', 'sop', 22, 0), ('SSOP', 'ssop', 28, 0), ('Malha', 'malha', 22, 0),
                  ('Status da malha', 'status', 18, 0), ('Malha em mais de uma SSOP', 'txt_multi', 12, 2)]
_SSOP = [('Unidade', 'unidade', 10, 0), ('SOP', 'sop', 22, 0), ('SSOP', 'ssop', 27, 0), ('Descrição da SSOP', 'descricao', 34, 0)]
COL_CCM = _SSOP + [('Status do CCM', 'txt_ccm_status', 18, 2)]
_SOP = [('Unidade', 'unidade', 10, 0), ('SOP', 'sop', 22, 0), ('Descrição do SOP', 'descricao', 46, 0)]
COL_TAP = _SOP + [('Status do TAP', 'txt_tap_status', 18, 2)]
COL_TTAS = _SOP + [('Situação TTAS 1', 'ttas1_status', 18, 2), ('Situação TTAS 2', 'ttas2_status', 18, 2)]
SEM_REGISTRO = 'Sem registro'
SEM_STATUS = 'Sem status'
OUTROS = 'Outros'
COL_HIST = [('Data de referência', 'data', 15, 0), ('Unidade', 'unidade', 11, 0), ('TAGs (FVI)', 'tags', 9, 0),
            ('FVI Campo II (%)', 'pct_cii', 14, 1), ('FVI Campo III (%)', 'pct_ciii', 15, 1), ('FVI Campo V (%)', 'pct_cv', 14, 1),
            ('Malhas (FVM)', 'malhas', 11, 0), ('FVM Campo III (%)', 'pct_fvm_ciii', 16, 1),
            ('FVM Campo V (%)', 'pct_fvm_cv', 15, 1), ('CCM (%)', 'pct_ccm', 9, 1), ('TAP (%)', 'pct_tap', 9, 1),
            ('TTAS 1 (%)', 'pct_ttas1', 10, 1), ('TTAS 2 (%)', 'pct_ttas2', 10, 1), ('TAGs sem FVI gerada', 'tags_sem_fvi', 17, 0)]
COL_VER = [('Grupo', 'g', 10, 0), ('Teste', 't', 66, 0), ('Esperado', 'e', 9, 0), ('Apurado', 'a', 9, 0),
           ('Resultado', 'r', 11, 2)]
COL_PEND = [('Documento', 'd', 10, 0), ('Linha no arquivo', 'l', 13, 0), ('Identificação', 'i', 22, 0),
            ('SSOP / SOP', 's', 28, 0), ('Problema', 'p', 64, 0)]
COL_CRIT = [('Etapa', 'etapa', 11, 0), ('Código de status', 'codigo', 18, 0), ('Origem / justificativa', 'origem', 52, 0)]
COL_STATUS = [('Documento', 'doc', 10, 0), ('Código', 'codigo', 16, 0), ('Descrição', 'descricao', 26, 0)]
COL_IMP = [('Nº', 'id', 5, 0), ('Tipo', 'tipo', 7, 0), ('Importado em', 'importado_em', 18, 0),
           ('Data de referência', 'data_br', 15, 0), ('Linhas', 'linhas', 8, 0), ('Arquivo', 'arquivo', 52, 0),
           ('Observação', 'observacao', 44, 0)]

# (chave do indicador, título por extenso)
INDICADORES = [('pct_cii', 'FVI Campo II'), ('pct_ciii', 'FVI Campo III'), ('pct_cv', 'FVI Campo V'),
               ('pct_fvm_ciii', 'FVM Campo III'), ('pct_fvm_cv', 'FVM Campo V'), ('pct_ccm', 'CCM'),
               ('pct_tap', 'TAP'), ('pct_ttas1', 'TTAS 1'), ('pct_ttas2', 'TTAS 2')]

NOTAS_VISAO = [
    'FVI e FVM: aprovados ÷ aplicáveis ("Não se aplica" e cancelados ficam fora da conta).',
    'CCM: SSOP com CCM aprovado ÷ total de SSOP da unidade. TAP e TTAS: SOP com documento concluído/aprovado ÷ total de SOP.',
    'Os totais são ponderados (não é a média dos percentuais). Itens sem SSOP ou com SSOP fora da base não entram – ver Verificação.',
]


def campo(status, ok, situacao):
    """Texto de um campo: 'Aprovado (APROVADO)', 'Pendente (ACAMPO3E4)', 'Não se aplica', 'Cancelada'."""
    st = (status or '').strip()
    if situacao == 'CANCELADA':
        return 'Cancelada'
    if ok is None:
        return 'Não se aplica'
    if ok:
        return 'Aprovado' if st.upper() == 'APROVADO' else f'Aprovado ({st})'
    return f'Pendente ({st})' if st else 'Pendente (sem status)'


def sim_nao(v):
    return 'Sim' if v else 'Não'


def distribuicao(valores, max_fatias=7):
    """[(status, quantidade)] do maior para o menor; acima de max_fatias, o resto vira 'Outros'."""
    from collections import Counter
    c = Counter((v or '').strip() or SEM_STATUS for v in valores)
    itens = c.most_common()
    if len(itens) > max_fatias:
        itens = itens[:max_fatias - 1] + [(OUTROS, sum(n for _, n in itens[max_fatias - 1:]))]
    return itens


def preparar(R):
    """Acrescenta aos resultados os textos de exibição (chaves txt_*). Idempotente."""
    for t in R.tags:
        for c in ('cii', 'ciii', 'cv'):
            t['txt_' + c] = campo(t['st_' + c], t['ok_' + c], t['situacao'])
        t['txt_situacao'] = 'Sem FVI gerada' if t['situacao'] == 'VÁLIDA' and t['sem_fvi'] else SITUACAO_ITEM[t['situacao']]
    for m in R.malhas:
        for c in ('ciii', 'cv'):
            m['txt_' + c] = campo(m['status'], m['ok_' + c], m['situacao'])
        m['txt_situacao'] = SITUACAO_ITEM[m['situacao']]
        m['txt_multi'] = 'Sim' if m.get('multi_ssop') else ''
    for s in R.ssop:
        s['txt_ccm_status'] = s['ccm_status'] or SEM_REGISTRO
        s['txt_ccm_aprovado'] = sim_nao(s['ccm_aprovado'])
        s['txt_ccm_situacao'] = SITUACAO_DOC[s['ccm_situacao']]
    for s in R.sop:
        s['txt_tap_status'] = s['tap_status'] or SEM_REGISTRO
        s['txt_tap_concluido'] = sim_nao(s['tap_concluido'])
        s['txt_tap_situacao'] = SITUACAO_DOC[s['tap_situacao']]
        s['txt_ttas_status'] = s['ttas_status'] or SEM_REGISTRO
        andamento = 'Em andamento' if s['ttas_reg'] else 'Sem registro'
        s['txt_ttas1'] = 'Aprovado' if s['ttas1_aprovado'] else andamento
        s['txt_ttas2'] = 'Aprovado' if s['ttas2_aprovado'] else andamento
        s['txt_ttas_situacao'] = SITUACAO_DOC[s['ttas_situacao']]
    return R
