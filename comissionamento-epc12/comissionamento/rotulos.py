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
             ('SSOP com TAP concluído', 'tap', 20, 0), ('TAP (%)', 'pct_tap', 9, 1),
             ('SOP com TTAS 1 aprovado', 'ttas1', 21, 0), ('TTAS 1 (%)', 'pct_ttas1', 10, 1),
             ('SOP com TTAS 2 aprovado', 'ttas2', 21, 0), ('TTAS 2 (%)', 'pct_ttas2', 10, 1),
             ('TAGs sem FVI gerada', 'tags_sem_fvi', 17, 0)]
COL_TAGS = [('TAG', 'tag', 18, 0), ('SSOP', 'ssop', 27, 0),
            ('FVI Campo II', 'txt_cii', 24, 0), ('FVI Campo III', 'txt_ciii', 24, 0), ('FVI Campo V', 'txt_cv', 24, 0),
            ('Situação', 'txt_situacao', 18, 0)]
COL_TAGS_XLS = [('Unidade', 'unidade', 10, 0), ('SOP', 'sop', 22, 0), ('SSOP', 'ssop', 28, 0), ('TAG', 'tag', 20, 0),
                ('Descrição da TAG', 'descricao', 36, 0), ('Nº da FVI', 'n_fvi', 10, 0), ('FVI Campo II', 'txt_cii', 24, 0),
                ('FVI Campo III', 'txt_ciii', 24, 0), ('FVI Campo V', 'txt_cv', 24, 0), ('Situação', 'txt_situacao', 18, 0)]
COL_MALHAS = [('Malha', 'malha', 22, 0), ('SSOP', 'ssop', 27, 0), ('FVM Campo III', 'txt_ciii', 24, 0),
              ('FVM Campo V', 'txt_cv', 24, 0), ('Situação', 'txt_situacao', 16, 0)]
COL_MALHAS_XLS = [('Unidade', 'unidade', 10, 0), ('SOP', 'sop', 22, 0), ('SSOP', 'ssop', 28, 0), ('Malha', 'malha', 22, 0),
                  ('FVM Campo III', 'txt_ciii', 26, 0), ('FVM Campo V', 'txt_cv', 26, 0), ('Situação', 'txt_situacao', 18, 0)]
_SSOP = [('Unidade', 'unidade', 10, 0), ('SOP', 'sop', 22, 0), ('SSOP', 'ssop', 27, 0), ('Descrição da SSOP', 'descricao', 34, 0)]
COL_CCM = _SSOP + [('CCM registrados', 'ccm_reg', 14, 0), ('CCM aprovado', 'txt_ccm_aprovado', 13, 2),
                   ('Situação do CCM', 'txt_ccm_situacao', 16, 0)]
COL_TAP = _SSOP + [('TAP registrados', 'tap_reg', 14, 0), ('TAP concluído', 'txt_tap_concluido', 13, 2),
                   ('Situação do TAP', 'txt_tap_situacao', 16, 0)]
COL_TTAS = [('Unidade', 'unidade', 10, 0), ('SOP', 'sop', 22, 0), ('Descrição do SOP', 'descricao', 40, 0),
            ('Quantidade de SSOP', 'qtd_ssop', 16, 0), ('TTAS registrados', 'ttas_reg', 15, 0),
            ('TTAS 1 aprovado', 'txt_ttas1', 14, 2), ('TTAS 2 aprovado', 'txt_ttas2', 14, 2),
            ('Situação do TTAS', 'txt_ttas_situacao', 17, 0)]
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
    'CCM e TAP: SSOP com documento aprovado/concluído ÷ total de SSOP da unidade. TTAS: SOP com TTAS aprovado ÷ total de SOP.',
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
    for s in R.ssop:
        s['txt_ccm_aprovado'] = sim_nao(s['ccm_aprovado'])
        s['txt_ccm_situacao'] = SITUACAO_DOC[s['ccm_situacao']]
        s['txt_tap_concluido'] = sim_nao(s['tap_concluido'])
        s['txt_tap_situacao'] = SITUACAO_DOC[s['tap_situacao']]
    for s in R.sop:
        s['txt_ttas1'] = sim_nao(s['ttas1_aprovado'])
        s['txt_ttas2'] = sim_nao(s['ttas2_aprovado'])
        s['txt_ttas_situacao'] = SITUACAO_DOC[s['ttas_situacao']]
    return R
