"""Critérios e listas oficiais de status usados quando o banco é criado.

Origem: aba CRITERIOS da planilha CONTROLE_GERENCIAL_EPC12.xlsx (28/10/2026) e arquivos STATUS_*.xlsx da FIC.
Tudo pode ser alterado pelo usuário no programa (aba Critérios) ou reimportando as listas de status."""

U = 'critério definido pelo usuário'
CONFIRMAR_147 = 'A CONFIRMAR - fluxo FIC slide 147: status posterior à aprovação dos campos III/IV'
TTAS1_219 = 'fluxo FIC slide 219: registro único por SOP; status posterior ao TTAS1 aprovado'
TTAS2_219 = 'fluxo FIC slide 219: TTAS2 Aprov. Engª → APROVADO → TTAS CONCLUÍDO (não existe código "TTAS2APROVADO")'

# (ETAPA, CÓDIGO, ORIGEM). Origem começando com "A CONFIRMAR" = inferência ainda não confirmada (destacada em amarelo).
CRITERIOS = (
    [('FVI_CII', 'APROVADO', U)]
    + [('FVI_CIII', c, U) for c in ('APROVADO', 'ACAMPO5', 'ALIBTFUNC')]
    + [('FVI_CV', 'APROVADO', U)]
    + [('FVM_CIII', c, U) for c in ('ACAMPO5', 'ALIBTFUNC', 'APRVCQ')]
    + [('FVM_CIII', c, CONFIRMAR_147) for c in ('AAPROV', 'APRVSUP', 'APROVADO')]
    + [('FVM_CIII', 'LIBTFUNC', 'A CONFIRMAR - lista oficial STATUS_FVM: "FVM Liberada para Testes Funcionais" (posterior ao III/IV)')]
    + [('FVM_CV', 'APROVADO', U)]
    + [('CCM', c, U + ' (fluxo FIC slide 156)') for c in ('APROVADO', 'CCMCONCLUIDO')]
    + [('TAP', 'TAP CONCLUIDO', U + ' (fluxo FIC slide 208)')]
    + [('TTAS1', 'TTAS1APROVADO', U)]
    + [('TTAS1', c, TTAS1_219) for c in ('TTAS2AGAPROVA', 'TTAS2APROVENG', 'TTAS2REPROVADO', 'APROVADO', 'TTASCONCLUIDO')]
    + [('TTAS2', c, TTAS2_219) for c in ('APROVADO', 'TTASCONCLUIDO')]
    + [('CANCELADO', c, 'registro desconsiderado (cancelados podem ser desconsiderados)')
       for c in ('CANCELADO', 'CANCEL', 'TAPCANCELADO', 'TTASCANCELADO')]
)

ETAPAS = ('FVI_CII', 'FVI_CIII', 'FVI_CV', 'FVM_CIII', 'FVM_CV', 'CCM', 'TAP', 'TTAS1', 'TTAS2', 'CANCELADO')

# Listas oficiais de status da FIC (STATUS_FVI/FVM/CCM/TAP/TTAS.xlsx): documento -> [(código, descrição)]
STATUS_FIC = {
    'FVI': [('AAPROV', 'Aguardando Aprovação'), ('CONCLUIDO', 'Concluído'),
            ('ALIBTFUNC', 'Aguardando Liberação para Testes Funcionais'), ('CANCELADO', 'Cancelado'),
            ('AINSPTEC', 'Aguardando Inspeção Técnica'), ('ARECEB', 'Aguardando Recebimento'),
            ('APRVSUP', 'Aprovado pelo Supervisor'), ('APRVCQ', 'Aprovado pelo CQ'),
            ('ACAMPO3E4', 'Aguardando Campos III e IV'), ('ACAMPO5', 'Aguardando Campo V'), ('APROVADO', 'Aprovado')],
    'FVM': [('AAPROV', 'Aguardando Aprovação'), ('APRV', 'Aprovado'), ('CANCEL', 'Cancelado'),
            ('LIBEXEC', 'FVM liberada para preenchimento'), ('LIBTFUNC', 'FVM Liberada para Testes Funcionais'),
            ('CONCLUIDO', 'Concluído'), ('ALIBTFUNC', 'Aguardando Liberação para Testes Funcionais'),
            ('APRVSUP', 'Aprovado pelo Supervisor'), ('APRVCQ', 'Aprovado pelo CQ'), ('RASCUNHO', 'Rascunho'),
            ('ACAMPO5', 'Aguardando Campo V'), ('APROVADO', 'Aprovado'),
            ('AAPRV3E4CQ', 'Campo III e IV Aguardando Aprovação CQ'),
            ('AAPRV3E4SUP', 'Campo III e IV Aguardando Aprovação Supervisor'),
            ('AAPRV3E4FIS', 'Campo III e IV Aguardando Aprovação Fiscal')],
    'CCM': [('CCMOCULTO', 'Oculto CCM'), ('CCMCONCLUIDO', 'Concluído CCM'), ('CANCELADO', 'Cancelado'),
            ('AGAPROVA', 'Aguardando Aprovação'), ('RASCUNHO', 'Rascunho'), ('REPROVADA', 'Reprovada'),
            ('APROVADO', 'Aprovado')],
    'TAP': [('TAP2APROVACQ', 'TAP2 Aprovado pelo CQ'), ('TAP2APROVSUP', 'TAP2 Aprovado pelo Supervisor'),
            ('TAP2APROVFIS', 'TAP2 Aprovado pelo Fiscal'), ('TAP1ANDAMENTO', 'TAP1 em Andamento'),
            ('TAP1CONCLUIDO', 'TAP1 concluído'), ('TAP2LIBEXEC', 'Liberado para Execução do TAP2'),
            ('TAP2ANDAMENTO', 'TAP2 em Andamento'), ('TAPSOLICEXEC', 'Solicitação para Execução'),
            ('TAP2APROVCQ', 'TAP2 Aprovado pelo CQ'), ('TAP2REPROVADO', 'TAP2 reprovado'),
            ('TAP1REPROVADO', 'TAP1 reprovado'), ('TAPAGAPROVPROC', 'Aguardando Aprovação do Procedimento'),
            ('TAPPROCAPROV', 'Procedimento Aprovado'), ('TAPAGPROGRAM', 'Aguardando Programação'),
            ('TAPAGCONRECUR', 'Aguardando Confirmação de Recursos'), ('TAP1APROVSUP', 'TAP1 Aprovado pelo Supervisor'),
            ('TAP1APROVCQ', 'TAP1 Aprovado pelo CQ'), ('TAP1APROVFIS', 'TAP1 Aprovado pelo Fiscal'),
            ('TAP1AGAPROVA', 'TAP1 Aguardando Aprovação'), ('TAP2AGAPROVA', 'TAP2 Aguardando Aprovação'),
            ('TAPPROCREPFISCAL', 'Procedimento TAP Reprovado pelo Fiscal'), ('TAP2AGPROGRAM', 'Aguardando Programação TAP2'),
            ('TAP2AGCONRECUR', 'Aguardando Confirmação de Recursos TAP2'), ('TAP2SOLICEXEC', 'Solicitação para Execução TAP2'),
            ('RASCUNHO', 'Rascunho'), ('TAP CONCLUIDO', 'TAP Concluído'), ('TAPCANCELADO', 'TAP cancelado'),
            ('TAP1LIBEXEC', 'Liberado para Execução do TAP1')],
    'TTAS': [('TTASOCULTO', 'TTAS Oculto'), ('TTAS1APROVENG', 'TTAS1 Aprovado pela Engenharia'),
             ('TTAS2APROVENG', 'TTAS2 Aprovado pela Engenharia'), ('TTAS1APROVADO', 'TTAS1 Aprovado'),
             ('TTAS1REPROVADO', 'TTAS1 Reprovado'), ('TTAS2AGAPROVA', 'TTAS2 Aguardando Aprovação'),
             ('TTAS1AGAPROVA', 'TTAS1 Aguardando Aprovação'), ('TTAS2REPROVADO', 'TTAS2 Reprovado'),
             ('TTASCONCLUIDO', 'TTAS Concluído'), ('RASCUNHO', 'Rascunho'), ('APROVADO', 'Aprovado'),
             ('TTASCANCELADO', 'TTAS Cancelado')],
}
