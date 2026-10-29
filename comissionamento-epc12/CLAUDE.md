# Controle Gerencial de Comissionamento – EPC 12 (programa desktop)

Programa Windows (Python + CustomTkinter/Pillow + SQLite + openpyxl, empacotado com PyInstaller em um .exe único) que substitui a
geração da planilha `CONTROLE_GERENCIAL_EPC12.xlsx`, mantendo as mesmas regras e números. Idioma: **português do Brasil**.

## Protocolo de trabalho do usuário (obrigatório)
- **Evidência em vez de suposição.** Não conserte o que não foi comprovado quebrado. **Reproduza antes de corrigir.**
- **Hipótese antes do código.** Isole variáveis até restar um único ponto de falha. **Fatos, não intuição.**
- **Nunca corrigir texto das fontes** (LI, cadastro). Inconsistências são reportadas (aba Verificação/pendências).
  Inferências entram marcadas "A CONFIRMAR" (amarelo) com a origem da evidência; o usuário decide.

## Dados – não versionar
O repositório é **público**. Arquivos da FIC/LI são classificação NP-1: ficam em `dados/` (no .gitignore). Testes do git
usam só dados fictícios (`testes/fabrica.py`).

## Estrutura
```
main.py                      entrada: sem argumentos abre a interface; com --li/--importar/--exportar roda por linha de comando
comissionamento/texto.py     TRIM e comparação sem caixa (equivalentes ao Excel)
comissionamento/padroes.py   critérios padrão e listas oficiais de status (semeiam o banco novo)
comissionamento/banco.py     esquema SQLite, critérios, parâmetros, log de importações, histórico, backup
comissionamento/importacao.py leitura da LI, do Relatório de Cadastro Geral e dos layouts provisórios (FVM/CCM/TAP/TTAS)
comissionamento/calculo.py   reprodução das fórmulas da planilha (IMP_* auxiliares, abas por SSOP/SOP, visão, verificação)
comissionamento/servico.py   importar -> gravar -> recalcular -> histórico
comissionamento/exportacao.py relatório Excel (valores formatados) e planilhas-modelo
comissionamento/gui.py       interface CustomTkinter: barra lateral com logo, Visão Gerencial (só os 5 gráficos: FVI, FVM,
                             TTAS, CCM, TAP – por unidade + Total), FVI TAG a TAG, FVM malha a malha, CCM/TAP/TTAS,
                             Verificação, Histórico, Importações, Critérios. Filtros: Unidade → SOP → SSOP em cascata,
                             + TAG/malha (FVI/FVM) ou situação (CCM/TAP/TTAS)
comissionamento/recursos/    logo.png (Consórcio HC4), icone.ico/icone.png – embutidos no .exe pelo .spec
comissionamento/rotulos.py   TODOS os textos e colunas exibidos (tela e Excel), por extenso – sem abreviações
comissionamento/tema.py      paleta suave, fontes e estilo das tabelas; cores de série validadas p/ daltonismo (3 cores)
comissionamento/icones.py    ícones de linha padronizados (Pillow, 4x e reduzidos), recurso() e logo()
comissionamento/graficos.py  gráficos em tk.Canvas (colunas agrupadas, linhas) com dica ao passar o mouse
testes/                      sintético (12 casos de borda à mão), importação, interface, dados reais (pulado sem dados/)
ci/teste_exe.py              testa o .exe gerado; .github/workflows/comissionamento-windows.yml gera o .exe no Windows
```
Rodar: `python -m pytest testes -q` (interface precisa de display; no Linux use `xvfb-run`).

## Regras de negócio (vindas da planilha – não mudar sem evidência)
- Base mestre = LI (aba LISTA). Chave = SSOP. Checagens da LI (SSOP único, prefixo "S"+SOP, unidade = código, ITEM 1..n) param a importação.
- Granularidade: FVI TAGs por SSOP e SOP · FVM malhas por SSOP e SOP · CCM e TAP por SSOP · TTAS1/TTAS2 por SOP · Visão por unidade.
- Cadastro FVI: linha de dado = TAG (A) e CONTRATO (AC) preenchidos e TAG ≠ "TAG". Layout validado pelos nomes das colunas-chave
  (A TAG, D SOP, E SSOP, F CLASSE, AC CONTRATO, AE Nº FVI, AJ..AN STATUS FVI/CAMPO II/III/IV/V); se mudar, para.
- Cancelados (etapa CANCELADO) fora de todas as contas; no FVI cancelamento verificado em AJ, AK, AL e AN.
- FVI: "N/A" tira a TAG do denominador do campo; status vazio = aplicável e não aprovado. TAGs sem FVI no denominador
  conforme parâmetro (SIM padrão). TAGs sem SSOP não contam avanço. Campo IV não é acompanhado (aviso se aparecer).
- TTAS é registro único por SOP que avança: TTAS1 aprovado inclui estágios posteriores; TTAS2 = APROVADO/TTASCONCLUIDO.
- % = aprovados ÷ aplicáveis; % CCM/TAP por SSOP; % TTAS por SOP; totais ponderados (não média de %).
- Comparações como no Excel: TRIM (só espaço) e sem diferença de maiúsculas.
- Diferenças deliberadas em relação à planilha: FVM/CCM/TAP/TTAS ignoram linhas sem SSOP/SOP **e** sem status (rodapé) e
  contam "registros sem SSOP" à parte; colunas são achadas pelo nome. Sem limites de linhas (a planilha tinha 10.000 no FVI).
- Textos ao usuário sempre por extenso ("FVI Campo II", "FVM Campo III", "TTAS 1"), definidos em rotulos.py e
  usados igualmente na tela e no Excel. Ícones só do conjunto de icones.py (mesmo traço).
- ci/teste_exe.py abre o .exe por linha de comando E a interface (variável EPC12_TESTE_INICIO); não remover.
- Gráficos: no máximo 3 séries por gráfico (paleta validada só até 3 cores em todos os pares); legenda com 2+ séries
  (com uma só, o título identifica).
- Layout das telas definido pelo usuário (29/10): Visão Gerencial só com gráficos; FVI/FVM sem coluna Situação;
  CCM/TAP só com a situação; TTAS com "Situação TTAS 1/2"; menu Importar = SSOP, FVI, FVM, CCM, TAP, TTAS.
- Banco: cada importação substitui a anterior do tipo; o histórico guarda contagens por unidade por data de referência.

## Pendências (perguntas abertas ao usuário)
1. FVM `LIBTFUNC`, `AAPROV`, `APRVSUP`, `APROVADO` no Campo III (incluídos, amarelo) – confirmar.
2. FVM `APRV` ≠ `APROVADO`: em que etapa aparece? Hoje fora dos critérios.
3. TAP `TAP2APROVACQ` e `TAP2APROVCQ` com a mesma descrição – equivalentes?
4. Parâmetro de TAGs sem FVI no denominador: SIM ou NÃO – qual a gerência usa?
5. Templates FVI/FVM/Preservação ainda não foram trazidos para o programa (continuam na planilha).
6. TAGS SEM FVI por SSOP soma também TAGs canceladas (como a planilha); hoje não há TAG cancelada, sem efeito – confirmar a regra.
7. Layouts reais de FVM, CCM, TAP e TTAS: ajustar `COLUNAS_PROVISORIAS` quando chegarem as exportações.
