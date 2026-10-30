# Controle Gerencial de Comissionamento – EPC 12 (programa desktop)

Programa Windows (Python + CustomTkinter/Pillow + SQLite + openpyxl + fpdf2 + pypdf, empacotado com PyInstaller em um .exe único) que substitui a
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
comissionamento/importacao.py leitura das exportações da FIC: Cadastros Básicos (SOP/SSOP; LI antiga também), Cadastro Geral
                             (FVI), FVM BASE (Hierarquia de Malhas), Lista de FVM, listas de CCM, TAP e TTAS,
                             PDFs "Template de FVI" e "Template de FVM" (perguntas por campo, lidos com pypdf)
comissionamento/calculo.py   reprodução das fórmulas da planilha (IMP_* auxiliares, abas por SSOP/SOP, visão, verificação)
comissionamento/servico.py   importar -> gravar -> recalcular -> histórico
comissionamento/exportacao.py relatório Excel + PDF: só Visão Gerencial, FVI, FVM, CCM, TAP, TTAS, com a logo
comissionamento/gui.py       interface CustomTkinter: barra lateral com logo, botão Importar atualização (SOP/SSOP + FIC),
                             Visão Gerencial (5 pizzas de status da FIC: FVI, FVM, CCM, TAP, TTAS – geral ou por unidade),
                             FVI TAG a TAG, FVM malha a malha, CCM/TAP/TTAS (status da FIC), Verificação, Histórico,
                             Importações, Critérios. Filtros: Unidade → SOP → SSOP em cascata + TAG/malha ou situação
comissionamento/recursos/    logo.png (Consórcio HC4, monocromática), icone.ico/icone.png e fontes/ (OFL) – embutidos no .exe pelo .spec
comissionamento/rotulos.py   TODOS os textos e colunas exibidos (tela e Excel), por extenso – sem abreviações
comissionamento/tema.py      identidade HUD (tokens hud-mono-light) + paleta categórica colorida dos gráficos
comissionamento/icones.py    ícones de linha padronizados (Pillow, 4x e reduzidos), recurso() e logo()
comissionamento/graficos.py  gráficos em tk.Canvas: Pizza (imagem_pizza, Pillow 4x – usada também no PDF) e linhas
ferramentas/fonte_display.py gera EPC12 Display (Orbitron com o til redesenhado; nome trocado por exigência da OFL)
testes/                      sintético (12 casos de borda à mão), importação, layouts da FIC, interface, dados reais (pulado sem dados/)
ci/teste_exe.py              testa o .exe gerado; .github/workflows/comissionamento-windows.yml gera o .exe no Windows
```
Rodar: `python -m pytest testes -q` (interface precisa de display; no Linux use `xvfb-run`).

## Regras de negócio (vindas da planilha – não mudar sem evidência)
- Base mestre = lista SOP / SSOP. Modelo (29/10): Lista de Cadastros Básicos da FIC (Código, Descrição, Tipo SOP/SSOP), sem
  coluna de unidade: unidade e SOP da SSOP vêm do código (SSOP-<unidade>-<grupo>-<nº>-<seq>). Código duplicado, tipo
  inválido ou código sem unidade param a importação; SSOP cujo SOP não está na lista (SOP incluído sem descrição, A CONFIRMAR)
  e código de SSOP sem "SSOP-" viram avisos na Verificação. A LI antiga (aba LISTA) segue aceita com as checagens da
  planilha (SSOP único, prefixo "S"+SOP, unidade = código, ITEM 1..n) – é ela que o teste de dados reais usa.
- Granularidade: FVI TAGs por SSOP e SOP · FVM malhas por SSOP e SOP · CCM por SSOP · TAP e TTAS1/TTAS2 por SOP · Visão por unidade.
  TAP por SOP (mudou em 29/10 pela lista da FIC: Número do TAP = "TAP-" + SOP, 14 TAP = 14 SOP); sem coluna SOP, vem do número.
- Cadastro FVI: linha de dado = TAG (A) e CONTRATO (AC) preenchidos e TAG ≠ "TAG". Layout validado pelos nomes das colunas-chave
  (A TAG, D SOP, E SSOP, F CLASSE, AC CONTRATO, AE Nº FVI, AJ..AN STATUS FVI/CAMPO II/III/IV/V); se mudar, para.
- Cancelados (etapa CANCELADO) fora de todas as contas; no FVI cancelamento verificado em AJ, AK, AL e AN.
- FVI: "N/A" tira a TAG do denominador do campo; status vazio = aplicável e não aprovado. TAGs sem FVI no denominador
  conforme parâmetro (SIM padrão). TAGs sem SSOP não contam avanço. Campo IV não é acompanhado (aviso se aparecer).
- TTAS é registro único por SOP que avança: TTAS1 aprovado inclui estágios posteriores; TTAS2 = APROVADO/TTASCONCLUIDO.
- % = aprovados ÷ aplicáveis; % CCM por SSOP; % TAP e TTAS por SOP; totais ponderados (não média de %).
- Comparações como no Excel: TRIM (só espaço) e sem diferença de maiúsculas.
- Diferenças deliberadas em relação à planilha: CCM/TAP/TTAS ignoram linhas sem SSOP/SOP **e** sem status (rodapé); linha
  sem número só vale com SSOP/SOP E status (CCMOCULTO); colunas são achadas pelo nome. Sem limites de linhas.
- FVM (definido pelo usuário em 29/10, 3ª versão): FVM BASE = Relatório de Hierarquia de Malhas (Importar atualização ▸
  FVM BASE; tipo FVM_BASE, tabela fvm; sem data/histórico): define as malhas, uma linha por TAG/LINHA com SSOP/SOP
  ("-" = vazio). Atualização FVM = Lista de FVM da FIC (Código FVM, Malha, Status em CÓDIGO; tipo FVM, tabela fvm_status);
  o Relatório de Status de Malhas (status em descrição) também é aceito. Hierarquia enviada como atualização é recusada;
  atualização sem FVM BASE é recusada. FVM de malha fora da base NÃO vira linha (vai para a Verificação). Malha com mais de
  uma FVM: vale a aprovada, senão a última válida, senão a cancelada. Status vale pelo código ou pela descrição oficial.
  Avanço: cada malha uma vez por SSOP. Rodapé ("Emissão", com o nome do emissor) é descartado. Página FVM: Malha, TAG,
  SSOP, Status da FVM; busca por malha ou TAG; contador = malhas distintas. Pizza da Visão: cada malha uma vez.
  Modelo real: 505 malhas na base = 505 FVMs na lista, todas APROVADO.
- Templates de FVM (30/10): Importar atualização ▸ Templates de FVM (PDF) lê o "Template de FVM - Folha de Verificação de
  Malhas" da FIC (uma página por template: código, revisão, disciplina, perguntas III.n do Campo III/IV e V.n do Campo V)
  para a tabela template_pergunta. O PDF é NP-1: as perguntas ficam só no banco, nunca no repositório. A malha liga ao
  template pela coluna "Código do Template" da Lista de FVM. Clique numa linha da página FVM → JanelaFolha com as perguntas.
  O PDF cola o título na 1ª pergunta do Campo V ("TESTES FUNCIONAISV.7"): o leitor separa antes (teste cobre). Leitor
  conferido contra pdftotext nos 8 templates do modelo (itens e textos iguais). Template usado sem estar no PDF → Verificação.
- Templates de FVI (30/10, PDF "Template de FVI - Folha de Verificação de Item", 627 páginas, 284 templates, 6564
  perguntas): mesmo menu ("Templates de FVI / FVM (PDF)", tipo reconhecido no texto). Cada template: CLASSIFICAÇÃO = classe,
  seções II) INSPEÇÃO TÉCNICA → Campo II, III) INSPEÇÃO MECÂNICA → Campo III, V) TESTES FUNCIONAIS → Campo V (não há IV),
  perguntas "N)" (item II.N etc.). Cabeçalho/rodapé de cada página são removidos antes (seção pode mudar de página).
  Leitor conferido contra pdftotext: 281/284 iguais; os 3 restantes (FVI-011, -272, -280) vêm SEM classificação no PDF.
  TAG → template (calculo.template_da_classe), nesta ordem: texto da classe igual · mesmo código "(NNN[-X])" · família
  (código principal, 007-C → 007; A CONFIRMAR) · número do template sem classificação = código (FVI-011 ↔ (011);
  A CONFIRMAR – evidência: nos 222 templates com código, número = código sempre). Cobertura: modelo 827 TAGs (485 classe,
  306 família, 36 número, 0 sem template); cadastro antigo 3915 (1195 / 2612 / 108 / 0). Clique numa TAG da página FVI →
  JanelaFolha com as perguntas por campo e o status da FIC de cada campo; inferência aparece com aviso A CONFIRMAR.
- Textos ao usuário sempre por extenso ("FVI Campo II", "FVM Campo III", "TTAS 1"), definidos em rotulos.py e
  usados igualmente na tela e no Excel. Ícones só do conjunto de icones.py (mesmo traço).
- ci/teste_exe.py abre o .exe por linha de comando E a interface (variável EPC12_TESTE_INICIO); não remover.
- Identidade visual (29/10, arquivo hud-mono-light.tokens.json do usuário): interface preto/cinzas sobre branco; painéis retos
  com cantoneiras; botões raio 2; títulos/rótulos/botões em caixa-alta; display EPC12 Display, títulos Rajdhani, corpo IBM Plex
  Mono, dados Share Tech Mono. Usar só os tokens de tema.py.
- Gráficos COLORIDOS (pedido do usuário, 29/10): tema.CATEGORICAS = paleta categórica validada (scripts do skill dataviz:
  daltonismo e visão normal OK; 4 cores < 3:1 no fundo → legenda com quantidade e % sempre visível). Ordem fixa, não mudar
  sem revalidar. Linhas: as 3 primeiras cores (+ traço e marcador). Texto sobre a fatia: preto ou branco pelo contraste.
- Layout das telas definido pelo usuário (29/10): Visão Gerencial = pizzas FVI, FVM, CCM, TAP, TTAS (nesta ordem, caixas
  iguais); páginas de acompanhamento mostram o STATUS DA FIC como vem no arquivo (o avanço % continua pelos Critérios);
  FVI/FVM sem coluna Situação; TTAS com "Situação TTAS 1/2". Em 29/10 (2ª): botão Importar base removido; um só botão
  Importar atualização = SOP / SSOP, FVM BASE | FVI, FVM, CCM, TAP, TTAS. LI de Itens / LI de Malhas e planilhas-modelo
  removidas.
- Status exibido por SSOP/SOP (CCM, TAP, TTAS): registro aprovado > último válido do arquivo > cancelado > "Sem registro".
  TTAS: status de etapa 2 (TTAS1 sem TTAS1APROVADO, ou TTAS2) ⇒ coluna TTAS 1 = TTAS1APROVADO (inferência, a confirmar).
- FVM: uma malha PODE, mas NÃO DEVE, pertencer a mais de uma SSOP – conta em cada SSOP, é destacada e vai para a
  Verificação (teste que deve dar 0).
- Estilo das pizzas: tema.cores_pizza devolve (tipo, cor) – a cor segue o código na lista oficial (mesmo status, mesma cor);
  N/A cinza, sem status/registro branco com contorno, Outros hachura cinza; tela, Excel e PDF iguais.
- Exportação (29/10): só Visão Gerencial, FVI, FVM, CCM, TAP, TTAS – Excel (destino .xlsx; logo em cada aba; números das
  pizzas abaixo delas, A60+) e PDF de mesmo nome (A4 paisagem, logo/título/data em cada página, fontes embutidas, colunas
  da tela). Verificação, Pendências, Histórico e Critérios não são mais exportados.
- Banco: cada importação substitui a anterior do tipo; o histórico guarda contagens por unidade por data de referência.

## Pendências (perguntas abertas ao usuário)
1. FVM `LIBTFUNC`, `AAPROV`, `APRVSUP`, `APROVADO` no Campo III (incluídos, amarelo) – confirmar.
2. FVM `APRV` ≠ `APROVADO`: em que etapa aparece? Hoje fora dos critérios.
3. TAP `TAP2APROVACQ` e `TAP2APROVCQ` com a mesma descrição – equivalentes?
4. Parâmetro de TAGs sem FVI no denominador: SIM ou NÃO – qual a gerência usa?
5. Templates FVI/FVM/Preservação ainda não foram trazidos para o programa (continuam na planilha).
6. TAGS SEM FVI por SSOP soma também TAGs canceladas (como a planilha); hoje não há TAG cancelada, sem efeito – confirmar a regra.
7. Modelos da FIC recebidos em 29/10 (SOP/SSOP, FVI, FVM, CCM, TAP, TTAS) – leitores ajustados e testados com eles.
8. SOP / SSOP com status PRONTO (11 SSOP no modelo, exatamente as sem CCM): entram no denominador? Hoje entram.
9. FVM: resolvida – FVM BASE (hierarquia) + Lista de FVM (status em código).
12. Templates de FVI: confirmar a regra de família (ex.: (007-C) TRANSMISSOR DE PRESSÃO usa o template (007) TRANSMISSOR
    DE PROCESSO) e o FVI-011 (sem classificação no PDF) para a classe (011) PAINEL ELÉTRICO AUXILIAR.
10. 5 SSOP do modelo têm SOP fora da lista (incluídos sem descrição) e diluem o % de TAP/TTAS (14 de 19 SOP) – confirmar.
11. CCMOCULTO: 3 registros com SSOP fora da lista de SOP / SSOP – SSOP removidas ou faltando na lista?
