# Controle Gerencial de Comissionamento – EPC 12

Programa para Windows (executável único, **não precisa instalar**) que mantém um **banco de dados** do avanço do
comissionamento e é alimentado pelas **planilhas Excel** exportadas da FIC. Continua o trabalho da planilha
`CONTROLE_GERENCIAL_EPC12.xlsx`: mesmas regras de negócio, mesmos números (conferido célula a célula).

## Como obter o executável
- **GitHub Actions** → workflow *Comissionamento EPC12 – executável Windows* → artefato `ComissionamentoEPC12-windows`
  (contém `ComissionamentoEPC12.exe`). O workflow roda os testes, gera o .exe e testa o .exe no Windows antes de publicar.
- Ou, num computador com Python 3.10+: `build_windows.bat` (gera `dist\ComissionamentoEPC12.exe`).

## Uso
1. Copie `ComissionamentoEPC12.exe` para uma pasta (pode ser de rede) e abra. O banco é criado em `dados\comissionamento_epc12.db`
   ao lado do .exe (se a pasta não aceitar gravação: `%LOCALAPPDATA%\ComissionamentoEPC12`). **Banco de dados** (barra
   lateral, embaixo) abre outro banco ou faz cópia de segurança.
2. **Importar atualização** – os arquivos exportados da FIC, sem editar:
   | Opção | Arquivo da FIC | Observação |
   |---|---|---|
   | **SOP / SSOP** | Lista de Cadastros Básicos (Código, Descrição, Tipo) | base mestre; importe primeiro. Unidade e SOP de cada SSOP vêm do código |
   | **FVI** | Relatório de Cadastro Geral | por TAG |
   | **FVM** | Relatório de Hierarquia de Malhas | a SSOP da malha vem das suas TAGs/linhas; o status é o da malha |
   | **CCM** | Lista de CCM | por SSOP (inclui registros sem número, ex.: CCMOCULTO) |
   | **TAP** | Lista de TAP | por SOP (o número do TAP é "TAP-" + SOP) |
   | **TTAS** | Lista de TTAS | por SOP |

   O programa lê a data de emissão do rodapé (quando existe) e pede confirmação – é a data em que o avanço fica no
   **Histórico**. A LI antiga (aba LISTA) ainda é aceita em SOP / SSOP.

### Páginas
- **Visão Gerencial**: uma pizza por documento, na ordem **FVI, FVM, CCM, TAP, TTAS**, todas do mesmo tamanho, com a
  distribuição dos **status da FIC** (geral ou por unidade). No FVI dá para alternar Status FVI / Campo II / III / V.
  Passe o mouse para ver a descrição oficial do status.
- **FVI**: cada TAG com a SSOP e o status da FIC dos Campos II, III e V. Filtros: unidade, SOP, SSOP (em cascata) e TAG.
- **FVM**: cada malha com a SSOP e o status da malha na FIC. Uma malha **pode, mas não deve**, pertencer a mais de uma SSOP:
  ela aparece em cada SSOP (linha em amarelo), conta no avanço de cada uma e entra na Verificação.
- **CCM** (por SSOP), **TAP** e **TTAS** (por SOP): status da FIC (com mais de um registro vale o aprovado, senão o
  último válido; "Sem registro" quando não há documento). No TTAS, que é um registro único que avança, a coluna
  Situação TTAS 1 mostra TTAS1APROVADO quando o SOP já está na etapa de TTAS 2. Filtros: unidade, SOP, SSOP e situação.
- **Verificação** (testes e pendências linha a linha, inclusive os avisos da lista de SOP / SSOP),
  **Histórico**, **Importações** e **Critérios** (o que conta como aprovado no avanço; listas oficiais de status).
- **Exportar Excel**: Visão Gerencial (as 5 pizzas), FVI, FVM, CCM, TAP, TTAS, Verificação, Pendências, Histórico,
  Critérios e Dados dos gráficos (contagem de status e avanço por unidade).

**Logo:** a logo do Consórcio HC4 vem embutida no programa (barra lateral, ícone da janela e do .exe). Um arquivo
`logo.png` colocado na pasta do programa a substitui.

**Visual:** identidade HUD monocromática (base branca, preto e cinzas, painéis retos com cantoneiras, textos em caixa-alta).
As fontes vêm embutidas e só são usadas pelo programa (nada é instalado no Windows): EPC12 Display (derivada da Orbitron,
com o til corrigido para ler "VISÃO" e não "VISÀO"), Rajdhani, IBM Plex Mono e Share Tech Mono – todas sob licença SIL OFL
(textos em `comissionamento/recursos/fontes`). Nas pizzas, as quatro primeiras categorias são cinzas sólidos e as demais
usam texturas (hachura, pontos), iguais na tela e no Excel. O Excel usa Bahnschrift e Consolas, que já vêm no Windows.

Cada importação substitui a anterior do mesmo tipo (a FIC exporta sempre a lista completa) e fica registrada em **Importações**.

### Linha de comando (automação)
```
ComissionamentoEPC12.exe --banco epc12.db --li SOP_E_SSOP.xlsx --status STATUS_*.xlsx --importar FVI FVI.xlsx --importar FVM FVM.xlsx --importar TAP TAP.xlsx --exportar relatorio.xlsx
```

## Desenvolvimento
```
pip install -r requirements-dev.txt
python main.py                 # interface
python -m pytest testes -q     # testes (sintéticos + interface; dados reais se existirem em dados/)
```
`dados/` é ignorado pelo git: os arquivos da FIC são classificados (NP-1) e **não** devem ser versionados.
Para validar contra a planilha entregue, coloque em `dados/` a LI, o `FIC_Cadastro_Geral.xlsx`, o
`CONTROLE_GERENCIAL_EPC12.xlsx` e `status_fic/` – `testes/test_dados_reais.py` compara SSOP, SOP, unidades e verificação.
Detalhes técnicos e regras em `CLAUDE.md`.
