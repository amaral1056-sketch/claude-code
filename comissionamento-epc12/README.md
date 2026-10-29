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
   lateral, embaixo) abre outro banco, faz cópia de segurança e exporta as planilhas-modelo.
2. **Importar base** (bases de dados, uma vez ou quando forem revisadas):
   - **LI de SOP / SSOP** (aba LISTA): base mestre. As checagens de integridade param a importação se algo não bater.
   - **LI de Itens** (TAG → SSOP) e **LI de Malhas** (malha → SSOP): layout provisório, colunas achadas pelo nome
     (`TAG`/`MALHA`, `SOP`, `SSOP`, `DESCRIÇÃO`). São conferidas com a FIC na página Verificação; não mudam o avanço.
3. **Importar atualização** (exportações da FIC): **FVI** (Relatório de Cadastro Geral, sem editar), **FVM**, **CCM**,
   **TAP** e **TTAS** (colunas achadas pelo nome; *Banco de dados ▸ Exportar planilhas-modelo* mostra o formato).
   O programa lê a data de emissão do rodapé e pede confirmação – é a data em que o avanço fica no **Histórico**.

### Páginas
- **Visão Gerencial**: uma pizza por documento, na ordem **FVI, FVM, CCM, TAP, TTAS**, todas do mesmo tamanho, com a
  distribuição dos **status da FIC** (geral ou por unidade). No FVI dá para alternar Status FVI / Campo II / III / V.
  Passe o mouse para ver a descrição oficial do status.
- **FVI**: cada TAG com a SSOP e o status da FIC dos Campos II, III e V. Filtros: unidade, SOP, SSOP (em cascata) e TAG.
- **FVM**: cada malha com a SSOP e o status da FIC. Uma malha **pode, mas não deve**, pertencer a mais de uma SSOP:
  ela aparece em cada SSOP (linha em amarelo), conta no avanço de cada uma e entra na Verificação.
- **CCM** e **TAP** (por SSOP) e **TTAS** (por SOP): status da FIC (com mais de um registro vale o aprovado, senão o
  último válido; "Sem registro" quando não há documento). No TTAS, que é um registro único que avança, a coluna
  Situação TTAS 1 mostra TTAS1APROVADO quando o SOP já está na etapa de TTAS 2. Filtros: unidade, SOP, SSOP e situação.
- **Verificação** (testes e pendências linha a linha, inclusive conferências das LIs de Itens e de Malhas),
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
ComissionamentoEPC12.exe --banco epc12.db --li LI.xlsx --status STATUS_*.xlsx --importar FVI Cadastro.xlsx --data 28/10/2026 --exportar relatorio.xlsx
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
