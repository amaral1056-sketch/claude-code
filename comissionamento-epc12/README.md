# Controle Gerencial de Comissionamento – EPC 12

Programa para Windows (executável único, **não precisa instalar**) que mantém um **banco de dados** do avanço do
comissionamento e é alimentado pelas **planilhas Excel** exportadas da FIC. Continua o trabalho da planilha
`CONTROLE_GERENCIAL_EPC12.xlsx`: mesmas regras de negócio, mesmos números (conferido célula a célula).

## Como obter o executável
- **GitHub Actions** → workflow *Comissionamento EPC12 – executável Windows* → artefato `ComissionamentoEPC12-windows`
  (contém `ComissionamentoEPC12.exe`). O workflow roda os testes, gera o .exe e testa o .exe no Windows antes de publicar.
- Ou, num computador com Python 3.10+: `build_windows.bat` (gera `dist\ComissionamentoEPC12.exe`).
- **Aviso "O Windows protegeu o computador" (SmartScreen, fornecedor desconhecido):** o .exe ainda não tem assinatura
  digital. Antes de extrair o .zip baixado: botão direito ▸ Propriedades ▸ marque **Desbloquear** ▸ OK (ou, no
  PowerShell, `Unblock-File .\ComissionamentoEPC12.exe`). Para eliminar o aviso de vez é preciso assinar o .exe com um
  certificado de assinatura de código da empresa (decisão pendente).

## Uso
1. Copie `ComissionamentoEPC12.exe` para uma pasta (pode ser de rede) e abra. O banco é criado em `dados\comissionamento_epc12.db`
   ao lado do .exe (se a pasta não aceitar gravação: `%LOCALAPPDATA%\ComissionamentoEPC12`). **Banco de dados** (barra
   lateral, embaixo) abre outro banco ou faz cópia de segurança.
2. **Importar atualização** – os arquivos exportados da FIC, sem editar:
   | Opção | Arquivo da FIC | Observação |
   |---|---|---|
   | **SOP / SSOP** | Lista de Cadastros Básicos (Código, Descrição, Tipo) | base mestre; importe primeiro. Unidade e SOP de cada SSOP vêm do código |
   | **FVI** | Relatório de Cadastro Geral | por TAG |
   | **FVM BASE** | Relatório de Hierarquia de Malhas | base das FVMs: as malhas, suas TAGs e a SSOP de cada uma; importe antes da FVM |
   | **FVM** | Lista de FVM | atualização: Nº, template e status da FVM de cada malha da base |
   | **Templates de FVI / FVM (PDF)** | Template de FVI – Folha de Verificação de Item e/ou Template de FVM – Folha de Verificação de Malhas (um PDF de cada vez) | as perguntas de cada campo; nas páginas FVI e FVM, clique numa TAG ou malha para vê-las. Na FVI o template é achado pela classe da TAG |
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
- **FVM**: uma linha por TAG da malha – Malha, TAG, SSOP e Status da FVM. Busca por malha ou TAG. Uma malha **pode, mas não deve**, pertencer a mais de uma SSOP:
  ela aparece em cada SSOP (linha em amarelo), conta no avanço de cada uma e entra na Verificação.
- **CCM** (por SSOP), **TAP** e **TTAS** (por SOP): status da FIC (com mais de um registro vale o aprovado, senão o
  último válido; "Sem registro" quando não há documento). No TTAS, que é um registro único que avança, a coluna
  Situação TTAS 1 mostra TTAS1APROVADO quando o SOP já está na etapa de TTAS 2. Filtros: unidade, SOP, SSOP e situação.
- **Verificação** (testes e pendências linha a linha, inclusive os avisos da lista de SOP / SSOP),
  **Histórico**, **Importações** e **Critérios** (o que conta como aprovado no avanço; listas oficiais de status).
- **Exportar Excel e PDF**: gera os dois arquivos com o mesmo nome, com a logo, só com Visão Gerencial (as 5 pizzas e
  os números delas), FVI, FVM, CCM, TAP e TTAS. O PDF é A4 paisagem, pronto para envio.

**Logo:** a logo do Consórcio HC4 vem embutida no programa (barra lateral, ícone da janela e do .exe). Um arquivo
`logo.png` colocado na pasta do programa a substitui.

**Visual:** identidade HUD (base branca, preto e cinzas, painéis retos com cantoneiras, textos em caixa-alta) e gráficos
coloridos com uma paleta testada para daltonismo – o mesmo status tem sempre a mesma cor, na tela, no Excel e no PDF.
As fontes vêm embutidas e só são usadas pelo programa (nada é instalado no Windows): EPC12 Display (derivada da Orbitron,
com o til corrigido para ler "VISÃO" e não "VISÀO"), Rajdhani, IBM Plex Mono e Share Tech Mono – todas sob licença SIL OFL
(textos em `comissionamento/recursos/fontes`), embutidas também no PDF. O Excel usa Bahnschrift e Consolas, que já vêm
no Windows.

Cada importação substitui a anterior do mesmo tipo (a FIC exporta sempre a lista completa) e fica registrada em **Importações**.

### Linha de comando (automação)
```
ComissionamentoEPC12.exe --banco epc12.db --li SOP_E_SSOP.xlsx --fvm-base FVM_BASE.xlsx --status STATUS_*.xlsx --importar FVI FVI.xlsx --importar FVM FVM.xlsx --importar TAP TAP.xlsx --exportar relatorio.xlsx
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
