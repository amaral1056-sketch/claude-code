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
2. **Importar planilha ▸ SSOP**: a LI (aba LISTA) com a base de SOP e SSOP. As checagens de integridade da LI param a importação se algo não bater.
3. **Importar planilha ▸ FVI**: o Relatório de Cadastro Geral, o arquivo exportado da FIC, sem editar (título, cabeçalho, rodapé).
   O programa lê a data de emissão do rodapé e pede confirmação – é a data em que o avanço fica no **Histórico**.
4. **FVM, CCM, TAP, TTAS**: layout ainda provisório; as colunas são achadas pelo **nome do cabeçalho**
   (*Banco de dados ▸ Exportar planilhas-modelo* mostra o formato: `MALHA | SOP | SSOP | STATUS`, `Nº CCM | SSOP | STATUS`, …).

### Páginas
- **Visão Gerencial**: gráficos de acompanhamento de **FVI** (Campos II, III e V), **FVM** (Campos III e V), **TTAS**
  (1 e 2), **CCM** e **TAP**, por unidade e com o **Total** destacado. Passe o mouse nas colunas para ver valores e contas.
- **FVI**: cada TAG com a SSOP e a situação dos Campos II, III e V – "Aprovado", "Pendente (status da FIC)" ou
  "Não se aplica". Filtros: unidade, SOP, SSOP (em cascata) e TAG. **FVM**: o mesmo por malha (Campos III e V).
- **CCM** e **TAP** (por SSOP) e **TTAS** (por SOP, com Situação TTAS 1 e Situação TTAS 2): filtros por unidade, SOP,
  SSOP e situação.
- **Verificação** (testes e pendências linha a linha), **Histórico**, **Importações** e **Critérios** (editáveis; as
  listas oficiais de status da FIC são importadas nesta página).
- **Importar planilha**: SSOP (a LI), FVI, FVM, CCM, TAP e TTAS.
- **Exportar Excel** gera a mesma organização: Visão Gerencial (os 5 gráficos), FVI, FVM, CCM, TAP, TTAS, Verificação,
  Pendências, Histórico, Critérios e Dados dos gráficos (números que alimentam os gráficos).

**Logo:** a logo do Consórcio HC4 vem embutida no programa (barra lateral, ícone da janela e do .exe). Um arquivo
`logo.png` colocado na pasta do programa a substitui.

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
