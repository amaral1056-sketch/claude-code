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
2. **Importar planilha ▸ LI – base de SOP e SSOP** (aba LISTA). As checagens de integridade da LI param a importação se algo não bater.
3. **Importar planilha ▸ Relatório de Cadastro Geral (FVI)**: o arquivo exportado da FIC, sem editar (título, cabeçalho, rodapé).
   O programa lê a data de emissão do rodapé e pede confirmação – é a data em que o avanço fica no **Histórico**.
4. **FVM, CCM, TAP, TTAS**: layout ainda provisório; as colunas são achadas pelo **nome do cabeçalho**
   (*Banco de dados ▸ Exportar planilhas-modelo* mostra o formato: `MALHA | SOP | SSOP | STATUS`, `Nº CCM | SSOP | STATUS`, …).

### Páginas
- **Visão Gerencial**: os 9 indicadores por extenso (FVI Campo II, FVI Campo III, FVI Campo V, FVM Campo III, FVM Campo V,
  CCM, TAP, TTAS 1, TTAS 2) em cartões com anel de progresso; gráficos de avanço geral, FVI por unidade, FVM por unidade,
  CCM e TAP por unidade, TTAS por unidade e evolução no histórico (passe o mouse para ver os valores); tabela por unidade.
- **FVI**: cada TAG com a SSOP e a situação de cada campo – "Aprovado", "Pendente (status da FIC)" ou "Não se aplica".
  Filtros por unidade, situação (com pendência, sem FVI gerada, canceladas, sem SSOP…) e busca por TAG, SSOP ou descrição.
- **FVM**: o mesmo, malha a malha (Campos III e V). **CCM**, **TAP** (por SSOP) e **TTAS** (por SOP).
- **Verificação** (testes e pendências linha a linha), **Histórico**, **Importações** e **Critérios** (editáveis).
- **Exportar Excel** gera a mesma organização: Visão Gerencial (indicadores, tabela e gráficos), FVI, FVM, CCM, TAP, TTAS,
  Verificação, Pendências, Histórico e Critérios.

**Logo:** coloque um arquivo `logo.png` na mesma pasta do programa para ele aparecer no lugar do quadrado "LOGO".

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
