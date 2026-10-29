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
   ao lado do .exe (se a pasta não aceitar gravação: `%LOCALAPPDATA%\ComissionamentoEPC12`). *Arquivo > Abrir ou criar banco* troca o arquivo.
2. **Importar > LI – base SOP/SSOP** (aba LISTA). Checagens de integridade da LI param a importação se algo não bater.
3. **Importar > Relatório de Cadastro Geral (FVI)**: o arquivo exportado da FIC, sem editar (título, cabeçalho, rodapé).
   O programa lê a data de emissão do rodapé e pede confirmação – é a data em que o avanço fica no **Histórico**.
4. **FVM, CCM, TAP, TTAS**: layout ainda provisório; as colunas são achadas pelo **nome do cabeçalho**
   (*Arquivo > Exportar planilhas-modelo* mostra o formato: `MALHA | SOP | SSOP | STATUS`, `Nº CCM | SSOP | STATUS`, …).
5. Confira a aba **Verificação** (testes devem dar OK; pendências linha a linha) e use **Arquivo > Exportar relatório Excel**.

Cada importação substitui a anterior do mesmo tipo (a FIC exporta sempre a lista completa) e fica registrada na aba
**Importações**. Critérios de "aprovado" e o parâmetro *TAGs sem FVI no denominador* são editáveis na aba **Critérios**.
Faça cópias do banco em *Arquivo > Salvar cópia de segurança*.

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
