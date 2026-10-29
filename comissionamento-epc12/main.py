"""Ponto de entrada do Controle Gerencial de Comissionamento – EPC 12.

Sem argumentos abre a interface gráfica. Também funciona por linha de comando (automação/conferência):
  ComissionamentoEPC12.exe --banco C:\\pasta\\epc12.db --li LI.xlsx --importar FVI Cadastro.xlsx --exportar relatorio.xlsx
  python main.py --importar FVM fvm.xlsx --data 04/11/2026 --exportar rel.xlsx"""
import argparse
import sys

from comissionamento import __version__, configuracao, exportacao, servico
from comissionamento.banco import Banco
from comissionamento.importacao import ErroImportacao


def linha_de_comando(a):
    b = Banco(a.banco or configuracao.banco_padrao())
    try:
        if a.li:
            print(servico.importar_li(b, a.li))
        for arq in a.status or []:
            print('status %s: %d códigos' % servico.importar_status(b, arq))
        for tipo, arq in a.importar or []:
            regs, data, _ = servico.importar_fic(b, tipo.upper(), arq, a.data)
            print(f'{tipo.upper()}: {len(regs)} registros, data de referência {servico.data_br(data)}')
        R = servico.calcular(b)
        t = R.total
        p = lambda v: '-' if v is None else f'{v * 100:.1f}%'
        print(f'TOTAL: {t["tags"]} TAGs · FVI CII {p(t["pct_cii"])} · CIII {p(t["pct_ciii"])} · CV {p(t["pct_cv"])} · '
              f'FVM CIII {p(t["pct_fvm_ciii"])} · CCM {p(t["pct_ccm"])} · TAP {p(t["pct_tap"])} · TTAS1 {p(t["pct_ttas1"])}')
        for _, teste, _, apurado, res in R.verificacao:
            if res == 'VERIFICAR':
                print(f'VERIFICAR: {teste} = {apurado}')
        if a.exportar:
            exportacao.exportar_relatorio(R, a.exportar, servico.data_br(b.data_referencia_atual()),
                                          servico.historico_com_percentuais(b), b.criterios(), b.incluir_sem_fvi)
            print('relatório:', a.exportar)
    except ErroImportacao as e:
        print('ERRO:', e, file=sys.stderr)
        return 2
    finally:
        b.fechar()
    return 0


def main():
    ap = argparse.ArgumentParser(description='Controle Gerencial de Comissionamento – EPC 12')
    ap.add_argument('--banco', help='arquivo do banco de dados (.db)')
    ap.add_argument('--li', help='importar a LI (base SOP/SSOP)')
    ap.add_argument('--status', nargs='+', help='importar listas oficiais STATUS_*.xlsx')
    ap.add_argument('--importar', nargs=2, action='append', metavar=('TIPO', 'ARQUIVO'),
                    help='importar exportação da FIC: TIPO = FVI, FVM, CCM, TAP ou TTAS')
    ap.add_argument('--data', help='data de referência dd/mm/aaaa (padrão: rodapé do arquivo ou hoje)')
    ap.add_argument('--exportar', help='gerar o relatório Excel neste arquivo')
    ap.add_argument('--versao', action='version', version=__version__)
    a = ap.parse_args()
    if a.li or a.status or a.importar or a.exportar:
        sys.exit(linha_de_comando(a))
    from comissionamento import gui
    gui.executar(a.banco)


if __name__ == '__main__':
    main()
