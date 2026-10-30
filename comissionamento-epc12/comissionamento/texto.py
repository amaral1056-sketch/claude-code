"""Normalização de texto equivalente às fórmulas da planilha original.

A planilha usava TRIM() e comparações do Excel (COUNTIFS, SUMIFS, "="), que:
- removem espaços nas pontas e reduzem espaços repetidos internos a um só (apenas o caractere espaço);
- não diferenciam maiúsculas de minúsculas.
O programa reproduz essas regras para chegar aos mesmos números da planilha."""
import datetime
import re

_ESPACOS = re.compile(' {2,}')


def texto(v):
    """Valor de célula como o Excel o converteria para texto (None -> '')."""
    if v is None:
        return ''
    if isinstance(v, bool):
        return 'VERDADEIRO' if v else 'FALSO'
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime('%d/%m/%Y')
    return str(v)


def trim(v):
    """TRIM() do Excel."""
    return _ESPACOS.sub(' ', texto(v).strip(' '))


def chave(v):
    """Chave de comparação: TRIM + maiúsculas (comparação do Excel não diferencia caixa)."""
    return trim(v).upper()
