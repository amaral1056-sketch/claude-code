"""Gera a fonte de display "EPC12 Display" a partir das instâncias estáticas da Orbitron (OFL).

Motivo: na Orbitron o acento til combinante (tildecomb) é um traço inclinado – "VISÃO" lê-se "VISÀO". Aqui o contorno
do til é trocado pela onda do próprio "~" (asciitilde, mesma largura), centralizada na altura do acento original.
Como "Orbitron" é Nome de Fonte Reservado na OFL, a versão modificada recebe outro nome.

Uso: python ferramentas/fonte_display.py Orbitron-Bold.ttf Orbitron-Medium.ttf
     (grava EPC12Display-Bold.ttf / EPC12Display-Medium.ttf em comissionamento/recursos/fontes)"""
import sys
from pathlib import Path

from fontTools.ttLib import TTFont

FAMILIA = 'EPC12 Display'
DESTINO = Path(__file__).resolve().parents[1] / 'comissionamento' / 'recursos' / 'fontes'


def converter(origem):
    f = TTFont(origem)
    g = f['glyf']
    onda, til = g['asciitilde'], g['tildecomb']
    onda.recalcBounds(g)
    til.recalcBounds(g)
    dx = (til.xMin + til.xMax) // 2 - (onda.xMin + onda.xMax) // 2
    dy = (til.yMin + til.yMax) // 2 - (onda.yMin + onda.yMax) // 2
    novo = g['asciitilde'].__class__()
    novo.numberOfContours = onda.numberOfContours
    novo.coordinates = onda.getCoordinates(g)[0].copy()
    novo.coordinates.translate((dx, dy))
    novo.endPtsOfContours = list(onda.endPtsOfContours)
    novo.flags = onda.getCoordinates(g)[2].copy()
    novo.program = onda.program
    g['tildecomb'] = novo

    estilo = f['name'].getDebugName(17) or f['name'].getDebugName(2)   # Bold / Medium
    ps = f'EPC12Display-{estilo}'
    for rec in list(f['name'].names):
        if rec.nameID in (1, 16):
            rec.string = FAMILIA if estilo in ('Regular', 'Bold') or rec.nameID == 16 else f'{FAMILIA} {estilo}'
        elif rec.nameID == 3:
            rec.string = f'{rec.toUnicode().split(";")[0]};EPC12;{ps}'
        elif rec.nameID == 4:
            rec.string = f'{FAMILIA} {estilo}'
        elif rec.nameID == 6:
            rec.string = ps
    saida = DESTINO / f'{ps}.ttf'
    f.save(saida)
    return saida


if __name__ == '__main__':
    for arq in sys.argv[1:]:
        print(converter(arq))
