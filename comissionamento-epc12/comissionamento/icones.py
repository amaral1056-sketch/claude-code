"""Ícones de linha padronizados (mesmo traço, cantos arredondados), desenhados com Pillow em 4x e reduzidos
(bordas suaves), e a logo do programa (embutida em recursos/; um logo.png ao lado do programa a substitui)."""
from functools import lru_cache
from pathlib import Path

import customtkinter as ctk
from PIL import Image, ImageDraw

ESCALA = 4
_TRACO = 1.7  # px no tamanho final


def _cor(hexa):
    hexa = hexa.lstrip('#')
    return tuple(int(hexa[i:i + 2], 16) for i in (0, 2, 4)) + (255,)


class _Pincel:
    def __init__(self, tamanho, cor):
        self.s = tamanho * ESCALA
        self.k = self.s / 24.0           # desenho em grade 24x24
        self.w = max(1, round(_TRACO * ESCALA * tamanho / 20))
        self.cor = _cor(cor)
        self.img = Image.new('RGBA', (self.s, self.s), (0, 0, 0, 0))
        self.d = ImageDraw.Draw(self.img)

    def p(self, x, y):
        return (x * self.k, y * self.k)

    def linha(self, *pts):
        q = [self.p(*pt) for pt in pts]
        self.d.line(q, fill=self.cor, width=self.w, joint='curve')
        r = self.w / 2
        for x, y in (q[0], q[-1]):
            self.d.ellipse((x - r, y - r, x + r, y + r), fill=self.cor)

    def ret(self, x0, y0, x1, y1, raio=2.5):
        self.d.rounded_rectangle((*self.p(x0, y0), *self.p(x1, y1)), radius=raio * self.k, outline=self.cor, width=self.w)

    def circ(self, cx, cy, r, cheio=False):
        caixa = (*self.p(cx - r, cy - r), *self.p(cx + r, cy + r))
        if cheio:
            self.d.ellipse(caixa, fill=self.cor)
        else:
            self.d.ellipse(caixa, outline=self.cor, width=self.w)

    def arco(self, cx, cy, r, a0, a1):
        self.d.arc((*self.p(cx - r, cy - r), *self.p(cx + r, cy + r)), a0, a1, fill=self.cor, width=self.w)

    def elipse(self, cx, cy, rx, ry, a0=0, a1=360):
        caixa = (*self.p(cx - rx, cy - ry), *self.p(cx + rx, cy + ry))
        self.d.arc(caixa, a0, a1, fill=self.cor, width=self.w)

    def imagem(self, tamanho):
        return self.img.resize((tamanho, tamanho), Image.LANCZOS)


def _desenhar(nome, p):
    if nome == 'visao':            # painel: 4 blocos
        p.ret(3.5, 3.5, 10.5, 10.5, 2)
        p.ret(13.5, 3.5, 20.5, 10.5, 2)
        p.ret(3.5, 13.5, 10.5, 20.5, 2)
        p.ret(13.5, 13.5, 20.5, 20.5, 2)
    elif nome == 'fvi':            # prancheta com visto
        p.ret(5, 4.5, 19, 21)
        p.ret(9, 2.8, 15, 6.2, 1.2)
        p.linha((8.5, 13.5), (11, 16), (15.8, 11))
    elif nome == 'fvm':            # malha: nós ligados
        p.circ(6, 6, 2.4)
        p.circ(18, 6, 2.4)
        p.circ(12, 18, 2.4)
        p.linha((8.4, 6), (15.6, 6))
        p.linha((7.2, 8.1), (10.8, 15.9))
        p.linha((16.8, 8.1), (13.2, 15.9))
    elif nome == 'ccm':            # selo de conclusão
        p.circ(12, 9.5, 6)
        p.linha((9.5, 9.5), (11.3, 11.3), (14.6, 7.9))
        p.linha((8.6, 14.5), (7.5, 21), (12, 18.6), (16.5, 21), (15.4, 14.5))
    elif nome == 'tap':            # frasco de teste
        p.linha((9.5, 3), (14.5, 3))
        p.linha((10.3, 3), (10.3, 9.5), (4.8, 19), (5.8, 21), (18.2, 21), (19.2, 19), (13.7, 9.5), (13.7, 3))
        p.linha((7.5, 15), (16.5, 15))
    elif nome == 'ttas':           # bandeira (aceitação)
        p.linha((5.5, 21), (5.5, 3.5))
        p.linha((5.5, 4.5), (10, 3.5), (14, 5.5), (18.5, 4.5), (18.5, 13), (14, 14), (10, 12), (5.5, 13))
    elif nome == 'verificacao':    # escudo com visto
        p.linha((12, 2.8), (19.5, 5.8), (19.5, 11.5), (17.5, 16.5), (12, 21.2), (6.5, 16.5), (4.5, 11.5), (4.5, 5.8), (12, 2.8))
        p.linha((8.8, 12), (11.2, 14.4), (15.4, 9.6))
    elif nome == 'historico':      # relógio
        p.circ(12, 12, 8.5)
        p.linha((12, 7.5), (12, 12), (15.2, 14))
    elif nome == 'criterios':      # controles deslizantes
        for y, x in ((6, 15), (12, 8.5), (18, 13)):
            p.linha((3.5, y), (20.5, y))
            p.circ(x, y, 2.3, cheio=True)
    elif nome == 'importacoes':    # lista
        for y in (6.5, 12, 17.5):
            p.circ(5, y, 1.1, cheio=True)
            p.linha((9, y), (20, y))
    elif nome == 'importar':       # seta para dentro da bandeja
        p.linha((12, 3.5), (12, 14.5))
        p.linha((7.5, 10.5), (12, 15), (16.5, 10.5))
        p.linha((4, 15.5), (4, 20), (20, 20), (20, 15.5))
    elif nome == 'exportar':       # seta saindo da bandeja
        p.linha((12, 15), (12, 3.8))
        p.linha((7.5, 8), (12, 3.5), (16.5, 8))
        p.linha((4, 15.5), (4, 20), (20, 20), (20, 15.5))
    elif nome == 'busca':
        p.circ(10.5, 10.5, 6)
        p.linha((15, 15), (20, 20))
    elif nome == 'calendario':
        p.ret(3.5, 5, 20.5, 20.5)
        p.linha((3.5, 9.5), (20.5, 9.5))
        p.linha((8, 3), (8, 6.5))
        p.linha((16, 3), (16, 6.5))
    elif nome == 'banco':          # banco de dados
        p.elipse(12, 6, 7.5, 2.8)
        p.linha((4.5, 6), (4.5, 18))
        p.linha((19.5, 6), (19.5, 18))
        p.elipse(12, 18, 7.5, 2.8, 0, 180)
        p.elipse(12, 12, 7.5, 2.8, 0, 180)
    elif nome == 'ajuda':
        p.circ(12, 12, 8.5)
        p.arco(12, 9.6, 2.8, 180, 405)
        p.linha((12, 12.4), (12, 13.6))
        p.circ(12, 16.6, 0.9, cheio=True)
    elif nome == 'limpar':
        p.linha((7, 7), (17, 17))
        p.linha((17, 7), (7, 17))
    elif nome == 'mais':
        p.linha((12, 5), (12, 19))
        p.linha((5, 12), (19, 12))
    elif nome == 'lixeira':
        p.linha((4, 6.5), (20, 6.5))
        p.linha((9.5, 6.5), (9.5, 3.8), (14.5, 3.8), (14.5, 6.5))
        p.linha((6, 6.5), (7, 20.5), (17, 20.5), (18, 6.5))
    elif nome == 'restaurar':
        p.arco(12, 12, 8, 200, 520)
        p.linha((3.6, 4.5), (4.3, 9), (8.7, 8.2))
    else:
        raise KeyError(nome)


@lru_cache(maxsize=None)
def _pil(nome, cor, tamanho):
    p = _Pincel(tamanho, cor)
    _desenhar(nome, p)
    return p.imagem(tamanho)


def icone(nome, cor, tamanho=18):
    # só o desenho fica em cache: a CTkImage pertence à janela que a usa
    return ctk.CTkImage(light_image=_pil(nome, cor, tamanho), size=(tamanho, tamanho))


def recurso(nome):
    """Arquivo embutido no programa (comissionamento/recursos), inclusive dentro do .exe."""
    return Path(__file__).resolve().parent / 'recursos' / nome


def logo(pastas, altura=64):
    """CTkImage da logo: logo.png ao lado do programa (substitui) ou a logo embutida. None se não houver."""
    for arq in [Path(p) / 'logo.png' for p in pastas] + [recurso('logo.png')]:
        if arq.exists():
            try:
                img = Image.open(arq).convert('RGBA')
            except OSError:
                continue
            return ctk.CTkImage(light_image=img, size=(round(img.width * altura / img.height), altura))
    return None
