"""Gráficos desenhados em tk.Canvas (sem dependências): colunas agrupadas, barras de avanço, barras 100% e linhas.

Regras seguidas: um só eixo (0–100%), grade discreta, cores de série em ordem fixa (no máximo 3 por gráfico),
legenda sempre presente com 2+ séries, texto em tinta neutra, dica ao passar o mouse em cada marca."""
import tkinter as tk

from . import tema


def pct_txt(v):
    return '–' if v is None else f'{v * 100:.1f}%'.replace('.', ',')


class Grafico(tk.Canvas):
    """Área de gráfico com título, subtítulo e legenda; redesenha ao redimensionar."""

    def __init__(self, master, titulo, subtitulo='', altura=290):
        super().__init__(master, background=tema.SUPERFICIE, highlightthickness=0, height=int(altura * tema.escala))
        self.titulo, self.subtitulo = titulo, subtitulo
        self._desenhar = None
        self._dicas = {}
        self.bind('<Configure>', lambda e: self.redesenhar())
        self.bind('<Motion>', self._mover)
        self.bind('<Leave>', lambda e: self.delete('dica'))

    def definir(self, funcao):
        self._desenhar = funcao
        self.redesenhar()

    def redesenhar(self):
        self.delete('all')
        self._dicas = {}
        w, h = self.winfo_width(), self.winfo_height()
        if w < 60 or h < 60:
            return
        e = tema.escala
        self.create_text(4 * e, 4 * e, text=self.titulo, anchor='nw', font=tema.fonte(12, 'bold'), fill=tema.TEXTO)
        if self.subtitulo:
            self.create_text(4 * e, 28 * e, text=self.subtitulo, anchor='nw', font=tema.fonte(9), fill=tema.TEXTO_MUDO)
        if self._desenhar:
            self._desenhar(w, h, e)

    def vazio(self, w, h, msg='Sem dados – importe as planilhas para ver este gráfico.'):
        self.create_text(w / 2, h / 2 + 14, text=msg, font=tema.fonte(10), fill=tema.TEXTO_MUDO, width=w - 40)

    def legenda(self, series, e):
        x, y = 4 * e, 52 * e
        for nome, cor in series:
            self.create_oval(x, y + 2 * e, x + 10 * e, y + 12 * e, fill=cor, outline='')
            t = self.create_text(x + 16 * e, y + 7 * e, text=nome, anchor='w', font=tema.fonte(9), fill=tema.TEXTO_2)
            x = self.bbox(t)[2] + 18 * e

    def dica(self, item, texto):
        self._dicas[item] = texto

    def _mover(self, ev):
        self.delete('dica')
        alvo = None
        for item in reversed(self.find_overlapping(ev.x - 3, ev.y - 3, ev.x + 3, ev.y + 3)):
            if item in self._dicas:
                alvo = item
                break
        if alvo is None:
            return
        t = self.create_text(ev.x + 16, ev.y + 14, text=self._dicas[alvo], anchor='nw', font=tema.fonte(9),
                             fill='#FFFFFF', tags='dica')
        x0, y0, x1, y1 = self.bbox(t)
        w, h = self.winfo_width(), self.winfo_height()
        self.move(t, (w - 10 - x1) if x1 > w - 10 else 0, -(y1 - y0) - 34 if y1 > h - 8 else 0)
        x0, y0, x1, y1 = self.bbox(t)
        r = self.create_rectangle(x0 - 10, y0 - 7, x1 + 10, y1 + 7, fill='#2B3445', outline='', tags='dica')
        self.tag_raise(t, r)

    def eixo_pct(self, x0, y0, x1, y1, horizontal=False):
        for p in (0, .25, .5, .75, 1):
            if horizontal:
                x = x0 + (x1 - x0) * p
                self.create_line(x, y0, x, y1, fill=tema.EIXO if p == 0 else tema.GRADE)
                self.create_text(x, y1 + 6, text=f'{int(p * 100)}%', anchor='n', font=tema.fonte(8), fill=tema.TEXTO_MUDO)
            else:
                y = y1 - (y1 - y0) * p
                self.create_line(x0, y, x1, y, fill=tema.EIXO if p == 0 else tema.GRADE, dash=() if p == 0 else (2, 3))
                self.create_text(x0 - 8, y, text=f'{int(p * 100)}%', anchor='e', font=tema.fonte(8), fill=tema.TEXTO_MUDO)
        if horizontal:
            return lambda v: x0 + (x1 - x0) * max(0, min(1, v))
        return lambda v: y1 - (y1 - y0) * max(0, min(1, v))

    # --------------------------------------------------------------- tipos
    def colunas_agrupadas(self, categorias, series, rodape=None, destaque_ultimo=False):
        """series: [(nome, cor, [valor 0..1 | None por categoria])] – até 3 séries. Com uma série só não há legenda
        (o título identifica). destaque_ultimo: a última categoria (Total) vem separada por uma linha e em negrito."""
        def desenhar(w, h, e):
            if not categorias or all(v is None for _, _, vals in series for v in vals):
                return self.vazio(w, h)
            if len(series) > 1:
                self.legenda([(n, c) for n, c, _ in series], e)
            x0, y0, x1, y1 = 44 * e, (86 if len(series) > 1 else 62) * e, w - 6 * e, h - 42 * e
            ymap = self.eixo_pct(x0, y0, x1, y1)
            grupo = (x1 - x0) / len(categorias)
            larg = max(3, min(16 * e, (grupo - 12 * e) / len(series) - 3))
            for i, cat in enumerate(categorias):
                cx = x0 + grupo * i + grupo / 2
                inicio = cx - (larg + 3) * len(series) / 2
                for j, (nome, cor, vals) in enumerate(series):
                    v = vals[i]
                    if v is None:
                        continue
                    bx = inicio + j * (larg + 3)
                    topo = min(ymap(v), y1 - 2)
                    item = self.create_rectangle(bx, topo, bx + larg, y1, fill=cor, outline='')
                    alvo = self.create_rectangle(bx - 1, y0, bx + larg + 1, y1, fill='', outline='')
                    texto = f'{cat}\n{nome}: {pct_txt(v)}' + (f'\n{rodape(i, j)}' if rodape else '')
                    self.dica(item, texto)
                    self.dica(alvo, texto)
                ultimo = destaque_ultimo and i == len(categorias) - 1
                if ultimo:
                    xs = x0 + grupo * i
                    self.create_line(xs, y0, xs, y1 + 4 * e, fill=tema.EIXO, dash=(3, 3))
                desloca = 14 * e if (grupo < len(cat) * 7 * e + 6 and i % 2 and not ultimo) else 0
                self.create_text(cx, y1 + 7 * e + desloca, text=cat, anchor='n',
                                 font=tema.fonte(8, 'bold') if ultimo else tema.fonte(8),
                                 fill=tema.TEXTO if ultimo else tema.TEXTO_2)
        self.definir(desenhar)

    def linhas(self, rotulos, series):
        """series: [(nome, cor, [valor 0..1 | None por rótulo])] – evolução no tempo."""
        def desenhar(w, h, e):
            if not rotulos:
                return self.vazio(w, h, 'A evolução aparece a partir da primeira importação com data de referência.')
            self.legenda([(n, c) for n, c, _ in series], e)
            x0, y0, x1, y1 = 44 * e, 86 * e, w - 60 * e, h - 30 * e
            ymap = self.eixo_pct(x0, y0, x1, y1)
            passo = (x1 - x0) / max(1, len(rotulos) - 1)
            xs = [x0 + passo * i if len(rotulos) > 1 else (x0 + x1) / 2 for i in range(len(rotulos))]
            mostrar = max(1, len(rotulos) // 8)
            for i, r in enumerate(rotulos):
                if i % mostrar == 0 or i == len(rotulos) - 1:
                    self.create_text(xs[i], y1 + 7 * e, text=r, anchor='n', font=tema.fonte(8), fill=tema.TEXTO_MUDO)
            for nome, cor, vals in series:
                pts = [(xs[i], ymap(v)) for i, v in enumerate(vals) if v is not None]
                if len(pts) > 1:
                    self.create_line(*[c for p in pts for c in p], fill=cor, width=max(2, round(2 * e)))
                for i, v in enumerate(vals):
                    if v is None:
                        continue
                    x, y = xs[i], ymap(v)
                    r = 4.5 * e
                    item = self.create_oval(x - r, y - r, x + r, y + r, fill=cor, outline=tema.SUPERFICIE, width=2)
                    alvo = self.create_oval(x - 10, y - 10, x + 10, y + 10, fill='', outline='')
                    self.dica(item, f'{rotulos[i]}\n{nome}: {pct_txt(v)}')
                    self.dica(alvo, f'{rotulos[i]}\n{nome}: {pct_txt(v)}')
                if pts:
                    ult = next(v for v in reversed(vals) if v is not None)
                    self.create_text(pts[-1][0] + 10 * e, pts[-1][1], text=pct_txt(ult), anchor='w',
                                     font=tema.fonte(9, 'bold'), fill=tema.TEXTO_2)
        self.definir(desenhar)
