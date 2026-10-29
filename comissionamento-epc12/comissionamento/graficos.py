"""Gráficos desenhados em tk.Canvas (sem dependências): colunas agrupadas, barras 100% empilhadas, linhas e medidor.

Regras seguidas: um só eixo (0–100%), grade discreta, cores de série em ordem fixa, legenda sempre presente com
2+ séries, texto em tinta neutra (nunca na cor da série), dica ao passar o mouse em cada marca."""
import tkinter as tk

from . import tema


def _pct(v):
    return '–' if v is None else f'{v * 100:.1f}%'.replace('.', ',')


class Grafico(tk.Canvas):
    """Cartão com título, subtítulo, legenda e área de plotagem; redesenha ao redimensionar."""

    def __init__(self, master, titulo, subtitulo='', altura=200):
        super().__init__(master, background=tema.SUPERFICIE, highlightthickness=1, highlightbackground=tema.BORDA,
                         height=altura)
        self.titulo, self.subtitulo = titulo, subtitulo
        self._desenhar = None
        self._dicas = {}
        self.bind('<Configure>', lambda e: self.redesenhar())
        self.bind('<Motion>', self._mover)
        self.bind('<Leave>', lambda e: self.delete('dica'))

    # --------------------------------------------------------------- estrutura comum
    def definir(self, funcao):
        self._desenhar = funcao
        self.redesenhar()

    def redesenhar(self):
        self.delete('all')
        self._dicas = {}
        w, h = self.winfo_width(), self.winfo_height()
        if w < 50 or h < 50:
            return
        self.create_text(16, 16, text=self.titulo, anchor='nw', font=tema.fonte(11, 'bold'), fill=tema.TEXTO)
        if self.subtitulo:
            self.create_text(16, 38, text=self.subtitulo, anchor='nw', font=tema.fonte(8), fill=tema.TEXTO_MUDO)
        if self._desenhar:
            self._desenhar(w, h)

    def vazio(self, w, h, msg='Sem dados – importe as planilhas para ver este gráfico.'):
        self.create_text(w / 2, h / 2 + 10, text=msg, font=tema.fonte(9), fill=tema.TEXTO_MUDO, width=w - 40)

    def legenda(self, w, series, y=56):
        x = 16
        for nome, cor in series:
            self.create_rectangle(x, y + 1, x + 10, y + 11, fill=cor, outline='')
            t = self.create_text(x + 15, y + 6, text=nome, anchor='w', font=tema.fonte(8), fill=tema.TEXTO_2)
            x = self.bbox(t)[2] + 16

    def dica(self, item, texto):
        self._dicas[item] = texto

    def _mover(self, e):
        self.delete('dica')
        alvo = None
        for item in reversed(self.find_overlapping(e.x - 3, e.y - 3, e.x + 3, e.y + 3)):
            if item in self._dicas:
                alvo = item
                break
        if alvo is None:
            return
        texto = self._dicas[alvo]
        t = self.create_text(e.x + 14, e.y + 12, text=texto, anchor='nw', font=tema.fonte(9), fill=tema.TEXTO, tags='dica')
        x0, y0, x1, y1 = self.bbox(t)
        w = self.winfo_width()
        dx = (w - 6 - x1) if x1 > w - 6 else 0
        dy = -(y1 - y0) - 30 if y1 > self.winfo_height() - 6 else 0
        self.move(t, dx, dy)
        x0, y0, x1, y1 = self.bbox(t)
        r = self.create_rectangle(x0 - 8, y0 - 6, x1 + 8, y1 + 6, fill=tema.SUPERFICIE, outline=tema.EIXO, tags='dica')
        self.tag_raise(t, r)

    def eixo_pct(self, x0, y0, x1, y1, horizontal=False):
        """Grade 0–100% (y0 topo, y1 base). Devolve função valor->coordenada."""
        for p in (0, .25, .5, .75, 1):
            if horizontal:
                x = x0 + (x1 - x0) * p
                self.create_line(x, y0, x, y1, fill=tema.EIXO if p == 0 else tema.GRADE)
                self.create_text(x, y1 + 6, text=f'{int(p * 100)}%', anchor='n', font=tema.fonte(8), fill=tema.TEXTO_MUDO)
            else:
                y = y1 - (y1 - y0) * p
                self.create_line(x0, y, x1, y, fill=tema.EIXO if p == 0 else tema.GRADE)
                self.create_text(x0 - 6, y, text=f'{int(p * 100)}%', anchor='e', font=tema.fonte(8), fill=tema.TEXTO_MUDO)
        if horizontal:
            return lambda v: x0 + (x1 - x0) * max(0, min(1, v))
        return lambda v: y1 - (y1 - y0) * max(0, min(1, v))

    # --------------------------------------------------------------- tipos de gráfico
    def colunas_agrupadas(self, categorias, series, rodape=None):
        """series: [(nome, cor, [valor 0..1 | None por categoria])]."""
        def desenhar(w, h):
            if not categorias or all(v is None for _, _, vals in series for v in vals):
                return self.vazio(w, h)
            self.legenda(w, [(n, c) for n, c, _ in series])
            x0, y0, x1, y1 = 52, 86, w - 16, h - 44
            ymap = self.eixo_pct(x0, y0, x1, y1)
            grupo = (x1 - x0) / len(categorias)
            larg = max(3, min(18, (grupo - 10) / len(series) - 2))
            for i, cat in enumerate(categorias):
                cx = x0 + grupo * i + grupo / 2
                inicio = cx - (larg + 2) * len(series) / 2
                for j, (nome, cor, vals) in enumerate(series):
                    v = vals[i]
                    bx = inicio + j * (larg + 2)
                    if v is None:
                        continue
                    topo = min(ymap(v), y1 - 1.5)  # 0% ainda aparece como traço
                    item = self.create_rectangle(bx, topo, bx + larg, y1, fill=cor, outline='')
                    alvo = self.create_rectangle(bx - 1, y0, bx + larg + 1, y1, fill='', outline='')
                    texto = f'{cat}\n{nome}: {_pct(v)}' + (f'\n{rodape(i, j)}' if rodape else '')
                    self.dica(item, texto)
                    self.dica(alvo, texto)
                desloca = 14 if (grupo < len(cat) * 6.5 + 6 and i % 2) else 0
                self.create_text(cx, y1 + 6 + desloca, text=cat, anchor='n', font=tema.fonte(8), fill=tema.TEXTO_2)
        self.definir(desenhar)

    def barras_empilhadas(self, categorias, segmentos):
        """Barras horizontais 100%. segmentos: [(nome, cor, [contagem por categoria])]."""
        def desenhar(w, h):
            totais = [sum(vals[i] for _, _, vals in segmentos) for i in range(len(categorias))]
            if not categorias or not any(totais):
                return self.vazio(w, h)
            self.legenda(w, [(n, c) for n, c, _ in segmentos])
            x0, y0, x1, y1 = 96, 84, w - 20, h - 30
            xmap = self.eixo_pct(x0, y0, x1, y1, horizontal=True)
            faixa = (y1 - y0) / len(categorias)
            alt = min(34, faixa - 14)
            for i, cat in enumerate(categorias):
                cy = y0 + faixa * i + faixa / 2
                self.create_text(x0 - 10, cy, text=cat, anchor='e', font=tema.fonte(9), fill=tema.TEXTO_2)
                acum = 0
                for nome, cor, vals in segmentos:
                    n = vals[i]
                    if not totais[i] or not n:
                        continue
                    a, b = xmap(acum / totais[i]), xmap((acum + n) / totais[i])
                    item = self.create_rectangle(a, cy - alt / 2, max(a + 1, b - 2), cy + alt / 2, fill=cor, outline='')
                    p = n / totais[i]
                    self.dica(item, f'{cat} – {nome}\n{n} TAG(s) · {_pct(p)}')
                    if b - a > 46:
                        claro = cor != tema.NEUTRO
                        self.create_text((a + b) / 2, cy, text=_pct(p), font=tema.fonte(8, 'bold'),
                                         fill='#FFFFFF' if claro else tema.TEXTO)
                    acum += n
        self.definir(desenhar)

    def linhas(self, rotulos, series):
        """series: [(nome, cor, [valor 0..1 | None por rótulo])] – evolução no tempo."""
        def desenhar(w, h):
            if not rotulos:
                return self.vazio(w, h, 'O histórico aparece a partir da primeira importação com data de referência.')
            self.legenda(w, [(n, c) for n, c, _ in series])
            x0, y0, x1, y1 = 52, 86, w - 70, h - 38
            ymap = self.eixo_pct(x0, y0, x1, y1)
            passo = (x1 - x0) / max(1, len(rotulos) - 1)
            xs = [x0 + passo * i if len(rotulos) > 1 else (x0 + x1) / 2 for i in range(len(rotulos))]
            mostrar = max(1, len(rotulos) // 8)
            for i, r in enumerate(rotulos):
                if i % mostrar == 0 or i == len(rotulos) - 1:
                    self.create_text(xs[i], y1 + 6, text=r, anchor='n', font=tema.fonte(8), fill=tema.TEXTO_MUDO)
            for nome, cor, vals in series:
                pts = [(xs[i], ymap(v)) for i, v in enumerate(vals) if v is not None]
                if len(pts) > 1:
                    self.create_line(*[c for p in pts for c in p], fill=cor, width=2, smooth=False)
                for i, v in enumerate(vals):
                    if v is None:
                        continue
                    x, y = xs[i], ymap(v)
                    item = self.create_oval(x - 4, y - 4, x + 4, y + 4, fill=cor, outline=tema.SUPERFICIE, width=2)
                    alvo = self.create_oval(x - 9, y - 9, x + 9, y + 9, fill='', outline='')
                    self.dica(item, f'{rotulos[i]}\n{nome}: {_pct(v)}')
                    self.dica(alvo, f'{rotulos[i]}\n{nome}: {_pct(v)}')
                if pts:  # rótulo direto no último ponto
                    ult = next(v for v in reversed(vals) if v is not None)
                    self.create_text(pts[-1][0] + 10, pts[-1][1], text=_pct(ult), anchor='w', font=tema.fonte(8, 'bold'),
                                     fill=tema.TEXTO_2)
        self.definir(desenhar)


class Medidor(tk.Canvas):
    """Barra de progresso fina (trilha + preenchimento) para os cartões de indicador."""

    def __init__(self, master, cor=tema.DESTAQUE):
        super().__init__(master, height=8, background=tema.SUPERFICIE, highlightthickness=0)
        self.cor, self.valor = cor, None
        self.bind('<Configure>', lambda e: self._desenhar())

    def definir(self, v):
        self.valor = v
        self._desenhar()

    def _desenhar(self):
        self.delete('all')
        w = self.winfo_width()
        self.create_rectangle(0, 1, w, 7, fill=tema.TRILHA, outline='')
        if self.valor:
            self.create_rectangle(0, 1, max(4, w * min(1, self.valor)), 7, fill=self.cor, outline='')
