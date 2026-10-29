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


class Pizza(tk.Canvas):
    """Gráfico de pizza (Pillow em 4x, bordas suaves) com legenda ao lado: status, quantidade e %.
    Passar o mouse numa fatia ou na legenda mostra a descrição do status."""

    def __init__(self, master, altura=300):
        super().__init__(master, background=tema.SUPERFICIE, highlightthickness=0, height=int(altura * tema.escala))
        self.fatias, self._img, self._geo = [], None, None
        self._dicas = {}
        self.bind('<Configure>', lambda e: self.redesenhar())
        self.bind('<Motion>', self._mover)
        self.bind('<Leave>', lambda e: self.delete('dica'))

    def definir(self, fatias):
        """fatias: [(rótulo, quantidade, cor, descrição)] – na ordem em que serão desenhadas."""
        self.fatias = [f for f in fatias if f[1] > 0]
        self.redesenhar()

    def redesenhar(self):
        from PIL import Image, ImageDraw, ImageTk
        self.delete('all')
        self._dicas, self._geo = {}, None
        w, h = self.winfo_width(), self.winfo_height()
        if w < 80 or h < 80:
            return
        e = tema.escala
        total = sum(n for _, n, _, _ in self.fatias)
        if not total:
            self.create_text(w / 2, h / 2, text='Sem dados – importe as planilhas para ver este gráfico.',
                             font=tema.fonte(10), fill=tema.TEXTO_MUDO, width=w - 40)
            return
        d = int(min(h - 16 * e, w * 0.46))
        cx, cy = 8 * e + d / 2, h / 2
        k = 4
        img = Image.new('RGBA', (d * k, d * k), (0, 0, 0, 0))
        dr = ImageDraw.Draw(img)
        ang, cortes = -90.0, []
        for rot, n, cor, _ in self.fatias:
            fim = ang + 360.0 * n / total
            dr.pieslice((0, 0, d * k - 1, d * k - 1), ang, fim, fill=cor,
                        outline=tema.SUPERFICIE if len(self.fatias) > 1 else None, width=2 * k)
            cortes.append((ang, fim))
            ang = fim
        self._img = ImageTk.PhotoImage(img.resize((d, d), Image.LANCZOS), master=self)
        self.create_image(cx, cy, image=self._img)
        self._geo = (cx, cy, d / 2, cortes)
        import math
        for (a0, a1), (rot, n, cor, _) in zip(cortes, self.fatias):
            if (a1 - a0) / 360 >= 0.06:
                meio = math.radians((a0 + a1) / 2)
                x, y = cx + d * 0.31 * math.cos(meio), cy + d * 0.31 * math.sin(meio)
                escuro = cor.lower() in ('#eda100', '#d9dee6', '#c3c9d2')
                self.create_text(x, y, text=pct_txt(n / total), font=tema.fonte(9, 'bold'),
                                 fill=tema.TEXTO if escuro else '#FFFFFF')
        # legenda
        x0 = cx + d / 2 + 26 * e
        passo = min(28 * e, (h - 20 * e) / max(1, len(self.fatias)))
        y = cy - passo * len(self.fatias) / 2 + passo / 2
        for rot, n, cor, desc in self.fatias:
            self.create_oval(x0, y - 5 * e, x0 + 10 * e, y + 5 * e, fill=cor, outline='')
            t = self.create_text(x0 + 18 * e, y, text=rot, anchor='w', font=tema.fonte(10, 'bold'), fill=tema.TEXTO)
            self.create_text(w - 4 * e, y, text=f'{n:,}'.replace(',', '.') + f'  ·  {pct_txt(n / total)}', anchor='e',
                             font=tema.fonte(10), fill=tema.TEXTO_2)
            alvo = self.create_rectangle(x0 - 4, y - passo / 2, w, y + passo / 2, fill='', outline='')
            self._dicas[alvo] = self._texto(rot, n, total, desc)
            self._dicas[t] = self._dicas[alvo]
            y += passo

    @staticmethod
    def _texto(rot, n, total, desc):
        return (f'{rot} – {desc}' if desc else rot) + f'\n{n:,}'.replace(',', '.') + f' de {total:,}'.replace(',', '.') + \
            f' ({pct_txt(n / total)})'

    def fatia_em(self, x, y):
        """Índice da fatia sob o ponto (x, y) ou None."""
        import math
        if not self._geo:
            return None
        cx, cy, r, cortes = self._geo
        if (x - cx) ** 2 + (y - cy) ** 2 > r * r:
            return None
        ang = math.degrees(math.atan2(y - cy, x - cx))
        for i, (a0, a1) in enumerate(cortes):
            for a in (ang, ang + 360, ang - 360):
                if a0 <= a < a1:
                    return i
        return None

    def _mover(self, ev):
        self.delete('dica')
        texto = None
        i = self.fatia_em(ev.x, ev.y)
        if i is not None:
            total = sum(n for _, n, _, _ in self.fatias)
            rot, n, _, desc = self.fatias[i]
            texto = self._texto(rot, n, total, desc)
        else:
            for item in reversed(self.find_overlapping(ev.x - 2, ev.y - 2, ev.x + 2, ev.y + 2)):
                if item in self._dicas:
                    texto = self._dicas[item]
                    break
        if not texto:
            return
        t = self.create_text(ev.x + 16, ev.y + 14, text=texto, anchor='nw', font=tema.fonte(9), fill='#FFFFFF', tags='dica')
        x0, y0, x1, y1 = self.bbox(t)
        w, h = self.winfo_width(), self.winfo_height()
        self.move(t, (w - 10 - x1) if x1 > w - 10 else 0, -(y1 - y0) - 34 if y1 > h - 8 else 0)
        x0, y0, x1, y1 = self.bbox(t)
        r = self.create_rectangle(x0 - 10, y0 - 7, x1 + 10, y1 + 7, fill='#2B3445', outline='', tags='dica')
        self.tag_raise(t, r)
