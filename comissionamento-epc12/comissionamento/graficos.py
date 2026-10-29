"""Gráficos em tk.Canvas no estilo HUD monocromático: grade de fundo de 16 px, linhas finas, escalas de cinza e
texturas (hachura, halftone) para distinguir séries sem cor. Dica (bloco preto, texto branco) ao passar o mouse."""
import math
import tkinter as tk

from PIL import Image, ImageDraw, ImageTk

from . import tema


def pct_txt(v):
    return '–' if v is None else f'{v * 100:.1f}%'.replace('.', ',')


def _grade_fundo(canvas, x0, y0, x1, y1, passo=16):
    """effect.gridPattern: grade de 16 px em line.faint."""
    x = x0
    while x <= x1:
        canvas.create_line(x, y0, x, y1, fill=tema.LINHA_FRACA)
        x += passo
    y = y0
    while y <= y1:
        canvas.create_line(x0, y, x1, y, fill=tema.LINHA_FRACA)
        y += passo


def _dica(canvas, ev, texto):
    canvas.delete('dica')
    if not texto:
        return
    t = canvas.create_text(ev.x + 16, ev.y + 14, text=texto, anchor='nw', font=tema.fonte(9, tipo='dado'),
                           fill=tema.SOBRE_INVERSO, tags='dica')
    x0, y0, x1, y1 = canvas.bbox(t)
    w, h = canvas.winfo_width(), canvas.winfo_height()
    canvas.move(t, (w - 10 - x1) if x1 > w - 10 else 0, -(y1 - y0) - 34 if y1 > h - 8 else 0)
    x0, y0, x1, y1 = canvas.bbox(t)
    r = canvas.create_rectangle(x0 - 8, y0 - 6, x1 + 8, y1 + 6, fill=tema.INVERSO, outline='', tags='dica')
    canvas.tag_raise(t, r)


# ---------------------------------------------------------------- texturas (Pillow)
def textura(estilo, largura, altura, k=1):
    """Imagem RGBA preenchida com o estilo (tipo, cor) – sólido, hachura, pontos, cruzado, linhas, grade ou vazio."""
    tipo, cor = estilo
    if tipo == 'solido':
        return Image.new('RGBA', (largura, altura), cor)
    img = Image.new('RGBA', (largura, altura), tema.BASE)
    d = ImageDraw.Draw(img)
    p, w = 6 * k, max(1, int(1.4 * k))
    if tipo in ('hachura', 'cruzado'):
        for x in range(-altura, largura + altura, p):
            d.line((x, 0, x + altura, altura), fill=cor, width=w)
    if tipo == 'cruzado':
        for x in range(-altura, largura + altura, p):
            d.line((x, altura, x + altura, 0), fill=cor, width=w)
    if tipo == 'pontos':                 # effect.halftoneDots
        r = 1.3 * k
        for y in range(0, altura, int(4.5 * k)):
            for x in range(0, largura, int(4.5 * k)):
                d.ellipse((x - r, y - r, x + r, y + r), fill=cor)
    if tipo == 'linhas':
        for y in range(0, altura, int(5 * k)):
            d.line((0, y, largura, y), fill=cor, width=w)
    if tipo == 'grade':
        for x in range(0, largura, p):
            d.line((x, 0, x, altura), fill=cor, width=max(1, w - 1))
        for y in range(0, altura, p):
            d.line((0, y, largura, y), fill=cor, width=max(1, w - 1))
    return img


def texto_sobre(estilo):
    """Cor de texto legível sobre o estilo (branco só sobre os dois cinzas mais escuros)."""
    return tema.SOBRE_INVERSO if estilo[0] == 'solido' and estilo[1] in (tema.DADOS[0], tema.DADOS[1]) else tema.TEXTO


class Grafico(tk.Canvas):
    """Área de gráfico com título (caixa-alta), subtítulo e legenda; redesenha ao redimensionar."""

    SERIES = ((tema.DADOS[0], (), 'circulo'), (tema.DADOS[1], (6, 3), 'quadrado'), (tema.DADOS[2], (2, 3), 'losango'))

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
        self.create_text(2 * e, 2 * e, text=tema.caixa_alta(self.titulo), anchor='nw',
                         font=tema.fonte(12, 'bold', 'titulo'), fill=tema.TEXTO)
        if self.subtitulo:
            self.create_text(2 * e, 24 * e, text=self.subtitulo, anchor='nw', font=tema.fonte(8), fill=tema.TEXTO_MUDO)
        if self._desenhar:
            self._desenhar(w, h, e)

    def vazio(self, w, h, msg='SEM DADOS // importe as planilhas para ver este gráfico.'):
        self.create_text(w / 2, h / 2 + 14, text=msg, font=tema.fonte(9, tipo='dado'), fill=tema.TEXTO_MUDO,
                         width=w - 40)

    def legenda(self, series, e):
        x, y = 2 * e, 46 * e
        for nome, estilo in series:
            self._marcador(x + 6 * e, y + 6 * e, estilo, e)
            t = self.create_text(x + 16 * e, y + 6 * e, text=tema.caixa_alta(nome), anchor='w',
                                 font=tema.fonte(9, tipo='dado'), fill=tema.TEXTO_2)
            x = self.bbox(t)[2] + 16 * e

    def _marcador(self, x, y, estilo, e):
        """Série = (cor, traço, forma do marcador)."""
        cor, _, forma = estilo
        r = 4 * e
        if forma == 'quadrado':
            return self.create_rectangle(x - r, y - r, x + r, y + r, fill=cor, outline=tema.SUPERFICIE)
        if forma == 'losango':
            return self.create_polygon(x, y - r - 1, x + r + 1, y, x, y + r + 1, x - r - 1, y, fill=cor,
                                       outline=tema.SUPERFICIE)
        return self.create_oval(x - r, y - r, x + r, y + r, fill=cor, outline=tema.SUPERFICIE)

    def _mover(self, ev):
        alvo = None
        for item in reversed(self.find_overlapping(ev.x - 3, ev.y - 3, ev.x + 3, ev.y + 3)):
            if item in self._dicas:
                alvo = item
                break
        _dica(self, ev, self._dicas.get(alvo))

    def eixo_pct(self, x0, y0, x1, y1):
        _grade_fundo(self, x0, y0, x1, y1)
        for p in (0, .25, .5, .75, 1):
            y = y1 - (y1 - y0) * p
            self.create_line(x0, y, x1, y, fill=tema.LINHA if p == 0 else tema.LINHA_SUTIL, dash=() if p == 0 else (2, 3))
            self.create_text(x0 - 6, y, text=f'{int(p * 100)}%', anchor='e', font=tema.fonte(8, tipo='dado'),
                             fill=tema.TEXTO_MUDO)
        return lambda v: y1 - (y1 - y0) * max(0, min(1, v))

    def linhas(self, rotulos, series):
        """series: [(nome, _, [valor 0..1 | None por rótulo])] – evolução no tempo; séries por traço e marcador."""
        def desenhar(w, h, e):
            if not rotulos:
                return self.vazio(w, h, 'SEM HISTÓRICO // a evolução aparece a partir da primeira importação com data.')
            estilos = [self.SERIES[i % 3] for i in range(len(series))]
            self.legenda([(n, est) for (n, _, _), est in zip(series, estilos)], e)
            x0, y0, x1, y1 = 40 * e, 76 * e, w - 60 * e, h - 26 * e
            ymap = self.eixo_pct(x0, y0, x1, y1)
            passo = (x1 - x0) / max(1, len(rotulos) - 1)
            xs = [x0 + passo * i if len(rotulos) > 1 else (x0 + x1) / 2 for i in range(len(rotulos))]
            mostrar = max(1, len(rotulos) // 8)
            for i, r in enumerate(rotulos):
                if i % mostrar == 0 or i == len(rotulos) - 1:
                    self.create_text(xs[i], y1 + 6 * e, text=r, anchor='n', font=tema.fonte(8, tipo='dado'),
                                     fill=tema.TEXTO_MUDO)
            for (nome, _, vals), est in zip(series, estilos):
                cor, traco, _ = est
                pts = [(xs[i], ymap(v)) for i, v in enumerate(vals) if v is not None]
                if len(pts) > 1:
                    self.create_line(*[c for p in pts for c in p], fill=cor, width=max(2, round(2 * e)), dash=traco)
                for i, v in enumerate(vals):
                    if v is None:
                        continue
                    item = self._marcador(xs[i], ymap(v), est, e)
                    alvo = self.create_oval(xs[i] - 10, ymap(v) - 10, xs[i] + 10, ymap(v) + 10, fill='', outline='')
                    self._dicas[item] = self._dicas[alvo] = f'{rotulos[i]}\n{nome}: {pct_txt(v)}'
                if pts:
                    ult = next(v for v in reversed(vals) if v is not None)
                    self.create_text(pts[-1][0] + 10 * e, pts[-1][1], text=pct_txt(ult), anchor='w',
                                     font=tema.fonte(9, tipo='dado'), fill=tema.TEXTO)
        self.definir(desenhar)


class Pizza(tk.Canvas):
    """Pizza monocromática (Pillow em 4x): 4 cinzas + texturas, separadores brancos e contorno preto fino.
    Legenda ao lado (status, quantidade, %); passar o mouse mostra a descrição oficial do status."""

    def __init__(self, master, altura=300):
        super().__init__(master, background=tema.SUPERFICIE, highlightthickness=0, height=int(altura * tema.escala))
        self.fatias, self._imgs, self._geo = [], [], None
        self._dicas = {}
        self.bind('<Configure>', lambda e: self.redesenhar())
        self.bind('<Motion>', self._mover)
        self.bind('<Leave>', lambda e: self.delete('dica'))

    def definir(self, fatias):
        """fatias: [(rótulo, quantidade, estilo (tipo, cor), descrição)] – na ordem em que serão desenhadas."""
        self.fatias = [f for f in fatias if f[1] > 0]
        self.redesenhar()

    def redesenhar(self):
        self.delete('all')
        self._dicas, self._geo, self._imgs = {}, None, []
        w, h = self.winfo_width(), self.winfo_height()
        if w < 80 or h < 80:
            return
        e = tema.escala
        total = sum(n for _, n, _, _ in self.fatias)
        if not total:
            self.create_text(w / 2, h / 2, text='SEM DADOS // importe as planilhas para ver este gráfico.',
                             font=tema.fonte(9, tipo='dado'), fill=tema.TEXTO_MUDO, width=w - 40)
            return
        d = int(min(h - 16 * e, w * 0.44))
        cx, cy = 8 * e + d / 2, h / 2
        k = 4
        img = Image.new('RGBA', (d * k, d * k), (0, 0, 0, 0))
        dr = ImageDraw.Draw(img)
        caixa = (k, k, d * k - k - 1, d * k - k - 1)
        ang, cortes = -90.0, []
        for rot, n, estilo, _ in self.fatias:
            fim = ang + 360.0 * n / total
            mascara = Image.new('L', img.size, 0)
            ImageDraw.Draw(mascara).pieslice(caixa, ang, fim, fill=255)
            img.paste(textura(estilo, *img.size, k=k), (0, 0), mascara)
            cortes.append((ang, fim))
            ang = fim
        if len(self.fatias) > 1:     # separadores brancos de 2 px
            r = (caixa[2] - caixa[0]) / 2
            for a0, _ in cortes:
                a = math.radians(a0)
                dr.line((d * k / 2, d * k / 2, d * k / 2 + r * math.cos(a), d * k / 2 + r * math.sin(a)),
                        fill=tema.SUPERFICIE, width=2 * k)
        dr.ellipse(caixa, outline=tema.LINHA_FORTE, width=k)        # contorno (line.strong, 1 px)
        foto = ImageTk.PhotoImage(img.resize((d, d), Image.LANCZOS), master=self)
        self._imgs.append(foto)
        self.create_image(cx, cy, image=foto)
        self._geo = (cx, cy, d / 2, cortes)
        for (a0, a1), (rot, n, estilo, _) in zip(cortes, self.fatias):
            if (a1 - a0) / 360 >= 0.06:
                meio = math.radians((a0 + a1) / 2)
                x, y = cx + d * 0.30 * math.cos(meio), cy + d * 0.30 * math.sin(meio)
                t = self.create_text(x, y, text=pct_txt(n / total), font=tema.fonte(9, 'bold', 'dado'),
                                     fill=texto_sobre(estilo))
                if estilo[0] != 'solido':   # rótulo sobre textura ganha fundo branco
                    x0_, y0_, x1_, y1_ = self.bbox(t)
                    self.tag_lower(self.create_rectangle(x0_ - 3, y0_ - 1, x1_ + 3, y1_ + 1, fill=tema.BASE,
                                                         outline=tema.LINHA_FORTE), t)
        # legenda
        x0 = cx + d / 2 + 22 * e
        passo = min(26 * e, (h - 20 * e) / max(1, len(self.fatias)))
        y = cy - passo * len(self.fatias) / 2 + passo / 2
        q = int(12 * e)
        for rot, n, estilo, desc in self.fatias:
            amostra = ImageTk.PhotoImage(textura(estilo, q * 4, q * 4, k=4).resize((q, q), Image.LANCZOS), master=self)
            self._imgs.append(amostra)
            self.create_image(x0 + q / 2, y, image=amostra)
            self.create_rectangle(x0, y - q / 2, x0 + q, y + q / 2, outline=tema.LINHA_FORTE)
            t = self.create_text(x0 + q + 8 * e, y, text=rot, anchor='w', font=tema.fonte(9, 'bold'), fill=tema.TEXTO)
            self.create_text(w - 4 * e, y, text=f'{n:,}'.replace(',', '.') + f'  {pct_txt(n / total):>6}', anchor='e',
                             font=tema.fonte(10, tipo='dado'), fill=tema.TEXTO_2)
            alvo = self.create_rectangle(x0 - 4, y - passo / 2, w, y + passo / 2, fill='', outline='')
            self._dicas[alvo] = self._dicas[t] = self._texto(rot, n, total, desc)
            y += passo

    @staticmethod
    def _texto(rot, n, total, desc):
        return (f'{rot} // {desc}' if desc else rot) + f'\n{n:,}'.replace(',', '.') + \
            f' de {total:,}'.replace(',', '.') + f' ({pct_txt(n / total)})'

    def fatia_em(self, x, y):
        """Índice da fatia sob o ponto (x, y) ou None."""
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
        _dica(self, ev, texto)
