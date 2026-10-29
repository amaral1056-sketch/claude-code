"""Banco de dados SQLite do programa (um único arquivo .db).

Conteúdo:
- base mestre da LI (unidades, SOP, SSOP);
- dados importados das exportações da FIC (cadastro FVI, FVM, CCM, TAP, TTAS) – cada importação substitui a anterior
  do mesmo tipo, porque as exportações da FIC são sempre a lista completa;
- critérios de "aprovado", listas oficiais de status e parâmetros (editáveis);
- registro de todas as importações e histórico do avanço por unidade em cada data de referência."""
import datetime
import json
import shutil
import sqlite3
from pathlib import Path

from . import padroes

VERSAO_ESQUEMA = 1

ESQUEMA = """
CREATE TABLE IF NOT EXISTS meta (chave TEXT PRIMARY KEY, valor TEXT);
CREATE TABLE IF NOT EXISTS unidade (nome TEXT PRIMARY KEY, ordem INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS sop (
    sop TEXT PRIMARY KEY, unidade TEXT NOT NULL, descricao TEXT, grupo TEXT, ordem INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS ssop (
    ssop TEXT PRIMARY KEY, item INTEGER, unidade TEXT NOT NULL, sop TEXT NOT NULL, descricao TEXT,
    documentos TEXT, grupo TEXT, ordem INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS importacao (
    id INTEGER PRIMARY KEY AUTOINCREMENT, tipo TEXT NOT NULL, arquivo TEXT, importado_em TEXT NOT NULL,
    data_referencia TEXT, linhas INTEGER, observacao TEXT);
CREATE TABLE IF NOT EXISTS fvi_tag (
    linha INTEGER PRIMARY KEY, tag TEXT, unidade TEXT, sop TEXT, ssop TEXT, classe TEXT, disciplina TEXT,
    malha TEXT, contrato TEXT, n_fvi TEXT, st_fvi TEXT, st_cii TEXT, st_ciii TEXT, st_civ TEXT, st_cv TEXT,
    dados TEXT);
CREATE TABLE IF NOT EXISTS fvm (linha INTEGER PRIMARY KEY, malha TEXT, sop TEXT, ssop TEXT, status TEXT);
CREATE TABLE IF NOT EXISTS ccm (linha INTEGER PRIMARY KEY, numero TEXT, ssop TEXT, status TEXT);
CREATE TABLE IF NOT EXISTS tap (linha INTEGER PRIMARY KEY, numero TEXT, ssop TEXT, status TEXT);
CREATE TABLE IF NOT EXISTS ttas (linha INTEGER PRIMARY KEY, numero TEXT, sop TEXT, status TEXT);
CREATE TABLE IF NOT EXISTS li_item (linha INTEGER PRIMARY KEY, tag TEXT, sop TEXT, ssop TEXT, descricao TEXT);
CREATE TABLE IF NOT EXISTS li_malha (linha INTEGER PRIMARY KEY, malha TEXT, sop TEXT, ssop TEXT, descricao TEXT);
CREATE TABLE IF NOT EXISTS criterio (
    id INTEGER PRIMARY KEY AUTOINCREMENT, etapa TEXT NOT NULL, codigo TEXT NOT NULL, origem TEXT,
    UNIQUE (etapa, codigo));
CREATE TABLE IF NOT EXISTS status_fic (
    documento TEXT NOT NULL, codigo TEXT NOT NULL, descricao TEXT, PRIMARY KEY (documento, codigo));
CREATE TABLE IF NOT EXISTS historico (
    data_referencia TEXT NOT NULL, unidade TEXT NOT NULL, registrado_em TEXT NOT NULL,
    sop INTEGER, ssop INTEGER, tags INTEGER, cii_aplic INTEGER, cii_aprov INTEGER, ciii_aplic INTEGER,
    ciii_aprov INTEGER, cv_aplic INTEGER, cv_aprov INTEGER, malhas INTEGER, fvm_ciii INTEGER, fvm_cv INTEGER,
    ccm INTEGER, tap INTEGER, ttas1 INTEGER, ttas2 INTEGER, tags_sem_fvi INTEGER,
    PRIMARY KEY (data_referencia, unidade));
"""

TABELAS_IMPORTACAO = {'FVI': 'fvi_tag', 'FVM': 'fvm', 'CCM': 'ccm', 'TAP': 'tap', 'TTAS': 'ttas',
                      'LI_ITENS': 'li_item', 'LI_MALHAS': 'li_malha'}
TIPOS_FIC = ('FVI', 'FVM', 'CCM', 'TAP', 'TTAS')
COLUNAS_HISTORICO = ('sop', 'ssop', 'tags', 'cii_aplic', 'cii_aprov', 'ciii_aplic', 'ciii_aprov', 'cv_aplic',
                     'cv_aprov', 'malhas', 'fvm_ciii', 'fvm_cv', 'ccm', 'tap', 'ttas1', 'ttas2', 'tags_sem_fvi')


def agora():
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')


class Banco:
    def __init__(self, caminho):
        self.caminho = Path(caminho)
        self.caminho.parent.mkdir(parents=True, exist_ok=True)
        self.con = sqlite3.connect(str(self.caminho))
        self.con.row_factory = sqlite3.Row
        self.con.execute('PRAGMA foreign_keys = ON')
        self._preparar()

    def fechar(self):
        self.con.close()

    # ------------------------------------------------------------------ criação
    def _preparar(self):
        with self.con:
            self.con.executescript(ESQUEMA)
            if self.meta('versao_esquema') is None:
                self.definir_meta('versao_esquema', str(VERSAO_ESQUEMA))
                self.definir_meta('criado_em', agora())
                self.definir_meta('incluir_sem_fvi', 'SIM')
                self._criterios_padrao()
                self._status_padrao()

    def _criterios_padrao(self):
        self.con.execute('DELETE FROM criterio')
        self.con.executemany('INSERT INTO criterio (etapa, codigo, origem) VALUES (?, ?, ?)', padroes.CRITERIOS)

    def _status_padrao(self):
        self.con.execute('DELETE FROM status_fic')
        for doc, lista in padroes.STATUS_FIC.items():
            self.con.executemany('INSERT INTO status_fic VALUES (?, ?, ?)', [(doc, c, d) for c, d in lista])

    # ------------------------------------------------------------------ meta / parâmetros
    def meta(self, chave, padrao=None):
        r = self.con.execute('SELECT valor FROM meta WHERE chave = ?', (chave,)).fetchone()
        return r[0] if r else padrao

    def definir_meta(self, chave, valor):
        self.con.execute('INSERT OR REPLACE INTO meta VALUES (?, ?)', (chave, valor))

    @property
    def incluir_sem_fvi(self):
        """Parâmetro da planilha (CRITERIOS!F2): TAGs sem FVI gerada entram no denominador do FVI?"""
        return self.meta('incluir_sem_fvi', 'SIM') != 'NÃO'

    def definir_incluir_sem_fvi(self, sim):
        with self.con:
            self.definir_meta('incluir_sem_fvi', 'SIM' if sim else 'NÃO')

    # ------------------------------------------------------------------ critérios e status
    def criterios(self):
        return [tuple(r) for r in self.con.execute('SELECT id, etapa, codigo, origem FROM criterio ORDER BY id')]

    def adicionar_criterio(self, etapa, codigo, origem=''):
        with self.con:
            self.con.execute('INSERT INTO criterio (etapa, codigo, origem) VALUES (?, ?, ?)',
                             (etapa.strip().upper(), codigo.strip(), origem.strip()))

    def remover_criterio(self, id_):
        with self.con:
            self.con.execute('DELETE FROM criterio WHERE id = ?', (id_,))

    def restaurar_criterios(self):
        with self.con:
            self._criterios_padrao()

    def status_fic(self):
        return [tuple(r) for r in self.con.execute('SELECT documento, codigo, descricao FROM status_fic ORDER BY rowid')]

    def substituir_status_fic(self, documento, lista):
        with self.con:
            self.con.execute('DELETE FROM status_fic WHERE documento = ?', (documento,))
            self.con.executemany('INSERT OR REPLACE INTO status_fic VALUES (?, ?, ?)',
                                 [(documento, c, d) for c, d in lista])

    # ------------------------------------------------------------------ base LI
    def tem_base(self):
        return self.con.execute('SELECT COUNT(*) FROM ssop').fetchone()[0] > 0

    def substituir_base(self, linhas, arquivo):
        """linhas: dicts com item, unidade, sop, descricao_sop, ssop, descricao_ssop, documentos, grupo (ordem da LI)."""
        unidades = list(dict.fromkeys(r['unidade'] for r in linhas))
        sops = {}
        for r in linhas:
            sops.setdefault(r['sop'], r)
        with self.con:
            self.con.execute('DELETE FROM ssop')
            self.con.execute('DELETE FROM sop')
            self.con.execute('DELETE FROM unidade')
            self.con.executemany('INSERT INTO unidade VALUES (?, ?)', [(u, i) for i, u in enumerate(unidades)])
            self.con.executemany('INSERT INTO sop VALUES (?, ?, ?, ?, ?)',
                                 [(s, r['unidade'], r['descricao_sop'], r['grupo'], i) for i, (s, r) in enumerate(sops.items())])
            self.con.executemany('INSERT INTO ssop VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
                                 [(r['ssop'], r['item'], r['unidade'], r['sop'], r['descricao_ssop'], r['documentos'],
                                   r['grupo'], i) for i, r in enumerate(linhas)])
            self._registrar('LI', arquivo, None, len(linhas),
                            f'{len(unidades)} unidades, {len(sops)} SOP, {len(linhas)} SSOP')

    def base(self):
        """(unidades em ordem, lista de SOP, lista de SSOP) como dicts."""
        un = [r[0] for r in self.con.execute('SELECT nome FROM unidade ORDER BY ordem')]
        sops = [dict(r) for r in self.con.execute('SELECT sop, unidade, descricao, grupo FROM sop ORDER BY ordem')]
        ssops = [dict(r) for r in self.con.execute(
            'SELECT ssop, item, unidade, sop, descricao, documentos, grupo FROM ssop ORDER BY ordem')]
        return un, sops, ssops

    # ------------------------------------------------------------------ importações da FIC
    def substituir_dados(self, tipo, registros, arquivo, data_referencia, observacao=''):
        """Troca todo o conteúdo do tipo (FVI/FVM/CCM/TAP/TTAS) pelos registros (dicts com as colunas da tabela)."""
        tabela = TABELAS_IMPORTACAO[tipo]
        with self.con:
            self.con.execute(f'DELETE FROM {tabela}')
            if registros:
                cols = list(registros[0].keys())
                sql = f'INSERT INTO {tabela} ({", ".join(cols)}) VALUES ({", ".join("?" * len(cols))})'
                self.con.executemany(sql, [[(json.dumps(r[c], ensure_ascii=False, default=str) if c == 'dados' else r[c])
                                            for c in cols] for r in registros])
            self._registrar(tipo, arquivo, data_referencia, len(registros), observacao)

    def dados(self, tipo):
        tabela = TABELAS_IMPORTACAO[tipo]
        return [dict(r) for r in self.con.execute(f'SELECT * FROM {tabela} ORDER BY linha')]

    def _registrar(self, tipo, arquivo, data_referencia, linhas, observacao):
        self.con.execute('INSERT INTO importacao (tipo, arquivo, importado_em, data_referencia, linhas, observacao) '
                         'VALUES (?, ?, ?, ?, ?, ?)', (tipo, str(arquivo or ''), agora(), data_referencia, linhas, observacao))

    def importacoes(self):
        return [dict(r) for r in self.con.execute('SELECT * FROM importacao ORDER BY id DESC')]

    def ultima_importacao(self, tipo):
        r = self.con.execute('SELECT * FROM importacao WHERE tipo = ? ORDER BY id DESC LIMIT 1', (tipo,)).fetchone()
        return dict(r) if r else None

    def data_referencia_atual(self):
        """Data de referência mais recente entre as importações da FIC (texto AAAA-MM-DD) ou None."""
        r = self.con.execute(f"SELECT MAX(data_referencia) FROM importacao WHERE tipo IN ({', '.join('?' * len(TIPOS_FIC))})",
                             TIPOS_FIC).fetchone()
        return r[0] if r else None

    # ------------------------------------------------------------------ histórico
    def gravar_historico(self, data_referencia, por_unidade):
        """por_unidade: {unidade: {coluna: valor}} com as COLUNAS_HISTORICO. Substitui o que houver na mesma data."""
        with self.con:
            self.con.execute('DELETE FROM historico WHERE data_referencia = ?', (data_referencia,))
            sql = (f'INSERT INTO historico (data_referencia, unidade, registrado_em, {", ".join(COLUNAS_HISTORICO)}) '
                   f'VALUES (?, ?, ?, {", ".join("?" * len(COLUNAS_HISTORICO))})')
            ts = agora()
            self.con.executemany(sql, [(data_referencia, u, ts, *[v[c] for c in COLUNAS_HISTORICO])
                                       for u, v in por_unidade.items()])

    def historico(self):
        return [dict(r) for r in self.con.execute('SELECT * FROM historico ORDER BY data_referencia, unidade')]

    def remover_historico(self, data_referencia):
        with self.con:
            self.con.execute('DELETE FROM historico WHERE data_referencia = ?', (data_referencia,))

    # ------------------------------------------------------------------ cópia de segurança
    def copiar_para(self, destino):
        """Cópia consistente do banco (API de backup do SQLite)."""
        destino = Path(destino)
        if destino.resolve() == self.caminho.resolve():
            raise ValueError('O destino é o próprio banco.')
        tmp = destino.with_suffix(destino.suffix + '.tmp')
        out = sqlite3.connect(str(tmp))
        with out:
            self.con.backup(out)
        out.close()
        shutil.move(str(tmp), str(destino))
