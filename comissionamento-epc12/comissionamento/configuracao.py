"""Onde fica o banco de dados e o registro de erros.

Programa portátil (não precisa instalar): por padrão o banco fica na pasta 'dados' ao lado do executável, para
poder ser levado junto ou ficar numa pasta de rede. Se essa pasta não aceitar gravação, usa a pasta do usuário
(%LOCALAPPDATA%\\ComissionamentoEPC12). O último banco aberto é lembrado."""
import datetime
import json
import os
import sys
from pathlib import Path

NOME_BANCO = 'comissionamento_epc12.db'


def pasta_programa():
    if getattr(sys, 'frozen', False):          # executável gerado pelo PyInstaller
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def pasta_usuario():
    base = os.environ.get('LOCALAPPDATA') or os.environ.get('XDG_DATA_HOME') or str(Path.home() / '.local' / 'share')
    p = Path(base) / 'ComissionamentoEPC12'
    p.mkdir(parents=True, exist_ok=True)
    return p


def _gravavel(pasta):
    try:
        pasta.mkdir(parents=True, exist_ok=True)
        teste = pasta / '.teste_gravacao'
        teste.write_text('ok')
        teste.unlink()
        return True
    except OSError:
        return False


def _arquivo_config():
    return pasta_usuario() / 'config.json'


def _ler_config():
    try:
        return json.loads(_arquivo_config().read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return {}


def banco_padrao():
    ultimo = _ler_config().get('banco')
    if ultimo and Path(ultimo).exists():
        return Path(ultimo)
    pasta = pasta_programa() / 'dados'
    return (pasta if _gravavel(pasta) else pasta_usuario()) / NOME_BANCO


def lembrar_banco(caminho):
    cfg = _ler_config()
    cfg['banco'] = str(Path(caminho).resolve())
    try:
        _arquivo_config().write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding='utf-8')
    except OSError:
        pass


def arquivo_log():
    return pasta_usuario() / 'erros.log'


def registrar_erro(texto):
    try:
        with open(arquivo_log(), 'a', encoding='utf-8') as f:
            f.write(f'\n===== {datetime.datetime.now():%d/%m/%Y %H:%M:%S}\n{texto}')
    except OSError:
        pass
