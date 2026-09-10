import json
from pathlib import Path

# Credenciais do SUAP (usuário/senha do login que assina e tramita) isoladas do código-fonte.
# O arquivo credenciais_suap.json NÃO vai para o Git (ver .gitignore), igual ao credenciais.json:
# numa máquina nova, copie-o manualmente para a pasta do projeto (e para dentro de dist/ quando
# for rodar o .exe). Modelo em credenciais_suap.exemplo.json.
#
# A leitura é preguiçosa (só na 1ª chamada de usuario()/senha()) e em cache, pra que só
# importar este módulo - como o gui.py faz com todos os scripts - não tenha efeito colateral.

_CAMINHO = Path(__file__).with_name("credenciais_suap.json")
_cache = None


def _carregar():
    global _cache
    if _cache is None:
        try:
            with open(_CAMINHO, encoding="utf-8") as arquivo:
                dados = json.load(arquivo)
            _cache = {"usuario": str(dados["usuario"]), "senha": str(dados["senha"])}
        except (OSError, ValueError, KeyError) as erro:
            raise RuntimeError(
                f"Não foi possível ler {_CAMINHO.name}. Crie o arquivo na pasta do projeto no "
                'formato {"usuario": "...", "senha": "..."} (modelo: credenciais_suap.exemplo.json).'
            ) from erro
    return _cache


def usuario():
    return _carregar()["usuario"]


def senha():
    return _carregar()["senha"]
