"""Utilitários compartilhados pelos leitores.

Os relatórios do Certronic vêm em dois formatos disfarçados:
- documentacao_enviada: arquivo .xls que na verdade é um .xlsx (zip OOXML)
- estado_vencimento / documentacao_faltante / lista_pessoal: arquivo .xls que na verdade é uma tabela HTML
Este módulo detecta o formato real e oferece parsing para ambos.
"""

from __future__ import annotations

import re
import zipfile
from html.parser import HTMLParser
from pathlib import Path


def normalizar_cpf(valor) -> str:
    """Converte CPF de qualquer origem para 11 dígitos com zeros à esquerda.

    Na Matriz o CPF aparece como número (perde zeros à esquerda, ex: 3898590607);
    nos relatórios vem como texto. Normalizar é obrigatório antes de cruzar.
    """
    if valor is None:
        return ""
    digitos = re.sub(r"\D", "", str(valor))
    # CNPJ (14 dígitos) passa direto; CPF é completado até 11
    if len(digitos) > 11:
        return digitos
    return digitos.zfill(11)


def detectar_formato(caminho: Path) -> str:
    """Retorna 'xlsx', 'html' ou 'desconhecido' olhando os bytes iniciais."""
    with open(caminho, "rb") as f:
        inicio = f.read(8)
    if inicio[:4] == b"PK\x03\x04":
        return "xlsx"
    if inicio.lstrip()[:1] == b"<":
        return "html"
    if zipfile.is_zipfile(caminho):
        return "xlsx"
    return "desconhecido"


class _TabelaHTMLParser(HTMLParser):
    """Extrai a primeira <table> de um HTML como lista de linhas (listas de célula)."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.linhas: list[list[str]] = []
        self._linha_atual: list[str] | None = None
        self._celula_atual: list[str] | None = None

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._linha_atual = []
        elif tag in ("td", "th"):
            self._celula_atual = []

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._celula_atual is not None:
            texto = "".join(self._celula_atual).strip()
            if self._linha_atual is not None:
                self._linha_atual.append(texto)
            self._celula_atual = None
        elif tag == "tr" and self._linha_atual is not None:
            if self._linha_atual:
                self.linhas.append(self._linha_atual)
            self._linha_atual = None

    def handle_data(self, data):
        if self._celula_atual is not None:
            self._celula_atual.append(data)


def ler_tabela_html(caminho: Path) -> list[list[str]]:
    """Lê um 'falso .xls' (HTML) e devolve as linhas da tabela.

    A primeira linha retornada é o cabeçalho.
    """
    # Os exports declaram UTF-8 no <meta>, mas alguns vêm com mojibake;
    # utf-8 com fallback latin-1 cobre os dois casos observados.
    raw = Path(caminho).read_bytes()
    try:
        conteudo = raw.decode("utf-8")
    except UnicodeDecodeError:
        conteudo = raw.decode("latin-1")

    parser = _TabelaHTMLParser()
    parser.feed(conteudo)
    return parser.linhas
