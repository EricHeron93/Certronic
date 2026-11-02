"""estado_vencimento — Estado general de vencimiento.

Documentos com prazo vencido ou a vencer nos próximos 7-30 dias,
por colaborador (CPF) ou contratista (CNPJ).

Formato real do arquivo: HTML disfarçado de .xls.
Colunas: Plantas | Recurso | Detalhe | Identificador | Tipo |
         Documento | Vencimento | Período

Vencimento '0000-00-00' significa "sem data definida" — é ignorado
para o cálculo de proximidade, mas reportado.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from .util import ler_tabela_html, normalizar_cpf


@dataclass
class Vencimento:
    cpf: str            # normalizado, 11 dígitos (ou 14 se CNPJ)
    nome: str
    documento: str
    vencimento: date | None   # None quando '0000-00-00'
    eh_pessoa_fisica: bool


def _parse_data(texto: str) -> date | None:
    texto = texto.strip()
    if not texto or texto == "0000-00-00":
        return None
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(texto, fmt).date()
        except ValueError:
            continue
    return None


def ler(caminho: str | Path) -> list[Vencimento]:
    linhas = ler_tabela_html(Path(caminho))
    if not linhas:
        return []

    resultado = []
    for celulas in linhas[1:]:  # pula cabeçalho
        if len(celulas) < 8:
            continue
        _, recurso, detalhe, identificador, tipo, documento, vencimento, _ = celulas[:8]
        resultado.append(
            Vencimento(
                cpf=normalizar_cpf(identificador),
                nome=detalhe.strip(),
                documento=documento.strip(),
                vencimento=_parse_data(vencimento),
                eh_pessoa_fisica=(recurso.strip().lower() == "empleado"),
            )
        )
    return resultado
