"""lista_pessoal — Lista do Pessoal (Lista de Ativos).

Situação de acesso de cada colaborador no cliente:
- 'Autorizada'   → acesso liberado para trabalhar na unidade
- 'Desabilitada' → sem acesso (mas ainda é colaborador ativo)
- 'Baixas'       → ex-funcionário (pendências dele são ignoradas)

Formato real do arquivo: HTML disfarçado de .xls.
Colunas: Planta | Contratante | Sobrenome | Nome | CNPJ | Tipo | Acesso
(a coluna 'CNPJ' na prática carrega o CPF quando Tipo == 'Recurso')
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .util import ler_tabela_html, normalizar_cpf


@dataclass
class Colaborador:
    cpf: str
    nome_completo: str
    acesso: str  # Autorizada | Desabilitada | Baixas


def ler(caminho: str | Path) -> dict[str, Colaborador]:
    """Devolve {cpf_normalizado: Colaborador}."""
    linhas = ler_tabela_html(Path(caminho))
    if not linhas:
        return {}

    resultado: dict[str, Colaborador] = {}
    for celulas in linhas[1:]:  # pula cabeçalho
        if len(celulas) < 7:
            continue
        _, _, sobrenome, nome, identificador, _tipo, acesso = celulas[:7]
        cpf = normalizar_cpf(identificador)
        resultado[cpf] = Colaborador(
            cpf=cpf,
            nome_completo=f"{sobrenome.strip()} {nome.strip()}".strip(),
            acesso=acesso.strip(),
        )
    return resultado
