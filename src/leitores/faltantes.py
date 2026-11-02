"""documentacao_faltante — Documentação faltante que impede acesso (documentacionAdeudada).

O relatório mais crítico: documentação vencida ('Pendiente'), nunca enviada
('Faltante') ou com data futura de vencimento. Falta de documento em
colaborador ativo pode bloquear o pagamento da IN HAUS.

Formato real do arquivo: HTML disfarçado de .xls.
Colunas: Planta | Contratante | Detalhe | Identificador | Alta | Baixas |
         Documento | Vencimento | Quantidade

Coluna 'Vencimento' tem 3 padrões:
- 'Pendiente'  → já venceu e não foi renovado
- 'Faltante'   → nunca foi enviado
- data futura  → vai vencer

Coluna 'Baixas' preenchida = colaborador já desligado (ignorar pendência).
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from .util import ler_tabela_html, normalizar_cpf


@dataclass
class Faltante:
    cpf: str
    nome: str
    documento: str
    situacao: str              # 'vencido' | 'nunca_enviado' | 'a_vencer'
    vencimento: date | None    # preenchido só quando situacao == 'a_vencer'
    data_baixa: str            # vazio se ainda ativo


def _parse_data(texto: str) -> date | None:
    texto = texto.strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%m-%Y"):
        try:
            return datetime.strptime(texto, fmt).date()
        except ValueError:
            continue
    return None


def ler(caminho: str | Path) -> list[Faltante]:
    linhas = ler_tabela_html(Path(caminho))
    if not linhas:
        return []

    resultado = []
    for celulas in linhas[1:]:  # pula cabeçalho
        if len(celulas) < 9:
            continue
        _, _, detalhe, identificador, _alta, baixas, documento, vencimento, _ = celulas[:9]

        venc = vencimento.strip()
        if venc.lower() == "pendiente":
            situacao, data = "vencido", None
        elif venc.lower() == "faltante":
            situacao, data = "nunca_enviado", None
        else:
            situacao, data = "a_vencer", _parse_data(venc)

        resultado.append(
            Faltante(
                cpf=normalizar_cpf(identificador),
                nome=detalhe.strip(),
                documento=documento.strip(),
                situacao=situacao,
                vencimento=data,
                data_baixa=baixas.strip(),
            )
        )
    return resultado
