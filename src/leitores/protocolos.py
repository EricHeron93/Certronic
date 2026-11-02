"""documentacao_enviada — Controle de Documentação por Estado (documentacionEnviada).

Todo documento enviado ao Certronic vira um protocolo (Id) com um Estado.
Este relatório NÃO tem CPF/nome do colaborador: o vínculo é feito pelo nº
do protocolo que a operação colou na célula da Matriz.

Formato real do arquivo: .xls que é um .xlsx (zip OOXML).
Colunas: Id | Planta | Contratante | Documento | Período | Enviado |
         Acusado | Observações | Comentário | Estado | Verificados | Total
"""

from __future__ import annotations

import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path

import openpyxl

from .util import detectar_formato


@dataclass
class Protocolo:
    id: int
    documento: str
    periodo: str
    enviado: str
    acusado: str
    observacoes: str
    comentario: str
    estado: str  # Aceptada | Rechazada | AceptadaParcial | Pendiente de Control


def ler(caminho: str | Path) -> dict[int, Protocolo]:
    """Lê o relatório e devolve um dicionário {id_protocolo: Protocolo}."""
    caminho = Path(caminho)
    formato = detectar_formato(caminho)
    if formato != "xlsx":
        raise ValueError(
            f"Formato inesperado em {caminho.name}: esperado xlsx (zip), obtido {formato}"
        )

    # openpyxl recusa a extensão .xls mesmo sendo xlsx por dentro —
    # copiamos para um temporário .xlsx antes de abrir.
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = Path(tmp.name)
    try:
        shutil.copyfile(caminho, tmp_path)
        wb = openpyxl.load_workbook(tmp_path, data_only=True, read_only=True)
        ws = wb.active

        protocolos: dict[int, Protocolo] = {}
        for linha in ws.iter_rows(min_row=2, values_only=True):
            if linha[0] is None:
                continue
            p = Protocolo(
                id=int(linha[0]),
                documento=str(linha[3] or "").strip(),
                periodo=str(linha[4] or "").strip(),
                enviado=str(linha[5] or "").strip(),
                acusado=str(linha[6] or "").strip(),
                observacoes=str(linha[7] or "").strip(),
                comentario=str(linha[8] or "").strip(),
                estado=str(linha[9] or "").strip(),
            )
            protocolos[p.id] = p
        wb.close()
        return protocolos
    finally:
        tmp_path.unlink(missing_ok=True)
