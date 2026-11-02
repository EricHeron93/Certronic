"""Relatório de execução: registro auditável do que mudou e por quê.

Gera um .xlsx por execução na pasta de relatórios, com 4 abas:
- Resumo               : cabeçalho da execução e totais por seção
- Alterações           : tudo que foi (ou será) escrito na Matriz
- Pendências críticas  : documentos vencidos/faltantes que exigem ação
- Casos pulados        : o que o programa não soube classificar
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .regras import ResultadoAnalise

AZUL_ESCURO = "1F4E78"
AZUL_CLARO = "DDEBF7"
CINZA = "F2F2F2"
VERDE = "C6EFCE"
AMARELO = "FFEB9C"

FONTE_TITULO = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
FONTE_CABECALHO = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONTE_ROTULO = Font(name="Calibri", size=11, bold=True)
FILL_TITULO = PatternFill(start_color=AZUL_ESCURO, end_color=AZUL_ESCURO, fill_type="solid")
FILL_CABECALHO = PatternFill(start_color=AZUL_ESCURO, end_color=AZUL_ESCURO, fill_type="solid")
FILL_ZEBRA = PatternFill(start_color=CINZA, end_color=CINZA, fill_type="solid")
BORDA_FINA = Border(*[Side(style="thin", color="BFBFBF")] * 4)


def _montar_aba_dados(wb, nome: str, cabecalho: list[str], linhas: list[list],
                      larguras: list[int]):
    """Cria uma aba de dados com cabeçalho formatado, zebra, filtro e congelamento."""
    ws = wb.create_sheet(nome)

    for i, titulo in enumerate(cabecalho, start=1):
        celula = ws.cell(row=1, column=i, value=titulo)
        celula.font = FONTE_CABECALHO
        celula.fill = FILL_CABECALHO
        celula.alignment = Alignment(horizontal="center", vertical="center")
        celula.border = BORDA_FINA
        ws.column_dimensions[get_column_letter(i)].width = larguras[i - 1]

    for num, linha in enumerate(linhas, start=2):
        for i, valor in enumerate(linha, start=1):
            celula = ws.cell(row=num, column=i, value=valor)
            celula.border = BORDA_FINA
            celula.alignment = Alignment(vertical="top", wrap_text=(larguras[i - 1] >= 40))
            if num % 2 == 0:
                celula.fill = FILL_ZEBRA

    ws.freeze_panes = "A2"
    if linhas:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cabecalho))}{len(linhas) + 1}"
    return ws


def gerar(resultado: ResultadoAnalise, pasta: str | Path, aplicado: bool) -> Path:
    pasta = Path(pasta)
    pasta.mkdir(parents=True, exist_ok=True)
    agora = datetime.now()
    caminho = pasta / f"atualizacao_{agora.strftime('%Y%m%d_%H%M%S')}.xlsx"

    wb = openpyxl.Workbook()

    # ------------------------------------------------------------ Resumo
    ws = wb.active
    ws.title = "Resumo"
    ws.merge_cells("A1:D2")
    titulo = ws["A1"]
    titulo.value = "ATUALIZAÇÃO DA MATRIZ CERTRONIC — IN HAUS"
    titulo.font = FONTE_TITULO
    titulo.fill = FILL_TITULO
    titulo.alignment = Alignment(horizontal="center", vertical="center")

    status_txt = "APLICADA NA MATRIZ" if aplicado else "PRÉVIA (nada foi gravado)"
    status_fill = PatternFill(start_color=VERDE if aplicado else AMARELO,
                              end_color=VERDE if aplicado else AMARELO, fill_type="solid")

    resumo = [
        ("Data e hora", agora.strftime("%d/%m/%Y %H:%M")),
        ("Status", status_txt),
        ("Alterações na Matriz", len(resultado.alteracoes)),
        ("Pendências críticas", len(resultado.pendencias)),
        ("Casos pulados / a verificar", len(resultado.incertos)),
    ]
    for i, (rotulo, valor) in enumerate(resumo, start=4):
        ws.cell(row=i, column=1, value=rotulo).font = FONTE_ROTULO
        ws.cell(row=i, column=1).border = BORDA_FINA
        ws.cell(row=i, column=1).fill = PatternFill(
            start_color=AZUL_CLARO, end_color=AZUL_CLARO, fill_type="solid")
        celula_valor = ws.cell(row=i, column=2, value=valor)
        celula_valor.border = BORDA_FINA
        if rotulo == "Status":
            celula_valor.fill = status_fill
            celula_valor.font = FONTE_ROTULO
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 35
    ws.cell(row=10, column=1, value="Cada aba deste arquivo traz o detalhe de uma seção. "
            "Use os filtros no cabeçalho para localizar um colaborador ou documento.")

    # -------------------------------------------------------- Alterações
    _montar_aba_dados(
        wb, "Alterações",
        ["Colaborador", "CPF", "Documento", "De", "Para", "Motivo", "Observação Certronic"],
        [[a.nome, a.cpf, a.documento, a.valor_atual, a.valor_novo, a.motivo, a.comentario]
         for a in resultado.alteracoes],
        [32, 14, 30, 16, 18, 45, 40],
    )

    # ------------------------------------------------ Pendências críticas
    _montar_aba_dados(
        wb, "Pendências críticas",
        ["Colaborador", "CPF", "Documento", "Situação"],
        [[p.nome, p.cpf, p.documento, p.descricao] for p in resultado.pendencias],
        [32, 14, 40, 60],
    )

    # --------------------------------------------------- Casos pulados
    _montar_aba_dados(
        wb, "Casos pulados",
        ["Origem", "Referência", "Motivo"],
        [[c.origem, c.referencia, c.motivo] for c in resultado.incertos],
        [16, 45, 70],
    )

    wb.save(caminho)
    return caminho
