"""Leitura e escrita da Matriz (BASE CERTRONIC_GPS.xlsx).

Regras invioláveis deste módulo:
- Colunas A-I (cadastro) NUNCA são alteradas.
- Células contendo RECICLAR NUNCA são alteradas (nem para, nem a partir de) —
  decisão 100% manual da operação, incluindo o timing.
- A escrita preserva as cores que fazem parte da linguagem da operação:
  verde = OK, laranja = protocolo em análise, vermelho = rejeitado/pendente.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string

from .leitores.util import normalizar_cpf

COR_VERDE = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
COR_LARANJA = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
COR_VERMELHO = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

CORES = {"verde": COR_VERDE, "laranja": COR_LARANJA, "vermelho": COR_VERMELHO}


@dataclass
class Alteracao:
    """Uma mudança proposta em uma célula da Matriz (ainda não aplicada)."""

    linha: int
    coluna: int
    cpf: str
    nome: str
    documento: str          # cabeçalho da coluna na Matriz
    valor_atual: str
    valor_novo: str
    cor: str                # 'verde' | 'laranja' | 'vermelho'
    motivo: str             # explicação legível para a prévia/relatório
    comentario: str = ""    # observações do Certronic (ex: AceptadaParcial)


class Matriz:
    def __init__(self, caminho: str | Path, cfg: dict):
        self.caminho = Path(caminho)
        self.cfg = cfg
        self.wb = openpyxl.load_workbook(self.caminho)
        if cfg["aba"] not in self.wb.sheetnames:
            raise ValueError(
                f"A aba '{cfg['aba']}' não existe nesta planilha.\n"
                f"Abas encontradas: {', '.join(self.wb.sheetnames)}.\n"
                "Confira o arquivo da Matriz apontado na aba Configurações."
            )
        self.ws = self.wb[cfg["aba"]]

        self._col_ini = column_index_from_string(cfg["primeira_coluna_documentos"])
        self._col_fim = column_index_from_string(cfg["ultima_coluna_documentos"])
        self._col_cpf = column_index_from_string(cfg["coluna_cpf"])
        self._col_nome = column_index_from_string(cfg["coluna_nome"])
        self._linha_cab = cfg["linha_cabecalho"]

        self._indexar()

    def _indexar(self):
        """Monta os índices CPF→linha e cabeçalho→coluna.

        A Matriz costuma declarar centenas de milhares de linhas "fantasma"
        (formatação residual sem dados). Depois de 100 linhas seguidas sem
        CPF, assumimos que os dados acabaram — varrer o resto custaria ~20s.
        """
        self.linha_por_cpf: dict[str, int] = {}
        vazias_seguidas = 0
        for linha in range(self._linha_cab + 1, self.ws.max_row + 1):
            cpf = self.ws.cell(row=linha, column=self._col_cpf).value
            if cpf is None:
                vazias_seguidas += 1
                if vazias_seguidas >= 100:
                    break
                continue
            vazias_seguidas = 0
            self.linha_por_cpf[normalizar_cpf(cpf)] = linha

        self.coluna_por_documento: dict[str, int] = {}
        for col in range(self._col_ini, self._col_fim + 1):
            cabecalho = self.ws.cell(row=self._linha_cab, column=col).value
            if cabecalho:
                self.coluna_por_documento[str(cabecalho).strip()] = col

    def validar_estrutura(self, documentos_de_para: dict) -> tuple[list[str], list[str]]:
        """Confere a estrutura da Matriz reaproveitando os índices já montados
        no __init__ (nenhuma leitura extra de arquivo — custo desprezível).

        Retorna (problemas_criticos, colunas_de_documento_faltando):
        - críticos impedem uma análise com sentido (sem colaboradores ou sem
          nenhuma coluna de documento);
        - faltando são colunas do de-para que não têm cabeçalho correspondente
          na Matriz (os documentos ligados a elas seriam ignorados).
        """
        criticos: list[str] = []
        if not self.linha_por_cpf:
            criticos.append(
                f"Nenhum colaborador lido na coluna de CPF ('{self.cfg['coluna_cpf']}'). "
                "A aba ou as colunas de cadastro podem estar diferentes do esperado."
            )
        if not self.coluna_por_documento:
            criticos.append(
                "Nenhuma coluna de documento encontrada no cabeçalho (faixa "
                f"{self.cfg['primeira_coluna_documentos']}-{self.cfg['ultima_coluna_documentos']})."
            )

        faltando: list[str] = []
        for destino in dict.fromkeys(documentos_de_para.values()):  # únicos, em ordem
            if self.localizar_coluna(destino) is None:
                faltando.append(destino)
        return criticos, faltando

    def localizar_coluna(self, inicio_cabecalho: str) -> int | None:
        """Encontra a coluna cujo cabeçalho começa com o texto do de-para."""
        alvo = inicio_cabecalho.strip().lower()
        for cabecalho, col in self.coluna_por_documento.items():
            if cabecalho.lower().startswith(alvo):
                return col
        return None

    def valor_celula(self, linha: int, coluna: int) -> str:
        v = self.ws.cell(row=linha, column=coluna).value
        return str(v).strip() if v is not None else ""

    def nome_colaborador(self, linha: int) -> str:
        v = self.ws.cell(row=linha, column=self._col_nome).value
        return str(v).strip() if v is not None else ""

    def pode_alterar(self, linha: int, coluna: int) -> bool:
        """Guarda-chuva das regras invioláveis."""
        if not (self._col_ini <= coluna <= self._col_fim):
            return False  # fora da faixa J-Y: cadastro é intocável
        if "RECICLAR" in self.valor_celula(linha, coluna).upper():
            return False  # RECICLAR é manual, sempre
        return True

    def aplicar(self, alteracoes: list[Alteracao]):
        """Escreve as alterações já confirmadas pela operação."""
        for alt in alteracoes:
            if not self.pode_alterar(alt.linha, alt.coluna):
                continue
            celula = self.ws.cell(row=alt.linha, column=alt.coluna)
            celula.value = alt.valor_novo
            celula.fill = CORES[alt.cor]

    def salvar(self):
        self.wb.save(self.caminho)
