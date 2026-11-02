"""Base estruturada de indicadores para o Power BI.

Gera/atualiza o base_indicadores.xlsx a cada análise, com 5 tabelas:

Dimensões (sobrescritas a cada execução — retrato do cadastro):
- colaboradores : cadastro da Matriz (A-I) + situação de acesso (lista_pessoal)
- documentos    : de-para de tipos de documento (config.yaml)

Fatos (históricos — acumulam por data de execução):
- protocolos      : todos os protocolos da documentacao_enviada, com o CPF descoberto
                    pelo cruzamento com a Matriz (upsert por id_protocolo)
- situacao_atual  : foto colaborador × documento da Matriz (acumula por data)
- pendencias      : pendências críticas apontadas pelo motor (acumula por data)

Rodar duas vezes no mesmo dia não duplica: as linhas do dia são substituídas.
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path

import openpyxl

from .leitores.ativos import Colaborador
from .leitores.protocolos import Protocolo
from .leitores.vencimentos import Vencimento
from .matriz import Matriz
from .regras import RE_OK_PROTOCOLO, RE_PROTOCOLO, ResultadoAnalise

ABAS = {
    "colaboradores": ["cpf", "nome", "funcao", "cr", "regional", "sitio",
                      "status_docs", "admissao", "status_certronic", "acesso"],
    "documentos": ["nome_certronic", "coluna_matriz"],
    "protocolos": ["id_protocolo", "cpf", "documento", "estado",
                   "data_enviado", "data_acusado", "observacoes", "comentario"],
    "situacao_atual": ["data_execucao", "cpf", "nome", "documento",
                       "status_celula", "data_vencimento"],
    "pendencias": ["data_execucao", "cpf", "nome", "documento", "descricao"],
}


def _classificar_celula(valor: str) -> str:
    """Normaliza o conteúdo da célula da Matriz em uma categoria estável."""
    v = valor.strip().upper()
    if not v:
        return "VAZIO"
    if "RESSALVA" in v:
        return "OK_RESSALVA"  # AceptadaParcial ("RESSALVA 12345") — conforme com ressalva
    if v.startswith("OK"):
        return "OK"
    if "RECICLAR" in v:
        return "RECICLAR"
    if v.startswith("VENCIDO"):
        return "VENCIDO"
    if v == "PENDENTE":
        return "PENDENTE"
    if v == "FALTANTE":
        return "FALTANTE"
    if v in ("N/A", "NA"):
        return "N/A"
    if "PACK" in v:
        return "PACK"
    if RE_PROTOCOLO.match(valor):
        return "REJEITADO" if "X" in v else "PROTOCOLO_EM_ANALISE"
    if v == "X":
        return "NAO_INICIADO"
    return "OUTRO"


def _reescrever_aba(wb, nome: str, linhas: list[list], cabecalho: list[str] | None = None):
    """Substitui a aba inteira pelo cabeçalho + linhas fornecidas."""
    if nome in wb.sheetnames:
        del wb[nome]
    ws = wb.create_sheet(nome)
    ws.append(cabecalho if cabecalho is not None else ABAS[nome])
    for linha in linhas:
        ws.append(linha)


def _linhas_existentes(wb, nome: str, excluir_data: str = "", chave_upsert: int | None = None,
                       chaves_novas: set | None = None) -> list[list]:
    """Lê as linhas atuais de uma aba, filtrando o que será substituído."""
    if nome not in wb.sheetnames:
        return []
    linhas = []
    for row in wb[nome].iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        if excluir_data and str(row[0]) == excluir_data:
            continue  # linhas da mesma data de execução serão regravadas
        if chave_upsert is not None and chaves_novas and row[chave_upsert] in chaves_novas:
            continue  # protocolo será regravado com estado atualizado
        linhas.append(list(row))
    return linhas


def gerar(
    caminho: str | Path,
    matriz: Matriz,
    protocolos: dict[int, Protocolo],
    vencimentos: list[Vencimento],
    ativos: dict[str, Colaborador],
    resultado: ResultadoAnalise,
    cfg: dict,
    hoje: date | None = None,
) -> Path:
    caminho = Path(caminho)
    hoje = hoje or date.today()
    data_exec = hoje.strftime("%Y-%m-%d")

    if caminho.exists():
        try:
            wb = openpyxl.load_workbook(caminho)
        except Exception:
            # Arquivo corrompido ou que não é um xlsx de verdade (ex: criado
            # vazio pelo Windows): recomeça a base do zero — o histórico
            # anterior desse arquivo inválido não existe de qualquer forma.
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

    # --- dimensão: colaboradores (Matriz A-I + acesso da lista_pessoal) -----
    linhas_colab = []
    for cpf, linha in matriz.linha_por_cpf.items():
        cadastro = [matriz.ws.cell(row=linha, column=c).value for c in range(1, 10)]
        acesso = ativos[cpf].acesso if cpf in ativos else ""
        # substitui o CPF bruto pelo normalizado e serializa datas
        valores = [cpf] + [
            v.strftime("%d/%m/%Y") if hasattr(v, "strftime") else (str(v).strip() if v is not None else "")
            for v in cadastro[0:1] + cadastro[2:9]
        ] + [acesso]
        # ordem final: cpf, nome, funcao, cr, regional, sitio, status_docs, admissao, status_certronic, acesso
        linhas_colab.append(valores)
    _reescrever_aba(wb, "colaboradores", linhas_colab)

    # --- dimensão: documentos (de-para do config) --------------------------
    linhas_docs = [[nome, coluna] for nome, coluna in cfg["documentos"].items()]
    _reescrever_aba(wb, "documentos", linhas_docs)

    # --- fato: protocolos (upsert por id, CPF via cruzamento com a Matriz) -
    cpf_por_protocolo: dict[int, str] = {}
    for cpf, linha in matriz.linha_por_cpf.items():
        for col in matriz.coluna_por_documento.values():
            valor = matriz.valor_celula(linha, col)
            # nº cru (protocolo em análise) OU nº preservado no OK ("OK 12345")
            m = RE_PROTOCOLO.match(valor) or RE_OK_PROTOCOLO.match(valor)
            if m:
                cpf_por_protocolo[int(m.group(1))] = cpf

    # CPFs descobertos em execuções anteriores são preservados: a célula da
    # Matriz vira "OK" depois da aprovação e o número some, mas o vínculo
    # protocolo→colaborador já registrado na base não pode se perder.
    cpf_ja_conhecido: dict = {}
    if "protocolos" in wb.sheetnames:
        for row in wb["protocolos"].iter_rows(min_row=2, values_only=True):
            if row[0] is not None and row[1]:
                cpf_ja_conhecido[row[0]] = str(row[1])

    novas_proto = [
        [p.id, cpf_por_protocolo.get(p.id) or cpf_ja_conhecido.get(p.id, ""),
         p.documento, p.estado,
         p.enviado, p.acusado, p.observacoes, p.comentario]
        for p in protocolos.values()
    ]
    antigas = _linhas_existentes(wb, "protocolos", chave_upsert=0,
                                 chaves_novas={p.id for p in protocolos.values()})
    _reescrever_aba(wb, "protocolos", antigas + novas_proto)

    # --- fato: situação atual (foto da Matriz, acumula por data) -----------
    venc_por_chave: dict[tuple[str, int], date] = {}
    de_para = cfg["documentos"]
    for v in vencimentos:
        if not v.eh_pessoa_fisica or v.vencimento is None:
            continue
        for inicio_rel, inicio_mat in de_para.items():
            if v.documento.lower().startswith(inicio_rel.lower()):
                col = matriz.localizar_coluna(inicio_mat)
                if col:
                    venc_por_chave[(v.cpf, col)] = v.vencimento
                break

    novas_sit = []
    for cpf, linha in matriz.linha_por_cpf.items():
        nome = matriz.nome_colaborador(linha)
        for cabecalho, col in matriz.coluna_por_documento.items():
            valor = matriz.valor_celula(linha, col)
            venc = venc_por_chave.get((cpf, col))
            novas_sit.append([
                data_exec, cpf, nome, cabecalho,
                _classificar_celula(valor),
                venc.strftime("%d/%m/%Y") if venc else "",
            ])
    antigas = _linhas_existentes(wb, "situacao_atual", excluir_data=data_exec)
    _reescrever_aba(wb, "situacao_atual", antigas + novas_sit)

    # --- visão por colaborador (foto larga, espelho da Matriz) -------------
    # Uma linha por colaborador, uma coluna por documento (como a tela
    # "Doc. Pessoal" do Certronic), com o valor da célula da Matriz e o
    # vencimento anotado quando conhecido. Sobrescrita a cada execução.
    docs_ordenados = sorted(matriz.coluna_por_documento.items(), key=lambda kv: kv[1])
    cab_visao = (["cpf", "nome", "acesso"]
                 + [cab for cab, _ in docs_ordenados]
                 + ["docs_ok", "docs_com_pendencia", "pct_conforme"])

    linhas_visao = []
    for cpf, linha in matriz.linha_por_cpf.items():
        nome = matriz.nome_colaborador(linha)
        acesso = ativos[cpf].acesso if cpf in ativos else ""
        celulas, n_ok, n_pend, n_aplicaveis = [], 0, 0, 0
        for _, col in docs_ordenados:
            valor = matriz.valor_celula(linha, col)
            venc = venc_por_chave.get((cpf, col))
            if venc and valor.strip().upper() == "OK":
                valor = f"OK (V:{venc.strftime('%d/%m/%Y')})"
            celulas.append(valor)
            categoria = _classificar_celula(valor)
            if categoria != "N/A":
                n_aplicaveis += 1
            if categoria in ("OK", "OK_RESSALVA"):  # ressalva conta como conforme
                n_ok += 1
            elif categoria in ("VENCIDO", "PENDENTE", "FALTANTE", "REJEITADO", "NAO_INICIADO"):
                n_pend += 1
        pct = round(100 * n_ok / n_aplicaveis) if n_aplicaveis else 0
        linhas_visao.append([cpf, nome, acesso] + celulas + [n_ok, n_pend, pct])
    _reescrever_aba(wb, "visao_colaborador", linhas_visao, cabecalho=cab_visao)

    # --- fato: pendências críticas (acumula por data) -----------------------
    novas_pend = [
        [data_exec, p.cpf, p.nome, p.documento, p.descricao]
        for p in resultado.pendencias
    ]
    antigas = _linhas_existentes(wb, "pendencias", excluir_data=data_exec)
    _reescrever_aba(wb, "pendencias", antigas + novas_pend)

    wb.save(caminho)
    return caminho
