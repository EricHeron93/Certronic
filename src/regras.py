"""Motor de regras: cruza os 4 relatórios e propõe alterações na Matriz.

Nada aqui escreve na planilha — a saída é uma lista de Alteracao (prévia)
e uma lista de Pendencia/CasoIncerto para o relatório. A escrita só
acontece depois que a operação confirma na GUI.

Ciclo de vida de uma célula colaborador × documento:
  X (vermelho) → PACK → protocolo (laranja) → OK (verde)
                                  └→ [ID] X (vermelho) → correção → protocolo...
  RECICLAR: fora do escopo da automação (manual, inclusive o timing).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date

from .leitores.ativos import Colaborador
from .leitores.faltantes import Faltante
from .leitores.protocolos import Protocolo
from .leitores.vencimentos import Vencimento
from .matriz import Alteracao, Matriz

# célula que contém apenas um nº de protocolo, com ou sem "X" de rejeição
# (convenção da operação: X colado na frente do número, ex: "X814525")
RE_PROTOCOLO = re.compile(r"^\s*X?\s*(\d{4,})\s*X?\s*$", re.IGNORECASE)


@dataclass
class Pendencia:
    """Item crítico para o relatório (não vira escrita automática)."""

    cpf: str
    nome: str
    documento: str
    descricao: str


@dataclass
class CasoIncerto:
    """Algo que o programa não soube classificar — pulado e listado."""

    origem: str      # qual planilha gerou o caso
    referencia: str  # CPF, protocolo ou documento envolvido
    motivo: str


@dataclass
class ResultadoAnalise:
    alteracoes: list[Alteracao] = field(default_factory=list)
    pendencias: list[Pendencia] = field(default_factory=list)
    incertos: list[CasoIncerto] = field(default_factory=list)


def _rotulo_vencimento(venc: date, hoje: date, proto: str = "") -> str:
    """Reproduz a convenção da operação: OK (X MESES) / OK (V:data).

    `proto` é o número do protocolo já aprovado, preservado na célula (fica
    entre o OK e o rótulo de vencimento, ex.: 'OK 12345 (3 MESES)').
    """
    prefixo = f"OK {proto}" if proto else "OK"
    dias = (venc - hoje).days
    if dias >= 45:
        meses = round(dias / 30)
        return f"{prefixo} ({meses} MESES)"
    return f"{prefixo} (V:{venc.strftime('%d/%m/%Y')})"


# nº do protocolo preservado numa célula OK (ex.: "OK 12345" ou "OK 12345 (3 MESES)")
RE_OK_PROTOCOLO = re.compile(r"^OK\s+(\d{4,})\b", re.IGNORECASE)


def analisar(
    matriz: Matriz,
    protocolos: dict[int, Protocolo],
    vencimentos: list[Vencimento],
    faltantes: list[Faltante],
    ativos: dict[str, Colaborador],
    cfg: dict,
    hoje: date | None = None,
) -> ResultadoAnalise:
    hoje = hoje or date.today()
    r = ResultadoAnalise()
    de_para: dict[str, str] = cfg["documentos"]
    ignorados = [d.lower() for d in cfg.get("documentos_ignorados", [])]
    desligados = {v.lower() for v in cfg["regras"]["acesso_desligado"]}

    def eh_ignorado(nome_certronic: str) -> bool:
        """Documento de folha mensal sem coluna na Matriz — fora do escopo."""
        nome = nome_certronic.lower()
        return any(nome.startswith(i) for i in ignorados)

    def mapear_documento(nome_certronic: str) -> int | None:
        """Nome no relatório → coluna na Matriz, via de-para do config."""
        for inicio_relatorio, inicio_matriz in de_para.items():
            if nome_certronic.lower().startswith(inicio_relatorio.lower()):
                return matriz.localizar_coluna(inicio_matriz)
        return None

    # ------------------------------------------------------------------
    # Regra 1 — protocolos: células da Matriz que contêm um nº de protocolo
    # são atualizadas conforme o Estado na documentacao_enviada.
    # ------------------------------------------------------------------
    for cpf, linha in matriz.linha_por_cpf.items():
        nome = matriz.nome_colaborador(linha)
        for cabecalho, col in matriz.coluna_por_documento.items():
            if not matriz.pode_alterar(linha, col):
                continue
            valor = matriz.valor_celula(linha, col)
            m = RE_PROTOCOLO.match(valor)
            if not m:
                continue
            pid = int(m.group(1))
            proto = protocolos.get(pid)
            if proto is None:
                continue  # protocolo fora do range de datas exportado: nada a fazer

            estado = proto.estado.strip().lower()
            if estado == "aceptada":
                # Preserva o nº do protocolo na célula (decisão da operação em
                # 30/07/2026): o protocolo aprovado vira "OK 12345", não só "OK".
                r.alteracoes.append(Alteracao(
                    linha=linha, coluna=col, cpf=cpf, nome=nome,
                    documento=cabecalho, valor_atual=valor, valor_novo=f"OK {pid}",
                    cor="verde",
                    motivo=f"Protocolo {pid} aprovado pelo cliente (Aceptada)",
                ))
            elif estado == "rechazada":
                novo = f"X{pid}"
                if valor.upper().replace(" ", "") != novo.replace(" ", ""):
                    r.alteracoes.append(Alteracao(
                        linha=linha, coluna=col, cpf=cpf, nome=nome,
                        documento=cabecalho, valor_atual=valor, valor_novo=novo,
                        cor="vermelho",
                        motivo=f"Protocolo {pid} reprovado (Rechazada) — corrigir",
                        comentario=proto.observacoes or proto.comentario,
                    ))
            elif estado == "aceptadaparcial":
                # Aceito com ressalva: célula vira "RESSALVA 12345" (formato da
                # operação, decisão de 30/07/2026), preservando o nº do protocolo.
                r.alteracoes.append(Alteracao(
                    linha=linha, coluna=col, cpf=cpf, nome=nome,
                    documento=cabecalho, valor_atual=valor,
                    valor_novo=f"RESSALVA {pid}", cor="verde",
                    motivo=f"Protocolo {pid} aceito parcialmente — ver observação",
                    comentario=proto.observacoes or proto.comentario,
                ))
            # 'Pendiente de Control': célula já está certa (nº em laranja) — sem ação

    # ------------------------------------------------------------------
    # Regra 2 — vencimentos: documento OK com data próxima vira OK (X MESES)
    #
    # Duas proteções contra o histórico da estado_vencimento:
    # a) Se houver mais de um registro do mesmo CPF+documento (histórico),
    #    só o de data MAIS RECENTE vale — registros antigos são descartados.
    # b) "VENCIDO" só é gravado na célula se a documentacao_faltante corroborar
    #    (documento cobrado lá). Vencido lá mas ausente da documentacao_faltante indica
    #    que provavelmente existe um documento mais novo vigente — nesse
    #    caso não alarmamos: vira caso a verificar no relatório.
    # ------------------------------------------------------------------
    mais_recente: dict[tuple[str, int], Vencimento] = {}
    for v in vencimentos:
        if not v.eh_pessoa_fisica or v.vencimento is None or eh_ignorado(v.documento):
            continue
        col = mapear_documento(v.documento)
        linha = matriz.linha_por_cpf.get(v.cpf)
        if col is None:
            r.incertos.append(CasoIncerto(
                "vencimentos", v.documento,
                "Documento sem correspondência no de-para do config.yaml",
            ))
            continue
        if linha is None:
            r.incertos.append(CasoIncerto(
                "vencimentos", f"{v.nome} ({v.cpf})",
                "CPF não encontrado na Matriz",
            ))
            continue
        chave = (v.cpf, col)
        atual = mais_recente.get(chave)
        if atual is None or v.vencimento > atual.vencimento:
            mais_recente[chave] = v

    # pares (cpf, coluna) que a documentacao_faltante está cobrando (vencido ou com data)
    cobrados_p3: set[tuple[str, int]] = set()
    for f in faltantes:
        col = mapear_documento(f.documento)
        if col is not None:
            cobrados_p3.add((f.cpf, col))

    for (cpf, col), v in mais_recente.items():
        linha = matriz.linha_por_cpf[cpf]
        if not matriz.pode_alterar(linha, col):
            continue

        valor = matriz.valor_celula(linha, col)
        if not valor.upper().startswith("OK"):
            continue  # só anota vencimento em documento que está OK

        if v.vencimento < hoje:
            data_fmt = v.vencimento.strftime("%d/%m/%Y")

            if (cpf, col) not in cobrados_p3:
                # Vencido no estado_vencimento mas a documentacao_faltante não está cobrando: indício de que
                # existe documento mais novo vigente (o estado_vencimento guarda histórico).
                # Não gravamos VENCIDO — vai como caso a verificar.
                r.incertos.append(CasoIncerto(
                    "vencimentos", f"{v.nome} — {v.documento}",
                    f"Consta vencido em {data_fmt} no estado_vencimento, mas a documentacao_faltante não cobra "
                    "este documento — provável renovação vigente. Verificar.",
                ))
                continue

            # Data já passou E a documentacao_faltante confirma a cobrança:
            # quem grava VENCIDO/PENDENTE/FALTANTE na célula é a Regra 3
            # (documentacao_faltante é a fonte da verdade do que está devido).
            continue

        # Preserva o nº do protocolo se a célula já o carrega ("OK 12345" →
        # "OK 12345 (3 MESES)"), para não perder a rastreabilidade.
        m_proto = RE_OK_PROTOCOLO.match(valor.strip())
        novo = _rotulo_vencimento(v.vencimento, hoje, m_proto.group(1) if m_proto else "")
        if valor.strip().upper() != novo.upper():
            r.alteracoes.append(Alteracao(
                linha=linha, coluna=col, cpf=v.cpf, nome=v.nome,
                documento=v.documento, valor_atual=valor, valor_novo=novo,
                cor="verde",
                motivo=f"Vence em {v.vencimento.strftime('%d/%m/%Y')} (estado_vencimento)",
            ))

    # ------------------------------------------------------------------
    # Regra 3 — faltantes: a documentacao_faltante é o alarme oficial e
    # AGORA TAMBÉM ESCREVE na Matriz (tudo em vermelho):
    #   data < hoje   → "VENCIDO (data)"
    #   'Pendiente'   → "PENDENTE"
    #   'Faltante'    → "FALTANTE"
    #   data futura   → não mexe (o aviso OK (V:...) vem do estado_vencimento)
    # Só para colaborador ativo e sem protocolo em análise cobrindo o
    # documento (correção já correndo não é alarmada nem sobrescrita).
    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Checagem de cadastro: todo colaborador ATIVO na lista_pessoal deve
    # existir na Matriz. Quem não existe vira caso a resolver (cadastrar),
    # independentemente de aparecer ou não nos outros relatórios.
    # ------------------------------------------------------------------
    for cpf, colab in ativos.items():
        if colab.acesso.lower() in desligados:
            continue
        if cpf not in matriz.linha_por_cpf:
            r.incertos.append(CasoIncerto(
                "lista_pessoal", f"{colab.nome_completo} ({cpf})",
                f"Colaborador ativo no Certronic (acesso: {colab.acesso}) "
                "não cadastrado na Matriz — cadastrar nas colunas A-I.",
            ))

    for f in faltantes:
        if f.data_baixa:
            continue  # já desligado — sem pendência
        if eh_ignorado(f.documento):
            continue
        colab = ativos.get(f.cpf)
        if colab and colab.acesso.lower() in desligados:
            continue

        # Classifica pela coluna "Vencimento" da documentacao_faltante.
        # 'a_vencer' com data já passada também é vencido — a data pode ter
        # ficado para trás entre o export e a execução da ferramenta.
        if f.situacao == "nunca_enviado":
            novo, rotulo = "FALTANTE", "FALTANTE — nunca enviado"
        elif f.situacao == "vencido" and f.vencimento is None:
            novo, rotulo = "PENDENTE", "PENDENTE — documento vencido"
        elif f.vencimento is not None and f.vencimento < hoje:
            data_fmt = f.vencimento.strftime("%d/%m/%Y")
            novo, rotulo = f"VENCIDO ({data_fmt})", f"VENCIDO em {data_fmt}"
        else:
            continue  # data futura: sem ação aqui

        linha = matriz.linha_por_cpf.get(f.cpf)
        col = mapear_documento(f.documento)
        valor = matriz.valor_celula(linha, col) if (linha is not None and col is not None) else ""
        valor_upper = valor.strip().upper()

        # "Correção em andamento" é decidido POR PESSOA, não pelo tipo de
        # documento: só suprime o alarme quando a PRÓPRIA célula do colaborador
        # guarda o número do protocolo (laranja). Um protocolo do mesmo
        # documento no nome de OUTRA pessoa não esconde o vencimento deste
        # colaborador — antes, uma checagem global por nome escondia centenas
        # de vencidos reais (célula OK/PACK/X) só porque alguém tinha protocolo
        # daquele tipo em análise.
        if valor and RE_PROTOCOLO.match(valor):
            continue  # protocolo próprio em análise — não alarma nem sobrescreve

        # Células em controle manual da operação NÃO são sobrescritas, mas o
        # vencimento ENTRA na pendência com o motivo explícito de não ter
        # alterado (decisão da operação em 23/07/2026 — antes o PACK era
        # sobrescrito; deixou de ser):
        #   PACK     → colaborador em fase de assinatura dos certificados
        #   RECICLAR → reciclagem/retreinamento, timing 100% manual
        if "PACK" in valor_upper:
            r.pendencias.append(Pendencia(
                cpf=f.cpf, nome=f.nome, documento=f.documento,
                descricao=(f"{rotulo}. Alteração NÃO feita: célula em PACK "
                           "(colaborador em fase de assinatura de certificados) "
                           "— etapa manual da operação. Risco de bloqueio."),
            ))
            continue
        if "RECICL" in valor_upper:  # RECICLAR / RECICLAGEM
            r.pendencias.append(Pendencia(
                cpf=f.cpf, nome=f.nome, documento=f.documento,
                descricao=(f"{rotulo}. Alteração NÃO feita: célula em RECICLAR "
                           "(reciclagem manual da operação). Risco de bloqueio."),
            ))
            continue

        r.pendencias.append(Pendencia(
            cpf=f.cpf, nome=f.nome, documento=f.documento,
            descricao=f"{rotulo} — sem protocolo em análise. Risco de bloqueio.",
        ))

        # --- escrita na célula da Matriz ---------------------------------
        # (N/A vira divergência sem tocar a célula; nº de protocolo fora do
        #  padrão é preservado. PACK e RECICLAR já foram tratados acima.)
        if linha is None or col is None:
            continue  # CPF/documento fora da Matriz: fica só na pendência
        if not matriz.pode_alterar(linha, col):
            continue

        # Célula "X" (não iniciado): permanece X quando o Certronic apenas cobra
        # sem data (PENDENTE/FALTANTE). Só vira VENCIDO quando há data de
        # vencimento concreta (decisão da operação em 28/07/2026). A pendência
        # já foi registrada acima, então o risco continua visível na lista.
        if valor_upper == "X" and not novo.upper().startswith("VENCIDO"):
            continue

        if valor_upper in ("N/A", "NA"):
            # Matriz diz "não se aplica", Certronic está cobrando: divergência
            # de cadastro — a operação decide qual lado corrigir.
            r.incertos.append(CasoIncerto(
                "faltantes", f"{f.nome} — {f.documento}",
                "Matriz marca N/A mas o Certronic cobra o documento "
                "(documentacao_faltante). Verificar cadastro.",
            ))
            continue

        if re.search(r"\d{5,}", valor) and not valor_upper.startswith(("OK", "VENCIDO")):
            # nº de protocolo com anotação fora do padrão (ex: "RECUSADO 678287"):
            # correção em andamento — preservar o número.
            continue
        if valor_upper == novo.upper():
            continue  # já está correto

        # a documentacao_faltante tem precedência sobre proposta anterior
        # (ex: rótulo de vencimento da Regra 2) para a mesma célula
        r.alteracoes = [
            a for a in r.alteracoes if not (a.linha == linha and a.coluna == col)
        ]
        r.alteracoes.append(Alteracao(
            linha=linha, coluna=col, cpf=f.cpf, nome=f.nome,
            documento=f.documento, valor_atual=valor, valor_novo=novo,
            cor="vermelho",
            motivo=f"{rotulo} (documentacao_faltante)",
        ))

    return r
