# -*- coding: utf-8 -*-
"""Gera o Manual da Automação (Word + Markdown) a partir de um conteúdo único.

Uso:  python docs/gerar_manual.py
Saída: C:/Projeto Certronic/Manual Atualizacao Matriz Certronic.docx
       docs/MANUAL.md (cópia em texto, versionada junto com o código)

Quando uma regra mudar, atualize o CONTEUDO abaixo e rode de novo.
"""

from pathlib import Path

from docx import Document
from docx.shared import Pt

VERSAO = "1.4"
DATA = "30/07/2026"

# Cada item: ("h1"|"h2"|"p"|"bullets"|"table", conteúdo)
CONTEUDO = [
    ("h1", "Manual da Automação Certronic → Matriz (BASE CERTRONIC_GPS)"),
    ("p", f"Versão {VERSAO} — {DATA} — IN HAUS / Operação Mercado Livre"),

    ("h2", "1. O que é este programa"),
    ("p", "Este programa atualiza automaticamente a planilha BASE CERTRONIC_GPS "
          "(que chamamos de \"Matriz\"), usando as informações dos 4 relatórios "
          "que exportamos do Certronic. Tudo o que hoje é feito célula por célula, "
          "à mão — conferir protocolo aprovado, marcar vencimento, apontar documento "
          "faltante — o programa faz de uma vez, em segundos."),
    ("p", "Importante: o programa NUNCA grava nada sozinho. Ele primeiro mostra na "
          "tela TUDO o que pretende mudar (a \"prévia\"), e só escreve na Matriz "
          "depois que você clica em Aplicar e confirma."),

    ("h2", "2. O que o programa NUNCA faz (garantias)"),
    ("bullets", [
        "Nunca altera as colunas A até I da Matriz (o cadastro do colaborador: nome, "
        "CPF, função, CR, regional, sítio, status). Essas colunas continuam sendo "
        "preenchidas pela operação, como sempre foi.",
        "Nunca mexe em células com RECICLAR — nem para colocar, nem para tirar. A "
        "decisão de reciclar um treinamento (e QUANDO reciclar) é 100% da operação. "
        "Se o documento estiver vencido, o programa não altera a célula, mas registra "
        "na aba de pendências com o motivo.",
        "Nunca mexe em células com PACK — o colaborador está em fase de assinatura dos "
        "certificados (etapa manual). Se o documento estiver vencido, o programa não "
        "altera a célula, mas registra na aba de pendências deixando claro o motivo.",
        "Nunca grava nada sem a sua confirmação na tela.",
        "Nunca \"chuta\": se encontra algo que não sabe classificar (um CPF que não "
        "está na Matriz, um documento desconhecido), ele pula e lista no relatório "
        "para a operação resolver.",
    ]),

    ("h2", "3. Os 4 relatórios do Certronic"),
    ("p", "Antes de usar o programa, exporte os 4 relatórios do Certronic e salve na "
          "pasta de entrada (C:\\Projeto Certronic\\entrada) com EXATAMENTE estes nomes:"),
    ("table",
     ["Nome do arquivo", "O que é", "Para que o programa usa"],
     [
         ["documentacao_enviada.xls",
          "Lista de todos os protocolos (IDs) de documentos que a operação enviou ao "
          "Certronic, com o estado de cada um (Aceptada, Rechazada, AceptadaParcial, "
          "Pendiente de Control).",
          "Atualizar as células da Matriz que contêm número de protocolo: aprovado "
          "vira OK + número (ex: OK 12345), reprovado vira X+número."],
         ["estado_vencimento.xls",
          "Documentos com prazo de validade vencido ou prestes a vencer, por "
          "colaborador (CPF), com a data.",
          "Avisar na Matriz quando um documento OK está perto de vencer."],
         ["documentacao_faltante.xls",
          "O relatório mais crítico: documentos que o Certronic está COBRANDO — "
          "vencidos (Pendiente), nunca enviados (Faltante) ou com data marcada.",
          "Marcar na Matriz os documentos vencidos, pendentes e faltantes (em "
          "vermelho) e montar a lista de pendências críticas."],
         ["lista_pessoal.xls",
          "Situação de acesso de cada colaborador no cliente: Autorizada, "
          "Desabilitada ou Baixas (ex-funcionário).",
          "Ignorar pendências de quem já foi desligado (Baixas)."],
     ]),
    ("p", "Dica: se o programa disser \"arquivo não encontrado\", confira se o nome do "
          "arquivo está exatamente igual ao da tabela acima."),

    ("h2", "4. A linguagem da Matriz (o que cada valor significa)"),
    ("p", "A Matriz usa códigos e cores que contam a história de cada documento. O "
          "programa fala essa mesma língua:"),
    ("table",
     ["Valor na célula", "Cor", "Significado"],
     [
         ["X", "Vermelho", "Documento ainda não iniciado (colaborador acabou de entrar)."],
         ["PACK", "—", "Kit de certificados está com o colaborador, aguardando assinatura."],
         ["Número (ex: 849204)", "Laranja", "Protocolo enviado ao Certronic, aguardando o cliente avaliar."],
         ["OK 12345", "Verde", "Documento aprovado e vigente — o número do protocolo aprovado fica preservado na célula."],
         ["OK 12345 (V:29/07/2026) ou OK 12345 (3 MESES)", "Verde", "Aprovado (protocolo preservado), mas vence na data indicada — programar renovação."],
         ["RESSALVA 12345", "Verde", "Cliente aprovou com ressalva (AceptadaParcial), com o nº do protocolo — ver observação no relatório."],
         ["X849204 (X colado no número)", "Vermelho", "Protocolo REPROVADO pelo cliente — corrigir e reenviar."],
         ["VENCIDO (18/06/2026)", "Vermelho", "Documento venceu na data indicada e o Certronic está cobrando."],
         ["PENDENTE", "Vermelho", "Certronic marca como Pendiente: venceu e não foi renovado."],
         ["FALTANTE", "Vermelho", "Certronic marca como Faltante: nunca foi enviado."],
         ["RECICLAR", "Laranja", "Precisa refazer o treinamento — controle MANUAL da operação (o programa não mexe)."],
         ["N/A", "—", "Documento não se aplica a esse colaborador."],
     ]),

    ("h2", "5. As regras de preenchimento — relatório por relatório"),

    ("h2", "5.1 documentacao_enviada (protocolos)"),
    ("p", "Como funciona: o programa varre a Matriz procurando células que contêm um "
          "número de protocolo. Para cada número encontrado, ele busca esse protocolo "
          "no relatório e olha o Estado:"),
    ("table",
     ["Estado no Certronic", "O que o programa faz na célula"],
     [
         ["Aceptada (aprovado)", "Célula vira OK + número do protocolo (ex: OK 12345), em "
          "verde. O número aprovado fica preservado na célula para rastreabilidade."],
         ["Rechazada (reprovado)", "Célula vira X + número (ex: X814525), em vermelho. O motivo "
          "da recusa (campo Observações do Certronic) aparece no relatório."],
         ["AceptadaParcial", "Célula vira RESSALVA + número (ex: RESSALVA 12345), em verde, e a "
          "observação do cliente aparece no relatório."],
         ["Pendiente de Control", "Não mexe — a célula já está certa (número em laranja, "
          "aguardando avaliação)."],
     ]),
    ("p", "Se o número da célula não estiver no relatório, o programa não faz nada — "
          "significa apenas que o protocolo é de um período diferente do exportado."),

    ("h2", "5.2 estado_vencimento (avisos de vencimento)"),
    ("p", "Esta regra só olha documentos que estão OK na Matriz e coloca o aviso de prazo:"),
    ("bullets", [
        "Data de vencimento no FUTURO: a célula ganha o aviso — OK (V:data) quando está "
        "perto (menos de 45 dias) ou OK (X MESES) quando falta mais tempo. Continua verde. "
        "Se a célula já traz o nº do protocolo (OK 12345), ele é preservado: OK 12345 (X MESES).",
        "Data de vencimento que JÁ PASSOU: quem resolve é a regra da documentacao_faltante "
        "(item 5.3). Mas se o documento vencido NÃO estiver sendo cobrado lá, o programa "
        "desconfia que existe uma renovação vigente (o estado_vencimento guarda histórico "
        "antigo) e NÃO marca nada — o caso vai para a lista \"a verificar\" do relatório.",
        "Se o mesmo documento aparecer mais de uma vez para a mesma pessoa (histórico), "
        "só vale o registro de data mais recente.",
    ]),

    ("h2", "5.3 documentacao_faltante (o alarme oficial — escreve em vermelho)"),
    ("p", "Este é o relatório que mostra o que o Certronic está cobrando de verdade. "
          "Falta de documento aqui pode bloquear o acesso do colaborador e o pagamento "
          "da IN HAUS. Por isso ele manda na Matriz:"),
    ("table",
     ["Coluna \"Vencimento\" no relatório", "O que o programa faz na célula"],
     [
         ["Data menor que hoje (já venceu)", "Célula vira VENCIDO (data), em vermelho."],
         ["Pendiente", "Célula vira PENDENTE, em vermelho."],
         ["Faltante", "Célula vira FALTANTE, em vermelho."],
         ["Data futura", "Não mexe aqui — o aviso de prazo vem do estado_vencimento (5.2)."],
     ]),
    ("p", "Exceção da célula X (não iniciado): se a célula já está como X e o "
          "Certronic cobra SEM data (Pendiente ou Faltante), a célula permanece X — "
          "só é trocada por VENCIDO (data) quando existe data de vencimento. Em todos "
          "os casos o documento continua na lista de pendências (ver proteções abaixo)."),
    ("p", "Além de marcar a célula, cada caso entra na aba \"Pendências críticas\" da "
          "tela e no relatório, para a operação agir."),
    ("p", "Proteções desta regra (quando o programa NÃO sobrescreve a célula):"),
    ("bullets", [
        "Colaborador desligado (Baixas na lista_pessoal): ignorado — ex-funcionário não gera pendência.",
        "Célula com PACK: NÃO é sobrescrita (decisão da operação em 23/07/2026, que "
        "substituiu a regra anterior). PACK = colaborador em fase de assinatura dos "
        "certificados, etapa manual. O vencimento continua entrando na aba de pendências, "
        "deixando claro que a alteração não foi feita por causa do PACK.",
        "Célula com RECICLAR: NÃO é sobrescrita — reciclagem é 100% manual da operação. "
        "Também entra nas pendências com o motivo (\"alteração não feita: célula em RECICLAR\").",
        "Célula com X (não iniciado): permanece X quando o Certronic cobra sem data "
        "(PENDENTE/FALTANTE); só vira VENCIDO (data) quando há data de vencimento "
        "concreta. Em todos os casos o documento segue na lista de pendências críticas.",
        "Documento com o protocolo NA PRÓPRIA célula do colaborador (número em laranja): "
        "a correção dessa pessoa já está correndo, então o programa não alarma nem apaga "
        "o número. (Um protocolo do mesmo tipo de documento no nome de OUTRA pessoa NÃO "
        "esconde o vencimento deste colaborador — a checagem é por pessoa.)",
        "Célula com anotação de recusa e número (ex: \"RECUSADO 678287\"): o número é "
        "preservado — correção em andamento.",
        "Célula com N/A: a Matriz diz \"não se aplica\", mas o Certronic está cobrando — "
        "alguém está errado. O programa NÃO mexe e lista como divergência para a operação "
        "verificar o cadastro.",
    ]),

    ("h2", "5.4 lista_pessoal (filtro de desligados)"),
    ("p", "Usada como filtro: colaboradores com acesso \"Baixas\" (ex-funcionários) não "
          "geram pendência nem alteração. \"Autorizada\" e \"Desabilitada\" são tratados "
          "como ativos — Desabilitada significa sem acesso à unidade, mas a pessoa ainda "
          "é colaboradora e a documentação continua sendo cobrada."),

    ("h2", "6. Casos que o programa não decide sozinho"),
    ("p", "Tudo que o programa pula aparece na seção \"Casos pulados\" do relatório, "
          "com o motivo. Os tipos mais comuns e o que fazer:"),
    ("table",
     ["Caso", "O que significa", "O que a operação deve fazer"],
     [
         ["CPF não encontrado na Matriz",
          "A pessoa está nos relatórios do Certronic mas não na Matriz — geralmente "
          "contratação nova ainda não cadastrada.",
          "Cadastrar o colaborador na Matriz (colunas A a I)."],
         ["Documento sem correspondência no de-para",
          "Apareceu um tipo de documento novo que o programa não sabe em qual coluna "
          "da Matriz colocar.",
          "Avisar o responsável pelo sistema para incluir o documento no arquivo de "
          "configuração (config.yaml)."],
         ["Vencido sem cobrança (a verificar)",
          "O estado_vencimento diz que venceu, mas a documentacao_faltante não cobra — "
          "provavelmente existe renovação vigente.",
          "Conferir no Certronic; se estiver tudo certo, nenhuma ação."],
         ["Divergência N/A",
          "A Matriz diz que o documento não se aplica, mas o Certronic cobra.",
          "Verificar qual cadastro está errado: o N/A da Matriz ou a exigência no Certronic."],
     ]),

    ("h2", "7. Como usar — passo a passo"),
    ("bullets", [
        "1. Exporte os 4 relatórios do Certronic e salve na pasta "
        "C:\\Projeto Certronic\\entrada com os nomes exatos da tabela do item 3.",
        "2. FECHE a planilha Matriz se ela estiver aberta no Excel (o Excel trava o "
        "arquivo e impede a gravação).",
        "3. Abra o programa e clique em \"1. Analisar relatórios\". Aguarde alguns segundos.",
        "4. Confira a aba \"Alterações na Matriz\": cada linha mostra o colaborador, o CPF, "
        "o documento, o valor atual, o valor novo e o motivo. Nada foi gravado ainda.",
        "5. Confira a aba \"Pendências críticas\": é a lista do que está vencido/faltante "
        "e precisa de ação da operação.",
        "6. Se estiver de acordo, clique em \"2. Aplicar na Matriz\" e confirme. Você pode "
        "desmarcar linhas que não quer gravar (caixa na primeira coluna) e usar os filtros "
        "no topo de cada coluna para localizar colaborador, documento etc. O programa grava "
        "só as linhas marcadas, com valores e cores.",
        "7. Ao final, um relatório em Excel (.xlsx) é salvo em C:\\Projeto Certronic\\relatorios, "
        "com tudo o que mudou, as pendências e os casos pulados. Guarde como auditoria.",
    ]),

    ("h2", "8. O que o programa produz a cada execução"),
    ("table",
     ["Arquivo", "O que é"],
     [
         ["Matriz atualizada", "A própria BASE CERTRONIC_GPS com os valores e cores novos."],
         ["Relatório da execução (relatorios\\atualizacao_DATA.xlsx)",
          "Registro auditável em Excel com abas: alterações aplicadas (só as marcadas), "
          "pendências críticas e casos pulados."],
         ["base_indicadores.xlsx",
          "Base estruturada que alimenta o Power BI: cadastro de colaboradores, tipos de "
          "documento, histórico de protocolos, foto diária da situação e pendências por data. "
          "Atualizada automaticamente ao Analisar e também ao Aplicar — não precisa fazer nada."],
     ]),

    ("h2", "9. Perguntas frequentes"),
    ("table",
     ["Pergunta", "Resposta"],
     [
         ["Rodei o programa duas vezes no mesmo dia. Duplicou algo?",
          "Não. A prévia pode ser rodada quantas vezes quiser, e a base de indicadores "
          "substitui os dados do dia em vez de somar."],
         ["O programa mexe no RECICLAR?",
          "Nunca. Nem coloca, nem tira, nem calcula prazo de reciclagem. É 100% manual."],
         ["Cliquei em Analisar e não quero aplicar. E agora?",
          "Só feche o programa (ou não clique em Aplicar). Nada foi gravado."],
         ["Deu erro \"planilha aberta no Excel\".",
          "Feche a Matriz no Excel e clique em Aplicar de novo."],
         ["Apareceu um documento novo no Certronic. O programa trata?",
          "Ele vai listar como \"caso pulado\" até alguém incluir o documento novo na "
          "configuração (de-para) ou na lista de ignorados."],
         ["Quem decide o que o programa pode sobrescrever?",
          "A operação. As decisões atuais estão no item 5.3 deste manual. Para mudar, "
          "fale com o responsável pelo sistema."],
     ]),

    ("h2", "Histórico de versões deste manual"),
    ("table",
     ["Versão", "Data", "O que mudou"],
     [["1.0", "14/07/2026", "Primeira versão."],
      ["1.1", "14/07/2026", "Relatório de execução passou de texto (.txt) para Excel (.xlsx) com abas."],
      ["1.2", "23/07/2026",
       "PACK deixou de ser sobrescrito (agora é preservado e vira pendência com motivo); "
       "RECICLAR vencido também entra nas pendências com motivo; supressão por protocolo "
       "passou a ser por pessoa (só a própria célula com o número); botão Aplicar passou a "
       "atualizar também a base de indicadores; abas ganharam filtro por coluna e seleção "
       "por linha."],
      ["1.3", "28/07/2026",
       "Célula X (não iniciado) passou a ser protegida: só vira VENCIDO quando há data de "
       "vencimento; PENDENTE/FALTANTE não sobrescrevem o X (o documento segue nas "
       "pendências). Validação de estrutura da Matriz e aba de Configurações (caminhos por "
       "máquina) adicionadas."],
      [VERSAO, DATA,
       "Protocolo aprovado passou a preservar o número na célula: Aceptada vira 'OK 12345' "
       "(e mantém o número no aviso de vencimento, 'OK 12345 (3 MESES)'); AceptadaParcial "
       "vira 'RESSALVA 12345'."]]),
]


def gerar_docx(destino: Path):
    doc = Document()
    estilo = doc.styles["Normal"]
    estilo.font.name = "Calibri"
    estilo.font.size = Pt(11)

    for bloco in CONTEUDO:
        tipo = bloco[0]
        if tipo == "h1":
            doc.add_heading(bloco[1], level=0)
        elif tipo == "h2":
            doc.add_heading(bloco[1], level=1)
        elif tipo == "p":
            doc.add_paragraph(bloco[1])
        elif tipo == "bullets":
            for item in bloco[1]:
                doc.add_paragraph(item, style="List Bullet")
        elif tipo == "table":
            cabecalho, linhas = bloco[1], bloco[2]
            tabela = doc.add_table(rows=1, cols=len(cabecalho))
            tabela.style = "Table Grid"
            for i, titulo in enumerate(cabecalho):
                celula = tabela.rows[0].cells[i]
                celula.text = titulo
                for run in celula.paragraphs[0].runs:
                    run.bold = True
            for linha in linhas:
                cells = tabela.add_row().cells
                for i, valor in enumerate(linha):
                    cells[i].text = valor
            doc.add_paragraph()

    destino.parent.mkdir(parents=True, exist_ok=True)
    doc.save(destino)


def gerar_markdown(destino: Path):
    partes = []
    for bloco in CONTEUDO:
        tipo = bloco[0]
        if tipo == "h1":
            partes.append(f"# {bloco[1]}\n")
        elif tipo == "h2":
            partes.append(f"## {bloco[1]}\n")
        elif tipo == "p":
            partes.append(bloco[1] + "\n")
        elif tipo == "bullets":
            partes.extend(f"- {item}" for item in bloco[1])
            partes.append("")
        elif tipo == "table":
            cabecalho, linhas = bloco[1], bloco[2]
            partes.append("| " + " | ".join(cabecalho) + " |")
            partes.append("|" + "---|" * len(cabecalho))
            for linha in linhas:
                partes.append("| " + " | ".join(v.replace("\n", " ") for v in linha) + " |")
            partes.append("")
    destino.write_text("\n".join(partes), encoding="utf-8")


if __name__ == "__main__":
    docx_destino = Path("C:/Projeto Certronic/Manual Atualizacao Matriz Certronic.docx")
    md_destino = Path(__file__).parent / "MANUAL.md"
    gerar_docx(docx_destino)
    gerar_markdown(md_destino)
    print(f"Word:     {docx_destino}")
    print(f"Markdown: {md_destino}")
