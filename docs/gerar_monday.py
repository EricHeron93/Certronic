# -*- coding: utf-8 -*-
"""Gera os artefatos para registro do projeto no monday.com:

1. monday_import.xlsx — planilha para importar no monday (cada linha = task)
2. Especificacao Projeto Certronic.docx — documento de especificação

Uso: python docs/gerar_monday.py
"""

from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from docx import Document
from docx.shared import Pt

DESTINO = Path("C:/Projeto Certronic")

# ---------------------------------------------------------------------------
# TASKS — (fase, nome, status, inicio, fim, prioridade, descricao)
# ---------------------------------------------------------------------------
TASKS = [
    # ----- FASE 1 ----------------------------------------------------------
    ("1. Levantamento & Análise",
     "Reunião de kickoff e levantamento do processo com a operação",
     "Concluído", "08/07/2026", "08/07/2026", "Alta",
     "Entendimento do processo de gestão documental no Certronic: fluxo de envio, "
     "protocolos, validação pelo cliente e o controle manual feito na planilha "
     "BASE CERTRONIC_GPS (Matriz). Alinhamento de acessos e responsabilidades."),
    ("1. Levantamento & Análise",
     "Análise da planilha Matriz (BASE CERTRONIC_GPS)",
     "Concluído", "08/07/2026", "08/07/2026", "Alta",
     "Mapeamento da estrutura: cadastro (colunas A-I), 16 colunas de documentos, "
     "convenções de preenchimento (OK, PACK, protocolos, X, RECICLAR, N/A) e cores."),
    ("1. Levantamento & Análise",
     "Análise dos 4 relatórios exportados do Certronic",
     "Concluído", "08/07/2026", "08/07/2026", "Alta",
     "Engenharia reversa dos formatos: documentacao_enviada (xlsx disfarçado de .xls), "
     "estado_vencimento, documentacao_faltante e lista_pessoal (HTML disfarçados de .xls). "
     "Levantamento de colunas, estados e volumes."),
    ("1. Levantamento & Análise",
     "Definição das regras de negócio da automação",
     "Concluído", "08/07/2026", "10/07/2026", "Alta",
     "Regras de atualização por relatório, ciclo de vida da célula (X → PACK → protocolo "
     "→ OK/rejeitado), restrições invioláveis (colunas A-I e RECICLAR intocáveis, "
     "prévia obrigatória antes de gravar, casos incertos pulados e reportados)."),
    ("1. Levantamento & Análise",
     "Definição da arquitetura da solução",
     "Concluído", "08/07/2026", "08/07/2026", "Média",
     "Python + interface Tkinter empacotados em executável; leitura dos exports em pasta "
     "padronizada; escrita na Matriz local sincronizada via OneDrive; config.yaml "
     "editável pela operação (caminhos e de-para de documentos)."),

    # ----- FASE 2 ----------------------------------------------------------
    ("2. Desenvolvimento",
     "Estrutura do projeto e arquivo de configuração",
     "Concluído", "08/07/2026", "08/07/2026", "Alta",
     "Esqueleto do projeto (main, src/leitores, matriz, regras, relatorio, gui) e "
     "config.yaml com caminhos, nomes padrão dos arquivos e de-para de documentos."),
    ("2. Desenvolvimento",
     "Leitores dos 4 relatórios do Certronic",
     "Concluído", "08/07/2026", "08/07/2026", "Alta",
     "Parsers com detecção do formato real dos arquivos (xlsx ou HTML disfarçados de "
     ".xls), normalização de CPF (zeros à esquerda) e tratamento de datas."),
    ("2. Desenvolvimento",
     "Módulo de leitura/escrita da Matriz",
     "Concluído", "08/07/2026", "08/07/2026", "Alta",
     "Indexação por CPF e por coluna de documento; escrita restrita à faixa J-Y; "
     "bloqueio de células RECICLAR; aplicação das cores (verde/laranja/vermelho)."),
    ("2. Desenvolvimento",
     "Motor de regras (cruzamento dos relatórios com a Matriz)",
     "Concluído", "08/07/2026", "10/07/2026", "Alta",
     "Regra de protocolos (Aceptada→OK, Rechazada→X+ID, AceptadaParcial→OK RESSALVA), "
     "regra de vencimentos (avisos OK (V:data)/OK (X MESES)) e regra de faltantes "
     "(VENCIDO/PENDENTE/FALTANTE em vermelho), com filtros de desligados e proteção "
     "contra alarmes falsos (protocolo em análise, histórico do estado_vencimento)."),
    ("2. Desenvolvimento",
     "Interface gráfica com prévia e confirmação",
     "Concluído", "08/07/2026", "10/07/2026", "Alta",
     "Tela Tkinter com fluxo Analisar → prévia em tabela (com CPF) → Aplicar com "
     "confirmação. Aba adicional de Pendências críticas com contador."),
    ("2. Desenvolvimento",
     "Relatório de execução em Excel",
     "Concluído", "14/07/2026", "14/07/2026", "Média",
     "Arquivo por execução com abas Resumo (status prévia/aplicada e totais), "
     "Alterações, Pendências críticas e Casos pulados — formatado com filtros."),
    ("2. Desenvolvimento",
     "Base de indicadores para o Power BI",
     "Concluído", "09/07/2026", "14/07/2026", "Alta",
     "base_indicadores.xlsx gerada a cada análise: dimensões (colaboradores, documentos), "
     "fatos históricos (protocolos com vínculo de CPF preservado, situacao_atual e "
     "pendencias acumuladas por data) e visão larga por colaborador com % conformidade."),

    # ----- FASE 3 ----------------------------------------------------------
    ("3. Testes & Ajustes",
     "Montagem do ambiente de teste com dados reais",
     "Concluído", "08/07/2026", "08/07/2026", "Alta",
     "Pasta padronizada com os 4 exports reais renomeados e cópia da Matriz "
     "(MATRIZ_TESTE.xlsx) para validação sem risco à planilha oficial."),
    ("3. Testes & Ajustes",
     "Primeira análise completa com dados reais (dry-run)",
     "Concluído", "08/07/2026", "08/07/2026", "Alta",
     "Validação do cruzamento ponta a ponta; calibragem do de-para de documentos e "
     "criação da lista de documentos ignorados (folha mensal sem coluna na Matriz)."),
    ("3. Testes & Ajustes",
     "Ajuste das regras de vencimento com validação cruzada",
     "Concluído", "09/07/2026", "10/07/2026", "Alta",
     "Correção de documentos vencidos rotulados como OK; proteção contra o histórico do "
     "estado_vencimento (registro antigo vencido com renovação vigente) via validação "
     "cruzada com a documentacao_faltante; deduplicação pela data mais recente."),
    ("3. Testes & Ajustes",
     "Tratamento de casos-limite definidos com a operação",
     "Concluído", "10/07/2026", "10/07/2026", "Média",
     "Decisões registradas: PACK é sobrescrito pela cobrança; N/A vira divergência "
     "sem tocar a célula; anotações de recusa fora do padrão são preservadas; formato "
     "do reprovado padronizado como X+número (ex: X814525)."),
    ("3. Testes & Ajustes",
     "Checagem de completude de cadastro",
     "Concluído", "14/07/2026", "14/07/2026", "Média",
     "Todo colaborador ativo no Certronic que não existe na Matriz é apontado no "
     "relatório com instrução de cadastro (independe de aparecer em outros relatórios)."),
    ("3. Testes & Ajustes",
     "Desempenho e robustez da interface",
     "Concluído", "14/07/2026", "14/07/2026", "Alta",
     "Análise movida para segundo plano com indicador de progresso (Matriz com 359 mil "
     "linhas residuais levava ~77s congelando a tela); indexação otimizada; tratamento "
     "de arquivos abertos no Excel e de base de indicadores corrompida."),
    ("3. Testes & Ajustes",
     "Testes de aplicação na cópia da Matriz",
     "Concluído", "10/07/2026", "15/07/2026", "Alta",
     "Aplicação real das alterações na MATRIZ_TESTE com verificação de valores, cores e "
     "idempotência (rodar de novo não propõe nem duplica nada)."),

    # ----- FASE 4 ----------------------------------------------------------
    ("4. Documentação",
     "Manual de regras e uso para a operação",
     "Concluído", "14/07/2026", "14/07/2026", "Alta",
     "Documento em linguagem não técnica: garantias, os 4 relatórios, legenda da Matriz, "
     "regras por relatório, casos não decididos automaticamente, passo a passo e FAQ. "
     "Gerado por script para manter versão junto do código (v1.1)."),
    ("4. Documentação",
     "Documentação da base de indicadores para o analista de BI",
     "Concluído", "14/07/2026", "14/07/2026", "Alta",
     "Contexto completo do negócio, linhagem de cada dado, dicionário de todas as "
     "tabelas, modelo relacional sugerido, 10 indicadores propostos, cuidados de "
     "modelagem e glossário."),
    ("4. Documentação",
     "Registro do projeto no monday.com",
     "Concluído", "15/07/2026", "15/07/2026", "Média",
     "Criação do board com as fases e tasks do projeto e publicação da especificação."),

    # ----- FASE 5 ----------------------------------------------------------
    ("5. Implantação (pendente)",
     "Limpeza das linhas residuais da Matriz",
     "A fazer", "", "", "Média",
     "Remover as ~359 mil linhas de formatação residual da planilha para reduzir o tempo "
     "de análise de ~1 minuto para segundos. Validar na cópia antes de aplicar na oficial."),
    ("5. Implantação (pendente)",
     "Geração do executável (PyInstaller)",
     "A fazer", "", "", "Alta",
     "Empacotar o programa em .exe com config.yaml externo para a máquina da operação, "
     "sem necessidade de instalar Python."),
    ("5. Implantação (pendente)",
     "Testes da Matriz e kick-off com a planilha Matriz real",
     "A fazer", "", "", "Alta",
     "Apontar o programa para a Matriz oficial (OneDrive), executar o ciclo completo de "
     "testes com dados de produção e realizar o kick-off da operação assistida."),
    ("5. Implantação (pendente)",
     "Treinamentos com a operação",
     "A fazer", "", "", "Alta",
     "Sessões de treinamento com as responsáveis pela operação usando o manual como "
     "guia: preparo dos exports, análise, conferência da prévia, aplicação e leitura "
     "dos relatórios."),
    ("5. Implantação (pendente)",
     "Homologação com a operação",
     "A fazer", "", "", "Alta",
     "Rodadas assistidas de validação com a operação e coleta de ajustes finais antes "
     "da rotina oficial."),
    ("5. Implantação (pendente)",
     "Reunião de alinhamento com o analista de BI",
     "A fazer", "", "", "Média",
     "Apresentar a base_indicadores.xlsx e a documentação de BI ao analista, alinhar o "
     "modelo de dados, os indicadores prioritários e a fonte no OneDrive/SharePoint."),
    ("5. Implantação (pendente)",
     "Reunião de apresentação do projeto",
     "A fazer", "", "", "Média",
     "Apresentação da solução para as partes interessadas (operação e gestão): fluxo, "
     "regras, ganhos e próximos passos."),
    ("5. Implantação (pendente)",
     "Início da rotina oficial + handoff do Power BI",
     "A fazer", "", "", "Média",
     "Definir a frequência de execução, ativar a rotina e entregar a base de indicadores "
     "ao analista de BI para construção do dashboard."),
]


def gerar_xlsx():
    """Formato espelhado no padrão de board usado pela empresa no monday:
    Nome | Grupo | Pessoas | Status | Cronograma Início | Cronograma Fim |
    Data Final Real | Dias de Atraso | Descrição
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tasks"
    cab = ["Nome", "Grupo", "Pessoas", "Status", "Cronograma Início",
           "Cronograma Fim", "Data Final Real", "Dias de Atraso", "Descrição"]
    ws.append(cab)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="323338", end_color="323338", fill_type="solid")
        c.alignment = Alignment(horizontal="center")
    for fase, nome, status, ini, fim, prio, desc in TASKS:
        status_monday = "Concluído" if status == "Concluído" else "Não Iniciado"
        data_final_real = fim if status == "Concluído" else ""
        dias_atraso = 0 if status == "Concluído" else ""
        ws.append([nome, fase, "", status_monday, ini, fim,
                   data_final_real, dias_atraso, desc])
    larguras = [55, 26, 16, 14, 16, 16, 15, 13, 100]
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    destino = DESTINO / "monday_import.xlsx"
    wb.save(destino)
    return destino


ESPEC = [
    ("h1", "Especificação — Automação do Controle Documental Certronic"),
    ("p", "Projeto: Automação da planilha BASE CERTRONIC_GPS (Matriz) | "
          "Área: Operação Mercado Livre — IN HAUS | Data: 30/07/2026 | Versão 1.3"),

    ("h2", "1. Contexto e problema"),
    ("p", "A IN HAUS mantém colaboradores alocados em unidades do Mercado Livre. Cada "
          "colaborador precisa de ~16 documentos regulatórios válidos (ASO, NRs, EPI, "
          "integração Meli etc.), validados pelo cliente através do sistema Certronic. "
          "Documentação vencida ou faltante pode bloquear o acesso do colaborador à "
          "unidade e travar o pagamento dos serviços."),
    ("p", "O Certronic não oferece visão consolidada. O controle era feito manualmente "
          "numa planilha (Matriz): a operação cruzava, célula por célula, 4 relatórios "
          "exportados do sistema — processo lento, repetitivo e sujeito a erro."),

    ("h2", "2. Objetivo"),
    ("p", "Automatizar a atualização da Matriz a partir dos 4 relatórios do Certronic, "
          "com prévia e confirmação antes de qualquer gravação, alertando pendências "
          "críticas e gerando uma base estruturada de indicadores para o Power BI."),

    ("h2", "3. Escopo"),
    ("p", "DENTRO do escopo:"),
    ("bullets", [
        "Leitura dos 4 relatórios do Certronic (documentacao_enviada, estado_vencimento, "
        "documentacao_faltante, lista_pessoal) em pasta padronizada.",
        "Atualização automática das colunas de documentos da Matriz (J a Y), com cores.",
        "Prévia de todas as alterações e confirmação obrigatória antes de gravar.",
        "Alerta de pendências críticas (vencidos/faltantes sem correção em andamento).",
        "Relatório auditável em Excel por execução.",
        "Base de indicadores (base_indicadores.xlsx) para consumo do Power BI.",
        "Manual de uso para a operação e documentação para o analista de BI.",
    ]),
    ("p", "FORA do escopo:"),
    ("bullets", [
        "Alterar as colunas de cadastro da Matriz (A a I) — permanecem manuais.",
        "Gerenciar o ciclo de RECICLAR (reciclagem de treinamentos) — decisão e registro "
        "100% manuais da operação, incluindo o cálculo de prazos.",
        "Integração direta com o Certronic (API) — a entrada são os relatórios exportados.",
        "Construção do dashboard Power BI (responsabilidade do analista de BI, com base "
        "e documentação fornecidas pelo projeto).",
    ]),

    ("h2", "4. Regras de negócio (resumo)"),
    ("table",
     ["Origem", "Regra"],
     [
      ["documentacao_enviada",
       "Células da Matriz com nº de protocolo são atualizadas conforme o estado: "
       "Aceptada→OK + nº (ex.: OK 12345, preserva o protocolo); Rechazada→X+número "
       "(vermelho) com observações no relatório; AceptadaParcial→RESSALVA + nº (ex.: "
       "RESSALVA 12345); Pendiente de Control→sem ação. O nº do protocolo é mantido "
       "mesmo quando entra o aviso de vencimento (OK 12345 (3 MESES))."],
      ["estado_vencimento",
       "Documento OK com data futura ganha aviso: OK (V:data) se < 45 dias, senão "
       "OK (N MESES). Registros históricos: vale a data mais recente por CPF+documento; "
       "vencido sem cobrança na documentacao_faltante vira caso 'a verificar' "
       "(provável renovação vigente), sem alteração."],
      ["documentacao_faltante",
       "Fonte oficial de cobrança — escreve em vermelho: data passada→VENCIDO (data); "
       "Pendiente→PENDENTE; Faltante→FALTANTE. Exceção: célula X (não iniciado) só vira "
       "VENCIDO quando há data; PENDENTE/FALTANTE não sobrescrevem o X. A correção só é "
       "considerada em andamento (e a pendência suprimida) quando o nº do protocolo está "
       "na PRÓPRIA célula do colaborador — protocolo do mesmo documento no nome de outra "
       "pessoa não esconde o vencimento deste. Cada caso vira pendência crítica."],
      ["lista_pessoal",
       "Colaboradores com acesso 'Baixas' (desligados) não geram pendência nem alteração. "
       "Ativo no Certronic sem cadastro na Matriz é apontado para cadastro."],
      ["Proteções gerais",
       "Colunas A-I nunca são alteradas. Células RECICLAR e PACK também não são "
       "sobrescritas (etapas manuais da operação): quando o documento está vencido, a "
       "célula é preservada e o caso entra nas pendências deixando claro que a alteração "
       "não foi feita por causa do PACK/RECICLAR. N/A cobrado vira divergência a "
       "verificar; anotações de recusa fora do padrão são preservadas; casos não "
       "classificáveis são pulados e listados, nunca 'chutados'."],
     ]),
    ("p", "A referência completa e atualizada das regras é o 'Manual da Automação "
          "Certronic → Matriz' (v1.4), mantido junto do sistema."),

    ("h2", "5. Solução entregue"),
    ("bullets", [
        "Aplicação desktop em Python com interface gráfica (Tkinter): fluxo Analisar → "
        "prévia em tabela → Aplicar com confirmação; aba de pendências críticas.",
        "Prévia com filtro por coluna (uma caixa sob cada título) e seleção por linha "
        "(checkbox): a operação pode desmarcar alterações que não quer gravar e o "
        "programa aplica apenas as marcadas.",
        "Configuração externa (config.yaml): caminhos, nomes padrão dos arquivos, "
        "de-para de documentos e lista de ignorados — ajustável sem mexer em código.",
        "Relatório de execução em Excel (Resumo, Alterações, Pendências, Casos pulados), "
        "registrando apenas as alterações efetivamente aplicadas.",
        "base_indicadores.xlsx com 6 tabelas: colaboradores, documentos, protocolos "
        "(com vínculo CPF preservado ao longo do tempo), situacao_atual e pendencias "
        "(históricas, acumuladas por data) e visao_colaborador (% de conformidade). "
        "Atualizada tanto ao Analisar quanto ao Aplicar.",
        "Processamento em segundo plano com indicador de progresso; tratamento de "
        "arquivos abertos/corrompidos; execuções repetidas no mesmo dia não duplicam dados.",
    ]),

    ("h2", "6. Fluxo de uso"),
    ("bullets", [
        "1. Operação exporta os 4 relatórios do Certronic para a pasta de entrada com os nomes padrão.",
        "2. Abre o programa e clica em 'Analisar relatórios' (≈1 minuto).",
        "3. Confere a prévia de alterações e a aba de pendências críticas.",
        "4. Opcionalmente filtra por coluna e desmarca linhas que não quer gravar; clica "
        "em 'Aplicar na Matriz' e confirma — só então a planilha é gravada (apenas as marcadas).",
        "5. Arquiva o relatório Excel gerado; a base de indicadores é atualizada "
        "automaticamente (no Analisar e no Aplicar).",
    ]),

    ("h2", "7. Decisões registradas"),
    ("table",
     ["Data", "Decisão"],
     [
      ["08/07/2026", "Prévia + confirmação obrigatórias; casos incertos pulados e listados."],
      ["09/07/2026", "RECICLAR fora do escopo da automação, incluindo cálculo de prazos."],
      ["10/07/2026", "Formato do reprovado: X colado ao número (X814525). PACK é sobrescrito "
       "pela cobrança; N/A vira divergência; recusas fora do padrão preservadas."],
      ["10/07/2026", "Relatórios passam a ser referidos pelos nomes dos arquivos "
       "(documentacao_enviada etc.), abandonando a numeração 01-04."],
      ["14/07/2026", "Relatório de execução em Excel; documento vencido marca a célula "
       "como VENCIDO (data) em vermelho além de constar nas pendências."],
      ["21/07/2026", "Supressão de vencido por protocolo passou a ser por pessoa (só a "
       "própria célula com o número, não por nome de documento). Botão Aplicar passou a "
       "atualizar também a base de indicadores. Prévia ganhou filtro por coluna e "
       "seleção por linha."],
      ["23/07/2026", "PACK deixou de ser sobrescrito (substitui a decisão de 10/07): é "
       "preservado e o vencimento entra nas pendências com o motivo. RECICLAR vencido "
       "também passa a constar nas pendências com o motivo. Ambas são etapas manuais."],
      ["28/07/2026", "Célula X (não iniciado) protegida: só vira VENCIDO quando há data "
       "de vencimento; PENDENTE/FALTANTE não sobrescrevem o X. O documento continua na "
       "lista de pendências críticas (risco visível)."],
      ["30/07/2026", "Protocolo aprovado passa a preservar o número na célula: Aceptada "
       "vira 'OK 12345' (número mantido inclusive no aviso de vencimento); AceptadaParcial "
       "vira 'RESSALVA 12345'. Nova categoria de BI OK_RESSALVA (conforme com ressalva)."],
     ]),

    ("h2", "8. Riscos e pontos de atenção"),
    ("bullets", [
        "A Matriz oficial contém ~359 mil linhas de formatação residual que tornam a "
        "análise lenta (~1 min). Limpeza recomendada antes da implantação.",
        "Os nomes dos arquivos exportados precisam seguir o padrão — nome errado "
        "interrompe a análise (com mensagem orientativa).",
        "Protocolos anteriores à automação não têm dono identificável (o relatório do "
        "Certronic não traz CPF) — indicadores por pessoa só valem do início da "
        "automação em diante.",
        "Mudanças de layout nos exports do Certronic exigem ajuste nos leitores.",
    ]),

    ("h2", "9. Critérios de aceite"),
    ("bullets", [
        "Análise completa dos 4 relatórios reais sem erro, com prévia fiel às regras.",
        "Nenhuma escrita fora das colunas J-Y; nenhuma célula RECICLAR alterada.",
        "Reexecução imediata não propõe alterações (idempotência).",
        "Pendências críticas conferem com a cobrança do Certronic (sem falso positivo "
        "de documento com correção em andamento).",
        "Base de indicadores íntegra após múltiplas execuções (sem duplicação).",
        "Operação executa o fluxo completo apoiada apenas no manual.",
    ]),

    ("h2", "10. Próximas etapas"),
    ("bullets", [
        "Limpeza das linhas residuais da Matriz.",
        "Geração do executável (PyInstaller) e instalação na máquina da operação.",
        "Apontamento para a Matriz oficial no OneDrive.",
        "Homologação com a operação e início da rotina oficial.",
        "Handoff da base de indicadores ao analista de BI para construção do dashboard.",
    ]),
]


def gerar_docx():
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    for bloco in ESPEC:
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
            cab, linhas = bloco[1], bloco[2]
            t = doc.add_table(rows=1, cols=len(cab))
            t.style = "Table Grid"
            for i, titulo in enumerate(cab):
                c = t.rows[0].cells[i]
                c.text = titulo
                for run in c.paragraphs[0].runs:
                    run.bold = True
            for linha in linhas:
                cells = t.add_row().cells
                for i, v in enumerate(linha):
                    cells[i].text = v
            doc.add_paragraph()
    destino = DESTINO / "Especificacao Projeto Certronic.docx"
    doc.save(destino)
    return destino


if __name__ == "__main__":
    print("Tasks:", gerar_xlsx())
    print("Especificação:", gerar_docx())
