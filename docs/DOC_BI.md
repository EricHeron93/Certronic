# Base de Indicadores Certronic — Documentação para BI

Versão 1.3 — 30/07/2026 — IN HAUS / Operação Mercado Livre

Este documento explica, do zero, o processo de negócio por trás do arquivo base_indicadores.xlsx, de onde nasce cada informação, o que significa cada tabela e coluna, e como modelar tudo no Power BI. Foi escrito para quem não conhece nada da operação — leia a Parte 1 antes de abrir o arquivo.

## PARTE 1 — O NEGÓCIO (leia antes de tudo)

## 1.1 Quem somos e o que está em jogo

A IN HAUS presta serviços de manutenção predial ao Mercado Livre. Nossos colaboradores (eletricistas, oficiais de manutenção etc.) trabalham DENTRO das unidades do Mercado Livre, espalhadas pelo Brasil em regionais (SP, MG, RJES, NENO...) e sítios (SMG11 - Patos de Minas, SSP34 - Cotia...).

Para um colaborador poder entrar e trabalhar numa unidade, ele precisa ter TODA a documentação regulatória válida: exame médico (ASO), treinamentos de segurança (NR10 eletricidade, NR35 altura, NR18, NR26...), ficha de EPI, contrato, integração com o Mercado Livre (Onboarding Meli), entre outros. São ~16 tipos de documento por pessoa.

SE A DOCUMENTAÇÃO FALTA OU VENCE: o colaborador pode ter o acesso bloqueado na unidade e o PAGAMENTO da IN HAUS pelo serviço pode ser travado pelo cliente. É por isso que este controle existe e é crítico. Os indicadores que você vai construir medem exatamente esse risco.

## 1.2 O Certronic e o fluxo de validação

O Mercado Livre valida os documentos através de um sistema de terceiros chamado CERTRONIC. O fluxo é assim:

- 1. A operação da IN HAUS envia um documento (ex: ASO renovado) para o Certronic.
- 2. O Certronic gera um número sequencial chamado PROTOCOLO (ou ID) para esse envio.
- 3. O cliente analisa e dá um veredito: Aceptada (aprovado), Rechazada (reprovado), AceptadaParcial (aprovado com ressalva) ou fica Pendiente de Control (em análise). Os nomes são em espanhol porque o sistema é latino-americano.
- 4. Se reprovado, a operação corrige e reenvia (gera novo protocolo).

Um detalhe que molda os dados: o relatório de protocolos do Certronic NÃO diz de quem é cada protocolo (não tem CPF nem nome). Guarde essa informação — ela explica um comportamento importante da tabela de protocolos (item 3.4).

## 1.3 A Matriz — a planilha de controle da operação

Como o Certronic dá pouca visibilidade do conjunto, a operação mantém uma planilha de controle chamada BASE CERTRONIC_GPS (apelidada de "Matriz"): uma linha por colaborador, uma coluna por tipo de documento, e em cada cruzamento um código que conta a história daquele documento:

| Código na Matriz | Significado |
|---|---|
| X | Documento ainda não iniciado (pessoa acabou de ser contratada). |
| PACK | Kit de certificados aguardando a assinatura do colaborador. |
| Número (ex: 849204) | Protocolo enviado ao Certronic, aguardando análise do cliente. |
| OK 12345 | Documento aprovado e vigente — o nº do protocolo aprovado fica preservado. |
| OK 12345 (V:13/05/2027) / OK 12345 (3 MESES) | Vigente (protocolo preservado), mas vence na data indicada. |
| RESSALVA 12345 | Aprovado com ressalva (AceptadaParcial), com o nº do protocolo. |
| X849204 | Protocolo REPROVADO — em correção. |
| VENCIDO (18/06/2026) | Venceu na data indicada e está sendo cobrado. |
| PENDENTE | O Certronic marca como Pendiente: venceu e não foi renovado. |
| FALTANTE | O Certronic marca como Faltante: nunca foi enviado. |
| RECICLAR | Treinamento precisa ser refeito (controle manual da operação). |
| N/A | Documento não se aplica àquela função/pessoa. |

A Matriz é a FONTE DA VERDADE de quem a operação controla: só quem está cadastrado nela entra no controle (e na maioria das tabelas da base).

## 1.4 A automação — de onde vem o arquivo que você vai usar

Um programa interno (Python) automatiza esse controle. Periodicamente, a operação exporta 4 relatórios do Certronic e roda o programa, que: (1) cruza tudo, (2) atualiza a Matriz, (3) aponta pendências críticas e (4) GERA/ATUALIZA o arquivo base_indicadores.xlsx — que é a sua fonte de dados. Você nunca deve editar esse arquivo à mão: ele é 100% gerado pelo programa, e edições manuais são sobrescritas.

Os 4 relatórios de origem (para você entender a linhagem do dado):

| Relatório do Certronic | O que traz | Alimenta |
|---|---|---|
| documentacao_enviada | Protocolos e seus estados (SEM CPF). | Tabela protocolos |
| estado_vencimento | Datas de vencimento por CPF e documento. | data_vencimento nas tabelas |
| documentacao_faltante | O que o Certronic está cobrando: vencido (Pendiente), nunca enviado (Faltante) ou com data marcada. | Tabela pendencias e células PENDENTE/FALTANTE/VENCIDO |
| lista_pessoal | Situação de acesso: Autorizada, Desabilitada, Baixas (desligado). | Coluna acesso |

## PARTE 2 — O ARQUIVO base_indicadores.xlsx

## 2.1 Visão geral das tabelas

| Aba | Granularidade | Comportamento a cada execução | Papel no modelo |
|---|---|---|---|
| colaboradores | 1 linha por colaborador da Matriz | SOBRESCRITA (retrato atual) | Dimensão |
| documentos | 1 linha por tipo de documento | SOBRESCRITA | Dimensão |
| protocolos | 1 linha por protocolo (ID) | UPSERT: estado atualiza, histórico não some | Fato |
| situacao_atual | 1 linha por colaborador × documento × data de execução | ACUMULA por data (série histórica) | Fato principal |
| visao_colaborador | 1 linha por colaborador (formato largo) | SOBRESCRITA (foto atual) | Consulta/consumo direto |
| pendencias | 1 linha por pendência crítica × data de execução | ACUMULA por data | Fato |

Regra geral: rodar o programa duas vezes no mesmo dia NÃO duplica — as linhas do dia são substituídas.

## 2.2 Dicionário: colaboradores (dimensão)

| Coluna | Significado |
|---|---|
| cpf | CHAVE. Texto de 11 dígitos COM zeros à esquerda (ex: 03898590607). Nunca converta para número — perde o zero e quebra os relacionamentos. |
| nome | Nome completo do colaborador. |
| funcao | Função (ELETR., OMP, 1/2 OMP...). |
| cr | Centro de resultado — código + descrição do contrato (ex: 50996 - MG - MAP...). |
| regional | Regional do Mercado Livre onde atua (MG, SPC, SPI, RJES, NENO...). |
| sitio | Unidade específica (ex: SMG11 - PATOS DE MINAS). |
| status_docs | Texto livre escrito pela operação sobre a situação geral da pessoa. Não padronizado — use com cautela, bom para tooltip, ruim para agrupar. |
| admissao | Data de admissão (texto dd/mm/aaaa). |
| status_certronic | Status do cadastro no Certronic (texto livre da operação). |
| acesso | Da lista_pessoal: Autorizada (pode trabalhar), Desabilitada (sem acesso, mas ainda é colaborador) ou Baixas (desligado). Vazio = não localizado no Certronic. |

## 2.3 Dicionário: documentos (dimensão)

| Coluna | Significado |
|---|---|
| nome_certronic | Início do nome do documento como aparece nos relatórios do Certronic. |
| coluna_matriz | Início do cabeçalho da coluna correspondente na Matriz. |

Atenção: esta tabela é um "de-para" de prefixos de nome, não uma dimensão perfeita. O relacionamento com as tabelas de fato funciona por correspondência de início de texto, não por chave exata. Para agrupar por documento no BI, muitas vezes é mais prático usar direto o campo "documento" da própria tabela de fato.

## 2.4 Dicionário: protocolos (fato)

| Coluna | Significado |
|---|---|
| id_protocolo | CHAVE. Número sequencial gerado pelo Certronic no envio. |
| cpf | Dono do protocolo. PODE ESTAR VAZIO — leia o quadro abaixo. |
| documento | Nome do documento enviado. |
| estado | Aceptada / Rechazada / AceptadaParcial / Pendiente de Control. |
| data_enviado | Quando a operação enviou (texto dd-mm-aaaa). |
| data_acusado | Quando o cliente deu o veredito (texto dd-mm-aaaa). |
| observacoes / comentario | Textos do avaliador — úteis para entender rejeições. |

POR QUE EXISTEM PROTOCOLOS SEM CPF? O relatório de protocolos do Certronic não informa o dono. O programa descobre o CPF "espiando" a Matriz: enquanto um protocolo está em análise, o número dele fica anotado na célula do colaborador — e nesse momento o programa grava o vínculo, que fica preservado para sempre. Protocolos antigos (anteriores à automação) nunca passaram por essa captura, por isso estão sem dono — e não há como recuperar. A cobertura de CPF só cresce daqui para frente — e melhorou: como o protocolo aprovado agora preserva o número na célula (ex.: 'OK 12345'), o programa consegue vincular o CPF também depois da aprovação, não só enquanto estava em análise. Consequência prática: indicadores POR PESSOA usam apenas protocolos com CPF; indicadores AGREGADOS (tempo médio de aprovação, taxa de rejeição por documento) podem usar a tabela inteira.

## 2.5 Dicionário: situacao_atual (fato principal)

| Coluna | Significado |
|---|---|
| data_execucao | Data em que o programa rodou (aaaa-mm-dd). É o eixo do tempo do seu modelo. |
| cpf | Colaborador (relaciona com a dimensão colaboradores). |
| nome | Nome (redundante, para conferência). |
| documento | Cabeçalho da coluna na Matriz (nome longo do documento). |
| status_celula | Categoria padronizada do status — ver tabela 2.6. |
| data_vencimento | Data de vencimento quando conhecida (texto dd/mm/aaaa). |

Esta é a tabela para séries históricas: cada execução adiciona uma "foto" completa (colaborador × documento). Para ver a situação ATUAL, filtre sempre pela última data_execucao. Para evolução, agrupe por data_execucao.

## 2.6 As categorias de status_celula

| Categoria | Significado | Conta como pendência? |
|---|---|---|
| OK | Documento aprovado e vigente (inclui os OK com aviso de vencimento e os OK com nº de protocolo, ex.: 'OK 12345'). | Não |
| OK_RESSALVA | Aprovado com ressalva (AceptadaParcial, ex.: 'RESSALVA 12345'). Conta como conforme (somado ao OK no % de conformidade), mas é distinguível para a operação acompanhar. | Não (requer atenção — ver observação) |
| PROTOCOLO_EM_ANALISE | Enviado, aguardando veredito do cliente. | Não (em andamento) |
| PACK | Kit aguardando assinatura do colaborador. | Não (em andamento) |
| REJEITADO | Protocolo reprovado, em correção. | Sim |
| VENCIDO | Venceu e está sendo cobrado. | Sim |
| PENDENTE | Vencido não renovado (Pendiente no Certronic). | Sim |
| FALTANTE | Nunca enviado (Faltante no Certronic). | Sim |
| NAO_INICIADO | Ainda não começou (X na Matriz). | Sim |
| RECICLAR | Treinamento a refazer — controle manual da operação. | Sim (atenção da operação) |
| N/A | Não se aplica à pessoa. EXCLUIR dos denominadores de conformidade. | Não se aplica |
| VAZIO / OUTRO | Célula vazia ou texto não padronizado da operação. | Analisar caso a caso |

## 2.7 Dicionário: visao_colaborador (foto larga)

Espelho da tela "Doc. Pessoal" do Certronic: uma linha por colaborador, uma coluna por documento com o valor real da célula da Matriz (enriquecido com a data de vencimento quando conhecida), e três colunas de resumo:

| Coluna | Significado |
|---|---|
| docs_ok | Quantos documentos estão OK. |
| docs_com_pendencia | Quantos estão em VENCIDO, PENDENTE, FALTANTE, REJEITADO ou NAO_INICIADO. |
| pct_conforme | docs_ok ÷ documentos aplicáveis (N/A fora da conta), em %. |

É a tabela mais fácil de consumir diretamente (ranking de risco, lista de 100% conformes), mas é SOBRESCRITA a cada execução — sem histórico. Para histórico, use situacao_atual.

## 2.8 Dicionário: pendencias (fato)

| Coluna | Significado |
|---|---|
| data_execucao | Data da execução (acumula — serve para evolução). |
| cpf / nome | Colaborador. |
| documento | Documento cobrado. |
| descricao | Situação detalhada (ex: VENCIDO em 17/12/2025 — sem protocolo em análise). |

Importante: só entra aqui pendência REAL — colaborador ativo e documento cobrado pelo Certronic. A correção só é considerada "em andamento" (e a pendência suprimida) quando o número do protocolo está NA PRÓPRIA célula do colaborador — um protocolo do mesmo documento no nome de outra pessoa não esconde o vencimento deste. Pendência de gente desligada (Baixas) não entra.

Células em controle manual ou não iniciadas (PACK = em fase de assinatura dos certificados; RECICLAR = reciclagem; X = não iniciado) NÃO são alteradas pela cobrança sem data, mas o documento AINDA aparece aqui. A célula X só vira VENCIDO quando há data de vencimento concreta; se o Certronic apenas cobra (Pendiente/Faltante), a célula permanece X e o caso segue como pendência. Ou seja: nesta aba pode haver pendência cujo status_celula na situacao_atual ainda é PACK, RECICLAR ou NAO_INICIADO (X) — é esperado, é um documento pendente aguardando ação manual.

## PARTE 3 — MODELAGEM E INDICADORES

## 3.1 Modelo relacional sugerido (estrela)

- colaboradores[cpf] 1 → * situacao_atual[cpf]
- colaboradores[cpf] 1 → * pendencias[cpf]
- colaboradores[cpf] 1 → * protocolos[cpf]  (só os com CPF preenchido)
- Crie uma dCalendario e relacione com data_execucao (situacao_atual e pendencias).
- visao_colaborador pode ficar desconectada (tabela de consumo direto) ou relacionada por cpf para enriquecer tooltips.

## 3.2 Indicadores sugeridos (o que a gestão vai querer ver)

| Indicador | Como calcular | Fonte |
|---|---|---|
| % de conformidade geral | Linhas OK ÷ linhas aplicáveis (excluir N/A), na última data_execucao. | situacao_atual |
| Conformidade por regional / sítio / função | Mesma conta, agrupada pelos atributos da dimensão. | situacao_atual + colaboradores |
| Colaboradores 100% conformes | Contagem de pct_conforme = 100. | visao_colaborador |
| Ranking de risco | Ordenar por docs_com_pendencia (desc) / pct_conforme (asc). | visao_colaborador |
| Evolução de pendências | Contagem de linhas por data_execucao (tendência semanal). | pendencias |
| Pendências por tipo de documento | Contagem agrupada por documento (ex: Onboarding Meli costuma dominar). | pendencias |
| Tempo médio de aprovação | Média de (data_acusado − data_enviado) onde estado = Aceptada. | protocolos |
| Taxa de rejeição | Rechazada ÷ (Aceptada + Rechazada + AceptadaParcial), por documento. | protocolos |
| Vencimentos próximos | data_vencimento entre hoje e hoje+30/60/90, status OK, última execução. | situacao_atual |
| Colaboradores bloqueáveis | Pessoas com ≥1 pendência crítica na última execução (risco de bloqueio de acesso/pagamento). | pendencias |

## 3.3 Cuidados na modelagem (leitura obrigatória)

- CPF é TEXTO com zeros à esquerda. No Power Query, defina o tipo como Texto logo no primeiro passo — se virar número, o zero some e os relacionamentos quebram.
- As datas vêm como texto e em formatos diferentes por origem: situacao_atual.data_execucao é aaaa-mm-dd; vencimentos e admissão são dd/mm/aaaa; datas de protocolos são dd-mm-aaaa. Converta com locale Português (Brasil) e trate erros.
- situacao_atual e pendencias ACUMULAM: quase toda medida precisa de filtro de data (última execução para foto atual; agrupamento por data para tendência). Medida sem filtro de data soma fotos repetidas e infla os números.
- colaboradores, documentos e visao_colaborador são SOBRESCRITAS: não têm histórico. Se um colaborador for desligado e removido da Matriz, ele sai dessas tabelas, mas o histórico dele permanece em situacao_atual e pendencias (nas datas antigas) — relacionamentos podem ter CPFs órfãos do lado histórico. Trate com relacionamento de cardinalidade adequada ou dimensão calculada.
- protocolos sem CPF: use a tabela inteira para métricas agregadas; filtre cpf preenchido para métricas por pessoa (a cobertura cresce com o tempo).
- O arquivo é gerado pelo programa — nunca edite à mão e nunca salve por cima. Se o Power BI apontar para ele no OneDrive/SharePoint, a atualização é automática.
- Rodadas no mesmo dia substituem o dia (sem duplicação) — mas se a operação rodar em horários diferentes, vale alinhar um horário padrão para a série ficar regular.

## 3.4 Glossário rápido

| Termo | Significado |
|---|---|
| Certronic | Sistema de validação documental usado pelo Mercado Livre. |
| Protocolo / ID | Número gerado pelo Certronic a cada envio de documento. |
| Matriz / BASE CERTRONIC_GPS | Planilha de controle da operação (colaborador × documento). |
| Aceptada / Rechazada / AceptadaParcial / Pendiente de Control | Vereditos do cliente sobre um protocolo (aprovado / reprovado / com ressalva / em análise). |
| Pendiente (relatório) vs PENDENTE (Matriz) | Pendiente é o rótulo do Certronic para documento vencido não renovado; PENDENTE é como isso aparece na Matriz. |
| PACK | Kit de certificados enviado ao colaborador para assinatura. |
| RECICLAR | Refazer treinamento vencido — decisão e registro manuais da operação. |
| Onboarding Meli | Integração obrigatória do colaborador com o Mercado Livre (Meli). |
| SSMA | Área de Saúde, Segurança e Meio Ambiente da IN HAUS (aplica os treinamentos). |
| Regional / Sítio | Divisões geográficas do Mercado Livre onde a IN HAUS atua / unidade específica. |
| Baixas / Desabilitada / Autorizada | Acesso no Certronic: desligado / sem acesso (mas ativo) / liberado. |
| NRs (NR10, NR35...) | Normas Regulamentadoras de segurança do trabalho — cada uma exige treinamento/certificado. |

## Histórico de versões

| Versão | Data | O que mudou |
|---|---|---|
| 1.0 | 14/07/2026 | Primeira versão. |
| 1.1 | 23/07/2026 | Pendências: supressão por protocolo passou a ser por pessoa (só a própria célula com o número); documentos vencidos em células PACK/RECICLAR agora aparecem nas pendências (célula não é alterada). base_indicadores também é atualizada ao Aplicar. |
| 1.2 | 28/07/2026 | Célula X (não iniciado) só vira VENCIDO quando há data; se o Certronic cobra sem data (Pendiente/Faltante), permanece X e segue como pendência. Pode haver pendência com status_celula = NAO_INICIADO. |
| 1.3 | 30/07/2026 | Protocolo aprovado preserva o número na célula ('OK 12345'); AceptadaParcial vira 'RESSALVA 12345', com nova categoria status_celula = OK_RESSALVA (conforme com ressalva). Cobertura de protocolo→CPF melhorada (lê também as células 'OK 12345'). |
