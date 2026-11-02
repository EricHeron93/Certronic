# Automação Certronic → Matriz (BASE CERTRONIC_GPS)

Ferramenta interna da IN HAUS para atualizar automaticamente a planilha de
controle de documentação dos colaboradores alocados no Mercado Livre,
a partir dos 4 relatórios exportados do Certronic.

## Como funciona

1. A operação exporta os 4 relatórios do Certronic e salva na pasta de
   entrada com os nomes padrão (ver `config.yaml`).
2. Abre o programa e clica em **Analisar relatórios**: o programa cruza
   tudo e mostra a prévia do que vai mudar — nada é escrito ainda.
3. Confere a prévia e clica em **Aplicar na Matriz** para confirmar.
4. Cada execução gera um relatório em texto com as alterações, as
   pendências críticas e os casos que o programa pulou.

## Regras invioláveis

- Colunas **A–I** (cadastro) nunca são alteradas — a operação segue
  preenchendo manualmente.
- Células com **RECICLAR** nunca são tocadas, e o programa **não** calcula
  prazo de reciclagem — decisão 100% manual da operação.
- Nada é escrito na Matriz sem confirmação explícita na tela.
- Caso não classificável (documento fora do de-para, CPF fora da Matriz)
  é pulado e listado no relatório — nunca "chutado".

## O que cada relatório alimenta

| Relatório | Papel |
|---|---|
| `documentacao_enviada` | Estado dos protocolos: Aceptada→OK, Rechazada→[ID] X, AceptadaParcial→OK (RESSALVA) + observações |
| `estado_vencimento` | Documentos OK vencendo em 7–30 dias → OK (X MESES) / OK (V:data) |
| `documentacao_faltante` | Pendências críticas (vencido/nunca enviado) — cruzadas com protocolos em análise para não dar alarme falso |
| `lista_pessoal` | Filtra desligados (Baixas) para não gerar pendência de ex-funcionário |

## Rodando em desenvolvimento

```
pip install openpyxl pyyaml
python main.py            # usa o config.yaml ao lado do main.py
python main.py outro.yaml # ou aponta outro config
```

## Gerando o executável (para o computador da Aline)

```
pip install pyinstaller
pyinstaller --onefile --windowed --add-data "config.yaml;." main.py
```

O executável fica em `dist/main.exe`. O `config.yaml` pode ficar ao lado
dele para a operação ajustar caminhos e o de-para sem recompilar.
