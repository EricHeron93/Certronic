# -*- coding: utf-8 -*-
"""Gera a apresentação corporativa do projeto (PowerPoint).

Uso:  python docs/gerar_apresentacao.py
Saída: C:/Projeto Certronic/Apresentacao Projeto Certronic.pptx

Fonte única do conteúdo (SLIDES abaixo). Ao mudar algo, edite e rode de novo.
"""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

# ---------------------------------------------------------------- paleta
# Paleta de Cores GPS (fornecida pela área)
AZUL = RGBColor(0x44, 0x72, 0xC4)       # azul primário GPS (4472C4)
NAVY = RGBColor(0x0B, 0x25, 0x36)       # azul escuro (0B2536) — capa/encerramento
AZUL_MEDIO = RGBColor(0x44, 0x54, 0x6A)  # azul acinzentado (44546A)
ACCENT = RGBColor(0x2B, 0x8D, 0xCF)     # azul claro de destaque (2B8DCF)
VERDE = RGBColor(0x70, 0xAD, 0x47)      # farol verde (70AD47) — ganhos/positivo
AMBER = RGBColor(0xFF, 0xC0, 0x00)      # farol amarelo (FFC000)
VERMELHO = RGBColor(0xDA, 0x29, 0x1C)   # farol vermelho (DA291C)
CINZA = RGBColor(0x55, 0x55, 0x55)      # cinza (555555)
CINZA_CLARO = RGBColor(0xF2, 0xF2, 0xF2)  # F2F2F2
AZUL_CLARO = RGBColor(0xD6, 0xDC, 0xE5)  # cinza azulado claro (D6DCE5)
BRANCO = RGBColor(0xFF, 0xFF, 0xFF)
GRAFITE = RGBColor(0x0B, 0x25, 0x36)    # texto principal = navy GPS

LARG = Inches(13.333)
ALT = Inches(7.5)
FONTE = "Calibri"

# Logo/imagem institucional do Grupo GPS (capa e encerramento).
# Salve a imagem exatamente neste caminho para ela ser incorporada.
IMG_CAPA = Path("C:/Projeto Certronic/gps_capa.png")

RODAPE = "Automação Certronic  ·  IN HAUS — Operação Mercado Livre"


def _fundo(slide, cor):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = cor


def _retangulo(slide, x, y, w, h, cor):
    from pptx.enum.shapes import MSO_SHAPE
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = cor
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def _texto(slide, x, y, w, h, texto, tam, cor, *, bold=False,
           align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, fonte=FONTE):
    cx = slide.shapes.add_textbox(x, y, w, h)
    tf = cx.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = texto
    r.font.size = Pt(tam)
    r.font.bold = bold
    r.font.name = fonte
    r.font.color.rgb = cor
    return cx


def _bullets(slide, x, y, w, h, itens, tam=18, cor=GRAFITE, cor_marca=AZUL):
    cx = slide.shapes.add_textbox(x, y, w, h)
    tf = cx.text_frame
    tf.word_wrap = True
    for i, item in enumerate(itens):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(10)
        p.line_spacing = 1.08
        # marcador
        m = p.add_run()
        m.text = "▪  "
        m.font.size = Pt(tam)
        m.font.name = FONTE
        m.font.color.rgb = cor_marca
        m.font.bold = True
        # negrito opcional: item pode ser (titulo, resto)
        if isinstance(item, tuple):
            forte, resto = item
            a = p.add_run()
            a.text = forte
            a.font.size = Pt(tam)
            a.font.name = FONTE
            a.font.bold = True
            a.font.color.rgb = cor
            b = p.add_run()
            b.text = resto
            b.font.size = Pt(tam)
            b.font.name = FONTE
            b.font.color.rgb = cor
        else:
            a = p.add_run()
            a.text = item
            a.font.size = Pt(tam)
            a.font.name = FONTE
            a.font.color.rgb = cor
    return cx


def _rodape(slide, numero):
    _texto(slide, Inches(0.6), Inches(7.05), Inches(10), Inches(0.35),
           RODAPE, 9, CINZA)
    _texto(slide, Inches(12.4), Inches(7.05), Inches(0.6), Inches(0.35),
           str(numero), 9, CINZA, align=PP_ALIGN.RIGHT)


# ----------------------------------------------------------- construtores
def slide_capa(prs, titulo, subtitulo, rodape):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _fundo(s, BRANCO)
    tem_img = IMG_CAPA.exists()
    larg_txt = Inches(7.1) if tem_img else Inches(11.5)
    _retangulo(s, Inches(0.9), Inches(2.35), Inches(1.4), Inches(0.09), VERDE)
    _texto(s, Inches(0.9), Inches(2.55), larg_txt, Inches(1.9),
           titulo, 38, NAVY, bold=True)
    _texto(s, Inches(0.95), Inches(4.45), larg_txt, Inches(0.9),
           subtitulo, 20, AZUL_MEDIO)
    _texto(s, Inches(0.95), Inches(6.45), larg_txt, Inches(0.5),
           rodape, 14, CINZA)
    if tem_img:
        s.shapes.add_picture(str(IMG_CAPA), Inches(8.45), Inches(0.9), height=Inches(5.7))


def slide_secao(prs, numero, titulo):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _fundo(s, CINZA_CLARO)
    _retangulo(s, 0, Inches(3.0), Inches(0.28), Inches(1.5), AZUL)
    _texto(s, Inches(0.9), Inches(2.9), Inches(2), Inches(1.5),
           f"{numero:02d}", 54, AZUL_CLARO, bold=True)
    _texto(s, Inches(2.3), Inches(3.05), Inches(10), Inches(1.4),
           titulo, 34, AZUL, bold=True, anchor=MSO_ANCHOR.MIDDLE)


def slide_conteudo(prs, numero, titulo, itens, *, tam=18, rotulo=None):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _fundo(s, BRANCO)
    _retangulo(s, 0, 0, LARG, Inches(1.25), AZUL)
    if rotulo:
        _texto(s, Inches(0.6), Inches(0.18), Inches(11), Inches(0.35),
               rotulo.upper(), 11, AZUL_CLARO, bold=True)
    _texto(s, Inches(0.6), Inches(0.48), Inches(12.1), Inches(0.7),
           titulo, 26, BRANCO, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    _bullets(s, Inches(0.8), Inches(1.7), Inches(11.7), Inches(5.0), itens, tam=tam)
    _rodape(s, numero)


def slide_metricas(prs, numero, titulo, cartoes, rodape_txt=""):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _fundo(s, BRANCO)
    _retangulo(s, 0, 0, LARG, Inches(1.25), AZUL)
    _texto(s, Inches(0.6), Inches(0.48), Inches(12.1), Inches(0.7),
           titulo, 26, BRANCO, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    n = len(cartoes)
    gap = Inches(0.35)
    marg = Inches(0.8)
    largura_util = LARG - marg * 2 - gap * (n - 1)
    cw = int(largura_util / n)
    y = Inches(2.1)
    ch = Inches(3.0)
    for i, (numero_g, rotulo) in enumerate(cartoes):
        x = int(marg) + i * (cw + int(gap))
        _retangulo(s, x, y, cw, ch, CINZA_CLARO)
        _retangulo(s, x, y, cw, Inches(0.12), VERDE)
        _texto(s, x, y + Inches(0.55), cw, Inches(1.2),
               numero_g, 40, AZUL, bold=True, align=PP_ALIGN.CENTER,
               anchor=MSO_ANCHOR.MIDDLE)
        _texto(s, x + Inches(0.2), y + Inches(1.75), cw - Inches(0.4), Inches(1.1),
               rotulo, 15, GRAFITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
    if rodape_txt:
        _texto(s, Inches(0.8), Inches(5.5), Inches(11.7), Inches(0.8),
               rodape_txt, 14, CINZA, align=PP_ALIGN.CENTER)
    _rodape(s, numero)


# ------------------------------------------------------------------ deck
def construir():
    prs = Presentation()
    prs.slide_width = LARG
    prs.slide_height = ALT

    slide_capa(
        prs,
        "Automação do Controle Documental Certronic",
        "Da planilha manual à gestão inteligente de conformidade",
        "IN HAUS  ·  Operação Mercado Livre  ·  Julho / 2026",
    )

    slide_conteudo(
        prs, 2, "O desafio", [
            ("Documentação é pré-requisito de acesso.  ",
             "Cada colaborador precisa de ~16 documentos regulatórios válidos "
             "(ASO, NRs, EPI, integração Meli...) para entrar nas unidades do Mercado Livre."),
            ("Documento vencido ou faltante custa caro.  ",
             "Pode bloquear o acesso do colaborador e travar o pagamento da IN HAUS pelo serviço."),
            ("O Certronic não dá visão do conjunto.  ",
             "O controle era feito à mão, cruzando 4 relatórios do sistema, célula por célula."),
            ("Processo lento e sujeito a erro.  ",
             "Horas de conferência manual repetitiva, sem histórico e sem visibilidade do risco."),
        ], rotulo="Contexto",
    )

    slide_conteudo(
        prs, 3, "O que estava em jogo", [
            ("Risco financeiro.  ", "Pagamento do serviço bloqueado pelo cliente."),
            ("Risco operacional.  ", "Colaborador impedido de acessar a unidade."),
            ("Falta de visibilidade.  ", "Difícil saber o que estava pendente, vencido ou a vencer."),
            ("Dependência do controle manual.  ",
             "Conhecimento concentrado e conferência demorada a cada rodada."),
        ], rotulo="O risco", tam=19,
    )

    slide_conteudo(
        prs, 4, "A solução", [
            ("Cruzamento automático.  ",
             "O programa lê os 4 relatórios do Certronic e atualiza a Matriz de controle sozinho."),
            ("Prévia antes de gravar.  ",
             "Mostra tudo o que vai mudar; nada é escrito sem a confirmação da operação."),
            ("Alerta de pendências críticas.  ",
             "Aponta documentos vencidos/faltantes que representam risco de bloqueio."),
            ("Base pronta para o Power BI.  ",
             "Gera indicadores de conformidade para a gestão acompanhar o risco."),
        ], rotulo="A proposta",
    )

    slide_conteudo(
        prs, 5, "Como funciona — 4 passos", [
            ("1.  Exportar.  ", "A operação baixa os 4 relatórios do Certronic para uma pasta padrão."),
            ("2.  Analisar.  ", "O programa cruza tudo em cerca de 1 minuto e monta a prévia."),
            ("3.  Conferir.  ", "A operação revisa as alterações e as pendências críticas."),
            ("4.  Aplicar.  ",
             "Grava na Matriz (SharePoint) e atualiza a base do Power BI — com registro de auditoria."),
        ], rotulo="Fluxo", tam=19,
    )

    slide_conteudo(
        prs, 6, "Principais recursos", [
            ("Prévia com filtro e ordenação por coluna.  ", "Localiza qualquer colaborador ou documento em segundos."),
            ("Seleção por linha.  ", "A operação aplica só as alterações que quiser."),
            ("Validação automática da planilha.  ", "Confere a estrutura da Matriz e evita rodar sobre o arquivo errado."),
            ("Feedback claro.  ", "Barra de progresso, mensagens amigáveis e aba de pendências críticas."),
            ("Configuração simples.  ", "Aponta as pastas uma vez por máquina — sem editar arquivos técnicos."),
        ], rotulo="Produto", tam=17,
    )

    slide_conteudo(
        prs, 7, "Inteligência das regras", [
            ("Análise por pessoa.  ",
             "Não esconde vencimentos reais — revelou documentos vencidos que estavam invisíveis no controle."),
            ("Respeita as etapas manuais.  ",
             "Células em PACK (assinatura) e RECICLAR não são sobrescritas, mas entram no alerta."),
            ("Nunca toca o cadastro.  ", "As colunas de dados do colaborador ficam protegidas."),
            ("Nunca “chuta”.  ", "Casos duvidosos são listados para a operação decidir, nunca adivinhados."),
            ("Prévia + confirmação obrigatórias.  ", "A operação mantém o controle da decisão final."),
        ], rotulo="Diferencial", tam=17,
    )

    slide_metricas(
        prs, 8, "Ganhos com o projeto", [
            ("minutos", "De horas de conferência manual para uma análise em minutos"),
            ("100%", "Rastreabilidade: relatório de auditoria a cada execução"),
            ("+risco\nvisível", "Pendências críticas destacadas antes do bloqueio"),
            ("Power BI", "Visibilidade da conformidade para a gestão"),
        ],
        rodape_txt="Menos esforço manual, menos erro, mais visibilidade e proteção do pagamento.",
    )

    slide_conteudo(
        prs, 9, "Indicadores para a gestão (Power BI)", [
            ("% de conformidade.  ", "Geral e por regional, sítio e função."),
            ("Ranking de risco.  ", "Colaboradores com mais documentos pendentes."),
            ("Evolução de pendências.  ", "Tendência ao longo do tempo."),
            ("Vencimentos próximos.  ", "Documentos a vencer em 30 / 60 / 90 dias."),
            ("Risco de bloqueio.  ", "Colaboradores com pendência crítica na última execução."),
        ], rotulo="Indicadores", tam=17,
    )

    slide_conteudo(
        prs, 10, "Segurança e governança", [
            ("Nada é gravado sem confirmação.  ", "A prévia sempre vem antes da escrita."),
            ("Auditoria por execução.  ", "Registro do que mudou, das pendências e dos casos pulados."),
            ("Cadastro intocável.  ", "As colunas de identificação do colaborador nunca são alteradas."),
            ("Dados no SharePoint.  ", "Matriz e indicadores na pasta corporativa, via OneDrive."),
        ], rotulo="Governança", tam=18,
    )

    slide_conteudo(
        prs, 11, "Onde estamos", [
            ("Solução construída e testada com dados reais.  ", "Fluxo completo validado ponta a ponta."),
            ("Integração com o SharePoint validada.  ", "Gravação na pasta corporativa funcionando."),
            ("Documentação completa.  ", "Manual da operação, documentação de BI e especificação do projeto."),
            ("Próxima etapa: executável + homologação.  ", "Instalação na máquina da operação e rodada assistida."),
        ], rotulo="Status", tam=18,
    )

    slide_conteudo(
        prs, 12, "Próximos passos", [
            ("Empacotar o executável.  ", "Instalação simples na máquina da operação."),
            ("Homologação assistida.  ", "Rodada real acompanhada com a operação."),
            ("Início da rotina oficial.  ", "Definição da frequência de execução."),
            ("Handoff do Power BI.  ", "Entrega da base ao analista para construção do dashboard."),
        ], rotulo="Roadmap", tam=18,
    )

    # encerramento
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _fundo(s, BRANCO)
    tem_img = IMG_CAPA.exists()
    larg_txt = Inches(6.9) if tem_img else Inches(11.5)
    _retangulo(s, Inches(0.9), Inches(2.75), Inches(1.4), Inches(0.09), VERDE)
    _texto(s, Inches(0.9), Inches(2.95), larg_txt, Inches(1.0),
           "Obrigado", 40, NAVY, bold=True)
    _texto(s, Inches(0.95), Inches(4.05), larg_txt, Inches(1.6),
           "A conformidade documental deixa de ser um controle manual e passa a "
           "ser um processo rápido, seguro e visível.", 20, AZUL_MEDIO)
    if tem_img:
        s.shapes.add_picture(str(IMG_CAPA), Inches(8.7), Inches(1.35), height=Inches(4.8))

    destino = Path("C:/Projeto Certronic/Apresentacao Projeto Certronic.pptx")
    destino.parent.mkdir(parents=True, exist_ok=True)
    prs.save(destino)
    return destino


if __name__ == "__main__":
    print("Apresentação:", construir())
