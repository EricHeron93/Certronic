"""Front simples (Tkinter) para a operação rodar a atualização.

Fluxo: [Analisar] → tabela de prévia com tudo que vai mudar → [Aplicar].
Nada é escrito na Matriz sem a confirmação explícita da operação.
"""

from __future__ import annotations

import re
import threading
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import yaml

from . import exportador
from .leitores import ativos, faltantes, protocolos, vencimentos
from .matriz import Matriz
from .regras import ResultadoAnalise, analisar
from .relatorio import gerar as gerar_relatorio


class App:
    def __init__(self, caminho_config: str | Path):
        self.caminho_config = Path(caminho_config)
        self.cfg = yaml.safe_load(self.caminho_config.read_text(encoding="utf-8"))
        self.resultado: ResultadoAnalise | None = None
        self.matriz: Matriz | None = None
        # Uma marca por alteração proposta (True = será gravada na Matriz).
        # A operação pode desmarcar linhas individuais antes de aplicar.
        self.alt_selecionado: list[bool] = []
        # (x_root inicial, largura inicial) durante o arraste da alça de coluna
        self._resize_estado: tuple[int, int] = (0, 0)
        # Ordenação atual de cada aba: (chave_coluna, crescente) ou None
        self._sort_alt: tuple[str, bool] | None = None
        self._sort_pend: tuple[str, bool] | None = None
        # Colunas de documento esperadas que faltaram na última análise
        self._avisos_estrutura: list[str] = []

        self.root = tk.Tk()
        self.root.title("Atualização Matriz Certronic — IN HAUS")
        self.root.geometry("1000x600")
        self._montar_tela()

        # Cria as pastas de saída se possível e valida os caminhos. Numa máquina
        # nova (caminhos de outra máquina), leva direto para a aba Configurações.
        self._garantir_pastas()
        self._validar_config()

    def _montar_tela(self):
        topo = ttk.Frame(self.root, padding=10)
        topo.pack(fill="x")

        self.btn_analisar = ttk.Button(topo, text="1. Analisar relatórios", command=self.analisar)
        self.btn_analisar.pack(side="left", padx=5)

        self.btn_aplicar = ttk.Button(
            topo, text="2. Aplicar na Matriz", command=self.aplicar, state="disabled"
        )
        self.btn_aplicar.pack(side="left", padx=5)

        self.lbl_status = ttk.Label(topo, text="Aguardando análise…")
        self.lbl_status.pack(side="left", padx=15)

        # Barra de progresso animada (indeterminada) — empacotada só enquanto
        # uma tarefa longa roda (ver _iniciar_progresso/_parar_progresso).
        self.progress = ttk.Progressbar(topo, mode="indeterminate", length=160)

        abas = ttk.Notebook(self.root)
        abas.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # --- aba 1: alterações que serão escritas na Matriz -----------------
        frame_alt = ttk.Frame(abas)
        abas.add(frame_alt, text="Alterações na Matriz")

        ctrl_alt = ttk.Frame(frame_alt, padding=(0, 6))
        ctrl_alt.pack(fill="x")
        self.var_sel_todos = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            ctrl_alt, text="Selecionar todos (visíveis)",
            variable=self.var_sel_todos, command=self._toggle_todos_alt,
        ).pack(side="left")
        self.lbl_sel = ttk.Label(ctrl_alt, text="")
        self.lbl_sel.pack(side="right")

        # tabela sem cabeçalho nativo (show="") — usamos um cabeçalho próprio,
        # com título em cima e a caixa de filtro logo abaixo (montado adiante).
        corpo_alt = ttk.Frame(frame_alt)
        colunas = ("sel", "colaborador", "cpf", "documento", "de", "para", "motivo")
        self.tabela = ttk.Treeview(corpo_alt, columns=colunas, show="")
        rol_alt = ttk.Scrollbar(corpo_alt, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=rol_alt.set)
        rol_alt.pack(side="right", fill="y")
        self.tabela.pack(side="left", fill="both", expand=True)
        self.tabela.column("sel", width=36, anchor="center", stretch=False)
        for c in ("colaborador", "cpf", "documento", "de", "para", "motivo"):
            self.tabela.column(c, stretch=False)
        # clique na coluna do checkbox marca/desmarca aquela linha
        self.tabela.bind("<Button-1>", self._on_click_alt)

        self.filtros_alt, self._labels_alt, self._titulos_alt = self._montar_cabecalho(
            frame_alt, self.tabela,
            [("sel", "✓", 36, False), ("colaborador", "Colaborador", 190, True),
             ("cpf", "CPF", 95, True), ("documento", "Documento", 170, True),
             ("de", "De", 95, True), ("para", "Para", 95, True),
             ("motivo", "Motivo", 280, True)],
            self._popular_alteracoes,
            lambda c: self._ordenar(c, "_sort_alt", self._labels_alt,
                                    self._titulos_alt, self._popular_alteracoes),
        )
        corpo_alt.pack(fill="both", expand=True)  # abaixo do cabeçalho

        # --- aba 2: pendências críticas (documentacao_faltante) -------------
        frame_pend = ttk.Frame(abas)
        abas.add(frame_pend, text="Pendências críticas")

        corpo_pend = ttk.Frame(frame_pend)
        colunas_p = ("colaborador", "cpf", "documento", "situacao")
        self.tabela_pend = ttk.Treeview(corpo_pend, columns=colunas_p, show="")
        rol_pend = ttk.Scrollbar(corpo_pend, orient="vertical", command=self.tabela_pend.yview)
        self.tabela_pend.configure(yscrollcommand=rol_pend.set)
        rol_pend.pack(side="right", fill="y")
        self.tabela_pend.pack(side="left", fill="both", expand=True)
        for c in colunas_p:
            self.tabela_pend.column(c, stretch=False)

        self.filtros_pend, self._labels_pend, self._titulos_pend = self._montar_cabecalho(
            frame_pend, self.tabela_pend,
            [("colaborador", "Colaborador", 220, True), ("cpf", "CPF", 100, True),
             ("documento", "Documento", 280, True), ("situacao", "Situação", 370, True)],
            self._popular_pendencias,
            lambda c: self._ordenar(c, "_sort_pend", self._labels_pend,
                                    self._titulos_pend, self._popular_pendencias),
        )
        corpo_pend.pack(fill="both", expand=True)  # abaixo do cabeçalho

        # --- aba 3: configurações (caminhos por máquina) --------------------
        self._tab_cfg = self._montar_aba_config(abas)

        self._abas = abas

    # ------------------------------------------------------------------
    # Aba de configurações — caminhos são específicos de cada máquina
    # (usuário + pasta sincronizada do SharePoint). Configura-se uma vez.
    # ------------------------------------------------------------------
    _CAMPOS_CONFIG = [
        ("pasta_entrada", "Pasta de entrada (onde ficam os 4 relatórios do Certronic)", "dir"),
        ("matriz", "Arquivo da Matriz (BASE CERTRONIC_GPS.xlsx)", "file_open"),
        ("base_indicadores", "Arquivo da base de indicadores (fonte do Power BI)", "file_save"),
        ("pasta_relatorios", "Pasta dos relatórios de execução", "dir"),
    ]

    def _montar_aba_config(self, abas):
        frame = ttk.Frame(abas, padding=15)
        abas.add(frame, text="Configurações")

        ttk.Label(
            frame, wraplength=940, justify="left",
            text=("Aponte cada caminho para a pasta do SharePoint sincronizada nesta "
                  "máquina (aparece no Explorer como uma pasta local). O programa grava "
                  "no caminho local e o OneDrive envia para o SharePoint sozinho. Estes "
                  "caminhos valem só para ESTA máquina — em outro computador, configure "
                  "de novo aqui."),
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        self.cfg_vars: dict[str, tk.StringVar] = {}
        for i, (chave, rotulo, tipo) in enumerate(self._CAMPOS_CONFIG):
            linha = 1 + i * 2
            ttk.Label(frame, text=rotulo).grid(row=linha, column=0, columnspan=2,
                                               sticky="w", pady=(8, 0))
            var = tk.StringVar(value=str(self.cfg["caminhos"].get(chave, "")))
            self.cfg_vars[chave] = var
            ttk.Entry(frame, textvariable=var).grid(row=linha + 1, column=0,
                                                    sticky="we", padx=(0, 6))
            ttk.Button(frame, text="Procurar…",
                       command=lambda c=chave, t=tipo: self._procurar_caminho(c, t)
                       ).grid(row=linha + 1, column=1)

        frame.columnconfigure(0, weight=1)

        barra = ttk.Frame(frame)
        barra.grid(row=1 + len(self._CAMPOS_CONFIG) * 2, column=0, columnspan=2,
                   sticky="w", pady=16)
        ttk.Button(barra, text="Salvar configurações", command=self._salvar_config).pack(side="left")
        ttk.Button(barra, text="Abrir pasta de entrada",
                   command=self._abrir_pasta_entrada).pack(side="left", padx=8)
        self.lbl_cfg = ttk.Label(barra, text="")
        self.lbl_cfg.pack(side="left", padx=10)
        return frame

    def _procurar_caminho(self, chave: str, tipo: str):
        atual = self.cfg_vars[chave].get().strip()
        inicial = atual if Path(atual).parent.exists() else ""
        if tipo == "dir":
            escolhido = filedialog.askdirectory(title="Escolha a pasta", initialdir=inicial or None)
        elif tipo == "file_open":
            escolhido = filedialog.askopenfilename(
                title="Escolha o arquivo", initialdir=inicial or None,
                filetypes=[("Planilha Excel", "*.xlsx"), ("Todos", "*.*")])
        else:  # file_save
            escolhido = filedialog.asksaveasfilename(
                title="Escolha onde salvar", initialdir=inicial or None,
                defaultextension=".xlsx",
                filetypes=[("Planilha Excel", "*.xlsx"), ("Todos", "*.*")])
        if escolhido:
            self.cfg_vars[chave].set(escolhido.replace("\\", "/"))

    def _salvar_config(self):
        novos = {c: v.get().strip().replace("\\", "/") for c, v in self.cfg_vars.items()}
        if any(not p for p in novos.values()):
            messagebox.showwarning("Configuração incompleta",
                                   "Preencha todos os quatro caminhos antes de salvar.")
            return
        self.cfg["caminhos"].update(novos)
        # Regrava só as linhas dos caminhos, preservando o resto do arquivo
        # (comentários, de-para de documentos, regras) intacto.
        texto = self.caminho_config.read_text(encoding="utf-8")
        for chave, valor in novos.items():
            texto = re.sub(rf'^(\s*{chave}:\s*).*$', rf'\g<1>"{valor}"',
                           texto, count=1, flags=re.M)
        self.caminho_config.write_text(texto, encoding="utf-8")
        self._garantir_pastas()
        self.lbl_cfg.config(text="Configurações salvas.")
        if self._validar_config():
            self.lbl_status.config(text="Configuração OK. Clique em Analisar relatórios.")

    def _abrir_pasta_entrada(self):
        pasta = Path(self.cfg_vars["pasta_entrada"].get().strip())
        if pasta.parent.exists():
            pasta.mkdir(parents=True, exist_ok=True)
            import os
            os.startfile(pasta)  # abre no Explorer (Windows)
        else:
            messagebox.showwarning(
                "Pasta indisponível",
                "A pasta-mãe não existe nesta máquina. Confira se o SharePoint está "
                "sincronizado e ajuste o caminho.")

    def _garantir_pastas(self):
        """Cria a pasta de entrada e a de relatórios se ainda não existirem —
        mas só quando a pasta-mãe existe (evita criar uma árvore inválida a
        partir do caminho de outra máquina)."""
        for chave in ("pasta_entrada", "pasta_relatorios"):
            p = Path(self.cfg["caminhos"][chave])
            if p.parent.exists() and not p.exists():
                try:
                    p.mkdir(parents=True, exist_ok=True)
                except OSError:
                    pass

    def _validar_config(self) -> bool:
        """Confere se a Matriz e a pasta de entrada existem. Se não, leva para
        a aba Configurações e bloqueia o Analisar. Retorna True se estiver ok."""
        c = self.cfg["caminhos"]
        faltas = []
        if not Path(c["matriz"]).exists():
            faltas.append("a Matriz")
        if not Path(c["pasta_entrada"]).exists():
            faltas.append("a pasta de entrada")
        if faltas:
            self.btn_analisar.config(state="disabled")
            self._abas.select(self._tab_cfg)
            self.lbl_status.config(
                text="Configure os caminhos na aba Configurações — não encontrei "
                     + " e ".join(faltas) + " nesta máquina.")
            return False
        self.btn_analisar.config(state="normal")
        return True

    def _montar_cabecalho(self, parent, tree, colspec, on_change, on_sort):
        """Cabeçalho próprio para a tabela: por coluna, um título em cima e a
        caixa de filtro logo abaixo, com uma alça à direita para redimensionar.
        Clicar no título chama on_sort(chave) para ordenar por essa coluna.

        Substitui o cabeçalho nativo do Treeview (que não deixa inserir a linha
        de filtros entre o título e os dados). Mantém a largura das colunas da
        tabela em sincronia com o arraste da alça.

        colspec: lista ordenada de (chave, titulo, largura_px, filtravel).
        Retorna dict chave->StringVar dos filtros.
        """
        ordem = [c[0] for c in colspec]
        larguras = {c[0]: c[2] for c in colspec}
        titulos = {c[0]: c[1] for c in colspec}
        filtravel = {c[0]: c[3] for c in colspec}

        head = tk.Frame(parent, height=52)
        head.pack(fill="x")
        head.pack_propagate(False)  # respeita a altura fixa

        variaveis, labels, entries, seps = {}, {}, {}, {}

        def layout(_e=None):
            x = 0
            for col in ordem:
                w = larguras[col]
                labels[col].place(x=x, y=0, width=w, height=23)
                if col in entries:
                    entries[col].place(x=x + 2, y=25, width=w - 4, height=24)
                seps[col].place(x=x + w - 3, y=0, width=6, height=52)
                tree.column(col, width=w)  # mantém a tabela do mesmo tamanho
                x += w
            for s in seps.values():
                s.tkraise()  # alças sempre clicáveis, acima dos títulos

        for col in ordem:
            labels[col] = tk.Label(
                head, text=titulos[col], anchor="w", padx=5,
                relief="raised", bd=1, bg="#e1e1e1", font=("Segoe UI", 9, "bold"),
            )
            if filtravel[col]:
                var = tk.StringVar()
                caixa = ttk.Entry(head, textvariable=var)
                caixa.bind("<KeyRelease>", lambda _e: on_change())
                variaveis[col] = var
                entries[col] = caixa
                # clicar no título ordena por essa coluna (só colunas de dados)
                labels[col].config(cursor="hand2")
                labels[col].bind("<Button-1>", lambda _e, c=col: on_sort(c))
            seps[col] = tk.Frame(head, cursor="sb_h_double_arrow")

            def _press(ev, c=col):
                self._resize_estado = (ev.x_root, larguras[c])

            def _drag(ev, c=col):
                x0, w0 = self._resize_estado
                larguras[c] = max(40, w0 + (ev.x_root - x0))
                layout()

            seps[col].bind("<Button-1>", _press)
            seps[col].bind("<B1-Motion>", _drag)

        layout()
        return variaveis, labels, titulos

    def _ordenar(self, chave, estado_attr, labels, titulos, popular):
        """Alterna a ordenação da aba pela coluna `chave` (1º clique crescente,
        clique seguinte na mesma coluna inverte) e redesenha a tabela."""
        estado = getattr(self, estado_attr)
        if estado and estado[0] == chave:
            novo = (chave, not estado[1])
        else:
            novo = (chave, True)
        setattr(self, estado_attr, novo)
        # atualiza a setinha ▲/▼ nos títulos
        for c, lbl in labels.items():
            if c in titulos:
                seta = ""
                if c == novo[0]:
                    seta = "  ▲" if novo[1] else "  ▼"
                lbl.config(text=titulos[c] + seta)
        popular()

    # ------------------------------------------------------------------
    # População das tabelas (com filtro) e seleção de linhas
    # ------------------------------------------------------------------
    # índice de cada campo na tupla de valores, por chave de coluna
    _CAMPOS_ALT = {"colaborador": 0, "cpf": 1, "documento": 2, "de": 3, "para": 4, "motivo": 5}
    _CAMPOS_PEND = {"colaborador": 0, "cpf": 1, "documento": 2, "situacao": 3}

    @staticmethod
    def _passa_filtros(campos, filtros, indices) -> bool:
        """True se a linha satisfaz TODOS os filtros de coluna preenchidos."""
        return all(
            texto in str(campos[indices[chave]]).lower()
            for chave, texto in filtros.items()
        )

    def _popular_alteracoes(self):
        """Redesenha a aba de alterações aplicando os filtros por coluna.
        O iid de cada linha é o índice na lista original, para o clique
        mapear de volta mesmo com os filtros ativos."""
        if not self.resultado:
            return
        self.tabela.delete(*self.tabela.get_children())
        filtros = {k: v.get().strip().lower() for k, v in self.filtros_alt.items() if v.get().strip()}
        itens = []
        for i, a in enumerate(self.resultado.alteracoes):
            campos = (a.nome, a.cpf, a.documento, a.valor_atual, a.valor_novo, a.motivo)
            if not self._passa_filtros(campos, filtros, self._CAMPOS_ALT):
                continue
            itens.append((i, campos))
        if self._sort_alt:
            chave, cresc = self._sort_alt
            pos = self._CAMPOS_ALT[chave]
            itens.sort(key=lambda t: str(t[1][pos]).lower(), reverse=not cresc)
        for i, campos in itens:
            glyph = "☑" if self.alt_selecionado[i] else "☐"
            self.tabela.insert("", "end", iid=str(i), values=(glyph, *campos))
        self._atualizar_status_sel()

    def _popular_pendencias(self):
        if not self.resultado:
            return
        self.tabela_pend.delete(*self.tabela_pend.get_children())
        filtros = {k: v.get().strip().lower() for k, v in self.filtros_pend.items() if v.get().strip()}
        itens = []
        for j, p in enumerate(self.resultado.pendencias):
            campos = (p.nome, p.cpf, p.documento, p.descricao)
            if not self._passa_filtros(campos, filtros, self._CAMPOS_PEND):
                continue
            itens.append((j, campos))
        if self._sort_pend:
            chave, cresc = self._sort_pend
            pos = self._CAMPOS_PEND[chave]
            itens.sort(key=lambda t: str(t[1][pos]).lower(), reverse=not cresc)
        for j, campos in itens:
            self.tabela_pend.insert("", "end", iid=str(j), values=campos)

    def _on_click_alt(self, event):
        """Alterna a marca da linha quando o clique cai na coluna do checkbox."""
        if self.tabela.identify_region(event.x, event.y) != "cell":
            return
        if self.tabela.identify_column(event.x) != "#1":
            return
        iid = self.tabela.identify_row(event.y)
        if not iid:
            return
        i = int(iid)
        self.alt_selecionado[i] = not self.alt_selecionado[i]
        self.tabela.set(iid, "sel", "☑" if self.alt_selecionado[i] else "☐")
        self._atualizar_status_sel()

    def _toggle_todos_alt(self):
        """Marca/desmarca todas as linhas atualmente visíveis (respeita o filtro)."""
        if not self.resultado:
            return
        marcar = self.var_sel_todos.get()
        for iid in self.tabela.get_children():
            i = int(iid)
            self.alt_selecionado[i] = marcar
            self.tabela.set(iid, "sel", "☑" if marcar else "☐")
        self._atualizar_status_sel()

    def _atualizar_status_sel(self):
        total = len(self.resultado.alteracoes) if self.resultado else 0
        marcadas = sum(self.alt_selecionado) if self.resultado else 0
        self.lbl_sel.config(text=f"{marcadas} de {total} marcadas para aplicar")

    def _iniciar_progresso(self, texto: str):
        """Desabilita os botões, mostra o status e liga a barra animada.
        Chamado sempre antes de disparar uma tarefa longa em segundo plano."""
        self.btn_analisar.config(state="disabled")
        self.btn_aplicar.config(state="disabled")
        self.lbl_status.config(text=texto)
        self.progress.pack(side="left", padx=8)
        self.progress.start(12)

    def _parar_progresso(self):
        """Para e esconde a barra animada. Chamado ao terminar (ok ou erro)."""
        self.progress.stop()
        self.progress.pack_forget()

    def analisar(self):
        # A Matriz é um arquivo grande (centenas de milhares de linhas de
        # formatação residual) e leva ~1 minuto para carregar. A análise roda
        # em segundo plano para a janela não congelar; a barra de progresso
        # mostra que está trabalhando.
        self._iniciar_progresso("Analisando… a Matriz é grande, isso leva cerca de 1 minuto.")
        threading.Thread(target=self._analisar_worker, daemon=True).start()

    def _analisar_worker(self):
        """Roda fora da thread da interface. Toda atualização de tela volta
        pela fila do Tkinter (root.after)."""
        try:
            pasta = Path(self.cfg["caminhos"]["pasta_entrada"])
            arq = self.cfg["arquivos"]

            self.dados_protocolos = protocolos.ler(pasta / arq["protocolos"])
            self.dados_vencimentos = vencimentos.ler(pasta / arq["vencimentos"])
            self.dados_ativos = ativos.ler(pasta / arq["ativos"])

            self.matriz = Matriz(self.cfg["caminhos"]["matriz"], self.cfg["matriz"])

            # Validação de estrutura (reaproveita a Matriz já carregada — sem
            # custo extra de leitura). Problema crítico interrompe antes de
            # analisar; colunas faltando viram aviso no fim.
            criticos, self._avisos_estrutura = self.matriz.validar_estrutura(
                self.cfg["documentos"])
            if criticos:
                self.root.after(0, self._analise_falhou,
                                "Matriz em formato inesperado",
                                "A planilha da Matriz não está no formato esperado:\n\n- "
                                + "\n- ".join(criticos)
                                + "\n\nConfira o arquivo e a aba na aba Configurações.")
                return

            self.resultado = analisar(
                matriz=self.matriz,
                protocolos=self.dados_protocolos,
                vencimentos=self.dados_vencimentos,
                faltantes=faltantes.ler(pasta / arq["faltantes"]),
                ativos=self.dados_ativos,
                cfg=self.cfg,
                hoje=date.today(),
            )
        except FileNotFoundError as e:
            self.root.after(0, self._analise_falhou, "Arquivo não encontrado",
                            f"{e}\n\nConfira se os 4 relatórios estão na pasta de "
                            "entrada com os nomes padrão (ver manual, item 3).")
            return
        except PermissionError as e:
            self.root.after(0, self._analise_falhou, "Arquivo aberto no Excel",
                            f"Não consegui ler um dos arquivos porque ele está aberto:\n"
                            f"{e.filename}\n\nFeche o arquivo no Excel e clique em "
                            "Analisar de novo.")
            return
        except Exception as e:
            self.root.after(0, self._analise_falhou, "Erro na análise", str(e))
            return

        # Base estruturada para o Power BI (dimensões + fatos históricos).
        # Qualquer problema aqui não pode derrubar a análise — só avisa.
        aviso_base = self._gerar_base_indicadores()

        self.root.after(0, self._analise_concluida, aviso_base)

    def _gerar_base_indicadores(self) -> str:
        """Atualiza o base_indicadores.xlsx a partir do estado atual da Matriz
        em memória (self.matriz). Retorna uma mensagem de aviso, vazia se deu
        tudo certo — nunca levanta exceção, para não derrubar quem chamou."""
        try:
            exportador.gerar(
                self.cfg["caminhos"]["base_indicadores"],
                self.matriz, self.dados_protocolos, self.dados_vencimentos,
                self.dados_ativos, self.resultado, self.cfg,
            )
            return ""
        except PermissionError:
            return (
                "A base_indicadores.xlsx está aberta no Excel e não pôde ser "
                "atualizada.\nFeche o arquivo e clique em Analisar de novo para "
                "atualizar a base."
            )
        except Exception as e:
            return f"Não foi possível atualizar a base_indicadores.xlsx:\n{e}"

    def _analise_falhou(self, titulo: str, mensagem: str):
        self._parar_progresso()
        self.btn_analisar.config(state="normal")
        self.lbl_status.config(text="Análise interrompida — veja a mensagem.")
        messagebox.showerror(titulo, mensagem)

    def _analise_concluida(self, aviso_base: str):
        self._parar_progresso()
        self.btn_analisar.config(state="normal")
        if aviso_base:
            messagebox.showwarning("Base de indicadores não atualizada", aviso_base)

        # Todas as alterações começam marcadas; a operação desmarca o que não quer.
        self.alt_selecionado = [True] * len(self.resultado.alteracoes)
        self.var_sel_todos.set(True)
        for v in self.filtros_alt.values():
            v.set("")
        for v in self.filtros_pend.values():
            v.set("")
        self._popular_alteracoes()
        self._popular_pendencias()

        r = self.resultado
        self._abas.tab(1, text=f"Pendências críticas ({len(r.pendencias)})")
        self.lbl_status.config(
            text=f"{len(r.alteracoes)} alterações | {len(r.pendencias)} pendências críticas | {len(r.incertos)} casos pulados"
        )
        self.btn_aplicar.config(state="normal" if r.alteracoes else "disabled")

        # Prévia sempre gera relatório, aplicada ou não
        gerar_relatorio(r, self.cfg["caminhos"]["pasta_relatorios"], aplicado=False)

        if self._avisos_estrutura:
            messagebox.showwarning(
                "Colunas de documento não encontradas",
                "Estas colunas esperadas não foram achadas no cabeçalho da Matriz — "
                "os documentos ligados a elas serão ignorados nesta análise:\n\n- "
                + "\n- ".join(self._avisos_estrutura)
                + "\n\nSe os nomes das colunas mudaram na Matriz, ajuste o de-para "
                "no config.yaml.",
            )

        if r.pendencias:
            messagebox.showwarning(
                "Pendências críticas",
                f"{len(r.pendencias)} documento(s) vencido(s)/faltante(s) sem protocolo em análise.\n"
                "Veja o relatório — risco de bloqueio de pagamento.",
            )

    def aplicar(self):
        if not self.resultado or not self.matriz:
            return
        selecionadas = [
            a for i, a in enumerate(self.resultado.alteracoes) if self.alt_selecionado[i]
        ]
        n, total = len(selecionadas), len(self.resultado.alteracoes)
        if n == 0:
            messagebox.showinfo(
                "Nada selecionado",
                "Nenhuma alteração está marcada. Marque ao menos uma linha para aplicar.",
            )
            return
        msg = f"Aplicar {n} de {total} alterações na Matriz?"
        if n < total:
            msg += f"\n({total - n} desmarcada(s) NÃO serão gravadas.)"
        if not messagebox.askyesno("Confirmar", msg):
            return
        # Gravar a Matriz (arquivo grande) e regerar a base demora alguns
        # segundos; roda em segundo plano com a barra de progresso ligada para
        # a operação ter feedback de que está trabalhando.
        self._iniciar_progresso("Aplicando na Matriz… gravando a planilha e atualizando a base. Aguarde.")
        threading.Thread(
            target=self._aplicar_worker, args=(selecionadas, n, total), daemon=True
        ).start()

    def _aplicar_worker(self, selecionadas, n, total):
        """Roda fora da thread da interface. Volta pela fila do Tkinter."""
        try:
            self.matriz.aplicar(selecionadas)
            self.matriz.salvar()
        except PermissionError:
            self.root.after(0, self._aplicar_falhou, "Planilha aberta",
                            "Não foi possível salvar: a Matriz está aberta no Excel. "
                            "Feche e tente de novo.")
            return
        except Exception as e:
            self.root.after(0, self._aplicar_falhou, "Erro ao aplicar", str(e))
            return

        # A Matriz em memória já reflete as alterações recém-gravadas, então a
        # base_indicadores.xlsx sai atualizada sem precisar reanalisar.
        aviso_base = self._gerar_base_indicadores()

        # O relatório de auditoria registra só o que foi de fato gravado.
        res_aplicado = ResultadoAnalise(
            alteracoes=selecionadas,
            pendencias=self.resultado.pendencias,
            incertos=self.resultado.incertos,
        )
        try:
            caminho = gerar_relatorio(
                res_aplicado, self.cfg["caminhos"]["pasta_relatorios"], aplicado=True
            )
        except Exception as e:
            self.root.after(0, self._aplicar_falhou, "Erro no relatório",
                            f"As alterações foram gravadas na Matriz, mas o relatório "
                            f"de auditoria falhou:\n{e}")
            return

        self.root.after(0, self._aplicar_concluido, n, total, str(caminho), aviso_base)

    def _aplicar_falhou(self, titulo: str, mensagem: str):
        self._parar_progresso()
        self.btn_analisar.config(state="normal")
        self.btn_aplicar.config(state="normal")  # permite tentar de novo
        self.lbl_status.config(text="Aplicação interrompida — veja a mensagem.")
        messagebox.showerror(titulo, mensagem)

    def _aplicar_concluido(self, n: int, total: int, caminho: str, aviso_base: str):
        self._parar_progresso()
        self.btn_analisar.config(state="normal")
        if aviso_base:
            messagebox.showwarning("Base de indicadores não atualizada", aviso_base)
        messagebox.showinfo("Concluído", f"{n} alteração(ões) aplicada(s).\nRelatório: {caminho}")
        self.btn_aplicar.config(state="disabled")
        self.lbl_status.config(text=f"{n} de {total} alterações aplicadas.")

    def executar(self):
        self.root.mainloop()
