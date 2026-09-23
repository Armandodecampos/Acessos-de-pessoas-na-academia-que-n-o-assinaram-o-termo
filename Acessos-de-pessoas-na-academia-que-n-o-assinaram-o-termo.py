import io
import os
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd


class AppCruzadorExcel(tk.Tk):
    """
    Aplicação Tkinter para cruzamento e tratamento automatizado e flexível
    de dados de acessos (eventos), cadastro de pessoas e filtragem por cartões e dispositivos.
    """

    def __init__(self):
        super().__init__()

        self.title("Cruzador de Eventos e Cartões de Acesso - Profissional")
        self.geometry("1180x820")
        self.minsize(980, 700)

        # Paleta de cores escura e profissional
        self.BG_DARK = "#1e1e2e"
        self.CARD_BG = "#2b2b3d"
        self.CARD_HEADER = "#36364c"
        self.ACCENT_BLUE = "#4a6fa5"
        self.ACCENT_GREEN = "#27ae60"
        self.ACCENT_RED = "#e74c3c"
        self.ACCENT_YELLOW = "#f39c12"
        self.TEXT_COLOR = "#f8f8f2"
        self.TEXT_MUTED = "#a0a0b0"

        self.configure(bg=self.BG_DARK)

        # DataFrames em memória
        self.df_guia1 = None  # Eventos de Acesso
        self.df_guia2 = None  # Cadastro de Pessoas
        self.df_guia3 = None  # Filtro / Lista de Cartões (Opcional)
        self.df_resultado = None  # Resultado Final Processado

        self._build_ui()

    def _build_ui(self):
        """Constrói o cabeçalho, botões de ação e abas principais."""
        header_frame = tk.Frame(self, bg=self.CARD_BG, pady=12, padx=20)
        header_frame.pack(fill=tk.X, side=tk.TOP)

        title_lbl = tk.Label(
            header_frame,
            text="Cruzador de Eventos, Cadastro e Cartões de Acesso",
            font=("Segoe UI", 16, "bold"),
            fg=self.TEXT_COLOR,
            bg=self.CARD_BG,
        )
        title_lbl.pack(anchor="w")

        sub_lbl = tk.Label(
            header_frame,
            text="Carregue os arquivos de cada guia e clique em EXECUTAR CRUZAMENTO.",
            font=("Segoe UI", 9),
            fg=self.TEXT_MUTED,
            bg=self.CARD_BG,
        )
        sub_lbl.pack(anchor="w")

        # Painel superior de controle
        control_frame = tk.Frame(self, bg=self.BG_DARK, padx=20, pady=8)
        control_frame.pack(fill=tk.X)

        self.btn_iniciar = tk.Button(
            control_frame,
            text="▶ EXECUTAR CRUZAMENTO",
            font=("Segoe UI", 11, "bold"),
            bg="#555566",
            fg="#888899",
            state=tk.DISABLED,
            padx=18,
            pady=8,
            bd=0,
            cursor="hand2",
            command=self._processar_dados,
        )
        self.btn_iniciar.pack(side=tk.LEFT)

        self.lbl_status = tk.Label(
            control_frame,
            text="Aguardando carregamento de arquivos: Guia 1 e Guia 2 são obrigatórias",
            font=("Segoe UI", 10, "bold"),
            fg=self.TEXT_MUTED,
            bg=self.BG_DARK,
        )
        self.lbl_status.pack(side=tk.LEFT, padx=15)

        # Configuração das Abas (Notebook)
        style = ttk.Style()
        style.theme_use("default")
        style.configure("TNotebook", background=self.BG_DARK, borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            padding=[14, 8],
            font=("Segoe UI", 9, "bold"),
            background=self.CARD_BG,
            foreground=self.TEXT_MUTED,
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", self.ACCENT_BLUE)],
            foreground=[("selected", "#ffffff")],
        )

        main_container = tk.Frame(self, bg=self.BG_DARK, padx=20)
        main_container.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.tab1 = tk.Frame(self.notebook, bg=self.CARD_BG)
        self.tab2 = tk.Frame(self.notebook, bg=self.CARD_BG)
        self.tab3 = tk.Frame(self.notebook, bg=self.CARD_BG)
        self.tab4 = tk.Frame(self.notebook, bg=self.CARD_BG)

        self.notebook.add(self.tab1, text="1. Eventos de Acesso (Guia 1)")
        self.notebook.add(self.tab2, text="2. Cadastro de Pessoas (Guia 2)")
        self.notebook.add(self.tab3, text="3. Filtro de Cartões (Guia 3)")
        self.notebook.add(self.tab4, text="4. Resultado Final")

        self._setup_tab_guia1()
        self._setup_tab_guia2()
        self._setup_tab_guia3()
        self._setup_tab_resultado()

    def _setup_tab_guia1(self):
        inst_frame = tk.LabelFrame(
            self.tab1,
            text=" Guia 1: Tabela de Eventos de Acesso (Obrigatória) ",
            bg=self.CARD_BG,
            fg=self.TEXT_COLOR,
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=8,
        )
        inst_frame.pack(fill=tk.X, padx=15, pady=10)

        lbl = tk.Label(
            inst_frame,
            text="Colunas identificadas automaticamente: ID do Evento | Horário | Nome do Dispositivo | ID Pessoal | Número do Cartão...\n"
            "📌 A aplicação buscará o 'Número do Cartão' correspondente no Cadastro (Guia 2) através do 'ID Pessoal'.",
            font=("Segoe UI", 8),
            fg=self.TEXT_MUTED,
            bg=self.CARD_BG,
            justify=tk.LEFT,
        )
        lbl.pack(anchor="w")

        btn_frame = tk.Frame(self.tab1, bg=self.CARD_BG)
        btn_frame.pack(fill=tk.X, padx=15, pady=5)

        self.lbl_guia1_status = tk.Label(
            btn_frame,
            text="✖ NENHUM ARQUIVO CARREGADO",
            font=("Segoe UI", 9, "bold"),
            fg=self.ACCENT_RED,
            bg=self.CARD_BG,
        )
        self.lbl_guia1_status.pack(side=tk.LEFT)

        btn_carregar = tk.Button(
            btn_frame,
            text="📁 Carregar Arquivo (Excel / CSV)",
            font=("Segoe UI", 9, "bold"),
            bg=self.ACCENT_BLUE,
            fg="#ffffff",
            padx=12,
            pady=4,
            command=lambda: self._carregar_arquivo(1),
        )
        btn_carregar.pack(side=tk.LEFT, padx=15)

        btn_limpar = tk.Button(
            btn_frame,
            text="🗑 Limpar Aba",
            font=("Segoe UI", 8),
            bg="#555566",
            fg="#ffffff",
            padx=8,
            command=lambda: self._limpar_aba(1),
        )
        btn_limpar.pack(side=tk.LEFT)

        self.tree1 = self._criar_treeview(self.tab1)

    def _setup_tab_guia2(self):
        inst_frame = tk.LabelFrame(
            self.tab2,
            text=" Guia 2: Cadastro de Pessoas / VLOOKUP (Obrigatória) ",
            bg=self.CARD_BG,
            fg=self.TEXT_COLOR,
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=8,
        )
        inst_frame.pack(fill=tk.X, padx=15, pady=10)

        lbl = tk.Label(
            inst_frame,
            text="Colunas identificadas automaticamente: ID Pessoal | Nome | Sobrenome | Número do Cartão...\n"
            "🔑 Tabela de Origem: Fornece o 'Número do Cartão' e nomes associados a cada 'ID Pessoal'.",
            font=("Segoe UI", 8),
            fg=self.TEXT_MUTED,
            bg=self.CARD_BG,
            justify=tk.LEFT,
        )
        lbl.pack(anchor="w")

        btn_frame = tk.Frame(self.tab2, bg=self.CARD_BG)
        btn_frame.pack(fill=tk.X, padx=15, pady=5)

        self.lbl_guia2_status = tk.Label(
            btn_frame,
            text="✖ NENHUM ARQUIVO CARREGADO",
            font=("Segoe UI", 9, "bold"),
            fg=self.ACCENT_RED,
            bg=self.CARD_BG,
        )
        self.lbl_guia2_status.pack(side=tk.LEFT)

        btn_carregar = tk.Button(
            btn_frame,
            text="📁 Carregar Arquivo (Excel / CSV)",
            font=("Segoe UI", 9, "bold"),
            bg=self.ACCENT_BLUE,
            fg="#ffffff",
            padx=12,
            pady=4,
            command=lambda: self._carregar_arquivo(2),
        )
        btn_carregar.pack(side=tk.LEFT, padx=15)

        btn_limpar = tk.Button(
            btn_frame,
            text="🗑 Limpar Aba",
            font=("Segoe UI", 8),
            bg="#555566",
            fg="#ffffff",
            padx=8,
            command=lambda: self._limpar_aba(2),
        )
        btn_limpar.pack(side=tk.LEFT)

        self.tree2 = self._criar_treeview(self.tab2)

    def _setup_tab_guia3(self):
        inst_frame = tk.LabelFrame(
            self.tab3,
            text=" Guia 3: Lista de Referência de Cartões (Opcional) ",
            bg=self.CARD_BG,
            fg=self.TEXT_COLOR,
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=8,
        )
        inst_frame.pack(fill=tk.X, padx=15, pady=10)

        lbl = tk.Label(
            inst_frame,
            text="Coluna esperada: ID / Número do Cartão.\n"
            "📌 Todos os cartões informados nesta guia serão AUTOMATICAMENTE EXCLUÍDOS/REMOVIDOS do resultado final.",
            font=("Segoe UI", 8),
            fg=self.TEXT_MUTED,
            bg=self.CARD_BG,
            justify=tk.LEFT,
        )
        lbl.pack(anchor="w")

        btn_frame = tk.Frame(self.tab3, bg=self.CARD_BG)
        btn_frame.pack(fill=tk.X, padx=15, pady=5)

        self.lbl_guia3_status = tk.Label(
            btn_frame,
            text="⚪ NENHUM ARQUIVO CARREGADO (OPCIONAL)",
            font=("Segoe UI", 9, "bold"),
            fg=self.TEXT_MUTED,
            bg=self.CARD_BG,
        )
        self.lbl_guia3_status.pack(side=tk.LEFT)

        btn_carregar = tk.Button(
            btn_frame,
            text="📁 Carregar Arquivo (Excel / CSV)",
            font=("Segoe UI", 9, "bold"),
            bg=self.ACCENT_BLUE,
            fg="#ffffff",
            padx=12,
            pady=4,
            command=lambda: self._carregar_arquivo(3),
        )
        btn_carregar.pack(side=tk.LEFT, padx=15)

        btn_limpar = tk.Button(
            btn_frame,
            text="🗑 Limpar Aba",
            font=("Segoe UI", 8),
            bg="#555566",
            fg="#ffffff",
            padx=8,
            command=lambda: self._limpar_aba(3),
        )
        btn_limpar.pack(side=tk.LEFT)

        self.tree3 = self._criar_treeview(self.tab3)

    def _setup_tab_resultado(self):
        top_frame = tk.Frame(self.tab4, bg=self.CARD_BG, padx=15, pady=10)
        top_frame.pack(fill=tk.X)

        self.lbl_res_summary = tk.Label(
            top_frame,
            text="O resultado do cruzamento aparecerá aqui após o processamento.",
            font=("Segoe UI", 10),
            fg=self.TEXT_MUTED,
            bg=self.CARD_BG,
        )
        self.lbl_res_summary.pack(side=tk.LEFT)

        self.btn_exportar = tk.Button(
            top_frame,
            text="💾 Exportar para Excel (.xlsx)",
            font=("Segoe UI", 9, "bold"),
            bg=self.ACCENT_GREEN,
            fg="#ffffff",
            state=tk.DISABLED,
            padx=12,
            pady=4,
            command=self._exportar_excel,
        )
        self.btn_exportar.pack(side=tk.RIGHT)

        self.tree_res = self._criar_treeview(self.tab4)

    def _criar_treeview(self, parent):
        """Cria um componente de tabela Treeview com barras de rolagem."""
        frame = tk.Frame(parent, bg=self.CARD_BG)
        frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(5, 15))

        tree = ttk.Treeview(frame, show="headings", selectmode="extended")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        tree.pack(fill=tk.BOTH, expand=True)

        return tree

    def _carregar_arquivo(self, num_guia):
        """Abre diálogo para seleção de arquivo (Excel ou CSV) e carrega os dados."""
        filepath = filedialog.askopenfilename(
            title=f"Selecionar arquivo para Guia {num_guia}",
            filetypes=[
                ("Arquivos de Tabela", "*.xlsx *.xls *.csv *.tsv *.txt"),
                ("Arquivos Excel", "*.xlsx *.xls"),
                ("Arquivos CSV/Texto", "*.csv *.tsv *.txt"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if not filepath:
            return

        try:
            df_raw = self._ler_arquivo_df(filepath)
            if df_raw is None or df_raw.empty:
                messagebox.showwarning(
                    "Aviso", "O arquivo selecionado está vazio ou não pôde ser lido."
                )
                return

            df_tratado = self._tratar_cabecalhos(df_raw, num_guia)
            nome_arquivo = os.path.basename(filepath)

            if num_guia == 1:
                self.df_guia1 = df_tratado
                self._atualizar_preview(
                    self.tree1, self.lbl_guia1_status, self.df_guia1, 1, nome_arquivo
                )
            elif num_guia == 2:
                self.df_guia2 = df_tratado
                self._atualizar_preview(
                    self.tree2, self.lbl_guia2_status, self.df_guia2, 2, nome_arquivo
                )
            elif num_guia == 3:
                self.df_guia3 = df_tratado
                self._atualizar_preview(
                    self.tree3, self.lbl_guia3_status, self.df_guia3, 3, nome_arquivo
                )

            self._validar_liberacao_botao()

        except Exception as e:
            messagebox.showerror(
                "Erro ao Carregar Arquivo",
                f"Não foi possível processar o arquivo selecionado:\n{str(e)}",
            )

    def _ler_arquivo_df(self, filepath):
        """Lê um arquivo Excel ou CSV/TSV e retorna um DataFrame com dtypes em string."""
        ext = os.path.splitext(filepath)[1].lower()
        if ext in [".xlsx", ".xls"]:
            df_raw = pd.read_excel(filepath, header=None, dtype=str)
        else:
            try:
                df_raw = pd.read_csv(
                    filepath,
                    sep=None,
                    dtype=str,
                    header=None,
                    keep_default_na=False,
                    engine="python",
                )
            except Exception:
                df_raw = pd.read_csv(
                    filepath,
                    sep="\t",
                    dtype=str,
                    header=None,
                    keep_default_na=False,
                    engine="python",
                )

        df_raw = df_raw.fillna("")

        for col in df_raw.columns:
            df_raw[col] = (
                df_raw[col]
                .astype(str)
                .str.strip()
                .str.replace(r'[\r\n"]', "", regex=True)
            )

        return df_raw

    def _tratar_cabecalhos(self, df_raw, num_guia):
        """Localiza a linha de cabeçalho através de palavras-chave nas primeiras 10 linhas."""
        header_idx = -1

        palavras_chave = {
            1: ["ID do Evento", "ID Pessoal", "Horário", "Nome do Dispositivo"],
            2: ["ID Pessoal", "Número do Cartão", "Nome", "Sobrenome"],
            3: ["ID", "Cartão", "Número do Cartão"],
        }

        alvos = palavras_chave.get(num_guia, [])

        for idx in range(min(10, len(df_raw))):
            linha_vals = [str(v).strip() for v in df_raw.iloc[idx].values]
            if any(term in linha_vals for term in alvos):
                header_idx = idx
                break

        if header_idx != -1:
            headers = [str(v).strip() for v in df_raw.iloc[header_idx].values]
            df = df_raw.iloc[header_idx + 1 :].copy().reset_index(drop=True)
            df.columns = headers
        else:
            df = df_raw.copy()
            df.columns = [f"Coluna_{i+1}" for i in range(df.shape[1])]

        # Converte todas as colunas para string por garantia
        for col in df.columns:
            df[col] = df[col].astype(str).str.strip()

        return df

    def _atualizar_preview(self, tree, label, df, num_guia, nome_arquivo=""):
        """Atualiza visualmente a tabela na interface gráfica prevenindo erros de tipo float."""
        tree.delete(*tree.get_children())
        cols = [str(c) for c in df.columns]
        tree["columns"] = cols

        for col in cols:
            tree.heading(col, text=col, anchor="w")
            serie_str = df[col].astype(str)
            max_len = (
                max(serie_str.map(lambda x: len(str(x))).max(), len(col))
                if len(df) > 0
                else len(col)
            )
            col_width = min(max(max_len * 9, 80), 250)
            tree.column(col, width=col_width, minwidth=60, anchor="w")

        for idx, row in df.head(300).iterrows():
            tree.insert("", "end", values=[str(v) for v in row])

        info_file = f" ({nome_arquivo})" if nome_arquivo else ""
        msg = f"✔ DADOS CARREGADOS{info_file} ({len(df)} registros)"
        color = self.ACCENT_GREEN if num_guia != 3 else self.ACCENT_BLUE
        label.config(text=msg, fg=color)

    def _limpar_aba(self, num_guia):
        """Reseta a aba selecionada."""
        if num_guia == 1:
            self.df_guia1 = None
            self.tree1.delete(*self.tree1.get_children())
            self.lbl_guia1_status.config(
                text="✖ NENHUM ARQUIVO CARREGADO", fg=self.ACCENT_RED
            )
        elif num_guia == 2:
            self.df_guia2 = None
            self.tree2.delete(*self.tree2.get_children())
            self.lbl_guia2_status.config(
                text="✖ NENHUM ARQUIVO CARREGADO", fg=self.ACCENT_RED
            )
        elif num_guia == 3:
            self.df_guia3 = None
            self.tree3.delete(*self.tree3.get_children())
            self.lbl_guia3_status.config(
                text="⚪ NENHUM ARQUIVO CARREGADO (OPCIONAL)", fg=self.TEXT_MUTED
            )

        self._validar_liberacao_botao()

    def _validar_liberacao_botao(self):
        """Habilita o botão de execução apenas quando Guia 1 e Guia 2 estiverem carregadas."""
        g1_ok = self.df_guia1 is not None and not self.df_guia1.empty
        g2_ok = self.df_guia2 is not None and not self.df_guia2.empty

        if g1_ok and g2_ok:
            self.btn_iniciar.config(
                state=tk.NORMAL, bg=self.ACCENT_GREEN, fg="#ffffff"
            )
            self.lbl_status.config(
                text="✔ Guia 1 e Guia 2 prontas! Clique em EXECUTAR CRUZAMENTO.",
                fg=self.ACCENT_GREEN,
            )
        else:
            pendentes = []
            if not g1_ok:
                pendentes.append("Guia 1")
            if not g2_ok:
                pendentes.append("Guia 2")
            self.btn_iniciar.config(state=tk.DISABLED, bg="#555566", fg="#888899")
            self.lbl_status.config(
                text=f"Aguardando carregamento de arquivos: {', '.join(pendentes)}",
                fg=self.TEXT_MUTED,
            )

    def _processar_dados(self):
        """
        Executa o cruzamento completo automatizado:
        1. VLOOKUP do Número do Cartão, Nome e Sobrenome da Guia 2 para a Guia 1 via 'ID Pessoal'.
        2. Padronização e filtro automático para 'Entrada Academia' e 'Saída Academia'.
        3. Exclusão automática de registros sem número de cartão.
        4. Exclusão automática dos cartões presentes na Guia 3 (exclui os cartões da Guia 3).
        """
        try:
            df1 = self.df_guia1.copy()
            df2 = self.df_guia2.copy()
            df3 = self.df_guia3.copy() if self.df_guia3 is not None else None

            # Localização flexível de colunas
            col_id_pessoal_g1 = self._encontrar_coluna(
                df1,
                ["ID Pessoal", "ID_Pessoal", "Id Pessoal", "ID"],
                idx_fallback=6,
            )
            col_cartao_g1 = self._encontrar_coluna(
                df1,
                ["Número do Cartão", "Numero do Cartao", "Cartão", "Cartao"],
                idx_fallback=9,
            )
            col_disp_g1 = self._encontrar_coluna(
                df1,
                ["Nome do Dispositivo", "Dispositivo", "Ponto do Evento"],
                idx_fallback=3,
            )
            col_nome_g1 = self._encontrar_coluna(df1, ["Nome"], idx_fallback=7)
            col_sobrenome_g1 = self._encontrar_coluna(
                df1, ["Sobrenome"], idx_fallback=8
            )

            col_id_pessoal_g2 = self._encontrar_coluna(
                df2,
                ["ID Pessoal", "ID_Pessoal", "Id Pessoal", "ID"],
                idx_fallback=0,
            )
            col_cartao_g2 = self._encontrar_coluna(
                df2,
                ["Número do Cartão", "Numero do Cartao", "Cartão", "Cartao"],
                idx_fallback=8,
            )
            col_nome_g2 = self._encontrar_coluna(df2, ["Nome"], idx_fallback=1)
            col_sobrenome_g2 = self._encontrar_coluna(
                df2, ["Sobrenome"], idx_fallback=2
            )

            def normalizar(val):
                if pd.isna(val) or val is None:
                    return ""
                s = str(val).strip()
                if s.endswith(".0"):
                    s = s[:-2]
                return s

            # Construção do mapa de busca da Guia 2
            mapa_cartao = {}
            mapa_nome = {}
            mapa_sobrenome = {}

            for _, row in df2.iterrows():
                id_p = normalizar(row[col_id_pessoal_g2])
                cartao = normalizar(row[col_cartao_g2])
                if id_p and id_p not in mapa_cartao:
                    mapa_cartao[id_p] = cartao
                    if col_nome_g2 in df2.columns:
                        mapa_nome[id_p] = str(row[col_nome_g2]).strip()
                    if col_sobrenome_g2 in df2.columns:
                        mapa_sobrenome[id_p] = str(row[col_sobrenome_g2]).strip()

            # Preenchimento de Número do Cartão, Nome e Sobrenome na Guia 1
            cartoes_preenchidos = 0
            for idx, row in df1.iterrows():
                id_p = normalizar(row[col_id_pessoal_g1])

                if id_p in mapa_cartao and mapa_cartao[id_p]:
                    df1.at[idx, col_cartao_g1] = mapa_cartao[id_p]
                    cartoes_preenchidos += 1

                if (
                    pd.isna(row[col_nome_g1]) or not str(row[col_nome_g1]).strip()
                ) and id_p in mapa_nome:
                    df1.at[idx, col_nome_g1] = mapa_nome[id_p]
                if (
                    pd.isna(row[col_sobrenome_g1])
                    or not str(row[col_sobrenome_g1]).strip()
                ) and id_p in mapa_sobrenome:
                    df1.at[idx, col_sobrenome_g1] = mapa_sobrenome[id_p]

            # 1. Padronização e Filtro dos Dispositivos para "Entrada Academia" / "Saída Academia"
            col_ponto = self._encontrar_coluna(
                df1, ["Ponto do Evento", "Ponto"], idx_fallback=4
            )
            col_leitor = self._encontrar_coluna(
                df1, ["Nome do Leitor", "Leitor"], idx_fallback=12
            )

            def padronizar_disp(row):
                d_val = str(row.get(col_disp_g1, "")).strip()
                p_val = str(row.get(col_ponto, "")).strip() if col_ponto in row else ""
                l_val = str(row.get(col_leitor, "")) if col_leitor in row else ""

                txt = f"{d_val} {p_val} {l_val}".lower()
                if "saída" in txt or "saida" in txt or " -- s" in txt or "--s" in txt:
                    return "Saída Academia"
                return "Entrada Academia"

            df1[col_disp_g1] = df1.apply(padronizar_disp, axis=1)

            # 2. Remoção OBRIGATÓRIA de eventos sem Número do Cartão
            df_filtrado = df1[
                df1[col_cartao_g1].apply(normalizar) != ""
            ].copy()

            # 3. Exclusão OBRIGATÓRIA dos cartões presentes na Guia 3 (exclui os cartões da Guia 3)
            if df3 is not None and not df3.empty:
                col_id_g3 = self._encontrar_coluna(
                    df3,
                    ["ID", "Cartão", "Número do Cartão", "Numero do Cartao"],
                    idx_fallback=0,
                )
                set_cartoes_g3 = set(
                    df3[col_id_g3]
                    .apply(normalizar)
                    .replace("", None)
                    .dropna()
                )

                df_filtrado["_TEMP_KEY_CARTAO"] = df_filtrado[
                    col_cartao_g1
                ].apply(normalizar)
                df_final = df_filtrado[
                    ~df_filtrado["_TEMP_KEY_CARTAO"].isin(set_cartoes_g3)
                ].copy()
                df_final = df_final.drop(columns=["_TEMP_KEY_CARTAO"])
            else:
                df_final = df_filtrado.copy()

            self.df_resultado = df_final

            # Exibição no Treeview do Resultado Final
            self._atualizar_preview(
                self.tree_res,
                self.lbl_res_summary,
                self.df_resultado,
                num_guia=4,
            )

            total_original = len(self.df_guia1)
            total_final = len(self.df_resultado)

            self.lbl_res_summary.config(
                text=f"✔ Concluído: {total_final} registros mantidos (de {total_original} originais na Guia 1). "
                f"Cartões vinculados do cadastro: {cartoes_preenchidos}.",
                fg=self.ACCENT_GREEN,
            )
            self.btn_exportar.config(state=tk.NORMAL)

            self.notebook.select(self.tab4)
            messagebox.showinfo(
                "Processamento Concluído",
                f"O cruzamento foi executado com sucesso!\n\n"
                f"📊 Resumo do Processamento:\n"
                f"• Registros Originais (Guia 1): {total_original}\n"
                f"• Cartões Vinculados do Cadastro: {cartoes_preenchidos}\n"
                f"• Registros Finais Mantidos: {total_final}\n\n"
                f"• Regras aplicadas:\n"
                f"  - Cartões da Guia 3 removidos/excluídos\n"
                f"  - Eventos sem cartão excluídos\n"
                f"  - Dispositivos padronizados para Entrada/Saída Academia",
            )

        except Exception as e:
            messagebox.showerror(
                "Erro de Processamento",
                f"Ocorreu um erro no cruzamento dos dados:\n{str(e)}",
            )

    def _encontrar_coluna(self, df, nomes_possiveis, idx_fallback):
        """Busca a coluna por nome exato/parcial ou usa o índice como fallback seguro."""
        for nome in nomes_possiveis:
            if nome in df.columns:
                return nome

        if idx_fallback < df.shape[1]:
            return df.columns[idx_fallback]

        return df.columns[0]

    def _exportar_excel(self):
        """Exporta o resultado processado para um arquivo .xlsx."""
        if self.df_resultado is None or self.df_resultado.empty:
            messagebox.showwarning("Aviso", "Não há resultados para exportar.")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            title="Salvar Resultado do Cruzamento",
        )

        if filepath:
            try:
                self.df_resultado.to_excel(filepath, index=False)
                messagebox.showinfo(
                    "Sucesso", f"Arquivo salvo com sucesso em:\n{filepath}"
                )
            except Exception as e:
                messagebox.showerror(
                    "Erro ao Salvar", f"Não foi possível salvar o arquivo:\n{e}"
                )


if __name__ == "__main__":
    app = AppCruzadorExcel()
    app.mainloop()
