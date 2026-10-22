from __future__ import annotations

import os
import csv
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd


CAMINHO_CSV_PADRAO = r"H:\DPR-DSU\PROGRAMAS\OF_Search\bd_apontamentos_modelo_CSV_2026.csv"

# =========================================================
# CONFIGURAÇÕES
# =========================================================

CSV_PADRAO = (
    r"H:\DPR-DSU\PROGRAMAS\OF_Search"
    r"\bd_apontamentos_modelo_CSV_2026.csv"
)

PASTA_RELATORIOS = (
    r"H:\DPR-DSU\PROGRAMAS\OF_Search\RELATORIOS"
)

ARQUIVO_FLUXO = "H:\DPR-DSU\PROGRAMAS\OF_Search\modelo_recurso.txt"


# =========================================================
# ESTRUTURA DO FLUXO
# =========================================================

@dataclass
class Etapa:
    """
    Representa uma etapa do processo.

    Exemplo:

        2101;2102;2103;2104;2105

    vira:

        recursos = ("2101", "2102", "2103", "2104", "2105")
        opcional = False

    Já:

        {101;102;103}

    vira:

        recursos = ("101", "102", "103")
        opcional = True
    """

    recursos: tuple[str, ...]
    opcional: bool


# =========================================================
# GERENCIADOR DOS DADOS
# =========================================================

class GerenciadorDados:

    COLUNAS_OBRIGATORIAS = {
        "OF+PC",
        "OF",
        "PC",
        "RECURSO",
        "DESCRICAO",
        "DATA APONTAMENTO.1",
        "OPERADOR",
        "MODELO",
    }

    def __init__(self):

        self.df = pd.DataFrame()

        # Índices para pesquisa rápida
        self.indice_ofpc = {}
        self.indice_of = {}

        # Recurso -> descrição
        self.descricoes = {}

        self.caminho_atual = None

    # -----------------------------------------------------
    # Detecta encoding
    # -----------------------------------------------------

    def detectar_encoding(self, caminho):

        encodings = [
            "utf-8-sig",
            "utf-8",
            "cp1252",
            "latin1",
        ]

        for encoding in encodings:

            try:

                with open(
                    caminho,
                    "r",
                    encoding=encoding,
                ) as arquivo:

                    arquivo.read(10000)

                return encoding

            except UnicodeDecodeError:

                continue

        return "latin1"

    # -----------------------------------------------------
    # Detecta separador
    # -----------------------------------------------------

    def detectar_separador(self, caminho, encoding):

        with open(
            caminho,
            "r",
            encoding=encoding,
            newline="",
        ) as arquivo:

            amostra = arquivo.read(20000)

        try:

            dialect = csv.Sniffer().sniff(
                amostra,
                delimiters=",;\t|",
            )

            return dialect.delimiter

        except csv.Error:

            return ","

    # -----------------------------------------------------
    # Carregar CSV
    # -----------------------------------------------------

    def carregar(self, caminho):

        caminho = Path(caminho)

        if not caminho.exists():

            raise FileNotFoundError(
                f"CSV não encontrado:\n\n{caminho}"
            )

        encoding = self.detectar_encoding(caminho)

        separador = self.detectar_separador(
            caminho,
            encoding,
        )

        df = pd.read_csv(
            caminho,
            sep=separador,
            encoding=encoding,
            dtype=str,
            keep_default_na=False,
            low_memory=False,
        )

        # Limpa nomes das colunas
        df.columns = [
            str(coluna).strip()
            for coluna in df.columns
        ]

        faltantes = (
            self.COLUNAS_OBRIGATORIAS
            - set(df.columns)
        )

        if faltantes:

            raise ValueError(
                "Colunas obrigatórias não encontradas:\n\n"
                + "\n".join(sorted(faltantes))
            )

        # -------------------------------------------------
        # Limpeza dos campos
        # -------------------------------------------------

        colunas_texto = [
            "OF+PC",
            "OF",
            "PC",
            "RECURSO",
            "DESCRICAO",
            "OPERADOR",
            "MODELO",
        ]

        for coluna in colunas_texto:

            df[coluna] = (
                df[coluna]
                .fillna("")
                .astype(str)
                .str.strip()
            )

        # Remove linhas sem OF+PC
        df = df[
            df["OF+PC"] != ""
        ].copy()

        # -------------------------------------------------
        # Data
        # -------------------------------------------------

        df["_DATA"] = pd.to_datetime(
            df["DATA APONTAMENTO.1"],
            dayfirst=True,
            errors="coerce",
        )

        # Ordem original do CSV
        df["_ORDEM"] = range(len(df))

        self.df = df

        self.caminho_atual = caminho

        # -------------------------------------------------
        # CRIA ÍNDICE OF+PC
        # -------------------------------------------------

        self.indice_ofpc = {
            chave: grupo.copy()

            for chave, grupo
            in df.groupby(
                "OF+PC",
                sort=False,
            )
        }

        # -------------------------------------------------
        # CRIA ÍNDICE OF
        # -------------------------------------------------

        grupos_of = (
            df.groupby(
                "OF",
                sort=False,
            )["OF+PC"]
            .unique()
        )

        self.indice_of = {
            chave: list(valores)

            for chave, valores
            in grupos_of.items()
        }

        # -------------------------------------------------
        # DESCRIÇÕES DOS RECURSOS
        # -------------------------------------------------

        self.descricoes = {}

        for linha in df[
            [
                "RECURSO",
                "DESCRICAO",
            ]
        ].itertuples(index=False):

            recurso = linha.RECURSO
            descricao = linha.DESCRICAO

            if recurso and recurso not in self.descricoes:

                self.descricoes[recurso] = (
                    descricao
                    or "SEM DESCRICAO"
                )

        return len(df)


# =========================================================
# GERENCIADOR DO FLUXO
# =========================================================

class GerenciadorFluxo:

    def __init__(self):

        self.fluxos = {}

    # -----------------------------------------------------
    # Verifica se é linha de modelo
    # -----------------------------------------------------

    def eh_modelo(self, linha):

        return bool(
            re.match(
                r"^[A-Za-zÀ-ÿ]",
                linha,
            )
        )

    # -----------------------------------------------------
    # Ordena recursos numericamente
    # -----------------------------------------------------

    def ordenar_recursos(self, recursos):

        def chave(valor):

            try:

                return (
                    0,
                    int(valor),
                )

            except ValueError:

                return (
                    1,
                    valor,
                )

        return tuple(
            sorted(
                recursos,
                key=chave,
            )
        )

    # -----------------------------------------------------
    # Carrega modelo_recurso.txt
    # -----------------------------------------------------

    def carregar(self, caminho):

        caminho = Path(caminho)

        if not caminho.exists():

            raise FileNotFoundError(
                f"Arquivo modelo_recurso.txt não encontrado:\n\n"
                f"{caminho}"
            )

        self.fluxos.clear()

        modelos_atuais = []

        with open(
            caminho,
            "r",
            encoding="utf-8-sig",
        ) as arquivo:

            for linha_original in arquivo:

                linha = linha_original.strip()

                if not linha:

                    continue

                # -------------------------------------------------
                # LINHA DE MODELOS
                # -------------------------------------------------

                if self.eh_modelo(linha):

                    modelos_atuais = [
                        modelo.strip()

                        for modelo in linha.split(";")

                        if modelo.strip()
                    ]

                    for modelo in modelos_atuais:

                        self.fluxos.setdefault(
                            modelo,
                            [],
                        )

                    continue

                # -------------------------------------------------
                # ETAPA OPCIONAL
                # -------------------------------------------------

                opcional = False

                if (
                    linha.startswith("{")
                    and linha.endswith("}")
                ):

                    opcional = True

                    linha = linha[1:-1].strip()

                # -------------------------------------------------
                # RECURSOS
                # -------------------------------------------------

                recursos = [
                    recurso.strip()

                    for recurso in linha.split(";")

                    if recurso.strip()
                ]

                if not recursos:

                    continue

                recursos = self.ordenar_recursos(
                    recursos
                )

                etapa = Etapa(
                    recursos=recursos,
                    opcional=opcional,
                )

                for modelo in modelos_atuais:

                    self.fluxos[modelo].append(
                        etapa
                    )


# =========================================================
# ANALISADOR
# =========================================================

class Analisador:

    def __init__(
        self,
        dados,
        fluxo,
    ):

        self.dados = dados
        self.fluxo = fluxo

    # -----------------------------------------------------
    # Primeiro apontamento por recurso
    # -----------------------------------------------------

    def separar_apontamentos(
        self,
        grupo,
    ):

        ordenado = grupo.sort_values(
            by=[
                "_DATA",
                "_ORDEM",
            ],
            na_position="last",
        )

        primeiros = {}
        repetidos = {}

        for _, linha in ordenado.iterrows():

            recurso = linha.RECURSO

            if not recurso:

                continue

            if recurso not in primeiros:

                primeiros[recurso] = linha

            else:

                repetidos.setdefault(
                    recurso,
                    [],
                )

                repetidos[recurso].append(
                    linha
                )

        return (
            primeiros,
            repetidos,
        )

    # -----------------------------------------------------
    # Cria linha do relatório
    # -----------------------------------------------------

    def criar_resultado(
        self,
        ofpc,
        modelo,
        recurso,
        descricao,
        data,
        operador,
        status,
        observacao,
    ):

        return {
            "OF+PC": ofpc,
            "MODELO": modelo,
            "RECURSO": recurso,
            "DESCRICAO": descricao,
            "DATA": data,
            "OPERADOR": operador,
            "STATUS": status,
            "OBSERVACAO": observacao,
        }

    # -----------------------------------------------------
    # ANÁLISE PRINCIPAL
    # -----------------------------------------------------

    def analisar(self, ofpc):

        grupo = self.dados.indice_ofpc.get(
            ofpc
        )

        if grupo is None or grupo.empty:

            return []

        modelo = (
            grupo.iloc[0]["MODELO"]
            .strip()
        )

        fluxo = self.fluxo.fluxos.get(
            modelo
        )

        if fluxo is None:

            raise ValueError(
                f"O fluxo do modelo "
                f"'{modelo}' não foi encontrado "
                f"no modelo_recurso.txt."
            )

        primeiros, repetidos = (
            self.separar_apontamentos(
                grupo
            )
        )

        # -------------------------------------------------
        # Último apontamento
        # -------------------------------------------------

        ordenado = grupo.sort_values(
            by=[
                "_DATA",
                "_ORDEM",
            ],
            na_position="last",
        )

        ultimo = ordenado.iloc[-1]

        ultimo_recurso = (
            str(
                ultimo["RECURSO"]
            ).strip()
        )

        # -------------------------------------------------
        # Descobre até qual etapa a peça chegou
        # -------------------------------------------------

        maior_etapa = -1

        for numero_etapa, etapa in enumerate(
            fluxo
        ):

            encontrou = any(
                recurso in primeiros

                for recurso
                in etapa.recursos
            )

            if encontrou:

                maior_etapa = numero_etapa

        resultados = []

        # -------------------------------------------------
        # Percorre fluxo
        # -------------------------------------------------

        for numero_etapa, etapa in enumerate(
            fluxo
        ):

            # Não precisamos analisar etapas futuras
            if numero_etapa > maior_etapa:

                break

            # =================================================
            # ETAPA OPCIONAL
            # =================================================

            if etapa.opcional:

                for recurso in etapa.recursos:

                    if recurso not in primeiros:

                        continue

                    linha = primeiros[recurso]

                    observacoes = []

                    if recurso in repetidos:

                        observacoes.append(
                            "REPETIDO"
                        )

                    if recurso == ultimo_recurso:

                        observacoes.append(
                            "ULTIMO APONTAMENTO"
                        )

                    status = (
                        "ERRO"
                        if recurso in repetidos
                        else "OK"
                    )

                    resultados.append(
                        self.criar_resultado(
                            ofpc,
                            modelo,
                            recurso,
                            linha.DESCRICAO
                            or self.dados.descricoes.get(
                                recurso,
                                "SEM DESCRICAO",
                            ),
                            linha["DATA APONTAMENTO.1"],
                            linha.OPERADOR,
                            status,
                            " | ".join(
                                observacoes
                            ),
                        )
                    )

                continue

            # =================================================
            # ETAPA OBRIGATÓRIA
            # =================================================

            presentes = [
                recurso

                for recurso
                in etapa.recursos

                if recurso in primeiros
            ]

            # -------------------------------------------------
            # Nenhum recurso da etapa foi apontado
            # -------------------------------------------------

            if not presentes:

                descricoes = []

                for recurso in etapa.recursos:

                    descricoes.append(
                        self.dados.descricoes.get(
                            recurso,
                            "SEM DESCRICAO",
                        )
                    )

                resultados.append(
                    self.criar_resultado(
                        ofpc,
                        modelo,
                        " / ".join(
                            etapa.recursos
                        ),
                        " / ".join(
                            descricoes
                        ),
                        "SEM DATA",
                        "SEM OPERADOR",
                        "ERRO",
                        "NAO APONTADO",
                    )
                )

                continue

            # -------------------------------------------------
            # Recursos presentes
            # -------------------------------------------------

            for recurso in presentes:

                linha = primeiros[recurso]

                observacoes = []

                # Repetição
                if recurso in repetidos:

                    observacoes.append(
                        "REPETIDO"
                    )

                # Último
                if recurso == ultimo_recurso:

                    observacoes.append(
                        "ULTIMO APONTAMENTO"
                    )

                # -------------------------------------------------
                # Verifica ordem cronológica
                # -------------------------------------------------

                data_atual = linha._DATA

                erros_fluxo = []

                if pd.notna(data_atual):

                    for etapa_anterior in fluxo[
                        :numero_etapa
                    ]:

                        # Etapas opcionais não bloqueiam
                        if etapa_anterior.opcional:

                            continue

                        for recurso_anterior in (
                            etapa_anterior.recursos
                        ):

                            if (
                                recurso_anterior
                                not in primeiros
                            ):

                                continue

                            linha_anterior = (
                                primeiros[
                                    recurso_anterior
                                ]
                            )

                            data_anterior = (
                                linha_anterior._DATA
                            )

                            if pd.isna(
                                data_anterior
                            ):

                                continue

                            if (
                                data_anterior
                                > data_atual
                            ):

                                erros_fluxo.append(
                                    recurso_anterior
                                )

                if erros_fluxo:

                    observacoes.append(
                        "FLUXO INCORRETO "
                        + ", ".join(
                            erros_fluxo
                        )
                    )

                status = "OK"

                if (
                    recurso in repetidos
                    or erros_fluxo
                ):

                    status = "ERRO"

                resultados.append(
                    self.criar_resultado(
                        ofpc,
                        modelo,
                        recurso,
                        linha.DESCRICAO
                        or self.dados.descricoes.get(
                            recurso,
                            "SEM DESCRICAO",
                        ),
                        linha["DATA APONTAMENTO.1"],
                        linha.OPERADOR,
                        status,
                        " | ".join(
                            observacoes
                        ),
                    )
                )

        return resultados


# =========================================================
# INTERFACE
# =========================================================

class Aplicacao(tk.Tk):

    def __init__(self):

        super().__init__()

        self.title(
            "OF_SEARCH - Consulta de Apontamentos"
        )

        self.geometry(
            "1500x850"
        )

        self.minsize(
            1100,
            650,
        )

        # ---------------------------------------------
        # Objetos principais
        # ---------------------------------------------

        self.dados = GerenciadorDados()

        self.fluxo = GerenciadorFluxo()

        self.analisador = Analisador(
            self.dados,
            self.fluxo,
        )

        # ---------------------------------------------
        # Estado
        # ---------------------------------------------

        self.resultados_relatorio = []

        self.caminho_csv = tk.StringVar(
            value=CSV_PADRAO
        )

        self.status = tk.StringVar(
            value="Aguardando carregamento dos dados."
        )

        self.pesquisa = tk.StringVar()

        self.configurar_estilo()

        self.criar_interface()

        # Tenta carregar automaticamente
        self.after(
            200,
            self.carregamento_automatico,
        )

    # =================================================
    # ESTILO
    # =================================================

    def configurar_estilo(self):

        estilo = ttk.Style(self)

        try:

            estilo.theme_use(
                "vista"
            )

        except tk.TclError:

            pass

        estilo.configure(
            "Titulo.TLabel",
            font=(
                "Segoe UI",
                20,
                "bold",
            ),
        )

        estilo.configure(
            "Subtitulo.TLabel",
            font=(
                "Segoe UI",
                10,
            ),
        )

        estilo.configure(
            "Treeview",
            rowheight=28,
            font=(
                "Segoe UI",
                9,
            ),
        )

        estilo.configure(
            "Treeview.Heading",
            font=(
                "Segoe UI",
                9,
                "bold",
            ),
        )

    # =================================================
    # INTERFACE
    # =================================================

    def criar_interface(self):

        # -------------------------------------------------
        # CABEÇALHO
        # -------------------------------------------------

        cabecalho = ttk.Frame(
            self,
            padding=(
                18,
                15,
                18,
                8,
            ),
        )

        cabecalho.pack(
            fill="x"
        )

        ttk.Label(
            cabecalho,
            text="OF_SEARCH",
            style="Titulo.TLabel",
        ).pack(
            anchor="w"
        )

        ttk.Label(
            cabecalho,
            text=(
                "Consulta e análise de "
                "apontamentos por OF / OF+PC"
            ),
            style="Subtitulo.TLabel",
        ).pack(
            anchor="w",
            pady=(
                2,
                10,
            ),
        )

        # -------------------------------------------------
        # BOTÕES
        # -------------------------------------------------

        controles = ttk.Frame(
            cabecalho
        )

        controles.pack(
            fill="x"
        )

        ttk.Button(
            controles,
            text="⟳ Atualizar dados",
            command=self.selecionar_csv,
        ).pack(
            side="left"
        )

        ttk.Button(
            controles,
            text="Exportar relatório",
            command=self.exportar_relatorio,
        ).pack(
            side="left",
            padx=8,
        )

        ttk.Label(
            controles,
            textvariable=self.caminho_csv,
        ).pack(
            side="left",
            fill="x",
            expand=True,
            padx=10,
        )

        # -------------------------------------------------
        # PESQUISA
        # -------------------------------------------------

        frame_pesquisa = ttk.LabelFrame(
            self,
            text="Pesquisa",
            padding=12,
        )

        frame_pesquisa.pack(
            fill="x",
            padx=18,
            pady=(
                5,
                10,
            ),
        )

        ttk.Label(
            frame_pesquisa,
            text=(
                "OF ou OF+PC "
                "(várias consultas separadas por ;):"
            ),
        ).pack(
            side="left"
        )

        self.caixa_pesquisa = ttk.Entry(
            frame_pesquisa,
            textvariable=self.pesquisa,
            font=(
                "Segoe UI",
                11,
            ),
        )

        self.caixa_pesquisa.pack(
            side="left",
            fill="x",
            expand=True,
            padx=10,
        )

        ttk.Button(
            frame_pesquisa,
            text="Pesquisar",
            command=self.pesquisar,
        ).pack(
            side="left"
        )

        self.caixa_pesquisa.bind(
            "<Return>",
            lambda evento: self.pesquisar(),
        )

        # -------------------------------------------------
        # STATUS
        # -------------------------------------------------

        ttk.Label(
            self,
            textvariable=self.status,
        ).pack(
            anchor="w",
            padx=18,
            pady=(
                0,
                5,
            ),
        )

        # -------------------------------------------------
        # TABELA
        # -------------------------------------------------

        frame_tabela = ttk.Frame(
            self,
            padding=(
                18,
                0,
                18,
                18,
            ),
        )

        frame_tabela.pack(
            fill="both",
            expand=True,
        )

        colunas = (
            "OF+PC",
            "MODELO",
            "RECURSO",
            "DESCRICAO",
            "DATA",
            "OPERADOR",
            "STATUS",
            "OBSERVACAO",
        )

        self.tabela = ttk.Treeview(
            frame_tabela,
            columns=colunas,
            show="headings",
        )

        larguras = {
            "OF+PC": 110,
            "MODELO": 110,
            "RECURSO": 130,
            "DESCRICAO": 320,
            "DATA": 150,
            "OPERADOR": 120,
            "STATUS": 90,
            "OBSERVACAO": 500,
        }

        for coluna in colunas:

            self.tabela.heading(
                coluna,
                text=coluna,
            )

            self.tabela.column(
                coluna,
                width=larguras[coluna],
                anchor="w",
            )

        # Scroll vertical
        scroll_vertical = ttk.Scrollbar(
            frame_tabela,
            orient="vertical",
            command=self.tabela.yview,
        )

        # Scroll horizontal
        scroll_horizontal = ttk.Scrollbar(
            frame_tabela,
            orient="horizontal",
            command=self.tabela.xview,
        )

        self.tabela.configure(
            yscrollcommand=scroll_vertical.set,
            xscrollcommand=scroll_horizontal.set,
        )

        self.tabela.grid(
            row=0,
            column=0,
            sticky="nsew",
        )

        scroll_vertical.grid(
            row=0,
            column=1,
            sticky="ns",
        )

        scroll_horizontal.grid(
            row=1,
            column=0,
            sticky="ew",
        )

        frame_tabela.rowconfigure(
            0,
            weight=1,
        )

        frame_tabela.columnconfigure(
            0,
            weight=1,
        )

        # -------------------------------------------------
        # CORES
        # -------------------------------------------------

        self.tabela.tag_configure(
            "erro",
            foreground="#C00000",
        )

        self.tabela.tag_configure(
            "ultimo",
            foreground="#008000",
        )

        self.tabela.tag_configure(
            "normal",
            foreground="#202020",
        )

        self.tabela.tag_configure(
            "separador",
            background="#D9D9D9",
        )

    # =================================================
    # CARREGAMENTO AUTOMÁTICO
    # =================================================

    def carregamento_automatico(self):

        caminho = Path(
            CSV_PADRAO
        )

        if caminho.exists():

            try:

                self.carregar_csv(
                    caminho
                )

            except Exception as erro:

                self.status.set(
                    f"Erro ao carregar: {erro}"
                )

        else:

            self.status.set(
                "CSV padrão não encontrado. "
                "Clique em 'Atualizar dados'."
            )

        self.caixa_pesquisa.focus_set()

    # =================================================
    # SELECIONAR CSV
    # =================================================

    def selecionar_csv(self):

        caminho = filedialog.askopenfilename(
            title=(
                "Selecione o arquivo "
                "de apontamentos"
            ),
            initialdir=str(
                Path(
                    CSV_PADRAO
                ).parent
            ),
            initialfile=Path(
                CSV_PADRAO
            ).name,
            filetypes=[
                (
                    "Arquivos CSV",
                    "*.csv",
                ),
                (
                    "Todos os arquivos",
                    "*.*",
                ),
            ],
        )

        if caminho:

            self.carregar_csv(
                Path(caminho)
            )

    # =================================================
    # CARREGAR CSV
    # =================================================

    def carregar_csv(
        self,
        caminho,
    ):

        self.status.set(
            "Carregando dados... aguarde."
        )

        self.update_idletasks()

        try:

            quantidade = (
                self.dados.carregar(
                    caminho
                )
            )

            # modelo_recurso fica junto do programa
            pasta_programa = (
                Path(__file__)
                .resolve()
                .parent
            )

            caminho_fluxo = (
                pasta_programa
                / ARQUIVO_FLUXO
            )

            self.fluxo.carregar(
                caminho_fluxo
            )

            self.analisador = Analisador(
                self.dados,
                self.fluxo,
            )

            self.caminho_csv.set(
                str(caminho)
            )

            self.status.set(
                f"Dados carregados: "
                f"{quantidade:,} apontamentos | "
                f"OF+PC: "
                f"{len(self.dados.indice_ofpc):,} | "
                f"OF: "
                f"{len(self.dados.indice_of):,}"
            )

        except Exception as erro:

            messagebox.showerror(
                "Erro ao carregar dados",
                str(erro),
            )

            self.status.set(
                "Falha ao carregar os dados."
            )

    # =================================================
    # PESQUISAR
    # =================================================

    def pesquisar(self):

        if self.dados.df.empty:

            messagebox.showwarning(
                "Dados não carregados",
                "Carregue o CSV antes de pesquisar.",
            )

            return

        texto = (
            self.pesquisa
            .get()
            .strip()
        )

        if not texto:

            return

        # Limpa tabela
        for item in self.tabela.get_children():

            self.tabela.delete(
                item
            )

        self.resultados_relatorio = []

        consultas = [
            consulta.strip()

            for consulta
            in texto.split(";")

            if consulta.strip()
        ]

        pecas_encontradas = 0
        pecas_nao_encontradas = 0
        quantidade_erros = 0

        for consulta in consultas:

            # ---------------------------------------------
            # OF+PC
            # ---------------------------------------------

            if "-" in consulta:

                ofpcs = [
                    consulta
                ]

            # ---------------------------------------------
            # SOMENTE OF
            # ---------------------------------------------

            else:

                ofpcs = (
                    self.dados
                    .indice_of
                    .get(
                        consulta,
                        [],
                    )
                )

            if not ofpcs:

                pecas_nao_encontradas += 1

                continue

            # ---------------------------------------------
            # Analisa cada peça
            # ---------------------------------------------

            for ofpc in ofpcs:

                # Separador visual entre peças
                if pecas_encontradas > 0:
                    self.tabela.insert(
                        "",
                        "end",
                        values=(
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                        ),
                        tags=("separador",),
                    )

                try:

                    resultados = (
                        self.analisador.analisar(
                            ofpc
                        )
                    )

                except Exception as erro:

                    messagebox.showerror(
                        "Erro na análise",
                        f"{ofpc}\n\n{erro}",
                    )

                    continue

                if not resultados:

                    pecas_nao_encontradas += 1

                    continue

                pecas_encontradas += 1

                for resultado in resultados:

                    self.resultados_relatorio.append(
                        resultado
                    )

                    status = (
                        resultado["STATUS"]
                    )

                    observacao = (
                        resultado["OBSERVACAO"]
                    )

                    if status == "ERRO":

                        tag = "erro"

                        quantidade_erros += 1

                    elif (
                        "ULTIMO APONTAMENTO"
                        in observacao
                    ):

                        tag = "ultimo"

                    else:

                        tag = "normal"

                    self.tabela.insert(
                        "",
                        "end",
                        values=(
                            resultado["OF+PC"],
                            resultado["MODELO"],
                            resultado["RECURSO"],
                            resultado["DESCRICAO"],
                            resultado["DATA"],
                            resultado["OPERADOR"],
                            resultado["STATUS"],
                            resultado["OBSERVACAO"],
                        ),
                        tags=(
                            tag,
                        ),
                    )

        self.status.set(
            f"Pesquisa concluída | "
            f"peças encontradas: "
            f"{pecas_encontradas:,} | "
            f"não encontradas: "
            f"{pecas_nao_encontradas:,} | "
            f"erros: "
            f"{quantidade_erros:,} | "
            f"linhas: "
            f"{len(self.resultados_relatorio):,}"
        )

    # =================================================
    # EXPORTAR RELATÓRIO
    # =================================================

    def exportar_relatorio(self):

        if not self.resultados_relatorio:

            messagebox.showinfo(
                "Relatório",
                "Faça uma pesquisa antes de exportar.",
            )

            return

        pasta = Path(
            PASTA_RELATORIOS
        )

        try:

            pasta.mkdir(
                parents=True,
                exist_ok=True,
            )

        except OSError:

            pasta_selecionada = (
                filedialog.askdirectory(
                    title=(
                        "Escolha a pasta "
                        "para salvar o relatório"
                    )
                )
            )

            if not pasta_selecionada:

                return

            pasta = Path(
                pasta_selecionada
            )

        data_hora = (
            datetime.now()
            .strftime(
                "%Y%m%d_%H%M%S"
            )
        )

        arquivo_saida = (
            pasta
            / f"resultado_fluxo_{data_hora}.csv"
        )

        try:

            pd.DataFrame(
                self.resultados_relatorio
            ).to_csv(
                arquivo_saida,
                index=False,
                encoding="utf-8-sig",
                sep=";",
            )

            messagebox.showinfo(
                "Relatório exportado",
                "Arquivo gerado com sucesso:\n\n"
                f"{arquivo_saida}",
            )

            self.status.set(
                f"Relatório exportado: "
                f"{arquivo_saida}"
            )

        except Exception as erro:

            messagebox.showerror(
                "Erro ao exportar",
                str(erro),
            )


# =========================================================
# INÍCIO
# =========================================================

def main():

    app = Aplicacao()

    app.mainloop()


if __name__ == "__main__":

    main()

