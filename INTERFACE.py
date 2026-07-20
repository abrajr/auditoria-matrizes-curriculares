"""
INTERFACE GRÁFICA — AUDITORIA REGULATÓRIA DE MATRIZES CURRICULARES
Front-end Tkinter para o motor de auditoria (PLANEJADOR.py).

Executar:  python INTERFACE.py
Empacotar: python build_exe.py   (gera dist/Auditoria_Matrizes.exe)
"""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import PLANEJADOR

# ---------------------------------------------------------------------------
# Paleta e constantes visuais
# ---------------------------------------------------------------------------
COR_NAVY = "#1F3864"
COR_AZUL = "#2E5496"
COR_FUNDO = "#F5F7FA"
COR_CARD = "#FFFFFF"
COR_ALERTA_BG = "#FFF3CD"
COR_ALERTA_BORDA = "#E0A800"
COR_ALERTA_TXT = "#7A5800"
COR_VERDE = "#1E7E34"
COR_VERMELHO = "#C62828"
COR_CINZA = "#5A6472"

ARQUIVO_TEMPLATE = "Dados_Exportacao_CSV.csv"
NOME_RELATORIO = "Relatorio_Auditoria_Matrizes.xlsx"


def caminho_recurso(nome_arquivo):
    """Resolve o caminho de um recurso, funcionando também dentro do .exe."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, nome_arquivo)


def pasta_padrao_saida():
    """Pasta sugerida para salvar arquivos (Documentos do usuário)."""
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    return docs if os.path.isdir(docs) else os.path.expanduser("~")


class RedirecionadorTexto:
    """Captura o que o motor escreve em stdout e entrega linha a linha à fila."""

    def __init__(self, fila):
        self.fila = fila
        self._buffer = ""

    def write(self, texto):
        self._buffer += texto
        while "\n" in self._buffer:
            linha, self._buffer = self._buffer.split("\n", 1)
            self.fila.put(("log", linha))

    def flush(self):
        if self._buffer:
            self.fila.put(("log", self._buffer))
            self._buffer = ""


class AplicacaoAuditoria(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Auditoria Regulatória de Matrizes Curriculares")
        self.geometry("1080x800")
        self.minsize(960, 720)
        self.configure(bg=COR_FUNDO)

        self.fila = queue.Queue()
        self.caminho_csv = tk.StringVar(value=caminho_recurso(ARQUIVO_TEMPLATE))
        self.caminho_relatorio = None
        self.em_execucao = False

        self._montar_estilos()
        self._montar_interface()
        self.after(100, self._drenar_fila)

    # ------------------------------------------------------------------ estilo
    def _montar_estilos(self):
        estilo = ttk.Style(self)
        try:
            estilo.theme_use("clam")
        except tk.TclError:
            pass
        estilo.configure("TFrame", background=COR_FUNDO)
        estilo.configure("Card.TFrame", background=COR_CARD, relief="flat")
        estilo.configure("TLabel", background=COR_FUNDO, font=("Segoe UI", 10))
        estilo.configure("Card.TLabel", background=COR_CARD, font=("Segoe UI", 10))
        estilo.configure("Titulo.TLabel", background=COR_FUNDO,
                         font=("Segoe UI", 17, "bold"), foreground=COR_NAVY)
        estilo.configure("Sub.TLabel", background=COR_FUNDO,
                         font=("Segoe UI", 10), foreground=COR_CINZA)
        estilo.configure("Secao.TLabel", background=COR_CARD,
                         font=("Segoe UI", 11, "bold"), foreground=COR_AZUL)
        estilo.configure("Acao.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8))
        estilo.configure("Primaria.TButton", font=("Segoe UI", 11, "bold"), padding=(18, 10))
        estilo.configure("TProgressbar", background=COR_AZUL, troughcolor="#E3E8EF")

    # -------------------------------------------------------------- construção
    def _montar_interface(self):
        raiz = ttk.Frame(self, padding=(18, 14, 18, 12))
        raiz.pack(fill="both", expand=True)

        ttk.Label(raiz, text="Auditoria Regulatória de Matrizes Curriculares",
                  style="Titulo.TLabel").pack(anchor="w")
        ttk.Label(raiz, text="Validação de carga horária com CP-SAT segundo as regras do MEC e as DCNs",
                  style="Sub.TLabel").pack(anchor="w", pady=(2, 10))

        self._montar_aviso(raiz)
        self._montar_entrada(raiz)
        # resultados é ancorado na base ANTES do log: assim reserva seu espaço e
        # o painel de execução fica com a sobra, sem empurrar nada para fora da tela
        self._montar_resultados(raiz)
        self._montar_execucao(raiz)

    def _montar_aviso(self, pai):
        aviso = tk.Frame(pai, bg=COR_ALERTA_BG, highlightbackground=COR_ALERTA_BORDA,
                         highlightthickness=2)
        aviso.pack(fill="x", pady=(0, 12))
        tk.Label(aviso, text="⚠  ATENÇÃO — DADOS FICTÍCIOS", bg=COR_ALERTA_BG,
                 fg=COR_ALERTA_TXT, font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=14, pady=(9, 0))
        tk.Label(aviso, bg=COR_ALERTA_BG, fg=COR_ALERTA_TXT, justify="left",
                 font=("Segoe UI", 9),
                 text=("A base que acompanha esta ferramenta é SINTÉTICA, criada apenas para demonstrar o "
                       "funcionamento do código.\nOs cursos, matrizes e cargas horárias não correspondem a "
                       "nenhuma instituição real e os resultados NÃO têm validade regulatória.\n"
                       "Para uso real, baixe o template, preencha com os dados da sua IES e execute a auditoria."
                       )).pack(anchor="w", padx=14, pady=(2, 10))

    def _montar_entrada(self, pai):
        card = tk.Frame(pai, bg=COR_CARD, highlightbackground="#DDE3EA", highlightthickness=1)
        card.pack(fill="x", pady=(0, 12))

        ttk.Label(card, text="1. Base de dados", style="Secao.TLabel").pack(anchor="w", padx=14, pady=(10, 6))

        linha = tk.Frame(card, bg=COR_CARD)
        linha.pack(fill="x", padx=14, pady=(0, 4))
        ttk.Button(linha, text="⬇  Baixar template CSV", style="Acao.TButton",
                   command=self.baixar_template).pack(side="left")
        ttk.Button(linha, text="📂  Selecionar arquivo CSV…", style="Acao.TButton",
                   command=self.selecionar_csv).pack(side="left", padx=(8, 0))
        ttk.Button(linha, text="↺  Usar base de exemplo", style="Acao.TButton",
                   command=self.usar_exemplo).pack(side="left", padx=(8, 0))

        # width=1 evita que um caminho longo force a janela a crescer; o fill="x" expande depois
        self.rotulo_arquivo = tk.Label(card, bg=COR_CARD, fg=COR_CINZA, anchor="w", width=1,
                                       font=("Consolas", 9), wraplength=980, justify="left")
        self.rotulo_arquivo.pack(fill="x", padx=14, pady=(4, 12))
        self._atualizar_rotulo_arquivo()

    def _montar_execucao(self, pai):
        card = tk.Frame(pai, bg=COR_CARD, highlightbackground="#DDE3EA", highlightthickness=1)
        card.pack(fill="both", expand=True)

        cabecalho = tk.Frame(card, bg=COR_CARD)
        cabecalho.pack(fill="x", padx=14, pady=(10, 6))
        ttk.Label(cabecalho, text="2. Execução — passo a passo", style="Secao.TLabel").pack(side="left")

        self.botao_executar = ttk.Button(cabecalho, text="▶  Executar Auditoria",
                                         style="Primaria.TButton", command=self.executar)
        self.botao_executar.pack(side="right")

        self.progresso = ttk.Progressbar(card, mode="determinate", maximum=100)
        self.progresso.pack(fill="x", padx=14, pady=(0, 8))

        moldura = tk.Frame(card, bg=COR_CARD)
        moldura.pack(fill="both", expand=True, padx=14, pady=(0, 12))

        # width/height pequenos: o widget expande pelo grid, mas não dita o tamanho da janela
        self.log = tk.Text(moldura, wrap="none", bg="#0F1B2D", fg="#D6E1F0",
                           insertbackground="#D6E1F0", font=("Consolas", 9),
                           relief="flat", padx=10, pady=8, state="disabled",
                           width=40, height=8)
        barra_v = ttk.Scrollbar(moldura, orient="vertical", command=self.log.yview)
        barra_h = ttk.Scrollbar(moldura, orient="horizontal", command=self.log.xview)
        self.log.configure(yscrollcommand=barra_v.set, xscrollcommand=barra_h.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        barra_v.grid(row=0, column=1, sticky="ns")
        barra_h.grid(row=1, column=0, sticky="ew")
        moldura.rowconfigure(0, weight=1)
        moldura.columnconfigure(0, weight=1)

        # realce por tipo de linha
        self.log.tag_configure("etapa", foreground="#7FD1FF", font=("Consolas", 9, "bold"))
        self.log.tag_configure("matriz", foreground="#FFFFFF", font=("Consolas", 9, "bold"))
        self.log.tag_configure("regra", foreground="#A8C4E0")
        self.log.tag_configure("valida", foreground="#6BE58A", font=("Consolas", 9, "bold"))
        self.log.tag_configure("inviavel", foreground="#FF8A80", font=("Consolas", 9, "bold"))
        self.log.tag_configure("falha", foreground="#FFC46B")
        self.log.tag_configure("titulo", foreground="#C9A227", font=("Consolas", 9, "bold"))

        self._escrever_boas_vindas()

    def _montar_resultados(self, pai):
        card = tk.Frame(pai, bg=COR_CARD, highlightbackground="#DDE3EA", highlightthickness=1)
        card.pack(side="bottom", fill="x", pady=(12, 0))

        ttk.Label(card, text="3. Resultado", style="Secao.TLabel").pack(anchor="w", padx=14, pady=(10, 6))

        painel = tk.Frame(card, bg=COR_CARD)
        painel.pack(fill="x", padx=14, pady=(0, 10))

        self.kpis = {}
        for chave, rotulo, cor in [("total", "Matrizes analisadas", COR_NAVY),
                                   ("conformes", "Conformes", COR_VERDE),
                                   ("nao_conformes", "Não conformes", COR_VERMELHO),
                                   ("taxa", "Taxa de conformidade", COR_AZUL)]:
            bloco = tk.Frame(painel, bg=COR_CARD)
            bloco.pack(side="left", expand=True, fill="x")
            valor = tk.Label(bloco, text="—", bg=COR_CARD, fg=cor, font=("Segoe UI", 20, "bold"))
            valor.pack(anchor="w")
            tk.Label(bloco, text=rotulo, bg=COR_CARD, fg=COR_CINZA,
                     font=("Segoe UI", 9)).pack(anchor="w")
            self.kpis[chave] = valor

        acoes = tk.Frame(card, bg=COR_CARD)
        acoes.pack(fill="x", padx=14, pady=(0, 12))
        self.botao_abrir = ttk.Button(acoes, text="📊  Abrir relatório Excel", style="Acao.TButton",
                                      command=self.abrir_relatorio, state="disabled")
        self.botao_abrir.pack(side="left")
        self.botao_salvar = ttk.Button(acoes, text="💾  Salvar relatório como…", style="Acao.TButton",
                                       command=self.salvar_relatorio_como, state="disabled")
        self.botao_salvar.pack(side="left", padx=(8, 0))

    # ------------------------------------------------------------------- log
    def _escrever(self, texto, tag=None):
        self.log.configure(state="normal")
        self.log.insert("end", texto + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _escrever_boas_vindas(self):
        self._escrever("Pronto para executar.", "titulo")
        self._escrever("")
        self._escrever("Regras aplicadas nesta auditoria:", "etapa")
        for regra in [
            "  • Extensão .......... mínimo de 10% da CH total (todos os cursos)",
            "  • EAD ............... máximo de 40% da CH (apenas cursos Presenciais)",
            "  • ASM ............... máximo de 20% da CH (quando o curso possui ASM)",
            "  • Estágio (Saúde) ... Medicina >= 35% | demais cursos da Saúde >= 20%",
            "  • Estágio (Engenh.) . mínimo de 160h absolutas",
            "  • Estágio (Licenc.) . mínimo de 400h | Prática Pedagógica >= 400h",
            "  • Prática indevida .. CH de prática deve ser zero em cursos cuja DCN não",
            "                        prevê formação prática (Economia, Ciência da",
            "                        Computação e Sistemas de Informação)",
            "  • Integridade ....... soma dos componentes deve fechar 100% da CH total",
        ]:
            self._escrever(regra, "regra")
        self._escrever("")
        self._escrever("Selecione uma base de dados e clique em 'Executar Auditoria'.", "regra")

    def _classificar(self, linha):
        if linha.startswith("[ETAPA"):
            return "etapa"
        if linha.startswith("(") and "Analisando Matriz" in linha:
            return "matriz"
        if "RESULTADO: VÁLIDA" in linha:
            return "valida"
        if "RESULTADO: INVIÁVEL" in linha:
            return "inviavel"
        if "NÃO CONFORMIDADE" in linha or "[Erro]" in linha:
            return "falha"
        if linha.startswith("=") or "AUDITORIA" in linha or "SISTEMA DE" in linha:
            return "titulo"
        return "regra"

    # -------------------------------------------------------------- comandos
    def _atualizar_rotulo_arquivo(self):
        caminho = self.caminho_csv.get()
        eh_exemplo = os.path.abspath(caminho) == os.path.abspath(caminho_recurso(ARQUIVO_TEMPLATE))
        marca = "  (base de exemplo — dados fictícios)" if eh_exemplo else ""
        self.rotulo_arquivo.configure(text=f"Arquivo selecionado:  {caminho}{marca}")

    def usar_exemplo(self):
        self.caminho_csv.set(caminho_recurso(ARQUIVO_TEMPLATE))
        self._atualizar_rotulo_arquivo()

    def selecionar_csv(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a base de matrizes (CSV separado por ponto e vírgula)",
            filetypes=[("Arquivos CSV", "*.csv"), ("Todos os arquivos", "*.*")])
        if caminho:
            self.caminho_csv.set(caminho)
            self._atualizar_rotulo_arquivo()

    def baixar_template(self):
        origem = caminho_recurso(ARQUIVO_TEMPLATE)
        if not os.path.exists(origem):
            messagebox.showerror("Template não encontrado",
                                 f"O arquivo {ARQUIVO_TEMPLATE} não foi localizado no pacote.")
            return
        destino = filedialog.asksaveasfilename(
            title="Salvar template CSV", defaultextension=".csv",
            initialfile=ARQUIVO_TEMPLATE, initialdir=pasta_padrao_saida(),
            filetypes=[("Arquivos CSV", "*.csv")])
        if not destino:
            return
        try:
            with open(origem, "rb") as f_origem, open(destino, "wb") as f_destino:
                f_destino.write(f_origem.read())
        except OSError as erro:
            messagebox.showerror("Erro ao salvar", str(erro))
            return
        if messagebox.askyesno(
                "Template salvo",
                f"Template salvo em:\n{destino}\n\n"
                "Ele contém a base fictícia de exemplo, que serve de modelo de preenchimento:\n"
                "mantenha o cabeçalho e o separador ponto e vírgula (;) e substitua as linhas "
                "pelos dados da sua instituição.\n\nDeseja abrir a pasta agora?"):
            self._abrir_no_sistema(os.path.dirname(destino))

    def executar(self):
        if self.em_execucao:
            return
        caminho = self.caminho_csv.get()
        if not os.path.exists(caminho):
            messagebox.showerror("Arquivo não encontrado",
                                 f"Não foi possível localizar:\n{caminho}")
            return

        self.em_execucao = True
        self.botao_executar.configure(state="disabled", text="⏳  Executando…")
        self.botao_abrir.configure(state="disabled")
        self.botao_salvar.configure(state="disabled")
        self.progresso.configure(value=0)
        for kpi in self.kpis.values():
            kpi.configure(text="—")

        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

        threading.Thread(target=self._rodar_auditoria, args=(caminho,), daemon=True).start()

    def _rodar_auditoria(self, caminho_csv):
        """Executa o motor numa thread separada, capturando a saída passo a passo."""
        stdout_original = sys.stdout
        sys.stdout = RedirecionadorTexto(self.fila)
        try:
            print("=" * 78)
            print("      SISTEMA DE AUDITORIA REGULATÓRIA DE MATRIZES CURRICULARES - V2")
            print("=" * 78)
            print("      ATENÇÃO: a base de exemplo é fictícia, apenas para demonstração.")
            print("=" * 78)

            print("\n[ETAPA 1/5] Ingestão de Dados...")
            df = PLANEJADOR.carregar_dados(caminho_csv)
            if df is None or df.empty:
                print("  [Erro]: não foi possível ler o arquivo ou ele está vazio.")
                self.fila.put(("erro", "Não foi possível ler o arquivo CSV ou ele está vazio."))
                return

            faltando = [c for c in ("ID_Matriz", "Curso", "Area", "Modalidade", "Carga_Horaria")
                        if c not in df.columns]
            if faltando:
                print(f"  [Erro]: colunas obrigatórias ausentes: {', '.join(faltando)}")
                self.fila.put(("erro",
                               "O arquivo não tem as colunas obrigatórias: "
                               + ", ".join(faltando)
                               + ".\n\nBaixe o template para conferir o formato esperado."))
                return

            total_matrizes = df["ID_Matriz"].nunique()
            print(f"  * Arquivo carregado: {os.path.basename(caminho_csv)}")
            print(f"  * Total de disciplinas importadas: {len(df)}")
            print(f"  * Total de matrizes curriculares: {total_matrizes}")
            self.fila.put(("progresso", 8))

            print("\n[ETAPA 2/5] Configuração dos Parâmetros (MEC & DCNs)...")
            print("  * EAD: máximo 40% (cursos Presenciais).")
            print("  * Extensão: mínimo 10%.")
            print("  * Saúde/Medicina: Estágio >= 35%; demais Saúde >= 20%.")
            print("  * Engenharias: Estágio >= 160h.")
            print("  * Licenciaturas: Estágio >= 400h | Prática >= 400h.")
            print("  * Prática indevida: CH de prática = 0 em cursos cuja DCN não prevê prática")
            print(f"    ({', '.join(sorted(PLANEJADOR.CURSOS_SEM_EXIGENCIA_PRATICA))}).")
            print("  * ASM: máximo 20%.")
            print("  * Regra matemática: soma da grade deve fechar 100%.")
            self.fila.put(("progresso", 15))

            print("\n[ETAPA 3/5] Processamento com CP-SAT Solver...")
            resultados = self._validar_com_progresso(df, total_matrizes)

            print("\n[ETAPA 4/5] Consolidação dos Resultados...")
            validas = sum(1 for r in resultados if "VÁLIDA" in r["Status_Geral"])
            invalidas = len(resultados) - validas
            taxa = (validas / len(resultados) * 100) if resultados else 0.0
            print(f"  * Matrizes em conformidade: {validas}")
            print(f"  * Matrizes inválidas: {invalidas}")
            print(f"  * Taxa geral de conformidade: {taxa:.1f}%")
            self.fila.put(("progresso", 92))

            print("\n[ETAPA 5/5] Geração de Relatório de Auditoria...")
            destino = os.path.join(pasta_padrao_saida(), NOME_RELATORIO)
            try:
                PLANEJADOR.gerar_relatorio_excel(resultados, destino)
            except PermissionError:
                print("  * [Erro]: permissão negada — o relatório provavelmente está aberto no Excel.")
                self.fila.put(("erro", f"Não foi possível gravar o relatório:\n{destino}\n\n"
                                       "Feche o arquivo no Excel e execute novamente."))
                return
            print(f"  * Relatório gerado em: {destino}")
            print("\n" + "=" * 78)
            print("AUDITORIA FINALIZADA")
            print("=" * 78)

            self.fila.put(("progresso", 100))
            self.fila.put(("concluido", {"total": len(resultados), "conformes": validas,
                                         "nao_conformes": invalidas, "taxa": taxa,
                                         "relatorio": destino}))
        except Exception as erro:  # falha inesperada não pode derrubar a janela
            print(f"\n  [Erro inesperado]: {erro}")
            self.fila.put(("erro", f"Falha inesperada durante a auditoria:\n\n{erro}"))
        finally:
            sys.stdout.flush()
            sys.stdout = stdout_original
            self.fila.put(("fim", None))

    def _validar_com_progresso(self, df, total_matrizes):
        """Roda a validação matriz a matriz para poder atualizar a barra de progresso."""
        resultados = []
        for indice, id_matriz in enumerate(df["ID_Matriz"].unique(), start=1):
            resultados.extend(PLANEJADOR.validar_todas_matrizes(
                df[df["ID_Matriz"] == id_matriz],
                indice_inicial=indice, total_geral=total_matrizes))
            self.fila.put(("progresso", 15 + int(indice / max(total_matrizes, 1) * 75)))
        return resultados

    # ------------------------------------------------------------------ fila
    def _drenar_fila(self):
        try:
            while True:
                tipo, dado = self.fila.get_nowait()
                if tipo == "log":
                    self._escrever(dado, self._classificar(dado))
                elif tipo == "progresso":
                    self.progresso.configure(value=dado)
                elif tipo == "erro":
                    messagebox.showerror("Erro na auditoria", dado)
                elif tipo == "concluido":
                    self._mostrar_resultado(dado)
                elif tipo == "fim":
                    self.em_execucao = False
                    self.botao_executar.configure(state="normal", text="▶  Executar Auditoria")
        except queue.Empty:
            pass
        self.after(100, self._drenar_fila)

    def _mostrar_resultado(self, dados):
        self.kpis["total"].configure(text=str(dados["total"]))
        self.kpis["conformes"].configure(text=str(dados["conformes"]))
        self.kpis["nao_conformes"].configure(text=str(dados["nao_conformes"]))
        self.kpis["taxa"].configure(text=f"{dados['taxa']:.1f}%")
        self.caminho_relatorio = dados["relatorio"]
        self.botao_abrir.configure(state="normal")
        self.botao_salvar.configure(state="normal")

    # ------------------------------------------------------------- relatório
    def _abrir_no_sistema(self, caminho):
        try:
            if sys.platform.startswith("win"):
                os.startfile(caminho)  # noqa: S606
            elif sys.platform == "darwin":
                subprocess.run(["open", caminho], check=False)
            else:
                subprocess.run(["xdg-open", caminho], check=False)
        except OSError as erro:
            messagebox.showerror("Não foi possível abrir", str(erro))

    def abrir_relatorio(self):
        if self.caminho_relatorio and os.path.exists(self.caminho_relatorio):
            self._abrir_no_sistema(self.caminho_relatorio)
        else:
            messagebox.showwarning("Relatório indisponível",
                                   "Execute a auditoria antes de abrir o relatório.")

    def salvar_relatorio_como(self):
        if not (self.caminho_relatorio and os.path.exists(self.caminho_relatorio)):
            messagebox.showwarning("Relatório indisponível",
                                   "Execute a auditoria antes de salvar o relatório.")
            return
        destino = filedialog.asksaveasfilename(
            title="Salvar relatório de auditoria", defaultextension=".xlsx",
            initialfile=NOME_RELATORIO, initialdir=pasta_padrao_saida(),
            filetypes=[("Planilha Excel", "*.xlsx")])
        if not destino:
            return
        try:
            with open(self.caminho_relatorio, "rb") as origem, open(destino, "wb") as saida:
                saida.write(origem.read())
            messagebox.showinfo("Relatório salvo", f"Relatório salvo em:\n{destino}")
        except OSError as erro:
            messagebox.showerror("Erro ao salvar", str(erro))


def autoteste():
    """Roda a auditoria sem abrir a janela e grava um relatório de diagnóstico.

    Útil para conferir, na máquina de quem recebeu o programa, se o pacote está
    íntegro (solver, base de exemplo e geração do Excel). Uso:
        Auditoria_Matrizes.exe --autoteste
    """
    destino_log = os.path.join(pasta_padrao_saida(), "Auditoria_autoteste.txt")
    linhas = []

    def registrar(texto):
        linhas.append(texto)

    try:
        registrar(f"Python .............. {sys.version.split()[0]}")
        registrar(f"Empacotado (frozen) . {getattr(sys, 'frozen', False)}")
        registrar(f"Base de recursos .... {getattr(sys, '_MEIPASS', 'execução a partir do código')}")

        base = caminho_recurso(ARQUIVO_TEMPLATE)
        registrar(f"Base de exemplo ..... {base}")
        registrar(f"  existe? ........... {os.path.exists(base)}")

        df = PLANEJADOR.carregar_dados(base)
        if df is None or df.empty:
            registrar("FALHA: não foi possível ler a base de exemplo.")
            return 1
        registrar(f"  disciplinas ....... {len(df)}")
        registrar(f"  matrizes .......... {df['ID_Matriz'].nunique()}")

        resultados = PLANEJADOR.validar_todas_matrizes(df)
        validas = sum(1 for r in resultados if "VÁLIDA" in r["Status_Geral"])
        registrar(f"Solver CP-SAT ....... OK ({len(resultados)} matrizes avaliadas)")
        registrar(f"  conformes ......... {validas}")
        registrar(f"  não conformes ..... {len(resultados) - validas}")

        saida = os.path.join(pasta_padrao_saida(), NOME_RELATORIO)
        PLANEJADOR.gerar_relatorio_excel(resultados, saida)
        registrar(f"Relatório Excel ..... OK ({os.path.getsize(saida) / 1024:.0f} KB)")
        registrar(f"  gravado em ........ {saida}")
        registrar("")
        registrar("RESULTADO: pacote íntegro, todas as etapas concluíram.")
        return 0
    except Exception as erro:
        registrar("")
        registrar(f"RESULTADO: FALHA — {type(erro).__name__}: {erro}")
        return 1
    finally:
        conteudo = "\n".join(["AUTOTESTE — Auditoria Regulatória de Matrizes Curriculares", ""] + linhas)
        try:
            # utf-8-sig: o BOM faz o Bloco de Notas e o PowerShell exibirem os acentos
            with open(destino_log, "w", encoding="utf-8-sig") as arquivo:
                arquivo.write(conteudo + "\n")
        except OSError:
            pass
        print(conteudo)


if __name__ == "__main__":
    if "--autoteste" in sys.argv:
        sys.exit(autoteste())
    AplicacaoAuditoria().mainloop()
