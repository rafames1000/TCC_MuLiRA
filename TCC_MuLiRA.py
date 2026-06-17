import shutil
import os.path
import pandas as pd
from pandas import Series
import statsmodels.api as sm
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
from tkinter.ttk import *
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox


class Janela(Tk):
    """Classe principal que gerencia o ciclo de vida do software e a troca de frames."""

    def __init__(self):
        Tk.__init__(self)
        self.title("MuLiRA Py-Ex")
        self._frame = None
        self.state("zoomed")  # Inicia o MuLiRA em tela cheia
        self.switch_frame(Menu)  # Inicia o Menu como tela inicial

    def switch_frame(self, frame_class):
        """Destrói o frame atual e renderiza o próximo."""

        new_frame = frame_class(self)

        if self._frame is not None:
            self._frame.destroy()

        self._frame = new_frame

        # Ajusta o layout dinamicamente conforme o tipo de tela (Ex: Menu, Formulário, etc.)
        if frame_class == Formulario:
            self.geometry("975x520+200+100")
            self.minsize(width=600, height=480)
            self._frame.grid(row=0, column=0, sticky="nsew")

        else:
            self.geometry("850x500+200+100")
            self.minsize(width=600, height=480)
            self.configure(bg="grey")
            self._frame.pack()


class Menu(Frame):
    """
    Classe responsável pela interface inicial do software.

    Gerencia a navegação principal e a escolha entre importar ou selecionar uma planilha.
    """

    def __init__(self, master: Janela):
        super().__init__(master)

        Frame.configure(self, bg="grey")

        Label(self, text="Menu", bg="grey", fg="navy blue", font=('Concert One', 50, "bold")
              ).pack(side="top", fill="x", pady=75)

        self.cor_original_bg = 'light grey'
        self.cor_hover_bg = '#F5F5F5'  # Um cinza um pouco mais claro
        self.cor_hover_fg = 'navy blue'
        self.cor_original_fg = 'black'

        self.btn_importar = Button(self, text="       Importar Planilha       ",
                                   bg=self.cor_original_bg, fg=self.cor_original_fg,
                                   font=('Concert One', 32, "bold"),
                                   command=lambda: importar_planilha(),
                                   relief="raised", borderwidth=3)

        self.btn_importar.pack(pady=25)

        self.btn_selecionar = Button(self, text="Selecionar Planilha Salva",
                                     bg=self.cor_original_bg, fg=self.cor_original_fg,
                                     font=('Concert One', 32, "bold"),
                                     command=lambda: master.switch_frame(Formulario),
                                     relief="raised", borderwidth=3)

        self.btn_selecionar.pack()

        # Vinculação de eventos (Binding) para melhorar a experiência do usuário
        self.btn_importar.bind("<Enter>", self.on_enter)
        self.btn_importar.bind("<Leave>", self.on_leave)
        self.btn_selecionar.bind("<Enter>", self.on_enter)
        self.btn_selecionar.bind("<Leave>", self.on_leave)

    # Requisito de Domínio (RD: 001)
    def on_enter(self, event):
        """Altera as cores do widget quando o mouse entra na área do botão."""

        widget = event.widget
        widget.config(bg=self.cor_hover_bg, fg=self.cor_hover_fg)

    def on_leave(self, event):
        """Restaura as cores originais quando o mouse sai da área do botão."""

        widget = event.widget
        widget.config(bg=self.cor_original_bg, fg=self.cor_original_fg)


def explorador_de_arquivos(n):
    """
    Interface de busca de arquivos no Sistema Operacional.

    Restringe o usuário ao formato Excel (.xlsx) compatível com a biblioteca Pandas.
    """

    if n == 0:
        # Define o caminho dinâmico para a área de trabalho do usuário logado
        desktop = os.path.expanduser("~\\Desktop")
        filename = filedialog.askopenfilename(initialdir=desktop,
                                              title="Selecione uma Planilha",
                                              filetypes=(("Arquivos Excel", "*.xlsx*"),))
        return filename

    elif n == 1:
        # Define o caminho para a pasta onde o executável/script está rodando
        desktop = os.path.expanduser(os.getcwd())
        filename = filedialog.askopenfilename(initialdir=desktop,
                                              title="Selecione uma Planilha",
                                              filetypes=(("Arquivos Excel", "*.xlsx*"),))
        return filename

    return None


def importar_planilha():
    """Localiza o arquivo original e cria uma cópia dentro do diretório do MuLiRA."""

    file_name = explorador_de_arquivos(0)

    a = "\\".join(os.path.split(os.getcwd()))  # Diretório raiz do projeto
    b = os.path.basename(file_name)  # Nome do arquivo com extensão
    c = "\\".join([a, b])  # Caminho de destino final
    shutil.copy2(file_name, c)  # Realiza a cópia física do arquivo

    if b != '':
        d = b[:-5]  # Remove a extensão .xlsx da string para exibir apenas o nome na mensagem
        messagebox.showinfo(title="Sucesso!", message="Planilha Salva: \n " + d)


class Formulario(Frame):
    """
    Classe responsável pela interface de seleção e tratamento de dados.

    Permite ao usuário escolher colunas e linhas específicas de uma planilha
    Excel para realizar análises estatísticas.
    """
    master: Janela

    # --- Front-end ---

    def __init__(self, master: Janela):
        """
        Inicializa a interface do formulário, carrega a planilha selecionada
        e configura o layout de scroll e variáveis de controle.
        """

        super().__init__(master)

        # Interface de seleção de arquivo
        self.file_name1 = explorador_de_arquivos(1)
        self.f_n1 = os.path.basename(self.file_name1)

        if self.f_n1 != '':
            self.f_n01 = self.f_n1[:-5]  # Remove a extensão .xlsx da string para exibir apenas o nome na mensagem
            messagebox.showinfo(title="Sucesso!", message="Planilha Selecionada: \n " + self.f_n01)

        self.df = pd.read_excel(self.file_name1)  # Carrega os dados para memória

        Frame.__init__(self, master)
        Frame.configure(self, bg='grey')
        master.grid_columnconfigure(0, weight=1)
        master.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- Declarando as variáveis da Classe ---
        (
            self.res_col, self.index_col, self.res_col_2, self.index_col_2, self.index_col_2_aux,
            self.res_ln, self.index_ln, self.res_ln_2, self.index_ln_2, self.index_ln_2_aux,
            self.nome_fatia_col, self.index_fatia_col, self.index_fatia_col_aux, self.aux_cb_values,
            self.nome_fatia_ln, self.index_fatia_ln, self.index_fatia_ln_aux, self.aux_cb_2_values
        ) = [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], []

        self.alt_cb = self.undo = self.ativado_2pts = self.fatiando = self.alt_cb_2 = self.undo_2 \
            = self.ativado_2pts_2 = self.fatiando_2 = self.tratar_dados_clicado = False

        self.a_col = self.b_col = self.a_col_nome = self.b_col_nome = self.a_ln = self.b_ln \
            = self.a_ln_nome = self.b_ln_nome = self.janela_secundaria \
            = self.rlm_txt_summary = self.rlm_btn_salvar = self.rlm_frame_graficos = self.rlm_btn_salvar_tabela \
            = self.rlm_combo_y = self.cols_para_verificar = self.summary_var = None

        (
            self.df_select_bruto, self.df_select_bruto_aux, self.df_select_bruto_aux_2,
            self.df_select_bruto_aux_3, self.df_todos_dados, self.df_validos
        ) = pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


        self.list_col, self.list_ln = list(self.df.columns), [row for idx, *row in self.df.itertuples()]
        self.list_ln_cb_2_format = [" ".join([str(x) for x in item]) for item in self.list_ln]
        # Dicionário para converter a string do Combobox(cb_2) para o formato de lista padrão
        self.dict_cb_2_str_list_ln = {string_item: original_item for string_item, original_item
                                      in zip(self.list_ln_cb_2_format, self.list_ln)}

        self.list_ln_str_format = [" | ".join([str(x) for x in item]) for item in self.list_ln]
        # Dicionário para converter a string de exibição(ex: res_ln_var) para o formato de lista padrão
        self.dict_str_list_ln = {string_item: original_item for string_item, original_item
                                 in zip(self.list_ln_str_format, self.list_ln)}

        # Dicionário para converter a string de exibição(ex: res_ln_var) para o formato de string do Combobox(cb_2)
        self.dict_convert_str_to_cb_2 = {string_item: original_item for string_item, original_item
                                         in zip(self.list_ln_str_format, self.list_ln_cb_2_format)}

        self.zip_col, self.zip_ln = [()], [()]
        self.lim_char_ln = 30  # Limite inicial de 30 caracteres por linha

        def criar_frame_rolavel(frame_pai, event=None):
            """
            Cria uma estrutura de scroll (barra de rolagem) customizada.

            Retorna o canvas, o frame interno e o seu ID.
            """
            canvas_ = Canvas(frame_pai, bg="light grey")
            scrollbar_v = Scrollbar(frame_pai, orient="vertical", command=canvas_.yview)
            scrollbar_h = Scrollbar(frame_pai, orient="horizontal", command=canvas_.xview)

            if frame_pai == self:
                canvas_.grid(row=0, column=0, sticky="nsew")
                scrollbar_v.grid(row=0, column=1, sticky="ns")
                scrollbar_h.grid(row=1, column=0, sticky="ew")

            elif event == 1:  # Layout para a seleção de colunas
                canvas_.grid(row=6, columnspan=5, sticky="nsew")
                scrollbar_v.grid(row=6, column=5, sticky="ns")
                scrollbar_h.grid(row=7, columnspan=5, sticky="ew")

            elif event == 2:  # Layout para a seleção de linhas
                canvas_.grid(row=12, columnspan=5, sticky="nsew")
                scrollbar_v.grid(row=12, column=5, sticky="ns")
                scrollbar_h.grid(row=13, columnspan=5, sticky="ew")

            canvas_.configure(yscrollcommand=scrollbar_v.set)
            canvas_.configure(xscrollcommand=scrollbar_h.set)

            # Cria o frame interno onde os widgets (checkboxes, labels) serão inseridos
            inner_frame = Frame(canvas_)
            canvas_frame_id_ = canvas_.create_window((0, 0), window=inner_frame, anchor="nw")

            return canvas_, inner_frame, canvas_frame_id_

        def update_scroll_region(canvas_, frame, frame_id, _=None):
            """
            Atualiza dinamicamente a área de rolagem.

            Calcula o tamanho do conteúdo interno e ajusta o canvas para que
            o usuário consiga rolar até o último item da lista.
            """

            # Atualiza o cálculo das dimensões reais dos widgets internos
            frame.update_idletasks()

            sf_actual_width = frame.winfo_reqwidth()
            sf_actual_height = frame.winfo_reqheight()
            canvas_current_width = canvas_.winfo_width()

            offset_x = 0

            if sf_actual_width < canvas_current_width:  # Centraliza o conteúdo caso o frame seja menor que a tela
                offset_x = (canvas_current_width - sf_actual_width) / 2

            # Reposiciona o frame interno dentro do canvas
            canvas_.coords(frame_id, offset_x, 0)
            canvas_.itemconfig(frame_id, width=sf_actual_width)

            # Define o limite de área do scroll
            if sf_actual_width <= canvas_current_width:
                canvas_.configure(scrollregion=(0, 0, canvas_current_width - 4, sf_actual_height))

            else:
                canvas_.configure(scrollregion=(0, 0, sf_actual_width, sf_actual_height))

        def bind_scroll_events_to_children(frame, enter_func, leave_func):
            """
            Aplica eventos de mouse a todos os elementos filhos de um frame.

            Isso permite que o scroll do mouse funcione mesmo quando o cursor
            estiver em cima de um Label ou Checkbutton específico.
            """

            for widget in frame.winfo_children():
                widget.bind("<Enter>", enter_func)
                widget.bind("<Leave>", leave_func)

        canvas, self.scrollable_frame, canvas_frame_id = criar_frame_rolavel(self)

        canvas.bind("<Configure>",
                    lambda e: update_scroll_region(canvas, self.scrollable_frame, canvas_frame_id, e))

        self.scrollable_frame.bind("<Configure>",
                                   lambda e: update_scroll_region(canvas, self.scrollable_frame, canvas_frame_id, e))

        canvas_col_sel, scrollable_frame_col_sel, canvas_frame_id_col_sel = criar_frame_rolavel(self.scrollable_frame,
                                                                                                1)

        canvas_col_sel.bind("<Configure>",
                            lambda e: update_scroll_region(canvas_col_sel, scrollable_frame_col_sel,
                                                           canvas_frame_id_col_sel, e))

        scrollable_frame_col_sel.bind("<Configure>",
                                      lambda e: update_scroll_region(canvas_col_sel, scrollable_frame_col_sel,
                                                                     canvas_frame_id_col_sel, e))

        canvas_ln_sel, scrollable_frame_ln_sel, canvas_frame_id_ln_sel = criar_frame_rolavel(self.scrollable_frame, 2)

        canvas_ln_sel.bind("<Configure>",
                           lambda e: update_scroll_region(canvas_ln_sel, scrollable_frame_ln_sel,
                                                          canvas_frame_id_ln_sel, e))

        scrollable_frame_ln_sel.bind("<Configure>",
                                     lambda e: update_scroll_region(canvas_ln_sel, scrollable_frame_ln_sel,
                                                                    canvas_frame_id_ln_sel, e))

        def rebind_scroll_wheel(canvas_):
            """
            Vincula novamente o evento de rolagem do mouse ao canvas específico.

            Necessário para garantir que o comando de rolagem volte ao canvas específico
            após o fechamento de menus suspensos (Comboboxes) ou trocas de foco.
            """
            master.bind("<MouseWheel>", lambda e: scroll_wheel(e, canvas_))

        def scroll_wheel(event, canvas_):
            """
            Gerencia o deslocamento do scroll e evita conflitos com Comboboxes.

            Verifica se alguma lista suspensa está aberta; se estiver, prioriza
            o fechamento/uso dela antes de permitir a rolagem do painel.
            """
            comboboxes = [self.cb, self.cb_2]
            prevent_canvas_scroll = False

            for cb in comboboxes:

                if 'pressed' in cb.state():
                    master.unbind("<MouseWheel>")
                    cb.event_generate('<Button-1>')
                    master.after(100, rebind_scroll_wheel, canvas_)
                    prevent_canvas_scroll = True
                    break

            if not prevent_canvas_scroll:
                canvas_.yview_scroll(int(-1 * (event.delta / 120)), "units")

        master.bind("<MouseWheel>", lambda e: scroll_wheel(e, canvas))

        def update_widgets_after_trace(event=None):
            """
            Atualiza a área de scroll e os eventos dos filhos após mudanças dinâmicas.

            Garante que, se novos itens forem adicionados (como novas colunas),
            a barra de rolagem e o foco do mouse se ajustem automaticamente.
            """

            if event == 1:  # Atualiza seção de seleção de colunas
                update_scroll_region(canvas_col_sel, scrollable_frame_col_sel, canvas_frame_id_col_sel)
                bind_scroll_events_to_children(scrollable_frame_col_sel, on_enter_col_sel, on_leave_func)

            elif event == 2:  # Atualiza seção de seleção de linhas
                update_scroll_region(canvas_ln_sel, scrollable_frame_ln_sel, canvas_frame_id_ln_sel)
                bind_scroll_events_to_children(scrollable_frame_ln_sel, on_enter_ln_sel, on_leave_func)

        def on_enter_col_sel(_):
            """Redireciona o scroll do mouse para o canvas de seleção de colunas ao passar o mouse."""
            master.unbind("<MouseWheel>")
            master.bind("<MouseWheel>", lambda e: scroll_wheel(e, canvas_col_sel))

        def on_enter_ln_sel(_):
            """Redireciona o scroll do mouse para o canvas de seleção de linhas ao passar o mouse."""
            master.unbind("<MouseWheel>")
            master.bind("<MouseWheel>", lambda e: scroll_wheel(e, canvas_ln_sel))

        def on_leave_func(_):
            """Retorna o controle do scroll para o canvas principal ao retirar o mouse das listas."""
            master.unbind("<MouseWheel>")
            master.bind("<MouseWheel>", lambda e: scroll_wheel(e, canvas))

        canvas_col_sel.bind("<Enter>", on_enter_col_sel)
        canvas_col_sel.bind("<Leave>", on_leave_func)
        canvas_ln_sel.bind("<Enter>", on_enter_ln_sel)
        canvas_ln_sel.bind("<Leave>", on_leave_func)

        bind_scroll_events_to_children(canvas_col_sel, on_enter_col_sel, on_leave_func)
        bind_scroll_events_to_children(canvas_ln_sel, on_enter_ln_sel, on_leave_func)

        Label(self.scrollable_frame, text="     Formulário     ", bg="light grey", fg="navy blue",
              font=('Concert One', 32, "bold")).grid(column=0, row=0, columnspan=5)

        Label(self.scrollable_frame, text="                              "
                                          'Selecione as "Colunas" do Excel que irá utilizar:'
                                          "                              ",
              font=('Concert One', 18, "bold")).grid(column=0, row=2, columnspan=5, pady=15)

        self.cb = Combobox(self.scrollable_frame, width=10, state='readonly', font=('Concert One', 18, "bold"))
        self.cb['values'] = self.list_col
        self.option_add('*TCombobox*Listbox.font', 18)
        self.cb.grid(column=1, row=3, pady=20)
        self.cb.bind('<<ComboboxSelected>>', self.exibir_selecao_coluna)

        self.botao_2pts_col = \
            (Button(self.scrollable_frame, text=":", command=self.fatiamento_coluna, font=('Concert One', 18, "bold")))

        self.botao_2pts_col.grid(column=2, row=3)

        self.des_col = Button(self.scrollable_frame, text="Desfazer", command=self.desfazer_selecao_coluna,
                              font=('Concert One', 18, "bold"))
        self.des_col.grid(column=3, row=3)

        self.checkvar1 = IntVar()
        self.checkvar1.set(0)

        self.tudo_col = Checkbutton(self.scrollable_frame, text="Tudo", command=self.selecionar_tudo_coluna,
                                    variable=self.checkvar1, font=('Concert One', 12), bg='light grey')
        self.tudo_col.grid(column=4, row=3, padx=15)

        self.alt_txt_cb = StringVar()  # Inicializando textvariable
        self.alt_txt_cb.set('Selecionar coluna ativado')

        Label(self.scrollable_frame, textvariable=self.alt_txt_cb, font=('Concert One', 12), bg='light grey', fg='red'
              ).grid(column=1, row=4)

        def qtde_col_sel(*_args):
            """
            Calcula e atualiza dinamicamente a quantidade de colunas selecionadas.

            Percorre a lista de resultados, verificando se há sublistas (seleções múltiplas)
            ou itens individuais, atualizando o StringVar do Label da interface.
            """
            counts = 0

            for sublista in self.res_col:

                if isinstance(sublista, list):
                    counts += len(sublista)

                else:
                    counts += 1

            self.qtde_col_sel.set(f'Colunas Selecionadas: {counts} ')  # Atualiza o texto na interface em tempo real

        def qtde_ln_sel(*_args):
            """
            Calcula e atualiza dinamicamente a quantidade de linhas selecionadas.

            Funciona de forma análoga à contagem de colunas, garantindo a precisão
            da amostra que será processada nos cálculos estatísticos.
            """
            counts = 0

            for sublista in self.res_ln:

                if isinstance(sublista, list):
                    counts += len(sublista)

                else:
                    counts += 1

            self.qtde_ln_sel.set(f'Linhas Selecionadas: {counts} ')

        self.qtde_col_sel = StringVar()  # Inicializando textvariable
        self.qtde_col_sel.set(f'Colunas Selecionadas: {len(self.res_col)} ')

        Label(self.scrollable_frame, textvariable=self.qtde_col_sel, font=('Concert One', 18, "bold")
              ).grid(column=0, row=5, columnspan=5, pady=40)

        self.qtde_ln_sel = StringVar()  # Inicializando textvariable
        self.qtde_ln_sel.set(f'Linhas Selecionadas: {len(self.res_ln)} ')

        Label(self.scrollable_frame, textvariable=self.qtde_ln_sel, font=('Concert One', 18, "bold")
              ).grid(column=0, row=11, columnspan=5, pady=40)

        self.res_col_var = StringVar()  # Inicializando textvariable
        self.res_col_var.trace("w", qtde_col_sel)  # Atualizando o num de colunas selecionadas (self.qtde_col_sel)
        self.res_col_var.trace_add("write", lambda name, index, mode: update_widgets_after_trace(1))

        self.res_ln_var = StringVar()
        self.res_ln_var.trace("w", qtde_ln_sel)  # Atualizando o num de linhas selecionadas (self.qtde_ln_sel)
        self.res_ln_var.trace_add("write", lambda name, index, mode: update_widgets_after_trace(2))

        Label(scrollable_frame_col_sel, textvariable=self.res_col_var, font=('Concert One', 22)
              ).grid(row=3, column=3, pady=10)

        Label(self.scrollable_frame, text="                              "
                                          'Selecione as "Linhas" do Excel que irá utilizar:'
                                          "                              ",
              font=('Concert One', 18, "bold")).grid(column=0, row=8, columnspan=5, pady=15)

        self.cb_2 = Combobox(self.scrollable_frame, width=10, state='readonly', font=('Concert One', 18, "bold"))
        self.cb_2['values'] = self.list_ln_cb_2_format
        self.option_add('*TCombobox*Listbox.font', 18)
        self.cb_2.grid(column=1, row=9, pady=20)
        self.cb_2.bind('<<ComboboxSelected>>', self.exibir_selecao_linha)

        self.botao_2pts_ln = Button(self.scrollable_frame, text=":", command=self.fatiamento_linha,
                                    font=('Concert One', 18, "bold"))
        self.botao_2pts_ln.grid(column=2, row=9)

        self.des_ln = Button(self.scrollable_frame, text="Desfazer", command=self.desfazer_selecao_linha,
                             font=('Concert One', 18, "bold"))
        self.des_ln.grid(column=3, row=9)

        self.checkvar2 = IntVar()
        self.checkvar2.set(0)

        self.tudo_ln = Checkbutton(self.scrollable_frame, text="Tudo", command=self.selecionar_tudo_linha,
                                   variable=self.checkvar2, font=('Concert One', 12), bg='light grey')
        self.tudo_ln.grid(column=4, row=9, padx=15)

        self.alt_txt_cb_2 = StringVar()  # Inicializando textvariable
        self.alt_txt_cb_2.set('Selecionar linha ativado')

        Label(self.scrollable_frame, textvariable=self.alt_txt_cb_2,
              font=('Concert One', 12), bg='light grey', fg='red').grid(column=1, row=10)

        Label(scrollable_frame_ln_sel, textvariable=self.res_ln_var, font=('Concert One', 22)
              ).grid(row=3, column=3, pady=10)

        def on_var_change(*_args):
            """
            Reseta o estado de tratamento de dados e revalida os botões.

            Sempre que uma seleção de linha ou coluna muda, o tratamento anterior
            torna-se inválido, exigindo um novo processamento.
            """
            self.tratar_dados_clicado = False  # Definindo o estado da variável

            # Verifica se os novos critérios atendem aos requisitos para habilitar os botões
            self.atualizar_estado_botoes()

        self.qtde_col_sel.trace_add("write", self.atualizar_estado_botoes)
        self.qtde_ln_sel.trace_add("write", self.atualizar_estado_botoes)

        Button(self.scrollable_frame, text="Tratar Dados Selecionados", font=('Concert One', 18, "bold"),
               fg='darkgoldenrod', command=self.tratar_dados_selecionados, state='disabled'
               ).grid(row=14, column=1, pady=50)

        self.res_col_var.trace_add('write', on_var_change)
        self.res_ln_var.trace_add('write', on_var_change)

        self.atualizar_estado_botoes()

        Button(self.scrollable_frame, text="Calcular RLM", font=('Concert One', 18, "bold"), fg='green',
               command=self.calcular_rlm, state='disabled').grid(row=14, column=3, pady=50)

        Button(self.scrollable_frame, text="Voltar ao Menu", font=('Concert One', 26, "bold"),
               command=lambda: master.switch_frame(Menu)).grid(row=15, column=2)

    # --- Back-end ---

    def txt_quebra_de_linha(self, itens, event):
        """
        Formata uma lista de strings para exibição na interface, aplicando quebras
        de linha e delimitadores visuais (sublinhados) para melhorar a leitura.

        Args:
            itens (list): Lista de nomes de colunas ou linhas selecionadas.
            event (int): Identificador do tipo de dado (1 para Colunas, 2 para Linhas).

        Returns:
            str: String formatada pronta para ser exibida em um Label ou Text widget.
        """
        itens_formatados = []

        if not itens:
            return ""

        # --- Caso 1: Formatação Exclusiva para LINHAS (event 2) ---
        # Foca em separar cada linha selecionada com molduras de sublinhado
        if event == 2:

            for item in itens:

                if item is itens[0]:  # A primeira linha ganha moldura em cima e embaixo
                    linha_formatada = f"{'_' * len(item)}\n\n{item}\n{'_' * len(item)}"

                else:  # O restante ganha moldura apenas embaixo
                    linha_formatada = f"\n{item}\n{'_' * len(item)}"

                itens_formatados.append(linha_formatada)

            return "\n".join(itens_formatados)

        # --- Caso 2: Formatação para COLUNAS (ou Geral) ---
        # Separa os nomes com ' | ', quebrando a linha quando necessário
        ln = [""]
        separador = " | "

        for item in itens:
            ln_atual = ln[-1]
            # Tenta adicionar o item à linha atual
            new_str = (ln_atual + separador + item) if ln_atual else item

            if len(new_str) > self.lim_char_ln:  # Pula linha se a resultante (atual) exceder o limite

                if ln_atual:  # Se há conteúdo na linha atual, cria outra nova linha para o item
                    ln.append(item)

                else:  # Ajusta o limite se um único item for maior que o limite atual

                    if len(item) > self.lim_char_ln:
                        self.lim_char_ln = len(item)

                    ln[-1] = item

            else:  # Se couber, mantém na mesma linha
                ln[-1] = new_str

        # --- Lógica de formatação visual das linhas usando underline ---
        item_ln_und = []

        if ln[-1] is ln[0]:
            # Caso o texto tenha apenas uma linha
            item_ln_und.append(f"{'_' * len(ln[0])}\n\n{ln[0]}\n{'_' * len(ln[0])}")

        else:

            for item_ln in ln:

                if item_ln is ln[0]:  # Cabeçalho do texto
                    item_ln_und.append(f"{'_' * len(ln[0])}\n\n{ln[0]}\n{'_' * len(ln[0])}")

                else:  # Linhas divisórias com underline
                    item_ln_und.append(f"\n{item_ln}\n{'_' * len(item_ln)}")

        return "\n".join(item_ln_und)

    def ordenar_lista_index(self, event):
        """
        Sincroniza e ordena as listas de índices e nomes após operações de seleção ou
        desseleção, garantindo que a ordem visual corresponda à ordem original do Excel.

        Args:
            event (int): 1 para processar Colunas, 2 para processar Linhas.
        """

        # Verifica se o último item é uma lista (fatiamento) e o "achata" para ordenação
        if event == 1 and self.index_col and isinstance(self.index_col[-1], list):
            a, b, d = self.index_col[-1], self.res_col[-1], self.list_col[-1]
            del self.index_col[-1], self.res_col[-1], self.list_col[-1]

            self.index_col.extend(a)
            self.res_col.extend(b)
            self.list_col.extend(d)

        elif event == 2 and self.index_ln and isinstance(self.index_ln[-1], list):
            a, b, d = self.index_ln[-1], self.res_ln[-1], self.list_ln_cb_2_format[-1]
            del self.index_ln[-1], self.res_ln[-1], self.list_ln_cb_2_format[-1]

            self.index_ln.extend(a)
            self.res_ln.extend(b)
            self.list_ln_cb_2_format.extend(d)

        # Ordenação com zip mantendo a sincronização (índice, nome, lista)
        if event == 1:
            self.zip_col = zip(self.index_col, self.res_col, self.list_col)
            self.zip_col = sorted(self.zip_col, key=lambda c: c[0])
            self.index_col = list(map(lambda c: c[0], self.zip_col))
            self.res_col = list(map(lambda c: c[1], self.zip_col))
            self.list_col = list(map(lambda c: c[2], self.zip_col))

        elif event == 2:
            self.zip_ln = zip(self.index_ln, self.res_ln, self.list_ln_cb_2_format)
            self.zip_ln = sorted(self.zip_ln, key=lambda c: c[0])
            self.index_ln = list(map(lambda c: c[0], self.zip_ln))
            self.res_ln = list(map(lambda c: c[1], self.zip_ln))
            self.list_ln_cb_2_format = list(map(lambda c: c[2], self.zip_ln))

    def exibir_selecao_coluna(self, event):
        """
                Gerencia a seleção ou exclusão de colunas via Combobox, ou fatiamento automático.

                Atualiza as variáveis, como listas de índices (index_col), nomes selecionados (res_col),
                colunas disponíveis (list_col), e a label de selecionados na interface.

                Args:
                    event: 0 para seleção automática (botão 'Tudo');

                        1 para conclusão do fatiamento (botão ':');

                        ou evento de clique na Combobox (cb).
                """
        nome_col = None

        # Se o botão ':' foi clicado, define o ponto final (B) do intervalo
        if self.fatiando:
            self.b_col_nome = self.cb.get()  # Seleção do Ponto B do intervalo feito na Combobox pelo usuário
            self.b_col = self.df.columns.get_loc(self.b_col_nome)
            self.ativado_2pts = False
            self.fatiando = False

            # Executa a lógica que agrupa as colunas entre o Ponto A e o Ponto B
            self.fatiamento_coluna()

            return

        # Define o nome da coluna com base na origem (Tudo ou Combobox)
        if event == 0:
            nome_col = self.list_col[-1]  # Seleção via botão 'Tudo'

        elif event != 1:
            nome_col = self.cb.get()  # Seleção via Combobox

        # --- Fluxo de REMOÇÃO (Checkbutton 'Tudo' está ativado) ---
        if self.alt_cb and (event != 1):
            # Transfere o item da lista de selecionados para a lista de excluídos (_2)
            self.index_col_2_aux.append(self.list_col.index(nome_col))
            self.index_col_2.append(self.df.columns.get_loc(nome_col))
            self.res_col_2.append(nome_col)
            # Remove a coluna das listas de selecionados
            self.index_col.remove(self.df.columns.get_loc(nome_col))
            self.res_col.remove(nome_col)
            self.list_col.remove(nome_col)  # Remove da lista da Combobox

            # Garante que a lista de selecionados siga a ordem original do DataFrame
            self.ordenar_lista_index(1)

        # --- Fluxo de REMOÇÃO EM MASSA (Fatiamento com 'Tudo' ativo) ---
        elif self.alt_cb and (event == 1):

            if (not self.index_fatia_col_aux) and self.index_fatia_col:
                self.index_fatia_col_aux = self.index_fatia_col.copy()

            for i, j in enumerate(self.nome_fatia_col):  # Percorre as colunas "fatiadas"
            # Garante que os índices correspondam à posição das colunas no DataFrame original (df.columns)

                if i == 0:  # O primeiro item do intervalo já existe na lista, apenas atualizamos o seu índice real

                    del self.index_fatia_col[i]

                    self.index_fatia_col.insert(i, self.df.columns.get_loc(j))
                    continue

                else:  # Remove o marcador temporário e insere o índice absoluto e o relativo (aux)

                    del self.index_fatia_col[i]
                    del self.index_fatia_col_aux[i]

                    self.index_fatia_col_aux.insert(i, self.list_col.index(j) + 1)
                    self.index_fatia_col.insert(i, self.df.columns.get_loc(j))

            # Move o grupo fatiado para a lista de "excluídos" (_2)
            del self.index_col_2_aux[-1]
            del self.index_col_2[-1]

            self.index_col_2_aux.append(self.index_fatia_col_aux.copy())
            self.index_col_2.append(self.index_fatia_col.copy())
            del self.res_col_2[-1]

            self.res_col_2.append(self.nome_fatia_col)

            # Limpa as colunas do intervalo das listas de seleção ativa
            for i, j in enumerate(self.nome_fatia_col):

                if i == 0: continue  # A primeira coluna do intervalo permanece como 'âncora'

                self.res_col.remove(j)
                self.list_col.remove(j)
                self.index_col.remove(self.df.columns.get_loc(j))

            # Reordena novamente após a limpeza para manter a integridade dos índices
            self.ordenar_lista_index(1)

        # --- Fluxo de ADIÇÃO (Checkbutton 'Tudo' desativado) ---
        elif event == 1:  # Adiciona um intervalo (fatiamento) aos selecionados

            del self.res_col[-1]
            del self.index_col[-1]

            self.index_col.append(self.index_fatia_col.copy())
            self.res_col.append(self.nome_fatia_col)

        else:  # Adiciona uma única coluna aos selecionados
            self.index_col.append(self.list_col.index(nome_col))
            self.res_col.append(nome_col)

        # --- Atualização da Interface Visual ---
        if (event == 0) and (nome_col == self.list_col[0]):
            # Inverte para exibir na ordem correta do DataFrame caso seja "Tudo"
            self.res_col.reverse()
            res_col = self.txt_quebra_de_linha(self.res_col, 1)
            self.res_col_var.set(res_col)
            self.res_col.reverse()

        elif event != 0:
            # Prepara a string para o Label, lidando com possíveis listas aninhadas
            join_list_res_col = []

            for col in self.res_col:

                if isinstance(col, list):
                    join_list_res_col.extend(col)

                else:
                    join_list_res_col.append(col)

            res_col = self.txt_quebra_de_linha(join_list_res_col, 1)
            self.res_col_var.set(res_col)

        if (not self.alt_cb) and (event != 1):
            self.list_col.remove(self.res_col[-1])  # Atualiza as opções disponíveis na Combobox

        elif not self.alt_cb:  # Remove todos os itens do intervalo da lista de opções da Combobox

            for i in self.res_col[-1][1:]:
                self.list_col.remove(i)

        self.cb['values'] = self.list_col

    def fatiamento_coluna(self):
        """
        Coordena o processo de fatiamento (seleção de intervalo) de colunas.

        A função funciona em duas etapas:

        1. Primeiro clique no botão ':': Com o Ponto A previamente selecionado,
        prepara a interface e limita as opções da Combobox para colunas à frente da selecionada.

        2. Segundo clique (seleção do Ponto B): Calcula o intervalo completo e
        solicita a atualização da interface através da função 'exibir_selecao_coluna'.
        """

        # --- Validação Inicial: Define a lista de referência para o fatiamento ---
        if self.res_col_2:

            if isinstance(self.res_col_2[-1], list):
                res_col_2 = self.res_col_2[-1]

            else:
                res_col_2 = self.res_col_2

        else:
            res_col_2 = [""]

        # Bloqueia a execução se não houver colunas para fatiar ou se a seleção for inválida
        if ((self.alt_cb and not self.res_col_2) or
                not (self.alt_cb or self.res_col) or
                not self.list_col or
                ((not self.alt_cb) and (self.index_col[-1] == len(self.list_col))) or
                ((not self.alt_cb) and (self.index_col[-1] == [len(self.list_col)])) or
                (self.alt_cb and (self.df.columns.get_loc(self.res_col_2[-1]) ==
                                  self.df.columns.get_loc(self.list_col[-1]) + 1))):
            return

        if self.list_col:

            # --- Caso de Cancelamento/Reset do Fatiamento ---
            if self.ativado_2pts:
                self.a_col_nome = self.df.columns.__getitem__(self.a_col)
                del self.aux_cb_values[self.aux_cb_values.index(self.a_col_nome)]

                self.cb['values'] = self.aux_cb_values  # Restaura os valores originais
                self.fatiando = False
                self.a_col = None
                self.aux_cb_values.clear()
                self.ativado_2pts = False

                # Restaura o texto abaixo da Combobox na interface
                if not self.alt_cb:
                    self.alt_txt_cb.set('Selecionar coluna ativado')

                else:
                    self.alt_txt_cb.set('Desselecionar coluna ativado')

                # Reativa o botão de desfazer e o checkbox 'Tudo'
                self.des_col['state'] = NORMAL
                self.tudo_col['state'] = NORMAL

            # --- Início do Fatiamento (Seleção do Ponto A) ---
            elif self.res_col and not ((isinstance(self.res_col[-1], list)) or (isinstance(res_col_2[-1], list))):

                # Identifica o índice da coluna que servirá como início do intervalo (Ponto A)
                if not self.a_col and self.a_col != 0:

                    if self.alt_cb:
                        self.a_col = self.df.columns.get_loc(self.res_col_2[-1])

                    else:
                        self.a_col = self.df.columns.get_loc(self.res_col[-1])

                    self.fatiando = True

                # Configura a interface para aguardar a seleção do Ponto B pelo usuário
                if not self.b_col:
                    self.des_col['state'] = DISABLED  # Bloqueia botões para evitar erros de fluxo
                    self.tudo_col['state'] = DISABLED
                    self.ativado_2pts = True
                    self.aux_cb_values.extend(self.cb['values'])  # Salva estado atual das opções

                    # Filtra a Combobox para mostrar apenas colunas que vêm DEPOIS do Ponto A
                    if self.alt_cb:
                        self.aux_cb_values.insert(self.index_col_2_aux[-1], self.res_col_2[-1])
                        self.cb['values'] = self.aux_cb_values[(self.index_col_2_aux[-1] + 1):]

                    else:
                        self.aux_cb_values.insert(self.index_col[-1], self.res_col[-1])
                        self.cb['values'] = self.aux_cb_values[(self.index_col[-1] + 1):]

                    self.alt_txt_cb.set('Fatiar coluna ativado (:)')

                # --- Conclusão do Processo de Fatiamento (Ponto B já selecionado) ---
                else:

                    # Guarda todos os nomes da fatia
                    if self.alt_cb:  # Quando o Checkbox 'Tudo' ativado
                        self.nome_fatia_col = \
                            self.aux_cb_values[self.index_col_2_aux[-1]:(self.list_col.index(self.b_col_nome) + 2)]

                    else:  # Quando o Checkbox 'Tudo' desativado
                        self.nome_fatia_col = \
                            self.aux_cb_values[self.index_col[-1]:(self.list_col.index(self.b_col_nome) + 2)]

                    self.index_fatia_col.clear()
                    self.index_fatia_col_aux.clear()

                    # Gera os índices para o novo grupo fatiado
                    for i, j in enumerate(self.nome_fatia_col):

                        if not i:
                            self.index_fatia_col.append(self.aux_cb_values.index(j))

                        else:
                            self.index_fatia_col.append(self.index_fatia_col[0] + i)

                    self.aux_cb_values.clear()
                    self.fatiando = False

                    # Envia a fatia para ser processada pela função de exibição
                    self.exibir_selecao_coluna(1)

                    self.a_col, self.b_col = None, None  # Resetando variáveis

                    if not self.alt_cb:
                        self.alt_txt_cb.set('Selecionar coluna ativado')

                    else:
                        self.alt_txt_cb.set('Desselecionar coluna ativado')

                    self.des_col['state'] = NORMAL
                    self.tudo_col['state'] = NORMAL

    def desfazer_selecao_coluna(self):
        """
        Remove a última seleção feita pelo usuário (seja individual ou fatiamento)
        e devolve as colunas para a lista de opções disponíveis (Combobox).

        Lida com dois estados:

        1. Modo Adição: Remove da lista de selecionados e devolve à Combobox.

        2. Modo 'Tudo' (alt_cb): Reativa uma coluna que havia sido descartada.
        """
        # Cria uma cópia da lista de resultados para manipulação visual
        rescol = self.res_col.copy()

        # --- Caso 1: Modo de Seleção Padrão ('alt_cb' desativado) ---
        if not self.alt_cb:

            # Limpeza total dos selecionados quando o Checkbox 'Tudo' é desmarcado
            if self.undo:
                del self.index_col[-1]
                del self.res_col[-1]
                del rescol[-1]

                if not self.res_col:
                    self.list_col = list(self.df.columns)
                    self.undo = False

            elif self.res_col:
                dp_on = False

                # Verifica se existe alguma fatia (lista) dentro das seleções
                for idx, col in enumerate(self.res_col):

                    if isinstance(col, list):
                        dp_on = True
                        break

                # Se houver fatiamento, reconstrói uma lista visual (rescol)
                if dp_on:
                    rescol = []
                    del dp_on

                    for idx, col in enumerate(self.res_col):

                        # Se não for o último item, apenas 'achata' para exibição
                        if col is not self.res_col[-1]:

                            if isinstance(col, list):
                                rescol.extend(col)

                            else:
                                rescol.append(col)

                        else:

                            # Se for o último item e for uma lista, devolve todos os itens à Combobox
                            if isinstance(col, list):

                                for (i, j) in zip(self.index_col[idx], col):
                                    self.list_col.insert(i, j)

                            else:  # Senão devolve o item individual na posição correta
                                self.list_col.insert(self.index_col[idx], col)

                    # Remove o último registro das listas de seleção
                    del self.index_col[-1]
                    del self.res_col[-1]

                else:  # Devolve a última coluna individual selecionada para a Combobox
                    self.list_col.insert(self.index_col[-1], self.res_col[-1])
                    del self.index_col[-1]
                    del self.res_col[-1]
                    del rescol[-1]

        # --- Caso 2: Modo de Desseleção (Checkbox 'Tudo' ativado) ---
        elif self.index_col_2:
            # "Desfaz a exclusão": Move os itens das listas de excluídos (_2) de volta para os selecionados
            self.list_col.append(self.res_col_2[-1])
            self.index_col.append(self.index_col_2[-1])
            self.res_col.append(self.res_col_2[-1])

            # Limpa os registros da lista de excluídos e auxiliar
            del self.index_col_2_aux[-1]
            del self.index_col_2[-1]
            del self.res_col_2[-1]

            # Reordena para garantir que a coluna reativada apareça na ordem certa
            self.ordenar_lista_index(1)

            rescol = self.res_col.copy()

        # --- Atualização da Interface ---
        # Formata a string com quebras de linha para o Label da interface
        res_col = self.txt_quebra_de_linha(rescol, 1)
        self.res_col_var.set(res_col)
        # Atualiza os valores da Combobox com o retorno do(s) nome(s) da(s) coluna(s)
        self.cb['values'] = self.list_col

    def selecionar_tudo_coluna(self):
        """
        Gerencia o estado global de seleção das colunas através do Checkbutton 'Tudo'.

        Se desmarcado (0): Limpa todas as seleções atuais, uma a uma.

        Se marcado (1): Adiciona todas as colunas do DataFrame à seleção,
        preparando o sistema para o modo de 'desseleção' (alt_cb).
        """

        # --- Caso 1: Checkbutton Desmarcado (Limpeza Total) ---
        if self.checkvar1.get() == 0:
            self.alt_cb = False
            self.undo = True  # Sinaliza para o 'desfazer' que a limpeza é total
            self.list_col = []
            self.alt_txt_cb.set('Selecionar coluna ativado')

        # Esvazia as listas de selecionados chamando a função de desfazer repetidamente
        while self.res_col:
            self.desfazer_selecao_coluna()

        # --- Caso 2: Checkbutton Marcado (Seleção Total) ---
        if self.checkvar1.get() == 1:

            # Simula a seleção de cada coluna disponível até que a lista (list_col) esvazie
            while self.list_col:
                self.exibir_selecao_coluna(0)  # O argumento 0 indica seleção automática

            # Após selecionar tudo, repopula as listas para permitir a 'desseleção'
            self.list_col = list(self.df.columns)
            self.cb['values'] = self.list_col
            # Reverte as listas para que a ordem visual no Label coincida com o DataFrame
            self.res_col.reverse()
            self.index_col.reverse()

            self.index_col_2_aux, self.index_col_2, self.res_col_2 = [], [], []  # Resetando variáveis
            self.alt_cb = True  # Ativa o modo onde selecionar na Combobox remove o item
            self.alt_txt_cb.set('Desselecionar coluna ativado')

    def exibir_selecao_linha(self, event):
        """
        Gerencia a seleção ou exclusão de linhas via Combobox, ou fatiamento automático.

        Semelhante à seleção de colunas, mas utiliza mapeamento via dicionários
        para converter as strings visualizadas na Combobox em listas e índices reais do DataFrame.

        Args:
            event: 0 para 'Selecionar Tudo',

                1 para conclusão de fatiamento (botão ':'),

                ou evento de clique na Combobox (cb_2).
        """

        nome_ln = None
        nome_ln_str = None
        nome_ln_cb_2_format = None
        index_ln_2_list_format = []

        # --- Lógica do Botão ':' (Modo Fatiamento de Linhas) ---
        if self.fatiando_2:
            self.b_ln_nome = self.cb_2.get()  # Captura o Ponto B
            # Converte a string da Combobox de volta para o formato de lista padrão
            b_ln_nome_list_format = self.dict_cb_2_str_list_ln.get(self.b_ln_nome)
            # Localiza o índice real da linha no DataFrame comparando os valores
            b_ln_idx_list = [idx for idx, *row in self.df.itertuples() if row == b_ln_nome_list_format]
            self.b_ln = b_ln_idx_list[0]
            self.ativado_2pts_2 = False
            self.fatiando_2 = False

            # Aciona o cálculo do intervalo A:B
            self.fatiamento_linha()

            return

        # --- Preparação dos Nomes (Formatação Visual) ---
        if event == 0:  # Quando 'Tudo' é ativado
            # Pega o último item disponível na lista
            nome_ln_cb_2_format = self.list_ln_cb_2_format[-1]
            nome_ln = self.dict_cb_2_str_list_ln.get(nome_ln_cb_2_format)
            nome_ln_str = " | ".join([str(item) for item in nome_ln])  # Formata para exibição

        elif event != 1:
            # Pega o item selecionado na Combobox
            nome_ln_cb_2_format = self.cb_2.get()
            nome_ln = self.dict_cb_2_str_list_ln.get(nome_ln_cb_2_format)
            nome_ln_str = " | ".join([str(item) for item in nome_ln])  # Formata para exibição

        # --- Fluxo de REMOÇÃO (Modo Desseleção ativado) ---
        if self.alt_cb_2 and (event != 1):
            self.index_ln_2_aux.append(self.list_ln_cb_2_format.index(nome_ln_cb_2_format))
            # Busca o índice real da linha com 'itertuples()'
            index_ln_2_list_format.extend([idx for idx, *row in self.df.itertuples() if row == nome_ln])
            self.index_ln_2.extend(index_ln_2_list_format)
            self.res_ln_2.append(nome_ln_cb_2_format)
            # Remove da lista de selecionados
            self.index_ln.remove(index_ln_2_list_format[-1])
            self.res_ln.remove(nome_ln_str)
            self.list_ln_cb_2_format.remove(nome_ln_cb_2_format)

            self.ordenar_lista_index(2)  # Reordena linhas (event 2)

        # --- Fluxo de REMOÇÃO EM MASSA (Fatiamento no modo Desseleção) ---
        elif self.alt_cb_2 and (event == 1):

            if (not self.index_fatia_ln_aux) and self.index_fatia_ln:
                self.index_fatia_ln_aux = self.index_fatia_ln.copy()

            for i, j in enumerate(self.nome_fatia_ln):  # Percorre cada item (linha) dentro do intervalo fatiado (A:B)
                # Converte a string da combobox de volta para o formato de lista padrão
                j_list_format = self.dict_cb_2_str_list_ln.get(j)
                # Busca o índice real no DataFrame que corresponde ao conteúdo da linha
                idx_list = [idx for idx, *row in self.df.itertuples() if row == j_list_format]

                if i == 0:  # Para o 'Ponto A' da fatia

                    # Remove o marcador temporário e insere o índice real do DataFrame
                    del self.index_fatia_ln[i]

                    self.index_fatia_ln.insert(i, idx_list[0])
                    continue

                else:  # Para o restante até o 'Ponto B'

                    del self.index_fatia_ln[i]
                    del self.index_fatia_ln_aux[i]

                    # Sincroniza índice visual (posição na Combobox + 1) e índice absoluto do DataFrame
                    self.index_fatia_ln_aux.insert(i, self.list_ln_cb_2_format.index(j) + 1)
                    self.index_fatia_ln.insert(i, idx_list[0])

            # Registra o grupo fatiado na pilha de exclusão (_2)
            del self.index_ln_2_aux[-1]
            del self.index_ln_2[-1]

            self.index_ln_2_aux.append(self.index_fatia_ln_aux.copy())
            self.index_ln_2.append(self.index_fatia_ln.copy())
            del self.res_ln_2[-1]

            self.res_ln_2.append(self.nome_fatia_ln)

            # Limpa a fatia de linhas das listas de seleção
            for i, j in enumerate(self.nome_fatia_ln):

                if i == 0:
                    continue

                j_list_format = self.dict_cb_2_str_list_ln.get(j)
                j_str_format = " | ".join([str(item) for item in j_list_format])
                self.res_ln.remove(j_str_format)
                self.list_ln_cb_2_format.remove(j)
                j_list_format_idx = [idx for idx, *row in self.df.itertuples() if row == j_list_format]
                self.index_ln.remove(j_list_format_idx[-1])

            self.ordenar_lista_index(2)  # Reordena linhas (event 2)

        # --- Fluxo de ADIÇÃO (Modo Seleção Padrão) ---
        elif event == 1:  # Finaliza o fatiamento substituindo o marcador individual pelo intervalo completo
            del self.res_ln[-1]
            del self.index_ln[-1]

            self.index_ln.append(self.index_fatia_ln.copy())
            # Converte a fatia de nomes para o formato de exibição separado por '|'
            nome_fatia_ln_list_format = [self.dict_cb_2_str_list_ln.get(item) for item in self.nome_fatia_ln]
            nome_fatia_ln_str_format = [" | ".join([str(x) for x in item]) for item in nome_fatia_ln_list_format]
            self.res_ln.append(nome_fatia_ln_str_format)

        else:  # Adição da linha selecionada
            self.index_ln.append(self.list_ln_cb_2_format.index(nome_ln_cb_2_format))
            self.res_ln.append(nome_ln_str)

        # --- Atualização do Label e Combobox ---
        if (event == 0) and (nome_ln_cb_2_format == self.list_ln_cb_2_format[0]):
            # Reverte para manter a ordem correta na exibição do Label
            self.res_ln.reverse()
            # Chama a formatação especial com molduras para linhas (event 2)
            res_ln = self.txt_quebra_de_linha(self.res_ln, 2)
            self.res_ln_var.set(res_ln)
            self.res_ln.reverse()

        elif event != 0:  # Achata possíveis listas de fatiamento para exibição no Label
            extend_list_res_ln = []

            for ln in self.res_ln:

                if isinstance(ln, list):
                    extend_list_res_ln.extend(ln)

                else:
                    extend_list_res_ln.append(ln)

            res_ln = self.txt_quebra_de_linha(extend_list_res_ln, 2)
            self.res_ln_var.set(res_ln)

        # Remove as linhas selecionadas da Combobox para evitar repetição
        if (not self.alt_cb_2) and (event != 1):
            self.list_ln_cb_2_format.remove(self.dict_convert_str_to_cb_2.get(self.res_ln[-1]))

        elif not self.alt_cb_2:
            res_ln_cb_2_format = [[self.dict_convert_str_to_cb_2.get(item) for item in itens] for itens in self.res_ln]

            for item in res_ln_cb_2_format[-1]:

                if item != res_ln_cb_2_format[-1][0]:
                    self.list_ln_cb_2_format.remove(item)

        self.cb_2['values'] = self.list_ln_cb_2_format  # Atualiza o combobox

    def fatiamento_linha(self):
        """
        Coordena o processo de fatiamento (seleção de intervalo) de linhas.

        Funciona em duas etapas:

        1. Primeiro clique (Ponto A): Identifica a linha inicial e limita a Combobox para opções posteriores.

        2. Segundo clique (Ponto B): Calcula o intervalo de linhas entre A e B e solicita a atualização da interface.
        """
        # Converte a lista visual da Combobox para o formato de lista padrão
        list_ln_list_format = [self.dict_cb_2_str_list_ln.get(item) for item in self.list_ln_cb_2_format]

        # --- Validação Inicial: Define a lista de referência para o fatiamento ---
        if self.res_ln_2 and len(self.res_ln_2) > 1:
            # Caso existam múltiplas exclusões, foca na última para definir o novo Ponto A
            res_ln_2 = [self.res_ln_2[-1]]

            if isinstance(self.res_ln_2[-1], list):
                res_ln_2_list_format = [self.dict_cb_2_str_list_ln.get(item) for item in res_ln_2[-1]]

            else:
                res_ln_2_list_format = [self.dict_cb_2_str_list_ln.get(item) for item in res_ln_2]

        elif self.res_ln_2:
            # Caso exista apenas uma exclusão anterior, prepara o seu formato para comparação
            res_ln_2 = self.res_ln_2.copy()
            res_ln_2_list_format = []

            for itens in self.res_ln_2:

                if isinstance(itens, list):
                    res_ln_2_list_format = [self.dict_cb_2_str_list_ln.get(item) for item in itens]

                else:
                    res_ln_2_list_format = [self.dict_cb_2_str_list_ln.get(item) for item in res_ln_2]

        else:
            # Caso não haja exclusões, inicializa variáveis vazias para evitar erros de referência
            res_ln_2 = [""]
            res_ln_2_list_format = []

        # Bloqueio de segurança: cancela o fatiamento se apresentar alguma das condições abaixo:
        # Fatiar exclusão se nada foi excluído ou seleção se nada foi selecionado, se a Combobox estiver vazia,
        # se o índice estiver fora de alcance na seleção ou em fatias, ou fatiar além da última linha disponível
        if ((self.alt_cb_2 and not self.res_ln_2) or
                not (self.alt_cb_2 or self.res_ln) or
                not self.list_ln_cb_2_format or
                ((not self.alt_cb_2) and (self.index_ln[-1] == len(self.list_ln_cb_2_format))) or
                ((not self.alt_cb_2) and (self.index_ln[-1] == [len(self.list_ln_cb_2_format)])) or
                (self.alt_cb_2 and (self.list_ln.index(res_ln_2_list_format[-1]) ==
                                    self.list_ln.index(list_ln_list_format[-1]) + 1))):
            return

        if self.list_ln_cb_2_format:

            # --- Caso de Cancelamento/Reset Manual do Fatiamento (Segunda chamada sem Ponto B) ---
            if self.ativado_2pts_2:

                if self.alt_cb_2:
                    a_ln_nome_list_format = [row for idx, *row in self.df.itertuples() if idx == self.a_ln]

                else:
                    a_ln_nome_list_format = [row for idx, *row in self.df.itertuples() if idx == self.a_ln[-1]]

                self.a_ln_nome = [" ".join([str(x) for x in item]) for item in a_ln_nome_list_format]
                del self.aux_cb_2_values[self.aux_cb_2_values.index(self.a_ln_nome[-1])]

                self.cb_2['values'] = self.aux_cb_2_values  # Restaura valores originais
                self.fatiando_2 = False
                self.a_ln = None
                self.aux_cb_2_values.clear()
                self.ativado_2pts_2 = False

                if not self.alt_cb_2:
                    self.alt_txt_cb_2.set('Selecionar linha ativado')

                else:
                    self.alt_txt_cb_2.set('Desselecionar linha ativado')

                self.des_ln['state'] = NORMAL
                self.tudo_ln['state'] = NORMAL

            # --- Início do Fatiamento (Seleção do Ponto A) ---
            elif self.res_ln and not ((isinstance(self.res_ln[-1], list)) or (isinstance(res_ln_2[-1], list))):

                # Identifica o índice real do 'Ponto A'
                if not self.a_ln and self.a_ln != 0:

                    if self.alt_cb_2:
                        a_ln_list = [idx for idx, *row in self.df.itertuples()
                                     if row == self.dict_cb_2_str_list_ln.get(self.res_ln_2[-1])]

                        self.a_ln = a_ln_list[0]

                    else:
                        self.a_ln = [idx for idx, *row in self.df.itertuples()
                                     if row == self.dict_str_list_ln.get(self.res_ln[-1])]

                    self.fatiando_2 = True

                # Prepara a Combobox para a seleção do Ponto B
                if not self.b_ln:
                    self.des_ln['state'] = DISABLED
                    self.tudo_ln['state'] = DISABLED
                    self.ativado_2pts_2 = True
                    self.aux_cb_2_values.extend(self.cb_2['values'])

                    # Filtra a lista para mostrar apenas as linhas posteriores ao Ponto A
                    if self.alt_cb_2:
                        self.aux_cb_2_values.insert(self.index_ln_2_aux[-1], self.res_ln_2[-1])
                        self.cb_2['values'] = self.aux_cb_2_values[(self.index_ln_2_aux[-1] + 1):]

                    else:
                        self.aux_cb_2_values.insert(
                            self.index_ln[-1], self.dict_convert_str_to_cb_2.get(self.res_ln[-1]))

                        self.cb_2['values'] = self.aux_cb_2_values[(self.index_ln[-1] + 1):]

                    self.alt_txt_cb_2.set('Fatiar linha ativado (:)')

                # --- Conclusão do Fatiamento (Ponto B selecionado) ---
                else:

                    # Captura todos os nomes de linha contidos no intervalo A:B
                    if self.alt_cb_2:
                        self.nome_fatia_ln = \
                            self.aux_cb_2_values[self.index_ln_2_aux[-1]:(self.list_ln_cb_2_format
                                                                          .index(self.b_ln_nome) + 2)]

                    else:
                        self.nome_fatia_ln = \
                            self.aux_cb_2_values[self.index_ln[-1]:(self.list_ln_cb_2_format
                                                                    .index(self.b_ln_nome) + 2)]

                    self.index_fatia_ln.clear()
                    self.index_fatia_ln_aux.clear()

                    # Mapeia os índices da fatia baseando-se na posição do Ponto A na lista auxiliar
                    for i, j in enumerate(self.nome_fatia_ln):

                        if not i:
                            self.index_fatia_ln.append(self.aux_cb_2_values.index(j))

                        else:
                            self.index_fatia_ln.append(self.index_fatia_ln[0] + i)

                    self.aux_cb_2_values.clear()
                    self.fatiando_2 = False

                    self.exibir_selecao_linha(1)  # Envia para processamento visual

                    self.a_ln = None
                    self.b_ln = None

                    if not self.alt_cb_2:
                        self.alt_txt_cb_2.set('Selecionar linha ativado')

                    else:
                        self.alt_txt_cb_2.set('Desselecionar linha ativado')

                    self.des_ln['state'] = NORMAL
                    self.tudo_ln['state'] = NORMAL

    def desfazer_selecao_linha(self):
        """
        Remove a última ação de seleção ou fatiamento de linhas.

        Restaura as linhas para a Combobox e atualiza o Label visual,
        gerenciando a conversão entre strings formatadas e objetos da lista.
        """
        # Cria uma cópia para manipular a exibição sem alterar a lista principal
        resln = self.res_ln.copy()

        # --- Caso 1: Modo de Seleção Padrão (Checkbox 'Tudo' desativado) ---
        if not self.alt_cb_2:

            if self.undo_2:  # Limpeza total quando o 'Tudo' é desmarcado
                del self.index_ln[-1]
                del self.res_ln[-1]
                del resln[-1]

                if not self.res_ln:  # Restaura a lista original formatada para a Combobox
                    self.list_ln_cb_2_format = [" ".join([str(x) for x in item]) for item in self.list_ln]
                    self.undo_2 = False

            elif self.res_ln:
                dp_on = False

                for idx, ln in enumerate(self.res_ln):  # Verifica se o último registro é uma fatia (lista)

                    if isinstance(ln, list):
                        dp_on = True
                        break

                if dp_on:
                    resln = []
                    del dp_on

                    for idx, ln in enumerate(self.res_ln):

                        if ln is not self.res_ln[-1]:  # Se não for o último, apenas achata para a exibição (resln)

                            if isinstance(ln, list):
                                resln.extend(ln)

                            else:
                                resln.append(ln)

                        else:  # Se for o último e for lista, converte cada item e devolve à Combobox

                            if isinstance(ln, list):
                                res_ln_cb_2_format = [self.dict_convert_str_to_cb_2.get(item) for item in ln]

                                for (i, j) in zip(self.index_ln[idx], res_ln_cb_2_format):
                                    self.list_ln_cb_2_format.insert(i, j)

                            else:  # Devolve o item individual na posição original
                                self.list_ln_cb_2_format.insert(
                                    self.index_ln[idx], self.dict_convert_str_to_cb_2.get(ln))

                    del self.index_ln[-1]
                    del self.res_ln[-1]

                else:  # Caso simples: Devolve a última linha individual para a Combobox
                    res_ln_cb_2_format = self.dict_convert_str_to_cb_2.get(self.res_ln[-1])
                    self.list_ln_cb_2_format.insert(self.index_ln[-1], res_ln_cb_2_format)
                    del self.index_ln[-1]
                    del self.res_ln[-1]
                    del resln[-1]

        # --- Caso 2: Modo de Desseleção (Checkbox 'Tudo' ativado) ---
        elif self.index_ln_2:
            # Reativa uma linha que havia sido descartada, movendo-a de volta para 'res_ln'
            self.list_ln_cb_2_format.append(self.res_ln_2[-1])
            self.index_ln.append(self.index_ln_2[-1])
            res_ln_2_list_format = []
            itens = self.res_ln_2[-1]

            # Formata o item (ou lista de itens) para voltar a ser exibido com "|"
            if isinstance(itens, list):
                res_ln_2_list_format.append([self.dict_cb_2_str_list_ln.get(item) for item in itens])
                res_ln_2_str_format_list = ([[" | ".join([str(x) for x in item]) for item in itens]
                                             for itens in res_ln_2_list_format])

                res_ln_2_str_format = res_ln_2_str_format_list[-1]

            else:
                res_ln_2_list_format = self.dict_cb_2_str_list_ln.get(itens)
                res_ln_2_str_format = " | ".join(str(item) for item in res_ln_2_list_format)

            self.res_ln.append(res_ln_2_str_format)
            # Limpa os registros das listas de exclusão (_2)
            del self.index_ln_2_aux[-1]
            del self.index_ln_2[-1]
            del self.res_ln_2[-1]

            self.ordenar_lista_index(2)  # Reordena para manter a sequência correta das linhas

            resln = self.res_ln.copy()

        # --- Atualização da Interface ---
        # Aplica a quebra de linha e as molduras (underline) específicas para linhas
        res_ln = self.txt_quebra_de_linha(resln, 2)
        self.res_ln_var.set(res_ln)
        self.cb_2['values'] = self.list_ln_cb_2_format

    def selecionar_tudo_linha(self):
        """
        Gerencia o estado global de seleção das linhas através do Checkbutton 'Tudo'.

        Se desmarcado (0): Inicia uma limpeza total da pilha de seleções.

        Se marcado (1): Adiciona todas as linhas do DataFrame à seleção de uma vez,
        preparando o sistema para o modo de 'desseleção' (alt_cb_2).
        """

        # --- Caso 1: Checkbutton Desmarcado (Limpeza Total) ---
        if self.checkvar2.get() == 0:
            self.alt_cb_2 = False
            self.undo_2 = True  # Sinaliza para o 'desfazer' que deve resetar o estado
            self.list_ln_cb_2_format = []
            self.alt_txt_cb_2.set('Selecionar linha ativado')

        # Esvazia a pilha de linhas selecionadas chamando o 'desfazer' repetidamente
        while self.res_ln:
            self.desfazer_selecao_linha()

        # --- Caso 2: Checkbutton Marcado (Seleção Total) ---
        if self.checkvar2.get() == 1:

            # Automatiza a seleção de linhas disponíveis
            while self.list_ln_cb_2_format:
                self.exibir_selecao_linha(0)  # O argumento 0 indica seleção automática

            # Reconstrói a lista visual da Combobox formatando cada linha
            self.list_ln_cb_2_format = [" ".join([str(x) for x in item]) for item in self.list_ln]
            self.cb_2['values'] = self.list_ln_cb_2_format
            # Reverte a ordem para que a exibição no Label siga a ordem do DataFrame
            self.res_ln.reverse()
            self.index_ln.reverse()
            self.index_ln_2_aux, self.index_ln_2, self.res_ln_2 = [], [], []  # Resetando variáveis
            self.alt_cb_2 = True  # Ativa o modo onde o clique na Combobox remove a linha
            self.alt_txt_cb_2.set('Desselecionar linha ativado')

    def tratar_dados_selecionados(self):
        """
        Executa o pipeline de pré-processamento, limpeza e normalização dos dados.

        Realiza a extração de índices e colunas selecionados na interface,
        seguido pela aplicação de filtros estatísticos para remoção de valores
        ausentes (NaN) e tratamento de variáveis não numéricas. Inclui a conversão
        de variáveis categóricas em indicadores binários (Dummies) e a validação
        dos graus de liberdade necessários para a Regressão Linear Múltipla (RLM).

        Processos principais:

        1. Mapeamento de índices reais do DataFrame original.

        2. Aplicação de critérios de exclusão por vacância (NaN) baseados no tamanho da amostra.

        3. Conversão de tipos e tratamento de colunas não numéricas.

        4. Geração de variáveis 'Dummy' para dados categóricos representativos.

        5. Verificação de consistência estatística (mínimo de 8 linhas e 3 colunas).
        """

        self.df_select_bruto_aux_2, self.df_select_bruto_aux_3, self.df_todos_dados, self.df_validos \
            = pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()  # Resetando variáveis.

        def achatar_lista(lista):
            """Função recursiva para transformar listas aninhadas (fatiamentos) em listas simples."""
            nova_lista = []

            for obj in lista:

                if isinstance(obj, list):

                    nova_lista.extend(achatar_lista(obj))

                else:

                    nova_lista.append(obj)

            return nova_lista

        # Extrai os nomes das colunas selecionadas, removendo o aninhamento dos fatiamentos
        cols_achatadas = achatar_lista(self.res_col)

        indexes_finais = []

        # --- Lógica de Mapeamento de Índices (Linhas) ---
        # Caso 1: Seleção Total (Quando 'Tudo' ativado)
        if self.checkvar2.get() == 1:
            click_positions = achatar_lista(self.index_ln)
            # Converte as posições relativas em índices reais do DataFrame original
            indexes_finais = [self.df.index[i] for i in click_positions]

        # Caso 2: Seleção Manual
        else:
            # Cria uma lista temporária com os índices originais
            indexes_temps = list(self.df.index)

            for item in self.index_ln:

                if isinstance(item, list):
                    # Inverte para a remoção (pop), mantendo a integridade das posições
                    list_sorted = sorted(item, reverse=True)
                    ids_da_lista = []

                    for pos in list_sorted:

                        if indexes_temps:  # Verificação de segurança
                            ids_reais = indexes_temps.pop(pos)  # Remove da lista temporária invertida (desempilhando)

                            ids_da_lista.append(ids_reais)

                    # Inverte novamente para manter a ordem cronológica da seleção do usuário
                    indexes_finais.extend(reversed(ids_da_lista))

                else:

                    if indexes_temps:
                        ids_reais = indexes_temps.pop(item)  # Remove o item da lista temporária
                        indexes_finais.append(ids_reais)

        # --- Criação do DataFrame Temporário ---
        # Filtra o DataFrame original com base nos índices e colunas processados
        self.df_select_bruto = self.df.loc[indexes_finais, cols_achatadas].copy()

        self.df_select_bruto_aux = self.df_select_bruto.copy()

        # --- Limpeza de Dados Faltantes (NaN) ---
        # Percentual de valores nulos por coluna
        nan_percent = (self.df_select_bruto_aux.isnull().sum() / len(self.df_select_bruto_aux)) * 100

        # Definindo critérios para remoção de colunas com base no tamanho do dataframe:
        # — Planilhas pequenas (>20 linhas): Remove colunas com >20% de vazios
        # — Planilhas grandes (>500 linhas): Remove colunas com >10% de vazios
        cols_para_remover = [col for col in self.df_select_bruto_aux.columns if
                             (nan_percent[col] > 30) or
                             (len(self.df_select_bruto_aux) > 20 and nan_percent[col] > 20) or
                             (len(self.df_select_bruto_aux) > 500 and nan_percent[col] > 10)]

        # Remove as colunas que não passaram no critério acima
        self.df_select_bruto_aux = self.df_select_bruto_aux.drop(columns=cols_para_remover)

        # Define quais colunas restantes precisam de uma limpeza de linhas (dropna)
        self.cols_para_verificar = [col for col in self.df_select_bruto.columns if
                                    col not in cols_para_remover]

        # Remove as linhas que ainda possuem valores nulos nas colunas selecionadas
        self.df_select_bruto_aux = self.df_select_bruto_aux.dropna(subset=self.cols_para_verificar)

        # --- Tratamento de Colunas "Não Numéricas" ---
        not_num_cols = self.df_select_bruto_aux.select_dtypes(exclude=['number']).columns

        if not not_num_cols.empty:

            for col in not_num_cols:
                # Tenta converter a coluna para número (ex: "10.5" -> 10.5)
                # 'coerce' transforma textos puramente literais em NaN temporariamente
                # noinspection PyTypeChecker
                numeric_col: Series = pd.to_numeric(self.df_select_bruto_aux[col], errors='coerce')
                not_num_bool = numeric_col.isnull()
                not_num_percent = (not_num_bool.sum() / len(self.df_select_bruto_aux)) * 100

                # Se a coluna tiver muito "lixo" (texto onde deveria ser número), ela é descartada
                if ((numeric_col.sum() > 100 and not_num_percent > 5 and not_num_percent != 100) or
                        (not_num_percent > 30 and not_num_percent != 100)):

                    self.df_select_bruto_aux = self.df_select_bruto_aux.drop(columns=[col])

                # Caso a coluna seja puramente texto (Categorias como 'Sim/Não' ou 'Cidades')
                elif not_num_percent == 100:
                    qtde_repeat = self.df_select_bruto_aux[col].value_counts()
                    percent_repeat = (len(qtde_repeat[qtde_repeat > 1]) / self.df_select_bruto_aux[col].nunique()) * 100

                    # Verifica se a repetição de dados justifica criar variáveis 'Dummy'
                    # Essencial para transformar 'Categorias' em 'Binários' para a Regressão
                    if ((numeric_col.sum() < 50 <= percent_repeat) or
                            (numeric_col.sum() >= 50 and percent_repeat >= 90)):

                        self.df_select_bruto_aux_2 = (
                            pd.get_dummies(self.df_select_bruto_aux, columns=[col], drop_first=True, dtype=int))

                    # Remove a coluna original após o tratamento (ou descarte)
                    self.df_select_bruto_aux = self.df_select_bruto_aux.drop(columns=[col])

                else:
                    # Se a coluna tiver pouco "lixo", remove apenas as linhas problemáticas
                    self.df_select_bruto_aux = self.df_select_bruto_aux[~not_num_bool]

        # --- Unificação e Sincronização de Dados ---
        cols_em_comum = list(set(self.df_select_bruto.columns) & set(self.df_select_bruto_aux.columns))
        cols_extras = []

        # Se foram geradas colunas 'Dummy'...
        if not self.df_select_bruto_aux_2.empty:
            cols_df_select_bruto_aux = set(self.df_select_bruto_aux.columns)
            cols_df_select_bruto_aux_2 = set(self.df_select_bruto_aux_2.columns)
            # Identifica apenas as colunas extras criadas pelo 'get_dummies'
            cols_extras = list(cols_df_select_bruto_aux_2.difference(cols_df_select_bruto_aux))
            self.df_select_bruto_aux_3 = self.df_select_bruto_aux_2[cols_extras]

        else:
            cols_extras = []

        # Cria uma máscara para identificar quais linhas 'sobreviveram' à limpeza (NaNs e textos)
        is_compatible_mask = self.df_select_bruto.index.isin(self.df_select_bruto_aux.index)
        self.df_select_bruto['is_compatible'] = is_compatible_mask
        col_order = list(self.df_select_bruto.columns)[:-1]

        if cols_extras:
            col_order.extend(cols_extras)

        # --- Construção do Dataframe Final (df_validos) ---
        if not self.df_select_bruto_aux_3.empty:

            # Filtra apenas linhas compatíveis e colunas que não foram descartadas
            df_left: pd.DataFrame \
                = self.df_select_bruto.loc[self.df_select_bruto['is_compatible'], cols_em_comum].copy()

            df_right: pd.DataFrame = self.df_select_bruto_aux_3.copy()

            # Backup do índice para garantir o alinhamento após o 'reset_index' (necessário para "concat axis=1")
            df_left['_index_backup_'] = df_left.index
            df_left_reset = df_left.reset_index(drop=True)
            df_right_reset = df_right.reset_index(drop=True)

            # Concatena as colunas originais com as colunas 'Dummy'
            self.df_validos = pd.concat([df_left_reset, df_right_reset], axis=1)
            self.df_validos.index = self.df_validos['_index_backup_']

            self.df_validos.drop(columns=['_index_backup_'], inplace=True)

            # Cria um DataFrame completo (incluindo linhas incompatíveis) para referência
            self.df_todos_dados = pd.concat([self.df_select_bruto, self.df_select_bruto_aux_3], axis=1)

        else:
            self.df_todos_dados = self.df_select_bruto.copy()
            self.df_validos = self.df_select_bruto.loc[self.df_select_bruto['is_compatible'], cols_em_comum].copy()

        # --- Limpeza Final e Formatação ---
        # Remove duplicatas de colunas e organiza a ordem conforme a seleção original
        self.df_todos_dados = self.df_todos_dados.loc[:, ~self.df_todos_dados.columns.duplicated()]
        cols_unicas_todos = list(dict.fromkeys(col_order + ['is_compatible']))
        self.df_todos_dados = self.df_todos_dados.reindex(columns=cols_unicas_todos)
        self.df_validos = self.df_validos.loc[:, ~self.df_validos.columns.duplicated()]
        self.df_validos = self.df_validos[self.df_validos.index.notna()]
        self.df_validos = self.df_validos.dropna(how='all')

        # Garante que o índice seja inteiro (evita problemas com float no Excel/Pandas)
        if self.df_validos.index.dtype == 'float64':
            self.df_validos.index = self.df_validos.index.astype(int)

        if 'is_compatible' in self.df_validos.columns:
            self.df_validos.drop(columns=['is_compatible'], inplace=True)

        # Reordena colunas e garante que todos os dados finais sejam numéricos
        cols_finais = [col for col in col_order if col in self.df_validos.columns]
        cols_finais = list(dict.fromkeys(cols_finais))
        self.df_validos = self.df_validos.reindex(columns=cols_finais)

        for col in self.df_validos.columns:
            self.df_validos[col] = pd.to_numeric(self.df_validos[col], errors='coerce')

        self.df_validos.dropna(inplace=True)

        # --- Validação Estatística Final (Graus de Liberdade) ---
        qtd_ln_finais = len(self.df_validos)
        qtd_col_finais = len(self.df_validos.columns)

        # Critério: Mínimo 8 linhas, 3 colunas e nº Linhas > nº Colunas
        if not (7 < qtd_ln_finais > qtd_col_finais > 2):

            messagebox.showerror(
                "Dados Insuficientes Pós-Tratamento",
                f"Após remover os textos e vazios, restaram apenas:\n"
                f"- {qtd_ln_finais} linhas válidas\n"
                f"- {qtd_col_finais} colunas válidas\n\n"
                "O cálculo de RLM requer pelo menos 8 linhas válidas "
                "e 3 colunas válidas (para normalidade) "
                "e o nº de Linhas > nº de Colunas (graus de liberdade)."
            )

            self.tratar_dados_clicado = False
            self.atualizar_estado_botoes()

            return

        # --- Interface de Visualização e Validação Final ---
        # Indica sucesso na limpeza para habilitar os próximos passos (Cálculo/Gráficos)
        self.tratar_dados_clicado = True
        self.atualizar_estado_botoes()

        def df_tabela():
            """
            Constrói uma tabela visual (Grid de Labels) dentro de um Canvas com scroll.

            Permite alternar entre a visão de todos os dados e apenas os dados válidos.
            """
            # Configuração do painel superior (Filtros e Botão Salvar)
            radio_frame = Frame(self.janela_secundaria)
            radio_frame.grid(row=0, column=1, sticky='ew', padx=10, pady=(5, 15))
            radio_frame.grid_columnconfigure(0, weight=1)
            radio_frame.grid_columnconfigure(1, weight=1)
            radio_frame.grid_columnconfigure(2, weight=1)

            radio_var = StringVar(value="todos")

            def atualizar_tabela(filtro):
                """
                Renderiza as células da tabela.

                Aplica a cor VERDE para dados preservados e VERMELHO para dados descartados.
                """

                # Limpa a tabela anterior antes de redesenhar
                for widget in frame_tabela.winfo_children():
                    widget.destroy()

                # Define qual conjunto de dados será exibido
                if filtro == "verde":
                    df_tabela_color = self.df_validos.copy()
                    list_df_col_nomes = list(df_tabela_color.columns)

                else:
                    df_tabela_color = self.df_todos_dados.copy()
                    # Oculta a coluna 'is_compatible' da visualização
                    list_df_col_nomes = [col for col in df_tabela_color.columns if col != 'is_compatible']

                # Cabeçalho da Tabela
                for j, col in enumerate(list_df_col_nomes):
                    label_header = Label(frame_tabela, text=str(col), font=('Arial', 10, 'bold'))
                    label_header.grid(column=j + 1, row=0, padx=5, pady=5, sticky='nsew')
                    label_header.bind("<MouseWheel>", on_mouse_wheel)

                # Coluna de Índices
                index_label_header = Label(frame_tabela, text=' ', font=('Arial', 10, 'bold'))
                index_label_header.grid(column=0, row=0)
                index_label_header.bind("<MouseWheel>", on_mouse_wheel)

                # Preenchimento das Células com Lógica de Cores (Verde/Vermelho)
                for i, (index, row) in enumerate(df_tabela_color.iterrows()):
                    index_label = Label(frame_tabela, text=str(index), font=('Arial', 10, 'bold'))
                    index_label.grid(column=0, row=i + 1, padx=5, pady=5, sticky='nsew')
                    index_label.bind("<MouseWheel>", on_mouse_wheel)

                    for j, col_name in enumerate(list_df_col_nomes):
                        valor_pos_cel = str(row[col_name])

                        if filtro == "verde":
                            cor_fonte = "green"

                        else:
                            # Identifica se o dado sobreviveu ao tratamento
                            is_compatible_row = row['is_compatible']

                            if col_name in cols_em_comum:
                                # Verde se a linha for compatível e a célula não for vazia, senão vermelho
                                is_valid_entry = is_compatible_row and not pd.isna(row[col_name])
                                cor_fonte = "green" if is_valid_entry else "red"

                            elif col_name in cols_extras:
                                # Colunas Dummy (extras) seguem a compatibilidade da linha
                                is_valid_entry = is_compatible_row
                                cor_fonte = "green" if is_valid_entry else "red"

                            else:
                                cor_fonte = "red"  # Colunas descartadas estatísticamente

                        label_data = Label(frame_tabela, text=valor_pos_cel, fg=cor_fonte)
                        label_data.grid(column=j + 1, row=i + 1, padx=5, pady=5, sticky='nsew')
                        label_data.bind("<MouseWheel>", on_mouse_wheel)

                # Atualiza a área de rolagem do Canvas
                frame_tabela.update_idletasks()

                canvas.configure(scrollregion=canvas.bbox("all"))

            # --- Botões de Controle ---
            radio_frame_interno = Frame(radio_frame)
            radio_frame_interno.grid(row=0, column=1)

            Radiobutton(radio_frame_interno, text="Mostrar todos", variable=radio_var, value="todos",
                        command=lambda: atualizar_tabela(radio_var.get())).pack(side='left', padx=10)

            Radiobutton(radio_frame_interno, text="Mostrar dados válidos", variable=radio_var, value="verde",
                        command=lambda: atualizar_tabela(radio_var.get())).pack(side='left', padx=10)

            def confirmar_e_salvar_btn():
                """Exporta o DataFrame limpo (self.df_validos) para um novo arquivo Excel."""
                nome_arquivo_excel = self.f_n1
                diretorio_atual = os.getcwd()

                # Lógica para evitar nomes de arquivos infinitamente repetidos
                if " Tratado.xlsx" in nome_arquivo_excel:
                    nome_arquivo_excel = nome_arquivo_excel.replace(" Tratado.xlsx", ".xlsx")

                nome, extensao = os.path.splitext(nome_arquivo_excel)
                contador = 1
                nome_tratado = f"{nome} Tratado{extensao}"

                while os.path.exists(nome_tratado):
                    nome_tratado = f"{nome} Tratado({contador}){extensao}"
                    contador += 1

                # Salva o arquivo e reseta o estado da janela
                self.df_validos.to_excel(nome_tratado, index=False)
                self.janela_secundaria.destroy()

                self.janela_secundaria = None  # Resetando variável

                messagebox.showinfo(
                    "Sucesso!", f"Arquivo salvo como '{nome_tratado}' em \n{diretorio_atual}")

            self.save_btn = Button(radio_frame_interno, text="Confirmar e Salvar", command=confirmar_e_salvar_btn,
                                   font=('Concert One', 9, "bold"), fg="dark green", state='normal')

            self.save_btn.pack(side='left', padx=50)

            # Se os dados tratados forem idênticos aos brutos, o botão de salvar pode ser desabilitado
            cols_para_comparar = list(set(self.df_validos.columns) & set(self.df_todos_dados.columns))
            df_todos_comparar = self.df_todos_dados[cols_para_comparar].drop(columns=['is_compatible'], errors='ignore')

            if self.df_validos.sort_index(axis=1).equals(df_todos_comparar.sort_index(axis=1)):
                self.save_btn.configure(state='disabled')

            else:
                self.save_btn.configure(state='normal')

            # --- Estrutura de Rolagem (Canvas + Scrollbars) ---
            container = Frame(self.janela_secundaria)
            container.grid(row=1, column=1, sticky='nsew')
            canvas = Canvas(container)
            v_scrollbar = Scrollbar(container, orient="vertical", command=canvas.yview)
            h_scrollbar = Scrollbar(container, orient="horizontal", command=canvas.xview)
            v_scrollbar.pack(side="right", fill="y")
            h_scrollbar.pack(side="bottom", fill="x")
            canvas.pack(side="left", fill="both", expand=True)
            canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
            frame_tabela = Frame(canvas)
            canvas_window = canvas.create_window((0, 0), window=frame_tabela, anchor="nw")

            # Função de rolagem do mouse
            def on_mouse_wheel(event):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

            canvas.bind("<MouseWheel>", on_mouse_wheel)
            frame_tabela.bind("<MouseWheel>", on_mouse_wheel)

            def centralizar_frame(event):
                """Mantém a tabela centralizada se ela for menor que a janela."""
                canvas_largura = event.width
                frame_largura = frame_tabela.winfo_reqwidth()

                if frame_largura < canvas_largura:
                    canvas.coords(canvas_window, (canvas_largura - frame_largura) / 2, 0)

                else:
                    canvas.coords(canvas_window, 0, 0)

            canvas.bind('<Configure>', centralizar_frame)

            atualizar_tabela("todos")  # Inicia a tabela exibindo todos os dados

        # Abre a janela e executa a construção da tabela
        self.abrir_janela_dependente(title="Tabela de Dados Com e Sem Filtro")

        df_tabela()

    def abrir_janela_dependente(self, title, geometry="700x500", minsize=(400, 300)):
        """
        Cria e configura uma nova janela secundária (Toplevel).

        Garante que apenas uma janela dependente exista por vez, destruindo
        qualquer versão anterior antes de abrir a nova para evitar sobrecarga.
        """

        # Se já existe uma janela secundária aberta, fecha-a
        if hasattr(self, 'janela_secundaria') and self.janela_secundaria and self.janela_secundaria.winfo_exists():
            self.janela_secundaria.destroy()

        # Cria a nova janela
        self.janela_secundaria = Toplevel(self)
        self.janela_secundaria.title(title)
        self.janela_secundaria.geometry(geometry)
        self.janela_secundaria.minsize(*minsize)
        self.janela_secundaria.resizable(True, True)
        # Configura o sistema de grid para que o conteúdo se expanda corretamente
        self.janela_secundaria.grid_rowconfigure(1, weight=1)
        self.janela_secundaria.grid_columnconfigure(0, weight=1)
        self.janela_secundaria.grid_columnconfigure(1, weight=1)

    def atualizar_estado_botoes(self, *_args):
        """
        Verifica se os requisitos mínimos de dados (Graus de Liberdade)
        foram atendidos para habilitar os botões de tratamento e cálculo.

        Requisito: Linhas > Colunas e mínimos de 8 linhas e 3 colunas.
        """
        count = qtd_col_total = qtd_ln_total = 0

        # --- Contagem de Colunas Selecionadas ---
        for item in self.res_col:

            if isinstance(item, list):  # Se for uma lista (fatiamento), soma o tamanho da lista
                qtd_col_total += len(item)

            else:
                qtd_col_total += 1

        # --- Contagem de Linhas Selecionadas ---
        for item in self.res_ln:

            if isinstance(item, list):
                qtd_ln_total += len(item)

            else:
                qtd_ln_total += 1

        # --- Varredura da Interface para Habilitar/Desabilitar Botões ---
        for widget in self.scrollable_frame.winfo_children():

            # Localiza os botões específicos pelo texto
            if isinstance(widget, Button) and (widget.cget('text') == "Tratar Dados Selecionados" or
                                               widget.cget('text') == "Calcular RLM"):

                count += 1

                # Aplica a regra matemática (Graus de Liberdade): Linhas > 7, Linhas > Colunas e Colunas > 2
                if 7 < qtd_ln_total > qtd_col_total > 2:
                    widget.config(state='normal')

                else:
                    widget.config(state='disabled')

                # Regra para o botão de Cálculo
                if widget.cget('text') == "Calcular RLM":

                    if self.tratar_dados_clicado:  # Só habilita se os dados já foram tratados e validados com sucesso
                        widget.config(state='normal')

                    else:
                        widget.config(state='disabled')

            if count == 2:  # Interrompe a busca após encontrar os dois botões
                break

    def calcular_rlm(self):
        """
        Inicia o módulo de Regressão Linear Múltipla.

        Configura a interface de cálculo, define funções para exportação de resultados
        e gera gráficos de diagnóstico (Real vs. Previsto e Resíduos) para validar
        a precisão do modelo estatístico.
        """

        # Abre a janela de cálculos com foco exclusivo
        self.abrir_janela_dependente(title="Cálculo de Regressão Linear Múltipla (RLM)",
                                     geometry="1100x750", minsize=(800, 600))

        # Configurações de grab e foco para garantir que a janela secundária fique no topo
        if self.janela_secundaria:
            self.janela_secundaria.transient(self.master)
            self.janela_secundaria.update_idletasks()
            self.janela_secundaria.wait_visibility()
            self.janela_secundaria.grab_set()
            self.janela_secundaria.focus_set()
            self.janela_secundaria.resizable(True, True)
            self.janela_secundaria.bind("<Map>", self._reaplicar_grab)

        self.rlm_txt_summary = self.summary_var = None  # Resetando variáveis
        self.rlm_btn_salvar = self.rlm_frame_graficos = self.rlm_combo_y = None  # Resetando variáveis
        list_colunas_validas = list(self.df_validos.columns)

        def salvar_imagem_tabela():
            """
            Converte o sumário estatístico (Tabela) numa imagem PNG/JPG.
            Utiliza Matplotlib para renderizar o texto com fonte monoespaçada,
            preservando o alinhamento das tabelas do Statsmodels.
            """

            if not self.summary_var:
                messagebox.showwarning("Nada para Salvar", "Nenhum modelo foi calculado ainda.")

                return

            file_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG Image", "*.png"), ("JPG Image", "*.jpg"), ("All Files", "*.*")],
                title="Salvar Tabela como Imagem"
            )

            if not file_path:
                return  # Usuário cancelou

            try:
                summary_text = self.summary_var.as_text()
                # Cria uma figura longa para acomodar o texto do sumário
                fig = plt.figure(figsize=(10, 14), dpi=150)
                ax = fig.add_subplot(111)

                # Renderiza o texto simulando um terminal/console
                ax.text(0.01, 0.99, summary_text, va='top', ha='left',
                        fontfamily='monospace', fontsize=9, wrap=False)

                ax.axis('off')  # Remove eixos do gráfico
                fig.tight_layout(pad=0.5)
                fig.savefig(file_path)
                plt.close(fig)

                messagebox.showinfo("Sucesso", f"Tabela salva como imagem em:\n{file_path}")

            except Exception as e:

                messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar a imagem:\n{e}")

        def plotar_graficos(modelo, y_real, x_para_previsao):
            """
            Gera os gráficos de diagnóstico essenciais para a Regressão:

            1. Real vs. Previsto: Indica a precisão global (proximidade da linha 45°).

            2. Resíduos vs. Previstos: Avalia a Homocedasticidade (erros devem ser aleatórios).
            """

            # Limpa o frame de gráficos antes de plotar novos
            for widget in self.rlm_frame_graficos.winfo_children():
                widget.destroy()

            try:

                Label(self.rlm_frame_graficos, text="Gráficos de Diagnóstico",
                      font=('Arial', 11, 'bold')).pack(pady=(5, 2))

                # Extrai dados do modelo para plotagem
                previsoes = modelo.predict(x_para_previsao)
                residuos = modelo.resid
                fig = Figure(figsize=(6, 6), dpi=100)

                # --- Gráfico 1: Real vs. Previsto ---
                ax1 = fig.add_subplot(211)

                ax1.scatter(y_real, previsoes, alpha=0.6, edgecolors='k', s=30)

                # Linha de identidade (onde o real é igual ao previsto)
                lims = [min(y_real.min(), previsoes.min()), max(y_real.max(), previsoes.max())]

                ax1.plot(lims, lims, 'r--', lw=2, label='Perfeito (Y = Previsto)')
                ax1.set_xlabel('Valores Reais (Y)')
                ax1.set_ylabel('Valores Previstos (Ŷ)')
                ax1.set_title('Reais vs. Previstos')
                ax1.legend()
                ax1.grid(True, linestyle='--', alpha=0.5)

                # --- Gráfico 2: Resíduos vs. Previstos ---
                ax2 = fig.add_subplot(212)

                ax2.scatter(previsoes, residuos, alpha=0.6, edgecolors='k', s=30)
                ax2.axhline(y=0, color='r', linestyle='--', lw=2, label='Resíduo Zero')
                ax2.set_xlabel('Valores Previstos (Ŷ)')
                ax2.set_ylabel('Resíduos (Y - Ŷ)')
                ax2.set_title('Resíduos vs. Previstos (Homocedasticidade)')
                ax2.legend()
                ax2.grid(True, linestyle='--', alpha=0.5)
                fig.tight_layout(pad=2.0)

                # Integra o gráfico do Matplotlib no Tkinter
                canvas = FigureCanvasTkAgg(fig, master=self.rlm_frame_graficos)
                canvas.draw()

                # Adiciona a barra de ferramentas padrão do Matplotlib (zoom, save, etc.)
                toolbar = NavigationToolbar2Tk(canvas, self.rlm_frame_graficos)
                toolbar.update()

                canvas.get_tk_widget().pack(side='top', fill='both', expand=True)

            except Exception as e:

                Label(self.rlm_frame_graficos, text=f"Erro ao gerar gráficos:\n{e}", fg="red").pack(padx=10, pady=10)

        def recalcular_rlm(_=None):
            """
            Executa o cálculo estatístico OLS (Mínimos Quadrados Ordinários).

            Separa a variável dependente (Y) das independentes (X), adiciona o
            intercepto (constante) e gera o sumário (tabela) detalhado do modelo.
            """

            # Obtém o alvo (Y) escolhido na interface
            y_col_nome = self.rlm_combo_y.get()

            if not y_col_nome:  # Retorna caso o usuário não selecione o alvo (Y)
                return

            try:

                # Todas as outras colunas válidas tornam-se preditores (X)
                x_cols_nomes = [col for col in list_colunas_validas if col != y_col_nome]

                if not x_cols_nomes:
                    self.rlm_txt_summary.config(state='normal')
                    self.rlm_txt_summary.delete('1.0', 'end')
                    self.rlm_txt_summary.insert('1.0', "Erro: Não há variáveis independentes (X) selecionadas.")
                    self.rlm_txt_summary.config(state='disabled')
                    self.rlm_btn_salvar_tabela.config(state='disabled')

                    return

                # --- Processamento Estatístico ---
                # Adicionando uma constante (intercepto) para que a reta não passe pela origem (0,0)
                x_com_constante = sm.add_constant(self.df_validos[x_cols_nomes])

                # Ajustando o modelo de regressão comparando a variável dependente (Y) com as independentes (X)
                model = sm.OLS(self.df_validos[y_col_nome], x_com_constante).fit()

                # Armazenando o sumário completo (Ex: R², p-valor, coeficientes, etc.) para exibição e salvamento
                self.summary_var = model.summary()

                # --- Atualização da Interface ---
                self.rlm_txt_summary.config(state='normal')
                self.rlm_txt_summary.delete('1.0', 'end')
                self.rlm_txt_summary.insert('1.0', self.summary_var.as_text())
                self.rlm_txt_summary.config(state='disabled')
                self.rlm_btn_salvar_tabela.config(state='normal')

                # Atualiza os gráficos de diagnóstico com o novo modelo
                plotar_graficos(model, self.df_validos[y_col_nome], x_com_constante)

            except Exception as e:  # Tratamento de erros matemáticos (Ex: Matriz singular ou Multicolinearidade)

                self.rlm_txt_summary.config(state='normal')
                self.rlm_txt_summary.delete('1.0', 'end')
                self.rlm_txt_summary.insert('1.0', f"ERRO NO CÁLCULO DA RLM:\n\n{e}\n\n"
                                                   "Isso pode ser causado por multicolinearidade perfeita "
                                                   "(ex: uma coluna ser duplicata de outra) ou "
                                                   "dados insuficientes.")

                self.rlm_txt_summary.config(state='disabled')
                self.rlm_btn_salvar_tabela.config(state='disabled')

                for widget in self.rlm_frame_graficos.winfo_children():
                    widget.destroy()

                Label(self.rlm_frame_graficos, text="Cálculo falhou. Verifique os dados.", fg="red").pack(padx=10,
                                                                                                          pady=10)

        def criar_layout_rlm():
            """
            Constrói a estrutura de widgets da janela de RLM.
            Organiza em duas colunas principais: Sumário (Texto) e Diagnóstico (Gráficos).
            """
            # Painel Superior: Selecionar Y (Variável Dependente)
            frame_combobox = Frame(self.janela_secundaria, pady=5)
            frame_combobox.grid(row=0, column=0, columnspan=2, sticky='ew', padx=10, pady=5)

            Label(frame_combobox, text="Selecione a Variável Dependente (Y):",
                  font=('Arial', 10, 'bold')).pack(side='left', padx=(5, 10))

            self.rlm_combo_y = Combobox(frame_combobox, values=list_colunas_validas,
                                        state="readonly", width=30, font=('Arial', 10))
            self.rlm_combo_y.pack(side='left')
            self.rlm_combo_y.bind("<<ComboboxSelected>>", recalcular_rlm)

            Button(frame_combobox, text="—", font=('Arial', 10, 'bold'), width=3,
                   command=self.minimizar_tudo).pack(side='right', padx=5)  # Botão para minimizar tudo (—)

            # Painel Esquerdo: Exibição do Sumário de Texto
            frame_resultados = Frame(self.janela_secundaria, relief='sunken', bd=1)
            frame_resultados.grid(row=1, column=0, sticky='nsew', padx=(10, 5), pady=10)
            frame_resultados.grid_rowconfigure(1, weight=1)
            frame_resultados.grid_columnconfigure(0, weight=1)

            Label(frame_resultados, text="Sumário do Modelo (Tabela)", font=('Arial', 11, 'bold')
                  ).grid(row=0, column=0, pady=(5, 2))

            # Área de texto para o sumário (tabela) do Statsmodels
            frame_tabela_texto = Frame(frame_resultados)
            frame_tabela_texto.grid(row=1, column=0, sticky='nsew', padx=5)
            frame_tabela_texto.grid_rowconfigure(0, weight=1)
            frame_tabela_texto.grid_columnconfigure(0, weight=1)

            self.rlm_txt_summary = Text(frame_tabela_texto, wrap='none', font=('Courier New', 9), state='disabled',
                                        tabs=("1c", "2c", "3c", "4c", "5c"))

            self.rlm_txt_summary.grid(row=0, column=0, sticky='nsew')

            # Scroll duplo para acomodar o sumário
            vscroll = Scrollbar(frame_tabela_texto, orient='vertical', command=self.rlm_txt_summary.yview)
            hscroll = Scrollbar(frame_tabela_texto, orient='horizontal', command=self.rlm_txt_summary.xview)
            vscroll.grid(row=0, column=1, sticky='ns')
            hscroll.grid(row=1, column=0, sticky='ew')
            self.rlm_txt_summary.config(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)

            # Botão "Salvar Tabela como Imagem" inicia desativado
            self.rlm_btn_salvar_tabela = Button(frame_resultados, text="Salvar Tabela como Imagem",
                                                command=salvar_imagem_tabela, state='disabled')

            self.rlm_btn_salvar_tabela.grid(row=2, column=0, pady=(5, 10))

            # Painel Direito: Espaço para os Gráficos
            self.rlm_frame_graficos = Frame(self.janela_secundaria, relief='sunken', bd=1)
            self.rlm_frame_graficos.grid(row=1, column=1, sticky='nsew', padx=(5, 10), pady=10)
            self.rlm_frame_graficos.grid_rowconfigure(0, weight=1)
            self.rlm_frame_graficos.grid_columnconfigure(0, weight=1)

            Label(self.rlm_frame_graficos, text="Aguardando seleção da variável dependente...",
                  font=('Arial', 10, 'italic'), fg='grey').pack(pady=20, padx=20)

        # Inicializa o layout ao abrir a janela secundária
        criar_layout_rlm()

    def minimizar_tudo(self):
        """
        Minimiza o software inteiro (Janela principal e secundárias).

        Libera o 'grab' (foco exclusivo) da janela secundária antes de minimizar
        para evitar que a interface fique travada ao retornar.
        """

        if self.janela_secundaria and self.janela_secundaria.winfo_exists():
            # Solta o foco exclusivo para permitir que o sistema operacional minimize a janela
            self.janela_secundaria.grab_release()

        # Minimiza a aplicação para a barra de tarefas
        self.master.iconify()

    def _reaplicar_grab(self, _=None):
        """
        Recupera o foco exclusivo da janela secundária.

        Utilizado após a janela ser restaurada (de-iconify) ou mapeada,
        garantindo que o usuário não consiga interagir com a janela principal
        enquanto a secundária estiver aberta.
        """

        if self.janela_secundaria and self.janela_secundaria.winfo_exists():
            current_grab = self.janela_secundaria.grab_current()

            # Se a janela secundária perdeu o controle do foco, força-o de volta
            if current_grab != self.janela_secundaria:
                self.janela_secundaria.grab_set()
                self.janela_secundaria.focus_set()


if __name__ == "__main__":  # Instancia a classe e abre a janela do sistema
    MuLiRA = Janela()
    MuLiRA.mainloop()
