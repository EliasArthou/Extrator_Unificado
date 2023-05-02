"""
Constrói a janela
"""

import sys
import tkinter as tk
from tkinter import ttk
import extracao


class App(tk.Tk):
    """
    Cria janela com retorno para o usuário
    """

    def __init__(self):
        super().__init__()

        # Largura da Janela

        self.labelradio = None
        self.cmbtiposervico = None
        self.labelservico = None
        self.cotaunica = None
        self.cotaparcelada = None
        self.data1 = None
        self.data2 = None
        self.data3 = None
        self.c1 = None
        self.c2 = None
        # self.c3 = None
        self.c4 = None
        self.tipoextracao = ''
        self.tiposervico = ''
        self.somentevalores = None
        self.codigosdebarra = None
        self.nadaconsta = None
        self.faltantes = True
        self.executar = None

        w = 500
        # Altura da Janela
        h = 162

        # Define a janela como não exclusiva (outras janelas podem sobrepor ela)
        self.acertaconfjanela(False)
        # Adiciona o cabeçalho da janela
        self.title('Andamento Extração')

        self.minsize(w, h)
        # Desenha a janela com a largura e altura definida e na posição calculada, ou seja, no centro da tela
        self.center()

        # Label de quantidade de extrações
        self.labelextracao = ttk.Label(self, text='Selecione a Extração será realizada:', font="Arial 10 bold")
        self.labelextracao.place(x=0, y=0)

        # Cria um Combobox
        self.tipoextracao = tk.StringVar()
        self.cmbtipoextracao = ttk.Combobox(self, textvariable=self.tipoextracao)
        self.cmbtipoextracao['state'] = 'readonly'
        self.cmbtipoextracao['values'] = ['', 'Prefeitura', 'Bombeiros', 'Condomínios']
        self.cmbtipoextracao.place(x=252, y=1, width=self.winfo_width()-252)
        self.cmbtipoextracao.current(0)
        self.cmbtipoextracao.bind('<<ComboboxSelected>>', self.extracao_changed)

        # Label número cliente
        self.labelcodigocliente = ttk.Label(self, text='Código Cliente:         ', font="Arial 15 bold")
        self.labelcodigocliente.place(x=0, y=17, width=self.winfo_width())
        self.labelcodigocliente.configure(anchor='center')

        # Label número da inscrição
        self.labelinscricao = ttk.Label(self, text='', font="Arial 15")
        self.labelinscricao.place(x=0, y=42, width=self.winfo_width())
        self.labelinscricao.configure(anchor='center')

        # ProgressBar de Quantidade de Transações (Views)
        self.barraextracao = ttk.Progressbar(self, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.barraextracao.place(x=(w-300)/2, y=70, width=300)

        # Label de quantidade de extrações
        self.labelquantidade = ttk.Label(self, text='', font="Arial 10")
        self.labelquantidade.place(x=0, y=100, width=self.winfo_width())
        self.labelquantidade.configure(anchor='center')

        # Label de estado do sistema
        self.labelstatus = ttk.Label(self, text='', font="Arial 15")
        self.labelstatus.place(x=0, y=120, width=self.winfo_width())
        self.labelstatus.configure(anchor='center')

        # Extrair somente valores
        self.somentevalores = tk.BooleanVar()

        # Salvar Código de Barras no BD
        self.codigosdebarra = tk.BooleanVar()

        # Extrair Nada Consta
        self.nadaconsta = tk.BooleanVar()

        # Extrair Nada Consta
        self.faltantes = tk.BooleanVar(value=True)

        self.tipopagamento = tk.IntVar()
        self.tipopagamento.set(2)  # Para a segunda opção ficar marcada
        self.manipularradio(self.tipoextracao.get())

        # button Iniciar Extração
        self.executar = ttk.Button(self, text='Executar', state='disabled')
        self.executar['command'] = self.executar_clicked
        self.executar.place(x=10, y=170)

        # button Fechar a Janela
        self.fechar = ttk.Button(self, text='Fechar')
        self.fechar['command'] = self.fechar_clicked
        self.fechar.place(x=self.winfo_width() - 90, y=170)

    def manipularradio(self, tipoextracao, criarradio=True):
        """
        : param criarradio: se cria ou destrói os "radio buttons".
        : param tipoextracao: ajuda a definir a tela que tem que ser carregada baseada na extração que será executada.
        : return:
        """

        self.mudartexto('labelcodigocliente', 'Código Cliente:     ')
        self.mudartexto('labelinscricao', '')
        self.mudartexto('labelquantidade', '')
        self.mudartexto('labelstatus', '')
        self.configurarbarra('barraextracao', 1, 0)
        if len(tipoextracao) > 0:
            if criarradio:
                self.c1 = tk.Checkbutton(self, text='Somente Valores', variable=self.somentevalores, onvalue=True, offvalue=False, font="Arial 10")
                self.c1.place(relx=0.040, y=142)
            self.executar.state(["!disabled"])
        else:
            if self.executar is not None:
                self.executar.state(["disabled"])

        match tipoextracao:
            case 'Bombeiros':
                if criarradio:
                    # Criando os radio buttons e o label
                    self.labelradio = ttk.Label(self, text='Extrair qual tipo de pagamento?')
                    self.labelradio.place(relx=0.05, y=100)
                    self.cotaunica = ttk.Radiobutton(self, text='Cota Única', variable=self.tipopagamento, value=1)
                    self.cotaunica.place(relx=0.05, y=120)
                    self.cotaparcelada = ttk.Radiobutton(self, text='Parcelado', variable=self.tipopagamento, value=2)
                    self.cotaparcelada.place(relx=0.25, y=120)
                    self.tipopagamento.set(2)

                else:
                    self.labelradio.destroy()
                    self.cotaunica.destroy()
                    self.cotaparcelada.destroy()
                    self.cmbtiposervico.destroy()
                    self.labelservico.destroy()
                    # self.c1.destroy()

            case 'Prefeitura':
                if criarradio:
                    # Criando os radio buttons e o label
                    self.labelradio = ttk.Label(self, text='Extrair qual data?')
                    self.labelradio.place(relx=0.05, y=100)
                    self.cotaunica = ttk.Radiobutton(self, text='Cota Única', variable=self.tipopagamento, value=1)
                    self.cotaunica.place(relx=0.05, y=120)
                    self.data1 = ttk.Radiobutton(self, text='Data 1', variable=self.tipopagamento, value=2)
                    self.data1.place(relx=0.25, y=120)
                    self.data2 = ttk.Radiobutton(self, text='Data 2', variable=self.tipopagamento, value=3)
                    self.data2.place(relx=0.4, y=120)
                    self.data3 = ttk.Radiobutton(self, text='Data 3', variable=self.tipopagamento, value=4)
                    self.data3.place(relx=0.55, y=120)
                    # "Checkbutton" para subir os códigos de barras
                    self.c2 = tk.Checkbutton(self, text='Subir Código de Barras', variable=self.codigosdebarra, onvalue=True, offvalue=False, font="Arial 10")
                    self.c2.place(relx=0.30, y=142)
                    self.tipopagamento.set(2)
                    # "Checkbutton" nada consta
                    # self.c3 = tk.Checkbutton(self, text='Nada Consta', variable=self.nadaconsta, onvalue=True, offvalue=False, font="Arial 10")
                    # self.c3.place(relx=0.74, y=142)
                    # "Checkbutton" de lista completa ou somente faltantes
                    self.c4 = tk.Checkbutton(self, text='Somente Faltantes?', variable=self.faltantes, onvalue=True, offvalue=False, font="Arial 10")
                    self.c4.place(relx=0.65, y=142)
                    # Cria a combobox pra selecionar o serviço da prefeitura
                    self.tiposervico = tk.StringVar()
                    self.labelservico = ttk.Label(self, text='Serviço:')
                    self.labelservico.place(relx=0.70, y=95)
                    self.cmbtiposervico = ttk.Combobox(self, textvariable=self.tiposervico)
                    self.cmbtiposervico['state'] = 'readonly'
                    self.cmbtiposervico['values'] = ['IPTU', 'Nada Consta', 'Certidão Negativa']
                    self.cmbtiposervico.place(relx=0.70, y=118, width=110)
                    self.cmbtiposervico.current(0)
                    self.cmbtiposervico.bind('<<ComboboxSelected>>', self.extracao_changed)

                else:
                    self.labelradio.destroy()
                    self.cotaunica.destroy()
                    self.labelservico.destroy()
                    self.cmbtiposervico.destroy()
                    self.data1.destroy()
                    self.data2.destroy()
                    self.data3.destroy()
                    self.c1.destroy()
                    self.c2.destroy()
                    # self.c3.destroy()
                    self.c4.destroy()

            case _:
                if self.labelradio is not None:
                    self.labelradio.destroy()
                if self.labelservico is not None:
                    self.labelservico.destroy()
                if self.cmbtiposervico is not None:
                    self.cmbtiposervico.destroy()
                if self.cotaunica is not None:
                    self.cotaunica.destroy()
                if self.data1 is not None:
                    self.data1.destroy()
                if self.data2 is not None:
                    self.data2.destroy()
                if self.data3 is not None:
                    self.data3.destroy()
                if self.cotaunica is not None:
                    self.cotaunica.destroy()
                if self.cotaparcelada is not None:
                    self.cotaparcelada.destroy()
                if self.c1 is not None:
                    self.c1.destroy()
                if self.c2 is not None:
                    self.c2.destroy()
                # if self.c3 is not None:
                #     self.c3.destroy()
                if self.c4 is not None:
                    self.c4.destroy()

        self.atualizatela()

    # bind the selected value changes
    def extracao_changed(self, event):
        """ Evento de mudança do combobox """
        self.manipularradio('', True)
        self.manipularradio(self.tipoextracao.get(), True)

    def executar_clicked(self):
        """
        Ação do botão
        """
        self.manipularradio(self.tipoextracao.get(), False)
        extrator = extracao.Extrator(self)
        extrator.controlaextracao()
        """
        self.manipularradio(self.tipoextracao.get(), False)
        match self.tipoextracao.get():
            case 'Bombeiros':
                extrairboletosbombeiro(self)

            case 'IPTU':
                extrairboletosiptu(self)

            case _:
                msg.msgbox(f'Opção Inválida!', msg.MB_OK, 'Tempo Decorrido')
        """

    def fechar_clicked(self):
        """
        Ação do botão
        """
        self.destroy()
        sys.exit()

    def mudartexto(self, nomelabel, texto):
        """
        :param nomelabel: nome do label a ter o texto alterado
        :param texto: texto a ser inserido
        """
        self.__getattribute__(nomelabel).config(text=texto)
        self.atualizatela()

    def configurarbarra(self, nomebarra, maximo, indicador):
        """
        : param nomebarra: nome da barra a ser atualizada.
        : param maximo: limite máximo da barra de progresso.
        : param indicador: variável
        """
        self.__getattribute__(nomebarra).config(maximum=maximo, value=indicador)
        self.atualizatela()

    def acertaconfjanela(self, exclusiva):
        """
        : param exclusiva: se a janela fica na frente das outras ou não
        """
        self.attributes("-topmost", exclusiva)
        self.atualizatela()

    def atualizatela(self):
        """
        Dá um 'refresh' na tela para modificar com alterações realizadas
        """
        self.update()

    def center(self):
        """
        :param: the main window or Toplevel window to center

        Apparently a common hack to get the window size. Temporarily hide the
        window to avoid update_idletasks() drawing the window in the wrong
        position.
        """

        self.update_idletasks()  # Update "requested size" from geometry manager

        # define window dimensions width and height
        width = self.winfo_width()
        frm_width = self.winfo_rootx() - self.winfo_x()
        win_width = width + 2 * frm_width

        height = self.winfo_height()
        titlebar_height = self.winfo_rooty() - self.winfo_y()
        win_height = height + titlebar_height + frm_width

        # Get the window position from the top dynamically as well as position from left or right as follows
        x = self.winfo_screenwidth() // 2 - win_width // 2
        y = self.winfo_screenheight() // 2 - win_height // 2

        # this is the line that will center your window
        self.geometry('{}x{}+{}+{}'.format(width, height, x, y))

        # This seems to draw the window frame immediately, so only call deiconify()
        # after setting correct window position
        self.deiconify()
