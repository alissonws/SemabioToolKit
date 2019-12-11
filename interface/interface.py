#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
import tkinter.ttk as ttk
from tkinter.filedialog import (askopenfilename, askdirectory)


import PIL.Image, PIL.ImageTk


import platform, re,os, traceback, logging,smtplib
from multiprocessing import Queue, Process, freeze_support
from pathlib import Path



from utils import emails
from utils.certs_generator import *

#start swriter --headless --convert-to pdf --outdir "c:\Users\aliss\Desktop" "c:\Users\aliss\Desktop\Novo(a) Documento do Microsoft Word.docx"
class Initialize(tk.Tk):
    def __init__(self,app_version, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        self.CONVERTER_PATH = ""


        #This container will stack all the frames (pages), allowing the screen selection
        self.container = tk.Frame(self)
        self.container.grid()
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        #STYLES FOR TTK
        style = ttk.Style()

        style.theme_use("clam")
        #Menu buttons
        style.configure("Menu.TButton", padding=8, relief="flat",
                              background="#52b8cf", font=("Noto Sans CJK JP Regular",10,"bold"),foreground="white")
        style.map("Menu.TButton",
        foreground=[('pressed', "white"), ('active', "white")],
        background=[('pressed', "#10274c"),('disabled',"grey"), ('active', "#1d799e")]
        )
        #Main buttons
        style.configure("Main.TButton", padding=4, relief= "flat",
                              background="#52b8cf", font=("Noto Sans CJK JP Regular",10,"bold"),foreground="white")
        style.map("Main.TButton",
        foreground=[('pressed', "white"), ('active', "white")],
        background=[('pressed',  "#10274c"),('disabled',"grey"), ('active', "#1d799e")]
        )
        #Common buttons
        style.configure("General.TButton", padding=1,
                              width= 12, font=("Noto Sans CJK JP Regular",10),foreground="black")

        #Title labels
        style.configure("Main.TLabel", background="white",
                        font=("Noto Sans CJK JP Regular",12,"bold"),foreground="black")
        #Common labels
        style.configure("General.TLabel", background="white",
                        font=("Chandas",11),foreground="black")


        style.configure("TCheckbutton", background="white",
                        font=("Noto Sans CJK JP Regular", 10), foreground="black")


        #Save the reference!

        self.frames = {}
        # Align classes to their differents geometries
        for F, geometry in zip((Login,Menu, Certificados, Email, Configuracoes),("290x400","290x400","520x440","600x630","540x170")):
            page_name = F.__name__
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = (frame, geometry)
            frame.configure(background="white")
            frame.grid(row=0, column=0, sticky="nsew")

        self.transition("Menu")

        #Check if the converter ir installed
        self.depCheck()


    def transition(self,page_name,i=1):
        """
        This function control the transition between frames
        :param page_name: the frame that will me shown
        :param i: orientation of the transition (don't change it)
        :return:
        """
        pad_info = self.container.grid_info()['pady']
        pad = pad_info +10*i
        if i==1:
            if pad_info < 300:
                self.container.grid(pady=pad)
                self.after(4,lambda :self.transition(page_name,i))
            else:
                self.after(4, lambda: self.transition(page_name, -1))
        elif pad_info > 0:
            self.show_frame(page_name)
            self.container.grid(pady=pad)
            self.after(4, lambda: self.transition(page_name,i))
        else:
            self.container.grid(pady=0)


    def show_frame(self, page_name):
        """
        Shows a frame
        :param page_name: name of the frame
        :return:
        """
        frame, geometry = self.frames[page_name]
        self.update_idletasks()
        self.geometry(str(frame.winfo_reqwidth())+"x"+str(frame.winfo_reqheight()))
        frame.tkraise()

    def depCheck(self, path=None):
        """
        Checks if LibreOffice is installed. It is needed for converting .docx to .pdf on "Certificados"
        :param path: User can define a specific path on "Configurações"
        :return:
        """

        logging.debug("Checking dependencies...")

        if path==None:
            path = [r'C:\\Program Files\\LibreOffice\\program\\soffice.exe']
        else:
            path = [path]

        for p in path:
            no_converter = 0
            p = p.replace("/",r"\\")

            converter = subprocess.check_output(
                "wmic datafile where name='"+p+"' get Version /value",
                shell=True)
            if "LibreOffice" in p:
                if "Version" in str(converter):
                    logging.info("Dependencies checked")
                    self.CONVERTER_PATH = p
                    return p
                else:
                    logging.warning("Path is not valid")
                    no_converter = 1
            else:
                logging.warning("Path provided does not lead to .exe")
                no_converter = 1

        if no_converter == 1:
            logging.warning("A .pdf converter was not found in the machine. 'Certificados' will not work well")
            popup("Foi detectado que você não tem o software LibreOffice instalado\n"
                  "Essa aplicação exige uma versão recente deste programa para\n"
                  "executar certas funções.", title="Dependências faltando")
            return False



class Login(tk.Frame):
    def __init__(self, parent, controller):
        #GLOBAL VARIABLES
        global GUI_PATH, APP_VERSION, PLATFORM
        tk.Frame.__init__(self, parent)
        self.controller = controller


        self.login = tk.Frame(self)
        self.login["bg"] = "white"
        self.login.grid()

        self.login.bind_all("<Return>", self.credentialsFilter)



        try:
            #for freeze version

            if PLATFORM == "Linux":
                raw_image = PIL.Image.open(sys._MEIPASS + "/icons/semabiologo.png")
            elif PLATFORM == "Windows":
                raw_image = PIL.Image.open(sys._MEIPASS + r"\icons\semabiologo.png")
            else:
                logging.critical("The program has not been tasted on this platform. Errors may occur")

        except:
            #for script version script
            if PLATFORM == "Linux":
                raw_image = PIL.Image.open(GUI_PATH.replace("/interface","/icons/semabiologo.png"))
            elif PLATFORM == "Windows":
                raw_image = PIL.Image.open(GUI_PATH.replace(r"\interface", r"\icons\semabiologo.png"))
            else:
                logging.critical("The program has not been tasted on this platform. Errors may occur")


        image = raw_image.resize((150, 229), PIL.Image.ANTIALIAS)
        self.photo = PIL.ImageTk.PhotoImage(image)


        self.logo = tk.Label(self.login, image=self.photo)
        self.logo["bg"] = "white"
        self.logo.image = self.photo
        self.logo.grid(row=0, column = 0, pady = 0)

        self.label_email = tk.Label(self.login, text="E-mail")
        self.label_email["font"] = ("Verdana", "7")
        self.label_email["bg"] = "white"
        self.label_email["fg"] = "black"
        self.label_email.grid(row=1, column = 0, sticky = "ws")

        self.entry_email = ttk.Entry(self.login)
        self.entry_email.grid(row=2, column =0, padx=20)
        self.entry_email["width"] = 30
        self.entry_email.focus_set()


        self.label_senha = tk.Label(self.login, text="Senha")
        self.label_senha["font"] = ("Verdana", "7")
        self.label_senha["fg"] = "black"
        self.label_senha["bg"] = "white"
        self.label_senha.grid(row=3, column = 0, sticky = "ws")

        self.entry_senha = ttk.Entry(self.login)
        self.entry_senha.grid(row=4, column=0)
        self.entry_senha["width"] = 30
        self.entry_senha["show"] = "*"


        self.button_login = ttk.Button(self.login,text="ENTRAR",style="Main.TButton",
                                       width=29,
                                       command = self.credentialsFilter)
        self.button_login.grid(row=5, column = 0,pady= 30)

        self.label_app_version = tk.Label(self.login, text=APP_VERSION, bg="white")
        self.label_app_version.grid(row=6,sticky="e")



    def credentialsFilter(self,event=None):
        email = self.entry_email.get()
        senha = self.entry_senha.get()

        #checando pela validade das credenciais
        if email == "" or senha =="":
            if email == "" and senha != "":
                self.label_dados_invalidos = tk.Label(self.login, text="Insira um E-mail válido")
                self.label_dados_invalidos["font"] = ("Verdana", "7")
                self.label_dados_invalidos["fg"] = "red"
                self.label_dados_invalidos["bg"] = "white"
                self.label_dados_invalidos.grid(row=3, column=0, sticky="es")
            elif senha == "" and email != "":
                self.label_dados_invalidos = tk.Label(self.login, text="Insira uma senha válida")
                self.label_dados_invalidos["font"] = ("Verdana", "7")
                self.label_dados_invalidos["fg"] = "red"
                self.label_dados_invalidos["bg"] = "white"
                self.label_dados_invalidos.grid(row=3, column=0, sticky="es")
            else:
                try:
                    self.label_dados_invalidos
                except:
                    self.label_dados_invalidos = tk.Label(self.login, text="Dados incorretos")
                    self.label_dados_invalidos["font"] = ("Verdana", "7")
                    self.label_dados_invalidos["fg"] = "red"
                    self.label_dados_invalidos["bg"] = "white"
                    self.label_dados_invalidos.grid(row=3, column=0, sticky="es")
                else:
                    self.label_dados_invalidos.grid_remove()
                    self.label_dados_invalidos = tk.Label(self.login, text="Dados incorretos")
                    self.label_dados_invalidos["font"] = ("Verdana", "7")
                    self.label_dados_invalidos["fg"] = "red"
                    self.label_dados_invalidos["bg"] = "white"
                    self.label_dados_invalidos.grid(row=3, column=0, sticky="es")
        else:
            self.button_login["text"] = "ENTRANDO..."
            self.login.unbind_all("<Return>")
            self.controller.update()
            #tentativa de login
            check = emails.EmailsSBChecking()
            try:
                login = check.check_login(email, senha)
            except smtplib.SMTPAuthenticationError as e:
                print(e)
                self.button_login["text"] ="ENTRAR"
                popup("Credenciais inválidas, revise seu e-mail e sua senha.",title="Login inválido")
                self.login.bind_all("<Return>", self.credentialsFilter)
            except socket.gaierror as e:
                print(e)
                self.button_login["text"] ="ENTRAR"
                popup("Não foi possível se conectar ao servidor. Cheque\nsua conexão com a internet e tente\nnovamente.",title="Login inválido")
                self.login.bind_all("<Return>", self.credentialsFilter)
            except Exception as e:
                print(e)
                self.button_login["text"] ="ENTRAR"
                popup("Um erro inesperado ocorreu. Contate o desenvolvedor",title="Erro")
                self.login.bind_all("<Return>", self.credentialsFilter)
            else:
                if login == True:
                    global fromaddr, pwd
                    fromaddr = email
                    pwd = senha

                    self.controller.transition("Menu")
                    self.button_login["text"] = "ENTRAR"
                    self.entry_email.delete(0, "end")
                    self.entry_senha.delete(0, "end")


class Menu(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        self.menu = tk.Frame(self, background = "white")
        self.menu.grid()


        self.label_menu1 = ttk.Label(self.menu, text = "MENU", style = "Main.TLabel")
        self.label_menu1.grid(row=0, pady = 40, padx = 125, sticky = "NSWE")

        #CERTIFICADOS

        self.button_menu1 = ttk.Button(self.menu,width = 17, text = "CERTIFICADOS",style="Menu.TButton",
                                   command = lambda : self.controller.transition("Certificados"))
        self.button_menu1.grid(row=1, pady = 0,padx = 30)

        #ENVIO DE EMAILS

        self.button_menu2 = ttk.Button(self.menu,width = 17, text = "ENVIO DE E-MAILS", style = "Menu.TButton",
                                   command = lambda:self.controller.transition("Email"))
        self.button_menu2.grid(row=2, pady = 20,padx = 30)


        # CONFIGURAÇÕES
        self.button_manual = ttk.Button(self.menu,width = 17, state = "disabled", text = "CONFIGURAÇÕES", style = "Menu.TButton",
                                   command = lambda:self.controller.transition("Configuracoes"))
        self.button_manual.grid(row=3, pady = 0,padx = 30)

        #SAIR

        self.button_manual = ttk.Button(self.menu,width = 17, text = "SAIR", style = "Menu.TButton",
                                   command = lambda :self.controller.transition("Login"))
        self.button_manual.grid(row=4, pady = 20,padx = 30)

        #MARGEM INFERIOR
        self.border = ttk.Label(self.menu,text="  ", style="General.TLabel")
        self.border.grid(row=5)


class Certificados(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)
        self.controller = controller

        #VARIÁVEIS
        self.PLAN_PATH = ""
        self.FILE_PATH = ""
        self.SAVE_PATH = ""
        self.TAGS = {}
        self.TUPLES = []

        #CABEÇALHO
        self.top = tk.Frame(self, background = "white")
        self.top.grid(sticky="w")

        self.top_title = ttk.Label(self.top, text = "Gerador de Certificados",style="Main.TLabel")
        self.top_title.grid(padx=10,pady=10,sticky="w")

        #PLANILHA DE DADOS
        self.top_label_planinsert = ttk.Label(self.top, text = "Planilha de dados", style = "General.TLabel")
        self.top_label_planinsert.grid(row = 2, pady=10, padx=20, sticky="w")

        self.top_button_planinsert = ttk.Button(self.top, text = "Carregar", style="General.TButton",
                                            command = lambda :self.fileInput(what=".xlsx"))
        self.top_button_planinsert.grid(row = 2, column = 1, padx=0, sticky="e")

        self.top_entry_showplan = ttk.Entry(self.top, width = 17, state = "readonly")
        self.top_entry_showplan.grid(row=2, column = 2, sticky="w", padx=10)

        #TEAMPLATE
        self.top_label_planinsert = ttk.Label(self.top, text = "Projeto de certificado",style="General.TLabel")
        self.top_label_planinsert.grid(row = 3,pady=10, padx=20, sticky="w")

        self.top_button_planinsert = ttk.Button(self.top, text = "Carregar", style="General.TButton",
                                            command = lambda :self.fileInput(what=".docx"))
        self.top_button_planinsert.grid(row = 3, column = 1, padx=0, sticky="e")

        self.top_entry_showfile = ttk.Entry(self.top, width = 17, state = "readonly")
        self.top_entry_showfile.grid(row=3, column = 2, sticky="w", padx=10)

        #PASTA DE SALVAMENTO
        self.top_label_savefolder = ttk.Label(self.top, text = "Salvar em:",style="General.TLabel")
        self.top_label_savefolder.grid(row = 4, pady=10, padx=20, sticky="w")

        self.top_button_savefolder = ttk.Button(self.top, text = "Abrir", style="General.TButton",
                                            command = self.dirInput)
        self.top_button_savefolder.grid(row = 4, column = 1, padx=0, sticky="e")

        self.top_entry_showsavefolder = ttk.Entry(self.top, width = 17, state = "readonly")
        self.top_entry_showsavefolder.grid(row=4, column = 2, sticky="w", padx=10)


        #HEADER
        self.top_checkbut_header_var = tk.IntVar()
        self.top_checkbutton_header = ttk.Checkbutton(self.top, text = "Primeira linha como cabeçalho", style="TCheckbutton",
                                                      variable = self.top_checkbut_header_var)
        self.top_checkbutton_header.grid(row=5,padx=20, column = 0, sticky="w")

        #ESCREVER ARQUIVO NA PLANILHA
        self.top_checkbut_write_var = tk.IntVar()
        self.top_checkbutton_write = ttk.Checkbutton(self.top, text = "Escrever nome do arquivo final na planilha", style="TCheckbutton",
                                                      variable = self.top_checkbut_write_var,
                                                     command = self.checkbutWriteSelect)
        self.top_checkbutton_write.grid(row=6,column=0,padx=20, sticky="w")

        self.top_checkbutton_strvar = tk.StringVar()
        self.top_checkbutton_strvar.trace("w", lambda a, b, c: self.entryTagCorrector(self.top_checkbutton_strvar))

        self.top_entry_write = ttk.Entry(self.top, width=2,textvariable=self.top_checkbutton_strvar)
        self.top_entry_write.grid(row = 6,column=1)
        self.top_entry_write['state'] = "disabled"

        #TAGS
        self.tags = tk.Frame(self, background = "white")
        self.tags.grid(row=5, sticky="w")

        self.name_strvar = tk.StringVar()
        self.name_strvar.trace("w", lambda a, b, c: self.entryTagCorrector(self.name_strvar))
        self.cpf_strvar = tk.StringVar()
        self.cpf_strvar.trace("w", lambda a, b, c: self.entryTagCorrector(self.cpf_strvar))

        self.TAGS = {1:["Nome", "<nome>",ttk.Label(self.tags, text = "Nome", style= "General.TLabel"),ttk.Entry(self.tags, width=5, textvariable=self.name_strvar)],
                     2:["CPF (opcional)","<cpf>",ttk.Label(self.tags, text = "CPF (opcional)",style="General.TLabel"),ttk.Entry(self.tags, width=5, textvariable=self.cpf_strvar)]}


        #gerando blocos personalizados
        ####
        #ADICIONAR SISTEMA DE ARMAZENAMENTO
        ####

        self.update_tags()

        #ADICIONAR/EXCLUIR TAGS

        self.tagbutton = tk.Frame(self, background = "white")
        self.tagbutton.grid(row=6,sticky="w")

        self.tagbutton_button_addtag = ttk.Button(self.tagbutton, style="General.TButton",width=15,
                                              text = "Adicionar tag", command = self.add_tag)
        self.tagbutton_button_addtag.grid(row=6, padx=20)

        self.tagbutton_button_deltag = ttk.Button(self.tagbutton, style="General.TButton",width=15,
                                              text = "Deletar tag", command = self.del_tags)
        self.tagbutton_button_deltag.grid(row=6,column=1, padx=10)

        #BOTÕES

        self.buttons = tk.Frame(self, background = "white")
        self.buttons.grid(row=7, padx=10,pady=20, sticky="w")


        self.buttons_generate = ttk.Button(self.buttons, text = "Gerar", style = "Main.TButton",
                                           command = self.genAttempt)

        self.buttons_generate.grid(row=0,column=0, padx=10)

        self.buttons_space = ttk.Label(self.buttons, width=30, background="white")
        self.buttons_space.grid(row=0, column=1, padx=10,sticky="es")

        self.buttons_backtomenu = ttk.Button(self.buttons, text="Voltar", style = "Main.TButton",
                                         command=lambda: self.controller.transition("Menu"))
        self.buttons_backtomenu.grid(row=0, column=3, padx=10,sticky="es")


    def entryTagCorrector(self, strvar):
        value = strvar.get()
        if len(value) > 1:
            value = value[0]
            strvar.set(value)
        if re.match("[a-z]",value) != None or re.match("[A-Z]",value) != None:
            strvar.set(value.capitalize())
        else: strvar.set("")


    def fileInput(self, what):
        path = filedialog()

        if path and what == ".xlsx":
            try:
                load = PyCertsSB()
                load.check_file(path,".xlsx")

                self.PLAN_PATH = path

                self.top_entry_showplan["state"] = "normal"
                self.top_entry_showplan.delete(0, 'end')
                self.top_entry_showplan.insert("end", path.split("/")[-1])

                self.top_entry_showplan["state"] = "readonly"
            except:
                popup("Selecione uma planilha .xlsx válida.",title="Arquivo inválido")

        if path and what == ".docx":
            try:
                load = PyCertsSB()
                load.check_file(path,".docx")

                self.FILE_PATH = path

                self.top_entry_showfile["state"] = "normal"
                self.top_entry_showfile.delete(0, 'end')
                self.top_entry_showfile.insert("end", path.split("/")[-1])
                self.top_entry_showfile["state"] = "readonly"
            except:
                popup("Selecione um arquivo válido.\nFormato aceito: .docx.",title="Arquivo inválido")

    def dirInput(self):
        path = dirdialog()

        if path != ():
            self.SAVE_PATH = path

            self.top_entry_showsavefolder["state"] = "normal"
            self.top_entry_showsavefolder.delete(0, 'end')
            self.top_entry_showsavefolder.insert("end", path.split("/")[-1])
            self.top_entry_showsavefolder["state"] = "readonly"


    def update_tags(self, event=None):
        #Inputs de colunas

        row=0

        # criando inputs personalizados

        if len(self.TAGS) > 0:
            for tags in self.TAGS:
                if tags > 3:
                    row=1
                label_ref = self.TAGS[tags][2]
                entry_ref = self.TAGS[tags][3]
                label_ref.grid(row=row, column=tags*2 - 2 -(6*row), pady=20, padx=20, sticky="w")
                entry_ref.grid(row=row, column=tags*2 - 1 - (6*row))

        self.controller.geometry(str(self.controller.winfo_reqwidth()) + "x" + str(self.controller.winfo_reqheight()))


    def add_tag(self, event = None):
        """
        This function opens a TopLevel interface so the user may insert custom tags to the certifies
        :param event:
        :return:
        """


        #CREATING A TOPLEVEL
        self.topl_add_tag = tk.Toplevel(bg="white")
        self.topl_add_tag.title("Adicione novas tags")

        #TopLevel layout
        top_label = tk.Label(self.topl_add_tag, bg="white",text="Adicione novas tags para personalizar o certificado.", font=("Arial",10,"bold"))
        top_label.grid(pady=10, padx=10, columnspan=3)

        #Name of the tag
        label1 = ttk.Label(self.topl_add_tag, text = "Título",style="General.TLabel")
        label1.grid(row=1,padx=10, sticky="w")

        self.add_tag_entry1 = ttk.Entry(self.topl_add_tag, width = 15)
        self.add_tag_entry1.grid(row=1, column=1, sticky="w")
        self.add_tag_entry1.focus_set()

        #"Tag" of the tag
        label2 = ttk.Label(self.topl_add_tag, text="Tag",style="General.TLabel")
        label2.grid(row=3, padx=10,sticky="w")

        self.add_tag_entry2 = ttk.Entry(self.topl_add_tag, width=15)
        self.add_tag_entry2.grid(row=3, column=1, sticky="w")


        B1 = ttk.Button(self.topl_add_tag, text="Pronto", style="Main.TButton",command = self.set_tags)
        B1.grid(row = 5, pady=20, padx=10, sticky="w")
        B1.bind_all("<Return>", self.set_tags)


        self.topl_add_tag.resizable(width=False, height=False)

        #Grab the TopLevel prevents user from interacting with the main window, it prevents duplicates
        self.topl_add_tag.grab_set()

        #Setting widith, heigh and coordinates for toplevel
        self.topl_add_tag.geometry(windowSizing(self.topl_add_tag))

        #Set the TopLevel icon
        try:
            # for freeze version
            if PLATFORM == "Linux":
                self.topl_add_tag.iconbitmap(sys._MEIPASS + "/icons/iconapp.ico")
            elif PLATFORM == "Windows":
                self.topl_add_tag.iconbitmap(sys._MEIPASS + r"\icons\iconapp.ico")
            else:
                logging.critical("The program has not been tested on this platform. Errors may occur")
        except:
            # for script version script
            if PLATFORM == "Linux":
                self.topl_add_tag.iconbitmap(GUI_PATH.replace("/interface", "/icons/iconapp.ico"))
            elif PLATFORM == "Windows":
                self.topl_add_tag.iconbitmap(GUI_PATH.replace(r"\interface", r"\icons\iconapp.ico"))
            else:
                logging.critical("The program has not been tested on this platform. Errors may occur")

        #This protocol prevents the universal bind to keep active after TopLevel close
        self.topl_add_tag.protocol("WM_DELETE_WINDOW", lambda: closeProtocol(self.topl_add_tag,"<Return>"))


    def del_tags(self, event = None):
        """
        This function opens a TopLevel interface so the user may select which custom tag will me deleted
        :param event:
        :return:
        """

        self.topl_listpopup = tk.Toplevel(bg="white")
        self.topl_listpopup.title("Excluir tags")


        self.list_box_tags = tk.Listbox(self.topl_listpopup,height=7,width=40, selectmode = "extended")
        self.list_box_tags.grid(padx=10,pady=10)

        for tags in self.TAGS:
            in_list = self.TAGS[tags][0] +" ("+self.TAGS[tags][1]+")"
            self.list_box_tags.insert("end",in_list)

        button_exclude_tag = ttk.Button(self.topl_listpopup,text = "Excluir",style="General.TButton")
        button_exclude_tag.grid(row=1,padx=10,pady=7,sticky="w")
        button_exclude_tag.bind("<ButtonRelease-1>", self.unset_tags)
        button_exclude_tag.bind_all("<Return>", self.unset_tags)

        self.topl_listpopup.resizable(width=False, height=False)

        #Grab the TopLevel prevents user from interacting with the main window, it prevents duplicates
        self.topl_listpopup.grab_set()

        #For some reason, del_tag toplevel doensn't grab the focus by defaul, so I do it manually
        self.topl_listpopup.focus_set()

        #Setting widith, heigh and coordinates for toplevel
        self.topl_listpopup.geometry(windowSizing(self.topl_listpopup,265))

        #Set the TopLevel icon
        try:
            # for freeze version
            if PLATFORM == "Linux":
                self.topl_listpopup.iconbitmap(sys._MEIPASS + "/icons/iconapp.ico")
            elif PLATFORM == "Windows":
                self.topl_listpopup.iconbitmap(sys._MEIPASS + r"\icons\iconapp.ico")
            else:
                logging.critical("The program has not been tested on this platform. Errors may occur")
        except:
            # for script version script
            if PLATFORM == "Linux":
                self.topl_listpopup.iconbitmap(GUI_PATH.replace("/interface", "/icons/iconapp.ico"))
            elif PLATFORM == "Windows":
                self.topl_listpopup.iconbitmap(GUI_PATH.replace(r"\interface", r"\icons\iconapp.ico"))
            else:
                logging.critical("The program has not been tested on this platform. Errors may occur")

        #This protocol prevents the universal bind to keep active after TopLevel close
        self.topl_listpopup.protocol("WM_DELETE_WINDOW", lambda: closeProtocol(self.topl_listpopup,"<Return>"))


    def unset_tags(self, event=None):
        """
        Excluir tags. Recupera o índice de seleção da lista e exclui as respectivas tags,
        então atualiza as tags e a lista.
        :param event:
        :return:
        """
        selection_tuple = self.list_box_tags.curselection()

        for tuple in selection_tuple:
            tuple+=1
            print(tuple)
            self.TAGS[tuple][2].grid_forget()
            self.TAGS[tuple][3].grid_forget()

            del self.TAGS[tuple]


        # Apagando a listbox
        self.list_box_tags.delete(0, "end")
        if len(self.TAGS) == 0: #Reiniciando tags
            self.topl_listpopup.unbind_all("<Return>")
            self.topl_listpopup.destroy()

            self.TAGS = {1: ["Nome", "<nome>", ttk.Label(self.tags, text="Nome", style="General.TLabel"),
                             ttk.Entry(self.tags, width=5, textvariable=self.name_strvar)],
                         2: ["CPF (opcional)", "<cpf>",
                             ttk.Label(self.tags, text="CPF (opcional)", style="General.TLabel"),
                             ttk.Entry(self.tags, width=5, textvariable=self.cpf_strvar)]}

        else:
            index = 1  # Normalizando o dicionário (ele fica "furado" depois das exclusões)
            new_tags = self.TAGS
            print(new_tags)
            self.TAGS = {}
            len(new_tags)
            for tag in new_tags:
                self.TAGS[index] = new_tags[tag]
                # reinserindo as tags na listbok
                in_list = self.TAGS[index][0] + " (" + self.TAGS[index][1] + ")"
                self.list_box_tags.insert("end", in_list)

                index +=1

        self.update_tags()


    def set_tags(self, event= None):
        """
        This function is responsible for defining new tags based on the inputs provided by the self.add_tag() function.
        :param event:
        :return:
        """

        logging.debug("Defining a new tag")
        #Getting the user input
        title = self.add_tag_entry1.get()
        tag = self.add_tag_entry2.get()

        #This String Var will be atributed to the tag. I prevents the user from inserting more than one letter as a sheet column
        strvar = tk.StringVar()
        strvar.trace("w",lambda a,b,c: self.entryTagCorrector(strvar))

        #Check if user left some of the entries empty
        if title == "":
            self.add_tag_warn1 = tk.Label(self.topl_add_tag,bg="white",text="Escolha um nome para a tag",fg = "red", font = ("Arian","8"))
            self.add_tag_warn1.grid(row=2,sticky="w", padx=10)
            logging.debug("A title for the tag hasn't been provided. The tag can not be created")
        elif tag =="":
            self.add_tag_warn2 = tk.Label(self.topl_add_tag,bg="white",text="Escolha um identificador",fg = "red", font = ("Arian","8"))
            self.add_tag_warn2.grid(row=4, sticky="w", padx=10)
            logging.debug("A tag for the tag hasn't been provided. The tag can not be created")

        #This block defines a new tag
        else:
            logging.debug("The entries are normal, proceeding the tag creation")
            len_tags = len(self.TAGS)
            if len_tags < 6:
                self.TAGS[len_tags+1] = [title,tag,ttk.Label(self.tags, text = title, style= "General.TLabel"),ttk.Entry(self.tags, width=5, textvariable=strvar)]
                logging.info("The tag has been sucefully added to the tag dictionary")

            logging.debug("Updating tags and wiping out the entries of the TopLevel")
            self.update_tags()
            self.add_tag_entry1.delete(0,"end")
            self.add_tag_entry1.focus_set()
            self.add_tag_entry2.delete(0,"end")
            self.topl_add_tag.update()

            if len(self.TAGS) == 6:
                self.topl_add_tag.unbind_all("<Return>")
                self.topl_add_tag.destroy()
                logging.info("All tag slots have been filled. The TopLevel has been destroyed")


    def genAttempt(self):
        """
        This function is responsible for processing user inputs and, if they were correct, starting the certificates generation
        :return:
        """
        logging.debug("Generating attempt: checking user inputs")
        go = 1
        self.TUPLES=[]

        while go == 1:
            #Checking dependencies
            deps = self.controller.depCheck()
            if deps == False:
                self.attemptFailed(error=7)
                go=0
                break

            # Chegando entrada de arquivos
            if self.PLAN_PATH == "":
                self.attemptFailed(error=0)
                go=0
                logging.info("A sheet has not been added. Aborting...")
                break
            elif self.FILE_PATH =="":
                self.attemptFailed(error=1)
                go=0
                logging.info("A teamplate has not been added. Aborting...")
                break
            elif self.SAVE_PATH =="":
                self.attemptFailed(error=6)
                go=0
                logging.info("A save directory has not been added. Aborting...")
                break

            logging.debug("Generating attempt: files and directories are filled")

            # testando colunas
            if self.TAGS[1][3].get() != "":
                col_name = self.TAGS[1][3].get()
                col_name = col_name.capitalize()
                if re.match("[A-Z]", col_name) == None or len(col_name) != 1:
                    self.attemptFailed(error=2)
                    go = 0
                    break
                else:
                    self.TUPLES.append((col_name,"<nome>"))
            else:
                self.attemptFailed(error=3)
                go = 0
                break

            if self.TAGS[2][3].get() != "":
                col_cpf = self.TAGS[2][3].get()
                col_cpf = col_cpf.capitalize()
                if re.match("[A-Z]", col_cpf) == None or len(col_cpf) != 1:
                    self.attemptFailed(error=4)
                    go = 0
                    break
                else:
                    self.TUPLES.append((col_cpf, "<cpf>"))

            #PARA ESCREVER NA PLANILHA O NOME DO ARQUIVO
            if self.top_checkbut_write_var.get() == 1:
                if re.match("[A-Z]",str(self.top_entry_write.get())):
                    self.WRITE_IN_CELL = str(self.top_entry_write.get())
                else:
                    self.attemptFailed(error = 8)
                    go = 0
                    break
            else: self.WRITE_IN_CELL = 0

            #para tags personalizadas

            for tag in self.TAGS:
                if tag >2:
                    col_tag = (self.TAGS[tag][3].get())
                    col_tag = col_tag.capitalize()

                    if re.match("[A-Z]", col_tag) != None and len(col_tag) == 1:
                        self.TUPLES.append((col_tag, self.TAGS[tag][1]))
                    else:
                        self.attemptFailed(error=5)
                        go = 0
                        break

            #GERANDO
            load = PyCertsSB()

            load.generateDocx(
                self.FILE_PATH, self.PLAN_PATH,self.SAVE_PATH, self.top_checkbut_header_var.get(),self.TUPLES, self.WRITE_IN_CELL
            )

            go = 0


    def attemptFailed(self,error):
        if error == 0:
            popup("Insira uma planilha de dados.",title="Erro")
        elif error == 1:
            popup("Insira um arquivo modelo para o certificado.",title="Erro")
        elif error == 2:
            popup("O valor de coluna inserido em 'Nome' não\né válido.",title="Erro")
        elif error == 3:
            popup("O preenchimento da coluna de nomes é obrigatório.",title="Erro")
        elif error == 4:
            popup("O valor de coluna inserido em 'CPF' não\né válido.", title="Erro")
        elif error == 5:
            popup("Um valor de coluna inserido para tags\npersonalizadas não é válido",title="Erro")
        elif error == 6:
            popup("Selecione uma pasta para salvar os certificados",title = "Erro")
        elif error == 7:
            popup("Essa função exige uma versão recente do LibreOffice ou Microsoft Word (ativado)\n"
                  "intalado na sua máquina. Se você já tem um desses programas instalados, verifique\n"
                  "o local de instalação em CONFIGURAÇÕES. Caso não, por favor, instale uma versão atualizada.",
                  title="Erro")
        elif error == 8:
            popup("Valor de coluna inserido para escrever nomes de arquivos inválido", title = "Erro")


    def checkbutWriteSelect(self):
        checkbut = self.top_entry_write

        if str(checkbut['state']) == "disabled":
            checkbut.configure(state = "normal")
        elif str(checkbut['state']) == "normal":
            checkbut.configure(state = "disabled")

class Email(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        #Definindo queues para a thread de envio de emails
        self.queue_progress = Queue()
        self.queue_data = Queue()

        #Variaveis de envio

        self.PLAN_PATH = ""
        self.ANNEX_PATH = ""
        self.COL_NAME = ""
        self.COL_EMAIL = ""
        self.COL_ANNEX = ""
        self.COL_TAGS = []
        self.TAGS = {}
        self.SUBJET = ""
        self.BODY = ""

        #Layout

        self.workspace1 = tk.Frame(self,background="white")
        self.workspace1.grid(sticky="w")

        #titulo

        self.label_top = ttk.Label(self.workspace1, text = "Envio de E-mails", style="Main.TLabel")
        self.label_top.grid(row=0, padx=10,pady=10, sticky="w")

        #inserir planilha

        self.label_workspace1 = ttk.Label(self.workspace1, text = "Planilha de dados",style="General.TLabel")
        self.label_workspace1.grid(row = 2, pady=20, padx=20, sticky="w")

        self.button_workspace1 = ttk.Button(self.workspace1, text = "Carregar", style="General.TButton")
        self.button_workspace1.grid(row = 2, column = 1, padx=0, sticky="e")
        self.button_workspace1.bind("<ButtonRelease-1>", self.fileinput)


        self.label_output_workspace1 = ttk.Entry(self.workspace1, width = 17, state = "readonly")
        self.label_output_workspace1.grid(row=2, column = 2, sticky="w", padx=10)

        self.checkbut_var = tk.IntVar()
        self.checkbutton_plan = ttk.Checkbutton(self.workspace1, text = "Primeira linha como cabeçalho",
                                                style="TCheckbutton",variable = self.checkbut_var)
        self.checkbutton_plan.grid(row=3,padx=20, column = 0)

        #criando um conteiner

        self.workspace2 = tk.Frame(self,bg="white")
        self.workspace2.grid(row=1,sticky="w", pady=10)

        #titulo

        self.label_top = ttk.Label(self.workspace2, text = "Dados das colunas",style="General.TLabel")
        self.label_top.grid(row=0, padx=10, sticky="w")


        #TAGS
        self.tags = tk.Frame(self, background = "white")
        self.tags.grid(row=4, sticky="w")

        #String vars
        #String vars processes user inputs in real time. I prevents the user from inserting more than one letter as a sheet column
        self.name_strvar = tk.StringVar()
        self.name_strvar.trace("w", lambda a, b, c: self.entryTagCorrector(self.name_strvar))
        self.email_strvar = tk.StringVar()
        self.email_strvar.trace("w", lambda a, b, c: self.entryTagCorrector(self.email_strvar))
        self.annex_strvar = tk.StringVar()
        self.annex_strvar.trace("w", lambda a, b, c: self.entryTagCorrector(self.annex_strvar))

        self.TAGS = {1:["Nome (opcional)", "<nome>",ttk.Label(self.tags, text = "Nome (opcional)", style= "General.TLabel"),ttk.Entry(self.tags, width=5, textvariable=self.name_strvar)],
                     2:["E-mail","<email>",ttk.Label(self.tags, text = "E-mail",style="General.TLabel"),ttk.Entry(self.tags, width=5, textvariable=self.email_strvar)],
                     3:["Anexo (opcional)","---",ttk.Label(self.tags, text = "Anexo (opcional)",style="General.TLabel"),ttk.Entry(self.tags, width=5, textvariable=self.annex_strvar)]}


        #gerando blocos personalizados
        ####
        #ADICIONAR SISTEMA DE ARMAZENAMENTO
        ####

        self.update_tags()

        #ADICIONAR/EXCLUIR TAGS

        self.tagbutton = tk.Frame(self, background = "white")
        self.tagbutton.grid(row=5,sticky="w")

        self.tagbutton_button_addtag = ttk.Button(self.tagbutton, style="General.TButton",width=15,
                                              text = "Adicionar tag", command = self.add_tag)
        self.tagbutton_button_addtag.grid(row=6, padx=20)

        self.tagbutton_button_deltag = ttk.Button(self.tagbutton, style="General.TButton",width=15,
                                              text = "Deletar tag", command = self.del_tags)
        self.tagbutton_button_deltag.grid(row=6,column=1, padx=10)

        #e-mail
        #criando um conteiner

        self.subject = tk.Frame(self,bg="white")
        self.subject.grid(row=6,sticky="w", pady=10)
        #titulo

        self.label_top_subject = ttk.Label(self.subject, text = "Corpo do e-mail", style="General.TLabel")
        self.label_top_subject.grid(row=6, padx=10, sticky="w")

        #assunto do email

        self.label_input_subject = ttk.Label(self.subject, text = "Assunto",style="General.TLabel")
        self.label_input_subject.grid(row = 8,column=0, padx=20, sticky="w")
        self.entry_name_subject = ttk.Entry(self.subject, width=40)
        self.entry_name_subject.grid(row=8, column=1)

        #corpo do emails
        #criando um conteiner

        self.body = tk.Frame(self,bg="white")
        self.body.grid(row=9,sticky="w")

        self.label_input_body = ttk.Label(self.body, text = "Corpo",style="General.TLabel")
        self.label_input_body.grid(row = 9,column=0, padx=20, sticky="w")
        self.entry_name_body = tk.Text(self.body, width=48, height =7)
        self.entry_name_body.grid(row=9, column=1)

        #botão de envio
        #criando um conteiner

        self.button_ready = tk.Frame(self,bg="white")
        self.button_ready.grid(row=10,sticky="w", pady=10)

        #inserir pasta de anexos

        self.button_annex = ttk.Button(self.button_ready, text = "Anexos",style="General.TButton",
                                       command=self.dirinput)
        self.button_annex.grid(row=10,column = 0,padx = 5, pady=10, sticky = "e")


        self.label_output_annex = ttk.Entry(self.button_ready, width = 17, state = "readonly")
        self.label_output_annex.grid(row=10, column = 1, sticky="w", padx=10)

        #Comboboox para extensões de anexo

        self.str_drop_var = tk.StringVar(self.button_ready)
        self.str_drop_var.set("")  # valor padrão

        filetypes = [".pdf",".png",".txt",".xlsx"]

        self.drop_annex = ttk.Combobox(self.button_ready, textvariable=self.str_drop_var, values=filetypes,width=5,state="readonly")
        self.drop_annex.grid(row=10,column=2)

        #espaçamento
        self.spacing = tk.Label(self.button_ready, text="",width=0,bg="white")
        self.spacing.grid(row=10, column=3, padx=5, pady=10)

        #enviar

        self.button_ready1 = ttk.Button(self.button_ready, text = "Enviar", style="Main.TButton",
                                       command=self.send_attempt)
        self.button_ready1.grid(row=10, column = 4, padx = 20, pady=10, sticky = "w")

        #voltar
        self.button_backtomenu = ttk.Button(self.button_ready, text = "Voltar", style="Main.TButton",
                                            command=lambda: self.controller.transition("Menu"))
        self.button_backtomenu.grid(row=10, column =5, padx = 0, pady=10, sticky = "w")

        #MARGEM LATERAL
        self.border = ttk.Label(self.button_ready,text="  ",style="General.TLabel")
        self.border.grid(row=10,column=6)


    def fileinput(self, event = None):
        path = filedialog()
        print(path)

        if path:
            try:
                load = emails.EmailsSBChecking()
                load.check_plan(path)

                self.PLAN_PATH = path


                self.label_output_workspace1["state"] = "normal"
                self.label_output_workspace1.delete(0, 'end')
                self.label_output_workspace1.insert("end", path.split("/")[-1])
                self.label_output_workspace1["state"] = "readonly"
            except:
                var = traceback.format_exc()
                print(var)
                popup("Selecione uma planilha válida. Formatos aceitos: .xlsx.\nOBS: Salve sua planilha com o formato mais recente de\nxlsx e me lembre de adicionar um suporte pra outras extenções.", title="Arquivo inválido")


    def dirinput(self, event=None):
        path = dirdialog()

        if path != ():
            self.ANNEX_PATH = path
            print("dirinput%s"%path)

            self.label_output_annex["state"] = "normal"
            self.label_output_annex.delete(0, 'end')
            self.label_output_annex.insert("end", path.split("/")[-1])
            self.label_output_annex["state"] = "readonly"


    def update_tags(self, event=None):
        #Inputs de colunas

        row=0

        # criando inputs personalizados

        if len(self.TAGS) > 0:
            for tags in self.TAGS:
                if tags > 3:
                    row=1
                label_ref = self.TAGS[tags][2]
                entry_ref = self.TAGS[tags][3]
                label_ref.grid(row=row, column=tags*2 - 2 -(6*row), pady=20, padx=20, sticky="w")
                entry_ref.grid(row=row, column=tags*2 - 1 - (6*row))

        self.controller.geometry(str(self.controller.winfo_reqwidth()) + "x" + str(self.controller.winfo_reqheight()))


    def add_tag(self, event = None):
        """
        This function opens a TopLevel interface so the user may insert custom tags to the certifies
        :param event:
        :return:
        """


        #CREATING A TOPLEVEL
        self.topl_add_tag = tk.Toplevel(bg="white")
        self.topl_add_tag.title("Adicione novas tags")

        #TopLevel layout
        top_label = tk.Label(self.topl_add_tag, bg="white",text="Adicione novas tags para personalizar o certificado.", font=("Arial",10,"bold"))
        top_label.grid(pady=10, padx=10, columnspan=3)

        #Name of the tag
        label1 = ttk.Label(self.topl_add_tag, text = "Título",style="General.TLabel")
        label1.grid(row=1,padx=10, sticky="w")

        self.add_tag_entry1 = ttk.Entry(self.topl_add_tag, width = 15)
        self.add_tag_entry1.grid(row=1, column=1, sticky="w")
        self.add_tag_entry1.focus_set()

        #"Tag" of the tag
        label2 = ttk.Label(self.topl_add_tag, text="Tag",style="General.TLabel")
        label2.grid(row=3, padx=10,sticky="w")

        self.add_tag_entry2 = ttk.Entry(self.topl_add_tag, width=15)
        self.add_tag_entry2.grid(row=3, column=1, sticky="w")


        B1 = ttk.Button(self.topl_add_tag, text="Pronto", style="Main.TButton",command = self.set_tags)
        B1.grid(row = 5, pady=20, padx=10, sticky="w")
        B1.bind_all("<Return>", self.set_tags)


        self.topl_add_tag.resizable(width=False, height=False)

        #Grab the TopLevel prevents user from interacting with the main window, it prevents duplicates
        self.topl_add_tag.grab_set()

        #Setting widith, heigh and coordinates for toplevel
        self.topl_add_tag.geometry(windowSizing(self.topl_add_tag))

        #Set the TopLevel icon
        try:
            # for freeze version
            if PLATFORM == "Linux":
                self.topl_add_tag.iconbitmap(sys._MEIPASS + "/icons/iconapp.ico")
            elif PLATFORM == "Windows":
                self.topl_add_tag.iconbitmap(sys._MEIPASS + r"\icons\iconapp.ico")
            else:
                logging.critical("The program has not been tested on this platform. Errors may occur")
        except:
            # for script version script
            if PLATFORM == "Linux":
                self.topl_add_tag.iconbitmap(GUI_PATH.replace("/interface", "/icons/iconapp.ico"))
            elif PLATFORM == "Windows":
                self.topl_add_tag.iconbitmap(GUI_PATH.replace(r"\interface", r"\icons\iconapp.ico"))
            else:
                logging.critical("The program has not been tested on this platform. Errors may occur")

        #This protocol prevents the universal bind to keep active after TopLevel close
        self.topl_add_tag.protocol("WM_DELETE_WINDOW", lambda: closeProtocol(self.topl_add_tag,"<Return>"))


    def del_tags(self, event = None):
        """
        This function opens a TopLevel interface so the user may select which custom tag will me deleted
        :param event:
        :return:
        """

        self.topl_listpopup = tk.Toplevel(bg="white")
        self.topl_listpopup.title("Excluir tags")


        self.list_box_tags = tk.Listbox(self.topl_listpopup,height=7,width=40, selectmode = "extended")
        self.list_box_tags.grid(padx=10,pady=10)

        for tags in self.TAGS:
            in_list = self.TAGS[tags][0] +" ("+self.TAGS[tags][1]+")"
            self.list_box_tags.insert("end",in_list)

        button_exclude_tag = ttk.Button(self.topl_listpopup,text = "Excluir",style="General.TButton")
        button_exclude_tag.grid(row=1,padx=10,pady=7,sticky="w")
        button_exclude_tag.bind("<ButtonRelease-1>", self.unset_tags)
        button_exclude_tag.bind_all("<Return>", self.unset_tags)

        self.topl_listpopup.resizable(width=False, height=False)

        #Grab the TopLevel prevents user from interacting with the main window, it prevents duplicates
        self.topl_listpopup.grab_set()

        #For some reason, del_tag toplevel doensn't grab the focus by defaul, so I do it manually
        self.topl_listpopup.focus_set()

        #Setting widith, heigh and coordinates for toplevel
        self.topl_listpopup.geometry(windowSizing(self.topl_listpopup,265))

        #Set the TopLevel icon
        try:
            # for freeze version
            if PLATFORM == "Linux":
                self.topl_listpopup.iconbitmap(sys._MEIPASS + "/icons/iconapp.ico")
            elif PLATFORM == "Windows":
                self.topl_listpopup.iconbitmap(sys._MEIPASS + r"\icons\iconapp.ico")
            else:
                logging.critical("The program has not been tested on this platform. Errors may occur")
        except:
            # for script version script
            if PLATFORM == "Linux":
                self.topl_listpopup.iconbitmap(GUI_PATH.replace("/interface", "/icons/iconapp.ico"))
            elif PLATFORM == "Windows":
                self.topl_listpopup.iconbitmap(GUI_PATH.replace(r"\interface", r"\icons\iconapp.ico"))
            else:
                logging.critical("The program has not been tested on this platform. Errors may occur")

        #This protocol prevents the universal bind to keep active after TopLevel close
        self.topl_listpopup.protocol("WM_DELETE_WINDOW", lambda: closeProtocol(self.topl_listpopup,"<Return>"))


    def unset_tags(self, event=None):
        """
        Excluir tags. Recupera o índice de seleção da lista e exclui as respectivas tags,
        então atualiza as tags e a lista.
        :param event:
        :return:
        """
        selection_tuple = self.list_box_tags.curselection()

        for tuple in selection_tuple:
            tuple+=1
            print(tuple)
            self.TAGS[tuple][2].grid_forget()
            self.TAGS[tuple][3].grid_forget()

            del self.TAGS[tuple]


        # Apagando a listbox
        self.list_box_tags.delete(0, "end")
        if len(self.TAGS) == 0: #Reiniciando tags
            self.topl_listpopup.unbind_all("<Return>")
            self.topl_listpopup.destroy()

            self.TAGS = {1: ["Nome", "<nome>", ttk.Label(self.tags, text="Nome", style="General.TLabel"),
                             ttk.Entry(self.tags, width=5, textvariable=self.name_strvar)],
                         2: ["CPF (opcional)", "<cpf>",
                             ttk.Label(self.tags, text="CPF (opcional)", style="General.TLabel"),
                             ttk.Entry(self.tags, width=5, textvariable=self.cpf_strvar)]}

        else:
            index = 1  # Normalizando o dicionário (ele fica "furado" depois das exclusões)
            new_tags = self.TAGS
            print(new_tags)
            self.TAGS = {}
            len(new_tags)
            for tag in new_tags:
                self.TAGS[index] = new_tags[tag]
                # reinserindo as tags na listbok
                in_list = self.TAGS[index][0] + " (" + self.TAGS[index][1] + ")"
                self.list_box_tags.insert("end", in_list)

                index +=1

        self.update_tags()


    def set_tags(self, event= None):
        """
        This function is responsible for defining new tags based on the inputs provided by the self.add_tag() function.
        :param event:
        :return:
        """

        logging.debug("Defining a new tag")
        #Getting the user input
        title = self.add_tag_entry1.get()
        tag = self.add_tag_entry2.get()

        #This String Var will be atributed to the tag. I prevents the user from inserting more than one letter as a sheet column
        strvar = tk.StringVar()
        strvar.trace("w",lambda a,b,c: self.entryTagCorrector(strvar))

        #Check if user left some of the entries empty
        if title == "":
            self.add_tag_warn1 = tk.Label(self.topl_add_tag,bg="white",text="Escolha um nome para a tag",fg = "red", font = ("Arian","8"))
            self.add_tag_warn1.grid(row=2,sticky="w", padx=10)
            logging.debug("A title for the tag hasn't been provided. The tag can not be created")
        elif tag =="":
            self.add_tag_warn2 = tk.Label(self.topl_add_tag,bg="white",text="Escolha um identificador",fg = "red", font = ("Arian","8"))
            self.add_tag_warn2.grid(row=4, sticky="w", padx=10)
            logging.debug("A tag for the tag hasn't been provided. The tag can not be created")

        #This block defines a new tag
        else:
            logging.debug("The entries are normal, proceeding the tag creation")
            len_tags = len(self.TAGS)
            if len_tags < 6:
                self.TAGS[len_tags+1] = [title,tag,ttk.Label(self.tags, text = title, style= "General.TLabel"),ttk.Entry(self.tags, width=5, textvariable=strvar)]
                logging.info("The tag has been sucefully added to the tag dictionary")

            logging.debug("Updating tags and wiping out the entries of the TopLevel")
            self.update_tags()
            self.add_tag_entry1.delete(0,"end")
            self.add_tag_entry1.focus_set()
            self.add_tag_entry2.delete(0,"end")
            self.topl_add_tag.update()

            if len(self.TAGS) == 6:
                self.topl_add_tag.unbind_all("<Return>")
                self.topl_add_tag.destroy()
                logging.info("All tag slots have been filled. The TopLevel has been destroyed")


    def entryTagCorrector(self, strvar):
        """This function defines the rules of entry in tag fields. It's only allowed a single capitalized caracter"""
        value = strvar.get()
        logging.debug("An input has been made on tag column field.")
        if len(value) > 1:
            value = value[0]
            strvar.set(value)
            logging.info("User tried to enter more than one character in a tag field. Wiping the field.")
        if re.match("[a-z]",value) != None or re.match("[A-Z]",value) != None:
            strvar.set(value.capitalize())
            logging.info("A new column for a tag value has been setted.")
        else:
            strvar.set("")
            logging.debug("User tried to enter a non alphabetical character. Wiping the field.")


    def send_attempt(self, event = None):
        """This function if very important. It checks the entries and calls the subsequent tasks."""
        logging.debug("The user has made a send e-mail attempt.")
        go = 1
        self.TUPLES=[]

        while go == 1:
            #Testing sheet
            try:
                load = emails.EmailsSBChecking()
                load.check_plan(self.PLAN_PATH)
            except:
                self.attemptFailed(error = 1)
                go = 0
                break

            logging.debug("Generating attempt: files and directories are filled")

            #STATIC TAGS
            if self.TAGS[1][3].get() != "":
                col_name = self.TAGS[1][3].get()
                self.TUPLES.append((col_name, "<nome>"))
                logging.debug("Names column: {}.".format(col_name))


            if self.TAGS[2][3].get() != "":
                col_email = self.TAGS[2][3].get()
                self.COL_EMAIL = col_email
                logging.debug("E-mails column: {}.".format(col_email))
            else:
                self.attemptFailed(error=2)
                go = 0
                logging.info("A column for e-mails has not been added. The task may not proceed. Aborting...")
                break

            #Eu deveria dar uma olhada nessa parte do código depois
            if self.TAGS[3][3].get() != "" or self.ANNEX_PATH != "":
                if self.TAGS[3][3].get() != "" and self.ANNEX_PATH == "":
                    self.attemptFailed(error=3)
                    go = 0
                    break
                elif self.TAGS[3][3].get() != "" and self.ANNEX_PATH != "":
                    self.COL_ANNEX = self.TAGS[3][3].get()
            else:
                self.COL_ANNEX = ""


            #CUSTOM TAGS
            logging.debug("Checking entries for custom tags...")
            print(self.TAGS)
            for tag in self.TAGS:
                print(tag)
                if tag >3: #Picking only custom (x > 3) tags
                    col_tag = (self.TAGS[tag][3].get())
                    col_tag = col_tag.capitalize()
                    self.TUPLES.append((col_tag, self.TAGS[tag][1]))
                    logging.debug("Custom tag: {}, {}".format(col_tag, tag))


            #Número de emails a serem enviados (para tempo de execução)

            self.MOUNT = load.check_plan(self.PLAN_PATH,self.COL_EMAIL,self.checkbut_var.get())


            self.SUBJET =  self.entry_name_subject.get()

            self.BODY = self.entry_name_body.get("1.0",'end-1c')


            #iniciando a janela de progesso
            self.checkWindow()


            go = 0
            break


    def emailTaskCheck(self):
        if self.p1.is_alive():
            if not self.queue_data.empty() and not self.queue_progress.empty():
                data = self.queue_data.get(block=False)

                self.activity_preview_toaddr["state"] = "normal"
                self.activity_preview_toaddr.delete(0, 'end')
                self.activity_preview_toaddr.insert("end", data[0])
                self.activity_preview_toaddr["state"] = "readonly"

                self.activity_preview_subject["state"] = "normal"
                self.activity_preview_subject.delete(0, 'end')
                self.activity_preview_subject.insert("end", data[1])
                self.activity_preview_subject["state"] = "readonly"

                self.activity_preview_body["state"] = "normal"
                self.activity_preview_body.delete(1.0, "end")
                self.activity_preview_body.insert("end", data[2])
                self.activity_preview_body["state"] = "disabled"

                if len(data)==4:
                    self.activity_preview_annex["state"] = "normal"
                    self.activity_preview_annex.delete(0, 'end')
                    self.activity_preview_annex.insert("end", data[3])
                    self.activity_preview_annex["state"] = "readonly"

                progress = self.queue_progress.get(block=False)
                self.activity_progressbar["value"] = progress


            self.activity.update()
            self.activity.after(500, self.emailTaskCheck)
        else:
            self.p1.terminate()
            self.activity_progressbar["value"] = 100
            self.activity_label_title['text'] = "Revisão"
            self.activity_preview_title['text'] = "Terminado"


    def attemptFailed(self, error):
        if error == 1:
            popup("Insira uma planilha contendo pelo\nmenos e-mails.",title="Erro")
        elif error == 2:
            popup("A coluna de e-mails não foi especificada", title="Erro")
        elif error == 3:
            popup("Uma coluna de anexos foi inserida porém não foi especicado a pasta\nonde encontram-se esses anexos, faça isso no botão no fim da página.", title="Erro")
        elif error == 4:
            popup("Existem valores inválidos nas colunas de\n tags personalizadas.", title = "Erro")
        elif error == 5:
            popup("Você adicionou um diretório de anexos mas\nnão especificou a coluna de anexos.", title="Erro")
        elif error == 6:
            popup("Você definiu uma coluna de anexos mas não\n especificou um diretório com os anexos a\n serem enviados.",title="Erro")


    def checkWindow(self):
        def ready():
            checkpoint.destroy()
            self.activityWindow()

        def not_ready():
            checkpoint.destroy()


        if self.SUBJET == "" or self.BODY == "":
            if self.SUBJET == "" and self.BODY == "":
                is_empty="corpo de e-mail e assunto"
            elif self.SUBJET == "":
                is_empty = "assunto"
            else: is_empty = "corpo de e-mail"

            checkpoint = tk.Toplevel(bg="white")

            check = tk.Frame(checkpoint,bg="white")
            check.grid()

            label_msg = "O(s) campo(s) %s parecem estar vazios.\nDeseja prosseguir mesmo assim?" % is_empty
            check_label = ttk.Label(check, text = label_msg, justify="left",style="General.TLabel")
            check_label.grid(columnspan=4,pady=10, padx=10)

            button_continue = ttk.Button(check, text = "Continuar", style="Main.TButton",command = ready)
            button_continue.grid(row =1,sticky="w", pady=20, padx=10)

            button_stop = ttk.Button(check, text = "Sair", style="Main.TButton",command = not_ready)
            button_stop.grid(row =1,column=1,sticky = "w",pady=20, padx=10)


            checkpoint.title("Aviso")
            checkpoint.grab_set()
            checkpoint.geometry(windowSizing(checkpoint))
        else:
            #iniciando o processo para envio de emails
            self.activityWindow()


    def activityWindow(self):
        global fromaddr, pwd


        def activityOnCreate():

            prev = emails.EmailsSBChecking()
            #The preview() function return the address, the subject and the body of the first e-mail that will be send
            data = prev.preview(
                header=self.checkbut_var.get(),
                plan_dir=self.PLAN_PATH,
                subject=self.SUBJET,
                body=self.BODY,
                col_email=self.COL_EMAIL,
                col_annex=self.COL_ANNEX,
                annex_folder=self.ANNEX_PATH,
                annex_ext=self.str_drop_var.get(),
                col_args=self.TUPLES
            )

            self.activity_preview_toaddr["state"] = "normal"
            self.activity_preview_toaddr.delete(0, 'end')
            self.activity_preview_toaddr.insert("end", data[0])
            self.activity_preview_toaddr["state"] = "readonly"

            self.activity_preview_subject["state"] = "normal"
            self.activity_preview_subject.delete(0, 'end')
            self.activity_preview_subject.insert("end", data[1])
            self.activity_preview_subject["state"] = "readonly"

            self.activity_preview_body["state"] = "normal"
            self.activity_preview_body.delete(1.0, "end")
            self.activity_preview_body.insert("end", data[2])
            self.activity_preview_body["state"] = "disabled"

            if len(data) == 4:
                self.activity_preview_annex["state"] = "normal"
                self.activity_preview_annex.delete(0, 'end')
                self.activity_preview_annex.insert("end", data[3])
                if data[3].find("ERRO") != -1:
                    self.activity_preview_annex.config(foreground= "red")
                self.activity_preview_annex["state"] = "readonly"


        def activityStart():
            self.p1 = Process(target=emails.EmailsSB, args=(self.queue_data,
                                                             self.queue_progress,
                                                             self.PLAN_PATH,
                                                             fromaddr,
                                                             pwd,
                                                             self.SUBJET,
                                                             self.BODY,
                                                             self.COL_EMAIL,
                                                             self.checkbut_var.get(),
                                                             self.COL_ANNEX,
                                                             self.ANNEX_PATH,
                                                             self.str_drop_var.get(),
                                                             self.TUPLES,))
            freeze_support()
            self.p1.start()

            self.activity_button_stop['state']="normal"


            self.p1_time = -1
            self.activity_panel_time_spent['text'] = "Tempo de execução: 0 segundos"



            activityClock()

            self.activity_button_start['state'] = "disabled"
            self.activity_label_title['text'] = "Envio"
            self.activity_preview_title['text'] = "Enviando e-mail para..."

            #monitorando o fluxo de informação do processo para a GUI e atualizando a barra de progresso
            self.emailTaskCheck()

            self.COL_NAME = ""
            self.COL_EMAIL = ""
            self.COL_ANNEX = ""
            self.SUBJET = ""
            self.BODY = ""

        def activityStop():
            self.p1.terminate()

            self.activity_label_title['text'] = "Revisão"
            self.activity_preview_title['text'] = "Parado"

            self.activity_button_stop['state'] = "disabled"

        def activityClose():
            if "self.p1" in locals():
                if self.p1.is_alive():
                    activityStop()
            self.activity.destroy()

        def activityClock():
            self.p1_time += 1
            if self.p1.is_alive():
                self.activity.after(1000, activityClock)
                self.activity_panel_time_spent["text"] = "Tempo de execução: %i segundos" % self.p1_time
                self.activity_panel_time_spent.update()
            else:
                self.p1_time += 1


        self.activity = tk.Toplevel(bg="white")


        self.activity.title("Enviando e-mails")

        self.activity_title = tk.Frame(self.activity,bg="white")
        self.activity_title.grid(sticky="w")

        self.activity_label_title = ttk.Label(self.activity_title, text ="Revisão", style="Main.TLabel")
        self.activity_label_title.grid(sticky="w",padx=10,pady=20)

        ##Preview
        self.activity_preview = tk.Frame(self.activity,bg="white")
        self.activity_preview.grid(row=1,pady=20, padx=10)

        self.activity_preview_title = tk.Label(self.activity_preview, bg="white",text="Pronto para enviar",font=("arial","11","italic"))
        self.activity_preview_title.grid(sticky="w")

        self.activity_preview_toaddr = ttk.Entry(self.activity_preview, width=40, state="readonly")
        self.activity_preview_toaddr.grid(row=1,sticky="w")

        self.activity_preview_subject = ttk.Entry(self.activity_preview,width=40, state = "readonly")
        self.activity_preview_subject.grid(row=2,sticky="w")

        self.activity_preview_body = tk.Text(self.activity_preview, width=40,height=17, state = "disabled")
        self.activity_preview_body.grid(row=3, sticky="w")

        self.activity_preview_annex = ttk.Entry(self.activity_preview, width=40, state = "readonly")
        self.activity_preview_annex.grid(row=4, sticky="w")


        #Dados de execução

        self.activity_panel_frame = tk.Frame(self.activity,bg="white")
        self.activity_panel_frame.grid(column=1,row=1,sticky="nw",pady=40)

        self.activity_panel_time_expected = ttk.Label(self.activity_panel_frame,style="General.TLabel")
        self.activity_panel_time_expected["text"] = "Tempo estimado: %i segundos." % int(self.MOUNT*1.7)
        self.activity_panel_time_expected.grid(sticky="w")

        self.activity_panel_time_spent = ttk.Label(self.activity_panel_frame,style="General.TLabel", text ="")
        self.activity_panel_time_spent.grid(row=1,pady = 10,sticky="w")


        #Barra de progresso

        self.activity_progress_frame = tk.Frame(self.activity,bg="white")
        self.activity_progress_frame.grid(column=0,row=2,columnspan=2,sticky="nw")

        self.activity_progressbar = ttk.Progressbar(self.activity_progress_frame, orient="horizontal", length=300, mode='determinate')
        self.activity_progressbar.grid(row=0, padx=160,pady=10)

        #Botões

        self.activity_buttons_frame = tk.Frame(self.activity,bg="white")
        self.activity_buttons_frame.grid(row=3,sticky="nw",columnspan=7)

        self.activity_button_start = ttk.Button(self.activity_buttons_frame, text="Enviar",
                                                style="Main.TButton",command=activityStart)
        self.activity_button_start["command"] = None
        self.activity_button_start.grid(row=1, sticky= "w", pady=10, padx=10)

        self.activity_button_stop = ttk.Button(self.activity_buttons_frame, style="Main.TButton",text="Parar", state="disabled", command=activityStop)
        self.activity_button_stop.grid(row=1,column=1, pady=10, padx=10)

        self.activity_button_close = ttk.Button(self.activity_buttons_frame,style="Main.TButton", text="Sair", command=activityClose)
        self.activity_button_close.grid(row=1,column=6, pady=10, padx=310)

        #Configuração de janela

        self.activity.protocol("WM_DELETE_WINDOW",activityClose)
        self.activity.resizable(width=True, height=True)
        self.activity.geometry(windowSizing(self.activity,640,590))
        self.activity.grab_set()

        #Adicionando preview
        activityOnCreate()


class Configuracoes(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        self.frame_title = tk.Frame(self, bg="white")
        self.frame_title.grid()

        self.label_title = ttk.Label(self.frame_title,text="Configurações",style="Main.TLabel")
        self.label_title.grid(sticky="w",padx=10)

        self.mainlabel_conversor = ttk.Label(self.frame_title, text = "Conversor", style = "Main.TLabel")
        self.mainlabel_conversor.grid(row = 1, pady=10, padx=20, sticky="w")

        self.label_conversor = ttk.Label(self.frame_title, text = "Conversor padrão", style = "General.TLabel")
        self.label_conversor.grid(row = 2, pady=10, padx=20, sticky="w")

        self.button_conversor = ttk.Button(self.frame_title, text = "Carregar", style="General.TButton",
                                            command = self.selectConverter)
        self.button_conversor.grid(row = 2, column = 1, padx=0, sticky="e")

        self.entry_showconversor = ttk.Entry(self.frame_title, width = 43, state = "readonly")
        self.entry_showconversor.grid(row=2, column = 2, sticky="w", padx=10)

        #BOTÕES

        self.buttons = tk.Frame(self, background = "white")
        self.buttons.grid(row=7, padx=10,pady=20, sticky="w")

        self.buttons_backtomenu = ttk.Button(self.buttons, text="Voltar", style = "Main.TButton",
                                         command=lambda: self.controller.transition("Menu"))
        self.buttons_backtomenu.grid(row=0, column=0, padx=10,sticky="e")


    def selectConverter(self):
        path = filedialog()

        if path:
            if platform.system() == "Windows":
                    result = self.controller.depCheck(path)
                    if result != False:
                        self.entry_showconversor["state"] = "normal"
                        self.entry_showconversor.delete(0, 'end')
                        self.entry_showconversor.insert("end", path)
                        self.entry_showconversor["state"] = "readonly"

def filedialog():

    options = {}

    options["title"] = "Selecione o arquivo"

    options["filetypes"] = {("Todos os arquivos", "*.*"),
                ("Planilhas", "*.csv *.xlsx *xls *.ods"),("Documentos","*.docx *.doc *.odt")}

    return askopenfilename(**options)


def dirdialog():

    options = {}

    options["title"] = "Selecione o diretório"

    return askdirectory(**options)


def popup(msg, width = None, height = None, title = "!"):
    """
    This is a generic function which displays a warning to the user by opening up a popup message
    :param msg: The text that will be displayed in the popup
    :param width: A specific width may be set. If None, the popup will be sizes automatically
    :param height: A specific height may be set. If None, the popup will be sizes automatically
    :param title: A specific title may be set.
    :return:
    """
    popup = tk.Toplevel()
    popup.title(title)
    popup.configure(bg="white")

    label = ttk.Label(popup, text=msg, style= "General.TLabel",justify = "left")
    label.grid(pady=10, padx = 10)
    B1 = ttk.Button(popup, style = "Main.TButton",text="OK", command = popup.destroy)
    B1.grid(sticky = "nw", pady = 10, padx = 10)

    # Setting widith, heigh and coordinates for toplevel
    popup.geometry(windowSizing(popup, width,height))

    popup.resizable(width=False, height=False)

    # Grab the TopLevel prevents user from interacting with the main window, it prevents duplicates
    popup.grab_set()

    #Somehow, the toplevel does not grab the focus by itself, so I do it manually
    popup.focus_set()

    # Set the TopLevel icon
    try:
        # for freeze version
        if PLATFORM == "Linux":
            popup.iconbitmap(sys._MEIPASS + "/icons/iconapp.ico")
        elif PLATFORM == "Windows":
            popup.iconbitmap(sys._MEIPASS + r"\icons\iconapp.ico")
        else:
            logging.critical("The program has not been tested on this platform. Errors may occur")
    except:
        # for script version script
        if PLATFORM == "Linux":
            popup.iconbitmap(GUI_PATH.replace("/interface", "/icons/iconapp.ico"))
        elif PLATFORM == "Windows":
            popup.iconbitmap(GUI_PATH.replace(r"\interface", r"\icons\iconapp.ico"))
        else:
            logging.critical("The program has not been tested on this platform. Errors may occur")


def closeProtocol(tk,bind=None):
    if bind != None:
        try:
            tk.unbind_all(bind)
        except: raise SyntaxError("This bind does not exist")

    tk.destroy()


def windowSizing(tk, width = None,height = None):
    """
    Posiciona o widget inserido no centro da tela considerando o tamanho mínimo de tela.
    Retorna string no padrão {}x{}+{}+{} para método geometry()
    :param tk: widget ou tk.Toplevel() associado.
    :param width: largura definida pelo usuário.
    :param height: altura definida pelo usuário.
    :return:
    """

    screen_width = int(tk.winfo_screenwidth())
    screen_height = int(tk.winfo_screenheight())


    # Posicionando tela
    if width == None and height == None:
        position = "+" + str(
            round((screen_width / 2) - (tk.winfo_reqwidth() / 2))) + "+" + str(
            round((screen_height / 2) - (tk.winfo_reqheight() / 2)))
        return position

    elif width != None and height != None:
        window_width = width
        window_height = height
    elif width != None:
        window_width = width
        window_height = tk.winfo_reqheight()
    else:
        window_height = height
        window_width = tk.winfo_reqwidth()




    position = str(str(window_width) + "x" + str(window_height) + "+" + str(
        round((screen_width / 2) - (window_width / 2))) + "+" + str(
        round((screen_height / 2) - (window_height / 2))))

    return position

def runMain(app_version):
    global GUI_PATH, APP_VERSION, PLATFORM
    GUI_PATH = os.path.dirname(os.path.abspath(__file__))
    APP_VERSION = app_version
    PLATFORM = platform.system()

    if platform.system() == "Windows":
        import comtypes.client

    root = Initialize(app_version)

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    root.configure(bg="white")

    try:
        # for freeze version
        if PLATFORM == "Linux":
            root.iconbitmap(sys._MEIPASS + "/icons/iconapp.ico")
        elif PLATFORM == "Windows":
            root.iconbitmap(sys._MEIPASS + r"\icons\iconapp.ico")
        else:
            logging.critical("The program has not been tested on this platform. Errors may occur")
    except:
        # for script version script
        if PLATFORM == "Linux":
            root.iconbitmap(GUI_PATH.replace("/interface", "/icons/iconapp.ico"))
        elif PLATFORM == "Windows":
            root.iconbitmap(GUI_PATH.replace(r"\interface", r"\icons\iconapp.ico"))
        else:
            logging.critical("The program has not been tested on this platform. Errors may occur")

    root.title("XX Semana Acadêmica da Biologia")
    root.resizable(width=True, height=True)

    root.mainloop()

if __name__ == "__main__":
    # Logging configuration
    logging.basicConfig(level=logging.DEBUG, format='%(process)d-%(levelname)s-%(message)s')

    runMain("debug version")