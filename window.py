#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from tkinter import *
from tkinter import filedialog
from carimbador_automatico_refat import stamp, adicionar_pagina, mesclar_pdf, numerar_paginas
from pathlib import Path
from typing import Union, Literal, List
from PyPDF2 import PdfWriter, PdfReader, PdfMerger
import glob, os, fitz
import sys


def btn_clicked():
    print("Button Clicked")

# Função que abre janela para selecionar diretório
def selecionar_diretorio():
    global diretorio_selecionado
    entry1.delete("1.0", END)
    diretorio_selecionado = filedialog.askdirectory()
    if diretorio_selecionado:
        #label_diretorio.config(text=f"Diretório Selecionado: {diretorio_selecionado}")
        entry1.insert(END, f'{diretorio_selecionado}')
    else:
        #label_diretorio.config(text="Nenhum diretório selecionado.")
        entry1.insert(END, f'Nenhum diretório selecionado.')
        diretorio_selecionado = ''

        
# Função que abre janela contendo instruções sobre o programa      
def abrir_janela_sobre():
    janela_sobre = Toplevel(window)
    janela_sobre.geometry("650x400")
    janela_sobre.title("Sobre o programa")
    janela_sobre.configure(bg='#b0e0e6')
    janela_sobre.resizable(False, False)

    # Adiciona texto à caixa de texto
    texto = """
    Versão 1.0
    Este programa permite que a partir de um conjunto de arquivos em formato PDF seja gerado um arquivo PDF final
    contendo todas as folhas carimbadas frente e verso e numeradas, caso o usuário deseje. O programa insere 
    carimbo na parte frontal de cada página, numera as  páginas, se desejado, e insere carimbo "em branco" no 
    verso de cada página restando apenas ao usuário imprimir o documento final em modo FRENTE E VERSO. Ao final 
    serão gerados dois arquivos em pdf: "Arquivo_pronto.pdf" (conterá todas as páginas carimbadas e numeradas) e
    "Arquivo_sem_numeracao.pdf" (conterá apenas as páginas carimbadas). 
    Para que as páginas sejam numeradas basta que o usuário preencha o campo "página inicial para numeração" 
    contendo o número a ser inserido na primeira página. Caso, esse campo seja deixado em branco, as páginas não 
    serão numeradas, mas apenas carimbadas. É importante destacar que ao final será gerado um arquivo em pdf 
    (Arquivo_final.pdf) que deverá ser impresso colorido e no modo modo FRENTE E VERSO. 
    Para gerar o documento carimbado e assinado siga os seguintes passos:
    1- Criar um diretório apenas com os documentos em pdf que se deseja carimbar. 
    2- Ordenar os arquivos do diretório inserindo uma numeração no nome deles 
    (Ex: 1.DIEx nr xyz, 2.Ofício nr abc, 3.Parecer técnico nr 2, 4. BI nr 234 etc)
    3- Clicar  em "Clique aqui para selecionar" e escolha o diretório que contenha os arquivos a serem carimbados.
    4- Se desejar numerá-los, inserir o número da primeira página no campo "Página inicial para numeração".
    5- Clique em "Carimbar".
    6- Serão gerados 2 arquivos em pdf na mesma pasta onde se encontra o arquivo executável.
    7- Clique em fechar para encerrar o programa.
    
    Para informar bugs ou retirar dúvidas, envie um e-mal para horstmann.alexandre@eb.mil.br
    """        
    
    label = Label(janela_sobre, text=texto, fg='black', bg='#b0e0e6', justify='left')
    label.grid(row=0, column=0, rowspan=4, padx=10, pady=10, sticky="NSEW")

    botao_fechar = Button(janela_sobre, text="Fechar", bg='#b0e0e6', command=janela_sobre.destroy)
    botao_fechar.grid(row=5, column=0, padx=10, pady=10, sticky="NSEW")
    
    
def carimbar():
    global diretorio_selecionado, entry0, entry1, entry2
    entry2.delete("1.0", END)  # Apaga a área de texto
    entry2.update()
    # Obtem lista de todos os arquivos pdf no diretório
    arquivos_pdf = glob.glob(diretorio_selecionado + "/*.pdf")
    if diretorio_selecionado:
        # Mescla os arquivos
        entry2.insert(END,'Mesclando arquivos...')
        entry2.update()
        mesclar_pdf(arquivos_pdf)
        # Carimba os arquivos
        entry2.delete("1.0", END)  # Apaga a área de texto
        entry2.insert(END,'Carimbando arquivos...')
        entry2.update()
        if checkbox_var.get():
            arquivo_carimbado = stamp('Arquivos_mesclados.pdf', 'CARIMBO_RESERVADO.pdf', 'Arquivo_carimbado.pdf')
        else:
            arquivo_carimbado = stamp('Arquivos_mesclados.pdf', 'CARIMBO.pdf', 'Arquivo_carimbado.pdf')
        # Adiciona páginas em branco
        entry2.delete("1.0", END)  # Apaga a área de texto
        entry2.insert(END,'Adicinando carimbo "em branco"...')
        entry2.update()
        arquivo_sem_numeracao = adicionar_pagina(arquivo_carimbado)
        if entry0.get():
            # Insere numeração nas páginas
            entry2.delete("1.0", END)  # Apaga a área de texto
            entry2.insert(END,'Numerando páginas...')
            entry2.update()
            numerar_paginas('Arquivo_sem_numeracao.pdf', int(entry0.get()))
        # Deleta arquivos desnecessários
        try:
            os.remove('Arquivos_mesclados.pdf')
            print("Arquivo deletado com sucesso!")
        except FileNotFoundError:
            print("O arquivo não foi encontrado.")
        except PermissionError:
            print("Você não possui permissão para deletar este arquivo.")
        print('Arquivo pronto!')
        entry2.delete("1.0", END)  # Apaga a área de texto
        entry2.insert(END,'Arquivo pronto!')
        entry2.update()
    else:
        entry2.insert(END,'Operação não realizada. É necessário selecionar um diretório!')
        entry2.update()


# Salva a saída padrão e a saída de erro em arquivos
original_stdout = sys.stdout
original_stderr = sys.stderr
sys.stdout = open('output.txt', 'w')
sys.stderr = open('error.txt', 'w')

    
window = Tk()
window.title("Carimbador Automático de Documentos")
window.geometry("750x420")
window.configure(bg = "#ffffff")
canvas = Canvas(window, bg = "#ffffff", height = 420, width = 750, bd = 0, highlightthickness = 0, relief = "ridge")
canvas.place(x = 0, y = 0)

# Adicionando Checkbox
checkbox_var = IntVar()
checkbox = Checkbutton(window, bg="#ffffff", text="Inserir carimbo de documento reservado", variable=checkbox_var)
checkbox.place(x=57, y=315)

background_img = PhotoImage(file = f"background.png")
background = canvas.create_image(375.0, 211.0, image=background_img)

# Numeração inicial a ser carimbada
entry0_img = PhotoImage(file = f"img_textBox0.png")
entry0_bg = canvas.create_image(195.0, 295.0, image = entry0_img)
entry0 = Entry(bd = 0, bg = "#fffdfd", highlightthickness = 0)
entry0.place(x = 57, y = 279, width = 276, height = 30)

# Botão "Fechar Janela"
img0 = PhotoImage(file = f"img0.png")
b0 = Button(image = img0, borderwidth = 0, highlightthickness = 0, command = window.destroy, relief = "flat")
b0.place(x = 562, y = 310, width = 157, height = 33)

# Botão "selecionar diretorio"
diretorio_selecionado = ''
img1 = PhotoImage(file = f"img1.png")
b1 = Button(image = img1, borderwidth = 0, highlightthickness = 0, command = selecionar_diretorio, relief = "flat")
b1.place(x = 526, y = 149, width = 193, height = 31)

# Botão "sobre o programa"
img2 = PhotoImage(file = f"img2.png")
b2 = Button(image = img2, borderwidth = 0, highlightthickness = 0, command = abrir_janela_sobre, relief = "flat")
b2.place(x = 562, y = 372, width = 157, height = 31)

# Botão "Carimbar"
img3 = PhotoImage(file = f"img3.png")
b3 = Button(image = img3, borderwidth = 0, highlightthickness = 0, command = carimbar, relief = "flat")
b3.place(x = 562, y = 250, width = 157, height = 31)

# Indica diretório selecionado
entry1_img = PhotoImage(file = f"img_textBox1.png")
entry1_bg = canvas.create_image(378.0, 205.0, image = entry1_img)
entry1 = Text(bd = 0, bg = "#fffdfd", highlightthickness = 0)
entry1.place(x = 37, y = 189, width = 682, height = 30)

# Mensagem de "arquivo pronto!"
entry2_img = PhotoImage(file = f"img_textBox2.png")
entry2_bg = canvas.create_image(191.0, 427.0, image = entry2_img)
entry2 = Text(bd = 0, bg = "#b0e0e6", highlightthickness = 0)
entry2.place(x = 37, y = 345, width = 500, height = 30)

window.resizable(False, False)
window.mainloop()

# Restaura a saída padrão e a saída de erro
sys.stdout.close()
sys.stderr.close()
sys.stdout = original_stdout
sys.stderr = original_stderr


# In[ ]:





# In[ ]:





# In[ ]:




