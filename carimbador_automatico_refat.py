#!/usr/bin/env python
# coding: utf-8

# In[1]:


from pathlib import Path
from typing import Union, Literal, List
from PyPDF2 import PdfWriter, PdfReader, PdfMerger
import glob, os, fitz
#from memory_profiler import profile


# Função para incluir carimbo no PDF extraída de https://pypdf2.readthedocs.io/en/latest/user/add-watermark.html.
# O carimbo foi preparado no paint, inserido e posicionado num documento do word, que foi exportado para PDF 
#@profile
def stamp(
    content_pdf: Path,
    stamp_pdf: Path,
    pdf_result: Path,
    page_indices: Union[Literal["ALL"], List[int]] = "ALL",
):
    reader = PdfReader(stamp_pdf)  # lê o arquivo com o carimbo
    image_page = reader.pages[0]  # pega a página contendo o carimbo
    writer = PdfWriter()  # Cria um objeto 
    reader = PdfReader(content_pdf)
    if page_indices == "ALL":
        page_indices = list(range(0, len(reader.pages)))
    for index in page_indices:
        content_page = reader.pages[index]
        mediabox = content_page.mediabox
        content_page.merge_page(image_page)
        content_page.mediabox = mediabox
        writer.add_page(content_page)
    
    #with open(pdf_result, "wb") as fp:
    #    writer.write(fp)
    return writer
    
    
# Função que adiciona uma página (página em branco) 
#@profile
def adicionar_pagina(arquivo):
    novo_arquivo = PdfWriter()  # cria novo arquivo pdf
    #arquivo_pdf = PdfReader(arquivo)
    pagina_em_branco = PdfReader('em_branco.pdf')
    #for pagina in arquivo_pdf.pages:
    for pagina in arquivo.pages:
        novo_arquivo.add_page(pagina)
        novo_arquivo.add_page(pagina_em_branco.pages[0])
    
    # Salva cada página num arquivo pdf dentro do diretório páginas
    with Path('Arquivo_sem_numeracao.pdf').open(mode="wb") as arquivo:
        novo_arquivo.write(arquivo)
    

# Função que mescla os arquivos pdf
#@profile
def mesclar_pdf(lista_arquivos_pdf):
    pdf_mesclado = PdfMerger()  # Cria um objeto para os pdfs mesclados
    for i, pdf_file in enumerate(lista_arquivos_pdf):
        if i == 0:
            pdf_mesclado.append(pdf_file)
        else:
            pdf_mesclado.append(pdf_file)
    # Salva os arquivos pdfs mesclados num arquivo
    with Path(f"Arquivos_mesclados.pdf").open(mode="wb") as arquivo:
        pdf_mesclado.write(arquivo)
    pdf_mesclado.close()
    

# Função para inserir números nas páginas
#@profile
def numerar_paginas(pdf_path, nr_inicial):
    pdf = fitz.open(pdf_path)
    nr_inicial = int(nr_inicial)
    for page_number, page in enumerate(pdf):
        if page_number == 0:
            # Obtém as dimensões da página
            width, height = page.rect.width, page.rect.height
            # Define as coordenadas para o canto superior direito
            x, y = width - 58, 30 
        if page_number%2 == 0:
            # Define as propriedades do texto
            text_properties = {"text": str(nr_inicial), "fontname": "helv", "fontsize": 12, "color": (0, 0, 1)}
            # Adiciona o texto à página
            page.insert_text((x, y), **text_properties)
            nr_inicial += 1

    # Salva as alterações
    pdf.save('Arquivo_pronto.pdf')
    pdf.close()

''' 
#carimbo = 'C:\\Users\\horstmann\\Downloads\\carimbador_automático\\CARIMBO.pdf'
#carimbo = 'C:\\Users\\Alexandre\\Dropbox\\Cursos\\Python\\Aplicações\\carimbador_automático\\CARIMBO.pdf'

# Diretório para buscar os arquivos pdf
#diretorio = 'C:\\Users\\horstmann\\Downloads\\carimbador_automático\\arquivos_pdf'
#diretorio = 'C:\\Users\\Alexandre\\Dropbox\\Cursos\\Python\\Aplicações\\carimbador_automático\\arquivos_pdf'

# Obtem lista de todos os arquivos pdf no diretório
arquivos_pdf = glob.glob(diretorio + "/*.pdf")

# Mescla os arquivos
mesclar_pdf(arquivos_pdf)

# Carimba os arquivos
#stamp('Arquivos_mesclados.pdf', carimbo, 'Arquivo_carimbado.pdf')
arquivo_carimbado = stamp('Arquivos_mesclados.pdf', carimbo, 'Arquivo_carimbado.pdf')

# Adiciona npáginas em branco
#adicionar_pagina('Arquivo_carimbado.pdf')
arquivo_sem_numeracao = adicionar_pagina(arquivo_carimbado)

# Insere numeração nas páginas
numerar_paginas('Arquivo_sem_numeracao.pdf','1710')
#numerar_paginas(arquivo_sem_numeracao,'171')

# Deleta arquivos desnecessários
try:
    os.remove('Arquivos_mesclados.pdf')
    print("Arquivo deletado com sucesso!")
except FileNotFoundError:
    print("O arquivo não foi encontrado.")
except PermissionError:
    print("Você não possui permissão para deletar este arquivo.")

print('Arquivo pronto!')'''


# In[ ]:





# In[ ]:




