from pathlib import Path
from typing import Union, Literal, List
from PyPDF2 import PdfWriter, PdfReader, PdfMerger
from PyPDF2.generic import RectangleObject
import glob, os, fitz, re
#from memory_profiler import profile


# Função para incluir carimbo no PDF extraída de https://pypdf2.readthedocs.io/en/latest/user/add-watermark.html.
# O carimbo foi preparado no paint, inserido e posicionado num documento do word, que foi exportado para PDF 
#@profile
def stamp(
    content_pdf: Path,
    stamp_pdf: Path,
    stamp_pdf_90: Path,
    pdf_result: Path,
    page_indices: Union[Literal["ALL"], List[int]] = "ALL",
):           
    writer = PdfWriter()  # Cria um objeto     
    reader_content = PdfReader(content_pdf)
    if page_indices == "ALL":
        page_indices = list(range(0, len(reader_content.pages)))
    for index in page_indices:        
        content_page = reader_content.pages[index]
        mediabox = content_page.mediabox
        if reader_content.pages[index].rotation == 90:
            reader = PdfReader(stamp_pdf_90)  # lê o arquivo com o carimbo
            image_page = reader.pages[0]  # pega a página contendo o carimbo
        else:
            reader = PdfReader(stamp_pdf)
            image_page = reader.pages[0]  # pega a página contendo o carimbo
        image_page = reader.pages[0]    
        content_page.merge_page(image_page)
        content_page.mediabox = mediabox
        writer.add_page(content_page)    
    #with open(pdf_result, "wb") as fp:   # alteração #
    #    writer.write(fp)
    return writer


# Função que adiciona uma página (página em branco) 
#@profile
def adicionar_pagina(arquivo):
    novo_arquivo = PdfWriter()  # cria novo arquivo pdf    
    pagina_em_branco = PdfReader('em_branco.pdf')    
    for pagina in arquivo.pages:
        novo_arquivo.add_page(pagina)
        novo_arquivo.add_page(pagina_em_branco.pages[0])    
    # Salva cada página num arquivo pdf dentro do diretório páginas
    with Path('Arquivo_sem_numeracao.pdf').open(mode="wb") as arquivo:
        novo_arquivo.write(arquivo)
    

# Função que mescla os arquivos pdf
#@profile   
def mesclar_pdf(lista_arquivos_pdf):
    pdf_mesclado = fitz.open()  # Cria um documento PDF vazio
    for pdf_file in lista_arquivos_pdf:
        documento_atual = fitz.open(pdf_file)  # Abre cada PDF individualmente
        pdf_mesclado.insert_pdf(documento_atual)  # Insere o PDF atual no PDF mesclado
    # Salva os arquivos PDFs mesclados num arquivo
    pdf_mesclado.save("Arquivos_mesclados.pdf")
        

# Função para inserir números nas páginas
#@profile
def numerar_paginas(pdf_path, nr_inicial):
    pdf = fitz.open(pdf_path)
    nr_inicial = int(nr_inicial)
    
    for page_number, page in enumerate(pdf):
        # Obtém as dimensões da página
        width, height = page.rect.width, page.rect.height
        # Define as coordenadas para o canto superior direito (quando não há rotação)
        x, y = width - 58, 30
        
        # Verifica a rotação da página
        rotation = page.rotation

        if rotation == 0:
            # Página não rotacionada
            x, y = width - 58, 30 
        elif rotation == 90:
            # Página rotacionada 90 graus (texto deve ser inserido como se estivesse em 0 graus)
            #x, y = height - 58, width - 30
            x, y = 30, 58
        elif rotation == 180:
            # Página rotacionada 180 graus (texto deve ser inserido como se estivesse em 0 graus)
            x, y = 58, height - 30
        elif rotation == 270:
            # Página rotacionada 270 graus (texto deve ser inserido como se estivesse em 0 graus)
            x, y = 30, 58
        
        if page_number%2 == 0:
            # Define as propriedades do texto
            text_properties = {"text": str(nr_inicial), "fontname": "helv", "fontsize": 12, "color": (0, 0, 1)}
            if rotation == 90:
                # Adiciona o texto rotacionado 90 graus para estar na horizontal
                page.insert_text((x, y), **text_properties, rotate=90)
            elif rotation == 180:
                # Adiciona o texto rotacionado 180 graus para estar na horizontal
                page.insert_text((x, y), **text_properties, rotate=180)
            elif rotation == 270:
                # Adiciona o texto rotacionado 270 graus para estar na horizontal
                page.insert_text((x, y), **text_properties, rotate=270)
            else:
                # Adiciona o texto sem rotação adicional
                page.insert_text((x, y), **text_properties)
            nr_inicial += 1
            
    # Salva as alterações
    pdf.save('Arquivo_pronto.pdf')
    pdf.close()

   
# Função que transforma as páginas de formato paisagem para retrato     
def rotate_landscape_pages(arquivo, arquivo_rotacionado):   
    # Abrir o documento PDF
    doc = fitz.open(arquivo)
    
    # Iterar sobre todas as páginas
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        
        # Obter as dimensões da página
        width, height = page.rect.width, page.rect.height
        
        # Verificar se a página está em formato landscape
        if width > height:
            # Girar a página em 90 graus no sentido horário
            page.set_rotation(90)
    
    # Salvar o documento modificado
    doc.save(arquivo_rotacionado)
    doc.close() 

    
# Coloca todas as páginas no formato A4.
def resize_to_a4(arquivo, arquivo_a4):
    # Definir as dimensões do A4 em pontos
    a4_width = 595  # largura em pontos
    a4_height = 842  # altura em pontos
    
    doc = fitz.open(arquivo)  # Abre o documento PDF
    new_doc = fitz.open()  # Cria um novo documento PDF
    
    # Itera por todas as páginas
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)  # Carrega a página atual
        original_width, original_height = page.rect.width, page.rect.height
        
        # Verifica a orientação original da página
        if original_width > original_height:
            # Página horizontal
            new_page = new_doc.new_page(width=a4_height, height=a4_width)
        else:
            # Página vertical
            new_page = new_doc.new_page(width=a4_width, height=a4_height)
        
        # Adiciona a página existente à nova página A4
        new_page.show_pdf_page(new_page.rect, doc, page_num)
    
    # Salva o PDF com as novas dimensões
    new_doc.save(arquivo_a4)
    doc.close()
    new_doc.close()

    
# Função para ordenar os aquivos dentro do diretório
def ordenar_arquivos_por_numero(diretorio):
    # Listar todos os arquivos no diretório
    arquivos = os.listdir(diretorio)
    
    # Filtrar apenas arquivos PDF
    arquivos_pdf = [arquivo for arquivo in arquivos if arquivo.lower().endswith('.pdf')]
    
    # Função para extrair a chave de ordenação do nome do arquivo
    def extrair_chave(nome_arquivo):
        # Usa expressão regular para capturar números e possíveis letras subsequentes
        match = re.match(r'^(\d+)([a-zA-Z]?)', nome_arquivo)
        if match:
            num = int(match.group(1))
            letra = match.group(2) if match.group(2) else ''
            return (num, letra)
        return (float('inf'), '')  # Retorna valor grande para números e string vazia para letras

    # Ordenar a lista de arquivos PDF com base na chave extraída
    arquivos_ordenados = sorted(arquivos_pdf, key=extrair_chave)
    
    arquivos_ordenados_2 = []
    for arquivo in arquivos_ordenados:
        arquivo = os.path.join(diretorio, arquivo)
        arquivos_ordenados_2.append(arquivo)
    
    return arquivos_ordenados_2    
    
    
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




