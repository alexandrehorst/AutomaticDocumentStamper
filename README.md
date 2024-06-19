<h1 align="center">:file_cabinet: Carimbador Automático de Documentos/ Automatic Document Stamper README.md</h1>

## :memo: Descrição/ Description
Este programa nasceu da necessidade de se otimizar os processos adminitrativos de onde trabalho, substituindo um trabalho manual (necessidade de carimbar algumas dezenas ou centenas de folhas), por outro totalmente automatizado.  

O programa permite que a partir de um conjunto de arquivos em formato PDF seja gerado um arquivo PDF final contendo todas as folhas carimbadas frente e verso e numeradas, caso o usuário deseje. 

A aplicação insere um carimbo na parte frontal de cada página, numera as mesmas, se desejado, e insere um carimbo "em branco" no verso de cada página. Ao final, o programa gera  dois arquivos em pdf: "Arquivo_pronto.pdf" (conterá todas as páginas carimbadas e numeradas) e "Arquivo_sem_numeracao.pdf" (conterá apenas as páginas carimbadas). Para ter a documentação pronta, resta apenas ao usuário, imprimir o documento final colorido e no modo FRENTE E VERSO. 

Para que as páginas sejam numeradas basta que o usuário preencha o campo "página inicial para numeração" contendo o número a ser inserido na primeira página. Caso, esse campo seja deixado em branco, as páginas não serão numeradas, mas apenas carimbadas.

A interface de usuário (GUI) (window.png) foi desenvolvida utilizando-se o Proxlight Designer onde se cria um projeto, a partir do qual são gerados um conjunto de imagens (contidas no diretório GUI images) e um arquivo .py utilizando-se o Figma (interface de desenvolvimento baseada no módulo Tkinter). 

=====================================================================================

This program was born out of the need to optimize administrative processes at my workplace, replacing manual tasks (the need to stamp dozens or hundreds of pages) with a completely automated solution.

The program allows the generation of a final PDF file from a set of PDF files, containing all pages stamped on both sides and numbered, if desired by the user.

The application inserts a stamp on the front of each page, numbers the pages if desired, and inserts a "blank" stamp on the back of each page. In the end, the program generates two PDF files: "Arquivo_pronto.pdf" (containing all stamped and numbered pages) and "Arquivo_sem_numeracao.pdf" (containing only the stamped pages). To have the documentation ready, the user only needs to print the final document in color and in double-sided mode.

To number the pages, the user just needs to fill in the "initial page for numbering" field with the number to be inserted on the first page. If this field is left blank, the pages will not be numbered, only stamped.

The user interface (GUI) (window.png) was developed using Proxlight Designer, creating a project from which a set of images (contained in the GUI images directory) and a .py file are generated using Figma (a development interface based on the Tkinter module).

## ⚓: Dependências/ Dependencies
Este código foi desenvolvido utilizando o Jupyter 6.4.12 e o Python versão 3.9.13. Além disso, o código possui as seguintes dependências: 
* PyPDF2==3.0.1; 
* PyMuPDF==1.24.5; e
* typing_extensions==4.12.2. 

This code was developed using Jupyter 6.4.12 e o Python version 3.9.13. Additionally, the code has the following dependencies:
* PyPDF2==3.0.1; 
* PyMuPDF==1.24.5; and
* typing_extensions==4.12.2. 

## :books: Funcionalidades/ Features
* Carimba todas as folhas frente e verso;
* Numera páginas; e
* Junta todos os documentos iniciais em um documento único.
  
===========================================================
* Stamps all pages front and back;
* Numbers pages; and
* Combines all initial documents into a single file.
 
## :wrench: Tecnologias utilizadas/ Technologies Used
* Python; and
* Tkinter.
  

## :rocket: Rodando o projeto/ Running the code
Antes de rodar o código é importante colocar todos os arquivos que precisam ser carimbados num único diretório e ordená-los (basta inserir um número inicial no nome de cada arquivo. Ex: 1.arquivo1, 2.arquivo2, 3.arquivo3 etc)

Para rodar o código é necessário o seguinte (usando o prompt do anaconda):
1) Criar o Ambiente virtual num determinado diretório: conda create -n <venv_name> python=3.9.13
2) Ativar o ambiente virtual: conda activate <venv_name>
3) Verificar se está tudo ok: python --version (deve ser mostrado a mensagem: Python 3.9.13)
4) Usando o prompt, buscar o diretório onde se encontram os arquivos do código.
5) Instalar as dependências usando o comando: pip install PyPDF2==3.0.1 PyMuPDF==1.24.5 typing_extensions==4.12.2
6) Executar o código: python "window.py"

Before running the code, it's important to place all the files that need to be stamped in a single directory and arrange them (simply insert an initial number in the name of each file. Example: 1.file1, 2.file2, 3.file3, etc.).

To run the code, the following steps are necessary (using the Anaconda prompt):

1) Create the virtual environment in a specific directory: conda create -n <venv_name> python=3.9.13
2) Activate the virtual environment: conda activate <venv_name>
3) Verify everything is okay: python --version (it should display the message: Python 3.9.13)
4) Using the prompt, navigate to the directory where the code files are located.
5) Install dependencies using the command: pip install PyPDF2==3.0.1 PyMuPDF==1.24.5 typing_extensions==4.12.2
6) Execute the code: python "window.py"

## :dart: Status do projeto/ Project Status
Deseja-se inserir assinatura digital numa próxima versão do programa.

It is desired to add digital signatures in a future version of the program.   
