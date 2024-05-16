# Todos os imports

import os
import glob
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox
import pyautogui     
import time
from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
import csv
import datetime
import subprocess
import platform
import re

# Configuração do caminho do executável do Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Definição das coordenadas e dimensões para a captura de tela dos erros e do nome do arquivo
x_nome, y_nome, largura_nome, altura_nome = 740, 472, 182, 49
x, y, largura, altura = 545, 445, 444, 191
# x, y, largura, altura = 590, 371, 274, 156
# x_nome, y_nome, largura_nome, altura_nome = 661, 385, 142, 31

# Função para pré-processamento da imagem para OCR
def preprocess_image_for_ocr(screenshot, enhance=True, save=True):
    # Converter para escala de cinza
    gray = screenshot.convert('L')
    # Se solicitado, melhora a imagem para detecção pelo OCR
    if enhance:      
        # Aumenta o tamanho da imagem
        gray = gray.resize((gray.width * 2, gray.height * 2), Image.LANCZOS)
        # Aumenta o contraste
        contrast = ImageEnhance.Contrast(gray).enhance(1.5)
        # Binariza a imagem
        threshold = contrast.point(lambda x: 0 if x < 150 else 255, '1')
        final_image = threshold
    else:
        final_image = gray
    return final_image


# Screenshot do numero da xml
def capturar_nr_arquivo(x_nome, y_nome, largura_nome, altura_nome, save=True):
    screenshot = pyautogui.screenshot(region=(x_nome, y_nome, largura_nome, altura_nome))
    processed_image = preprocess_image_for_ocr(screenshot)  # Processamento direto da imagem capturada
    nr_arquivo = pytesseract.image_to_string(processed_image)
    nr_arquivo_numeros = ''.join(re.findall(r'\d+', nr_arquivo))
    print("Número identificado pelo OCR:", nr_arquivo_numeros)
    if save:
        processed_image.save('Numero_Processado.png')  # Salva a imagem processada
    return nr_arquivo_numeros

# Screenshot dos erros
def capturar_texto_na_tela(x, y, largura, altura, save=True):
    screenshot = pyautogui.screenshot(region=(x, y, largura, altura))
    processed_image = preprocess_image_for_ocr(screenshot)  # Processamento direto da imagem capturada
    texto = pytesseract.image_to_string(processed_image)
    print("Texto identificado pelo OCR (Erro):", texto)
    if save:
        processed_image.save('Erro_Processado.png')  # Salva a imagem processada
    return texto


def get_log_path():
    # Retorna o caminho absoluto do arquivo de log
    return os.path.join(os.path.abspath(os.getcwd()), "log_erros.csv")

# Gerar um log
def log_to_csv(numero, tipo, erro):
    log_path = get_log_path()

    # Cria o cabeçalho se o arquivo não existir
    if not os.path.exists(log_path):
        with open(log_path, 'w', newline='', encoding='utf-8-sig') as file:
            writer = csv.writer(file, delimiter=';')  # Mudando para delimitador ';'
            writer.writerow(["Data", "Numero", "Tipo", "Erro"])  # Adicionando cabeçalhos

    # Adiciona uma nova entrada no arquivo
    with open(log_path, 'a', newline='', encoding='utf-8-sig') as file:
        writer = csv.writer(file, delimiter=';')  # Mudando para delimitador ';'
        timestamp = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        writer.writerow([timestamp, numero, tipo, erro])

# Abrir o log
def abrir_arquivo_log():
    log_path = get_log_path()
    try:
        if platform.system() == 'Windows':
            os.startfile(log_path)
        elif platform.system() == 'Darwin':
            subprocess.call(('open', log_path))
        else:
            subprocess.call(('xdg-open', log_path))
    except Exception as e:
        print(f"Erro ao abrir o arquivo {log_path}: {e}")

# Tabela dos possiveis erros
def verificar_erros_e_logar(x, y, largura, altura):
    texto = capturar_texto_na_tela(x, y, largura, altura, save=True)

    # Erros comuns
    erros_críticos = {
        "Transportadora nao foi informada": "Transportadora nao foi informada",
        "nao foi informada": "Transportadora nao foi informada",
        "Este conhecimento ja existe": "Este conhecimento ja existe no sistema para essa transportadora",
        "ja existe": "Este conhecimento ja existe no sistema para essa transportadora",
        "Produto <FRETE>": "Produto <FRETE> nao foi informado !!!",  
        "<FRETE>": "Produto <FRETE> nao foi informado !!!"
    }

    # Erro fatal
    erro_encontrado = None
    for erro, mensagem_erro in erros_críticos.items():
        if erro in texto:
            print(f"Erro encontrado: {mensagem_erro}")  # Debug
            erro_encontrado = mensagem_erro
            if erro_encontrado == "Produto <FRETE> nao foi informado !!!":
                # Interrompe a execução retornando o erro e o número do arquivo
                return erro_encontrado
            break

    # Retorna o erro e o número do arquivo (o erro pode ser None se não for encontrado)
    return erro_encontrado


# Função para criar uma pasta para arquivos processados
def criar_pasta_lancados(diretorio):
    pasta_lancados = os.path.join(diretorio, "NF Lançada")
    # Cria a pasta se ela não existir
    if not os.path.exists(pasta_lancados):
        os.makedirs(pasta_lancados)
    return pasta_lancados

# Função intermediária para selecionar o diretório e renomear arquivos XML
def renomear_arquivos_xml_intermediario():
    # Abre um diálogo para o usuário selecionar o diretório
    diretorio = filedialog.askdirectory()
    if diretorio:
        # Chama a função de renomeação para o diretório selecionado
        renomear_arquivos_xml(diretorio)

# Função para renomear arquivos XML com base em um critério específico
def renomear_arquivos_xml(diretorio):
    # Busca por todos os arquivos XML no diretório
    arquivos_xml = glob.glob(os.path.join(diretorio, '*.xml'))
    for arquivo_xml in arquivos_xml:
        try:
            # Analisa o arquivo XML para encontrar o elemento específico
            tree = ET.parse(arquivo_xml)
            root = tree.getroot()
            nCT_elemento = root.find('.//{http://www.portalfiscal.inf.br/cte}nCT')
            if nCT_elemento is not None and nCT_elemento.text:
                # Renomeia o arquivo com base no conteúdo do elemento nCT
                novo_nome = f"{nCT_elemento.text}.xml"
                novo_caminho = os.path.join(diretorio, novo_nome)
                os.rename(arquivo_xml, novo_caminho)
                print(f"Arquivo {arquivo_xml} renomeado para {novo_nome}")
            else:
                print(f"Arquivo {arquivo_xml} não contém o elemento <nCT> ou está vazio.")
        except Exception as e:
            # Mostra uma mensagem de erro se algo der errado
            messagebox.showerror("Erro", f"Erro ao processar o arquivo {arquivo_xml}: {str(e)}")
            print(f"Erro ao processar o arquivo {arquivo_xml}: {str(e)}")
    
    # Após processar todos os arquivos, uma mensagem de conclusão é exibida
    criar_pasta_lancados(diretorio)
    messagebox.showinfo("Concluído", "Arquivos XML foram renomeados com sucesso!\n \n                                   OBSERVAÇÃO \n \n Deixe selecionado no sistema o primeiro XML's para lançar. \n \n Após isso clique no 'SIM'!")

# Código principal do loop
def iniciar_programa_automacao():
    pyautogui.click(1566,20)
        # pyautogui.moveTo(830, 191, duration=1)  # Início (1)
        # pyautogui.click(830, 191) # Início (1)
        # pyautogui.press('Enter')
        # time.sleep(1)
        # pyautogui.doubleClick(696,391) # Nota Selecionada (2)
        # time.sleep(1)
        # pyautogui.click(521,643) # FRETE ISENTO
        # pyautogui.write("Frete", interval=0.1)
        # pyautogui.press('F1')
        # pyautogui.press('Enter')
        # pyautogui.click(523,673) # SERVICO DE TRANSPORTES
        # pyautogui.press('F1')
        # pyautogui.press('Enter')
        # time.sleep(1)
        # pyautogui.moveTo(869, 193, duration=1) # Fim (4)
    ## Frete não informado.
    erro_encontrado = verificar_erros_e_logar(x, y, largura, altura)
    if erro_encontrado == "Produto <FRETE> nao foi informado !!!":
        messagebox.showerror("Erro Crítico", f"{erro_encontrado} Programa encerrado.")
        log_to_csv("-", "Error", erro_encontrado)
        pyautogui.moveTo(0,0, duration=0.4)
    time.sleep(1)
    num_iteracoes = int(entry_iterations.get())
    for i in range(num_iteracoes):
        try:
            pyautogui.click(1273,20)
                # pyautogui.click(830, 191)  # Início (1)
                # pyautogui.press('Enter')
                # time.sleep(0.70)
                # inicio_x, inicio_y = 728, 396  # Nota Selecionada (2)
                # fim_x, fim_y = 723, 373  # Nota na Pasta (3)
                # pyautogui.mouseDown(inicio_x, inicio_y)
                # time.sleep(0.15)
                # pyautogui.moveTo(fim_x, fim_y, duration=0.15)  # 'duration' controla a velocidade do arraste
                # pyautogui.mouseUp()
            nr_arquivo_capturado = capturar_nr_arquivo(x_nome, y_nome, largura_nome, altura_nome)
            if not nr_arquivo_capturado:
                nr_arquivo_capturado = "0"
            time.sleep(1)
                # pyautogui.doubleClick(728, 396)  # Nota Selecionada (2)
                # pyautogui.click(869, 193)  # Fim (4)
            pyautogui.click(1420,20)
            erro_encontrado = verificar_erros_e_logar(x, y, largura, altura)
            if erro_encontrado:
                log_to_csv(nr_arquivo_capturado, "Warning", erro_encontrado)
            time.sleep(1)
                # time.sleep(0.4)
                # pyautogui.press('Enter')
                # time.sleep(0.2)
                # pyautogui.press('Enter')
                # time.sleep(0.2)
                # pyautogui.click(830, 191)  # Início (1)
                # pyautogui.press('Enter')
            if not erro_encontrado:
                log_to_csv(nr_arquivo_capturado, "Info", "Nenhum erro detectado")
        except Exception as e:
            print(f"Ocorreu um erro: {e}")
            log_to_csv("0", "Error", str(e))
            break  # Encerra o loop se houver exceção
    abrir_arquivo_log()

# Configurações da janela principal
root = tk.Tk()
root.title("AutomaçãoXML")

# Cria um botão para selecionar o diretório
botao_selecionar = tk.Button(root, text="Renomeie os Xml's", command=renomear_arquivos_xml_intermediario)
botao_selecionar.pack(pady=20)

# Criando um frame para organizar os widgets
frame = tk.Frame(root)
frame.pack(pady=1)

# Função para exibir a imagem em uma nova janela
def exibir_imagem():
    # Carrega a imagem
    imagem = Image.open("exemplo.png")  # Substitua "exemplo.png" pelo nome da sua imagem
    # Mostra a imagem em uma nova janela
    imagem.show()

# Adicionando um botão "Exemplo"
botao_exemplo = tk.Button(frame, text="Exemplo", command=exibir_imagem)
botao_exemplo.pack(pady=1)

# Label de instrução
label = tk.Label(frame, text="Deseja iniciar o lançamento automático?", font=("Helvetica", 14))
label.pack(pady=10)

# Campo de entrada para o número de iterações
entry_iterations = tk.Entry(frame, font=("Helvetica", 12), width=10)
entry_iterations.pack(pady=5)
entry_iterations.insert(0, "1")  # Valor padrão inicial
entry_iterations.focus_set()  # Define o foco inicial para o campo de entrada

# Botões de confirmação
button_yes = tk.Button(frame, text="Sim", font=("Helvetica", 12), padx=10, pady=5, command=iniciar_programa_automacao)
button_yes.pack(side=tk.LEFT, padx=20)

button_no = tk.Button(frame, text="Não", font=("Helvetica", 12), padx=10, pady=5, command=root.destroy)
button_no.pack(side=tk.RIGHT, padx=20)

# Inicia o loop principal da janela
root.mainloop()