import pyautogui
import flet as ft
import os
from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
import csv
import datetime
import time
import subprocess
import platform
import re
import logging

# Configuração do caminho do executável do Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Definição das coordenadas e dimensões para a captura de tela dos erros e do nome do arquivo
x, y, largura, altura = 1211, 254, 119, 44
x_nome, y_nome, largura_nome, altura_nome = 1133, 309, 107, 28

# Função para pré-processamento da imagem para OCR
def preprocess_image_for_ocr(screenshot, enhance=True):
    gray = screenshot.convert('L')
    if enhance:
        gray = gray.resize((gray.width * 2, gray.height * 2), Image.LANCZOS)
        contrast = ImageEnhance.Contrast(gray).enhance(1.5)
        threshold = contrast.point(lambda x: 0 if x < 150 else 255, '1')
        final_image = threshold
    else:
        final_image = gray
    return final_image

# Screenshot do numero da xml
def capturar_nr_arquivo(view, x_nome, y_nome, largura_nome, altura_nome):
    screenshot = pyautogui.screenshot(region=(x_nome, y_nome, largura_nome, altura_nome))
    processed_image = preprocess_image_for_ocr(screenshot)
    nr_arquivo = pytesseract.image_to_string(processed_image)
    nr_arquivo_numeros = ''.join(re.findall(r'\d+', nr_arquivo))
    view.controls[2].content = f"Número identificado pelo OCR: {nr_arquivo_numeros}"
    view.update()
    return nr_arquivo_numeros

# Screenshot dos erros
def capturar_texto_na_tela(x, y, largura, altura, save=True):
    screenshot = pyautogui.screenshot(region=(x, y, largura, altura))
    processed_image = preprocess_image_for_ocr(screenshot)  # Processamento direto da imagem capturada
    texto = pytesseract.image_to_string(processed_image)
    print("Texto identificado pelo OCR (Erro):", texto)
    if save:
        processed_image.save('Verificacao.png')  # Salva a imagem processada
    return texto

def get_log_path():
    # Retorna o caminho absoluto do arquivo de log
    return os.path.join(os.path.abspath(os.getcwd()), "log-Erros.csv")

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
        "Atualizada": "Atualizada",
    }

    for erro, mensagem_erro in erros_críticos.items():
        if erro in texto:
            print(f"Erro encontrado: {mensagem_erro}")  # Debug
            return mensagem_erro  # Retorna o erro encontrado para uso externo
    return None  # Retorna None se nenhum erro for encontrado

def execute_program(view, entry_iterations, status_label):
    num_iteracoes = int(entry_iterations.value)
    contador_atualizada = 0  # Variável para contar as ocorrências consecutivas de "Atualizada"
    logging.basicConfig()
        # pyautogui.moveTo(902, 245, duration=0.1)
        # time.sleep(1)
    for _ in range(num_iteracoes):
        try:
            pyautogui.click(1376, 23)
            time.sleep(1)
            nr_arquivo_capturado = capturar_nr_arquivo(view, x_nome, y_nome, largura_nome, altura_nome)
            if not nr_arquivo_capturado:
                nr_arquivo_capturado = "0"
            pyautogui.moveTo(800, 800, duration=0.5)
            time.sleep(1)

            erro_encontrado = verificar_erros_e_logar(x, y, largura, altura)
            if erro_encontrado == "Atualizada":
                contador_atualizada += 1
                status_label.value = f"'Atualizada' encontrada {contador_atualizada} vezes."
                if contador_atualizada >= 3:
                    status_label.value += " Programa encerrado. 'Atualizada' encontrada 3 vezes consecutivas."
                    break
                log_to_csv(nr_arquivo_capturado, "Warning", erro_encontrado)
            else:
                if contador_atualizada:
                    contador_atualizada = 0  # Redefine o contador se "Atualizada" não foi encontrada
            time.sleep(1)
            pyautogui.moveTo(450, 450, duration=2)
            if not erro_encontrado:
                log_to_csv(nr_arquivo_capturado, "Info", "Nenhum erro detectado")
        except Exception as e:
            status_label.value += f"\nErro durante a execução: {str(e)}"
            break
        finally:
            view.update()
    abrir_arquivo_log()
    status_label.value += "\nExecução concluída."
    view.update()

def show_alert(page, text):
    def close_alert(e):
        alert.open = False
        page.update()

    alert = ft.AlertDialog(
        title="Alerta",
        content=ft.Text(text),
        actions=[
            ft.TextButton("OK", on_click=close_alert),
        ],
        open=True,
    )
    page.dialog = alert
    page.update()

def main(page: ft.Page):
    page.title = "Atualizar XML"
    titulo = ft.Text("Quantas notas serão atualizadas?")
    valor_atualizar = ft.TextField(value="1", label='Número de iterações desejadas', hint_text="Digite o número")
    botão_inicio = ft.ElevatedButton(text="Iniciar Programa de Automação",
                                    on_click=lambda e: execute_program(page, valor_atualizar, texto_tamanho))
    botao_teste = ft.ElevatedButton(text="Testando a box",
                                    on_click=lambda e: show_alert(page, "Esta é uma caixa de diálogo de teste"))
    texto_tamanho = ft.Text(value="", size=16)

    page.add(titulo, valor_atualizar, botão_inicio, botao_teste, texto_tamanho)

ft.app(target=main)




