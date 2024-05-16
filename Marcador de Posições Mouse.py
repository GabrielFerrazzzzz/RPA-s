import tkinter as tk
import pyautogui
import threading
import os


# classe com interface
class MousePositionWidget(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Verificar Posição do Mouse")

        # Estilo para o botão
        self.button_style = {
            'background': '#4CAF50',  # Cor de fundo verde
            'foreground': 'white',     # Cor da fonte branca
            'font': ('Arial', 12),     # Fonte Arial tamanho 12
            'width': 10,               # Largura do botão
            'height': 1,               # Altura do botão
            'borderwidth': 0,          # Largura da borda
            'highlightthickness': 0,   # Espessura do destaque
            'padx': 5,                 # Espaçamento horizontal interno
            'pady': 3                  # Espaçamento vertical interno
        }

        self.label_position = tk.Label(self, text="Posição atual do mouse: X=?, Y=?")
        self.label_position.pack(pady=10)



        self.button_verify = tk.Button(self, text="INICIAR", command=self.start_update_thread, **self.button_style)
        self.button_verify.pack(pady=5)

        # Configura o evento de clique do botão para iniciar o programa
        self.button_verify.bind("<Button-1>", lambda event: self.start_update_thread())

        self.bind("<Return>", self.enter_pressed)  # Chama a função enter_pressed() quando a tecla Enter é pressionada

        self.update_thread = None  # Inicializa a variável para a thread de atualização
        self.log_index = 1  # Inicializa o índice do log
        self.first_enter_pressed = True  # Indica se o primeiro Enter foi pressionado


    # dar update da posição do mouse
    def update_mouse_position(self):
        while True:
            x, y = pyautogui.position()
            self.label_position.config(text=f"Posição atual do mouse: X={x}, Y={y}")
            self.update_idletasks()  # Atualiza a interface gráfica
            threading.Event().wait(0.1)  # Aguarda 0.1 segundos para atualizar novamente

    def start_update_thread(self):
        # Inicia a thread para atualizar a posição do mouse
        self.update_thread = threading.Thread(target=self.update_mouse_position)
        self.update_thread.start()

    # Apertando o enter ele grava a posição
    def enter_pressed(self, event):
        if self.first_enter_pressed:
            self.first_enter_pressed = False
            return

        # Verifica se a thread de atualização está em execução e a interrompe
        if self.update_thread and self.update_thread.is_alive():
            self.update_thread.join()
        
        # Atualiza a posição do mouse uma última vez antes de travar
        x, y = pyautogui.position()
        self.label_position.config(text=f"Posição atual do mouse: X={x}, Y={y}")

        # Se o botão ainda não mudou de nome, muda para "MARCAR"
        if self.button_verify['text'] == 'INICIAR':
            self.button_verify.config(text='MARCAR')
        # Gera o log da posição travada
        self.generate_log(x, y)
        # Incrementa o índice do log
        self.log_index += 1

    # Após você fechar o programa ele gerá o log com as posições obtidas
    def generate_log(self, x, y):
        log_message = f"{self.log_index}ª Posicao travada em: X={x}, Y={y}\n"
        with open("MousePosition.txt", "a") as log_file:
            log_file.write(log_message)

if __name__ == "__main__":
    # Limpar o conteúdo do arquivo de log ao iniciar o programa
    with open("MousePosition.txt", "w"):
        pass
    
    app = MousePositionWidget()
    app.protocol("WM_DELETE_WINDOW", lambda: (app.destroy(), os.system("notepad MousePosition.txt")))
    app.mainloop()