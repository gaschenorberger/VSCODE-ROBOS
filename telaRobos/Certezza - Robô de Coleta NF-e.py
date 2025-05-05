import customtkinter as ctk
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import filedialog
import os
import subprocess
import threading



# Inicialização da Janela
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("Painel de Download de Notas - Certezza")
root.geometry("520x600")
root.resizable(False, False)

# Função para logs
def adicionar_log(texto):
    log_textbox.configure(state="normal")
    log_textbox.insert("end", texto + "\n")
    log_textbox.configure(state="disabled")
    log_textbox.see("end")

# Função de iniciar robô
def iniciar_robo():
    def rodar_script(script):
        try:
            processo = subprocess.Popen(
                ["python", script],
                cwd=pasta_scripts,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1
            )

            for linha in processo.stdout:
                adicionar_log(linha.strip())

            processo.stdout.close()
            processo.wait()

        except Exception as e:
            adicionar_log(f"❌ Erro ao executar o robô: {e}")

    tipo_download = var_opcao.get()
    if tipo_download == "":
        adicionar_log("⚠️ Selecione PDF ou XML antes de iniciar.")
        return

    pasta_scripts = r"C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\BUSCA_XML"

    if tipo_download == "PDF":
        script = "BUSCA_XML_PDF.py"
        adicionar_log("☑️ Iniciando robô para PDF...")
    elif tipo_download == "XML":
        script = "BUSCA_XML - Backup.py"
        adicionar_log("☑️ Iniciando robô para XML...")
    else:
        adicionar_log("❌ Tipo de download inválido.")
        return

    caminho_script = os.path.join(pasta_scripts, script)

    # Executa o script em uma thread para não travar a interface
    thread = threading.Thread(target=rodar_script, args=(caminho_script,))
    thread.start()

# Funções auxiliares
def upload_arquivo():
    arquivo = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    if arquivo:
        adicionar_log(f"Arquivo selecionado: {os.path.basename(arquivo)}")

def abrir_vpn():
    os.startfile(r"c:\Program Files (x86)\Kaspersky Lab\Kaspersky VPN 5.21\ksdeui.exe")
    adicionar_log("Abrindo VPN...")

# Variável para opção selecionada
var_opcao = tk.StringVar(value="")

# Frame principal
main_frame = ctk.CTkFrame(master=root, fg_color="#D9D9D9", border_color="#F5C400", border_width=2)
main_frame.pack(padx=10, pady=(10, 0), fill="x")

# Radiobuttons PDF/XML
opcoes_frame = ctk.CTkFrame(master=main_frame, fg_color="transparent")
opcoes_frame.grid(row=0, column=0, columnspan=2, pady=(10, 0), padx=10, sticky="w")

ctk.CTkRadioButton(opcoes_frame, text="PDF", variable=var_opcao, value="PDF", radiobutton_height=15, radiobutton_width=15).pack(side="left", padx=(10, 20))
ctk.CTkRadioButton(opcoes_frame, text="XML", variable=var_opcao, value="XML", radiobutton_height=15, radiobutton_width=15).pack(side="left")

# Logo
try:
    logo = Image.open(r"telaRobos\img\logo.png").resize((100, 40))
    logo_image = ImageTk.PhotoImage(logo)
    logo_label = ctk.CTkLabel(main_frame, image=logo_image, text="")
    logo_label.place(x=400, y=10)
except:
    pass

# Instruções
instrucoes = (
    "1- ESCOLHER O TIPO DE DOWNLOAD (PDF/XML)\n"
    "2- FAZER UPLOAD DE ARQUIVO TXT CONTENDO AS NOTAS UMA EMBAIXO DA OUTRA\n"
    "3- ATIVAR A VPN\n"
    "4- CLICAR NO BOTÃO DE INICIAR"
)
ctk.CTkLabel(main_frame, text=instrucoes, justify="left", font=("Arial", 12)).grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="w")

# Botões
botao_upload = ctk.CTkButton(main_frame, text="UPLOAD", command=upload_arquivo, fg_color="#FFCC00", hover_color="#E6B800", text_color="black", width=100)
botao_upload.grid(row=1, column=1, padx=10, pady=5, sticky="e")

botao_iniciar = ctk.CTkButton(main_frame, text="INICIAR", command=iniciar_robo, fg_color="#FFCC00", hover_color="#E6B800", text_color="black", width=100)
botao_iniciar.grid(row=2, column=0, padx=10, pady=(5, 10), sticky="w")

botao_vpn = ctk.CTkButton(main_frame, text="ABRIR VPN", command=abrir_vpn, fg_color="#FFCC00", hover_color="#E6B800", text_color="black", width=100)
botao_vpn.grid(row=2, column=1, padx=10, pady=(5, 10), sticky="e")

# Área de Log
log_frame = ctk.CTkFrame(master=root, fg_color="#EDEDED", border_color="#F5C400", border_width=2)
log_frame.pack(padx=10, pady=(0, 10), fill="both", expand=True)

log_textbox = tk.Text(log_frame, bg="#EDEDED", font=("Courier New", 10), relief="flat", state="disabled", wrap="word", height=15)
log_textbox.pack(padx=10, pady=10, fill="both", expand=True)

root.mainloop()
