import customtkinter as ctk
from PIL import Image

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("Painel de Download de Notas - Certezza")
root.geometry("510x640")  # tamanho ajustado para o conteúdo

# Logo
# logo_img = ctk.CTkImage(light_image=Image.open("certezza_logo.png"), size=(130, 40))
# logo_label = ctk.CTkLabel(root, image=logo_img, text="")
# logo_label.place(x=360, y=15)

# Checkbox group (PDF/XML)
download_type = ctk.StringVar(value="")  # para controle único

radio_pdf = ctk.CTkRadioButton(root, text="PDF", variable=download_type, value="pdf", radiobutton_height=20, radiobutton_width=20, border_color="#000", hover_color="#ffcc00", fg_color="#ffcc00")
radio_pdf.place(x=120, y=30)

radio_xml = ctk.CTkRadioButton(root, text="XML", variable=download_type, value="xml", radiobutton_height=20, radiobutton_width=20, border_color="#000", hover_color="#ffcc00", fg_color="#ffcc00")
radio_xml.place(x=200, y=30)

# Instruções
instructions = """1- ESCOLHER O TIPO DE DOWNLOAD (PDF/XML)
2- FAZER UPLOAD DE ARQUIVO TXT CONTENDO AS NOTAS UMA EMBAIXO DA OUTRA
3- ATIVAR A VPN
4- CLICAR NO BOTÃO DE INICIAR"""
instructions_label = ctk.CTkLabel(root, text=instructions, justify="left", font=("Arial", 12))
instructions_label.place(x=30, y=70)

# Botões
btn_upload = ctk.CTkButton(root, text="UPLOAD", fg_color="#ffcc00", hover_color="#e6b800", text_color="black", width=100)
btn_upload.place(x=350, y=110)

btn_iniciar = ctk.CTkButton(root, text="INICIAR", fg_color="#ffcc00", hover_color="#e6b800", text_color="black", width=100)
btn_iniciar.place(x=100, y=160)

btn_vpn = ctk.CTkButton(root, text="ABRIR VPN", fg_color="#ffcc00", hover_color="#e6b800", text_color="black", width=100)
btn_vpn.place(x=300, y=160)

# Log
log_box = ctk.CTkTextbox(root, width=470, height=350, corner_radius=10)
log_box.place(x=20, y=220)
log_box.insert("end", "✅ Iniciando robô para PDF...\nChrome conectado com sucesso!\n...")

root.mainloop()
