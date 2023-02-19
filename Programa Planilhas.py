# Ler os dados do Excel
import pandas as pd
import pyautogui as pya
import pyperclip as ppc
import time
import customtkinter as tk
import os

tk.set_appearance_mode('dark')
tk.set_default_color_theme('dark-blue')

def abrir_criador_planilha():
   janela2 = tk.CTkToplevel()
   janela2.title('Criar Planilha')
   janela2.geometry('500x400')
   
   label_nome = tk.CTkLabel(janela2, text='Nome')
   label_nome.pack(padx=5, pady=5)
   entry_nome = tk.CTkEntry(janela2)
   entry_nome.pack(padx=10, pady=10)
   
   label_numero = tk.CTkLabel(janela2, text='Número')
   label_numero.pack(padx=5, pady=5)
   entry_numero = tk.CTkEntry(janela2)
   entry_numero.pack(padx=10, pady=10)
   
   num = 1
   while f"nomes_numeros.xlsx" in os.listdir():
       num += 1
       
   filename = f"nomes_numeros.xlsx"
   
   nomes = []
   numeros = []
   
   def adicionar():
       nome = entry_nome.get()
       numero = entry_numero.get()
       
       nomes.append(nome)
       numeros.append(numero)
       
       entry_nome.delete(0, tk.END)
       entry_numero.delete(0, tk.END)
       
       entry_nome.focus()

   def salvar():
       
       df = pd.DataFrame({'Nomes': nomes, 'Numeros': numeros})
       
       
       with pd.ExcelWriter(filename) as writer:
           df.to_excel(writer, index=False, sheet_name='Dados')
           
   
   button_adicionar = tk.CTkButton(janela2, text="Adicionar", command=adicionar)
   button_adicionar.pack(padx=10, pady=10)
   button_salvar = tk.CTkButton(janela2, text="Salvar", command=salvar)
   button_salvar.pack(padx=10, pady=10)
   button_cancelar = tk.CTkButton(janela2, text='Cancelar', command=janela2.destroy)
   button_cancelar.pack(padx=10, pady=10)
   
   entry_nome.focus()
   
   janela2.mainloop()


# programa
def enviar_mensagem():
    
    df = pd.read_excel('nomes_numeros.xlsx')  

    
    mensagem = mensagem_entry.get()
    nav = nav_entry.get()

    
    for i in range(len(df)):
        nome = df.loc[i, 'Nomes']
        numero = df.loc[i, 'Numeros']
        time.sleep(2)

        # Abrir o navegador e navegar até o WhatsApp Web
        pya.press('win')
        pya.write(nav)
        pya.press('enter')
        time.sleep(2)
        pya.hotkey('ctrl', 't')
        time.sleep(2)
        pya.write('https://wa.me/{}'.format(numero))
        pya.press('enter')
        time.sleep(5)
        pya.press('Tab')
        pya.press('Tab')
        pya.press('Tab')
        pya.press('Tab')
        pya.press('Tab')
        pya.press('Tab')
        pya.press('Tab')
        pya.press('Tab')
        pya.press('Tab')
        pya.press('enter')
        time.sleep(5)
        pya.press('Tab')
        pya.press('Tab')
        pya.press('enter')
        time.sleep(5)

        # Enviar a mensagem de boas-vindas
        boas_vindas = 'Olá '
        ppc.copy(boas_vindas)
        pya.hotkey('ctrl', 'v')
        ppc.copy(nome)
        pya.hotkey('ctrl', 'v')
        pya.press('enter')
        ppc.copy(mensagem)
        pya.hotkey('ctrl', 'v')
        pya.press('enter')
        time.sleep(1.2)
        pya.hotkey('ctrl', 'w')
        pya.hotkey('ctrl', 't')
        time.sleep(2)


janela = tk.CTk()
janela.geometry('500x300')


mensagem_label = tk.CTkLabel(janela, text="Digite uma mensagem")
mensagem_label.pack(padx=10, pady=10)
mensagem_entry = tk.CTkEntry(janela)
mensagem_entry.pack(padx=10, pady=10)

nav_label = tk.CTkLabel(janela, text="Digite seu navegador")
nav_label.pack(padx=5, pady=5)
nav_entry = tk.CTkEntry(janela)
nav_entry.pack(padx=10, pady=10)

enviar_button = tk.CTkButton(janela, text="Enviar Mensagens", command=enviar_mensagem)
enviar_button.pack(padx=10, pady=10)
button_inserir_planilha = tk.CTkButton(janela, text='Inserir dados na planilha', command=abrir_criador_planilha)
button_inserir_planilha.pack()


# Iniciar o loop principal do tkinter
janela.mainloop()

    