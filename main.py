from cProfile import label
from random import randint
from turtle import back
import win32com.client as client
import textwrap
import pandas as pd
import os
from tkinter.messagebox import showinfo
from cx_Freeze import setup, Executable
import sys
from tkinter import *
from PIL import Image, ImageFont, ImageDraw

# exe = Executable(
#     script=r"main.py",
#     base="Win32GUI",
# )

# setup(
#     name="main",
#     version="0.1",
#     description="An example",
#     executables=[exe]
# )

def sendMail():
    # Create a root TKinter
    root = Tk()
    # Create a Outlook Application
    outlook = client.Dispatch("Outlook.Application")

    # Configurate Title bar
    root.title('SR Brasília Sul - C150713')
    root.iconbitmap(r'sr2637.ico')
    appWidth = 600
    appHeight = 600

    screenWidth = root.winfo_screenwidth()
    screenHeight = root.winfo_screenheight()


    x = (screenWidth / 2) - (appWidth / 2)
    y = (screenHeight / 2) - (appHeight / 2)

    root.geometry(f'{appWidth}x{appHeight}+{int(x)}+{int(y)}')


    root.resizable(False, False)


    root.config(height=500, width=500)
    can = Canvas(root, bg='#ef9c00', height=400, width=420)
    can.place(relx=0.5, rely=0.5, anchor=CENTER)
    root.configure(background='#1c60ab')


    # Create Absolute Path
    file_path = os.path.abspath(os.path.dirname(__file__))
    absolutPath = "\"" + file_path.replace("\\", "\\\\") + "\\\\result.png" + "\""

    # Create Phrases to use on body in outlook e-mail
    phrases = [", hoje é um dia muito especial, pois você completa mais um ano de vida. Curta muito o seu aniversário ao lado das pessoas que mais ama. Felicidades!",
            ", espero que hoje você celebre seu dia especial junto dos que mais ama, e que o seu coração se aqueça com todo amor que receber. Feliz aniversário!",
            ", este é seu dia, e por isso deve festejar com alegria. Espero que receba muito carinho, homenagens e surpresas boas. Parabéns e muitas felicidades!"]

    # Open sheet to get data ASS
    ass = pd.read_excel(
        r"Empregados.xlsx", sheet_name='Assinatura')

    # Get datas to usage in footer as ASS outlook body e-mail
    srName = ass['Assinatura'][0]
    srEntity = ass['Assinatura'][1]
    office = ass['Assinatura'][3]

    # Concatenate ASS data
    textAss = srName + '\n' + office + '\n' + srEntity

    # Create a content

    # Root
    labelframe = LabelFrame(
        root, text="Olá, seja bem-vindo ao sistema de envio de e-mails", padx=50, pady=50)
    labelframe.pack(padx=100, pady=100, fill="both", expand="yes")


    # Label Inserir Matrícula
    insertMat = Label(labelframe, text='Inserir Matrícula', background='#ccc',
                    foreground='#009', anchor=W).place(x=50, y=0, width=200, height=25)
    # Input of Matrícula
    vMat = Entry(labelframe)
    vMat.place(x=50, y=25, width=200, height=20)

    # Label Inserir Nome
    insertNameLabel = Label(labelframe, text='Inserir Nome', background='#ccc',
                            foreground='#009', anchor=W).place(x=50, y=60, width=200, height=25)
    # Input of Name
    vNome = Entry(labelframe)
    vNome.place(x=50, y=85, width=200, height=20)


    def onClick():
        message = outlook.CreateItem(0)
        message.Display()
        message.To = vMat.get()
        message.BCC = ass['Assinatura'][2]
        message.Subject = "Feliz Aniversário!"
        firstName = vNome.get().split()

        # Transformar primeiro nome
        # # Usar Matrícula em um while como destinatário
        # # Converter data de nascimento em Brasil e verificar se há necessidade de enviar o e-mail de aniversário
        my_image = Image.open(
            "parabensind.jpg")

        title_text = firstName[0].capitalize() + phrases[randint(0, 2)]

        lines = textwrap.wrap(title_text, width=36)
        y_text = 100
        font = ImageFont.truetype(
            'BebasNeue-Regular.ttf', 28)

        fontAss = ImageFont.truetype(
            'BebasNeue-Regular.ttf', 16)

        image_editable = ImageDraw.Draw(my_image)

        for line in lines:
            width, height = font.getsize(line)
            image_editable.text(((650 - width) / 3, y_text),
                                line, font=font, fill="white", stroke_width=1, stroke_fill="white")
            y_text += height

        image_editable.text((130, 430),
                            text=textAss, fill='white', font=fontAss, anchor="ls", stroke_width=0, stroke_fill="white")

        my_image.save("result.png", optimize=True, quality=100)
        string = f"""
            <div>
                <img src={absolutPath}>
            </div> 
            """
        message.HTMLBody = string


    SendMail = Button(labelframe, text='Enviar e-mail', command=onClick)
    SendMail.place(x=45, y=170, anchor=W)
    buttonQuit = Button(labelframe, text='Sair do programa', command=root.quit)
    buttonQuit.place(x=150, y=170, anchor=W)


    root.mainloop()

if __name__ == "__main__":
    sendMail()