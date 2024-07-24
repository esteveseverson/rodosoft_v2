import os
import shutil
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from datetime import datetime, timedelta
import pandas as pd

def login(driver, key, user, password):  
    # encontra os elementos por xpath  
    boxKey = driver.find_element(By.XPATH, "//input[@id='chave']")
    boxLogin = driver.find_element(By.XPATH, "//input[@id='login']")
    boxPassword = driver.find_element(By.XPATH, "//input[@id='password']")
    btnLogin = driver.find_element(By.XPATH, "//input[@class='btn btn-block btn-lg btn-danger']")
    # envia os valores aos elementos
    boxKey.send_keys(key)
    boxLogin.send_keys(user)
    boxPassword.send_keys(password)
    sleep(0.2)
    btnLogin.click()
    driver.get('http://backofficevendaembarcada.rodosoft.com.br/Pages/Operacoes')
    sleep(1)

def search(driver, id, data):
    # encontra os elementos
    boxService = driver.find_element(By.XPATH, "//input[@id='MainContent_txtServico']")
    btnSearch = driver.find_element(By.XPATH, "//input[@id='MainContent_Button1']")
    boxDate = driver.find_element(By.XPATH, "//input[@id='MainContent_btnenableddate']")
    sleep(0.2)
    # limpa o campo para não escrever na frente
    boxService.clear()
    boxService.send_keys(id)
    boxDate.clear()
    boxDate.send_keys(data)
    sleep(0.2)
    btnSearch.click()
    sleep(0.5)

def dataMining(driver, serviço, horaInicio, horaInicioEfetivo, horaFinalizado):
    # encontra a tabela por xpath
    td_elements = driver.find_elements(By.XPATH, "//table[@class='table margin table-striped table-hover sources-table']//td")
    sleep(0.2)
    # recebe o texto por indexação da tabela
    serviço = td_elements[0].text.strip()
    horaInicio = td_elements[5].text.strip()
    horaInicioEfetivo = td_elements[6].text.strip()
    horaFinalizado = td_elements[8].text.strip()
    sleep(0.2)
    return serviço, horaInicio, horaInicioEfetivo, horaFinalizado

def save(filePathSave, service, dayHour, travelStart, travelEnd):
    # Carregar o arquivo Excel existente em um DataFrame
    df = pd.read_excel(filePathSave)
    
    # Adicionar a nova linha ao DataFrame
    new_row = {'SERVIÇO': service, 'HORA_PLANEJADA': dayHour, 'HORA_INICIADA': travelStart, 'HORA_FINALIZADA': travelEnd}
    df = df.append(new_row, ignore_index=True)
    
    # Salvar o DataFrame de volta ao arquivo Excel
    df.to_excel(filePathSave, index=False)

def renameAndMoveFile(filePathSave):
    saveDay = datetime.now() - timedelta(days=1)
    # salvando com ano na frente para uma melhor ordenação
    newFileName = f'planilhaDiariaMonitriip{saveDay.strftime("%Y-%m-%d")}.xlsx'
    
    # Verificar se o arquivo existe antes de renomear
    if not os.path.exists(filePathSave):
        raise FileNotFoundError(f"The file {filePathSave} does not exist.")
    
    os.rename(filePathSave, newFileName)
    
    # Criar diretório se não existir
    destinationDir = 'planilhasGeradas'
    if not os.path.exists(destinationDir):
        os.makedirs(destinationDir)
    
    shutil.move(newFileName, os.path.join(destinationDir, newFileName))

def generateFile():
    filename = 'planilhaDiariaMonitriip.xlsx'
    
    # Verifica se o arquivo existe
    if os.path.exists(filename):
        # Remove o arquivo existente
        os.remove(filename)
    
    # Cria um DataFrame com os cabeçalhos
    headers = ['SERVIÇO', 'HORA_PLANEJADA', 'HORA_INICIADA', 'HORA_FINALIZADA']
    df = pd.DataFrame(columns=headers)
    
    # Salva o DataFrame em um arquivo Excel
    df.to_excel(filename, index=False)

def credentials():
    def on_submit():
        nonlocal chave, login, password
        chave = chave_entry.get()
        login = login_entry.get()
        password = password_entry.get()
        
        if not chave or not login or not password:
            messagebox.showerror("Erro", "Todos os campos devem ser preenchidos!")
            return
        
        messagebox.showinfo("Info", "Dados armazenados com sucesso!")
        root.destroy()

    def toggle_password():
        if password_entry.cget('show') == '*':
            password_entry.config(show='')
            show_password_button.config(text='Ocultar')
        else:
            password_entry.config(show='*')
            show_password_button.config(text='Mostrar')
    
    # Criando a janela principal
    root = tk.Tk()
    root.title("Login Form")
    root.geometry("300x200")
    root.resizable(False, False)
    
    # Centralizando a janela
    window_width = 300
    window_height = 200
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int(screen_height/2 - window_height/2)
    position_right = int(screen_width/2 - window_width/2)
    root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

    # Criando e posicionando os campos de entrada
    tk.Label(root, text="Chave:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
    chave_entry = tk.Entry(root)
    chave_entry.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(root, text="Login:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
    login_entry = tk.Entry(root)
    login_entry.grid(row=1, column=1, padx=10, pady=5)

    tk.Label(root, text="Password:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
    password_entry = tk.Entry(root, show="*")
    password_entry.grid(row=2, column=1, padx=10, pady=5)

    show_password_button = tk.Button(root, text="Mostrar", command=toggle_password)
    show_password_button.grid(row=2, column=2, padx=5, pady=5)

    # Criando e posicionando o botão de envio
    submit_button = tk.Button(root, text="Enviar", command=on_submit)
    submit_button.grid(row=3, columnspan=3, pady=10)

    # Variáveis para armazenar os dados
    chave = login = password = None

    # Iniciando o loop principal
    root.mainloop()

    return chave, login, password

def main():
    generateFile()
    searchDay = lambda: (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")

    filePathServices = os.path.abspath('planilhaServicos2024.xlsx')

    sheet = pd.read_excel(filePathServices)
    sheetSearch = sheet.query('DATA_SERVIÇO == @searchDay()')

    chave, usuario, senha = credentials()

    nav = webdriver.Chrome()
    nav.get('http://backofficevendaembarcada.rodosoft.com.br/Pages/Login')
    login(nav, chave, usuario, senha)

    try:
        val = nav.find_element(By.XPATH, "//input[@id='MainContent_txtServico']")
        for _, line in sheetSearch.iterrows():
            DATA_SERVIÇO, CORRIDA_ID = line['DATA_SERVIÇO'], line['CORRIDA_ID']
            if DATA_SERVIÇO == searchDay():

                search(nav, CORRIDA_ID, DATA_SERVIÇO)
                idcorrida, previsao, inicio, fim = "", "", "", ""
                idcorrida, previsao, inicio, fim = dataMining(nav, idcorrida, previsao, inicio, fim)

                filePath = os.path.abspath('planilhaDiariaMonitriip.xlsx')
                save(filePath, idcorrida, previsao, inicio, fim)

        renameAndMoveFile(os.path.abspath('planilhaDiariaMonitriip.xlsx'))
        
        messagebox.showinfo("Info", "Informações coletadas com sucesso e planilha foi movida para a pasta Planilhas Geradas")

    except:
        messagebox.showinfo("Info", "O login está errado")    

if __name__ == "__main__":
    main()