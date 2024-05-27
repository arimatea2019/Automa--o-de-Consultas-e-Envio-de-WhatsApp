from pynput.mouse import Listener
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import threading
import time
import os
import pyautogui
import pandas as pd
from tkinter import filedialog, Tk
from PIL import Image, ImageFilter, ImageOps, ImageGrab
import pytesseract
import re
import random


file_path = "click_positions.txt"
running = True  
click_names = [
    "CAMPO_CPF", "CAMPO_AGENCIA",
    "CAMPO_CONTA", "BOTAO_SIMULAR","BOTAO_VOLTAR"
]
click_index = 0  

def IA(screenshot_path):
    image = Image.open(screenshot_path)
    image = image.convert('L')  
    image = ImageOps.autocontrast(image)  
    image = image.resize([int(s * 2) for s in image.size], Image.Resampling.LANCZOS)
    image = image.filter(ImageFilter.SHARPEN)
    text = pytesseract.image_to_string(image,  lang='por',config='--psm 6')
    print(text)
    if "Crédito Novo" in text:
        return "Com Crédito ou Renovação"
    if 'R$' in text:
        lines = text.split('\n')

        product_names = []
        values = []

        for line in lines:
            if 'CRÉDITO' in line or 'RENOVAÇÃO' in line: 
                products = line.split('>')  
                for product in products:
                    if product.strip():
                        product_names.append(product.strip())

            if 'R$' in line:  
                vals = re.findall(r'R\$\s*\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?', line)
                values.extend(vals)

        for name, value in zip(product_names, values):
            print(f'{name}: {value}')
        return "Com Crédito ou Renovação"
    
    else:
        return "Sem Crédito ou Renovação"


def checar_erro():
    regiao = (60, 120, 90 + 200, 100 + 80)

    imagem = ImageGrab.grab(bbox=regiao)
    image = imagem.convert('L') 
    image = ImageOps.autocontrast(image)  
    image = image.resize([int(s * 2.5) for s in image.size], Image.Resampling.LANCZOS)
    image = image.filter(ImageFilter.SHARPEN)
    text = pytesseract.image_to_string(image, lang='por', config='--psm 6')
        
    if "Dados do beneficiário" in text:
        return 1
    return 0 

def checalog():
    regiao = (55, 120, 55 + 200, 120 + 50)
    imagem = ImageGrab.grab(bbox=regiao)
    image = imagem.convert('L')
    image = ImageOps.autocontrast(image)  
    image = image.resize([int(s * 2.5) for s in image.size], Image.Resampling.LANCZOS)
    image = image.filter(ImageFilter.SHARPEN)
    text = pytesseract.image_to_string(image, lang='por', config='--psm 6')
    print(text)
        
    if "Acesso rápido" in text:
        return 1
    else:
        return 0

def logoff():
    x = 1300
    y = 90
    x += random.randint(-3, 3)
    y += random.randint(-3, 3)
    pyautogui.click(x, y)
    time.sleep(random.uniform(1, 2))
    x = 800
    y = 435
    x += random.randint(-10, 10)
    y += random.randint(-4, 4)
    pyautogui.click(x, y)
    print("deslogando")
    
def login():
    x = 730
    y = 390
    x += random.randint(-10, 10)
    y += random.randint(-4, 4)
    pyautogui.click(x, y)
    print("logando")
    

def automate_whatsapp(browser,celular,nome,cpf):
    cpf =cpf[-6:]
    elemento_span = WebDriverWait(browser, 120).until(EC.element_to_be_clickable((By.XPATH,'//*[@title="Nova conversa"]')))
    elemento_span.click()
    time.sleep(random.uniform(0,1))
    elemento_p = WebDriverWait(browser, 30).until(EC.element_to_be_clickable((By.XPATH, '(//*[@aria-label="Caixa de texto de pesquisa"])[1]')))
    elemento_p.send_keys(celular)
    time.sleep(random.uniform(0,1))
    try:
        resultado_nao_encontrado = WebDriverWait(browser, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@dir="auto" and contains(text(), "Nenhum resultado encontrado para")]'))
        )
        print("Cliente sem Whatsapp.")

        elemento_volta = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH,'//*[@aria-label="Voltar"]')))
        browser.execute_script("arguments[0].click();", elemento_volta)
        return "Sem Whatsapp"
    
    except:
        elemento_p.send_keys(Keys.ENTER)
        time.sleep(random.uniform(0,1))
    mensagem = f'*'
    message_box = WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH,'(//*[@aria-label="Digite uma mensagem"])[1]')))
    for char in mensagem:
        message_box = WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH,'(//*[@aria-label="Digite uma mensagem"])[1]')))
        message_box.send_keys(char)
        time.sleep(random.uniform(0.001, 0.002))  # Adiciona um pequeno atraso aleatório entre cada caractere
        
    message_box.send_keys(Keys.ENTER)
    print("Whatsapp enviado")
    return "Sim"
    



def carregar_excel():
    root = Tk()
    root.withdraw() 
    root.attributes('-topmost', True)
    nome_do_arquivo = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx;*.xls")])
    root.attributes('-topmost', False)
    if nome_do_arquivo:
        try:
            df = pd.read_excel(nome_do_arquivo)
            print("Arquivo Excel carregado com sucesso!")
            nome_do_arquivo= os.path.basename(nome_do_arquivo)
            return df, nome_do_arquivo
        except Exception as e:
            print(f"Erro ao carregar o arquivo Excel: {e}")
            return None
    else:
        print("Nenhum arquivo selecionado.")
        return None


def on_click(x, y, button, pressed):
    global running, click_index
    if pressed and running:
        if click_index < len(click_names):
            nome_clique = click_names[click_index]
            click_index += 1
            
            with open(file_path, "a") as file:
                file.write(f"{x},{y},{nome_clique}\n")
            print(f"Posição guardada: {(x, y)} - {nome_clique}")
        else:
            print("Todos os cliques necessários foram registrados.")
            running = False
            return False
    if not running:
        return False

def monitor_input():
    global running
    while True:
        inp = input("Digite '0' e pressione Enter para parar: ")
        if inp == "0":
            running = False
            print("Gravação de cliques será interrompida.")
            break

def criar_rotina():
    global running, click_index
    open(file_path, "w").close()
    click_index = 0  
    print("Clique no botão 'Pronto' nesta janela quando quiser começar a gravar cliques.")
    input("Pressione Enter para continuar...")
    print("Comece a clicar nas posições de acordo com a sequência predefinida de nomes.")
    running = True
    input_thread = threading.Thread(target=monitor_input)
    input_thread.start()
    with Listener(on_click=on_click) as listener:
        listener.join()
    input_thread.join()
    print("Gravação de cliques interrompida pelo usuário.")

def rodar_rotina(index, cpf, agencia, conta):
    print("Rodando a rotina de cliques...")
    skip_next_click = False  # Controle para pular o próximo clique
    with open(file_path, "r") as file:
        entries = [line.strip().split(',') for line in file]
        for x, y, nome_clique in entries:
            if skip_next_click:
                skip_next_click = False  # Reseta o controle
                print(f"Clique pulado: {nome_clique}")
                time.sleep(random.uniform(1, 2))
                pyautogui.hotkey('ctrl', 'r')
                time.sleep(random.uniform(1, 2))
                continue  # Pula este clique

            x, y = int(x), int(y)
            x += random.randint(-20, 20)
            y += random.randint(-3, 3)
            pyautogui.click(x, y)
            time.sleep(random.uniform(1, 2))
            print(f"Clique realizado em: {nome_clique} ({x}, {y})")
            if nome_clique == 'CAMPO_CPF':
                for char in cpf:
                    pyautogui.write(char, interval=0.001)  
                time.sleep(random.uniform(2, 3))
            elif nome_clique == 'CAMPO_AGENCIA':
                for char in agencia:
                    pyautogui.write(char, interval=0.001)
                time.sleep(random.uniform(2, 3))
            elif nome_clique == 'CAMPO_CONTA':
                for char in conta:
                    pyautogui.write(char, interval=0.001)
                time.sleep(random.uniform(2, 3))
            elif nome_clique == 'BOTAO_SIMULAR':
                time.sleep(4)
                x = checar_erro()
                if x == 0:
                    skip_next_click = True  # Seta para pular o próximo clique
                    margem = "Erro na consulta"
                else:
                    region = (50, 250, 1200, 500)
                    screenshot_path = f"margens/screenshot_{index}.png"
                    pyautogui.screenshot(screenshot_path, region=region)
                    margem = IA(screenshot_path)
            
    return margem


def consultar(browser):
    lista, nome_do_arquivo = carregar_excel()
    
    if 'Margem' not in lista.columns:
        lista['Margem'] = ''
        lista.to_excel(nome_do_arquivo, index=False)
    if 'disparado' not in lista.columns:
        lista['disparado'] = ''
        lista.to_excel(nome_do_arquivo, index=False)

    
    total_consultas = 0
    tempos_de_consulta = []
    
    # Abrindo o arquivo de texto para armazenar os tempos
    with open('tempos_de_consulta.txt', 'w') as file:
        for index, row in lista.iterrows():
            if row['Margem'] == '':
                cpf = str(row['cpf']).zfill(11)
                agencia = str(row['agencia'])
                conta = str(row['conta'])
                celular = str(row['celular']).rstrip('.0')
                nome = str(row['nome'])
                
                start_time = time.time()
                credito = rodar_rotina(index, cpf, agencia, conta)
                consulta_time = time.time() - start_time
                
                lista.at[index, 'Margem'] = credito
                if credito == 'Com Crédito ou Renovação':
                    lista.at[index, 'disparado'] = automate_whatsapp(browser,celular,nome,cpf)
                    
                lista.to_excel(nome_do_arquivo, index=False)

                
                total_consultas += 1
                tempos_de_consulta.append(consulta_time)
                print(f"Consulta {total_consultas} completa. Duração: {consulta_time:.2f} segundos.")
                
                # Gravando o tempo de consulta no arquivo e forçando a gravação imediata
                file.write(f'Consulta {total_consultas}: {consulta_time:.2f} segundos\n')
                file.flush()
                
                # Esperar de 100 a 160 segundos a cada 20 consultas
                if total_consultas % 20 == 0:
                    logoff()
                    pause_time = random.uniform(100,140)  # Calcula o tempo de pausa
                    for remaining in range(int(pause_time), 0, -1):  # Loop regressivo
                        print(f"Faltam {remaining} segundos para retomar.")
                        time.sleep(1)
                    login()
                    time.sleep(15)
                    print("Checando login")
                    x = checalog()
                    if x == 1:
                        print("login bem sucedido")
                        x = 130
                        y = 230
                        x += random.randint(-10, 10)
                        y += random.randint(-4, 4)
                        pyautogui.click(x, y)
                        time.sleep(random.uniform(3, 4))

                    else:
                        print("erro no login")
                        time.sleep(random.uniform(2, 3))
                        login()
                        time.sleep(15)
                        print("Checando login")
                        x = checalog()
                        if x == 1:
                            print("login bem sucedido")
                            x = 130
                            y = 230
                            x += random.randint(-10, 10)
                            y += random.randint(-4, 4)
                            pyautogui.click(x, y)
                            time.sleep(random.uniform(3, 4))
                        else:
                            breakpoint
                        
                                
        if total_consultas > 0:
            duracao_media = sum(tempos_de_consulta) / total_consultas
            print(f"Total de consultas realizadas: {total_consultas}")
            print(f"Duração média por consulta: {duracao_media:.2f} segundos")
            # Gravando a duração média no arquivo
            file.write(f"Duração média por consulta: {duracao_media:.2f} segundos\n")
            file.flush()
        else:
            print("Nenhuma consulta realizada.")
            file.write("Nenhuma consulta realizada.\n")
            file.flush()
        

while True:
    profile_directory = r'C:\Users\arima\AppData\Local\Microsoft\Edge\User Data\Profile 2'
    edge_options = Options()
    edge_options.add_argument('--start-maximized')
    edge_options.add_argument('--disable-infobars')
    edge_options.add_argument('--disable-extensions')
    edge_options.add_argument('--disable-gpu')
    edge_options.add_argument('--disable-dev-shm-usage')
    edge_options.add_argument('--no-sandbox')
    edge_options.add_argument('--enable-chrome-browser-cloud-management')
    edge_options.add_argument(f'--user-data-dir={profile_directory}')
    browser = webdriver.Edge(options=edge_options)
    browser.get('https://web.whatsapp.com/')
    print("Escolha uma opção:")
    print("1 - Criar rotina")
    print("2 - Rodar rotina")
    print("3 - Consultar")
    choice = input("Digite o número correspondente à sua escolha e pressione Enter: ").strip()
    if choice == '1':
        criar_rotina()
    elif choice == '2':
        #rodar_rotina()
        print("checar")
    elif choice == '3':
        consultar(browser)
    else:
        print("Opção inválida, tente novamente.")
