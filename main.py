import time
import sys
import os
import calendar
import tkinter as tk
from tkinter import simpledialog

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys

# Cria a janela Tkinter e a oculta
root = tk.Tk()
root.withdraw()

# Solicita ao usuário a competência (formato MM/YYYY)
input_date = simpledialog.askstring("Competência de Faturamento", "Informe a competência e ano (MM/YYYY):")
if not input_date:
    print("Nenhuma data informada. Encerrando...")
    exit(1)

try:
    month_str, year_str = input_date.split("/")
    month = int(month_str)
    year = int(year_str)
    
    # Define a data inicial como o primeiro dia do mês
    first_date = f"01/{month:02d}/{year}"
    
    # Determina o último dia do mês
    last_day = calendar.monthrange(year, month)[1]
    last_date = f"{last_day:02d}/{month:02d}/{year}"
    
    print("Data Inicial:", first_date)
    print("Data Final:", last_date)
except Exception as e:
    print("Erro ao processar a data:", e)
    exit(1)

# Solicita ao usuário o login e a senha
login = simpledialog.askstring("Login", "Informe seu login:")
password = simpledialog.askstring("Senha", "Informe sua senha:", show='*')
if not login or not password:
    print("Login ou senha não informados. Encerrando...")
    exit(1)

# Agora que os dados foram informados, inicia o Chrome

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS  # PyInstaller extrai arquivos temporários aqui
else:
    base_path = os.path.abspath(".")

service = Service(os.path.join(base_path, "chromedriver-win64", "chromedriver.exe"))
driver = webdriver.Chrome(service=service)

# Acessa a página de login
driver.get('http://c3189prd.cloudmv.com.br/mvsso/login?service=http%3A%2F%2Fc3189prd.cloudmv.com.br%2Fsigss%2Fauthentication%2Fcallback%3Fclient_name%3DCasClient')

# Preenche os campos de login e senha com os dados informados
try:
    username_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "username"))
    )
    username_input.clear()
    username_input.send_keys(login)
    print("Login preenchido com sucesso!")
except Exception as e:
    print("Erro ao preencher o campo de login:", e)

try:
    senha_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "password"))
    )
    senha_input.clear()
    senha_input.send_keys(password)
    print("Senha preenchida com sucesso!")
except Exception as e:
    print("Erro ao preencher o campo de senha:", e)

# Efetua o login (pressiona TAB e ENTER)
senha_input.send_keys(Keys.TAB)
senha_input.send_keys(Keys.ENTER)
time.sleep(1)


# Clicar em "ADMINISTRATIVO"
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//li[span[contains(text(), 'ADMINISTRATIVO')]]"))
).click()

# Esperar a página principal carregar
time.sleep(1)

# Clicar no botão do menu lateral
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "button.sigss-button.sigss-viewport-handler"))
).click()

# Clicar na opção "Sistema"
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//span[text()='Sistema']"))
).click()

# Clicar na opção "Sist. Min. da Saúde"
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//span[text()='Sist. Min. da Saúde']"))
).click()

# Clicar na opção "Faturamento SIA - Ajustes e Configurações"
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//span[text()='Faturamento SIA - Ajustes e Configurações']"))
).click()

# ✅ Marcar o checkbox
try:
    checkbox = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "cb_grid_unsa_regi"))
    )
    checkbox.click()
    print("Checkbox marcado com sucesso!")
except Exception as e:
    print("Erro ao marcar checkbox:", e)

# ✅ Preencher o campo de data inicial via JavaScript
try:
    dt_inicial = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "dtInicialRegi"))
    )
    # Remover o atributo readonly (caso exista)
    driver.execute_script("arguments[0].removeAttribute('readonly')", dt_inicial)
    # Definir o valor diretamente
    driver.execute_script("arguments[0].value = arguments[1];", dt_inicial, first_date)
    # Disparar o evento 'change' para garantir que a alteração seja capturada
    driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", dt_inicial)
    print("Data inicial definida:", first_date)
except Exception as e:
    print("Erro ao definir data inicial:", e)

# ✅ Preencher o campo de data final via JavaScript
try:
    dt_final = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "dtFinalRegi"))
    )
    # Remover o atributo readonly (caso exista)
    driver.execute_script("arguments[0].removeAttribute('readonly')", dt_final)
    # Definir o valor diretamente
    driver.execute_script("arguments[0].value = arguments[1];", dt_final, last_date)
    # Disparar o evento 'change' para garantir que a alteração seja capturada
    driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", dt_final)
    print("Data final definida:", last_date)
except Exception as e:
    print("Erro ao definir data final:", e)


# ✅ Abrir menu suspenso clicando no <a> com `class="chzn-single"`
try:
    dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//a[@class='chzn-single']"))
    )
    dropdown.click()

    # Digitar "AGENDAMENTO CONSULTA" e pressionar ENTER
    dropdown_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//div[@class='chzn-search']/input"))
    )
    dropdown_input.send_keys("AGENDAMENTO CONSULTA")
    dropdown_input.send_keys(Keys.ENTER)
    print(f"Iniciando Ajustes no agendamento de consultas'")
    # Aguardar 6 segundos
    time.sleep(6)
except Exception as e:
    print("Erro ao selecionar 'AGENDAMENTO CONSULTA':", e)
    

# Iterar sobre todas as linhas da tabela
try:
    # Esperar até a tabela estar presente e pegar todas as linhas dentro dela
    tabela = WebDriverWait(driver, 7).until(
        EC.presence_of_element_located((By.ID, "grid_prci_regi"))
    )
    linhas_tabela = tabela.find_elements(By.XPATH, ".//tr[@role='row']")
    print(f"Encontradas {len(linhas_tabela)} linhas para edição.")

    for i, linha in enumerate(linhas_tabela):
        try:
            # Encontrar o botão "Alterar" dentro da linha atual
            botao_alterar = WebDriverWait(linha, 1).until(
                EC.element_to_be_clickable((By.XPATH, ".//button[contains(@class, 'btnAlterar')]"))
            )
            botao_alterar.click()  # Clicar no botão "Alterar"

            # Esperar o menu suspenso carregar
            menu_suspenso = WebDriverWait(driver, 1).until(
                EC.visibility_of_element_located((By.ID, "regi_regiId_chzn"))
            )
            menu_suspenso.click()  # Clicar no menu suspenso

            # Encontrar o campo de pesquisa e buscar por "BPA (INDIVIDUALIZADO)"
            campo_pesquisa = WebDriverWait(driver, 1).until(
                EC.visibility_of_element_located((By.ID, "regi_regiId_chzn_search"))
            )
            campo_pesquisa.clear()  # Limpar o campo de pesquisa
            campo_pesquisa.send_keys("BPA (INDIVIDUALIZADO)")  # Digitar "BPA (INDIVIDUALIZADO)"
            time.sleep(1)  # Esperar para garantir que a pesquisa seja feita

            # Tentar selecionar o item "BPA (INDIVIDUALIZADO)"
            try:
                item_individualizado = WebDriverWait(driver, 1).until(
                    EC.element_to_be_clickable((By.XPATH, "//li[@title='BPA (INDIVIDUALIZADO)']"))
                )
                item_individualizado.click()
                print(f"[{i+1}] Item 'BPA (INDIVIDUALIZADO)' selecionado!")

                # Clicar no botão de "Gravar"
                try:
                    # Esperar até o botão "Gravar" estar visível e habilitado
                    botao_gravar = WebDriverWait(driver, 1).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[@class='ui-button ui-widget ui-state-default ui-corner-all ui-button-text-icon-primary']/span[text()='Gravar']"))
                    )
                    botao_gravar.click()
                    # Esperar até o botão "Gravar" desaparecer (garante que foi gravado)
                    WebDriverWait(driver, 1).until_not(
                        EC.presence_of_element_located((By.XPATH, "//button/span[text()='Gravar']"))
                    )
                except Exception as e:
                    print(f"Erro ao clicar no botão 'Gravar'")

            except Exception as e:
                # Se não encontrar "BPA (INDIVIDUALIZADO)", clicar no botão "Cancelar"
                print(f"[{i+1}] 'BPA (INDIVIDUALIZADO)' não encontrado. Clicando no botão 'Cancelar'.")
                botao_cancelar = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, "//button/span[text()='Cancelar']"))
                )
                botao_cancelar.click()
                time.sleep(2)  # Esperar o cancelamento para o próximo item
                
                 # Clicar no botão "OK" após a mensagem de sucesso
            try:
                botao_ok = WebDriverWait(driver, 1).until(
                    EC.element_to_be_clickable((By.XPATH, "//button/span[text()='OK']"))
                )
                botao_ok.click()
            except Exception as e:
                print("Erro ao clicar no botão 'OK':", e)

            # Esperar até o botão OK sumir (garante que foi confirmado)
            WebDriverWait(driver, 1).until_not(
                EC.presence_of_element_located((By.XPATH, "//span[text()='OK']"))
            )

        except Exception as e:
            print(f"Erro ao processar linha {i+1}: {e}")

    print(f"Fim dos ajustes de consultas")
except Exception as e:
    print("Erro ao processar tabela:", e)


# ✅ Abrir menu suspenso clicando no <a> com `class="chzn-single"`
try:
    dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//a[@class='chzn-single']"))
    )
    dropdown.click()

    # Digitar "AGENDAMENTO EXAME" e pressionar ENTER
    dropdown_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//div[@class='chzn-search']/input"))
    )
    dropdown_input.send_keys("AGENDAMENTO EXAME")
    dropdown_input.send_keys(Keys.ENTER)
    print(f"Iniciando os ajustes de exames")
    # Aguardar 20 segundos antes de começar a alterar
    print("Aguardando 20 segundos antes de iniciar as alterações...")
    time.sleep(20)
except Exception as e:
    print("Erro ao selecionar 'AGENDAMENTO EXAME':", e)

# Iterar sobre todas as linhas da tabela
try:
    # Esperar até a tabela estar presente e pegar todas as linhas dentro dela
    tabela = WebDriverWait(driver,15).until(
        EC.presence_of_element_located((By.ID, "grid_prci_regi"))
    )
    linhas_tabela = tabela.find_elements(By.XPATH, ".//tr[@role='row']")
    print(f"Encontradas {len(linhas_tabela)} linhas para edição.")

    for i, linha in enumerate(linhas_tabela):
        try:
            # Encontrar o botão "Alterar" dentro da linha atual
            botao_alterar = WebDriverWait(linha, 1).until(
                EC.element_to_be_clickable((By.XPATH, ".//button[contains(@class, 'btnAlterar')]"))
            )
            botao_alterar.click()  # Clicar no botão "Alterar"

            # Esperar o menu suspenso carregar
            menu_suspenso = WebDriverWait(driver, 1).until(
                EC.visibility_of_element_located((By.ID, "regi_regiId_chzn"))
            )
            menu_suspenso.click()  # Clicar no menu suspenso

            # Encontrar o campo de pesquisa e buscar por "BPA (INDIVIDUALIZADO)"
            campo_pesquisa = WebDriverWait(driver, 1).until(
                EC.visibility_of_element_located((By.ID, "regi_regiId_chzn_search"))
            )
            campo_pesquisa.clear()  # Limpar o campo de pesquisa
            campo_pesquisa.send_keys("BPA (INDIVIDUALIZADO)")  # Digitar "BPA (INDIVIDUALIZADO)"
            time.sleep(1)  # Esperar para garantir que a pesquisa seja feita

            # Tentar selecionar o item "BPA (INDIVIDUALIZADO)"
            try:
                item_individualizado = WebDriverWait(driver, 1).until(
                    EC.element_to_be_clickable((By.XPATH, "//li[@title='BPA (INDIVIDUALIZADO)']"))
                )
                item_individualizado.click()
                print(f"[{i+1}] Item 'BPA (INDIVIDUALIZADO)' selecionado!")

                # Clicar no botão de "Gravar"
                try:
                    # Esperar até o botão "Gravar" estar visível e habilitado
                    botao_gravar = WebDriverWait(driver, 1).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[@class='ui-button ui-widget ui-state-default ui-corner-all ui-button-text-icon-primary']/span[text()='Gravar']"))
                    )
                    botao_gravar.click()
                    # Esperar até o botão "Gravar" desaparecer (garante que foi gravado)
                    WebDriverWait(driver, 1).until_not(
                        EC.presence_of_element_located((By.XPATH, "//button/span[text()='Gravar']"))
                    )
                except Exception as e:
                    print(f"Erro ao clicar no botão 'Gravar'")

            except Exception as e:
                # Se não encontrar "BPA (INDIVIDUALIZADO)", clicar no botão "Cancelar"
                print(f"[{i+1}] 'BPA (INDIVIDUALIZADO)' não encontrado. Clicando no botão 'Cancelar'.")
                botao_cancelar = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, "//button/span[text()='Cancelar']"))
                )
                botao_cancelar.click()
                time.sleep(2)  # Esperar o cancelamento para o próximo item
                
                 # Clicar no botão "OK" após a mensagem de sucesso
            try:
                botao_ok = WebDriverWait(driver, 1).until(
                    EC.element_to_be_clickable((By.XPATH, "//button/span[text()='OK']"))
                )
                botao_ok.click()
            except Exception as e:
                print("Erro ao clicar no botão 'OK':", e)

            # Esperar até o botão OK sumir (garante que foi confirmado)
            WebDriverWait(driver, 1).until_not(
                EC.presence_of_element_located((By.XPATH, "//span[text()='OK']"))
            )

        except Exception as e:
            print(f"Erro ao processar linha {i+1}: {e}")
    print(f"Fim dos ajustes de exames")

except Exception as e:
    print("Erro ao processar tabela:", e)
    

# ✅ Abrir menu suspenso clicando no <a> com `class="chzn-single"`
try:
    dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//a[@class='chzn-single']"))
    )
    dropdown.click()

    # Digitar "PROCEDIMENTO AMBULATORIAL" e pressionar ENTER
    dropdown_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//div[@class='chzn-search']/input"))
    )
    dropdown_input.send_keys("PROCEDIMENTO AMBULATORIAL")
    dropdown_input.send_keys(Keys.ENTER)
    print(f"Iniciando ajustes de procedimentos")
    # Aguardar 10 segundos antes de começar a alterar
    print("Aguardando 10 segundos antes de iniciar as alterações...")
    time.sleep(10)
except Exception as e:
    print("Erro ao selecionar 'PROCEDIMENTO AMBULATORIAL':", e)

# Iterar sobre todas as linhas da tabela
try:
    # Esperar até a tabela estar presente e pegar todas as linhas dentro dela
    tabela = WebDriverWait(driver, 7).until(
        EC.presence_of_element_located((By.ID, "grid_prci_regi"))
    )
    linhas_tabela = tabela.find_elements(By.XPATH, ".//tr[@role='row']")
    print(f"Encontradas {len(linhas_tabela)} linhas para edição.")

    for i, linha in enumerate(linhas_tabela):
        try:
            # Encontrar o botão "Alterar" dentro da linha atual
            botao_alterar = WebDriverWait(linha, 1).until(
                EC.element_to_be_clickable((By.XPATH, ".//button[contains(@class, 'btnAlterar')]"))
            )
            botao_alterar.click()  # Clicar no botão "Alterar"

            # Esperar o menu suspenso carregar
            menu_suspenso = WebDriverWait(driver, 1).until(
                EC.visibility_of_element_located((By.ID, "regi_regiId_chzn"))
            )
            menu_suspenso.click()  # Clicar no menu suspenso

            # Encontrar o campo de pesquisa e buscar por "BPA (INDIVIDUALIZADO)"
            campo_pesquisa = WebDriverWait(driver, 1).until(
                EC.visibility_of_element_located((By.ID, "regi_regiId_chzn_search"))
            )
            campo_pesquisa.clear()  # Limpar o campo de pesquisa
            campo_pesquisa.send_keys("BPA (INDIVIDUALIZADO)")  # Digitar "BPA (INDIVIDUALIZADO)"
            time.sleep(1)  # Esperar para garantir que a pesquisa seja feita

            # Tentar selecionar o item "BPA (INDIVIDUALIZADO)"
            try:
                item_individualizado = WebDriverWait(driver, 1).until(
                    EC.element_to_be_clickable((By.XPATH, "//li[@title='BPA (INDIVIDUALIZADO)']"))
                )
                item_individualizado.click()
                print(f"[{i+1}] Item 'BPA (INDIVIDUALIZADO)' selecionado!")

                # Clicar no botão de "Gravar"
                try:
                    # Esperar até o botão "Gravar" estar visível e habilitado
                    botao_gravar = WebDriverWait(driver, 1).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[@class='ui-button ui-widget ui-state-default ui-corner-all ui-button-text-icon-primary']/span[text()='Gravar']"))
                    )
                    botao_gravar.click()
                    # Esperar até o botão "Gravar" desaparecer (garante que foi gravado)
                    WebDriverWait(driver, 1).until_not(
                        EC.presence_of_element_located((By.XPATH, "//button/span[text()='Gravar']"))
                    )
                except Exception as e:
                    print(f"Erro ao clicar no botão 'Gravar'")

            except Exception as e:
                # Se não encontrar "BPA (INDIVIDUALIZADO)", clicar no botão "Cancelar"
                print(f"[{i+1}] 'BPA (INDIVIDUALIZADO)' não encontrado. Clicando no botão 'Cancelar'.")
                botao_cancelar = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, "//button/span[text()='Cancelar']"))
                )
                botao_cancelar.click()
                time.sleep(2)  # Esperar o cancelamento para o próximo item
                
                 # Clicar no botão "OK" após a mensagem de sucesso
            try:
                botao_ok = WebDriverWait(driver, 1).until(
                    EC.element_to_be_clickable((By.XPATH, "//button/span[text()='OK']"))
                )
                botao_ok.click()
            except Exception as e:
                print("Erro ao clicar no botão 'OK':", e)

            # Esperar até o botão OK sumir (garante que foi confirmado)
            WebDriverWait(driver, 1).until_not(
                EC.presence_of_element_located((By.XPATH, "//span[text()='OK']"))
            )

        except Exception as e:
            print(f"Erro ao processar linha {i+1}: {e}")
    print(f"Fim dos ajustes de procedimentos")
except Exception as e:
    print("Erro ao processar tabela:", e)