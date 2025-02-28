import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
import sys
import os
import tkinter as tk
from tkinter import simpledialog
import calendar
# Cria a janela Tkinter e a oculta (não precisamos da janela principal)
root = tk.Tk()
root.withdraw()

# Solicita ao usuário a competência e o ano no formato MM/YYYY
input_date = simpledialog.askstring("Competência de Faturamento", "Informe a competência e ano (MM/YYYY):")

if input_date:
    try:
        # Separa o mês e o ano
        month_str, year_str = input_date.split("/")
        month = int(month_str)
        year = int(year_str)
        
        # Define a data inicial como o primeiro dia do mês
        first_date = f"01/{month:02d}/{year}"
        
        # Usa o módulo calendar para determinar o último dia do mês
        last_day = calendar.monthrange(year, month)[1]
        last_date = f"{last_day:02d}/{month:02d}/{year}"
        
        print("Data Inicial:", first_date)
        print("Data Final:", last_date)
        
        # Aqui você pode armazenar first_date e last_date para uso posterior na automação dos datepickers
        
    except Exception as e:
        print("Erro ao processar a data:", e)
        exit(1)
else:
    print("Nenhuma data informada. Encerrando...")
    exit(1)


if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS  # PyInstaller extrai arquivos temporários aqui
else:
    base_path = os.path.abspath(".")

service = Service(os.path.join(base_path, "chromedriver-win64", "chromedriver.exe"))


driver = webdriver.Chrome(service=service)

# Acessar o sistema
driver.get('http://c3189prd.cloudmv.com.br/mvsso/login?service=http%3A%2F%2Fc3189prd.cloudmv.com.br%2Fsigss%2Fauthentication%2Fcallback%3Fclient_name%3DCasClient')


# Localizar e preencher os campos de usuário e senha
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "username"))).send_keys("SANTANA")
senha_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "password")))
senha_input.send_keys("*261597Vg")

# Pressionar TAB para mover para o botão e ENTER para efetuar o login
senha_input.send_keys(Keys.TAB)
senha_input.send_keys(Keys.ENTER)

# Esperar a página trocar para o ponto de acesso
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

# ✅ Clicar no primeiro datepicker e aguardar 3 segundos
try:
    first_datepicker = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//img[@class='ui-datepicker-trigger']"))
    )
    first_datepicker.click()
    time.sleep(4)
except Exception as e:
    print("Erro ao clicar no primeiro datepicker:", e)

# ✅ Clicar no segundo datepicker e aguardar 3 segundos
try:
    second_datepicker = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "(//img[@class='ui-datepicker-trigger'])[2]"))
    )
    second_datepicker.click()
    time.sleep(4)
except Exception as e:
    print("Erro ao clicar no segundo datepicker:", e)

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