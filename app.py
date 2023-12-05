from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import *
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from time import sleep
from openpyxl import load_workbook

#configuração do Chrome
def iniciar_driver():
    chrome_options = Options()
    arguments = ['--lang=pt-BR', 'start-maximized', '--incognito']
    for argument in arguments:
        chrome_options.add_argument(argument)

    chrome_options.add_experimental_option('prefs', {
        'download.prompt_for_download': False,
        'profile.default_content_setting_values.notifications': 2,
        'profile.default_content_setting_values.automatic_downloads': 1,

    })
    driver = webdriver.Chrome(service=ChromeService(
        ChromeDriverManager().install()), options=chrome_options)

    return driver
#acessando Google Maps / reconhecendo os campos necessários para carregar rotas
def acessar_valores():
    driver = iniciar_driver()
    #navegar até o site
    driver.get('https://www.google.com.br/maps/preview')
    botao_rotas = driver.find_element(By.XPATH,'//*[@id="hArJGc"]').click()
    sleep(3)
    campo_de= driver.find_element(By.XPATH,'//*[@id="sb_ifc50"]/input')
    campo_para = driver.find_element(By.XPATH,'//*[@id="sb_ifc51"]/input')
    arquivo= load_workbook('C:/Users/2160011883/Desktop/base_triagem_teste2.xlsx',data_only=True)
    
    #nomeando as variáveis que vão representar as colunas que representam as rotas
    sheet= arquivo['F_ME5A']
    coluna_prestador = 'M'
    conta_coluna_loja = sheet['L']
    coluna_loja = 'L'
    coluna_km = 'O'
    contador =2 
    cont=2

    for celula1 in conta_coluna_loja:
     contador+= 1
        


    for celula in range(1,contador):
        #acessando valores da base
        celula_ativa_prestador = sheet[coluna_prestador+str(cont)].value
        celula_ativa_loja = sheet[coluna_loja+str(cont)].value 
        celula_ativa_km =sheet[coluna_km+str(cont)]   

        
        if celula_ativa_loja and celula_ativa_prestador != None:        
           campo_de.send_keys(str(celula_ativa_prestador))
           campo_para.send_keys(str(celula_ativa_loja))
           sleep(1)
           campo_carro = driver.find_element(By.XPATH,'//*[@id="omnibox-directions"]/div/div[2]/div/div/div/div[2]/button').click()
           sleep(1)
           pesquisa = driver.find_element(By.XPATH,'//*[@id="omnibox-directions"]/div/div[3]/div[2]/button').click()
           sleep(1)
        
           try:
               km = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="section-directions-trip-0"]/div[1]/div/div[1]/div[2]/div')))          
                        
               km_get = km.text 
               #preenchendo a coluna de KM                  
               cel = sheet[coluna_km+str(cont)]
               cel.value = km_get

               print(f'{contador} {celula_ativa_loja} {celula_ativa_prestador} {km_get}')

                    
               sleep(1)
               campo_de= driver.find_element(By.XPATH,'//*[@id="sb_ifc50"]/input')
               campo_de.clear()
               campo_para = driver.find_element(By.XPATH,'//*[@id="sb_ifc51"]/input')
               campo_para.clear()
           except StaleElementReferenceException:
               print('endereço não encontrado!')
           except TimeoutException:
               print('endereço não encontrado!')

           finally:                            
               campo_de.clear()
               campo_para = driver.find_element(By.XPATH,'//*[@id="sb_ifc51"]/input')
               campo_para.clear()                 
           
            
        else:
             pass  
           
            

        cont += 1
        #salvando arquivo de planilha com a coluna de KM preenchida
        arquivo.save('C:/Users/2160011883/Desktop/base_triagem_teste2.xlsx')
        arquivo.close()


acessar_valores()








