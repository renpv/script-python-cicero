import getpass
from modulos import config
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep

def run(driver: webdriver):
    print(f"\nFazendo login em {config.BASE_URL}")
    driver.get(config.BASE_URL + '/sipac')
    if(config.USER_LOGIN in ['', None]):
        config.USER_LOGIN = input('Informe seu usuario para login: ')
    el_user_login = driver.find_element(By.NAME, 'login')
    el_user_login.send_keys(config.USER_LOGIN)

    if(config.USER_PASS in ['', None]):
        config.USER_PASS = getpass.getpass(
            prompt='Informe sua senha para login:', stream=None)
    el_user_senha = driver.find_element(By.NAME, 'senha')
    el_user_senha.send_keys(config.USER_PASS)

    el_form_submit = driver.find_element(
        By.XPATH, '//*[@id="conteudo"]/div[3]/form/table/tfoot/tr/td/input')
    el_form_submit.click()
    
    sleep(2)

    if len(driver.find_elements(By.CLASS_NAME, 'sair-sistema')) == 0 :
        return False
    
    print(f'\nLogin realizado com sucesso. Usuário: {config.USER_LOGIN}\n')
    return True
