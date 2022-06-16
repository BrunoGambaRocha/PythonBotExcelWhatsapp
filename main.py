from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import time
import urllib
import pandas as pd

contatos = pd.read_excel("Enviar.xlsx")

navegador = webdriver.Chrome(executable_path=r"C:\Users\Cliente\Documents\BrunoTI\Cursos\IntensivaoPython\programas\chromedriver.exe")
navegador.get("https://web.whatsapp.com/")
actions = ActionChains(navegador)

while len(navegador.find_elements_by_id("side")) < 1:
    time.sleep(1)

# já estamos com o login feito no whatsapp web
for i, mensagem in enumerate(contatos["Mensagem"]):
    pessoa = contatos.loc[i, "Pessoa"]
    numero = contatos.loc[i, "Número"]
    tipo = contatos.loc[i, "Tipo"]

    if tipo == "Individual":
        # através de um link do whatsapp, enviamos uma mensagem personalizada com o número e mensagem
        texto = urllib.parse.quote(f"Oi {pessoa}! {mensagem}")
        link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
        navegador.get(link)

        while len(navegador.find_elements_by_id("side")) < 1:
            time.sleep(1)

    elif (tipo == "Grupo") or (tipo == "Sem Contato"):
        while len(navegador.find_elements_by_id("side")) < 1:
            time.sleep(1)

        # aguarda o painel das conversas ser exibido
        paneSide = navegador.find_elements_by_id("pane-side")
        while len(paneSide) < 1:
            time.sleep(1)
            paneSide = navegador.find_elements_by_id("pane-side")

        # localiza a conversa pelo texto igual ao celular ou ao nome do grupo
        while len(navegador.find_elements_by_xpath(f"//span[@title='{pessoa}']")) < 1:
            time.sleep(1)
            print("buscando conversa")
            navegador.execute_script("arguments[0].scrollTop = (arguments[0].scrollTop + 400 >= arguments[0].scrollHeight) ? arguments[0].scrollTo(0,0) : arguments[0].scrollTop + 400;", paneSide[0])

        conversa = navegador.find_element_by_xpath(f"//span[@title='{pessoa}']")
        conversa.click()

        # testa se ja carregou a conversa do contato
        while not(navegador.find_element_by_xpath("//*[@id='main']/header/div[2]/div[1]/div").text == pessoa):
            time.sleep(1)
            print("aguardando a conversa carregar")

        # define o campo onde digitamos a mensagem na conversa, clica, digita a mensagem
        chat_box = navegador.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[1]/div")
        chat_box.click()
        chat_box.send_keys(mensagem)

        # testa se mensagem foi digitada corretamente
        while not(chat_box.text == mensagem):
            navegador.execute_script(f"var ele=arguments[0]; ele.innerHTML = '{mensagem}';", chat_box)
            time.sleep(1)
            print("aguardando validação mensagem")

    # clica no botão enviar mensagem na conversa
    navegador.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[2]/button").click()
    time.sleep(10)

navegador.quit()