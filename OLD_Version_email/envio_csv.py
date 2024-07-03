import time
import pyautogui as bot

#Abrir firefox
#escrever gmail.com e dar enter
#Clicar em 'compose'
#Endereçar para suportebrazil@blackbox.com
#subject: Printer Report
#Clicar em anexo e selecionar o arquivo
#Enviar email

#Configurações Pyautogui
bot.PAUSE = 0.5
bot.FAILSAFE = True

#Constantes export_csv
navegador = 'chrome'
email = 'gmail.com'
endereco = 'suportebrazil@blackbox.com'
subject = 'Printer Report'
busca = 'teste'


def export_csv():
    time.sleep(2)
    #bot.press('win')
    #bot.write(navegador)
    #bot.press('enter')
    #time.sleep(2)
    #bot.write(email)
    #bot.press('enter')
    #time.sleep(2)
    #bot.click('escrever_email.PNG')
    #time.sleep(1)
    bot.write(endereco)
    time.sleep(1)
    bot.press('tab', presses=2)
    time.sleep(1)
    bot.write(subject)
    bot.click('anexo.PNG')
    time.sleep(2)
    bot.click('desktop_grey.PNG')
    time.sleep(1)
    bot.click('pesquisa.PNG')
    bot.write(busca)
    time.sleep(2)
    bot.click(x=354, y=296, clicks=2)
    time.sleep(3)
    bot.click('enviar.PNG')

    return


export_csv()