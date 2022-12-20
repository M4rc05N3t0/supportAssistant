import os
import smtplib, email.message
from PyQt5.QtWidgets import QMessageBox
import subprocess
from time import sleep

import desing2


def testeDeConectividade():# Valida se a maquina tem conexao com a internet:
    internet_teste = str(os.popen('ping google.com.br -n 1').read())
    if "Verifique" in internet_teste or "check" in internet_teste:
        connect = False
        return connect
    else:
        connect = True
        return connect

def testeRede():
    #comando que vai rodar no CMD:
    command = "ping vpn.escale.com.br -n 30"
    #Roda o comando no CMD e coleta as informaçoes do mesmo:
    teste_completo = str(os.popen(command).read())

    #Variaveis usadas para realizar o diagnostico do teste:
    len_teste_completo = int(len(teste_completo)) # --> variavel que identifica o comprimento do teste completo
    resultado = str(teste_completo[-225:len_teste_completo]) # --> variavel que captura apenas o final do teste, que mostra a perda de pacotes e o MS
    index_percent = int(resultado.index("%")) # --> variavel que identifica o index do % no teste para poder, essa informaçao é usada para odentificar a perda de pacotes

    #Soluçao caso a maquina esteja em Ingles:
    if "Average" in resultado:
        resultado = teste_completo[-194:len_teste_completo]
        index_percent = int(resultado.index("%"))  # --> variavel que identifica o index do % no teste para poder, essa informaçao é usada para odentificar a perda de pacotes
        if resultado[index_percent - 2] == "(":
            packet_loss_rate = int(resultado[index_percent - 1:index_percent])
        elif resultado[index_percent - 3] == "(":
            packet_loss_rate = int(resultado[index_percent - 2:index_percent])
        else:
            packet_loss_rate = int(resultado[index_percent - 3:index_percent])
    else:
    #Mostra o resultado do teste de PING:
    #Condicional que vai validar a porcentagem de perda de pacotes no resultado do teste:
        if resultado[index_percent-2] == "(":
            packet_loss_rate = int(resultado[index_percent-1:index_percent])
        elif resultado[index_percent-3] == "(":
            packet_loss_rate = int(resultado[index_percent-2:index_percent])
        else:
            packet_loss_rate = int(resultado[index_percent-3:index_percent])
    #Encontra a latencia dentro da String "resultado":
    lista_resultado = resultado.split()
    lista_latencia = [item for item in lista_resultado if "ms" in item]
    #Retira o valor numerico da latencia da conexao e adiciona os valores em variaveis:
    lista_latencia2 =[]
    for item in lista_latencia:
        item = item.replace(" ","")
        item = item.replace(",","")
        item = item.replace("ms","")
        lista_latencia2.append(int(item))
    latencia_minima = lista_latencia2[0]
    latencia_maxima = lista_latencia2[1]
    latencia_media = lista_latencia2[2]
    lista_progress = 70
    #Retorna as informaçoes necessarias para realiar o diagnostico:
    return packet_loss_rate, latencia_minima, latencia_media, latencia_maxima, lista_progress, resultado


def diagnosticoHardware():
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    lista_diagnostico = []
    lista_nome = ["Teste Windows Update","Teste Apps","Teste Audio","Teste Bits","Teste Bluetooth","Teste Device","Teste Keyboard","Teste Printer","Teste Speech","Teste Video"]
    teste_windows_update = str((subprocess.run(["powershell", "Get-TroubleshootingPack -Path 'C:\Windows\Diagnostics\System\WindowsUpdate' | Invoke-TroubleshootingPack"],capture_output=True, startupinfo=startupinfo, creationflags=subprocess.SW_HIDE)))
    lista_diagnostico.append(teste_windows_update)
    teste_apps = str((subprocess.run(["powershell", "Get-TroubleshootingPack -Path 'C:\Windows\Diagnostics\System\Apps' | Invoke-TroubleshootingPack"],capture_output=True, startupinfo=startupinfo, creationflags=subprocess.SW_HIDE)))
    lista_diagnostico.append(teste_apps)
    teste_audio = str((subprocess.run(["powershell", "Get-TroubleshootingPack -Path 'C:\Windows\Diagnostics\System\Audio' | Invoke-TroubleshootingPack"],capture_output=True, startupinfo=startupinfo, creationflags=subprocess.SW_HIDE)))
    lista_diagnostico.append(teste_audio)
    teste_bits = str((subprocess.run(["powershell", "Get-TroubleshootingPack -Path 'C:\Windows\Diagnostics\System\BITS' | Invoke-TroubleshootingPack"],capture_output=True, startupinfo=startupinfo, creationflags=subprocess.SW_HIDE)))
    lista_diagnostico.append(teste_bits)
    teste_bluetooth = str((subprocess.run(["powershell", "Get-TroubleshootingPack -Path 'C:\Windows\Diagnostics\System\Bluetooth' | Invoke-TroubleshootingPack"],capture_output=True, startupinfo=startupinfo, creationflags=subprocess.SW_HIDE)))
    lista_diagnostico.append(teste_bluetooth)
    teste_device = str((subprocess.run(["powershell", "Get-TroubleshootingPack -Path 'C:\Windows\Diagnostics\System\Device' | Invoke-TroubleshootingPack"],capture_output=True, startupinfo=startupinfo, creationflags=subprocess.SW_HIDE)))
    lista_diagnostico.append(teste_device)
    teste_keyboard = str((subprocess.run(["powershell", f"Get-TroubleshootingPack -Path 'C:\Windows\Diagnostics\System\Keyboard' | Invoke-TroubleshootingPack "],capture_output=True, startupinfo=startupinfo, creationflags=subprocess.SW_HIDE)))
    lista_diagnostico.append(teste_keyboard)
    teste_printer = str((subprocess.run(["powershell", f"Get-TroubleshootingPack -Path 'C:\Windows\Diagnostics\System\Printer' | Invoke-TroubleshootingPack "],capture_output=True, startupinfo=startupinfo, creationflags=subprocess.SW_HIDE)))
    lista_diagnostico.append(teste_printer)
    teste_speech = str((subprocess.run(["powershell", f"Get-TroubleshootingPack -Path 'C:\Windows\Diagnostics\System\Speech' | Invoke-TroubleshootingPack "],capture_output=True, startupinfo=startupinfo, creationflags=subprocess.SW_HIDE)))
    lista_diagnostico.append(teste_speech)
    teste_video = str((subprocess.run(["powershell", f"Get-TroubleshootingPack -Path 'C:\Windows\Diagnostics\System\Video' | Invoke-TroubleshootingPack "],capture_output=True, startupinfo=startupinfo, creationflags=subprocess.SW_HIDE)))
    lista_diagnostico.append(teste_video)
    return lista_nome, lista_diagnostico

def eventBotoonHardware():
    lista_nome, lista_diagnostico = diagnosticoHardware()
    lista_completa = []
    x = 0
    while x <= 9:
        if 'Nenhum problema foi detectado' in lista_diagnostico[x] or 'No problems were detected' in lista_diagnostico[x]:
            lista_completa.append(f'{lista_nome[x]}     --------->   Nenhum erro encontrado')
        else:
            lista_completa.append(f'Os seguintes erros foram encontrados:\n{lista_diagnostico[x]}\n Se necessario, contate o time de Suporte')

        x += 1

    return lista_completa


def sendEmail(nome, mail, texto, titulo):
    email_body = texto
    msg = email.message.Message()
    msg['Subject'] = f'{nome} {titulo}'
    msg['From'] = 'enviodeemailtestepython@gmail.com'
    msg['To'] = mail
    password = '?Z/FJs7\Knw&2Q{m'#?Z/FJs7\Knw&2Q{m
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(email_body)
    s = smtplib.SMTP('smtp.gmail.com:587')
    s.starttls()
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))

def cleanDesktop():
    comandoLimpeza = 'ipconfig /flushdns'#del /q/f/s %temp%*
    return os.popen(comandoLimpeza).read()


def cleanDNS():
    commandClean = 'ipconfig /flushdns'
    return os.popen(commandClean).read()


def updatePolice():
    atualizaPolice = str(os.popen('gpupdate /force').read())
    if not 'conclu¡da com ˆxito' in atualizaPolice:
        return False
    else:
        return True

def alert(title, text):
    msg = QMessageBox()
    msg.setWindowTitle(title)
    msg.setText(text)
    aviso = msg.exec_()

    return aviso



def vpnOrError():
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    var = str((subprocess.run(
        ["powershell", "Get-NetAdapter -InterfaceDescription 'Fortinet SSL VPN Virtual Ethernet*'"],
        capture_output=True, startupinfo=startupinfo, creationflags=subprocess.SW_HIDE)))
    if "Up" in var:
        vpn_value = True
    else:
        vpn_value = False
    return vpn_value



