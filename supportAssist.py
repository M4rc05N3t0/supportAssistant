#coding: Latin1
import sys
from PyQt5.QtWidgets import QMainWindow, QApplication
from PyQt5.QtCore import Qt
import random
import diverseFunction
from desing2 import *
import socket, getpass, os
from time import sleep
import os.path
import re
import ApiRequests
import datetime
import config_infra
from usuario_planilha_ramal import Dados


class EscaleSuport(QMainWindow, Ui_MainWindow):
    def __init__(self, parent = None):
        super().__init__(parent)
        super().setupUi(self)
        self.progressBar.close()
        self.conf = config_infra.InfraSIP()
        self.dados = Dados()
        self.frame_15.close()
        usuario = str(getpass.getuser())
        self.nameUser.setText(usuario)
        self.btn_Rede.clicked.connect(self.eventoTesteRede)
        self.btn_DiagnosticoHardware.clicked.connect(self.eventoDiagnosticoHardare)
        self.btn_InfoMaquina.clicked.connect(self.eventoInsertInfoMaquina)
        self.btn_InfoEnvio.clicked.connect(self.eventoSendEmail)
        self.btn_abrirChamado.clicked.connect(self.eventoAbrindoChamado)
        self.btn_enviarHelp.clicked.connect(self.eventoAjuda)
        self.btn_limpeza.clicked.connect(self.eventoLimpeza)
        self.btn_atualizaMaquina.clicked.connect(self.eventoAtualizandoPolice)
        self.r_sim.clicked.connect(self.eventoIndentificar)
        self.r_nao.clicked.connect(self.eventoNaoindentificar)
        self.btn_conf_infra.clicked.connect(self.conf_infra)
        self.btn_aut_enviar.clicked.connect(self.inicia_aut)
        self.input_api()


    def eventoTesteRede(self):
        self.info.setPlainText('')
        connect = diverseFunction.testeDeConectividade()
        if connect == False:
            diverseFunction.alert('Erro', 'Por favor verifique a sua conexão')
            self.info.setPlainText('Por favor verifique a sua conexão com a internet para prosseguir com o teste!')
        else:
            diverseFunction.alert('Aguarde!', 'Aguarde até que conclua o teste')
            self.info.setPlainText('O teste está sendo realizado')
            self.progressBar.show()
            self.progressBar.setValue(5)
            packet_loss_rate, latencia_minima, latencia_media, latencia_maxima, lista_progress, resultado = diverseFunction.testeRede()
            self.info.setPlainText('')
            if connect == True:
                # Retorna a taxa de perda de pacotes na conexao e analisa se está alta ou baixa
                self.info.insertPlainText(f"Taxa de perda de pacotes: {packet_loss_rate}%")
                if packet_loss_rate >= 5:
                    self.info.insertPlainText(
                        "\nSua conexao apresenta uma taxa alta de perda de pacotes, isso pode prejudicar seu acesso aos sistemas e servicos da empresa \n")
                else:
                    self.info.insertPlainText("\nSua conexao nao apresenta uma taxa de perda de pacotes alta\n")
                self.progressBar.setValue(20)
                sleep(3)
                # Retorna a taxa de MS media na conexao e analisa se está alta ou baixa
                self.info.insertPlainText(f"\nLatencia media: {latencia_media}ms")
                if latencia_media >= 150:
                    self.info.insertPlainText(
                        "\nSua conexao apresenta uma latencia media muito alta, isso pode prejudicar seu acesso aos sistemas e servicos da empresa\n")
                else:
                    self.info.insertPlainText("\nSua latencia media está ok\n")
                self.progressBar.setValue(50)
                sleep(3)
                self.progressBar.setValue(80)
                sleep(3)
                self.info.insertPlainText(f"\nlatencia maxima detectada: {latencia_maxima}ms")
                if latencia_maxima >= 200:
                    self.info.insertPlainText(
                        "\nSua latencia maxima está muito alta, isso pode prejudicar seu acesso aos sistemas e servicos da empresa\n")
                else:
                    self.info.insertPlainText("\nSua latencia maxima está ok\n")
                    self.progressBar.setValue(100)
                    sleep(5)
                    self.progressBar.close()

    def serialNumber(self):
        comando = 'wmic bios get serialnumber'
        executar = os.popen(comando).read()
        format = executar[14:100]
        serial = re.sub(r"\s+", "", format)
        return serial

    def modeloEquipamento(self):
        comando = 'wmic computersystem get model,manufacturer'
        executar = os.popen(comando).read()
        format = executar[25:100]
        serial = re.sub(r"\s+", "", format)
        return serial

    def infoMaquina(self):
        try:
            comando = 'systeminfo'
            return os.popen(comando).read()
        except Exception as erro:
            self.info.setPlainText(erro)

    def eventoDiagnosticoHardare(self):
        self.info.setPlainText('')
        self.info.setPlainText('Teste está sendo realizado')
        diverseFunction.alert('Aguarde', 'Estamos fazendo o diagnostico do seu equipamento')
        self.progressBar.show()
        self.progressBar.setValue(1)
        chamaBTN = diverseFunction.eventBotoonHardware()
        x = 0
        progresso = 1
        self.info.setPlainText('')
        while x <= 9:
            self.info.insertPlainText(f'{chamaBTN[x]}\n')
            x +=1
            progresso+=10
            sleep(3)
            self.progressBar.setValue(progresso)

        self.progressBar.setValue(100)
        sleep(5)
        self.progressBar.close()

    def eventoInsertInfoMaquina(self):
        info = str(self.infoMaquina())
        serial = str(self.serialNumber())
        total_info = serial + info
        self.info.setPlainText(total_info)

    def eventoSendEmail(self):
        connect = diverseFunction.testeDeConectividade()
        if connect == False:
            diverseFunction.alert('Alerta!', 'Por favor verifique a sua conexão com a internet')
            self.info.setPlainText(
                'Verifique a sua conexão com a internet, va na aba Help e selecione "Problemas com internet", la tera algum passo a passo para você seguir')
        else:
            msg_agradecimento = str('Dados enviados para a nossa base, o time de TI agradece o contato =)')
            msg_base = str('Dados já enviado a base')
            if msg_agradecimento in self.info.toPlainText() or msg_base in self.info.toPlainText():
                self.info.setPlainText(msg_base)
            else:
                self.input_api()
                self.info.setPlainText(msg_agradecimento)

    def input_api(self):
        data = datetime.datetime.today().strftime('%d/%m/%Y - %H:%M')
        serial = str(self.serialNumber())
        modelo = str(self.modeloEquipamento())
        usuario = str(getpass.getuser())
        ApiRequests.inputAPI(serial=serial, usuario=usuario, modelo=modelo, data=data)


    def eventoNaoindentificar(self):
        self.label_6.close()
        self.label_7.close()
        self.label_8.close()
        self.input_nome.close()
        self.input_equipe.close()
        self.input_bandeira.close()

    def eventoIndentificar(self):
        self.label_6.show()
        self.label_7.show()
        self.label_8.show()
        self.input_nome.show()
        self.input_equipe.show()
        self.input_bandeira.show()

    def eventoAbrindoChamado(self):
        connect = diverseFunction.testeDeConectividade()
        if connect ==  False:
            diverseFunction.alert('Alerta!', 'Por favor verifique a sua conexão com a internet')
            self.info_chamado.setPlainText(
                'Verifique a sua conexão com a internet, va na aba Help e selecione "Problemas com internet", la tera algum passo a passo para você seguir')
        else:
            info = str(self.info_chamado.toPlainText())
            r_button = self.r_sim.isChecked()
            if r_button == True:
                nome = str(self.input_nome.text())
                equipe = str(self.input_equipe.text())
                bandeira = str(self.input_bandeira.currentText())
                if nome == '' or equipe == '' or bandeira == '' or info == '' or bandeira == 'NENHUM':
                    self.info_chamado.setPlainText(
                        'Digite todas as informações acima, caso não queira se indentificar por favor selecione a opção "Não"')
                else:
                    self.input_nome.setText('')
                    self.input_equipe.setText('')
                    self.info_chamado.setPlainText('Obrigado pelo seu feedback, ele é muito importante para o nosso desenvolvimento ;)')
                    diverseFunction.sendEmail(nome=nome, titulo='Abriu uma solicitação', mail= 'helpdesk@escale.com.br', texto=
                    f"""
                    <p>Nome: {nome}<p>
                    <p>Equipe: {equipe}<p>
                    <p>Bandeira: {bandeira}<p>
                    <p>{info}<p>""")
            else:
                if info == '':
                    self.info_chamado.setPlainText('Por favor digite no campo')
                else:
                    self.info_chamado.setPlainText(
                        'Obrigado pelo seu feedback, ele é muito importante para o nosso desenvolvimento ;)')
                    diverseFunction.sendEmail(
                        nome= 'Anonimo', titulo='Solicitação aberta de forma anonima', mail='helpdesk@escale.com.br', texto=info)


    def eventoAjuda(self):
        solicitacao = self.problemasRelacionais.currentText()
        if solicitacao == 'Nenhum':
            self.help.setPlainText('Por favor selecione algum problema que podemos ajudar!')
        elif solicitacao == 'Problemas para conectar na VPN?':
            self.help.setPlainText(f"""
Por gentileza realizar os procedimentos abaixo

1-Por gentileza, feche todas as janelas e reinicie o computador.

2-Depois que colocar para reiniciar o computador, tire o modem da tomada, aguarde 10 segundos, e conecte novamente.

3-Ligue seu computador e garanta que ele está conectado na internet.

Após feito os procedimentos acima e não de certo por favor entrar em contato com o nosso time de suporte, ou pode ficar avontade para abrir um ticket

WhatsApp Suporte: 800 663 1515
            """)
        elif solicitacao == 'Problemas com Phonemanager?':
            self.help.setPlainText(f"""
Por gentileza realizar os procedimentos abaixo

1-Abra o seu navegador

2-Aperte simultaneamente as teclas "CTRL + SHIFT + DELETE"

3-Após feito esse procedimento ira ser redirecionado para outra página

4-Seleciona a aba "Avançado"

5-Em periodo selecione "Todo o periodo"

6-selecione todos os checkbox

7-Clique em "Limpar dados"

Após feito os procedimentos acima e não de certo por favor entrar em contato com o nosso time de suporte, ou pode ficar avontade para abrir um ticket

WhatsApp Suporte: 800 663 1515
              """)
        elif solicitacao == 'Problema com Infrasip?':
            self.help.setPlainText(f"""
Por gentileza realizar os procedimentos abaixo

1-Por favor verificar se o seu ramal está conectado no infrasip é tambem no Phonemanager 
(Você consegue confirmar clicando no ícone de telefone no canto superior direito da tela, nele irá aparecer o ramal conectado.)

2-Após a confirmação, deslogar-se do ramal eo seu usuario do phonemanager e logar novamente com o ramal que está configurado no seu infrasip

3-Caso ainda o problema persistir por favor reiniciar a sua maquina e seu roteador

pós feito os procedimentos acima e não de certo por favor entrar em contato com o nosso time de suporte, ou pode ficar avontade para abrir um ticket

WhatsApp Suporte: 800 663 1515

            """)

        elif solicitacao == 'Problemas com Internet?':
            self.help.setPlainText("""
Por gentileza realizar os procedimentos abaixo

1-Desligue completamente o seu equipamento

2-Tire da tomada o seu roteador durante 10 segundo

3-Ligue o seu roteador de internet e o seu equipamento

4-Após feito os procedimento acima ir na aba "Diagnostico Desktop" e executar o "Teste de Rede"

5-Caso o teste acuse que a sua internet está instavel por favor entra em contato com o suporte tecnico da sua internet
            
            """)

    def eventoLimpeza(self):
        self.info.setPlainText('')
        connectInVpn = diverseFunction.vpnOrError()
        if connectInVpn == True:
            self.progressBar.show()
            self.progressBar.setValue(10)
            self.info.insertPlainText(f'{getpass.getuser()}, aguarde estamos fazendo a limpeza de todo o seu equipamento')
            diverseFunction.alert('Aguarde!', 'Processo de limpeza em andamento')
            sleep(5)
            self.progressBar.setValue(33)
            sleep(5)
            diverseFunction.cleanDesktop()
            self.progressBar.setValue(40)
            sleep(5)
            diverseFunction.cleanDNS()
            self.progressBar.setValue(55)
            sleep(5)
            atualizaPolitica = diverseFunction.updatePolice()
            self.progressBar.setValue(70)
            if atualizaPolitica == False:
                self.info.setPlainText(f'Falha ao atualizar politica, por favor entrar em contato com o nosso time de suporte.')
                self.progressBar.setValue(100)
                sleep(5)
                self.progressBar.close()
            else:
                self.info.setPlainText('Limpeza concluida com sucesso!')
                self.progressBar.setValue(100)
                sleep(5)
                self.progressBar.close()
        else:
           diverseFunction.alert('Aviso', 'Por favor conecte na VPN para prosseguir com a limpeza')

    def eventoAtualizandoPolice(self):
        connect = diverseFunction.testeDeConectividade()
        if connect ==  False:
            diverseFunction.alert('Alerta!', 'Por favor verifique a sua conexão com a internet')
            self.info.setPlainText(
                'Verifique a sua conexão com a internet, va na aba Help e selecione "Problemas com internet", la tera algum passo a passo para você seguir')
        else:
            self.info.setPlainText('')
            connectInVPN = diverseFunction.vpnOrError()
            if connectInVPN == True:
                self.progressBar.show()
                self.progressBar.setValue(25)
                self.info.setPlainText('Aguarde as atualizações podem demorar alguns instantes')
                sleep(5)
                self.progressBar.setValue(55)
                sleep(5)
                diverseFunction.cleanDNS()
                self.progressBar.setValue(75)
                sleep(5)
                atualizaPolice = diverseFunction.updatePolice()
                self.progressBar.setValue(80)
                if atualizaPolice == True:
                    self.info.setPlainText('Atualizações feita com sucesso!')
                    self.progressBar.setValue(100)
                    sleep(5)
                    self.progressBar.close()
                else:
                    self.info.setPlainText(
                        'As Atualizações não foram concluidas por favor entre em contato com o nosso time de suporte')
                    self.progressBar(100)
                    sleep(5)
                    self.progressBar.close()
            else:
                diverseFunction.alert('Aviso', 'Por favor conecte na VPN para prosseguir com o teste')


    def conf_infra(self):
        self.frame_15.show()

    def inicia_aut(self):
        try:
            self.usuario = getpass.getuser()
            self.operacao = self.cob_operacao.currentText()

            if self.operacao == 'NENHUM':
                diverseFunction.alert('Alerta', 'Por favor selecione uma operação')

            else:
                self.validando = self.dados.obtendo_ramal(operacao=self.operacao, usuario=self.usuario)
                if self.validando == False:
                    self.info.setPlainText(f'''Ops :-(

Infelizmente não encontramos o seu ramal na operação {self.operacao}.

Por favor confirme com o seu lider a sua operação ou entre em contato com o nosso time de suporte

WhatsApp Suporte: 0800 663 1515 ''')
                else:
                    self.conf.acessando_infra()
                    self.conf.logando_infra()
                    self.conf.preenchendo_conta(self.usuario, self.operacao)
                    diverseFunction.alert('Sucesso', 'Configuração feita com sucesso')
                    self.close()
            self.frame_15.close()

        except Exception as erro:
            self.frame_15.close()
            print(erro)
            diverseFunction.alert('Alerta', 'Infrasip já configurado')




    # def f_bot_user(self, respostaUser):
    #     data = datetime.datetime.today().strftime('%H:%M')
    #     usuario = str(getpass.getuser())
    #     qtd = len(self.listWidget)
    #     add = QtWidgets.QListWidgetItem()
    #     self.listWidget.addItem(add)
    #     item = self.listWidget.item(qtd)
    #     item.setText(f' \n{data} Você: {respostaUser}\n')
    #     item.setForeground(Qt.green)
    #     item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
    #
    # def f_bot(self, respostaBoot):
    #     data = datetime.datetime.today().strftime('%H:%M')
    #     qtd = len(self.listWidget)
    #     add = QtWidgets.QListWidgetItem()
    #     self.listWidget.addItem(add)
    #     item = self.listWidget.item(qtd)
    #     item.setText(f'{data} Zoe: {respostaBoot}')
    #     item.setTextAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)



if __name__ == '__main__':
    qt = QApplication(sys.argv)
    suporte = EscaleSuport()
    suporte.show()
    qt.exec_()




#pyinstaller --onefile -w supportAssist.py
#pyinstaller --windowed supportAssist.py
#pyinstaller --windowed -i logoEsc.ico supportAssist.py
