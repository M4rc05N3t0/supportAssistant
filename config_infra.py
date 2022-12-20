from pywinauto.application import Application
import pyautogui
from usuario_planilha_ramal import Dados
from time import sleep

class InfraSIP:
    def __init__(self):
        self.titulo_janela = "InfraSIP"
        self.planilha = Dados()
        self.tabela_pabx = {'CLARO':'pabxclaro.cobmax.com.br',
                            'VIVO': 'pabxvivo.cobmax.com.br',
                            'ITAU':'finance.cobmax.com.br',
                            'HEALTH': 'health.cobmax.com.br',
                            'PDP': 'pabxpdp.cobmax.com.br',
                            'MULTIMARCAS': 'pabxmultimarcas.cobmax.com.br'
                            }
    def acessando_infra(self):
        try:
            self.app = Application(backend='uia').connect(title=self.titulo_janela, timeout=5)
        except:
            pass

    def logando_infra(self):
        self.app.InfraSIP.child_window(title="Menu", auto_id="1001", control_type="Button").click_input(double=True)
        pyautogui.typewrite('eSc4l3@17#')
        self.app.InfraSIP.child_window(title="Ok", auto_id="1", control_type="Button").click_input()

    def preenchendo_conta(self, usuario, operacao):
        self.ramal =self.planilha.obtendo_ramal(usuario=usuario, operacao=operacao)
        if self.ramal == False:
            return False
        else:
            self.app.InfraSIP.child_window(title="Nome da Conta", auto_id="1152", control_type="Edit").click_input()
            pyautogui.typewrite(self.ramal)
            self.app.InfraSIP.child_window(title="Servidor SIP", auto_id="1052", control_type="Edit").click_input()
            pyautogui.typewrite(self.tabela_pabx[f'{operacao}'])
            self.app.InfraSIP.child_window(title="Usuário", auto_id="1053", control_type="Edit").click_input()
            pyautogui.typewrite(self.ramal)
            self.app.InfraSIP.child_window(title="Domínio", auto_id="1046", control_type="Edit").click_input()
            pyautogui.typewrite(self.tabela_pabx[f'{operacao}'])
            self.app.InfraSIP.child_window(title="Login", auto_id="1043", control_type="Edit").click_input()
            pyautogui.typewrite(self.ramal)
            self.app.InfraSIP.child_window(title="Senha", auto_id="1049", control_type="Edit").click_input()
            pyautogui.typewrite('c0bm4xt3l3f0n1@')
            self.app.InfraSIP.child_window(title="Nome de exibição", auto_id="1045", control_type="Edit").click_input()
            pyautogui.typewrite(self.ramal)
            self.app.InfraSIP.child_window(title="Salvar", auto_id="1", control_type="Button").click_input()

# dado = InfraSIP()
# dado.acessando_infra()
# dado.logando_infra()
# dado.preenchendo_conta('marcos.neto', 'CLARO')