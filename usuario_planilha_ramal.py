from dados_planilha import ObterTabela

class Dados:
    def __init__(self):
        self.lista_dados = []

    def obtendo_ramal(self, operacao, usuario):
        self.dados = ObterTabela().obter_tabela(operacao)
        self.lista_dados.append(self.dados)
        contador = len(self.lista_dados[0])
        cont =0
        while cont <= contador:
            cont+=1
            try:
                self.usuario = str(self.lista_dados[0][cont][3]).upper()
                print(usuario, self.usuario, contador, cont)
                if self.usuario == str(usuario).upper():
                    self.ramal = self.lista_dados[0][cont][0]
                    print(self.ramal)
                    self.lista_dados = []
                    return self.ramal
            except IndexError as erro:
                print(erro)
                pass
        self.lista_dados= []
        return False






