import getpass
import os
import ApiRequests
import datetime
import re

data = datetime.datetime.today().strftime('%D-%M-%Y %H:%M')

def serialNumber():
    comando = 'wmic bios get serialnumber'
    executar = os.popen(comando).read()
    format = executar[14:100]
    serial = re.sub(r"\s+", "", format)
    return serial

def modeloEquipamento():
    comando = 'wmic computersystem get model,manufacturer'
    executar = os.popen(comando).read()
    format = executar[25:100]
    serial = re.sub(r"\s+", "", format)
    return serial

def usuarios():
    set_user= str(getpass.getuser())
    return set_user



print(modeloEquipamento())


ApiRequests.inputAPI(serial=serialNumber(), usuario= usuarios(), modelo=modeloEquipamento(), data= data)

