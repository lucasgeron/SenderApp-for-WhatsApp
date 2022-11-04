import pyautogui
# pyautogui.PAUSE = 0.5
import requests
import time
import json
import sys
import os
import traceback
import webbrowser
import subprocess
from requests.utils import quote

sends = 0
qs = 0 
ccs = 0

def updateToken():
    save_path = 'C:/Program Files/Common Files'
    full_path = os.path.join(save_path, "SenderApp" + ".conf")

    file_exists = os.path.isfile(full_path)

    if file_exists:
      f = open(full_path, "r")
      lines = f.readlines()
      email  = lines[0].split("\n")[0]
      token = lines[1].split("\n")[0]
      tcopy = token
      f.close()
      
      f = open(full_path, "w+")
      token = ""
      while (token == ""):
        token = pyautogui.prompt(text='Email:Informe o Novo Token', title='SenderApp - Alterar Token' , default='XXXX-XXXX-XXXX')

      if (not(token)):
        token=tcopy
        print("Operação não Realizada - Cancelado pelo usuário.")
        # pyautogui.alert(text='Cancelado pelo usuário.\nRestaurando último token cadastrado...', title='SenderApp - Operação Não Realizada', button='OK')
      else:
        print("Token atualizado com sucesso!")
        time.sleep(1)
        os.system('cls')
        printLogo()
        # pyautogui.alert(text='Token atualizado com sucesso!', title='SenderApp - Operação Realizada', button='OK')  


      f.write(email +"\n")
      f.write(token+"\n")
      f.close()
     
      startApp()

def setup():
    save_path = 'C:/Program Files/Common Files'
    full_path = os.path.join(save_path, "SenderApp" + ".conf")

    file_exists = os.path.isfile(full_path)

    if file_exists:
      f = open(full_path, "r")
      lines = f.readlines()
      email  = lines[0].split("\n")[0]
      token = lines[1].split("\n")[0]
      # print("Lendo Arquivo de Configuração...")
    else:
      f = open(full_path, "w+")

      email = ""
      token = ""
      
      while(email == ""):
        email = pyautogui.prompt(text='Informe seu email', title='SenderApp - Nova Configuração' , default='')
        
      if (not(email)):
        print("Operação não Realizada - Cancelado pelo usuário.")
        pyautogui.alert(text='Cancelado pelo usuário.', title='SenderApp - Operação Não Realizada', button='OK')
        sys.exit(0)

      print("Email: " + email)
      f.write(email+"\n")

      while(token == ""):
        token = pyautogui.prompt(text='Informe seu Token', title='SenderApp - '+ email , default='XXXX-XXXX-XXXX')

      if (not(token)):
        print("Operação não Realizada - Cancelado pelo usuário.")
        pyautogui.alert(text='Cancelado pelo usuário.', title='SenderApp - Operação Não Realizada', button='OK')
        sys.exit(0)

      print("Token: " + token)
      f.write(token+"\n")
      print("Configuração Realizada com sucesso!")

    f.close()
    
    time.sleep(1)
    os.system('cls')
    printLogo()

    return  [email, token]

def startApp():

    configs = setup();
    
    print("- Conectando ao servidor...")

    url = "https://script.google.com/macros/s/AKfycbyyVrXZ2nmgwuPBcrrL2OWQWVbLKf_PkVWNIXT_kZ4UAgkhk0HrGxm7MgvxVtMx9PePjg/exec?"

    payload = {
      'a': 'p',
      'email' : configs[0],
      'token': configs[1]
    }
    
    response = requests.request("POST", url, data=payload)
    print("Conexão estabelecida.")
    time.sleep(1)
    os.system('cls')

    client =  json.loads(response.text)

    printLogo()

    if isinstance(client, str): # access denied
      print(client)
      opt = pyautogui.confirm(text=client + "\n\n Token Atual:\n"+payload['token']+"\n\nDeseja alterar seu Token Atual?", title='SenderApp - Acesso Negado', buttons=['Sim', 'Não', 'Renovar Assinatura'])
      if (opt =="Sim"):
        updateToken()
        
      elif(opt == "Renovar Assinatura"):
         webbrowser.open(url + "a=s")
    else: # connectToApp :: 
      
      cardInfo = "" 
      print("Planilha: https://docs.google.com/spreadsheets/d/" + client['sheetId'])

      if(client['serverNote'] != ""):
        print("Notificação do Servidor:\n" + client['serverNote'])
        cardInfo += "Notificação do Servidor:\n" + client['serverNote']+"\n\n"

      sConfigs = {
        'showSignatureInfos' : client['sConfigs'][3][1],
        'openSheet' : client['sConfigs'][4][1],
        'openWhatsApp' : client['sConfigs'][5][1],
      }
      
      if sConfigs['openSheet']:
        print("Abrindo Planilha...")
        webbrowser.open("https://docs.google.com/spreadsheets/d/" + client['sheetId'] + "/edit?usp=drivesdk"),

      if sConfigs['showSignatureInfos']:
        print("Email: "+ client['email'])
        cardInfo += ("Email: "+ client['email']+"\n")
        print("Token: " + client['token']+" - " + client['validityInfo'])
        cardInfo += ("Token: " + client['token'] +" - " + client['validityInfo']) + "\n\n"


      # print quotes avaliable
      cardInfo += client['infoMsg'] + "\n"
      print(client['infoMsg'] )
 
      if client['onHold']:
        pyautogui.alert(text=client['infoMsg'], title='SenderApp - ' + client['email'], button='OK')
        sys.exit(0)
      elif (len(client['contacts']) < 10 ):
        pyautogui.alert(text="Atenção:\nPara um melhor desempenho, é necessário possuir ao menos 10 contatos válidos.\nAtualmente você possui apenas " +  str(client['ccs']) + " contatos váldios em sua planilha.\n\n- Considere preencher as primeiras linhas.", title='SenderApp - ' + client['email'], button='OK')
        if not(sConfigs['openSheet']):
          print("Abrindo Planilha..." )
          webbrowser.open("https://docs.google.com/spreadsheets/d/" + client['sheetId'] + "/edit?usp=drivesdk")
        sys.exit(0)
      else:  
        print("Lista de Envios: " + str(client['ccs']))
        cardInfo += ("Lista de Envios: " + str(client['ccs'])) + "\n\n"

      cardInfo += "- - - - - - - - - - - - - - - - - - - - \n"
      cardInfo += ("Mensagem: " + client['message']) + "\n"
      cardInfo += "- - - - - - - - - - - - - - - - - - - - \n"

      print("- - - - - - - - - - - - - - - - - - - - ")
      print("Mensagem: " + client['message'])
      print("- - - - - - - - - - - - - - - - - - - - ")

      if sConfigs['openWhatsApp']:
        print("Abrindo WhatsApp...")
        subprocess.Popen(["cmd", "/C", "start whatsapp://send?"], shell=True)

      opt = pyautogui.confirm(text=cardInfo, title='SenderApp - ' + client['email'], buttons=['Iniciar Envio',  'Abrir Planilha', 'Sair'])

      if (opt ==  "Abrir Planilha"):
        print("Abrindo Planilha...")
        webbrowser.open("https://docs.google.com/spreadsheets/d/" + client['sheetId'] + "/edit?usp=drivesdk"),
        opt = pyautogui.confirm(text=cardInfo, title='SenderApp - ' + client['email'], buttons=['Iniciar Envio',  'Sair'])
      
      if(opt == "Sair"):
        sys.exit(0)
      else:     
        startBot = pyautogui.confirm(text='Deseja iniciar o envio das Mensagens?\n\n- - - - - - - - - - - - - - - - - - - - \n' + client['message'] + "\n- - - - - - - - - - - - - - - - - - - - \n", title='SenderApp - ' + client['email'], buttons=['Sim', 'Não'])


        if startBot == "Sim":
          print("")
          print("Iniciando SenderApp...")
          print("")

          i=0
          global sends
          global qs 
          qs = client['qs']
          global ccs
          ccs = client['ccs']

          qsCounter = qs

          for i in range (len(client['contacts'])):
            contact = {
              'name' :  client['contacts'][i][0],
              'info' :  client['contacts'][i][1],
              'phone' :  str(client['contacts'][i][2]),
            }

            curMsg = client['message']
            curMsg = curMsg.replace("@Nome", contact['name'])
            curMsg = curMsg.replace("@Info", contact['info'])

            if(send(contact['name'], contact['info'], contact['phone'], client['message']) == 1):
              qsCounter = qsCounter - 1
              sends = sends + 1

            os.system('cls')
            printLogo()

            print("Nome: "+ contact['name'] + "\nInfo: "+contact['info'] + "\nTelefone: "+ contact['phone']+"\n\nMensagem:\n"+ curMsg+"\n\nTotal de Cotas Utilizadas: "+str(client['sends']+sends)+"\nCotas Utilizadas Nesta Execução: " + str(sends) +"\nCotas Disponíveis Restantes: " + str(qsCounter) +"\n\n")
        
          payload = {
            'a': 'e',
            'email' : configs[0],
            'token': configs[1],
            'qs' : client['qs'],
            'ccs' : client['ccs'],
            'sendSuccess' : sends
          }

          response = requests.request("POST", url, data=payload)
          pyautogui.alert(text= ' Mensagens Enviadas com Sucesso! ' + str(sends) + " Cotas Utilizadas", title='SenderApp - ' + client['email'], button='OK') 

def send(nome, info, tel, msg):

  msg = msg.replace("@Nome", nome)
  msg = msg.replace("@Info", info)

  link = "start whatsapp://send?phone=55" + tel +"^&text="+quote(msg, safe='~@#$&()*!+=:;,.?/\'')
  
  if(len(link) <= 2000 ):
    subprocess.Popen(["cmd", "/C", link], shell=True)
    time.sleep(1.5) 
    pyautogui.hotkey('enter')  
    return 1
  else:
    return 0

def printLogo():
  print("")
  print("███████ ███████ ███    ██ ██████  ███████ ██████   █████  ██████  ██████  ")
  print("██      ██      ████   ██ ██   ██ ██      ██   ██ ██   ██ ██   ██ ██   ██ ")
  print("███████ █████   ██ ██  ██ ██   ██ █████   ██████  ███████ ██████  ██████  ")
  print("     ██ ██      ██  ██ ██ ██   ██ ██      ██   ██ ██   ██ ██      ██      ")
  print("███████ ███████ ██   ████ ██████  ███████ ██   ██ ██   ██ ██      ██      ")
  print("")
  print("Para cancelar a execução do aplicativo, selecione esta janela e aperte Ctrl + C.")
  print("")
                                                                          
if __name__ == "__main__":
  try:
    printLogo()
    startApp()
  except PermissionError:
    pyautogui.alert(text='Operação Não Realiazada!\n  |' +  str(sys.exc_info()[0]) +  '\n  | O usuário atual não possui permissões suficiente.\n\nExecute a aplicação como Administrador.', title='SenderApp - Erro de Permissão', button='OK')
  except SystemExit as e:
    if e.code != 0:
      pyautogui.alert(text='Encerrando aplicação...\nErro Código: ' + str(e.code), title='SenderApp', button='OK') 
  except IndexError:
    print("Existem erros no arquivo de configuração do SenderApp.")
    print("")
    print('Acesse a pasta C:\Program Files\Common Files e exclua o arquivo "SenderApp.conf"')
    print('Reinicie o aplicativo e faça uma nova configuração')
    pyautogui.alert(text='Existem erros no arquivo de configuração do SenderApp. ', title='SenderApp - Operação Não Realizada', button='OK') 
  except KeyboardInterrupt:
    configs = setup()
    url = "https://script.google.com/macros/s/AKfycbyyVrXZ2nmgwuPBcrrL2OWQWVbLKf_PkVWNIXT_kZ4UAgkhk0HrGxm7MgvxVtMx9PePjg/exec?"
    payload = {
      'a': 'e',
      'email' : configs[0],
      'token': configs[1],
      'qs' : qs,
      'ccs' : ccs,
      'sendSuccess' : sends
    }
    print("SenderApp - Parando execução...")
    response = requests.request("POST", url, data=payload)
    pyautogui.alert(text='Por medida de segurança, vamos considerar o envio de ' + str(sends) + ' mensagens.\nRecomendamos que só envie mensagens quando as tiver certeza da informação.', title='SenderApp', button='OK')
  except:
    pyautogui.alert(text='Operação Não Realiazada!\n  |' +  str(sys.exc_info()[0]) +  '\n'+traceback.format_exc(), title='SenderApp - Erro', button='OK')
  finally:
    print("Saindo do SenderApp...")
    pyautogui.alert(text='SenderApp Encerrado.', title='SenderApp', button='OK') 

