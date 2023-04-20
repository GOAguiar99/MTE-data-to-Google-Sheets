from __future__ import print_function
import ScrapMTE
import os.path

from pprint import pprint
from googleapiclient import discovery
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
spreadsheetId='123445645523423'
def Atualiza_CCTS():
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    # Call the Sheets API
    sheet = service.spreadsheets()
    cnpj = sheet.values().get(spreadsheetId  ,
                                range='CCTs!F3:F100').execute()
    cnpjj = cnpj.get('values', [])

    result = sheet.values().get(spreadsheetId  ,
                                range='CCTs!I3:I100').execute()
    Fim_ant = result.get('values', [])

    result = sheet.values().get(spreadsheetId  ,
                                range='CCTs!J3:J100').execute()
    Tipo_ant = result.get('values', [])

    result = sheet.values().get(spreadsheetId  ,
                                range='CCTs!E3:E100').execute()
    Sind_Empresas = result.get('values', [])



    Modificado = [
        ["Sim"],
    ]
    NModificado = [
        ["Não"],
    ]
    Convencao = [
        ["Convenção"],
    ]
    Termo = [
        ["Termo"],
    ]

    for i in range(len(cnpjj)): #Numero de linhas da CNPJ
        if(i==i):
            now = datetime.now()
            Now = str(now.strftime("%d/%m/%Y %X"))
            Agora = [
                [Now],
            ]
            result = sheet.values().update(spreadsheetId  ,
                                        range='CCTs!P2',valueInputOption="USER_ENTERED", body={"values": Agora}).execute() #Atualizador de data de atualização
            print("Atualizando a linha: ",i+3)
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################

        try: #Primeiro requeste TRY
            A = ScrapMTE.scrap(cnpjj[i])
            for x in range(0,10):
                C_list = A[9+2+(13*x)].split('\n')
                if(Sind_Empresas[i]==['{}'.format(C_list[2])]):
                    B_list = A[9+(13*x)].split('-')
                    break
            #print("Na Planilha ", Sind_Empresas[i], "No MTE", C_list[2], sep=": ")
            print("scrap1", cnpjj[i], B_list[0], B_list[1], sep=": ")
            Inicio = str(B_list[0])
            Fim = str(B_list[1])
            inserir = [
                [Inicio,Fim],
            ]
            Fim_D = datetime.strptime(Fim,' %d/%m/%Y')
            try:
                Fim_antD = datetime.strptime(str(Fim_ant[i]),"['%d/%m/%Y']")
            except:
                Fim_antD = None #Tenta transformar string em DATA
            if (Fim_antD is None):# Se Celula vazia
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Convencao}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()

                print("Vigência era NULA, atualizou.")
                print("--------------------------------------------------")
            elif(Fim_D > Fim_antD): #Celula com valor menor do que o atual
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Convencao}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()

                print("Vigência Atualizada")
                print("--------------------------------------------------")
            elif((Fim_D == Fim_antD) and (Tipo_ant[i] is None or Tipo_ant[i] == [])):
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Convencao}).execute()
                print("Tipo alterado.")
                print("--------------------------------------------------")

            else: #Celula com o valor maior ou igual do que o atual
                # result = sheet.values().update(spreadsheetId  ,
                #                             range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": NModificado}).execute()
                print("Não atualizou")
                print("--------------------------------------------------")
        except Exception as xx:
            print(xx)
            print("Erro scrap1, tentando novamente...")
            Fim_antD = None
            try: #Tenta primeiro request denovo
                B = ScrapMTE.scrap(cnpjj[i])
                for x in range(0,5):
                    C_list = A[9+2+(13*x)].split('\n')
                    if(Sind_Empresas[i]==['{}'.format(C_list[2])]):
                        B_list = A[9+(13*x)].split('-')
                        break
                #print("Na Planilha ", Sind_Empresas[i], "No MTE", C_list[2], sep=": ")
                print("scrap1", cnpjj[i], B_list[0], B_list[1], sep=": ")
                Inicio = str(B_list[0])
                Fim = str(B_list[1])
                inserir = [
                    [Inicio,Fim],
                ]
                Fim_D = datetime.strptime(Fim,' %d/%m/%Y')
                try:
                    print(Fim_ant[i])
                    Fim_antD = datetime.strptime(str(Fim_ant[i]),"['%d/%m/%Y']")
                except:
                    Fim_antD = None
                if (Fim_antD is None):
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Convencao}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                    print("Vigência era NULA, atualizou.")
                    print("--------------------------------------------------")

                elif(Fim_D > Fim_antD):
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Convencao}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                    print("Vigência Atualizada")
                    print("--------------------------------------------------")
                elif((Fim_D == Fim_antD) and (Tipo_ant[i] is None or Tipo_ant[i] == [])):
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Convencao}).execute()
                    print("Tipo alterado.")
                    print("--------------------------------------------------")

                else:
                    # result = sheet.values().update(spreadsheetId  ,
                    #                             range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": NModificado}).execute()
                    print("Não Atualizou.")
                    print("--------------------------------------------------")
            except Exception as xx2:
                print(xx2)
                print("Convenção não foi atualizada, indo para proximo request.")
                print("--------------------------------------------------")
                Fim_antD = None
            try:
                B_list.clear()
            except:
                print("Nao limpou o vetor B")
            result = sheet.values().get(spreadsheetId  ,
                                        range='CCTs!I3:I100').execute()
            Fim_ant = result.get('values', []) #Atualiza vetor com os valores de

#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
        try: #Faz segundo request
            A = ScrapMTE.scrap2(cnpjj[i])
            for x in range(0,5):
                C_list = A[9+2+(13*x)].split('\n')
                if(Sind_Empresas[i]==['{}'.format(C_list[2])]):
                    B_list = A[9+(13*x)].split('-')
                    break
            #print("Na Planilha ", Sind_Empresas[i], "No MTE", C_list[2], sep=": ")
            print("scrap2", cnpjj[i], B_list[0], B_list[1], sep=": ")
            Inicio = str(B_list[0])
            Fim = str(B_list[1])
            inserir = [
                [Inicio,Fim],
            ]
            Fim_D = datetime.strptime(Fim,' %d/%m/%Y')
            try:
                print(Fim_ant[i])
                Fim_antD = datetime.strptime(str(Fim_ant[i]),"['%d/%m/%Y']")
            except:
                Fim_antD = None
            if (Fim_antD is None):
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Termo}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                print("Vigência era NULA, atualizou.")
                print("--------------------------------------------------")
            elif(Fim_D > Fim_antD):
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Termo}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                print("Vigência Atualizada.")
                print("--------------------------------------------------")
            elif((Fim_D == Fim_antD) and (not Tipo_ant[i] == ['Termo'])):
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Termo}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                print("Tipo alterado.")
                print("--------------------------------------------------")
            else:
                # result = sheet.values().update(spreadsheetId  ,
                #                             range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": NModificado}).execute()
                print("Não Atualizou.")
                print("--------------------------------------------------")

        except Exception as xx:
            print(xx)
            print("Erro scrap2, tentando novamente...")
            Fim_antD = None
            try:
                A = ScrapMTE.scrap2(cnpjj[i])
                for x in range(0,5):
                    C_list = A[9+2+(13*x)].split('\n')
                    if(Sind_Empresas[i]==['{}'.format(C_list[2])]):
                        B_list = A[9+(13*x)].split('-')
                        break
                #print("Na Planilha ", Sind_Empresas[i], "No MTE", C_list[2], sep=": ")
                print("scrap2", cnpjj[i], B_list[0], B_list[1], sep=": ")
                Inicio = str(B_list[0])
                Fim = str(B_list[1])
                inserir = [
                    [Inicio,Fim],
                ]
                Fim_D = datetime.strptime(Fim,' %d/%m/%Y')
                try:
                    print(Fim_ant[i])
                    Fim_antD = datetime.strptime(str(Fim_ant[i]),"['%d/%m/%Y']")
                except:
                    Fim_antD = None
                if (Fim_antD is None):
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Termo}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                    print("Vigência era NULA, atualizou.")
                    print("--------------------------------------------------")
                elif(Fim_D > Fim_antD):
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Termo}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                    print("Vigência Atualizada.")
                    print("--------------------------------------------------")
                elif((Fim_D == Fim_antD) and (not Tipo_ant[i] == ['Termo'])):
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Termo}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                    print("Tipo alterado.")
                    print("--------------------------------------------------")
                else:
                    # result = sheet.values().update(spreadsheetId  ,
                    #                             range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": NModificado}).execute()
                    print("Não Atualizou.")
                    print("--------------------------------------------------")
            except Exception as xx:
                Fim_antD = None
                print(xx)
                print("Termo não atualizado, indo para próximo CNPJ.")
                print("--------------------------------------------------")
def Atualiza_CCTS_V2(a, b):
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    # Call the Sheets API
    sheet = service.spreadsheets()

    cnpj = sheet.values().get(spreadsheetId  ,
                                range='CCTs!F3:F100').execute()
    cnpjj = cnpj.get('values', [])

    result = sheet.values().get(spreadsheetId  ,
                                range='CCTs!I3:I100').execute()
    Fim_ant = result.get('values', [])

    result = sheet.values().get(spreadsheetId  ,
                                range='CCTs!J3:J100').execute()
    Tipo_ant = result.get('values', [])

    result = sheet.values().get(spreadsheetId  ,
                                range='CCTs!E3:E100').execute()
    Sind_Empresas = result.get('values', [])



    Modificado = [
        ["Sim"],
    ]
    NModificado = [
        ["Não"],
    ]
    Convencao = [
        ["Convenção"],
    ]
    Termo = [
        ["Termo"],
    ]

    for i in range(a, b): #Numero de linhas da CNPJ
        if(i==i):
            now = datetime.now()
            Now = str(now.strftime("%d/%m/%Y %X"))
            Agora = [
                [Now],
            ]
            result = sheet.values().update(spreadsheetId  ,
                                        range='CCTs!P2',valueInputOption="USER_ENTERED", body={"values": Agora}).execute() #Atualizador de data de atualização
            print("Atualizando linha: ",i+3)
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################

        try: #Primeiro requeste TRY
            A = ScrapMTE.scrap(cnpjj[i])
            for x in range(0,10):
                C_list = A[9+2+(13*x)].split('\n')
                if(Sind_Empresas[i]==['{}'.format(C_list[2])]):
                    B_list = A[9+(13*x)].split('-')
                    break
            #print("Na Planilha ", Sind_Empresas[i], "No MTE", C_list[2], sep=": ")
            print("scrap1", cnpjj[i], B_list[0], B_list[1], sep=": ")
            Inicio = str(B_list[0])
            Fim = str(B_list[1])
            inserir = [
                [Inicio,Fim],
            ]
            Fim_D = datetime.strptime(Fim,' %d/%m/%Y')
            try:
                Fim_antD = datetime.strptime(str(Fim_ant[i]),"['%d/%m/%Y']")
            except:
                Fim_antD = None #Tenta transformar string em DATA
            if (Fim_antD is None):# Se Celula vazia
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Convencao}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()

                print("Vigência era NULA, atualizou.")
                print("--------------------------------------------------")
            elif(Fim_D > Fim_antD): #Celula com valor menor do que o atual
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Convencao}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()

                print("Vigência Atualizada")
                print("--------------------------------------------------")
            elif((Fim_D == Fim_antD) and (Tipo_ant[i] is None or Tipo_ant[i] == [])):
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Convencao}).execute()
                print("Tipo alterado.")
                print("--------------------------------------------------")

            else: #Celula com o valor maior ou igual do que o atual
                # result = sheet.values().update(spreadsheetId  ,
                #                             range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": NModificado}).execute()
                print("Não atualizou")
                print("--------------------------------------------------")
        except Exception as xx:
            print(xx)
            print("Erro scrap1, tentando novamente...")
            Fim_antD = None
            try: #Tenta primeiro request denovo
                B = ScrapMTE.scrap(cnpjj[i])
                for x in range(0,5):
                    C_list = A[9+2+(13*x)].split('\n')
                    if(Sind_Empresas[i]==['{}'.format(C_list[2])]):
                        B_list = A[9+(13*x)].split('-')
                        break
                #print("Na Planilha ", Sind_Empresas[i], "No MTE", C_list[2], sep=": ")
                print("scrap1", cnpjj[i], B_list[0], B_list[1], sep=": ")
                Inicio = str(B_list[0])
                Fim = str(B_list[1])
                inserir = [
                    [Inicio,Fim],
                ]
                Fim_D = datetime.strptime(Fim,' %d/%m/%Y')
                try:
                    print(Fim_ant[i])
                    Fim_antD = datetime.strptime(str(Fim_ant[i]),"['%d/%m/%Y']")
                except:
                    Fim_antD = None
                if (Fim_antD is None):
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Convencao}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                    print("Vigência era NULA, atualizou.")
                    print("--------------------------------------------------")

                elif(Fim_D > Fim_antD):
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Convencao}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                    print("Vigência Atualizada")
                    print("--------------------------------------------------")
                elif((Fim_D == Fim_antD) and (Tipo_ant[i] is None or Tipo_ant[i] == [])):
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Convencao}).execute()
                    print("Tipo alterado.")
                    print("--------------------------------------------------")

                else:
                    # result = sheet.values().update(spreadsheetId  ,
                    #                             range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": NModificado}).execute()
                    print("Não Atualizou.")
                    print("--------------------------------------------------")
            except Exception as xx2:
                print(xx2)
                print("Convenção não foi atualizada, indo para proximo request.")
                print("--------------------------------------------------")
                Fim_antD = None
            try:
                B_list.clear()
            except:
                print("Nao limpou o vetor B")
            result = sheet.values().get(spreadsheetId  ,
                                        range='CCTs!I3:I100').execute()
            Fim_ant = result.get('values', []) #Atualiza vetor com os valores de

#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
#######################################################################################################################
        try: #Faz segundo request
            A = ScrapMTE.scrap2(cnpjj[i])
            for x in range(0,5):
                C_list = A[9+2+(13*x)].split('\n')
                if(Sind_Empresas[i]==['{}'.format(C_list[2])]):
                    B_list = A[9+(13*x)].split('-')
                    break
            #print("Na Planilha ", Sind_Empresas[i], "No MTE", C_list[2], sep=": ")
            print("scrap2", cnpjj[i], B_list[0], B_list[1], sep=": ")
            Inicio = str(B_list[0])
            Fim = str(B_list[1])
            inserir = [
                [Inicio,Fim],
            ]
            Fim_D = datetime.strptime(Fim,' %d/%m/%Y')
            try:
                print(Fim_ant[i])
                Fim_antD = datetime.strptime(str(Fim_ant[i]),"['%d/%m/%Y']")
            except:
                Fim_antD = None
            if (Fim_antD is None):
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Termo}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                print("Vigência era NULA, atualizou.")
                print("--------------------------------------------------")
            elif(Fim_D > Fim_antD):
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Termo}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                print("Vigência Atualizada.")
                print("--------------------------------------------------")
            elif((Fim_D == Fim_antD) and (not Tipo_ant[i] == ['Termo'])):
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Termo}).execute()
                result = sheet.values().update(spreadsheetId  ,
                                            range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                print("Tipo alterado.")
                print("--------------------------------------------------")
            else:
                # result = sheet.values().update(spreadsheetId  ,
                #                             range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": NModificado}).execute()
                print("Não Atualizou.")
                print("--------------------------------------------------")

        except Exception as xx:
            print(xx)
            print("Erro scrap2, tentando novamente...")
            Fim_antD = None
            try:
                A = ScrapMTE.scrap2(cnpjj[i])
                for x in range(0,5):
                    C_list = A[9+2+(13*x)].split('\n')
                    if(Sind_Empresas[i]==['{}'.format(C_list[2])]):
                        B_list = A[9+(13*x)].split('-')
                        break
                #print("Na Planilha ", Sind_Empresas[i], "No MTE", C_list[2], sep=": ")
                print("scrap2", cnpjj[i], B_list[0], B_list[1], sep=": ")
                Inicio = str(B_list[0])
                Fim = str(B_list[1])
                inserir = [
                    [Inicio,Fim],
                ]
                Fim_D = datetime.strptime(Fim,' %d/%m/%Y')
                try:
                    print(Fim_ant[i])
                    Fim_antD = datetime.strptime(str(Fim_ant[i]),"['%d/%m/%Y']")
                except:
                    Fim_antD = None
                if (Fim_antD is None):
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Termo}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                    print("Vigência era NULA, atualizou.")
                    print("--------------------------------------------------")
                elif(Fim_D > Fim_antD):
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!H{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": inserir}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Termo}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                    print("Vigência Atualizada.")
                    print("--------------------------------------------------")
                elif((Fim_D == Fim_antD) and (not Tipo_ant[i] == ['Termo'])):
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!J{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Termo}).execute()
                    result = sheet.values().update(spreadsheetId  ,
                                                range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": Modificado}).execute()
                    print("Tipo alterado.")
                    print("--------------------------------------------------")
                else:
                    # result = sheet.values().update(spreadsheetId  ,
                    #                             range='CCTs!K{}'.format(i+3),valueInputOption="USER_ENTERED", body={"values": NModificado}).execute()
                    print("Não Atualizou.")
                    print("--------------------------------------------------")
            except Exception as xx:
                Fim_antD = None
                print(xx)
                print("Termo não atualizado, indo para próximo CNPJ.")
                print("--------------------------------------------------")
