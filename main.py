import quickstart
import ScrapMTE

if __name__ == '__main__':
    try:
        print("-----Bem vindo-----")
        while(True):
            print("#####################################################")
            print("Escolha uma das opções abaixo:")
            print("1- Atualizar todas as CCTs")
            print("2- Escolher linhas para atualizar")
            opcao1 = input("Digite aqui: ")
            if(opcao1 == "1"):
                quickstart.Atualiza_CCTS()

            else:
                a = input("Digite a linha que você quer começar: ")
                b = input("Digite a linha que você quer terminar: ")
                quickstart.Atualiza_CCTS_V2(int(a)-3,int(b)-2)
    except Exception as erro:
        print(erro)
        opcao2 = input("Enter para finalizar")
