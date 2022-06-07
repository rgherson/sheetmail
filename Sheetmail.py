# outras bibliotecas
from os import listdir, getcwd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pwinput import pwinput

# meus arquivos
import constantes as const
import funcoes as fun

#usamos a variavel erro para verificar execoes e rodar as listas de opções varias vezes
erro = True
    
# abrir o texto e copiar pra uma lista
while erro == True:
    try:
        erro = False
        txt_list = fun.abrir_txt()
        print("abrimos o seu texto, por enquanto tudo certo")

    except TypeError as err:
        print(err)
        erro = True
    except Exception as err:
        print(err)
        quit()
    except:
        print("ocorreu um erro inesperado ao tentar acessar seu arquivo de texto, chame alguem inteligente para dar uma checada em tudo")
        quit()

erro = True

# abrir excel
while erro == True:
    try:
        erro = False
        sheet = fun.abrir_xsl()
        print("abrimos seu excel, por enquanto tudo certo")

    except TypeError as err:
        print(err)
        print("vamos tentar de novo")
        erro = True
    except OSError as err:
        print("Um erro ocorreu!!!")
        print(err)
        quit()
    except fun.ErroCarregamentoExcel as err:
        print(err)
        erro = True
    except BaseException as err:
        print("ocorreu um erro inesperado ao tentar acessar seu arquivo excel, chame alguem inteligente para dar uma checada em tudo ( ˘︹˘)")
        print("para geeks: ", err)
        quit()

# construir um tuple de contatos, max_row da o numero de linhas do exel mais 1
contatos_list = [] 
for linha in sheet.rows:
    contatos_list.append(linha[const.COL_CONTATOS - 1].value)
contatos_list.pop(0)
contatos = tuple(contatos_list)


# abrir gmail e testar erros de entrada
erro = True

while erro == True:
    try:
        erro = False

        gmail_user = fun.gmail_user_reciver()
        
        gmail_password = pwinput("Preciso da sua senha, rapidinho: ")

        # vamos fazer login na conta gmail e sobreescrever a senha
        server = fun.gmail_loginer(gmail_user, gmail_password)
        gmail_password = "shananigans_de_segunranca_nao_gosto_de_hackeers"


    except TypeError as err:
        print(err)
        erro = True
        print("vamos de novo")
    except fun.ErroLoginGmail as err:
        print(err)
        print("vamos de novo!")
        erro = True
    except BaseException as err:
        print(err)
        quit()

# cria uma lista com todos os tipos de arquivos permitidos para anexar
loc_permitidos = getcwd() + "\permitidos.txt"
with open(loc_permitidos, "r") as f:
    arq_per = f.read()
    list_permitidos = arq_per.split() 
    f.close()

print("\n")

# criar e enviar um email por contato
# percorrer a lista de contatos adicionando uma linha a cada iteracao

linha = const.PRIMEIRA_LINHA

for email_reciver in contatos:

    coluna = const.COL_GAPS

    print("construindo texto", (linha - 1))
    
    # manipular o texto: construir uma matriz substituindo os @@@ pelos inputs do excel e recontruir como uma unica string
    texto = fun.construtor_texto(txt_list, linha, sheet)

    #construir o obj MIME
    msg = MIMEMultipart()
    msg['Subject'] = sheet.cell(row= linha, column= const.COL_ASSUNTOS).value
    msg['From'] = gmail_user
    msg['To'] = email_reciver
    body = MIMEText(texto)
    msg.attach(body)

    # anexar arquivos
    # olhar dentro da pasta arquivos e enviar de acordo com o tipo
    print("anexando arquivos")

    # abre a pasta de anexos e cria uma lista deles
    loc_arq = getcwd() + "\\anexos" 
    list_arq = listdir(loc_arq)
    
    # olhar se o arquivo eh permitido e colocar no mime
    i = 0
    for i in range(len(list_arq)):
        loc = loc_arq + "\\" + list_arq[i]
        fun.anexar(loc, list_permitidos, list_arq[i], msg)

    # enviar email
    print("enviando email...")

    try:
        r = server.sendmail(gmail_user, email_reciver, msg.as_string())
        print("enviado!")
        print("__________________")

    except smtplib.SMTPException as err:
        print ("algo deu errado no email para ", end= "")
        print (email_reciver)
        print("no excel eh a linha", linha)
        print("para geeks: SMTPException message - ", err, "\n")
        print("__________________\n")
    except RuntimeError as err:
        print ("algo deu errado no email para ", end= "")
        print (email_reciver)
        print("no excel eh a linha", linha)
        print("para geeks: RuntimeError message - ", err, "\n")
        print("__________________\n")
    except SystemExit as err:
        print ("algo deu errado no email para ", end= "")
        print (email_reciver)
        print("no excel eh a linha", linha)
        print("para geeks: SystemExit message - ", err)
        print("__________________\n")
    except BaseException as err:
        print(err)
        print ("algo deu errado no email para ", end= "")
        print (email_reciver)
        print("no excel eh a linha", linha)
        print("__________________\n")
        quit()
    
    linha = linha + 1
    del msg

server.close()

print("acabou!")



    




