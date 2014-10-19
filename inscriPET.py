#!/usr/bin/python
# -*- coding: UTF-8 -*-

'''
    SisPET - Sistema de Inscrição do PET Eng. Elétrica @Universidade Federal do Paraná {eletrica.ufpr.br/pet/}

    autor: Wendeurick Silverio
    repositório original: https://github.com/wsilverio/inscriPET

    licença: CC BY - http://creativecommons.org/licenses/

'''

'''
    A fazer:
        - Trocar para autorização via OAuth2
            - https://github.com/burnash/gspread#alternate-authorization-using-oauth2

        - Deixar o programa em loop
            - O usuário deve decidir quando fechar o programa

        - Função ler_dados()
            - Obter data e hora já formatadas, ex. "%d/%m/%y %H:%M"
                - https://docs.python.org/2/library/time.html#time.strftime
            - 'HORA' não precisa fazer parte da lista 'DADOS'
                - pode ser movida para a função enviar_dados()
            - O usuário deve confirmar antes de enviar os dados

        - Função ler_codigo_barras()
            - Timeout para a leitura
            - Fechar a janela do zbar após a leitura do código

        - Função enviar_dados()
            - Fazer a busca do GRR pelo comando 'find' e não "linha-a-linha"
                - https://github.com/burnash/gspread#finding-a-cell
'''

# planilhas, sistema, email, conexão (timeout), scanner
import gspread, os, smtplib, socket, zbar
# função teste_de_conexao()
from urllib2 import urlopen
# data e hora
from datetime import datetime
# função enviar_email()
from email.mime.text import MIMEText
# função remover_acentos()
from unicodedata import normalize
# usuário e senha contidos no arquivo PET.py
from PET import Sis_M, Sis_S, P_M, KEY_BD

def teste_de_conexao():
    socket.setdefaulttimeout(4)
    try :
        r = urlopen("http://google.com/")
    except:
        # fecha o programa caso nao haja conexao
        print "Sem conexão"
        exit()

def ler_codigo_barras():
    scanner.process_one() # aguarda o primeiro cod. barras
    for codigo in scanner.results:
        return str(codigo.data)

def ler_dados():

    COD = ""
    while True: # permanece em loop até que se obtenha um código valido

        print "Posicione a carteirinha em frente à webcam"

        COD = ler_codigo_barras()  # código de barras

        if COD.isdigit() and len(COD) == 12:
            print "Carteirinha: %s" %COD
            break
        else:
            print COD
            print "Erro de leitura. Tente novamente"

    NOME = raw_input("Nome completo: ")
    GRR = raw_input('GRR (ex.: 20119900): ')
    CURSO = raw_input('Curso (D/N/Física): ')
    PERIODO = raw_input('Período de ingresso (1/2): ')
    MAIL = raw_input('E-mail (ex.: aluno@mail.com): ')
    FONE = raw_input('Telefone (ex.: 41 9988-7766): ')

    # obtem data e hora no formato das planilhas Google
    t = datetime.now()
    HORA = str(str(t.day) + '/' + str(t.month) + '/' + str(t.year) +
               ' ' + str(t.hour) + ':' + str(t.minute) + ':' + str(t.second))

    NOME = remove_acentos(NOME).title()
    CURSO = remove_acentos(CURSO).title()

    if CURSO.upper() == 'D':
        CURSO = 'Eng Eletrica - Diurno'
    elif CURSO.upper() == 'N':
        CURSO = 'Eng Eletrica - Noturno'

    # verifica se o GRR é válido (somente números)
    while True:
        if GRR.isdigit(): break
        GRR = raw_input('\nGRR invalido. Por favor, corrija-o (ex.: 20119900): ')

    # verifica se o período é válido (/1 ou /2)
    while True:
        if PERIODO == '1' or PERIODO == '2': break
        PERIODO = raw_input('\nPeríodo inválido. Por favor, corrija-o (1/2): ')

    # verifica se o email é válido (se contém @ e .com ou .br)
    while True:
        if (('.com' in MAIL or '.br' in MAIL) and '@' in MAIL): break
        MAIL = raw_input('\nE-mail invalido. Por favor, corrija-o (ex.: aluno@mail.com): ')

    # captura o ano através do GRR: [2011]9900
    INGRESSO = GRR[0:4] + '/' + PERIODO

    # retorna lista com os dados
    return [HORA, COD, NOME, GRR, CURSO, INGRESSO, MAIL, FONE]

def remove_acentos(txt):
    # Substitui os caracteres por seus equivalentes "sem acento"
    return normalize('NFKD', txt.decode('utf-8')).encode('ASCII', 'ignore')

def enviar_dados(dados):
    print('Conectando ao Banco de Dados...')

    teste_de_conexao()

    try:
        # login no sistema Google
        cliente = gspread.login(Sis_M, Sis_S)
    except:
        print("Falha ao efetuar login no sistema Google\nVerifique os dados do arquivo PET.py")
        exit()
    else:
        try:
            # abre a planilha "Banco_de_Dados" pela "key"
            planilha = cliente.open_by_key(KEY_BD)
        except:
            print("Erro ao abrir o Banco de Dados\nVerifique a chave da planilha: ", KEY_BD)
            exit()
        else:
            # obtém a primeira página da planilha
            folha = planilha.get_worksheet(0)

            print('Conectado')
            
            i = 2  # índice que correrá a planilha. começa em 2 porque a primeira linha é o cabeçalho
            while True:
                celula = folha.cell(i, 2).value  # obtém o código de barras (coluna 2)
                
                # verifica se o cód. já está cadastrado
                if celula == dados[1]:
                    print("\nCódigo já cadastrado. Linha [%d]" % i)
                    s = raw_input("Substituir? (S/N): ").upper()
                    if s == 'S':
                        break
                    else:
                        exit()  # sai do programa
                # verifica se a celula está vazia
                elif not celula:  # celula == None
                    break  # 'i' contém o índice da linha em branco
                i += 1  # percorre a planilha linha a linha

            # define o intervalo para gravar os dados na linha 'i'
            # col A: "Indicação de data e hora" | ... | col H: "Telefone"
            # ex.: A6:H6
            cell_list = folha.range('A' + str(i) + ':H' + str(i))

            try:
                j = 0 # índice para as células
                for cell in cell_list:
                    cell.value = dados[j]  # escreve os valores nas colunas de A a H
                    j += 1
                # atualiza a planilha
                folha.update_cells(cell_list)
                print('\nDados gravados')
            except:
                print("Erro ao gravar dados")
                exit()

def envia_email(dados):
    print("\nPreparado mensagem de email")

    teste_de_conexao()

    # HORA, COD, NOME, GRR, CURSO, INGRESSO, MAIL, FONE

    nome = dados[2].split(' ')[0]  # obtém o primeiro nome
    nome_completo = dados[2]
    grr = dados[3]
    curso = dados[4]
    ingresso = dados[5]
    destinatario = dados[6]
    fone = dados[7]

    msg = MIMEText("Olá %s,\n"
                   "\nVocê foi cadastrado no SisPET - Sistema de Inscrição do PET Eng. Elétrica da UFPR.\n"
                   "Em caso de alteração em seus dados (telefone, email, etc), "
                   "envie um e-mail para ""pet@eletrica.ufpr.br"" ou dirija-se à sala da PET.\n\n"
                   "\tNome: %s\n"
                   "\tGRR: %s\n"
                   "\tCurso: %s\n"
                   "\tIngresso: %s\n"
                   "\tTelefone: %s\n"
                   "\nEsta é uma mensagem automática. Não responda.\n"
                   "\nGrupo PET Engenharia Elétrica"
                   "\nUniversidade Federal do Paraná"
                   "\n(41) 3361-3225"
                   "\nwww.eletrica.ufpr.br/pet" % (nome, nome_completo, grr, curso, ingresso, fone))

    msg['Subject'] = 'Cadastro - SisPET'
    msg['From'] = P_M
    msg['To'] = destinatario

    # servidor SMTP do GMAIL
    #print("Identificando servidor SMTP")
    mailServer = smtplib.SMTP('smtp.gmail.com', 587)
    # identificação de cliente
    #print("Identificando cliente")
    mailServer.ehlo()
    # criptografia tls
    #print("Criando conexão criptografada")
    mailServer.starttls()
    # identificação com conexão criptografada
    #print("Identificando conexão criptografada")
    mailServer.ehlo()

    try:
        print("Efetuando login")
        mailServer.login(Sis_M.replace('@gmail.com', ''), Sis_S)  # login com o nome de usuário
        print("Enviando email...")
        mailServer.sendmail(P_M, destinatario, msg.as_string())
    except:
        print("Erro ao enviar email\nVerifique os dados do arquivo PET.py e o email do destinatário")
        exit()

    mailServer.close()

    # notificação do notify-send
    #os.system('notify-send -i %s "sisPET" "Email enviado para %s"' %((os.popen('pwd').readline().rstrip() + "/emblem-mail.svg"),destinatario))

    print("\nEmail enviado para %s" % destinatario)

def configura_scanner():
    global scanner
    # cria um "Processor" zbar
    scanner = zbar.Processor()
    scanner.parse_config('enable')
    # substituir /dev/video0 pelo endereço do dispositivo da camera
    scanner.init('/dev/video0')
    # True: visualização habilitada
    # False: visualização desabilitada
    scanner.visible = True

if __name__ == '__main__':
    os.system('clear')

    print('\n### sisPET - V1.0 ###\n# Cadastro de usuário #\n')

    teste_de_conexao()

    configura_scanner()

    DADOS = ler_dados()

    # exibe os dados do usuário
    print('\n' + ' | '.join(DADOS) + '\n')
    # grava os dados no banco de dados
    enviar_dados(DADOS)
    # envia email com os dados do usuário
    envia_email(DADOS)
    # Sai do programa
    exit()