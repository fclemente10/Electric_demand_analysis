# -*- coding: utf-8 -*-
import matplotlib.pyplot as plt
from collections import OrderedDict
from xlrd import open_workbook
import time

import pymsgbox


class arquivo:
    #atributos do cabeçalho
    alimentador_01=0
    trafo_01 = 420                                              #define o valor do transformador dos alimentadores
    alimentador_02=0
    trafo_02 = 500
    alimentador_03=0
    trafo_03 = 320
    alimentador_04=0
    trafo_04 = 1800
    alimentador_05=0
    trafo_05 = 1800
    alimentador_06=0
    trafo_06 = 350
    alimentador_07=0
    trafo_07 = 250

    #atributos gerais
    data1 = 0
    ctrl_media =0

    #atributos controle de lista
    dic_dados_01 = {}
    lista_dados_01 = []
    dic_dados_02 = {}
    lista_dados_02 = []
    dic_dados_03 = {}
    lista_dados_03 = []
    dic_dados_04 = {}
    lista_dados_04 = []
    dic_dados_05 = {}
    lista_dados_05 = []
    dic_dados_06 = {}
    lista_dados_06 = []
    dic_dados_07 = {}
    lista_dados_07 = []

    intervalos=0
    qnt_car = 0
    result_pot = 0
    result_data = 0
    potencia = 0
    bloco_div = 0
    total_pot = 0
    nivel = 3

    valor_hora =0
    dic_aux_energia = {}


    D_bloco = {}
    C_bloco = {}
    B_bloco = {}
    A_bloco = {}
    total_bloco = {}

    def __init__(self):                     #extrai os dados da tabela para fazer processamento

        i=0
        wb = open_workbook('um dia.xlsx')

        for s in  wb.sheets():
            for row in range(s.nrows):
                values = []
                for col in range(s.ncols):
                    values.append(s.cell(row,col).value)
                i+=1
                if i<2:
                    self.busca_info_cabecalho(values)
                else:
                    self.busca_info_dados(values)
#        self.prepara_impressao()                           #comentado para teste

    def busca_info_cabecalho(self,valores):                 #retira da variavel valores o nome do cabeçalho da tabela

        self.alimentador_01 = valores[1]
        self.alimentador_02 = valores[2]
        self.alimentador_03 = valores[3]
        self.alimentador_04 = valores[4]
        self.alimentador_05 = valores[5]
        self.alimentador_06 = valores[6]
        self.alimentador_07 = valores[7]

    def busca_info_dados(self,valores):


        data = valores[0]                                                                                               #Bloco para trabalhar com a data
        data_hora1 = data.split(' ')
        self.data1 = (data_hora1[1])
        busca_00 = data.split(':')                                                                                      #coleta hora 00 para excluir amostra e deletar
                                                                                                                        #para calcular media certa
        self.valor_hora += 300                                                                                          #transforma string para hora
        self.hora_hora = time.strftime("%H:%M", time.gmtime(self.valor_hora))



        try:

#            if busca_00[1] != '00':                                                                                    #Bloco busca informaçao de energia ativa

            dados1 = float(valores[1])
            self.lista_dados_01.append(dados1)
            dados2 = float(valores[2])
            self.lista_dados_02.append(dados2)
            dados3 = float(valores[3])
            self.lista_dados_03.append(dados3)
            dados4 = float(valores[4])
            self.lista_dados_04.append(dados4)
            dados5 = float(valores[5])
            self.lista_dados_05.append(dados5)
            dados6 = float(valores[6])
            self.lista_dados_06.append(dados6)
            dados7 = float(valores[7])
            self.lista_dados_07.append(dados7)

            self.ctrl_media+=1                                                                          # controle agrupar dados de 5 em 5 minutos para 15 minutos
#                print "cont_media=",self.ctrl_media
            #else:
            print "Loading..."                                                                         #exclui amostra 00 para

        except(ValueError, NameError, IndexError):
            None

        if self.lista_dados_01 == []:   None                                                                            #Caso a lista esteja vazia não agrupa dados
        elif self.lista_dados_02 == []: None                                                                            #faz o programa competar a lista toda para média correta
        elif self.lista_dados_03 == []: None
        elif self.lista_dados_04 == []: None
        elif self.lista_dados_05 == []: None
        elif self.lista_dados_06 == []: None
        elif self.lista_dados_07 == []: None
        else: self.agrupa_dados()

    def agrupa_dados(self):

        if self.ctrl_media%3 == 0:                                                                                      #pega 3 amostras para calcular a media
            if self.hora_hora not in self.dic_dados_01:
#                print "hora in dic=",self.hora_hora
                self.dic_dados_01[self.hora_hora] = self.calc_media(self.lista_dados_01)                                    #cria dicionario de dados chave: data, conteudo: media
                self.lista_dados_01=[]                                                                                  #Limpa lista para buscar novas 3 amostras
#                print "dic dados 01=",self.dic_dados_01
            else:
                self.dic_aux = self.dic_dados_01.pop(self.hora_hora)
                teste = self.calc_media(self.lista_dados_01)
                self.lista_soma = (self.dic_aux + teste)/2                                                              #Soma média atual com media existente e divide por 2
#                print "hora in dic=",self.hora_hora
                self.dic_dados_01[self.hora_hora] = self.lista_soma
                self.lista_soma=[]
                self.lista_dados_01=[]
#-------------------------BLOCO GRUPO 2---------------------------------------------------------------------------------

            if self.hora_hora not in self.dic_dados_02:
                self.dic_dados_02[self.hora_hora] = self.calc_media(self.lista_dados_02)                                    #cria dicionario de dados chave: data, conteudo: media
                self.lista_dados_02=[]                                                                                  #Limpa lista para buscar novas 3 amostras
            else:
                self.dic_aux = self.dic_dados_02.pop(self.hora_hora)
                teste = self.calc_media(self.lista_dados_02)
                self.lista_soma = (self.dic_aux + teste)/2                                                              #Soma média atual com media existente e divide por 2
                self.dic_dados_02[self.hora_hora]= self.lista_soma
                self.lista_soma=[]
                self.lista_dados_02=[]
#-------------------------BLOCO GRUPO 3---------------------------------------------------------------------------------

            if self.hora_hora not in self.dic_dados_03:
                self.dic_dados_03[self.hora_hora] = self.calc_media(self.lista_dados_03)                                    #cria dicionario de dados chave: data, conteudo: media
                self.lista_dados_03=[]                                                                                  #Limpa lista para buscar novas 3 amostras
            else:
                self.dic_aux = self.dic_dados_03.pop(self.hora_hora)
                teste = self.calc_media(self.lista_dados_03)
                self.lista_soma = (self.dic_aux + teste)/2                                                              #Soma média atual com media existente e divide por 2
                self.dic_dados_03[self.hora_hora]= self.lista_soma
                self.lista_soma=[]
                self.lista_dados_03=[]
#-------------------------BLOCO GRUPO 4---------------------------------------------------------------------------------

            if self.hora_hora not in self.dic_dados_04:
                self.dic_dados_04[self.hora_hora] = self.calc_media(self.lista_dados_04)                                    #cria dicionario de dados chave: data, conteudo: media
                self.lista_dados_04=[]                                                                                  #Limpa lista para buscar novas 3 amostras
            else:
                self.dic_aux = self.dic_dados_04.pop(self.hora_hora)
                teste = self.calc_media(self.lista_dados_04)
                self.lista_soma = (self.dic_aux + teste)/2                                                              #Soma média atual com media existente e divide por 2
                self.dic_dados_04[self.hora_hora]= self.lista_soma
                self.lista_soma=[]
                self.lista_dados_04=[]
#-------------------------BLOCO GRUPO 5---------------------------------------------------------------------------------

            if self.hora_hora not in self.dic_dados_05:
                self.dic_dados_05[self.hora_hora] = self.calc_media(self.lista_dados_05)                                    #cria dicionario de dados chave: data, conteudo: media
                self.lista_dados_05=[]                                                                                  #Limpa lista para buscar novas 3 amostras
            else:
                self.dic_aux = self.dic_dados_05.pop(self.hora_hora)
                teste = self.calc_media(self.lista_dados_05)
                self.lista_soma = (self.dic_aux + teste)/2                                                                   #Soma média atual com media existente e divide por 2
                self.dic_dados_05[self.hora_hora]= self.lista_soma
                self.lista_soma=[]
                self.lista_dados_05=[]
#-------------------------BLOCO GRUPO 6---------------------------------------------------------------------------------
            if self.hora_hora not in self.dic_dados_06:
                self.dic_dados_06[self.hora_hora] = self.calc_media(self.lista_dados_06)                                    #cria dicionario de dados chave: data, conteudo: media
                self.lista_dados_06=[]                                                                                  #Limpa lista para buscar novas 3 amostras
            else:
                self.dic_aux = self.dic_dados_06.pop(self.hora_hora)
                teste = self.calc_media(self.lista_dados_06)
                self.lista_soma = (self.dic_aux + teste)/2                                                              #Soma média atual com media existente e divide por 2
                self.dic_dados_06[self.hora_hora]= self.lista_soma
                self.lista_soma=[]
                self.lista_dados_06=[]
#-------------------------BLOCO GRUPO 7---------------------------------------------------------------------------------
            if self.hora_hora not in self.dic_dados_07:
                self.dic_dados_07[self.hora_hora] = self.calc_media(self.lista_dados_07)                                    #cria dicionario de dados chave: data, conteudo: media
                self.lista_dados_07=[]                                                                                  #Limpa lista para buscar novas 3 amostras
            else:
                self.dic_aux = self.dic_dados_07.pop(self.hora_hora)
                teste = self.calc_media(self.lista_dados_07)
                self.lista_soma = (self.dic_aux + teste)/2                                                              #Soma média atual com media existente e divide por 2
                self.dic_dados_07[self.hora_hora]= self.lista_soma
                self.lista_soma=[]
                self.lista_dados_07=[]


    def prepara_impressao(self):

        self.dic_dados_01 = self.ordem(self.dic_dados_01)                                         #Ordena dicionario em função das chaves para plotar grafico correntamente
        self.soma_carga(1)
        self.plota(self.dic_dados_01,self.alimentador_01,1,self.trafo_01,'blue')   #self.plota(Valores Grafico, nome alimentador, numero do grafico, potencia do transformador,cor do grafico)

        self.dic_dados_02 = self.ordem(self.dic_dados_02)                                         #Ordena dicionario em função das chaves para plotar grafico correntamente
        self.soma_carga(2)
        self.plota(self.dic_dados_02,self.alimentador_02,2,self.trafo_02,'blue')

        self.dic_dados_03 = self.ordem(self.dic_dados_03)                                         #Ordena dicionario em função das chaves para plotar grafico correntamente
        self.soma_carga(3)
        self.plota(self.dic_dados_03,self.alimentador_03,3,self.trafo_03,'blue')

        self.dic_dados_04 = self.ordem(self.dic_dados_04)                                         #Ordena dicionario em função das chaves para plotar grafico correntamente
        self.soma_carga(4)
        self.plota(self.dic_dados_04,self.alimentador_04,4,self.trafo_04,'blue')

        self.dic_dados_05 = self.ordem(self.dic_dados_05)                                         #Ordena dicionario em função das chaves para plotar grafico correntamente
        self.soma_carga(5)
        self.plota(self.dic_dados_05,self.alimentador_05,5,self.trafo_05,'blue')

        self.dic_dados_06 = self.ordem(self.dic_dados_06)                                         #Ordena dicionario em função das chaves para plotar grafico correntamente
        self.soma_carga(6)
        self.plota(self.dic_dados_06,self.alimentador_06,6,self.trafo_06,'blue')

        self.dic_dados_07 = self.ordem(self.dic_dados_07)                                         #Ordena dicionario em função das chaves para plotar grafico correntamente
        self.soma_carga(7)
        self.plota(self.dic_dados_07,self.alimentador_07,7,self.trafo_07,'blue')


    def calc_media(self,media):                                                                                         #calcula a media de amostras
        self.media = media
        total = 0
        for i in self.media:
            total+=i
        return total/ len(self.media)

    def prepara_bloco(self,numero_blocos):
        self.numero_bloco = numero_blocos

        while self.numero_bloco != 0:
            try:
                if self.numero_bloco == 4:
                    self.potencia = float(pymsgbox.prompt("\nDigite a potencia dos veículos do Bloco 4 (5~100kw):"))
                    self.qnt_car = int(pymsgbox.prompt("\nDigite a quantidade de carros a modelar no Bloco 4:"))
                    self.hora = int(pymsgbox.prompt("\nDigite a Hora incial(0~24h)de carregamento do Bloco nº4:"))
                    self.calcula_bloco()
                    self.numero_bloco = 3
                elif self.numero_bloco == 3:
                    self.potencia = float(pymsgbox.prompt("\nDigite a potencia dos veículos do Bloco 3 (5~100kw):"))
                    self.qnt_car = int(pymsgbox.prompt("\nDigite a quantidade de carros a modelar no Bloco 3:"))
                    self.hora = int(pymsgbox.prompt("\nDigite a Hora incial(0~24h)de carregamento do Bloco nº3:"))
                    self.calcula_bloco()
                    self.numero_bloco = 2
                elif self.numero_bloco == 2:
                    self.potencia = float(pymsgbox.prompt("\nDigite a potencia dos veículos do Bloco 2 (5~100kw):"))
                    self.qnt_car = int(pymsgbox.prompt("\nDigite a quantidade de carros a modelar no Bloco 2:"))
                    self.hora = int(pymsgbox.prompt("\nDigite a Hora incial(0~24h)de carregamento do Bloco nº2:"))
                    self.calcula_bloco()
                    self.numero_bloco = 1
                elif self.numero_bloco == 1:
                    self.potencia = float(pymsgbox.prompt("\nDigite a potencia dos veículos do Bloco 1 (5~100kw):"))
                    self.qnt_car = int(pymsgbox.prompt("\nDigite a quantidade de carros a modelar no Bloco 1:"))
                    self.hora = int(pymsgbox.prompt("\nDigite a Hora incial(0~24h)de carregamento do Bloco nº1:"))
                    self.calcula_bloco()
                    self.numero_bloco = 0
#                    print "numero_bloco=",self.numero_bloco

            except (ValueError, NameError):
                pymsgbox.alert( "Preparando Bloco - Digite um numero inteiro positivo por favor")
            except (TypeError):
                pymsgbox.alert("O Programa foi finalizado!")
                break

    def calcula_bloco(self):

        self.nivel = 3
        check = 5

        while True:
            try:
                while self.nivel > 2:
                    self.nivel = int(pymsgbox.prompt("\nQual o nível de carregamento (1 ou 2) dos veículos?"))
                if self.nivel > 2:
                    pymsgbox.alert("Digite o 1 ou 2")
                else: break
                break
            except (ValueError, NameError):
                pymsgbox.alert("definindo nível de carregamento - Digite 1 ou 2 por favor")
            except (TypeError):
                pymsgbox.alert("O Programa foi finalizado!")
                break

        if self.nivel == 1:
            self.intervalos = int(self.potencia / 0.47)                                                    #Calcula a potencia para carregamento nivel 1
            self.total_pot = 0.47 * self.qnt_car                                                       #multiplica a potencia pela quantidade de veículos
        elif self.nivel == 2:
            self.intervalos = int(self.potencia / 1.87)                                                    #Calcula a potencia para carregamento nivel 1
            self.total_pot = 1.87 * self.qnt_car

        total = self.potencia*self.qnt_car
        pymsgbox.alert( ">>______________BLOCO=%s________________<<\n Potencia dos veículos %s KW \n Quantidade de veículos=%s \n \
hora inicial %s:00 \nTotal potencia veículos somados=%s Kw \n Serão distribuidas %s KWh em intervalos de \n 15 minutos \
totalizando %i horas pelo grafico" \
%(self.numero_bloco, self.potencia,self.qnt_car, self.hora, total, self.total_pot, self.intervalos/4))

        self.hora = self.hora*3600                                                            # converte inteiro para horas
#        print "hora total =",self.hora

        for i in range(self.intervalos):

#           c = time.strftime("%H:%M", time.gmtime(900))                                                #15 minutos
#           d = time.strftime("%H:%M", time.gmtime(3600))                                               # 1 hora
            if self.numero_bloco == 4:
                a = time.strftime("%H:%M", time.gmtime(self.hora))
                self.hora += 900                                                                   #soma 15 minutos na lista de horas
                self.D_bloco[a]= self.total_pot

            if self.numero_bloco == 3:
                a = time.strftime("%H:%M", time.gmtime(self.hora))
                self.hora += 900                                                                   #soma 15 minutos na lista de horas
                self.C_bloco[a]= self.total_pot

            if self.numero_bloco == 2:
                a = time.strftime("%H:%M", time.gmtime(self.hora))
                self.hora += 900                                                                   #soma 15 minutos na lista de horas
                self.B_bloco[a]= self.total_pot


            if self.numero_bloco == 1:
                a = time.strftime("%H:%M", time.gmtime(self.hora))
                self.hora += 900                                                                   #soma 15 minutos na lista de horas
                self.A_bloco[a] = self.total_pot

        self.total_bloco = {k: self.A_bloco.get(k,0) + self.B_bloco.get(k,0) + self.C_bloco.get(k,0) + self.D_bloco.get(k,0) for k in set(self.A_bloco) | set(self.B_bloco) | set(self.C_bloco) | set(self.D_bloco) }
#        print "total bloco=",self.total_bloco
        self.total_bloco = self.ordem(self.total_bloco)

        return

    def soma_carga(self,control):                                                                                    #soma valor da carga dos veiculos ao valor da energia ativa

        self.control = control
        lista_aux_carga = []
        valor_carga = 0
        valor_energia = 0
        teste1 = []
        espelho_bloco_total = self.total_bloco.copy()                                           #Cria lista espelho para retirar valor somar e no proximo grafico nao sumir o valor

        if self.control == 1 :
            self.dic_aux_energia = self.dic_dados_01.copy()
            lista_aux_carga = self.dic_dados_01.keys()

        if self.control == 2 :
            self.dic_aux_energia = self.dic_dados_02.copy()
            lista_aux_carga = self.dic_dados_02.keys()

        if self.control == 3 :
            self.dic_aux_energia = self.dic_dados_03.copy()
            lista_aux_carga = self.dic_dados_03.keys()

        if self.control == 4 :
            self.dic_aux_energia = self.dic_dados_04.copy()
            lista_aux_carga = self.dic_dados_04.keys()

        if self.control == 5 :
            self.dic_aux_energia = self.dic_dados_05.copy()
            lista_aux_carga = self.dic_dados_05.keys()

        if self.control == 6 :
            self.dic_aux_energia = self.dic_dados_06.copy()
            lista_aux_carga = self.dic_dados_06.keys()

        if self.control == 7 :
            self.dic_aux_energia = self.dic_dados_07.copy()
            lista_aux_carga = self.dic_dados_07.keys()

        for i in lista_aux_carga:

            if espelho_bloco_total.has_key(i):
                valor_carga = espelho_bloco_total.pop(i)
#                print "valor carga=",valor_carga
                valor_energia = self.dic_aux_energia.pop(i)
#                print "valor_energia=", valor_energia
                soma = valor_carga + valor_energia
                self.dic_aux_energia[i] = soma
            else:
                self.dic_aux_energia[i]=0

        self.dic_aux_energia = self.ordem(self.dic_aux_energia)                                         #Ordena dicionario em função das chaves para plotar grafico correntamente
#        self.plota(self.dic_aux_energia,self.alimentador_01,1,self.trafo_01,'red')
#        self.prepara_impressao()                                                      #comentado para teste

    def plota(self,graph,title,num_graph,transformador,cor):                                                                                               #função para plotar grafico

        self.graph = graph
        self.title = title
        self.num_graph = int(num_graph)
        self.transformador = transformador
        self.cor = cor
        lista_transformador = []
#        print"em plota graph=%s, title=%s, num_graph=%s " %(self.graph,self.title,self.num_graph)

        x = range(0,len(self.graph),1)
        for i in range(len(self.graph)):
            lista_transformador.append(self.transformador)

        plt.figure(self.num_graph)
        ax = plt.subplot(111)
        ax.bar(range(len(self.dic_aux_energia)),self.dic_aux_energia.values(),align='center',color='red')
        ax.bar(range(len(self.graph)), self.graph.values(), align='center',color=self.cor)
#        ax.bar(range(len(self.dic_dados_01)),self.dic_dados_01.values(),align='center',color='blue')

        plt.xticks(range(len(self.graph)), self.graph.keys(),rotation='90')
        plt.plot(x, lista_transformador, 'r--')
        plt.title(self.title+"  em um dia")
        plt.xlabel('Hora')
        plt.ylabel('KW')
        plt.grid('on')
        plt.margins(0)

        plt.subplots_adjust(left=0.06,bottom=0.11,right=0.99,top=0.95)           #        subplots_adjust(left=None, bottom=None, right=None, top=None, wspace=None, hspace=None)
        plt.show()               #colocado para teste de grafico potencia
#



    def ordem(self,a_ordenar):                                                                                          #função para ordenar dicionario
        dic_final = 0
        dic_final = OrderedDict(sorted(a_ordenar.items()))
        print "ordenado=", dic_final
        return dic_final


if __name__ == "__main__":

    pymsgbox.alert("        ESTE É UM PROGRAMA É PARA VERIFICAR O COMPORTAMENTO DA DEMANDA DE ENERGIA AO SE INSERIR VEICULOS ELÉTRICOS CARREGANDO \
EM DIFERENTES HORARIOS DO DIA COM DADOS DE DEMANDA REAL DA CIDADE DE NITEROI. \n    O programa funciona da seguinte maneira:\
 O usuário escolhera a quantidade de blocos de veículos que \
deseja modelar e a quantidade de veículos por blocos. \n        Os blocos serão divididos pelo gráfico de demanda real. A modelagem \
dos blocos será feita pelo usuário onde serão coletadas as informações de: --->A potência da bateria dos veiculos, --->A \
quantidade de veículos desejada por bloco --->O horário inicial de carregamento do bloco e ---> Nível de carregamento dos blocos.\
\n      A potência do transformador é uma variável definida no programa. Contudo, será possível criar um grafico de demanda\
 dos veículos juntamente com a demanda real de alguns bairros da cidade de Niteroi e analisar o comportamento da rede", title="INFORMAÇÕES")

    crtl = arquivo()

    bloco = 5

    while True:
        try:
            while bloco > 4 or bloco <= 0:
                bloco = int(pymsgbox.prompt('\nDigite a quantidade de blocos (1~4):'))

                if bloco > 4 or bloco <= 0:
                    pymsgbox.alert("Digite o valor de 1 a 4")
                else: break
            break

        except (ValueError, NameError):
            pymsgbox.alert("Digite um numero inteiro positivo por favor")
        except (TypeError):
            pymsgbox.alert("O Programa foi finalizado!")
            break

    crtl.prepara_bloco(bloco)
    crtl.prepara_impressao()

