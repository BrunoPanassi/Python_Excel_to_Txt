# coding: utf-8

import pandas as pd
import datetime as dt
import os
import time
from xlrd import XLRDError
import sys
import argparse

# Iniciando os argumentos do terminal
parser = argparse.ArgumentParser()

# Colocando os argumentos
parser.add_argument('-txt', help='O caminho do arquivo.txt', type=str)
parser.add_argument('-planilha', help='O caminho da planilha', type=str)
parser.add_argument('-aba', help='O nome da aba da planilha', type=str)

args = parser.parse_args()

SortedColumns = False
SortedColumnsFlag = ''
BreakLoop = False

def VerificaArquivo(Directory):
    Exists = False
    if os.path.isfile(Directory):
        Exists = True
        print('Diretório do arquivo encontrado!\n')
    else:
        print('Diretório do arquivo não encontrado!\n')
    return Exists

PathExists = False
while PathExists == False:
    Txt_Directory = str(args.txt)
    PathExists = VerificaArquivo(Txt_Directory)

PathExists = False
while PathExists == False:
    Excel_Directory = str(args.planilha)
    PathExists = VerificaArquivo(Excel_Directory)
    
PathExists = False
while PathExists == False:
    SheetName = str(args.aba)
    
    try:
        Excel_Data = pd.read_excel(Excel_Directory, sheet_name = SheetName)
        PathExists = True
    except XLRDError:
        print('Aba não encontrada!\n')
        
df = pd.DataFrame(Excel_Data, columns=['Empr', 'CL', 'Conta', 'Valor do Montante', 'Elemento PEP', 'Chv.ref.1',
                                       'Data do Doc', 'Contrato', 'Data Lançamento', 'Denominação', 'Interface'])

#Abro o arquivo para gravar
f = open(Txt_Directory, "w")

#Começo a contar o tempo de execução
BeginTime = time.perf_counter()

Linha = 0
for i in range(df.shape[0]):
    
    if (BreakLoop == True):
        break
    
    print('Linha ', Linha)

    if(Linha > 0):
        f.write('\n')
    
    Linha += 1
    
    for j in range(df.shape[1]):
        
        #Empresa | 1º Column
        if (j == 0):
            Empr = str(df.iloc[i, j])
            
            if (len(Empr) < 4): 
                Empr = Empr.zfill(5)

                print('Aviso: A informação de empresa está menor que 4! Empresa: ', str(Empr))

                f.write("&SdtTexto.Add('" + Empr)
            elif(len(Empr) == 4):
                Empr = Empr.zfill(5)

                f.write("&SdtTexto.Add('" + Empr)
            else:
                BreakLoop = True
                print('Codigo da Empresa está com o tamanho errado! Tamanho: ' + str(len(Empr)))
                
        #Credito ou Debito | 2º Column
        elif(j == 1):
            CD = str(df.iloc[i, j])
            
            if (len(CD) == 1):
                f.write(CD)
            else:
                BreakLoop = True
                print('Informação de Crédito ou Débito com tamanho errado! Tamanho: ' + str(len(CD)))

        #Conta | 3º Column
        elif(j == 2):
            Conta = str(df.iloc[i, j])
            
            if (len(Conta) == 10):
                f.write(Conta)
            else:
                BreakLoop = True
                print('Informação de Conta está com tamanho errado! Tamanho: ' + str(len(Conta)))

        #Valor do Montante | 4º Column
        elif(j == 3):
            ValorDoMontante = float(str(df.iloc[i, j]))
            
            ValorDoMontante = '{0:.2f}'.format(ValorDoMontante)
            
            ValorDoMontante = str(ValorDoMontante).replace(".","").zfill(15)

            if (len(ValorDoMontante) == 15):
                f.write(ValorDoMontante)
            else:
                BreakLoop = True
                print('O valor do montante está com o tamanho errado! Tamanho: ' + str(len(ValorDoMontante)))

        #PEP | 5º Column
        elif(j == 4):
            PEP = str(df.iloc[i, j])
            
            if (len(PEP) == 15):
                PEP = PEP + '        '
                f.write(PEP)
            else:
                BreakLoop = True
                print('A informação de PEP está com o tamanho errado! Tamanho: ' + str(len(PEP)))

        #Chave Referencia | 6º Column
        elif(j == 5):
            ChaveRef = str(df.iloc[i, j])
            
            if (ChaveRef == 'nan'):
                ChaveRef = '            '
                f.write(ChaveRef)
            elif (len(ChaveRef) == 12):
                f.write(ChaveRef)
            elif (len(ChaveRef) < 12):
                ChaveRef = ChaveRef.ljust(12)
                f.write(ChaveRef)
                print('Aviso! A informação de ChaveRef está menor que 12! ChaveRef: ', str(ChaveRef))
            else:
                BreakLoop = True
                print('A informação de Chave Ref. está com o tamanho errado! Tamanho: ' + str(len(ChaveRef)))

        #Data do Documento | 7º Column      
        elif(j == 6):
            DataDoDocumento = str(df.iloc[i, j])
            #Atraso na velocidade da execução
            DataDoDocumento = dt.datetime.strptime(DataDoDocumento, '%Y-%m-%d %H:%M:%S').strftime('%Y%m%d')
            
            if (len(DataDoDocumento) == 8):
                f.write(DataDoDocumento)
            else:
                BreakLoop = True
                print('A informação da Data do Documento está com o tamanho errado! Tamanho: ' + str(len(DataDoDocumento)))
                
        #Contrato | 8º Column
        elif(j == 7):
            Contrato = str(df.iloc[i, j])
            
            if (len(Contrato) < 6):
                print('Aviso: A informação de contrato está menor que 6 - Contrato: ', Contrato)
                Contrato = Contrato.zfill(6)
                f.write(Contrato)
            elif (len(Contrato) > 6):
                BreakLoop = True
                print('A informação de Contrato está menor que o normal! Tamanho: ', str(len(Contrato)))
            else:
                f.write(Contrato)
        
        #Data do Lançamento | 9º Column
        elif(j == 8):
            DataDoLancamento = str(df.iloc[i, j])
            
            DataDoLancamento = dt.datetime.strptime(DataDoLancamento, '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y')
            
            if (len(DataDoLancamento) == 10):
                f.write(DataDoLancamento)
            else:
                BreakLoop = True
                print('A informação de Data do Lancamento está errada! Tamanho: ', str(len(DataDoLancamento)))
        
        #Histórico | 10º Column
        elif(j == 9):              
            Historico = str(df.iloc[i, j])

            QuotationMark = "')"

            if (len(Historico) > 50):
                Historico = Historico.format(Historico, 50)
                Historico += QuotationMark
                print('Aviso! O tamanho da informação de historico veio maior que 50')
            elif (len(Historico) == 50):
                Historico += QuotationMark
            else:
                SpacesToFill = 52 - len(Historico)
                QuotationMark = QuotationMark.rjust(SpacesToFill)
                Historico += QuotationMark

            if (len(Historico) == 52):
                f.write(Historico)
            else:
                BreakLoop = True
                print('A informação de Histórico está errada! Tamanho: ', str(len(Historico)))

        #Interface | 11º Column
        elif(j == 10):
            Interface = str(df.iloc[i, j])
            Empr_OriginValue = str(df.iloc[i, 0])  

            if(Linha < df.shape[0]):
                NextInterface = str(df.iloc[i+1, j])
                NextEmpr      = str(df.iloc[i+1, 0])
                if(Interface != NextInterface) or (Empr_OriginValue != NextEmpr):
                    f.write("\nDo 'Processar'")
            elif(Linha == df.shape[0]):
                f.write("\nDo 'Processar'")
            
f.close()
EndTime = time.perf_counter()
ProcessTime = EndTime - BeginTime
FormatTime = '{0:.2f}'.format(ProcessTime)

print('Tempo de processamento: ' + str(FormatTime) + ' segundos.')