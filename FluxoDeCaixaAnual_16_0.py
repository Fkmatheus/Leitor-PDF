
import tabula
import pandas as pd
import time
import os
import win32com.client
import tkinter as tk
from tkinter import filedialog
from tkinter import font

teste_renda_agro_1 = False
teste_renda_agro_2 = False


# Variavel Global para o arquivo PDF
ARQUIVO_LOCAL_PDF = ""

''' Janela da interface gráfica feita com TKinter para o usuario
    Escolher o arquivo pdf para leitura '''

def selecionar_arquivo():
    global ARQUIVO_LOCAL_PDF
    
    # Abre a janela para selecionar o arquivo
    arquivo = filedialog.askopenfilename()
    if arquivo:
        # Aqui você pode adicionar o código para utilizar o arquivo selecionado
        print()
        ARQUIVO_LOCAL_PDF = arquivo
        
        root.destroy()

# Cria a janela principal
root = tk.Tk()
root.title("Selecionar Arquivo")

fonte_elegante = font.Font(family="Helvetica", size=12, weight="bold")

# Cria um botão para selecionar o arquivo
botao = tk.Button(root, text="Selecionar Arquivo", command=selecionar_arquivo, font=fonte_elegante, fg="black",bg="lightgreen")
botao.pack(pady=20)

# Inicia o loop da interface gráfica
root.mainloop()
'''------------------------------------------------------------------------------------------------------------'''

print('Lendo o PDF...')
print()
# Lendo o PDFt
try:
    lista_tabelas = tabula.read_pdf(ARQUIVO_LOCAL_PDF, pages="all")
except:
    print("Arquivo invalido")
    exit()
    
# Verificando o ultimo ano
base_de_anos = []
for tabela in lista_tabelas:
    
    df_tabela = pd.DataFrame(tabela)
    
    linhas, colunas = df_tabela.shape

    infocolunas = list(df_tabela.columns)

    if colunas == 9:
        base_de_anos.append(infocolunas[0])

# Verificando o indice do ultimo ano
indice_ultimo_ano = len(base_de_anos) - 1

# Fornecendo o ultimo ano a partir do ultimo indice da lista
ultimo_ano = base_de_anos[indice_ultimo_ano]


# Iterando sobre todas as tabelas do pdf
for tabela in lista_tabelas:
    
    df_tabela = pd.DataFrame(tabela)
    
    linhas, colunas = df_tabela.shape

    infocolunas = list(df_tabela.columns)

    if colunas == 9 and infocolunas[0] == ultimo_ano:
        infocolunas = list(df_tabela.columns)
        
        #Ano
        Ano = infocolunas[0]

        

        if Ano == ultimo_ano:

            #display(df_tabela)
            
            #Produto
            Empreendimento = infocolunas[1]
            Empreendimento = Empreendimento.split("-")
            Analise = Empreendimento[1]
            Analise = Analise.split(" ")
            Produto = Analise[0]

            #Area
            area = infocolunas[2]

            #Produtividade 
            produtividade = infocolunas[3]

            #Produção TOTAL
            producao_total = infocolunas[4]

            #Valor unitário
            valor_unitario = infocolunas[5]

            #Valor TOTAL
            valor_total = infocolunas[6]

            #Frust
            frust = infocolunas[7]

            #Seguro
            seguro = infocolunas[8]


            teste_renda_agro_1 = True

            
            #print(df_tabela.loc[0])
            #display(df_tabela)
            print()


    if colunas == 9 and linhas > 0:
        infocolunas = list(df_tabela.columns)
        
        
        #Ano
        Ano = infocolunas[0]

        

        if Ano == ultimo_ano:

            lista_de_anos = []
            lista_de_empreendimentos = []
            lista_de_areas = []
            lista_de_produtividade = []
            lista_de_producao_total = []
            lista_de_valor_unitario = []
            lista_de_valor_total = []
            lista_de_frust = []
            lista_de_seguros = []

            lista_de_anos.append(infocolunas[0])
            
            #Tirando os traços do produto
            if '-' in str(infocolunas[1]):
                analise_produto = str(infocolunas[1]).split("-")
                empreendimento_str = analise_produto[1]
                analise_produto = empreendimento_str.split(" ")
                empreendimento_str = analise_produto[0]
                
                lista_de_empreendimentos.append(empreendimento_str)
            
            lista_de_areas.append(infocolunas[2])
            lista_de_produtividade.append(infocolunas[3])
            lista_de_producao_total.append(infocolunas[4])
            lista_de_valor_unitario.append(infocolunas[5])
            lista_de_valor_total.append(infocolunas[6])
            lista_de_frust.append(infocolunas[7])
            lista_de_seguros.append(infocolunas[8])
            
            
            for idx, row in df_tabela.iterrows():

                # ANOS
                ano_str = str(row[ultimo_ano])
                ano_verificacao = ano_str.split('.')
                ano_str = ano_verificacao[0]

                if ano_str == ultimo_ano:
                    lista_de_anos.append(row[ultimo_ano])
                    #print(lista_de_anos)

                # Empreendimentos
                empreendimento_str = str(row[infocolunas[1]])

                if '-' in empreendimento_str:
                    analise_produto = empreendimento_str.split("-")
                    empreendimento_str = analise_produto[1]
                    analise_produto = empreendimento_str.split(" ")
                    empreendimento_str = analise_produto[0]
                
                    lista_de_empreendimentos.append(empreendimento_str)
                    print()
                    #print(lista_de_empreendimentos)

                # AREAS
                area_str = str(row[infocolunas[2]])
                if area_str != 'nan':
                    lista_de_areas.append(area_str)
                    #print(lista_de_areas)

                # Produtividade
                produtividade_str = str(row[infocolunas[3]])
                if produtividade_str != 'nan':
                    lista_de_produtividade.append(produtividade_str)
                    #print(lista_de_produtividade)

                # Produção TOTAL
                producao_total_str = str(row[infocolunas[4]])
                if producao_total_str != 'nan':
                    lista_de_producao_total.append(producao_total_str)
                    #print(lista_de_producao_total)

                # Valor Unitário
                valor_unitario_str = str(row[infocolunas[5]])
                if valor_unitario_str != 'nan':
                    lista_de_valor_unitario.append(valor_unitario_str)
                    #print(lista_de_valor_unitario)

                # Valor TOTAL
                valor_total_str = str(row[infocolunas[6]])
                if valor_total_str != 'nan':
                    lista_de_valor_total.append(valor_total_str)
                    #print(lista_de_valor_total)

                # Frust
                frust_str = str(row[infocolunas[7]])
                if frust_str != 'nan':
                    lista_de_frust.append(frust_str)
                    #print(lista_de_frust)

                seguro_str = str(row[infocolunas[8]])
                if seguro_str != 'nan':
                    lista_de_seguros.append(seguro_str)
                    #print(lista_de_seguros)


            teste_renda_agro_2 = True
                    
            
print('Leitura de PDF concluida ✅')
         


print("Conectando ao Excel...")
if teste_renda_agro_1 == True:
    # Conectar ao Excel
    try: 
        excel = win32com.client.Dispatch("Excel.Application")
    except:
        print('Por favor, verifique se você está com a planilha da ferramenta de análise aberta e tente novamente.')

    # Obter a planilha ativa
    try:
        workbook = excel.ActiveWorkbook
    except:
        print('Por favor, certifique-se de não estar editando nenhuma célula enquanto a automação está funcionando')
        print('Tente novamente após alguns segundos')
        time.sleep(3)
        exit()
    worksheet = workbook.Sheets('Fluxo_Agro')

    ''' Limpando os dados na planilha '''

    # Limpando Produto
    worksheet.Cells(370, 14).Value = ""
    worksheet.Cells(370, 21).Value = ""
    worksheet.Cells(370, 28).Value = ""
    worksheet.Cells(370, 35).Value = ""
    worksheet.Cells(370, 42).Value = ""

    # Limpando Àrea
    worksheet.Cells(371, 14).Value = ""
    worksheet.Cells(371, 21).Value = ""
    worksheet.Cells(371, 28).Value = ""
    worksheet.Cells(371, 35).Value = ""
    worksheet.Cells(371, 42).Value = ""
    
    
    
    ''' Modificando dados na planilha '''
    
    # Modificando Produto
    worksheet.Cells(370, 14).Value = Produto
    
    # Modificando Àrea
    worksheet.Cells(371, 14).Value = area.replace('.','').replace(',','.')
    
    
    ''' Salvando alterações'''
    
    # Salvar a planilha
    workbook.Save()
    




if teste_renda_agro_2 == True:
    contador_de_produtos = len(lista_de_empreendimentos)
    
    # Conectando ao Excel
    try: 
        excel = win32com.client.Dispatch("Excel.Application")
    except:
        print('Por favor, verifique se você está com a planilha da ferramenta de análise aberta e tente novamente.')

    # Obter a planilha ativa
    try:
        workbook = excel.ActiveWorkbook
    except:
        print('Por favor, certifique-se de não estar editando nenhuma célula enquanto a automação está funcionando')
        print('Tente novamente após alguns segundos')
        time.sleep(3)
        exit()
    
    worksheet = workbook.Sheets('Fluxo_Agro')
    
    ''' Limpando os dados na planilha '''

    # Limpando Produto
    worksheet.Cells(370, 14).Value = ""
    worksheet.Cells(370, 21).Value = ""
    worksheet.Cells(370, 28).Value = ""
    worksheet.Cells(370, 35).Value = ""
    worksheet.Cells(370, 42).Value = ""

    # Limpando Àrea
    worksheet.Cells(371, 14).Value = ""
    worksheet.Cells(371, 21).Value = ""
    worksheet.Cells(371, 28).Value = ""
    worksheet.Cells(371, 35).Value = ""
    worksheet.Cells(371, 42).Value = ""
    

    
    indice = 0

    for x in range(contador_de_produtos):
        

        if indice == 0:
            coluna_celula = 14

        if indice == 1:
            coluna_celula = 21

        if indice == 2:
            coluna_celula = 28

        if indice == 3:
            coluna_celula = 35

        if indice == 4:
            coluna_celula = 42


        linha_celula = 370
        
        worksheet.Cells(linha_celula, coluna_celula).Value = lista_de_empreendimentos[indice].replace('.','').replace(',','.')
        
        linha_celula = linha_celula + 1
        
        worksheet.Cells(linha_celula, coluna_celula).Value = lista_de_areas[indice].replace('.','').replace(',','.')
        

        indice += 1
        

            




    workbook.Save()
    

print('Os Dados foram exportados para a ferramenta de análise com sucesso! ✅')

print()




            
