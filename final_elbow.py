import pandas as pd
import numpy as np
from sklearn.cluster import KMeans
import openpyxl
import PySimpleGUI as sg
import matplotlib.pyplot as plt

#Define o caminho do arquivo Excel e a quantidade de grupos a serem analisados

arquivo = "bd_basic1.xlsx"


#Carrega o arquivo Excel em um DataFrame
dataframe = pd.read_excel(arquivo)
#Obtem os valores do DataFrame como uma matriz numpy
dados = dataframe.values
#Converte cada linha em um vetor distinto
vetores = []
    
for linha in dados:
    vetor = np.array(linha)
    vetores.append(vetor)

#Imprime os vetores
    for vetor in vetores:
        print(vetor)
        
#Atribue os dados da planilha (vetores) que serão analisados
X = np.array(vetores)

    
# Determinar o valor ideal de K usando o método do cotovelo
sse = []
k_values = range(1, 5)

for k in k_values:
    kmeans = KMeans(n_clusters=k, random_state=42)
    kmeans.fit(X)
    sse.append(kmeans.inertia_)


# Plotar a curva SSE versus K
plt.plot(k_values, sse, 'bo-')
plt.xlabel('Número de clusters (K)')
plt.ylabel('Soma dos quadrados das distâncias (SSE)')
plt.title('Método do Cotovelo')
plt.show()



