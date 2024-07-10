import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import xlwings as xw
import os

def motivacao_plot_heterorrelato():
    # Conectar ao workbook e à planilha ativa
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    # Ler os valores de norma
    Norma_autorrelato = sheet.range('D12').value
    Norma_Heterorrelato = sheet.range('F12').value
    
    # Ler os valores de norma
    Dados_brutos_autorrelato = sheet.range('C12').value
    Dados_brutos_Heterorrelato = sheet.range('E12').value

    intervalo_confianca_autorrelato = (Norma_autorrelato - 5, Norma_autorrelato + 5)  
    intervalo_confianca_heterorrelato = (Norma_Heterorrelato - 5, Norma_Heterorrelato + 5)  

    # Dados para a tabela
    sub_escalas = [
        'Pontuação bruta autorrelato', 
        'Valor da norma autorrelato',
        'Respostas faltantes autorrelato',
        'Int. confiança autorrelato',
        'Pontuação bruta heterorrelato', 
        'Valor da norma heterorrelato', 
        'Respostas faltantes heterorrelato',
        'Int. confiança heterorrelato'
        ]
    
    valores = [
        Dados_brutos_autorrelato, 
        Norma_autorrelato,
        0,
        f'[{intervalo_confianca_autorrelato[0]} - {intervalo_confianca_autorrelato[1]}]',
        
        Dados_brutos_Heterorrelato,
        Norma_Heterorrelato,
        0,
        f'[{intervalo_confianca_heterorrelato[0]} - {intervalo_confianca_heterorrelato[1]}]'        
        ]
    
    table_data = [
        ['Pontuação bruta autorrelato', Dados_brutos_autorrelato],
        ['Valor da norma autorrelato', Norma_autorrelato],
        ['Respostas faltantes autorrelato', 0],
        ['Int. confiança autorrelato', f'[{intervalo_confianca_autorrelato[0]} - {intervalo_confianca_autorrelato[1]}]'],
        ['Pontuação bruta heterorrelato', Dados_brutos_Heterorrelato], 
        ['Valor da norma heterorrelato', Norma_Heterorrelato],
        ['Respostas faltantes heterorrelato', 0],
        ['Int. confiança heterorrelato', f'[{intervalo_confianca_heterorrelato[0]} - {intervalo_confianca_heterorrelato[1]}]']
    ]
    
    table_df = pd.DataFrame(table_data, columns=['', 'Valor'])

    # Configuração da Figura
    fig, ax = plt.subplots(figsize=(13, 5))  # Reduzir o tamanho da figura

#    # Plotar a curva normal
    mean = 50
    std_dev = (80 - 20) / 6  # Aproximadamente 10
    x = np.linspace(20, 80, 200) # indica o inicio e o fim do gráfico
    y = (1/(np.sqrt(2*np.pi*std_dev**2))) * np.exp(-0.5*((x-50)/std_dev)**2)  # Distribuição normal com média=50, desvio padrão=15


    # Ajustar o valor do desvio padrão para que a cauda direita toque o eixo x no mesmo valor do ponto vermelho
    std_dev_adjusted = std_dev
    y_adjusted = (1/(np.sqrt(2*np.pi*std_dev_adjusted**2))) * np.exp(-0.5*((x-mean)/std_dev_adjusted)**2)
    ax.plot(x, y_adjusted, color='lightcyan')
    ax.fill_between(x, y_adjusted, color='lightcyan', alpha=0.3)

    # Destacar o valor da norma autorrelato
    # cor azul
    norm_y_value = (1/(np.sqrt(2*np.pi*std_dev**2))) * np.exp(-0.5*((Norma_autorrelato-mean)/std_dev)**2) # define a posição da curva do gráfico
    ax.plot(Norma_autorrelato, norm_y_value, 'ro', color = 'blue')  # Ponto Norma autorrelato
    ax.vlines(Norma_autorrelato, 0, norm_y_value, colors='blue', linestyles='dotted') # Linha Ponto Norma Autorrelato > Eixo X

    # Destacar o valor da norma heterorrelato
    # cor vermelha
    norm_y_value = (1/(np.sqrt(2*np.pi*std_dev**2))) * np.exp(-0.5*((Norma_Heterorrelato-mean)/std_dev)**2)
    ax.plot(Norma_Heterorrelato, norm_y_value, 'ro', color = 'red')  # Norma autorrelato
    ax.vlines(Norma_Heterorrelato, 0, norm_y_value, colors='red', linestyles='dotted')

    # Sombrear a área entre 40 e 60 com a cor #E5F9F7
    x_shade = np.linspace(40, 60, 100) # indica os ranges inicio e fim da area a ser sombreada
    y_shade = (1/(np.sqrt(2*np.pi*std_dev_adjusted**2))) * np.exp(-0.5*((x_shade-50)/std_dev_adjusted)**2)
    ax.fill_between(x_shade, y_shade, color='#E5F9F7', alpha=0.8)

    # Configurar os eixos
    ax.set_xlim(0, Norma_autorrelato)
    ax.set_ylim(0, max(y)*2.5)
    ax.set_xticks(np.arange(20, 100, 10))
    ax.set_yticks([])

    # Adicionar a tabela à esquerda - escala
    raw_data_table = [[raw] for raw in sub_escalas]
    cell_height = 1 / len(sub_escalas)
    raw_table_ax = fig.add_axes([0.03, 0.1, 0.15, 0.75])
    raw_table_ax.axis('off')
    raw_table = raw_table_ax.table(cellText=raw_data_table, cellLoc='center', loc='center', bbox=[0, 0, 1, 1])
    for key, cell in raw_table.get_celld().items():
        cell.set_height(cell_height)
        cell.set_facecolor('lightcyan')


    # Adicionar a tabela à esquerda - escala
    valores_data_table = [[valores] for valores in valores]
    cell_height = 1 / len(valores)
    valores_table_ax = fig.add_axes([0.16, 0.1, 0.15, 0.75])
    valores_table_ax.axis('off')
    valores_table = valores_table_ax.table(cellText=valores_data_table, cellLoc='center', loc='center', bbox=[0, 0, 1, 1])
    for key, cell in valores_table.get_celld().items():
        cell.set_height(cell_height)
        cell.set_facecolor('lightcyan')


    # Ajustar o layout para garantir que todos os elementos estejam bem posicionados
    plt.subplots_adjust(left=0.25, right=0.85, top=0.85, bottom=0.1)

    # Salvar o gráfico como imagem
    nome_arquivo_png = 'motivacao.png'
    plt.savefig(nome_arquivo_png, bbox_inches='tight')
    plt.show()


    sheet.pictures.add(nome_arquivo_png, name='motivacao', update=True, top=sheet.range('T135').top, left=sheet.range('T135').left, height=400, width=700)
    #Remover o arquivo temporário
    os.remove(nome_arquivo_png)

if __name__ == "__main__":
    xw.Book("SRS-2.xlsm").set_mock_caller()
    motivacao_plot_heterorrelato()
