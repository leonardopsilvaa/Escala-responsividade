import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import xlwings as xw
import os

Range_brutos_autorrelato = 'C9:C15'
Range_norma_autorrelato = 'D9:D15'
Range_brutos_heterorrelato = 'E9:E15'
Range_norma_heterorrelato = 'F9:F15'


def create_plot():
    # Conectar ao workbook e à planilha ativa
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    # Deletar todas as imagens da planilha, para gerar o novo plot
 #   for picture in sheet.pictures:
 #       picture.delete()

    # Ler os valores de normas da coluna D, linhas 9 a 15
    # Ler os valores de dados brutos da coluna C, linhas 9 a 15
    Norma_autorrelato = sheet.range(Range_norma_autorrelato).value
    Brutos_autorrelato = sheet.range(Range_brutos_autorrelato).value
    Norma_heterorrelato = sheet.range(Range_norma_heterorrelato).value
    Brutos_heterorrelato = sheet.range(Range_brutos_heterorrelato).value

    # Escalas do gráfico
    subscales = [
        'Percepção Social', 'Cognição Social', 'Comunicação Social', 
        'Motivação Social', 'Padrões Restritos e Repetitivos', 
        'Comunicação e Interação Social', 'Escore Total'
    ]
    
    # Calcular os intervalos de confiança como norm_scores ± 5 com base no valor de norma
    lower_bounds_autorrelato = [score - 5 for score in Norma_autorrelato]
    upper_bounds_autorrelato = [score + 5 for score in Norma_autorrelato]
    lower_bounds_heterorrelato = [score - 5 for score in Norma_heterorrelato]
    upper_bounds_heterorrelato = [score + 5 for score in Norma_heterorrelato]

    # Calcular as barras de erro sem valores negativos, considerando os intervalos superior e inferior
    xerr_lower_autorrelato = [max(0, norm - lower) for norm, lower in zip(Norma_autorrelato, lower_bounds_autorrelato)]
    xerr_upper_autorrelato = [max(0, upper - norm) for norm, upper in zip(Norma_autorrelato, upper_bounds_autorrelato)]
    xerr_lower_heterorrelato = [max(0, norm - lower) for norm, lower in zip(Norma_heterorrelato, lower_bounds_heterorrelato)]
    xerr_upper_heterorrelato = [max(0, upper - norm) for norm, upper in zip(Norma_heterorrelato, upper_bounds_heterorrelato)]

    # Configuração da imagem
    fig, ax = plt.subplots(figsize=(10, 7))  # Aumentar o tamanho da figura
    ax.set_xlim(20, 85)  # Ajustar os limites do eixo X

    # Adicionar o fundo colorido para cada range
    ax.axvspan(20, 40, color='darkgray', alpha=0.5)
    ax.axvspan(40, 60, color='lightgray', alpha=0.5)
    ax.axvspan(60, 105, color='darkgray', alpha=0.5)

    # Configurar a escala do eixo X para ser de 5 em 5
    ax.set_xticks(np.arange(20, 105, 5))

    # Gráfico de Erro
    y_pos = np.arange(len(subscales))
    ax.errorbar(Norma_autorrelato, y_pos, fmt='o', color='blue', ecolor='black')
    ax.errorbar(Norma_heterorrelato, y_pos, fmt='o', color='red', ecolor='black')
    ax.plot(Norma_autorrelato, y_pos, 'k-', alpha=0.5, color = 'blue')
    ax.plot(Norma_heterorrelato, y_pos, 'k-', alpha=0.5, color = 'red')

    # Configuração dos Eixos
    ax.set_yticks(y_pos)
    ax.set_yticklabels([])  # Remover as etiquetas do eixo Y do gráfico principal
    ax.invert_yaxis()  # Inverter o eixo Y para que o primeiro item apareça no topo
    ax.yaxis.tick_right()
    ax.xaxis.set_ticks_position('top')
    ax.xaxis.set_label_position('top')

    # Configuração do Grid
    ax.grid(True, which='both', linestyle='--', linewidth=0.5)

    # Adicionar cabeçalhos para colunas de dados brutos e normas
    fig.text(0.11, 0.90, 'Autorrelato', ha='right', va='center', color='black', fontsize=8, fontweight='bold')
    fig.text(0.22, 0.90, 'Heterorrelato', ha='right', va='center', color='black', fontsize=8, fontweight='bold')

    # Ajustar o layout para garantir que todos os elementos estejam bem posicionados
    plt.subplots_adjust(left=0.25, right=0.85, top=0.85, bottom=0.1)

    # Adicionar as tabelas de referência à esquerda
    tbl_brutos_autorrelato = [[raw] for raw in Brutos_autorrelato]
    tbl_normas_autorrelato = [[raw] for raw in Norma_autorrelato]
    tbl_brutos_heterorrelato = [[raw] for raw in Brutos_heterorrelato]
    tbl_normas_heterorrelato = [[raw] for raw in Norma_heterorrelato]

    # Calcular a altura das células para ajustar ao gráfico
    cell_height = 1 / len(subscales)

    
    #########################################
    #Inclusão das tabelas de referência
    # Adicionar a tabela ao lado do gráfico - Dados brutos autorrelato
    raw_table_ax = fig.add_axes([0.01, 0.1, 0.06, 0.75])
    raw_table_ax.axis('off')
    raw_table = raw_table_ax.table(cellText=tbl_brutos_autorrelato, cellLoc='center', loc='center', bbox=[0, 0, 1, 1])
    for key, cell in raw_table.get_celld().items():
        cell.set_height(cell_height)
        cell.set_facecolor('lightcyan')

    # Adicionar a tabela ao lado do gráfico - Norma autorrelato
    norm_table_ax = fig.add_axes([0.07, 0.1, 0.06, 0.75])
    norm_table_ax.axis('off')
    norm_table = norm_table_ax.table(cellText=tbl_normas_autorrelato, cellLoc='center', loc='center', bbox=[0, 0, 1, 1])
    for key, cell in norm_table.get_celld().items():
        cell.set_height(cell_height)
        cell.set_facecolor('lightcyan')

    # Adicionar a tabela ao lado do gráfico - Dados brutos heterorrelato
    raw_table_ax = fig.add_axes([0.13, 0.1, 0.06, 0.75])
    raw_table_ax.axis('off')
    raw_table = raw_table_ax.table(cellText=tbl_brutos_heterorrelato, cellLoc='center', loc='center', bbox=[0, 0, 1, 1])
    for key, cell in raw_table.get_celld().items():
        cell.set_height(cell_height)
        cell.set_facecolor('lightcyan')

    # Adicionar a tabela ao lado do gráfico - Norma heterorrelato
    norm_table_ax = fig.add_axes([0.19, 0.1, 0.06, 0.75])
    norm_table_ax.axis('off')
    norm_table = norm_table_ax.table(cellText=tbl_normas_heterorrelato, cellLoc='center', loc='center', bbox=[0, 0, 1, 1])
    for key, cell in norm_table.get_celld().items():
        cell.set_height(cell_height)
        cell.set_facecolor('lightcyan')



    # Adicionar a tabela de referência à direita sem a coluna "Escala"
    table_data = subscales

    # Adicionar a tabela ao lado do gráfico
    table_ax = fig.add_axes([0.86, 0.1, 0.20, 0.75])
    table_ax.axis('off')
    table = table_ax.table(cellText=[[item] for item in table_data], cellLoc='center', loc='center', bbox=[0, 0, 1, 1])
    for key, cell in table.get_celld().items():
        cell.set_height(cell_height)
        cell.set_facecolor('lightcyan')

    # Salvar o gráfico como imagem
    nome_arquivo_png = 'graph.png' 
    plt.savefig(nome_arquivo_png, bbox_inches='tight')  # Ajuste de bounding box
    plt.show()

    # Inserir a imagem no Excel a partir da célula T15
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    # Adicionar a nova imagem com tamanho específico
    sheet.pictures.add(nome_arquivo_png, name='Graph', update=True, top=sheet.range('T15').top, left=sheet.range('T15').left, height=400, width=700)
    
    # Remover o arquivo temporário
    os.remove(nome_arquivo_png)
if __name__ == "__main__":
    xw.Book("SRS-2.xlsm").set_mock_caller()
    create_plot()
