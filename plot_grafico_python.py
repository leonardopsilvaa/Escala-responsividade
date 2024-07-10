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

    # Ler os valores de normas da coluna D, linhas 9 a 15
    # Ler os valores de dados brutos da coluna C, linhas 9 a 15
    norm_scores = sheet.range(Range_norma_autorrelato).value
    raw_scores = sheet.range(Range_brutos_autorrelato).value

    # Escalas do gráfico
    subscales = [
        'Percepção Social', 'Cognição Social', 'Comunicação Social', 
        'Motivação Social', 'Padrões Restritos e Repetitivos', 
        'Comunicação e Interação Social', 'Escore Total'
    ]
    
    # Calcular os intervalos de confiança como norm_scores ± 5 com base no valor de norma
    lower_bounds = [score - 5 for score in norm_scores]
    upper_bounds = [score + 5 for score in norm_scores]

    # Calcular as barras de erro sem valores negativos, considerando os intervalos superior e inferior
    xerr_lower = [max(0, norm - lower) for norm, lower in zip(norm_scores, lower_bounds)]
    xerr_upper = [max(0, upper - norm) for norm, upper in zip(norm_scores, upper_bounds)]

    # Configuração da imagem
    fig, ax = plt.subplots(figsize=(13, 7))  # Aumentar o tamanho da figura
    ax.set_xlim(20, 85)  # Ajustar os limites do eixo X

    # Adicionar o fundo colorido para cada range
    ax.axvspan(20, 40, color='darkgray', alpha=0.5)
    ax.axvspan(40, 60, color='lightgray', alpha=0.5)
    ax.axvspan(60, 105, color='darkgray', alpha=0.5)

    # Configurar a escala do eixo X para ser de 5 em 5
    ax.set_xticks(np.arange(20, 105, 5))

    # Gráfico de Erro
    y_pos = np.arange(len(subscales))
    ax.errorbar(norm_scores, y_pos, xerr=[xerr_lower, xerr_upper], fmt='o', color='r', ecolor='black', capsize=5)
    ax.plot(norm_scores, y_pos, 'k-', alpha=0.5)

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
    fig.text(0.11, 0.90, 'Brutos', ha='right', va='center', color='black', fontsize=12, fontweight='bold')
    fig.text(0.22, 0.90, 'Normas', ha='right', va='center', color='black', fontsize=12, fontweight='bold')

    # Ajustar o layout para garantir que todos os elementos estejam bem posicionados
    plt.subplots_adjust(left=0.25, right=0.85, top=0.85, bottom=0.1)

    # Adicionar as tabelas de referência à esquerda
    raw_data_table = [[raw] for raw in raw_scores]
    norm_data_table = [[norm] for norm in norm_scores]

    # Calcular a altura das células para ajustar ao gráfico
    cell_height = 1 / len(subscales)

    # Adicionar a tabela ao lado do gráfico - Dados brutos
    raw_table_ax = fig.add_axes([0.07, 0.1, 0.09, 0.75])
    raw_table_ax.axis('off')
    raw_table = raw_table_ax.table(cellText=raw_data_table, cellLoc='center', loc='center', bbox=[0, 0, 1, 1])
    for key, cell in raw_table.get_celld().items():
        cell.set_height(cell_height)
        cell.set_facecolor('lightcyan')

    # Adicionar a tabela ao lado do gráfico - Norma
    norm_table_ax = fig.add_axes([0.16, 0.1, 0.09, 0.75])
    norm_table_ax.axis('off')
    norm_table = norm_table_ax.table(cellText=norm_data_table, cellLoc='center', loc='center', bbox=[0, 0, 1, 1])
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
