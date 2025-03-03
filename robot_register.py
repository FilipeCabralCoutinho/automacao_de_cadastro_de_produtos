# Esse é um projeto de automação de uma tarefa repetitiva de cadastro de um produto.
# O programava vai obter as insformações do arquivo em excel, acessará o site do formulário de cadastro,
# vai inserir as informações e enviar o formulário.

import pandas as pd
import pyautogui as auto
import time

# definir um tempo de intervalo padrão para cada etapa
auto.PAUSE = 0.5

# obter arquivo com os dados
table = pd.read_excel('Produtos.xlsx')

# abrir o formulário
auto.press('win')
auto.write('edge')
auto.press('enter')
auto.write('https://forms.office.com/r/yiDmntkg8v')
time.sleep(2)
auto.press('enter')

# laço para inserir os dados e enviar formulário
for row in table:
    cod_produto = table.loc['CodProduto']
    # descricao = table.loc['Descrição']
    # cod_marca = table.loc['CodMarca']
    # preco = table.loc['Preço Unitário']
    # categoria = table.loc['Categoria']

    auto.write()
    

