# Esse é um projeto de automação de uma tarefa repetitiva de cadastro de um produto.
# O programava vai obter as insformações do arquivo em excel, acessará o site do formulário de cadastro,
# vai inserir as informações e enviar o formulário.

import pandas as pd
import pyautogui as auto
import time

# definir um tempo de intervalo padrão para cada etapa
auto.PAUSE = 0.8

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
for row in table.itertuples():
    
    cod_produto = row.CodProduto
    descricao = row.descricao
    cod_marca = row.CodMarca
    preco = str(row.preco)
    categoria = row.Categoria
    
    time.sleep(1)
    auto.write(cod_produto)

    auto.press('tab', presses=2)
    auto.write(descricao)

    auto.press('tab', presses=2)
    auto.write(cod_marca)

    auto.press('tab', presses=2)
    auto.write(preco)

    auto.press('tab', presses=2)
    auto.write(categoria)
    
    auto.press('tab')
    auto.press('enter')

    time.sleep(2)
    auto.press('tab', presses=16)
    auto.press('enter')

    time.sleep(0.5)
    auto.press('tab', presses=4)
    
    

