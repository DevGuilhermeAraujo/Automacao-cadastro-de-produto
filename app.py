import pyautogui
import pyperclip
import openpyxl

workbook = openpyxl.load_workbook('produtos.xlsx')
produtos = workbook['Produtos']

for linha in produtos.iter_rows(min_row=2):
    nomeProduto = linha[0].value
    pyperclip.copy(nomeProduto)
    pyautogui.click(1274,338, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1248,455, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1268,558, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    codProduto = linha[3].value
    pyperclip.copy(codProduto)
    pyautogui.click(1271,648, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1260,739, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1292,822, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    
    pyautogui.click(1293,882, duration=1)
    
    
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1220,371, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    quantEstoque = linha[7].value
    pyperclip.copy(quantEstoque)
    pyautogui.click(1204,451, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    dtValidade = linha[8].value
    pyperclip.copy(dtValidade)
    pyautogui.click(1264,542, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1233,639, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    tamanho = linha[10].value
    pyautogui.click(1215,711, duration=1)
    
    if tamanho == 'Pequeno':
        pyautogui.click(1200,742, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1200,768, duration=1)
    else:
         pyautogui.click(1200,786, duration=1)
    
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1191,792, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    
    pyautogui.click(1357,852, duration=1)
    
    
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1237,388, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    paisOrigem = linha[13].value
    pyperclip.copy(paisOrigem)
    pyautogui.click(1252,477, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1231,597, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    codBarras = linha[15].value
    pyperclip.copy(codBarras)
    pyautogui.click(1248,686, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    locArmazem = linha[16].value
    pyperclip.copy(locArmazem)
    pyautogui.click(1246,778, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    
    pyautogui.click(1368,830, duration=1)

    pyautogui.click(1211,619, duration=1)
