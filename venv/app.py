import pyautogui
import pyperclip
import openpyxl
from time import sleep

workbook = openpyxl.load_workbook("dados.xlsx")
sheet_page = workbook["Produtos"]
for linha in sheet_page.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(193, 340, duration=1)
    pyautogui.hotkey("ctrl", "v")

    desc_produto = linha[1].value
    pyperclip.copy(desc_produto)
    pyautogui.click(195, 442, duration=1)
    pyautogui.hotkey("ctrl", "v")

    categoria_produto = linha[2].value
    pyperclip.copy(categoria_produto)
    pyautogui.click(199, 558, duration=1)
    pyautogui.hotkey("ctrl", "v")

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(183, 657, duration=1)
    pyautogui.hotkey("ctrl", "v")

    peso_produto = linha[4].value
    pyperclip.copy(peso_produto)
    pyautogui.click(186, 736, duration=1)
    pyautogui.hotkey("ctrl", "v")

    dimensoes_produto = linha[5].value
    pyperclip.copy(dimensoes_produto)
    pyautogui.click(185, 814, duration=1)
    pyautogui.hotkey("ctrl", "v")

    pyautogui.click(203, 883, duration=1)
    sleep(3)

    preco_produto = linha[6].value
    pyperclip.copy(preco_produto)
    pyautogui.click(180, 378, duration= 1)
    pyautogui.hotkey("ctrl", "v")

    quantidade_produto = linha[7].value
    pyperclip.copy(quantidade_produto)
    pyautogui.click(181, 448, duration= 1)
    pyautogui.hotkey("ctrl", "v")

    validade_produto = linha[8].value
    pyperclip.copy(validade_produto)
    pyautogui.click(190, 540, duration= 1)
    pyautogui.hotkey("ctrl", "v")

    cor_produto = linha[9].value
    pyperclip.copy(cor_produto)
    pyautogui.click(198, 618, duration= 1)
    pyautogui.hotkey("ctrl", "v")

    tamanho_produto = linha[10].value
    pyautogui.click(180, 717, duration= 1)
    if tamanho_produto == "Pequeno":
        pyautogui.click(189, 745, duration= 1)
    elif tamanho_produto == "Médio":
        pyautogui.click(189, 771, duration= 1)
    else:
        pyautogui.click(189, 791, duration= 1)
        
    material_produto = linha[11].value
    pyperclip.copy(material_produto)
    pyautogui.click(193, 795, duration= 1)
    pyautogui.hotkey("ctrl", "v")

    pyautogui.click(187, 863, duration= 1)


    fabricante_produto = linha[12].value
    pyperclip.copy(fabricante_produto)
    pyautogui.click(195, 406, duration= 1)
    pyautogui.hotkey("ctrl", "v")

    pais_origem_produto = linha[13].value
    pyperclip.copy(pais_origem_produto)
    pyautogui.click(191, 490, duration= 1)
    pyautogui.hotkey("ctrl", "v")

    obs_produto = linha[14].value
    pyperclip.copy(obs_produto)
    pyautogui.click(187, 604, duration= 1)
    pyautogui.hotkey("ctrl", "v")

    codbarras_produto = linha[15].value
    pyperclip.copy(codbarras_produto)
    pyautogui.click(180, 703, duration= 1)
    pyautogui.hotkey("ctrl", "v")

    localizacao_produto = linha[16].value
    pyperclip.copy(localizacao_produto)
    pyautogui.click(172, 800, duration= 1)
    pyautogui.hotkey("ctrl", "v")

    pyautogui.click(199, 853, duration= 1)
    pyautogui.click(676, 161, duration= 1)
    pyautogui.click(502, 621, duration= 1)




