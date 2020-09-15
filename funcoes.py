import pandas as pd
import pyautogui
import time
import pyperclip


def _escrever_relatorio (mensagem, mes, ano):
    file = open("relatorio_emitir_extrato_SN_" + mes + "-" + ano + ".txt", "a+")
    file.write(mensagem + "\n")
    file.close()

def _clicar_pela_imagem(imagem, offsetX = 0, offsetY = 0, tentativas = 10):
    conf = 0.95
    aux = None
    c = 0
    while aux is None and c <= tentativas:
        aux = pyautogui.locateCenterOnScreen(imagem, confidence=conf)
        if aux is not None:
            x = aux[0] + offsetX
            y = aux[1] + offsetY
            pyautogui.click(x, y)
            return True
        print("Procurando: ", imagem)
        conf -= 0.01
        time.sleep(1)
        c += 1
        
    return False


def _comecar(df, mes, ano):
    df = pd.read_excel(df)

    if not _clicar_pela_imagem("imgs/icone_fiscal.png"):
        pyautogui.alert(text='O módulo do Domínio Fiscal não foi encontrado. Certifique-se de que ele esteja '
                         'aberto. Se estiver, chame Thiago Madeira para solucionar '
                         'este mistério misterioso.', title='Domínio Fiscal não encontrado', button='OK')
        return False

    pyautogui.press('esc', presses=5)

    for i in df.index:
        pyautogui.press('esc', presses=5)
        pyautogui.press('f8')  # Abrir busca de empresas
        print(df.at[i, "Código"])
        pyautogui.typewrite(str(df.at[i, "Código"]))  # Digitar id da empresa
        pyautogui.press('enter')
        time.sleep(10)  # Esperar empresa carregar

        path = "P:/documentos/OneDrive - Novus Contabilidade/Doc Compartilhado/Fiscal/Impostos/Federais/Empresas/"
        path_aux = path + str(df.at[i, "Código"]) + '-' + str(df.at[i, "Nome da Empresa"]) + '/' + str(ano) + '/' + str(
            mes) + '.' + str(ano)

        # Salvar relatório de serviços
        pyautogui.hotkey('alt', 'r')
        pyautogui.press(['a', 'v'])

        _clicar_pela_imagem("imgs/serie_docto.png")
        _clicar_pela_imagem("imgs/ok.png")

        pyautogui.hotkey('alt', 'r')
        pyautogui.press(['u', 'f', 's'])
        pyautogui.typewrite(str(mes)+"/"+str(ano))

        if not _clicar_pela_imagem("imgs/codigo_acesso.png"):
            _escrever_relatorio('Não achou o campo Código de Acesso', mes, ano)
            continue

        if not _clicar_pela_imagem("imgs/informar_caracteres.png"):
            _escrever_relatorio('Não achou o campo Informar Caracteres', mes, ano)
            continue

        if not _clicar_pela_imagem("imgs/extrato.png"):
            _escrever_relatorio('Não achou o campo Extrato', mes, ano)
            continue

        if not _clicar_pela_imagem("imgs/ok.png"):
            _escrever_relatorio('Não achou o campo OK', mes, ano)
            continue

        if not _clicar_pela_imagem("imgs/outros_dados.png"):
            _escrever_relatorio('Não achou janela \"Outros dados não digitados!\"', mes, ano)
            continue
        else:
            pyautogui.press('enter')

        if _clicar_pela_imagem("imgs/periodo_nao_processado.png"):
            _escrever_relatorio('Período não processado', mes, ano)
            continue
        else:
            pyautogui.press('enter')

        extrato = True
        while _clicar_pela_imagem('imgs/espera.png', tentativas=2):
            time.sleep(5)

        if _clicar_pela_imagem("imgs/aviso_naoExisteExtrato.png"):
            extrato = False
            pyautogui.press('enter')
            time.sleep(1)

            # Não tem extrato, pegar então o recibo
            if not _clicar_pela_imagem("imgs/recibo.png"):
                _escrever_relatorio('Não achou o campo Recibo', mes, ano)
                continue

            if not _clicar_pela_imagem("imgs/ok.png"):
                _escrever_relatorio('Não achou o campo OK', mes, ano)
                continue

            if not _clicar_pela_imagem("imgs/outros_dados.png"):
                _escrever_relatorio('Não achou janela \"Outros dados não digitados!\"', mes, ano)
                continue
            else:
                pyautogui.press('enter')

        while _clicar_pela_imagem('imgs/espera.png', tentativas=2):
            time.sleep(5)

        time.sleep(2)
        pyautogui.click(500, 500)
        pyautogui.hotkey('ctrl', 'p')
        time.sleep(1)
        pyautogui.press('enter')

        _clicar_pela_imagem('imgs/salvar_como.png')

        path = "P:/documentos/OneDrive - Novus Contabilidade/Doc Compartilhado/Fiscal/Impostos/Federais/Empresas/"
        path_aux = path + str(df.at[i, "Código"]) + '-' + str(df.at[i, "Nome da Empresa"]) + '/' + str(ano) + '/' + str(mes) + '.' + str(ano)

        pyperclip.copy(path_aux)
        pyautogui.hotkey("ctrl", "v")


        time.sleep(666)






_comecar('P:/documentos/OneDrive - Novus Contabilidade/Doc Compartilhado/Sistemas Internos/Fiscal/Emitir Extrato do SN/'
        'lista_empresas.xlsx', '06', '2020')