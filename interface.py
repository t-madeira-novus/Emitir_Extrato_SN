from appJar import gui
from funcoes import _comecar


def _converte_mes(mes):
    if str(mes) == "Janeiro":
        return "01"
    elif str(mes) == "Fevereiro":
        return "02"
    elif str(mes) == "Março":
        return "03"
    elif str(mes) == "Abril":
        return "04"
    elif str(mes) == "Maio":
        return "05"
    elif str(mes) == "Junho":
        return "06"
    elif str(mes) == "Julho":
        return "07"
    elif str(mes) == "Agosto":
        return "08"
    elif str(mes) == "Setembro":
        return "09"
    elif str(mes) == "Outubro":
        return "10"
    elif str(mes) == "Novembro":
        return "11"
    elif str(mes) == "Dezembro":
        return "12"


def _ajuda(submenu):
    aux = ''
    if submenu == 'Como usar':
        app.infoBox('Como usar', aux)
    elif submenu == 'Versão':
        app.infoBox('Versão', 'Versão 1.0')


def _thread_carregar_empresas():
    app.thread(_carregar_empresas())


def _carregar_empresas():
    global df
    df = app.openBox(title=None, dirName=None, fileTypes=None, asFile=False, parent=None, multiple=False, mode='r')


def thread_comecar():
    ano = app.getOptionBox("Ano: ")
    mes = app.getOptionBox("Mês: ")
    if mes is not None and ano is not None:
        df = app.openBox(title=None, dirName=None, fileTypes=None, asFile=False, parent=None, multiple=False, mode='r')
        mes = _converte_mes(mes)
        app.thread(_comecar(df, ano, mes))
    else:
        app.infoBox("Erooou...", "Mês ou Ano não selecionado")



# Variáveis globais
df = ''

# Criando a interface Gráfica
app = gui("Emitir Extrato do Simples Nacional")
app.setFont(10)
app.addMenuList("Ajuda", ["Como usar", "Versão"], _ajuda)

coluna = 0
linha = 0
app.addLabelOptionBox("Mês: ", ["- Mês -", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 'Julho',
                                "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"],  row=linha, column=coluna)
linha += 1
app.addLabelOptionBox("Ano: ", ["- Ano -", "2020", "2021"],  row=linha, column=coluna)
linha += 1
########################################################################################################################
coluna = 1
linha = 0
# app.addButton("Carregar empresas", _thread_carregar_empresas, row=linha, column=coluna)
# linha += 1
app.addButton("Começar", thread_comecar, row=linha, column=coluna)
linha += 1
########################################################################################################################
# Inicializa a GUI
app.go()
