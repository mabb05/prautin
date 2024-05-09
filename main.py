import os,ctypes,easygui
import INS_ricoh as insr

ipins = ""
portausb = ""

opcoes = [
        "Ricoh 3710SF",
        "Ricoh 3510SF",
        "Epson L3150"
        ]

opcoes_c = [
        "Rede",
        "USB",
        "Atualizar Firmware"
]

print ("-------- PrAutIn --------")


escolha = easygui.choicebox("Selecione a impressora:\n\nLembre-se que o Spooler e Explorer será reiniciado.", "PrAutIn", opcoes)

if escolha:


    coneccao = easygui.choicebox("Selecione o tipo de conexão:", "PrAutIn", opcoes_c)

    #insr.restart_spooler()
    #insr.restart_explorer()

    if escolha == opcoes[0]:
        insr.ricoh3710()
    elif escolha == opcoes[1]:
        insr.ricoh3510()
    elif escolha == opcoes[2]:
        insr.epson3150()

#if ctypes.windll.shell32.IsUserAnAdmin():
#else:
#    easygui.msgbox("EXECUTE COMO ADMINISTRADOR!\nWindows não permite instalar sem admin.","ERRO ADMIN")