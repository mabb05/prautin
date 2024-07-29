import os,ctypes,easygui
import INS_ricoh as insr

ipins = ""

opcoes = [
        "Ricoh 3710SF",
        "Ricoh 3510SF"
        ]

opcoes_c = [
        "Rede",
        "USB"
]

print ("-------- PrAutIn --------")

escolha = easygui.choicebox("Selecione a impressora:\n\nLembre-se que o Spooler e Explorer será reiniciado.", "PrAutIn", opcoes)

if escolha:
    coneccao = easygui.choicebox("Selecione o tipo de conexão:", "PrAutIn", opcoes_c)
    if coneccao:
        ipins = easygui.enterbox("Digite o endereço IP para instalar:", "Endereço IP", "192.168.50.203")
        nomeimp = easygui.enterbox("Digite o nome da impressora para instalar:", "Nome da impressora", str(escolha))
        lines = [ipins, '\n', nomeimp]
        with open("param.txt", "w") as file: file.writelines(lines)

        # REINICIAR SERVIÇOS
        # insr.restart_spooler()
        # insr.restart_explorer()

        if escolha == opcoes[0]:
            insr.ricoh3710()
        elif escolha == opcoes[1]:
            insr.ricoh3510()

# if ctypes.windll.shell32.IsUserAnAdmin():
# else:
#    easygui.msgbox("EXECUTE COMO ADMINISTRADOR!\nWindows não permite instalar sem admin.","ERRO ADMIN")