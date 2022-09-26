
"""
from janela import App

app = App()
app.mainloop()

"""
import sensiveis as senha
import condominios
import auxiliares as aux
import shutil


try:
    shutil.rmtree(aux.caminhoprojeto('Profile'))

    # linha = ['2071', 'condominios@wmartins.com.br', '#123456Wm', 'CIPA', 'JOLIA', 'APT 104 / BL B']

    linha = ['3127', 'condominios@wmartins.com.br', '#123456Wm', 'CIPA', 'BARRA FAMILY', 'BL 1 APT 603']
    getattr(condominios, senha.retornaradministradora('nomereal', linha[condominios.Administradora], 'nomereduzido').lower())(linha)

    print(linha[6])
    #texto = "Teste/teste1"
    #ap, bloco = texto.split('/')
    #print(ap)
    #print(bloco)
except OSError as e:
    print(f"Error:{ e.strerror}")



# linha = ['2938', 'condominios@wmartins.com.br', '35433543', 'ABRJ ADMINISTRADORA DE BENS', 'CAMINO DEL SOL', '201', '']

# linha = ['1095', 'JOSE AROLDO', '010203', 'APSA', 'SOCIEDADE DOS AMIGOS DO DREAM VILLAGE', '', '']

# linha = ['27842784', '210100015', '5439', 'BAP', 'BARRA ZEN', '', '']

# linha = ['27842784', '210100015', '5439', 'BAP', 'BARRA ZEN', '', '']

# Sem boletos para teste (Sem boleto Ok)
# linha = ['2152', '211340090', '0339', 'BCF', 'MIRANTE CAMPESTRE', '', '']


