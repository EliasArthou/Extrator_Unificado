
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
    linha = ['1095', 'JOSE AROLDO', '010203', 'APSA', 'META OFFICE BUILDING I', '', '']
    linha = ['1095', 'JOSE AROLDO', '010203', 'APSA', 'SOCIEDADE DOS AMIGOS DO DREAM VILLAGE', '', '']

    getattr(condominios, senha.retornaradministradora('nomereal', linha[condominios.Administradora], 'nomereduzido').lower())(linha)
    print(linha[6])
except OSError as e:
    print(f"Error:{ e.strerror}")



# linha = ['2938', 'condominios@wmartins.com.br', '35433543', 'ABRJ ADMINISTRADORA DE BENS', 'CAMINO DEL SOL', '201', '']

# Sem boletos para teste (Sem boleto Ok)
# linha = ['2152', '211340090', '0339', 'BCF', 'MIRANTE CAMPESTRE', '', '']


