
"""
from janela import App

app = App()
app.mainloop()

"""
import sensiveis as senha
import condominios

# linha = ['2938', 'condominios@wmartins.com.br', '35433543', 'ABRJ ADMINISTRADORA DE BENS', 'CAMINO DEL SOL', '201', '']

# Sem boletos para teste (Sem boleto Ok)
# linha = ['2152', '211340090', '0339', 'BCF', 'MIRANTE CAMPESTRE', '', '']




getattr(condominios, senha.retornaradministradora('nomereal', linha[condominios.Administradora], 'nomereduzido').lower())(linha)
print(linha[6])
