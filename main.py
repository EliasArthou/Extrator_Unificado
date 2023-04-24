
"""
from janela import App

app = App()
app.mainloop()

"""
import sensiveis as senha
import condominios
import auxiliares as aux
import shutil
from twilio.rest import Client


client = Client(senha.account_sid, senha.auth_token)
try:
    aux.caminho = aux.caminhoselecionado(3, 'Selecione o caminho de destino dos boletos:')

    tabela = []
    shutil.rmtree(aux.caminhoprojeto('Profile'))

    if len(aux.caminho) > 0:
        # ICONDO
        # linha =['42144214', '', '', 'PROHOME ADM E CONSULTORIA IMÓVEIS', 'GLOBAL 7000 OFFICES', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['27252725', '205800111', '5049', 'CENTRIMÓVEIS', 'RESIDENCIAL AROAZES', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['43604360', '', '', 'PRED RIO ADMINISTRAÇÃO IMOBILIÁRIA', 'VILA DE ESPANA', '704', '', '', False, '']
        # tabela.append(linha)
        # linha =['20362036', '2272800020', '6639', 'PROTEST', 'SOLAR DOS COQUEIROS', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['25892589', '06914307', '3082', 'ESTASA', 'RESIDENCIAL FACILE', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['27132713', '', '', 'PROMENADE', 'LINK OFFICE MALL & STAY-OFFICE MALL', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['27512751', 'Est096602506', '5819', 'ESTASA', 'OCEAN DRIVE', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['42114211', '2278900289', '7289', 'PROTEST', 'RESIDENCIAL AQUAGREEN', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['27212721', '2004700061', '7939', 'PACIFICA ADMINSTRADORA DE IMOVEIS', 'MARAMBAIA', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['43514351', '2013100246', '4035', 'PACIFICA ADMINSTRADORA DE IMOVEIS', 'RESIDENCIAL MINHA PRAIA', '809', '', '', False, '']
        # tabela.append(linha)
        # linha =['40934093', '2000800072', '1672', 'JC CONSULTORIA', 'COND.E. CASA DAS VARANDAS', '001.869.607-43', '', '', False, '']
        # tabela.append(linha)
        # linha =['22162216', '07741002', '2023', 'ESTASA', 'BARRA QUALITY I', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['36903690', '6850596', '2175', 'ESTASA', 'CORES DA LAPA', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['38923892', '2409300046', '9366', 'PROTEST', 'SPAZIO REDENTORE', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['34103410', '2700100059', '6021', 'PROTEST', 'STELLA VITA', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['31643164', '', '', 'NOTA 10 X', 'CONDOMINIO RESIDENCIAL GALEOES', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['39773977', '', '', 'PROHOME ADM E CONSULTORIA IMÓVEIS', 'GLOBAL 7000 OFFICES', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['33053305', '2004700209', '4867', 'PACIFICA ADMINSTRADORA DE IMOVEIS', 'COND. MARAMBAIA - BARRA SUL', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['44424442', '', '', 'ZIRTAEB', 'BARRA BUSINESS', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['44434443', '', '', 'ZIRTAEB', 'BARRA BUSINESS', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['44414441', '', '', 'ZIRTAEB', 'BARRA BUSINESS', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['42744274', '', '', 'PERSONAL PRIME ADMINISTRADORA', 'PORTAL SHOPPING VARGEM GRANDE', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['40124012', '', '', 'PROMENADE', 'VIA BARRA', '1701 / 01', '', '', False, '']
        # tabela.append(linha)
        # linha =['39753975', '096503101', '9716', 'ESTASA', 'ED. HEAVEN', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['38883888', '2278900233', '5701', 'PROTEST', 'RESIDENCIAL AQUAGREEN', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['43764376', '', '', 'ZIRTAEB', 'CONDOMÍNIO DO EDIFÍCIO JULIO LIMA', '1408', '', '', False, '']
        # tabela.append(linha)
        # linha =['40924092', '2668400012', '9820', 'ZIRTAEB', 'PARQUE CHICO MENDES', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['38933893', '', '', 'PROMENADE', 'LINK OFFICE & MALL', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['43354335', '2279900129', '7679', 'PROTEST', 'INSIGHT OFFICE', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['39523952', '', '', 'PROHOME ADM E CONSULTORIA IMÓVEIS', 'GLOBAL 7000 OFFICES', '', '', '', False, '']
        # tabela.append(linha)
        # linha =['43814381', '2274600020', '3795', 'PROTEST', 'SAINT GERMAIN', '304', '', '', False, '']
        # tabela.append(linha)
        # linha =['43044304', 'CONDOMINIOS@WMARTINS.COM.BR', '442905', 'PROHOME ADM E CONSULTORIA IMÓVEIS', 'MERITO', '', '', '', False, '']
        # tabela.append(linha)

        # WEBWARE

        linha = ['40014001', '2822wm', '2822wmartins', 'LOWNDES', 'LIBERTA RESORT', '', '', '', False, '']
        tabela.append(linha)
        linha = ['39953995', '', '', 'CONAC ADMINISTRADORA', 'RIO RESIDENCIAL', '', '', '', False, '']
        tabela.append(linha)
        linha = ['42324232', 'RIO MALL LOJA E', '35603560', 'CONAC ADMINISTRADORA', 'AMERICA CONDOMINIO CLUBE', '', '', '', False, '']
        tabela.append(linha)
        linha = ['24772477', '', '', 'CONAC ADMINISTRADORA', 'BRITANIA', 'APT 203', '', '', False, '']
        tabela.append(linha)
        linha = ['24582458', 'BL09APTO609', 'APTO609', 'CONAC ADMINISTRADORA', 'SUBLIME MAX CONDOMINIUM', '', '', '', False, '']
        tabela.append(linha)
        linha = ['39243924', 'rio residencial', '010203', 'CONAC ADMINISTRADORA', 'RIO RESIDENCIAL', '', '', '', False, '']
        tabela.append(linha)
        linha = ['30283028', 'APTO.410', 'BL01APTO410', 'LOWNDES', 'COND. ED. HIGH RESIDENCE', '', '', '', False, '']
        tabela.append(linha)
        linha = ['36883688', 'america condominio', '010203', 'CONAC ADMINISTRADORA', 'AMERICA CONDOMINIO CLUB', '', '', '', False, '']
        tabela.append(linha)
        linha = ['30753075', 'APTO.409', 'APTO409', 'CONAC ADMINISTRADORA', 'PALAZZO LARANJEIRAS', '', '', '', False, '']
        tabela.append(linha)
        linha = ['36723672', '', '', 'CONAC ADMINISTRADORA', 'NOBRE NORTE CLUBE RESIDENCIAL', '', '', '', False, '']
        tabela.append(linha)
        linha = ['28892889', 'APTO.501', 'BL01APTO501', 'LOWNDES', 'LIBERTA RESORT', '', '', '', False, '']
        tabela.append(linha)
        linha = ['31393139', 'APTO.204', 'APTO204', 'CONAC ADMINISTRADORA', 'PALAZZO LARANJEIRAS', '', '', '', False, '']
        tabela.append(linha)
        linha = ['37403740', 'rio_residencial01', '010203', 'CONAC ADMINISTRADORA', 'RIO RESIDENCIAL', '', '', '', False, '']
        tabela.append(linha)
        linha = ['42824282', 'ROSSIPARQUE', '010203', 'CONAC ADMINISTRADORA', 'ROSSI PARQUE DAS ROSAS', '', '', '', False, '']
        tabela.append(linha)
        linha = ['37073707', '', '', 'CONAC ADMINISTRADORA', 'AMERICA CONDOMINIO CLUBE', '', '', '', False, '']
        tabela.append(linha)
        linha = ['34143414', 'america cond', '010203', 'CONAC ADMINISTRADORA', 'AMERICA CONDOMINIO CLUBE', '', '', '', False, '']
        tabela.append(linha)
        linha = ['43784378', '', '', 'CONAC ADMINISTRADORA', 'ONDA CARIOCA CONDOMINIUM CLUB', '303 BL 04', '', '', False, '']
        tabela.append(linha)
        linha = ['34333433', 'PROKO I', '34333433', 'CONAC ADMINISTRADORA', 'PROKO I', '', '', '', False, '']
        tabela.append(linha)
        linha = ['34343434', 'STA.TEREZINHA', 'APTO902', 'CONAC ADMINISTRADORA', 'CONDOMINIO SANTA TEREZINHA', '', '', '', False, '']
        tabela.append(linha)
        linha = ['40954095', '', '', 'CONAC ADMINISTRADORA', 'PALAZZO LARANJEIRAS', '', '', '', False, '']
        tabela.append(linha)
        linha = ['36273627', 'rio_residencial02', '010203', 'CONAC ADMINISTRADORA', 'RIO RESIDENCIAL', '', '', '', False, '']
        tabela.append(linha)
        linha = ['40234023', '', '', 'CONAC ADMINISTRADORA', 'RIO RESIDENCIAL', '', '', '', False, '']
        tabela.append(linha)
        linha = ['37643764', '', '', 'CONAC ADMINISTRADORA', 'RIO RESIDENCIAL', '', '', '', False, '']
        tabela.append(linha)
        linha = ['38753875', 'verano.2919', '29192919', 'CONAC ADMINISTRADORA', 'VERANO', '', '', '', False, '']
        tabela.append(linha)
        linha = ['60896089', 'interativa@interativaimoveis.com.br', '383801', 'LOWNDES', 'PORTO ROMAZZINO', '', '', '', False, '']
        tabela.append(linha)
        linha = ['43304330', 'qi75qa5119', 'ne25hi', 'LOWNDES', 'LIBERTA RESORT', '', '', '', False, '']
        tabela.append(linha)
        linha = ['39923992', '', '', 'CONAC ADMINISTRADORA', 'RIO RESIDENCIAL', '', '', '', False, '']
        tabela.append(linha)
        linha = ['41024102', 'LIBERTA RESORT', '34413441', 'LOWNDES', 'LIBERTÁ RESORT', '', '', '', False, '']
        tabela.append(linha)
        linha = ['42334233', 'Renata Izidoro', '', 'CONAC ADMINISTRADORA', 'VERANO', '3118-0374', '', '', False, '']
        tabela.append(linha)
        linha = ['43924392', '', '', 'CONAC ADMINISTRADORA', 'VILA DO LARGO', 'CASA 28 APTO 102', '', '', False, '']
        tabela.append(linha)

        for item in tabela:
            if len(item[condominios.Usuario].strip()) > 0 and len(item[condominios.Senha].strip()) > 0:
                multiplas = aux.encontrar_administradora(item[condominios.Administradora])
                if multiplas is None:
                    getattr(condominios, senha.retornaradministradora('nomereal', item[condominios.Administradora], 'nomereduzido').lower())(item)
                else:
                    getattr(condominios, multiplas['Site'])(item)
            else:
                item[condominios.Resposta] = 'Usuário e/ou senha não preenchido!'
            print('Resposta: ' + item[condominios.Resposta])

except OSError as e:
    print(f"Error:{e.strerror}")
    # client.messages.create(body='A função minha_funcao falhou: {}'.format(str(e)), from_=senha.twilio_number, to=senha.my_number)



# linha = ['2938', 'condominios@wmartins.com.br', '35433543', 'ABRJ ADMINISTRADORA DE BENS', 'CAMINO DEL SOL', '201', '']

# linha = ['1095', 'JOSE AROLDO', '010203', 'APSA', 'SOCIEDADE DOS AMIGOS DO DREAM VILLAGE', '', '']

# linha = ['27842784', '210100015', '5439', 'BAP', 'BARRA ZEN', '', '']

# linha = ['27842784', '210100015', '5439', 'BAP', 'BARRA ZEN', '', '']

# Sem boletos para teste (Sem boleto Ok)
# linha = ['2152', '211340090', '0339', 'BCF', 'MIRANTE CAMPESTRE', '', '']


