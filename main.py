"""
from janela import App

app = App()
app.mainloop()

"""
import sensiveis as senha
import condominios
import auxiliares as aux

import os
import messagebox

# import shutil


try:
    aux.caminho = aux.caminhoselecionado(3, 'Selecione o caminho de destino dos boletos:')

    if os.path.isfile(aux.caminhoprojeto() + '/' + 'Scai.WMB'):
        caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                              tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')],
                                              caminhoini=aux.caminhoprojeto(), arquivoinicial='Scai.WMB')
    else:
        if os.path.isdir(aux.caminhoprojeto()):
            caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                                  tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')],
                                                  caminhoini=aux.caminhoprojeto())
        else:
            caminhobanco = aux.caminhoselecionado(titulojanela='Selecione o arquivo de banco de dados:',
                                                  tipoarquivos=[('Banco ' + senha.empresa, '*.WMB'), ('Todos os Arquivos:', '*.*')])

    tabela = []
    # shutil.rmtree(aux.caminhoprojeto('Profile'))

    if len(aux.caminho) > 0 and len(caminhobanco) > 0:
        bd = aux.Banco(caminhobanco)
        tabela = bd.consultar(aux.retornarlistaboletos(['HFLEX', 'VÓRTEX']))

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
    else:
        if len(aux.caminho) > 0:
            messagebox.msgbox('Escolha o caminho para salvar os boletos!', messagebox.MB_OK, 'Caminho não definido!')

        if len(caminhobanco) > 0:
            messagebox.msgbox('Arquivo do banco não selecionado!', messagebox.MB_OK, 'Sem Banco!')

except Exception as e:
    print(f"Error:{str(e)}")
