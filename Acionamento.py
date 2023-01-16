import os
import time
import pandas as pd
import win32clipboard


# chrome_options = selenium.webdriver.chrome.options.Options()
# chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9223")

# driver = webdriver.Chrome(chrome_driver, chrome_options = chrome_options)



def testeretorno( dcheckretorno ):
    portas = dcheckretorno.drop_duplicates( subset = 6 )
    print( 'dentro da funcão',portas )

    pass


### versão 2, corrigido a limpeza dos contratos e parte da caracterização

while True:
    try:
        print( 'Ferramenta desenvolvida por Daniel Ledezma Vieira\n\n' )
        ctts = input( 'Cole a linha de contratos da URA:' )  # Receber Control+C Control+V com os contratos
        ctts = ctts.split( ',' or '.' )  # Realizar a separação dos contratos e abaixo a limpeza

        for n,item in enumerate( ctts ):
            while ":" in item:
                item = item[ 0:-1 ]
            item = item.lstrip( ' ' )
            item = item.rstrip( ' ' )
            ctts[ n ] = item

        if len( ctts[ 0 ] ) > 9:
            cod = ctts[ 0 ]
            cod = cod[ 0:3 ]
            print( cod,'É o código da cidade.' )
            for i,limpinho in enumerate( ctts ):
                if limpinho[ 0:3 ] == cod:
                    limpinho = limpinho[ 3:: ]
                limpinho = limpinho.lstrip( ' ' )
                limpinho = limpinho.rstrip( ' ' )
                limpinho = limpinho.lstrip( '0' )
                limpinho = limpinho.rstrip( '.' )
                ctts[ i ] = limpinho
        else:
            print( 'Segue o Baile' )

        print( 'Contrato(s) limpos(s):',ctts,'\n\nVerificando contratos no arquivo:' )

        # OK PEGAR O CAMINHO FUNCÇÃO
        pastadownloads = os.environ[
            'USERPROFILE' ]  # USERPROFILE mostra a pasta do usuario, podendo facilmente localizad a pasta downloads
        ### concatenar forma não garantida, é melhor usando pathjoin = ### sistema = sistema + '/Downloads'
        pastadownloads = os.path.join( pastadownloads,'Downloads' )
        # print( os.listdir( pastadownloads ) )
        # print(os.getcwd()) #localizar onde se encontra
        listaxlsrecente = [ ]
        listaarquivospastadownloads = os.listdir( pastadownloads )
        for arquivo in listaarquivospastadownloads:
            if '.xls' in arquivo:
                dataarquivo = os.path.getmtime( f'{pastadownloads}/{arquivo}' )
                listaxlsrecente.append( (dataarquivo,arquivo) )
        listaxlsrecente.sort( reverse = True )
        xlsrecente = listaxlsrecente[ 0 ]
        # print('O arquivo .XLS mais recente é: ',xlsrecente[1])
        # print('\nOs demais arquivos .XLS encontrado: ',listaxlsrecente)
        caminho = os.path.join( pastadownloads,xlsrecente[ 1 ] )
        print( caminho )
        time.sleep( 1 )

        planilha = pd.read_html( caminho )  # encoding='ISO-8859-1'
        df = pd.DataFrame( planilha[ 0 ],columns = [ 6,8,24,25,28,29,30,31,32,33,34,35,36,37,38 ] )

        if ctts != [ '' ]:
            reclamantes = [ ]
            for z,cttlista in enumerate( ctts ):
                valx = df.loc[ df[ 25 ] == cttlista ]
                reclamantes.append( valx )
            reclamantes = pd.concat( reclamantes )
            reclamantes = reclamantes.drop_duplicates( subset = 25 )
            print( '\nRECLAMANTES >>>:\n',
                   reclamantes[ [ 8,28,29,30,31,32,33,34,35,36,37,38 ] ].dropna( axis = 'columns',how = 'all' ).fillna(
                       ' ' ).to_string( index = False ) )
            # print(resultado[28].iloc[1]) #exibe a coluna especifica e iloc linha
            reclamantes.insert( 4,26,'-',allow_duplicates = False )
            reclamantes.insert( 5,'n','n°',allow_duplicates = False )
            reclamantes.insert( 14,'bairro','- Bairro:',allow_duplicates = False )

            checkretorno = reclamantes
            checkretorno = checkretorno.drop_duplicates( subset = 8 )
            resultado = [ ]
            portas = [ ]
            ruas = [ ]

            testeretorno( checkretorno )

            if len( checkretorno ) == 1:
                print( '\nReclamantes pertencem ao mesmo retorno:',checkretorno[ 8 ].iloc[ 0 ] )
                checkretorno = checkretorno[ 8 ].iloc[ 0 ]
                resultado = df.groupby( 8 ).get_group( checkretorno )
                testeretorno( resultado )
                portas = resultado.drop_duplicates( subset = 6 )
                portas = portas[ 6 ].tolist()
                ruas = resultado.drop_duplicates( subset = 28 )
                ruas = ruas.dropna( axis = 0,thresh = 7 )
                ruas = ruas.drop_duplicates( subset = 28 ).dropna( axis = 'columns',how = 'all' ).fillna( '' )
                ruas = ruas[ 28 ].to_string( index = False,header = False )
                retorno = resultado[ 8 ].iloc[ 0 ]

            else:
                checkretorno = checkretorno[ 8 ].tolist()
                print( '\nReclamantes em retornos diferentes:',checkretorno )
                checkretorno = checkretorno[ 0 ]
                resultado = df.groupby( 8 ).get_group( checkretorno )
                portas = resultado.drop_duplicates( subset = 6 )
                portas = portas[ 6 ].tolist()
                # for i,retornos in enumerate(checkretorno):
                # valx=df.loc[df[8]==retornos]
                # resultado.append(valx)
                # resultado=pd.concat(resultado)

                # resultado = resultado.drop_duplicates( subset=25 )
        else:
            print( 'Não foi inserido contratos, puxando node/retorno do arquivo12' )
            noderetorno = df = pd.DataFrame( planilha[ 0 ],columns = [ 6,8,28 ] )
            rretornos = noderetorno.drop_duplicates( subset = 8 ).dropna( axis = 'columns',how = 'all' ).fillna( '' )
            rretornos = rretornos[ 8 ].tolist()
            for n,item in enumerate( rretornos ):
                if n != 0:
                    print( 'eis os retornos disponiveis:',item )

        print(
            "\n\nOpção 1: Sem Sinal TOTAL\nOpção 2: Sem sinal PARCIAL NO RETORNO\nOpção 3: Outage Degradação de FEC SNR\nOpção 4: SATURAÇÃO DEGRADAÇÃO DE RX TX MER" )
        otgsintoma = input( '\nDigite o Sintoma:' )

        inputretorno = ''

        ruido = ''

        if otgsintoma == '1':
            otgsintoma = "   ### OUTAGE SEM SINAL TOTAL ###"
            retorno = 'TOTAL'
            ruas = 'TOTAL'

        elif otgsintoma == '2':
            otgsintoma = '### OUTAGE SEM SINAL PARCIAL NO RETORNO'
            inputretorno = input( 'Digite o RETORNO, ou ENTER para retorno único:' )
            if inputretorno == '':
                inputretorno = 'ÚNICO ###'
                retorno = 'ÚNICO'
            else:
                inputretorno = inputretorno.upper()
                retorno = inputretorno
                inputretorno = inputretorno + ' ###'

        elif otgsintoma == '3':
            otgsintoma = '### OUTAGE DEGRADAÇÃO DE FEC SNR ###'
            ruido = '''

SNR MÉDIO: 35DB
SNR ATUAL: 25B
FEC CORRIGIDO: 80.00%
FEC NÃO CORRIGIDO: 5.0%'''
        elif otgsintoma == '4':
            otgsintoma = '### OUTAGE SATURAÇÃO DEGRADAÇÃO DE RX TX MER ###'

        print( "Opção 1: URA MP\nOpção 2: URA \nOpção 3: SCAN\nOpção 4: Cliente BSOD\nOpção 5: IE" )
        otgreclamante = input( 'Digite o Sintoma:' )

        if otgreclamante == '1':
            otgreclamante = "#OUTAGE ABERTO POR RECLAMAÇÃO URA MP#"
        elif otgreclamante == '2':
            otgreclamante = '# OUTAGE ABERTO POR RECLAMAÇÃO URA #'
        elif otgreclamante == '3':
            otgreclamante = '## OUTAGE ABERTO POR ALARME SCAN ##'
        elif otgreclamante == '4':
            otgreclamante = '## OUTAGE ABERTO POR CLIENTE BSOD ##'
        elif otgreclamante == '5':
            otgreclamante = '# OUTAGE ABERTO POR RECLAMAÇÃO IE #'

        os.system( 'cls' )

        # deixar somente o node
        nodefinal = resultado[ 8 ].iloc[ 0 ]
        while "-" in nodefinal or "#" in nodefinal:
            nodefinal = nodefinal[ 0:-1 ]

        caracterizacao = (f'''{otgreclamante} 
{otgsintoma} {inputretorno}

NODE: {nodefinal}
RETORNO: {retorno}

CMTS: 
PORTAS: {portas[ 0 ]} AO {portas[ -1 ]} {ruido}

RECLAMANTES:
{reclamantes[ [ 25,26,28,'n',29,30,31,32,33,34,35,36,37,'bairro',38 ] ].dropna( axis = 'columns',how = 'all' ).fillna( ' ' ).to_string( index = False,header = False )}

TRECHO AFETADO:
{ruas}

ATLAS: 
PORTAL: 

RECORRÊNCIA NOS ULTIMOS 30 DIAS: 0 OUTAGE(S)
''')

        print( caracterizacao )
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText( caracterizacao )
        win32clipboard.CloseClipboard()

    except:
        print( '\nREINICIANDO\n\n' )
        break