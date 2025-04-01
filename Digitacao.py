from subprocess import Popen
from pywinauto import Application
import RepositorioDAO
import pyautogui
import FuncoesAuxiliares
import pyperclip
import os
import cv2
import time
from decimal import Decimal
import Calculos


class Digitacao:
   
    
    def acertoOperador(self,dta):
        pyautogui.PAUSE=0
        
        varCalculos = Calculos.Operacoes()
        varExecuteDAO = RepositorioDAO.DAO()
        varQueryLojas ="select nroempresa, empresa from CONSINCO.dim_empresa where nroempresa not in (99, 800, 986, 987, 989, 999, 10, 13, 20, 25, 29)  order by NROEMPRESA"
        #varQueryLojas ="select nroempresa, empresa from CONSINCO.dim_empresa where nroempresa not in (99, 800, 986, 987, 989, 999, 13, 20, 25,29) AND NROEMPRESA > 18 order by NROEMPRESA"
       
        varFuncao = FuncoesAuxiliares.Funcao_Apoio()
        varJn_acerto_operador="C:\\Projetos_Python\\TESOURARIA\\img\\janelas\\jn_acerto_operador.png"
        Gt_inicial_Formulario =None
        Gt_Final_Formulario = Decimal(0.0)
        tx_field_gt_inicial =None
        tx_field_gt_final =None
        vlrOutros_Sistema = Decimal(0.0)
        vlrOutros_Tabela_String ="0"
        srtDiferenca= Decimal(0.0)

        #Sequencia de atalhos que abre a tela Acerto de Operador.
        pyautogui.press("ctrl")
        pyautogui.press("ctrl")
        pyautogui.sleep(1.5)
        pyautogui.press("alt")
        pyautogui.sleep(1.5)
        pyautogui.press("m")
        pyautogui.sleep(1.5)
        pyautogui.press("enter")
                                                                                                                            
        varFuncao.aguardar_janela_por_imagem(varJn_acerto_operador, "Acerto de Operador")
                         
        varLojas=varExecuteDAO.executaQuery(varQueryLojas)
        #LOOP DAS LOJAS
        for row in varLojas:
            lj=row[0]
            strLoja = str(lj)
            nm_lj = row[1]

            print("Mudando Loja")

            if strLoja == "23":
               print("pare")

            #validacaoxlsx =f'C:/Projetos_Python/TESOURARIA/arquivos/download/movimentolj{strLoja}.xlsx'
            validacaoxlsx = f"\\\\10.11.10.3\\arcomixfs$\\Financeiro\\digitacao\\movimentolj{strLoja}.xlsx"
            validacaoLojas =f"C:\\Projetos_Python\\TESOURARIA\\arquivos\\execucao\\movimentolj{strLoja}.csv"
            print("Testou o arquivo csv")           


            if os.path.exists(validacaoxlsx): #Testa para ver se o arquivo existe.
               print("Testou se o arquivo xlsx existe antes de ler a data")
               resul_data =varFuncao.capturar_primeira_data(strLoja)
               print("Data Lida")

               if resul_data:
                  dta = resul_data

               if not os.path.exists(validacaoLojas):                                                                                             

                 # varFuncao.excluir_coluna_11(strLoja)
                  pyautogui.sleep(2)
                  varFuncao.xlsx_to_csv(strLoja) 

            """
            if os.path.exists(validacaoxlsx): #Testa para ver se o arquivo existe.
               if not os.path.exists(validacaoLojas):
                varFuncao.excluir_coluna_11(strLoja)
                pyautogui.sleep(2)
                varFuncao.xlsx_to_csv(strLoja)
            """ 
            pyautogui.sleep(1)
            
            #validacaoLojas =f"C:\\Projetos_Python\\TESOURARIA\\arquivos\\digitacao\\movimentolj{strLoja}.csv"
            if not os.path.exists(validacaoLojas): #Testa para ver se o arquivo existe.
               #varFuncao.GeraLogsInfo(f" Loja : {strLoja} não existe ")
               continue # Chamar outra loja, arquivo não existe.
               
            pyautogui.press("ctrl")  
            print("Efetuando a troca de loja")                                        
            pyautogui.hotkey("ctrl","shift","t") 
     
            varFuncao.AguardaAberturaJanela("Empresas")

            #Conectando em empresas administradoras
            appEmpre = Application().connect(title_re=".*Empresas Administradoras.*", class_name="Gupta:Dialog") 
            dlgEmpresas = appEmpre.window(class_name="Gupta:Dialog")

            dlgEmpresas['Edit0'].click_input() # Clicando na data do campo numeroLoja
            pyautogui.sleep(1)
            pyautogui.keyDown("ctrl") #Selecionando a informacao
            pyautogui.keyDown("a")
            pyautogui.keyDown("a")
            pyautogui.keyDown("ctrl")
            dlgEmpresas['Edit0'].type_keys("{DELETE}")
            dlgEmpresas['Edit0'].type_keys(strLoja)
            dlgEmpresas['Button0'].click_input() # Clicando em ok para na janelas empresas.

           # varFuncao.esperar_fechamento_janela("Empresas")

            #pyautogui.sleep(0.35)
            jnErro = varFuncao.check_window_exists("Atenção") # Se a consiliação do ultimo pdv da loja anterior falhar, esse bloco será chamado
            if jnErro == True:
               appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
               window = appDlgConf_Exclusao.Dialog
               window.Wait('ready')
               button = window[u'&Sim']
               button.click_input()
               varFuncao.esperar_fechamento_janela("Atenção")
          
            varFuncao.esperar_fechamento_janela("Empresas")
                                                                                                                
            varFuncao.aguardar_janela_por_imagem(varJn_acerto_operador, "Janela Acerto do Operador") 
            #time.sleep(2)
            
            #Conectando na janela acerto de operador para executar a pesquisa
            appAcertoOperador = Application().connect(title_re=".*Tesouraria.*")
            guptamdiframeAcer_Operador = appAcertoOperador.window(class_name='Gupta:MDIFrame') 
            guptamdiframeAcer_Operador.Wait('ready')
            guptamdiframeAcer_Operador[u'Edit1'].click_input() #Click na no campo data
            pyautogui.sleep(1)
            #Selecionando a informação que será apagada
            pyautogui.keyDown('ctrl')
            pyautogui.sleep(0.25)
            pyautogui.keyDown('a')
            pyautogui.sleep(0.25)
            pyautogui.keyUp('a')
            pyautogui.sleep(0.25)
            pyautogui.keyUp('ctrl')
            time.sleep(0.25)
           
           #Apagando a informação na caixa de data.
            guptamdiframeAcer_Operador[u'Edit1'].type_keys("{DELETE}")
            pyautogui.sleep(1)
            guptamdiframeAcer_Operador[u'Edit1'].type_keys(dta) # Informando a data
            pyautogui.sleep(1)
            guptamdiframeAcer_Operador[u'Button33'].click_input() #Clicando na lupa da janela  
           
            pyautogui.sleep(2.5)
                                                                                                
            varLoopBanco = 0 # Variavel que controla o retorno do loop

            #Capturando o movimento das lojas no banco de Dados:
            reult_movimento = varExecuteDAO.RetornaMovimento(strLoja,dta)
            for row_movimento_banco in reult_movimento: #Loop do banco de dados
               bd_nroempresa = row_movimento_banco[0]
               bd_data = row_movimento_banco[1]
               bd_turno = row_movimento_banco[2]
               bd_gtInicial = row_movimento_banco[3]
               bd_gtFinal = row_movimento_banco[4]
               bd_encargos = row_movimento_banco[5]
               bd_vlrCesat = row_movimento_banco[6]   
               bd_Total = row_movimento_banco[7]
               bd_vlrBanca = row_movimento_banco[8]
               bd_qtBanc = row_movimento_banco[9]
               nro_semUso = row_movimento_banco[10]
               bd_acertado = row_movimento_banco[11]
               bd_fechado = row_movimento_banco[12]
               bd_usuFechou = row_movimento_banco[13]
               bd_data_semUso = row_movimento_banco[14]
               bd_alteracao = row_movimento_banco[15]
               bd_versao = row_movimento_banco[16]
               bd_identificador = row_movimento_banco[17]
               aleatorio = row_movimento_banco[18]
               bd_nropdv = row_movimento_banco[19]
               softpdv = row_movimento_banco[20]
               bd_codOperador = row_movimento_banco[21]
               bd_nvl = row_movimento_banco[22]
               bd_usoxxxxxxxxx = row_movimento_banco[23]
               bd_dtadddddddd = row_movimento_banco[24]
               bd_nomeOperadora = row_movimento_banco[25]
                                                                                                                                   
               bd_nroPdv_String = str(bd_nropdv)# Convertendo o numero do pdv para string
               bd_codOperador_String = str(bd_codOperador)# Convertendo o codigo do operador para string
               bd_turno_string = str(bd_turno) # Convertendo o Turno
               vlrDinheiroBanco_Srt =str(bd_Total)
              # pyautogui.sleep(2)
               
              # mensagem =f"PREPARANDO A DIGITAÇÃO DO PDV : NRO PDV : {bd_nropdv} -- TURNO : {bd_turno} -- OPERADOR : {bd_nomeOperadora}"
               #varFuncao.show_popup(mensagem)

              # TESTA SE O PDV JÁ FOI CONCILIADO
               if bd_acertado == "S": # SE JA FOI CONCILIADO CHAMA OUTRO  CAIXA
                  
                  varFuncao.GeraLogsInfo(f"LOJA : {bd_nroempresa} --  PDV JÁ  CONCILIADO : NRO PDV : {bd_nropdv} -- TURNO : {bd_turno} -- OPERADOR : {bd_nomeOperadora}")
                                                                                                                                                     
                  guptamdiframeAcer_Operador[u'Button31'].click_input() #Chama o proximo pdv no formulario
                  
                  jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 
                  if jnErro == True:
                     appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
                     window = appDlgConf_Exclusao.Dialog
                     window.Wait('ready')
                     button = window[u'&Sim']
                     button.click_input()
         
                  pyautogui.sleep(0.3)
                  continue
               
              # Se o GTInicial for igual a zero chamar outro pdv 
               if bd_gtInicial == 0 or bd_gtFinal == 0:

                #  mensagem =f"GT DO PDV NÃO PODE SER ZERO : GT INICIAL : {bd_gtInicial} -- GT FINAL : {bd_gtFinal} NRO PDV : {bd_nropdv} -- TURNO : {bd_turno} -- OPERADOR : {bd_nomeOperadora}"
                #  varFuncao.show_popup(mensagem)

                  guptamdiframeAcer_Operador[u'Button31'].click_input() #Chama o proximo pdv no formulario
                  pyautogui.sleep(0.25)

                  jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 
                  if jnErro == True:
                     appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
                     window = appDlgConf_Exclusao.Dialog
                     window.Wait('ready')
                     button = window[u'&Sim']
                     button.click_input()
                  pyautogui.sleep(1)
                  continue
                  
               else:
                    if bd_gtInicial == bd_gtFinal:
                      varFuncao.GeraLogsInfo(f"LOJA : {bd_nroempresa} GT INICIAL E FINAL SÃO IGUAIS - CONCILIANDO PDV !!! : GT INICIAL : {bd_gtInicial} -- GT FINAL : {bd_gtFinal} NRO PDV : {bd_nropdv} -- TURNO : {bd_turno} -- OPERADOR : {bd_nomeOperadora}")
        
                      guptamdiframeAcer_Operador[u'Button40'].click_input() # Clicando em Conciliar
                      pyautogui.sleep(2)
                      jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 
                      if jnErro == True:
                          appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
                          window = appDlgConf_Exclusao.Dialog
                          window.Wait('ready')
                          button = window[u'Button']
                          button.click_input()
                          pyautogui.sleep(0.25)

                      guptamdiframeAcer_Operador[u'Button31'].click_input() #Chama o proximo pdv no formulario
                      jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 
                      if jnErro == True:
                         appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
                         window = appDlgConf_Exclusao.Dialog
                         window.Wait('ready')
                         button = window[u'&Sim']
                         button.click_input()
                           
                      pyautogui.sleep(2)
                      continue
                     
              ##############################################################  
               varFuncao.GeraSeqTurnoCSV(strLoja,dta) 
                        
               print("Abrindo arquivo csv das lojas")
               varArquivosLojas =f"C:\\Projetos_Python\\TESOURARIA\\arquivos\\execucao\\movimentolj{strLoja}.csv"
               if os.path.exists(varArquivosLojas): #Testa para ver se o arquivo existe.
            
                  with open(varArquivosLojas, 'r') as c:
                     lines = c.readlines()
                     
                  indices_para_remover = set()

                  #LOOP DA PLANILHA
                  for index, line in enumerate (lines[1:], start=1):
                      line_data = line.strip().split(";")
                      
                   
                      codigo=line_data[0]
                      usuario=line_data[1]
                      pdv=line_data[2]
                      data=line_data[3]
                      dinheiro=line_data[4].strip()
                      devolucao=line_data[5]
                      sobra=line_data[6]
                      quebra=line_data[7]
                      loja=line_data[8]
                      coo=line_data[9]
                      key=line_data[10]
                                     
                      turnoCsv=str(line_data[11])
                    
                        

                                           
                      
                      #Se a linha referente ao código estiver  vazia, passar para a próxima linha   
                      if codigo == "":
                           
                           if index == len(lines) -1: # Se for a ultima interação , pdv do formulario e pode chamar o continue mesmo porque o fluxo seguirá normalmente.
                               
                               guptamdiframeAcer_Operador[u'Button31'].click_input()
                               varFuncao.GeraLogsInfo(f"LOJA : {bd_nroempresa} PDV NÃO FOI ENCONTRADO --  OPERADORA : {bd_nomeOperadora} -- PDV BANCO:{bd_nropdv} -- TURNO : {bd_turno}")

                               continue
                           else:
                               continue

                      
                      try:
                       int_pdv = int(pdv.strip()) 
                      except:
                          guptamdiframeAcer_Operador[u'Button31'].click_input() #Chama o proximo pdv no formulario
                          jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 
                          if jnErro == True:
                              appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
                              window = appDlgConf_Exclusao.Dialog
                              window.Wait('ready')
                              button = window[u'&Sim']
                              button.click_input()

                          varFuncao.GeraLogsInfo(f"LOJA : {bd_nroempresa} NUMERO DO PDV INVÁLIDO : {bd_nomeOperadora} - PDV BANCO:{bd_nropdv} - TURNO : {bd_turno}")
                          continue   

                                                       
                      if bd_nroPdv_String.strip() == pdv.strip(): # Primeira faze da validação dos pdv´s, nesta fase está testando se os pdv´s são iguais
                      
                        if bd_codOperador_String.strip() == codigo.strip() and bd_turno_string.strip() == turnoCsv.strip(): # Testa se o código do operador e o turno de trabalho são os mesmos
                           
                           tx_field_gt_inicial = guptamdiframeAcer_Operador.Edit3 
                           Gt_inicial_Formulario = tx_field_gt_inicial.window_text()
                           pyautogui.sleep(1)
                           #tx_field_gt_inicial.SetFocus()
                           #tx_field_gt_inicial.type_keys("^c") # Executa a copia
                           #Gt_inicial_Formulario = pyperclip.paste()
                           gt_inicio_decimal = varCalculos.convertStringPadraoUS(Gt_inicial_Formulario)

                           fl_gtInicial_csv = float(gt_inicio_decimal)
                           fl_gtInicial_bd_dados = float(bd_gtInicial)                                                 
                           #(Formulário)
                           if fl_gtInicial_csv != fl_gtInicial_bd_dados:#(Bando de Dados): #Verifica se o GT do Formulario que está sendo digitado, é igual ao do banco de dados Se não for, aborta a digitação
                                                             
                                varFuncao.GeraLogsInfo(f"LOJA : {bd_nroempresa} FALHA NA VALIDAÇÃO DO GT_INICAL : GT NO BANCO : {bd_gtInicial} -- GT NO FORMULÁRIO : {fl_gtInicial_csv} --OPERADOR BANCO : {bd_nomeOperadora} -- PDV BANCO:{bd_nropdv} -- TURNO : {bd_turno}")
                                guptamdiframeAcer_Operador[u'Button31'].click_input() #Chama o proximo pdv no formulario

                                jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 
                                if jnErro == True:
                                  appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
                                  window = appDlgConf_Exclusao.Dialog
                                  window.Wait('ready')
                                  button = window[u'&Sim']
                                  button.click_input()
                                  # Duas linhas comentadas abaixo não serão necessárias pois o break após o if irá retornar a execução para o loop do movimento do banco de dados.
                                  #pyautogui.sleep(3)
                                  #continue
                                break # CHAMA O PROXIMO PDV NO BANCO
                           else:
                              #SE O GT INICIAL DO FORMULARIO ESTIVER BATENDO COM O GT INICAL DO BANCO, A EXECUÇÃO VEM PARA ESSA LINHA, O PDV PODERÁ SER DIGITADO NORMALMENTE
                            #  mensagem =f"DIGITAÇÃO NORMAL : {bd_gtInicial} -- GT NO FORMULÁRIO : {gt_inicio_decimal} --OPERADOR BANCO : {bd_nomeOperadora} -- PDV BANCO:{bd_nropdv}"
                            #  varFuncao.show_popup(mensagem)

                              #Captura no formulário acerto do operador o status que diz se o modulo precisa ser conciliado ou não
                              varControleDigitacao = None
                              varFuncao.GeraLogsInfo(f"LOJA : {bd_nroempresa} DIGITAÇÃO NORMAL DO GT_INICAL : GT NO BANCO : {bd_gtInicial} -- GT NO FORMULÁRIO : {fl_gtInicial_csv} --OPERADOR BANCO : {bd_nomeOperadora} -- PDV BANCO:{bd_nropdv} -- TURNO : {bd_turno}")
                              varFuncao.GeraLogsInfo(f"LOJA : {bd_nroempresa} CONTINUANDO DIGITAÇÃO DO GT_INICAL : GT NO BANCO : {bd_gtInicial} -- GT NO FORMULÁRIO : {fl_gtInicial_csv} --OPERADOR BANCO : {bd_nomeOperadora} -- PDV BANCO:{bd_nropdv} -- TURNO : {bd_turno}")
                              
                              #COMEÇANDO A DIGITAÇÃO DO DINHEIRO ###############################################################################################################
                              #Captura o valor do dinheiro do sistema para para o calculo            
                              
                              srtDinheiroSistema_01 = guptamdiframeAcer_Operador.Edit15
                              srtDinheiroSistema = srtDinheiroSistema_01.window_text()

                              if dinheiro != "" and dinheiro > "0" :
                                dinheiro = varCalculos.removeCifraoRetornaString(dinheiro) 
                                guptamdiframeAcer_Operador[u'&Dinheiro'].click_input() #  Clicando em no botão dinheiro 
                                #pyautogui.sleep(0.45)

                                varFuncao.AguardaAberturaJanela("Movimento")
                                                                                                                                                                                              
                                guptamdiframeAcer_dinheiro = Application().connect(title_re="Movimento")
                                guptadialog_dinheiro = guptamdiframeAcer_dinheiro[u'Gupta:Dialog']
                                guptadialog_dinheiro.Wait('ready')

                                guptadialog_dinheiro[u'Edit2'].type_keys(dinheiro) # Informando o dinheiro
                                pyautogui.sleep(0.25)
                              
                                guptadialog_dinheiro[u'Button1'].click_input() # clicando em ok

                                pyautogui.sleep(0.25)
                                jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção
                                if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Tesouraria.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'&OK']
                                    button.click_input()

                                    varFuncao.esperar_fechamento_janela("Atenção")
                                    
                                    pyautogui.sleep(0.25)
                                    guptadialog_dinheiro[u'Edit2'].click_input() #Clicando no campo dinheiro
                                    pyautogui.sleep(0.25)

                                    pyautogui.press("ctrl")   #Selecionando o conteudo do campo
                                    pyautogui.sleep(0.25)  
                                    pyautogui.keyDown("ctrl")
                                    pyautogui.sleep(0.25)  
                                    pyautogui.keyDown("a") 
                                    pyautogui.sleep(0.25) 
                                    pyautogui.keyUp("a") 
                                    pyautogui.sleep(0.25) 
                                    pyautogui.keyUp("ctrl") 

                                   #pyautogui.hotkey("ctrl","a") 

                                    guptadialog_dinheiro[u'Edit2'].type_keys("1,00") # Informando um valor aleatorio no dinheiro, não ficará salvo
                                    pyautogui.sleep(0.25)
                                    guptadialog_dinheiro[u'Button1'].click_input() # clicando em ok para fechar a caixa de dinheiro
                                    pyautogui.sleep(0.25)

                                    varFuncao.esperar_fechamento_janela("Movimento")

                                    guptamdiframeAcer_Operador[u'Button31'].click_input() #Chama o proximo pdv no formulario
                                    jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 
                                    if jnErro == True:
                                       appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
                                       window = appDlgConf_Exclusao.Dialog
                                       window.Wait('ready')
                                       button = window[u'&Sim']
                                       button.click_input()

                                       varFuncao.esperar_fechamento_janela("Atenção")

                                       break
                                    varFuncao.esperar_fechamento_janela("Movimento")

                              devolucao_float =0.0           
                              #COMEÇANDO A DIGITAÇÃO DE OUTROS ######################################################################################################################
                              if devolucao != "":
                                 try:
                                  devolucao_float = varCalculos.convertStringPadraoUS(devolucao) # Removendo o cifrão, se existir, e convertendo o valor da devoucao da planilha em float para a comparacao
                                 except:
                                     pyautogui.sleep(0.25)
                                                                                                                
                                     guptamdiframeAcer_Operador[u'Button31'].click_input() #Chama o proximo pdv no formulario
                                     jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 
                                     if jnErro == True:
                                          appDlgConf_Exclusao = Application().connect(title_re=".*Tesouraria.*")
                                          window = appDlgConf_Exclusao.Dialog
                                          window.Wait('ready')
                                          button = window[u'&Sim']
                                          button.click_input()

                                          varFuncao.esperar_fechamento_janela("Atenção")

                                          break

                                 guptamdiframeAcer_Operador[u'Edit32'].type_keys("^c")
                                 controleOutros_box =  pyperclip.paste() # Insere na variável

                                 if devolucao_float > 0.0:
                                    guptamdiframeAcer_Operador[u'&Outros'].click_input() 
                                    varFuncao.AguardaAberturaJanela("Movimento Detalhado")

                                    varFuncao.InsertOutros(devolucao.strip(), controleOutros_box)
                                    #pyautogui.sleep(2)

                              #Captura o valor da diferença e verifica se ela é igual a difereça entre o dinheiro da tebela  e o dinheiro do sistema
                              vlr_diferanca_capturado = guptamdiframeAcer_Operador.Edit49 #[u'Edit49'] .type_keys("^c")  Executa a copia  do controle digitacao 
                              srtDiferenca = vlr_diferanca_capturado.window_text() 
                              #srtDiferenca = pyperclip.paste() # Insere na variável
                              
                              #Captura o valor do dinheiro do sistema para para o calculo            
                           #   srtDinheiroSistema_2 = guptamdiframeAcer_Operador.Edit15#[u'&DinheiroEdit3'].type_keys("^c") 
                           #   srtDinheiroSistema = srtDinheiroSistema_2.window_text()
                              #srtDinheiroSistema = pyperclip.paste() # Insere na variável
                                                             
                              float_diferenca = round(varCalculos.convertStringPadraoUS(srtDiferenca),2)

                              float_dinheiro_tabela = 0
                              if dinheiro.strip() != " " and dinheiro.strip() != "":
                                  float_dinheiro_tabela = round(varCalculos.convertStringPadraoUS(dinheiro),2)  

                              float_dinheiroSistema = round(varCalculos.convertStringPadraoUS(srtDinheiroSistema),2)

                              float_diferencaCalculada = float_dinheiroSistema -float_dinheiro_tabela 
                              float_diferencaCalculada = round(float_diferencaCalculada,2)

                              if float_diferenca != float_diferencaCalculada:
                                 varFuncao.GeraLogsInfo(f"PDV COM DIFERGENCIA DE VALORES -- LOJA : {bd_nroempresa} -- PDV BANCO:{bd_nropdv} -- TURNO : {bd_turno} -- VALOR DA DIFERENÇA SISTEMA : {float_diferenca} -- DIFERENCA CALCULADA :{float_diferencaCalculada}")
                                 guptamdiframeAcer_Operador[u'Button31'].click_input() #Chama o proximo pdv no formulario
                                 
                                 jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'&Sim']
                                    button.click_input()

                                    varFuncao.esperar_fechamento_janela("Atenção")
                                 
                                 pyautogui.sleep(0.25)
                                 break

                             # pyautogui.sleep(1)      
                              guptamdiframeAcer_Operador[u'&OutrosEdit2'].type_keys("^c") # Executa a copia  do controle digitacao   
                              varControleDigitacao = pyperclip.paste() # Insere na variável

                             # guptadialog_01 = None   #linha inserida para teste
                             # window = None                                                                                                                                          
                                                         
                              if varControleDigitacao >= "1":

                                 guptamdiframeAcer_Operador[u'Button11'].click_input() #Clica em Outros se o módulo precisar ser conciliado  

                                 varFuncao.AguardaAberturaJanela("Movimento Detalhado")
                                
                                 guptamdiframeAcer_Detalhamento_01 = Application().connect(title_re=".*Tesouraria.*") 
                                 guptadialog_01 = guptamdiframeAcer_Detalhamento_01[u'Gupta:Dialog']
                                 guptadialog_01.Wait('ready')  
                                 
                                 guptadialog_01[u'Button3'].double_click_input() #Click na no  botão conciliar janela outros
                                 pyautogui.sleep(0.25)
                                 guptadialog_01[u'Valor TotalButton2'].double_click_input() # Confirma e fecha a janela outros    
                                 varControleDigitacao =None 
                                 
                                 """
                                 jnErro = varFuncao.check_window_exists("Movimento Detalhado")
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Movimento Detalhado.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'Button6']
                                    button.click_input()
                                 """
                                 varFuncao.esperar_fechamento_janela("Movimento Detalhado")
                              #pyautogui.sleep(1.5)

                              #Cheque a Vista ##############################################################################################################################################
                              guptamdiframeAcer_Operador[u'Edit17'].type_keys("^c") 
                              varControleDigitacao = pyperclip.paste() # Insere na variável
                          
                              if varControleDigitacao >= "1":

                                 guptamdiframeAcer_Operador[u'Button6'].click_input() #Click no botão cheque a vista  
                                 varFuncao.AguardaAberturaJanela("Movimento Detalhado") 
                                
                                 guptamdiframeAcer_Detalhamento_01 = Application().connect(title_re=".*Tesouraria.*")
                                 guptadialog_01 = guptamdiframeAcer_Detalhamento_01[u'Gupta:Dialog']
                                 guptadialog_01.Wait('ready')         
                                 
                                 guptadialog_01[u'Button2'].click_input() #Click na no  botão conciliar cheque a vista
                                 pyautogui.sleep(1.5)
                                 guptadialog_01[u'Valor TotalButton2'].double_click_input() # Confirma e fecha a cheque a vista
                                 varControleDigitacao =None 

                                 """ 
                                
                                 jnErro = varFuncao.check_window_exists("Movimento Detalhado")
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Movimento Detalhado.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'Button5']
                                    button.click_input()
                                 """
                                 varFuncao.esperar_fechamento_janela("Movimento Detalhado")
                              pyautogui.sleep(1) 

                              #Conciliando cheque a Prazo #######################################################################################################################################
                              guptamdiframeAcer_Operador[u'Edit20'].type_keys("^c") 
                              varControleDigitacao = pyperclip.paste() # Insere na variável
                          
                              if varControleDigitacao >= "1":

                                 guptamdiframeAcer_Operador[u'Button7'].click_input() #Click no botão cheque a prazo  
                                 varFuncao.AguardaAberturaJanela("Movimento Detalhado")
                                
                                 guptamdiframeAcer_Detalhamento_01 = Application().connect(title_re=".*Tesouraria.*")
                                 guptadialog_01 = guptamdiframeAcer_Detalhamento_01[u'Gupta:Dialog']
                                 guptadialog_01.Wait('ready')        
                                 
                                 guptadialog_01[u'Button2'].click_input() #Click na no  botão conciliar cheque a prazo
                                 pyautogui.sleep(1.5)
                                 guptadialog_01[u'Valor TotalButton2'].double_click_input() # Confirma e fecha a cheque a prazo
                                  
                                 varControleDigitacao =None 
                                 """
                                 jnErro = varFuncao.check_window_exists("Movimento Detalhado")
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Movimento Detalhado.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'Button5']
                                    button.click_input()
                                 """
                              varFuncao.esperar_fechamento_janela("Movimento Detalhado")
                              pyautogui.sleep(1.5) 

                              
                              #Conciliando Credito ###############################################################################################################################################
                              #Captura o controle do modo Credito
                              guptamdiframeAcer_Operador[u'Edit23'].type_keys("^c") 
                              varControleDigitacao = pyperclip.paste() # Insere na variável
                              pyautogui.sleep(0.25) 
                              if varControleDigitacao >= "1":

                                 guptamdiframeAcer_Operador[u'Button8'].click_input() #Click no botão crédito  
                                 varFuncao.AguardaAberturaJanela("Movimento Detalhado") 
                                 
                                 guptamdiframeAcer_Detalhamento_01 = Application().connect(title_re=".*Movimento Detalhado.*")
                                 guptadialog_01 = guptamdiframeAcer_Detalhamento_01[u'Gupta:Dialog']
                                 guptadialog_01.Wait('ready')         
                                
                                 guptadialog_01[u'Button3'].click_input() #Click na no  botão conciliar janela credito

                                 pyautogui.sleep(3)

                                 guptadialog_01[u'Button6'].double_click_input() # Confirma e fecha a janela credito

                                # pyautogui.sleep(4) 
                                                                
                                 varControleDigitacao =None 
                                 """
                                 jnErro = varFuncao.check_window_exists("Movimento Detalhado")
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Movimento Detalhado.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'Button6']
                                    button.click_input()
                              """
                              varFuncao.esperar_fechamento_janela("Movimento Detalhado")   
                              
                              pyautogui.sleep(1.5) 

                              #Conciliando Cartao Débito ############################################################################################################################################
                              guptamdiframeAcer_Operador[u'Edit26'].type_keys("^c") 
                              varControleDigitacao = pyperclip.paste() # Insere na variável
                          
                              if varControleDigitacao >= "1":

                                 guptamdiframeAcer_Operador[u'Button9'].click_input() #Click no botão Débito  
                                 varFuncao.AguardaAberturaJanela("Movimento Detalhado") 
                                 
                                 guptamdiframeAcer_Detalhamento_01 = Application().connect(title_re=".*Tesouraria.*")
                                 guptadialog_01 = guptamdiframeAcer_Detalhamento_01[u'Gupta:Dialog']
                                 guptadialog_01.Wait('ready')         
                                 
                                 guptadialog_01[u'Button3'].click_input() #Click na no  botão conciliar debito
                                 pyautogui.sleep(1.5)
                                 guptadialog_01[u'Valor TotalButton2'].double_click_input() # Confirma e fecha a janela debito
                                 #pyautogui.sleep(3)
                                 varControleDigitacao =None
                                 
                                 """
                                 
                                 jnErro = varFuncao.check_window_exists("Movimento Detalhado")
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Movimento Detalhado.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'Button6']
                                    button.click_input()
                                """
                              varFuncao.esperar_fechamento_janela("Movimento Detalhado")
                             # pyautogui.sleep(1.5)

                              #Conciliando Ticket ######################################################################################################################################################
                              guptamdiframeAcer_Operador[u'Edit29'].type_keys("^c") 
                              varControleDigitacao = pyperclip.paste() # Insere na variável
                          
                              if varControleDigitacao >= "1":

                                 guptamdiframeAcer_Operador[u'Button10'].click_input() #Click no botão Ticket  
                                 varFuncao.AguardaAberturaJanela("Movimento Detalhado")  
                                 
                                 guptamdiframeAcer_Detalhamento_01 = Application().connect(title_re=".*Tesouraria.*")
                                 guptadialog_01 = guptamdiframeAcer_Detalhamento_01[u'Gupta:Dialog']
                                 guptadialog_01.Wait('ready')         
                                 
                                 guptadialog_01[u'Button3'].click_input() #Click na no  botão conciliar Ticket
                                 pyautogui.sleep(1.5)
                                 guptadialog_01[u'Valor TotalButton2'].double_click_input() # Confirma e fecha a janela Ticket
                                 pyautogui.sleep(2)
                                  
                                 varControleDigitacao =None 
                                 """
                                 jnErro = varFuncao.check_window_exists("Movimento Detalhado")
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Movimento Detalhado.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'Button6']
                                    button.click_input()
                                 """
                              varFuncao.esperar_fechamento_janela("Movimento Detalhado")
                             # pyautogui.sleep(1)

                              #Conciliando Cancelamento #######################################################################################################################################################
                              guptamdiframeAcer_Operador[u'Edit36'].type_keys("^c") #[u'Edit35'].
                              varControleDigitacao = pyperclip.paste() # Insere na variável
                          
                              if varControleDigitacao >= "1":

                                 guptamdiframeAcer_Operador[u'Button12'].click_input() #Click no botão Cancelamento  
                                 varFuncao.AguardaAberturaJanela("Movimento Detalhado")  
                                
                                 guptamdiframeAcer_Detalhamento_01 = Application().connect(title_re=".*Tesouraria.*")
                                 guptadialog_01 = guptamdiframeAcer_Detalhamento_01[u'Gupta:Dialog']
                                 guptadialog_01.Wait('ready')       
                                 
                                 guptadialog_01[u'Button2'].click_input() #Click na no  botão conciliar Cancelamento
                                 pyautogui.sleep(1)
                                 guptadialog_01[u'Valor TotalButton2'].double_click_input() # Confirma e fecha a janela Cancelamento
                                 pyautogui.sleep(1)

                                 #Ocorreu um bug no modulo cancelamento no qual o campo que mostra se o ítem deve ser conciliado veio com uma informação errada, fazendo o robo da um click desnecessario              
                                 jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'&Sim']
                                    button.click_input()

                                    appDlgConf_Exclusao = Application().connect(title_re=".*Movimento Detalhado.*")
                                    window_e = appDlgConf_Exclusao.Dialog
                                    window_e.Wait('ready')
                                    window_e.close()
                                 varControleDigitacao =None  

                                 """ 
                                 
                                 jnErro = varFuncao.check_window_exists("Movimento Detalhado")
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Movimento Detalhado.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'Button5']
                                    button.click_input()
                                    """
                                    
                              varFuncao.esperar_fechamento_janela("Movimento Detalhado")
                              pyautogui.sleep(1)

                              #Conciliando convênio #############################################################################################################################################################
                              guptamdiframeAcer_Operador[u'Edit38'].type_keys("^c") 
                              varControleDigitacao = pyperclip.paste() # Insere na variável
                          
                              if varControleDigitacao >= "1":

                                 guptamdiframeAcer_Operador[u'Button13'].click_input() #Click no botão Convenio  
                                 varFuncao.AguardaAberturaJanela("Movimento Detalhado") 
                             
                                 guptamdiframeAcer_Detalhamento_01 = Application().connect(title_re=".*Tesouraria.*")
                                 guptadialog_01 = guptamdiframeAcer_Detalhamento_01[u'Gupta:Dialog']
                                 guptadialog_01.Wait('ready')       
                                 
                                 guptadialog_01[u'Button3'].click_input() #Click na no  botão conciliar Convenio
                                 pyautogui.sleep(1.5)
                                 guptadialog_01[u'Valor TotalButton2'].double_click_input() # Confirma e fecha a janela Convenio
                                # pyautogui.sleep(1.5)
                                                                  
                                 varControleDigitacao =None
                                 """ 
                                 jnErro = varFuncao.check_window_exists("Movimento Detalhado")
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Movimento Detalhado.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'Button6']
                                    button.click_input()
                                 """
                              #pyautogui.sleep(1)
                              varFuncao.esperar_fechamento_janela("Movimento Detalhado")

                              #Conciliando Carteira Digital #####################################################################################################################################################
                              guptamdiframeAcer_Operador[u'Edit41'].type_keys("^c") 
                              varControleDigitacao = pyperclip.paste() # Insere na variável
                          
                              if varControleDigitacao >= "1":

                                 guptamdiframeAcer_Operador[u'Button14'].click_input() #Click no botão Carteira Digital  
                                 varFuncao.AguardaAberturaJanela("Movimento Detalhado")  
                                 
                                 guptamdiframeAcer_Detalhamento_01 = Application().connect(title_re=".*Tesouraria.*")
                                 guptadialog_01 = guptamdiframeAcer_Detalhamento_01[u'Gupta:Dialog']
                                 guptadialog_01.Wait('ready')         
                                
                                 guptadialog_01[u'Button2'].click_input() #Click na no  botão conciliar Carteira Digital
                                 pyautogui.sleep(1.5)
                                 guptadialog_01[u'Valor TotalButton2'].double_click_input() # Confirma e fecha a janela Carteira Digital
                                 pyautogui.sleep(1)
                                  
                                 varControleDigitacao =None
                                 """
 
                                 jnErro = varFuncao.check_window_exists("Movimento Detalhado")
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Movimento Detalhado.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'Button5']
                                    button.click_input()
                                 """                             
                              varFuncao.esperar_fechamento_janela("Movimento Detalhado")
                              #pyautogui.sleep(1)

                              #Verificando se existe quebra ou sobra : 
                              if float_diferenca < 0.0:

                                 guptamdiframeAcer_Operador[u'SButton'].click_input()
                                 pyautogui.sleep(0.45)

                                 jnErro = varFuncao.check_window_exists("Informe")
                                 if jnErro == True:
                                    guptamdiframeAcer_Detalhamento_01 = Application().connect(title_re=".*Tesouraria.*")
                                    guptadialog_01 = guptamdiframeAcer_Detalhamento_01[u'Gupta:Dialog']
                                    guptadialog_01.Wait('ready')       
                                    guptadialog_01[u'&Sim'].click_input() #Click na no  botão confirma sobra
                               
                              
                                 guptamdiframeAcer_Operador[u'Button40'].click_input() # Clicando em Conciliar
                                 pyautogui.sleep(1)
                                 jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'Button'] #u'&OK'
                                    button.click_input()
                                    pyautogui.sleep(0.25)

                                 guptamdiframeAcer_Operador[u'Button31'].click_input() #Chama o proximo pdv no formulario
                                 pyautogui.sleep(1.5)
                                 
                                 jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'Button'] #u'&Sim'
                                    button.click_input()
                                 
                                 #Após o pdv ser conciliado, a linha digita vai ser apagada pelo codigo abaixo
                                 indices_para_remover.add(index)
                                 with open(varArquivosLojas, 'w') as c:
                                    c.write(lines[0])  # Escreve o cabeçalho
                                    for index, line in enumerate(lines[1:], start=1):
                                          if index not in indices_para_remover:
                                             c.write(line)


                                 pyautogui.sleep(1)
                                 
                                 break
                                                        
                              elif float_diferenca == 0.0:
                                                                                           
                                 guptamdiframeAcer_Operador[u'Button40'].click_input() # Clicando em Conciliar
                               
                                 pyautogui.sleep(1)
                                 jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 

                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'Button'] #[u'&Sim', u'Button', u'Button1', u'Button0', u'&SimButton']
                                    button.click_input()
                                  #  window.close()

                                    varFuncao.esperar_fechamento_janela("Atenção")

                                    pyautogui.sleep(0.25)

                                 guptamdiframeAcer_Operador[u'Button31'].click_input() #Chama o proximo pdv no formulario
                                 pyautogui.sleep(1.5)
                                 
                                 jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'&Sim']
                                    button.click_input()
                                    varFuncao.esperar_fechamento_janela("Atenção")

                                 #Após o pdv ser conciliado, a linha digita vai ser apagada pelo codigo abaixo
                                 indices_para_remover.add(index)
                                 with open(varArquivosLojas, 'w') as c:
                                    c.write(lines[0])  # Escreve o cabeçalho
                                    for index, line in enumerate(lines[1:], start=1):
                                          if index not in indices_para_remover:
                                             c.write(line)

                                 pyautogui.sleep(1)
                              
                                 break
                              else:
                                 guptamdiframeAcer_Operador[u'QButton'].click_input()

                                 pyautogui.sleep(0.45)

                                 jnErro = varFuncao.check_window_exists("Informe")
                                 if jnErro == True:
                                    guptamdiframeAcer_Detalhamento_01 = Application().connect(title_re=".*Informe.*")
                                    guptadialog_01 = guptamdiframeAcer_Detalhamento_01[u'Gupta:Dialog']
                                    guptadialog_01.Wait('ready')       
                                    guptadialog_01[u'&Sim'].click_input() #Click na no  botão confirma a quebra  
                                    varFuncao.esperar_fechamento_janela("Informe")

                                 guptamdiframeAcer_Operador[u'Button40'].click_input() # Clicando em Conciliar   

                                 pyautogui.sleep(1.5)
                                 jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'Button']
                                    button.click_input()
                                    pyautogui.sleep(1)
                                                                
                                 #_______________________________________________________________________________________________________________________________________________________________   
                                
                                 guptamdiframeAcer_Operador[u'Button31'].click_input() #Chama o proximo pdv no formulario
                                 
                                 jnErro = varFuncao.check_window_exists("Atenção") # Se abrir uma caixa de diálogo com o nome atenção, 
                                 if jnErro == True:
                                    appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
                                    window = appDlgConf_Exclusao.Dialog
                                    window.Wait('ready')
                                    button = window[u'&Sim']
                                    button.click_input()

                                 #Após o pdv ser conciliado, a linha digita vai ser apagada pelo codigo abaixo
                                 indices_para_remover.add(index)
                                 with open(varArquivosLojas, 'w') as c:
                                    c.write(lines[0])  # Escreve o cabeçalho
                                    for index, line in enumerate(lines[1:], start=1):
                                          if index not in indices_para_remover:
                                             c.write(line)

                                 pyautogui.sleep(1)
                                                            

                                 break

                        #VERIFICAR BUG, QAUANDO O PDV NÃO É ENCONTRATO É CHAMADO UM OUTRO PDV NO BANCO MAS NÃO NO FORMULÁRIO           
                        else:
                           #CASO FALHE A VALIDAÇÃO DO CRUZAMENTO CODIGO DO CODIGO E O TURNO, ESTE ESLE VAI SER EXECUTADO, CHAMANDO OUTRA LINHA DA PLANILHA 
                         #  mensagem =f"FALHA NA VALIDAÇÃO PDV TURNO: PDV_PLANILHA : {bd_nropdv} -- TURNO_PLANILHA : {bd_turno} --OPERADOR : {bd_nomeOperadora} --PLANILHA :{pdv} - {turnoCsv} "
                          # varFuncao.show_popup(mensagem) bd_nropdv
                            if index == len(lines) -1: # Se for a ultima interação , pdv do formulario e pode chamar o continue mesmo porque o fluxo seguirá normalmente.
                               guptamdiframeAcer_Operador[u'Button31'].click_input()

                               varFuncao.GeraLogsInfo(f"LOJA : {bd_nroempresa} PDV NÃO FOI ENCONTRADO --  OPERADORA : {bd_nomeOperadora} -- PDV BANCO:{bd_nropdv} -- TURNO : {bd_turno}")

                               continue
                            else:
                               continue
                                              
                      elif bd_nropdv < int_pdv:# SE O NUMERO DO PDV NÃO FOR IGUAL, ESSa LINHA TESTA SE O NUMERO DO PDV DESEJADO É MENOR QUE O LISTADO NO MOMENTO, SE FOR, CHAMA UM OUTRO PDV NO BANCO
                        #CHAMAR O PROXIMO PDV NO BANCO E EXECUTA O CLICK O PROXIMO PDV NO FORMULARIO
                        guptamdiframeAcer_Operador[u'Button31'].click_input() #Chama o proximo pdv no formulario
                        pyautogui.sleep(2)
                        break
                      else:
                         if index == len(lines) -1: # Se for a ultima interação , pdv do formulario e pode chamar o continue mesmo porque o fluxo seguirá normalmente.
                            guptamdiframeAcer_Operador[u'Button31'].click_input() 

                            varFuncao.GeraLogsInfo(f"LOJA : {bd_nroempresa} PDV NÃO FOI ENCONTRADO --  OPERADORA : {bd_nomeOperadora} -- PDV BANCO:{bd_nropdv} -- TURNO : {bd_turno}")
                           
                            continue
                         else:
                            print(f"Procurando...{pdv}")
                            continue
                     
               else:
                   #Esta linha é chamada se o arquivo da loja, não for encontrado. Chama outra loja
                   #break
                   continue
               
            #Metodo que apaga o arquivo que foi digitado
            varFuncao.verifica_e_apaga_arquivo(validacaoLojas)   
        
        #FECHA A JANELA ACERTO DE OPERADOR AO FINAL DO LOOP DO MOVIMENTO DA LOJA DO BANCO DE DADOS
        jnErro = varFuncao.check_window_exists("Tesouraria") # Se a consiliação do ultimo pdv da loja anterior falhar, esse bloco será chamado
        if jnErro == True:
               fechaJanela = Application().connect(title_re=".*Tesouraria.*")
               fecha = fechaJanela.window(class_name='Gupta:MDIFrame')
               fecha.Wait('ready')
               button = fecha[u'Button24']#Fecha a janela fora do loop
               button.click_input()
        #guptamdiframeAcer_Operador[u'Button24'].click_input() #Clicando na lupa da janela   
          
        #FECHA A JANELA ACERTO DE OPERADOR AO FINAL DO LOOP DO MOVIMENTO DAs LOJAs DO BANCO DE DADOS
        pyautogui.sleep(2)
        #achou = varFuncao.SeJanelaExiste_porImagem(varJn_acerto_operador)
        #if achou == True:
         #guptamdiframeAcer_Operador[u'Button24'].click_input() #Clicando na lupa da janela  

        pyautogui.sleep(0.35)
        jnErro = varFuncao.check_window_exists("Atenção") # Se a consiliação do ultimo pdv da loja anterior falhar, esse bloco será chamado
        if jnErro == True:
               appDlgConf_Exclusao = Application().connect(title_re=".*Tesouraria.*")
               window = appDlgConf_Exclusao.Dialog
               window.Wait('ready')
               button = window[u'&Sim']
               button.click_input()
      
            

            
                

       