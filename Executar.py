from subprocess import Popen
from pywinauto import Application
import RepositorioDAO
import pyautogui
import FuncoesAuxiliares
import Digitacao
import os
import cv2
import time


if __name__  == '__main__':
    print('EXECUTANDO')

    #pyautogui.PAUSE=1.5
    pyautogui.FAILSAFE=False
    varExecuteDAO = RepositorioDAO.DAO()

    varDtainicial = varExecuteDAO.executaQuery("SELECT TO_CHAR(SYSDATE - 2, 'DD/MM/YYYY') AS data_dia_anterior FROM dual")[0][0] # Os dois zeros faz com que seja retornado apenas a data
       
          
    varNome="automacao"
    varSenha = os.getenv('pw_automacao')
    
    #Variaveis tela de login C5
    varImgLoginC5 ="C:\\Projetos_Python\\Tesouraria\\Img\\Janelas\\jn_login_consinco.png"
    varModuloTesouraria ="C:\\C5Client\\Tesouraria\\Tesouraria.exe \n"
         
    #ImgJanela

   # varImgJanelaFiscal="C:\\Projetos_Python\\FISCAL\\Img\\Janelas\\jn_ModuloFiscal.png"
    #Variaveis gerais da aplicacao
    varFuncao = FuncoesAuxiliares.Funcao_Apoio() 
    varExecuteDAO = RepositorioDAO.DAO()
    vardigitacao = Digitacao.Digitacao()
                  
    pyautogui.hotkey("win","r")
    pyautogui.sleep(2)
    pyautogui.typewrite(varModuloTesouraria,interval=0.015)
    
    varFuncao.aguardar_janela_por_imagem(varImgLoginC5,"Janela Login")

    time.sleep(2)
     #iniciando aplicativo da consinco
    app = Application().connect(class_name="Gupta:Dialog")
    pyautogui.sleep(2)
    #O formulário em questão não tem título, sendo assim foi identificado desta forma.
    dlg = app.window() 

    #Digitando os dados no formulário de login
    dlg['Edit4'].click_input() # Caixa de texto usuário
    dlg['Edit4'].type_keys(varNome) #Escrevendo na caixa de texto usuário
    pyautogui.sleep(1)
    dlg['Edit5'].click_input() # Click dentro da caixa de texto senha
    dlg['Edit5'].type_keys(varSenha) #Escrevendo na caixa de texto senhaouraria.exe 
    dlg['Button0'].click_input() # Click no butão entrar

    time.sleep(1)
    
    vardigitacao.acertoOperador(varDtainicial)
    #varFuncao.AguardaAberturaJanela("Tesouraria")
                    
     
    print("CONCLUIDO")
    
    

   
   
  