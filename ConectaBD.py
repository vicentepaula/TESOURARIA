import os
import oracledb
oracledb.init_oracle_client(lib_dir=r"C:\\Projetos_Python\\BibliotecasOraclePython\\instantclient_21_13")

class ConexaoOracle:
   

    def __init__(self):
           self.conexao = None

    def conectar(self):
        try: 
            self.conexao = oracledb.connect(
            user="consinco",
            password= os.getenv('pw_bd'),
            dsn="10.102.227.2/arcomix.subnetarcomixda.vcnrootskyoneda.oraclevcn.com")
            print("Conexão bem-sucedida!")
            return self.conexao
            
        except self.conexao.Error as e:
            print(f"Erro ao conectar ao Oracle: {e}")


    def conectar_base_teste(self):
        try: 
            self.conexao = oracledb.connect(
            user="consinco",
            password=os.getenv('pw_bd'),
            dsn="10.10.10.114/c5db.arcoiris.local")
            print("Conexão bem-sucedida!")
            return self.conexao
            
        except self.conexao.Error as e:
            print(f"Erro ao conectar ao Oracle: {e}")
        
 
         
    def desconectar(self):
        if self.conexao:
            self.conexao.close()
            print("Conexão fechada !!!.")
        else:
            print("Nenhuma conexão ativa.")

