import ConectaBD
#from ConnectionBD import Connect_BD

class DAO:

        def __init__(self):
                self.con = ConectaBD.ConexaoOracle

        def getConection(self):
                return  ConectaBD.ConexaoOracle()

        def executaQuery(self, query):
                conexao = self.con.conectar(self)
                resultados = []

                # Verifica se a conexão foi bem-sucedida
                if conexao is not None:
                        try:
                                cursor = conexao.cursor()

                                # Executa a consulta
                                for row in cursor.execute(query):
                                        resultados.append(row)

                                # Fecha o cursor
                                cursor.close()

                        except Exception as e:
                         print(f"Erro ao executar a consulta: {e}")

                        finally:
                        # Sempre desconecta, mesmo em caso de exceção
                         self.con.desconectar(self)
                         

                else:
                        print("Falha ao conectar ao banco de dados. CONEXAO É NULL.")

                return resultados
        

        def RetornaMovimento(self,nro_loja, dta_movimento):
                query = f"""
                        SELECT 
                        FI_TSMOVTOOPERADOR.NROEMPRESA, 
                        TO_CHAR(FI_TSMOVTOOPERADOR.DTAMOVIMENTO, 'dd/mm/yyyy') AS data,
                        FI_TSMOVTOOPERADOR.NROTURNO,
                        FI_TSMOVTOOPERADOR.GTINICIO,
                        FI_TSMOVTOOPERADOR.GTFINAL,
                        FI_TSMOVTOOPERADOR.ENCARGOS,
                        FI_TSMOVTOOPERADOR.VLRTOTALNFENFCESAT,
                        FI_TSMOVTOOPERADOR.TOTAL,
                        FI_TSMOVTOOPERADOR.VLRBANCARIO,
                        FI_TSMOVTOOPERADOR.QTDEDOCBANCARIO,
                        FI_TSMOVTOOPERADOR.NROEMPRESAMAE,
                        FI_TSMOVTOOPERADOR.ACERTADO,
                        FI_TSMOVTOOPERADOR.FECHADO,
                        FI_TSMOVTOOPERADOR.USUFECHOU,
                        FI_TSMOVTOOPERADOR.DTAFECHOU,
                        FI_TSMOVTOOPERADOR.USUALTERACAO,
                        FI_TSMOVTOOPERADOR.VERSAO,
                        FI_TSMOVTOOPERADOR.SEQIDENTIFICA,
                        NVL(FI_TSMOVTOOPERADOR.INDQUEBRAECF, 'N'),
                        FI_TSMOVTOOPERADOR.NROPDV,
                        FI_TSMOVTOOPERADOR.SOFTPDV,
                        FI_TSMOVTOOPERADOR.CODOPERADOR,
                        NVL(FI_TSMOVTOOPERADOR.SOFTPDV, 'DIG.MANUALMENTE'), 
                        FI_TSMOVTOOPERADOR.USUACERTOU,
                        FI_TSMOVTOOPERADOR.DTAHORAACERTOU,
                        usu.NOME
                        FROM FI_TSMOVTOOPERADOR, ge_usuario usu
                        WHERE FI_TSMOVTOOPERADOR.NROEMPRESA = :nro_loja
                        AND NVL(VERSAO, 'A') = 'N'
                        AND FI_TSMOVTOOPERADOR.DTAMOVIMENTO = TO_DATE(:dta_movimento, 'dd/mm/yyyy')
                        AND usu.SEQUSUARIO = FI_TSMOVTOOPERADOR.CODOPERADOR
                        ORDER BY NROPDV, NROTURNO
                """

                conexao = self.con.conectar(self)
                resultados = []

                # Verifica se a conexão foi bem-sucedida
                if conexao is not None:
                     try:
                           cursor = conexao.cursor()

                        # Executa a consulta com parâmetros
                           cursor.execute(query, {'nro_loja': nro_loja, 'dta_movimento': dta_movimento})
                           resultados = cursor.fetchall()

                        # Fecha o cursor 
                           cursor.close()

                     except Exception as e:
                          print(f"Erro ao executar a consulta: {e}")

                     finally:
                        # Sempre desconecta, mesmo em caso de exceção
                        self.con.desconectar(self)

                else:
                        print("Falha ao conectar ao banco de dados. CONEXAO É NULL.")

                return resultados
             
             
             
                
                
                    

        
                     