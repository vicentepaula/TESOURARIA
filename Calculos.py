from decimal import Decimal

class Operacoes:

    def convertStringEmDecimal(self, valor_ptbr):

        if 'R$' in valor_ptbr:
            # Remover o cifrão
            valor_ptbr = valor_ptbr.replace('R$','').lstrip()

        # Remover separador de milhares (pontos)
        valor_sem_pontos = valor_ptbr.replace('.', '').strip()
        # Substituir vírgula por ponto (separador decimal)
        valor_com_ponto = valor_sem_pontos.replace(',', '.')
        # Converter para Decimal
        return Decimal(valor_com_ponto)
    

    def convertStringPadraoUS(self, valor_ptbr):

        if 'R$' in valor_ptbr:
             # Remover o cifrão
            valor_ptbr = valor_ptbr.replace('R$','').lstrip()

        # Remover separador de milhares (pontos)
        valor_sem_pontos = valor_ptbr.replace('.','').strip()
        # Substituir vírgula por ponto (separador decimal)
        valor_com_ponto = valor_sem_pontos.replace(',','.')
        # Converter para Decimal
        return float(valor_com_ponto)
    

    def removeCifraoRetornaString(self, valor_ptbr):
        if 'R$' in valor_ptbr:
             # Remover o cifrão
            valor_ptbr = valor_ptbr.replace('R$','').lstrip()

        return valor_ptbr

