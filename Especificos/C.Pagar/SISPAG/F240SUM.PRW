#include "protheus.ch"
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  � F240SUM  �Autor  �                      � Data � 14/02/19  ���
�������������������������������������������������������������������������͹��
���Desc.     � Ponto de entrada no tratamento do valor na gera��o do arq. ���
���          � do CNAB a Pagar/SISPAG                                     ���
�������������������������������������������������������������������������͹��
���Uso       �                                                            ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
User Function F240SUM()                   

Local nValor := SE2->E2_SALDO+SE2->E2_SDACRES-SE2->E2_SDDECRE
                                                                                                     
//Efetuado tratamento direto no arquivo do SISPAG.PAG sem necessidade de retirar o valor do acrescimo da soma, que estava sendo utilizado para somar outro campo de total pago.
//If SEA->EA_MODELO == "17" //GPS -> N�o soma acrescimo no total pois se trata de outros entidades.
//
//   nValor := SE2->E2_SALDO
//
//EndIf

Return nValor