#include "rwmake.ch" 

User Function ItauCC() 

// Itau - Posicoes( 024 - 043 )

Local _cConta
Local _cCnt
Local _Conta   
Local _cDig

/*
  
Campo de conta preencher com o digito da conta no final
      
Formato: CCCCCX

Ex.:
079393

neste exemplo o ultimo digito 3 � o digito da conta.     

*/
//Retira espacos, tracos e barras do numero da conta
//_cConta := STRTRAN(STRTRAN(STRTRAN(SA2->A2_NUMCON," ",""),"-",""),"/","")

//Separa Conta e Digito
//_cCnt   := SUBSTR(_cConta, 1, LEN(_cConta) - 1)  
//_cDig   := SUBSTR(_cConta, LEN(_cConta), 1)

If RTRIM(SA2->A2_BANCO) $ "341/409"
   /*
   Nome do Campo 	Significado 					Posi��o 	Picture 	Conte�do
   ZEROS 			COMPLEMENTO DE REGISTRO 		024 024 	9(01)
   AG�NCIA 			N�MERO AG�NCIA CREDITADA 		025 028 	9(04)
   BRANCOS 			COMPLEMENTO DE REGISTRO 		029 029 	x(01)
   ZEROS 			COMPLEMENTO DE REGISTRO 		030 035 	9(06)
   CONTA 			N�MERO DE C/C CREDITADA 		036 041 	9(06)
   BRANCOS 			COMPLEMENTO DE REGISTRO 		042 042 	X(01)
   DAC 				DAC DA AG�NCIA/CONTA CREDITADA 	043 043     9(01)
   */
   //_Conta :="0"+STRZERO(VAL(SA2->A2_AGENCIA),4)+" "+"0000000"+STRZERO(VAL(_cCnt),5,0)+" "+ _cDig
   _Conta :="0"+STRZERO(VAL(SA2->A2_AGENCIA),4)+" "+"0000000"+STRZERO(VAL(SA2->A2_NUMCON),5,0)+" "+Trim(SA2->A2_DVCTA)

ElseIf RTRIM(SA2->A2_BANCO) $ "104" //Tratamento para conta da caixa com tipo de operacao da conta
   /*
   Nome do Campo 	Significado 					Posi��o 	Picture 	Conte�do
   ZEROS 			COMPLEMENTO DE REGISTRO 		024 024 	9(01)
   AG�NCIA 			N�MERO AG�NCIA CREDITADA 		025 028 	9(04)
   BRANCOS 			COMPLEMENTO DE REGISTRO 		029 029 	x(01)
   ZERO 			COMPLEMENTO DE REGISTRO 		030 030 	9(01)
   OPERACAO 	    TIPO DE CONTA CAIXA  		    031 033 	9(03)       (*)Tabela de numero de opera��o
   CONTA 			N�MERO DE C/C CREDITADA 		034 041 	9(08)
   BRANCOS 			COMPLEMENTO DE REGISTRO 		042 042 	X(01)
   DAC 				DAC DA AG�NCIA/CONTA CREDITADA 	043 043     9(01)                              
   

   (*) Tabela de numero de opera��o / Tipo de Conta
   001 � Conta Corrente de Pessoa F�sica
   002 � Conta Simples de Pessoa F�sica
   003 � Conta Corrente de Pessoa Jur�dica
   006 � Entidades P�blicas
   007 � Dep�sitos Institui��es Financeiras
   013 � Poupan�a de Pessoa F�sica
   022 � Poupan�a de Pessoa Jur�dica
   023 � Conta Caixa F�cil
   028 � Poupan�a de Cr�dito Imobili�rio
   032 � Conta Investimento Pessoa F�sica
   034 � Conta Investimento Pessoa Jur�dica
   037 � Conta Sal�rio
   043 � Dep�sitos Lot�ricos
   131 � Poupan�a Integrada
   */
  // _Conta :="0"+STRZERO(VAL(SA2->A2_AGENCIA),4)+" "+"0"+SA2->A2_XOPECEF+STRZERO(VAL(_cCnt),8,0)+" "+ _cDig
     _Conta :="0"+STRZERO(VAL(SA2->A2_AGENCIA),4)+" "+"0"+SA2->A2_XOPECEF+STRZERO(VAL(SA2->A2_NUMCON),8,0)+" "+Trim(SA2->A2_DVCTA)
Else                                             
   /*
   Nome do Campo 	Significado 					Posi��o 	Picture 	Conte�do
   Ag�ncia 			N�mero ag�ncia CREDITADA 		024 028 	9(05)
   brancos 			Complemento de registro 		029 029 	X(01)
   Conta 			N�mero de C/C CREDITADA 		030 041 	9(12)
   brancos 			Complemento de registro 		042 042 	X(01)
   DAC 				DAC DA AG�NCIA/Conta CREDITADA 	043 043 	X(01)
   */
  //_Conta := STRZERO(VAL(SA2->A2_AGENCIA),5)+" "+STRZERO(VAL(_cCnt),12,0)+" "+ _cDig
    _Conta := STRZERO(VAL(SA2->A2_AGENCIA),5)+" "+STRZERO(VAL(SA2->A2_NUMCON),12,0)+" "+ Trim(SA2->A2_DVCTA)
EndIf
  
Return(_Conta)