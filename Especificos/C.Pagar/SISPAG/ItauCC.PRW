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

neste exemplo o ultimo digito 3 é o digito da conta.     

*/
//Retira espacos, tracos e barras do numero da conta
//_cConta := STRTRAN(STRTRAN(STRTRAN(SA2->A2_NUMCON," ",""),"-",""),"/","")

//Separa Conta e Digito
//_cCnt   := SUBSTR(_cConta, 1, LEN(_cConta) - 1)  
//_cDig   := SUBSTR(_cConta, LEN(_cConta), 1)

If RTRIM(SA2->A2_BANCO) $ "341/409"
   /*
   Nome do Campo 	Significado 					Posição 	Picture 	Conteúdo
   ZEROS 			COMPLEMENTO DE REGISTRO 		024 024 	9(01)
   AGÊNCIA 			NÚMERO AGÊNCIA CREDITADA 		025 028 	9(04)
   BRANCOS 			COMPLEMENTO DE REGISTRO 		029 029 	x(01)
   ZEROS 			COMPLEMENTO DE REGISTRO 		030 035 	9(06)
   CONTA 			NÚMERO DE C/C CREDITADA 		036 041 	9(06)
   BRANCOS 			COMPLEMENTO DE REGISTRO 		042 042 	X(01)
   DAC 				DAC DA AGÊNCIA/CONTA CREDITADA 	043 043     9(01)
   */
   //_Conta :="0"+STRZERO(VAL(SA2->A2_AGENCIA),4)+" "+"0000000"+STRZERO(VAL(_cCnt),5,0)+" "+ _cDig
   _Conta :="0"+STRZERO(VAL(SA2->A2_AGENCIA),4)+" "+"0000000"+STRZERO(VAL(SA2->A2_NUMCON),5,0)+" "+Trim(SA2->A2_DVCTA)

ElseIf RTRIM(SA2->A2_BANCO) $ "104" //Tratamento para conta da caixa com tipo de operacao da conta
   /*
   Nome do Campo 	Significado 					Posição 	Picture 	Conteúdo
   ZEROS 			COMPLEMENTO DE REGISTRO 		024 024 	9(01)
   AGÊNCIA 			NÚMERO AGÊNCIA CREDITADA 		025 028 	9(04)
   BRANCOS 			COMPLEMENTO DE REGISTRO 		029 029 	x(01)
   ZERO 			COMPLEMENTO DE REGISTRO 		030 030 	9(01)
   OPERACAO 	    TIPO DE CONTA CAIXA  		    031 033 	9(03)       (*)Tabela de numero de operação
   CONTA 			NÚMERO DE C/C CREDITADA 		034 041 	9(08)
   BRANCOS 			COMPLEMENTO DE REGISTRO 		042 042 	X(01)
   DAC 				DAC DA AGÊNCIA/CONTA CREDITADA 	043 043     9(01)                              
   

   (*) Tabela de numero de operação / Tipo de Conta
   001 – Conta Corrente de Pessoa Física
   002 – Conta Simples de Pessoa Física
   003 – Conta Corrente de Pessoa Jurídica
   006 – Entidades Públicas
   007 – Depósitos Instituições Financeiras
   013 – Poupança de Pessoa Física
   022 – Poupança de Pessoa Jurídica
   023 – Conta Caixa Fácil
   028 – Poupança de Crédito Imobiliário
   032 – Conta Investimento Pessoa Física
   034 – Conta Investimento Pessoa Jurídica
   037 – Conta Salário
   043 – Depósitos Lotéricos
   131 – Poupança Integrada
   */
  // _Conta :="0"+STRZERO(VAL(SA2->A2_AGENCIA),4)+" "+"0"+SA2->A2_XOPECEF+STRZERO(VAL(_cCnt),8,0)+" "+ _cDig
     _Conta :="0"+STRZERO(VAL(SA2->A2_AGENCIA),4)+" "+"0"+SA2->A2_XOPECEF+STRZERO(VAL(SA2->A2_NUMCON),8,0)+" "+Trim(SA2->A2_DVCTA)
Else                                             
   /*
   Nome do Campo 	Significado 					Posição 	Picture 	Conteúdo
   Agência 			Número agência CREDITADA 		024 028 	9(05)
   brancos 			Complemento de registro 		029 029 	X(01)
   Conta 			Número de C/C CREDITADA 		030 041 	9(12)
   brancos 			Complemento de registro 		042 042 	X(01)
   DAC 				DAC DA AGÊNCIA/Conta CREDITADA 	043 043 	X(01)
   */
  //_Conta := STRZERO(VAL(SA2->A2_AGENCIA),5)+" "+STRZERO(VAL(_cCnt),12,0)+" "+ _cDig
    _Conta := STRZERO(VAL(SA2->A2_AGENCIA),5)+" "+STRZERO(VAL(SA2->A2_NUMCON),12,0)+" "+ Trim(SA2->A2_DVCTA)
EndIf
  
Return(_Conta)