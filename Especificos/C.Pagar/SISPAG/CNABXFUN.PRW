#INCLUDE "RWMAKE.CH"
#INCLUDE "PROTHEUS.CH"

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³ CNABREGN ºAutor  ³                    º Data ³  21/09/20   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Monta complemento de detalhamento dos dados de tributo no   º±±
±±º          ³registro N do Cnab a pagar                                  º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³                                                            º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
User Function CNABREGN()

Local cReturn := ""
Local cPerApu := ""
Local cDtVenc := ""
Local cValTit := ""
Local cTpInsc := ""

If SEA->EA_MODELO $ "16" //16-PAGAMENTO DE TRIBUTOS - DARF NORMAL - SEGMENTO N2
                     
	cPerApu := DTOS(SE2->E2_XAPUR) //Periodo de Apuração
	cDtVenc := DTOS(SE2->E2_VENCTO)
	cValTit	:= STRZERO(((SE2->E2_SALDO+SE2->E2_ACRESC-SE2->E2_DECRESC-SE2->E2_MULTA-SE2->E2_JUROS)*100),14)
   //cTpInsc := Iif(Len(AllTrim(SE2->E2_XCGC))==14,"2","1") // Alteração p/ carregar correto a variavel - Edelcio Cano 14/11/19   
   
	If Len(AllTrim(SE2->E2_XCNPJC)) < 14
	   	cTpInsc := "1"
	Else
    	cTpInsc := "2"     
    Endif
	
	cReturn := "02"																//Identificacao do Tributo - 02 = DARF
	cReturn += STRZERO(VAL(SE2->E2_XCODREC),4)									//Codigo da Receita
	cReturn += cTpInsc															//Tipo Identificação do Contribuinte - 2 CNPJ , 1 = CPF 
	cReturn += iIf(Empty(SE2->E2_XCNPJC),Space(14),AllTrim(SE2->E2_XCNPJC))			//Identificacao do Contribuinte
	cReturn += SUBSTR(cPerApu,7,2)+SUBSTR(cPerApu,5,2)+SUBSTR(cPerApu,1,4)	//Periodo de Apuracao	- DDMMAAAA  
	cReturn += PADR(SE2->E2_E2_XREFER,17)											//Numero de Referencia
	cReturn += cValTit 															//Valor Principal
	cReturn += STRZERO(SE2->E2_MULTA,14)										//Valor da Multa
	cReturn += STRZERO(SE2->E2_JUROS,14)										//Valor dos Juros / Encargos
	cReturn += cValTit 															//Valor Pagamento
	cReturn += SUBSTR(cDtVenc,7,2)+SUBSTR(cDtVenc,5,2)+SUBSTR(cDtVenc,1,4)	//Data do Vencimento - DDMMAAAA
	cReturn += SUBSTR(cDtVenc,7,2)+SUBSTR(cDtVenc,5,2)+SUBSTR(cDtVenc,1,4)	//Data do Pagamento - DDMMAAAA
	cReturn += SPACE(30)														//Brancos
	cReturn += SE2->E2_XCONTR													//Nome do Contribuinte                                                      


ElseIf SEA->EA_MODELO $ "17" //17-PAGAMENTO DE TRIBUTOS - GPS  - SEGMENTO N1
	
	// --------------------------------------------------------------------------	
	// WRN 20180710 - Implantação da remessa de GPS
	// --------------------------------------------------------------------------
	
	cPerApu := DTOS(SE2->E2_XAPUR)
	cValArr	:= StrZero(((SE2->E2_SALDO+SE2->E2_ACRESC-SE2->E2_DECRESC-SE2->E2_MULTA-SE2->E2_JUROS)*100),14)
	cDatArr	:= DToS(SE2->E2_VENCREA)
	
	// VAlor do INSS -> Valor do título
	cValTit := StrZero(((SE2->E2_SALDO)*100),14)
	// Outras entidades -> Acrescimos do título
	cValOut := StrZero(((SE2->E2_ACRESC)*100),14)
	// Atualização monetária -> Juros e Multas
	cValAtu := StrZero(((SE2->E2_MULTA+SE2->E2_JUROS)*100),14)

	
	cReturn := '01'															// Identificação do Tributo - 01 = GPS
	cReturn += StrZero(Val(SE2->E2_XCODREC),4)								// Código do Pagamento -> campo Código da Receita no Protheus
	cReturn += SUBSTR(cPerApu,5,2)+SUBSTR(cPerApu,1,4)						// Periodo da Apuração -> Mes e ano da competencia -> MMAAAA
	cReturn += iIf(Empty(SE2->E2_XCNPJC),Space(14),AllTrim(E2_XCNPJC))		// Identificação do Contribuinte
	cReturn += cValTit														// VAlor previsto do INSS
	cReturn += cValOut														// Valor de Outras entidades
	cReturn += cValAtu														// Atualização monetária
	cReturn += cValArr														// Valor arrecadado
	cReturn += SubStr(cDatArr,7,2)+SubSTr(cDatArr,5,2)+SubStr(cDatArr,1,4)	// Data da Arrecadação -> DDMMAAAA
	cReturn += Space(8) + Space(50)											// Brancos + Informações complementares
	cReturn += SE2->E2_XCONTR												// Nome do Contribuinte 


ElseIf SEA->EA_MODELO $ "18" //18-PAGAMENTO DE TRIBUTOS - DARF SIMPLES   - SEGMENTO N3

   //Não definido
   
ElseIf SEA->EA_MODELO $ "22" //22-PAGAMENTO DE TRIBUTOS - GARE ICMS SP - SEGMENTO N4

	cPerApu := DTOS(SE2->E2_E2_XAPUR)
	cValImp	:= StrZero(((SE2->E2_SALDO+SE2->E2_ACRESC+SE2->E2_MULTA+SE2->E2_JUROS)*100),14)
	cDatArr	:= DToS(SE2->E2_VENCREA)
	
	cReturn := '05'															// Identificação do tributo - 05 -> ICMS
	cReturn += StrZero(Val(SE2->E2_XCODREC),4)								// Código da Receita -> campo  no Protheus 
	cReturn += '1'															// Tipo de identificação do Contribuinte - 1: CNPJ; 2: CEI
	cReturn += iIf(Empty(SE2->E2_XCNPJC),Space(14),AllTrim(SE2->E2_XCNPJC))		// Identificação do Contribuiente
	cReturn += PadL(SM0->M0_INSC,12)										// Inscrição Estadual
	cReturn += SUBSTR(SE2->E2_XDIVID,1,13)									// Divida Ativa/No Etiqueta
	cReturn += SUBSTR(cPerApu,5,2)+SUBSTR(cPerApu,1,4)						// Mes e ano de referencia -> MMAAAA
	cReturn += PadL(SE2->E2_XPARCE,13,'0')                                  //Parcelamento
	cReturn += StrZero(((SE2->E2_SALDO)*100),14)							// Valor da Receita
	cReturn += StrZero(((SE2->E2_JUROS)*100),14)							// Valor dos Juros
	cReturn += StrZero(((SE2->E2_MULTA)*100),14)							// Valor das Multas
	cReturn += cValImp														// Valor do pagamento
	cReturn += SubStr(cDatArr,7,2)+SubSTr(cDatArr,5,2)+SubStr(cDatArr,1,4)	// Data Vencimento
	cReturn += SubStr(cDatArr,7,2)+SubSTr(cDatArr,5,2)+SubStr(cDatArr,1,4)	// Data Pagamento
	cReturn += Space(11)													// Branco
	cReturn += SE2->E2_XCONTR												// Nome do Contribuinte
	
ElseIf SEA->EA_MODELO $ "35" //35-FGTS – GRF/GRRF/GRDE - SEGMENTO N7

	cValImp	:= StrZero(((SE2->E2_SALDO+SE2->E2_ACRESC+SE2->E2_MULTA+SE2->E2_JUROS)*100),14)
	cDatArr	:= DToS(SE2->E2_VENCREA)
	
	cReturn := '11'															// Identificação do tributo - 11=FGTS-GFIP
	cReturn += STRZERO(VAL(SE2->E2_XCODREC),4)								// Código da Receita -> campo  no Protheus 
	cReturn += '1'															// Tipo de identificação do Contribuinte - 1: CNPJ; 2: CEI
	cReturn += IIF(Empty(SE2->E2_XCNPJC),Space(14),AllTrim(SE2->E2_XCNPJC))	// Identificação do Contribuiente
	cReturn += PADR(SE2->E2_CODBAR,48)										// CODIGO DE BARRAS
	cReturn += PADL(RTRIM(SE2->E2_XDIVID),16,'0')						        // IDENTIFICADOR DO FGTS
	cReturn += PADL('',09,'0')										        // LACRE DE CONECTIVIDADE SOCIAL
	cReturn += PADL('',02,'0')										        // DIGITO DO LACRE DE CONECTIVIDADE SOC.
	cReturn += PADR(SE2->E2_XCONTR,30)									    // NOME DO CONTRIBUINTE
	cReturn += SubStr(cDatArr,7,2)+SubSTr(cDatArr,5,2)+SubStr(cDatArr,1,4)	// DATA DO PAGAMENTO
	cReturn += cValImp							                            // VALOR DO PAGAMENTO 
	cReturn += Space(30)													// BRANCOS COMPLEMENTO DE REGISTRO
  
EndIf                         

Return(cReturn)     