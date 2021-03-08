#INCLUDE "PROTHEUS.CH" 
#INCLUDE "RWMAKE.CH"
 
/*
=====================================================================================
Programa............: SPDPIS07
Autor...............: Elton Zaniboni
Data................: 26/04/2019
Descrição / Objetivo: SPDPIS07 - Trazer CONTA contábil de Doc Entrada quando há rateio
                      da tabela SDE.
Solicitante.........: Rosimeire
Uso.................: Victorinox
Obs.................: http://tdn.totvs.com/pages/releaseview.action?pageId=307833861
=====================================================================================
*/

User Function SPDPIS07()

Local	cFilial		:=	PARAMIXB[1]	//FT_FILIAL
Local	cTpMov		:=	PARAMIXB[2]	//FT_TIPOMOV
Local	cSerie		:=	PARAMIXB[3]	//FT_SERIE
Local	cDoc		:=	PARAMIXB[4]	//FT_NFISCAL
Local	cClieFor	:=	PARAMIXB[5]	//FT_CLIEFOR
Local	cLoja		:=	PARAMIXB[6]	//FT_LOJA
Local	cItem		:=	PARAMIXB[7]	//FT_ITEM
Local	cProd		:=	PARAMIXB[8]	//FT_PRODUTO	 	
Local cRateio		:= 	GETADVFVAL("SD1","D1_RATEIO",PARAMIXB[1]+cDoc+cSerie+cClieFor,1)      
Local cConta		:= 	GETADVFVAL("SFT","FT_CONTA",PARAMIXB[1]+cTpMov+cSerie+cDoc+cClieFor+cLoja+cItem+cProd,1)

If cTpMov == "E" .AND. cRateio	== "1"

	DbSelectArea("SDE")
	DbSetOrder(1)                                
	
		If DbSeek(xFilial("SDE") + cDoc + cSerie + cClieFor + cLoja)  //Busca a Conta Contábil do primeiro Item da tabela de Rateio (SDE)
   		cConta := SDE->DE_CONTA
		EndIf

EndIf
          
Return cConta