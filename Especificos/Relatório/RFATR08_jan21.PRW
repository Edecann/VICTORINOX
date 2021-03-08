#Include "TopConn.ch"   
#include "tbiconn.ch"
#DEFINE ENTER CHR(13) + CHR(10)
/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³RFATR04   º Autor ³Cristian Gutierrerz º Data ³  07/12/18   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDescricao ³Relatorio de itens de notas fiscais de faturamento com      º±±
±±º          ³grupo especifico do cliente                                 º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³Uso Exclusivo Cliente Victorinox                            º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
User Function RFATR08()
//ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ//
//³ Declaracao de Variaveis                                             ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
Local aMensagem := {}
Local aBotoes   := {}

Private cCadastro := "Faturamento por Cliente em Excel"

	AAdd( aMensagem, "Geracao de itens faturados por cliente x grupos" )
	AAdd( aMensagem, "Ao final será aberto o Ms-Excel com este arquivo gerado." )
	
	AAdd( aBotoes, { 01, .T., { || fParam(), FechaBatch() } } )
	AAdd( aBotoes, { 22, .T., { || FechaBatch() } } )

	//FormBatch(<cTitulo>,<aMensagem>,<aBotoes>,<bValid>,nAltura,nLargura)
	FormBatch( cCadastro, aMensagem, aBotoes, , 160 )
Return
/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³PARAM     º Autor ³Cristian Gutierrerz º Data ³  07/12/18   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDescricao ³Solicitacao de Parametros ao usuarios						     º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³Uso Exclusivo Cliente Victorinox                            º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static Function fParam()
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Declaracao de Variaveis                                             ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
Local aPar := {}
Local aRet := {}
Local bOk  := {|| MsgYesNo( "Confirma o processamento?", cCadastro ) }
Local bValid := {|| Iif(ApOleClient("MsExcel"),.T.,(MsgInfo("MsExcel não instalado"),)) }
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Parametros					                                             ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
AAdd( aPar, { 1, "Cliente de " 	,Space( TamSx3( "A1_COD" 	)[1] ) 	,"","","SA1"	,"",TamSx3( "A1_COD" 	)[1],.T.})
AAdd( aPar, { 1, "Loja de "		,Space( TamSx3( "A1_LOJA"	)[1] ) 	,"","",""		,"",TamSx3( "A1_LOJA"	)[1],.T.}) 
AAdd( aPar, { 1, "Cliente ate "	,Space( TamSx3( "A1_COD" 	)[1] ) 	,"","","SA1"	,"",TamSx3( "A1_COD" 	)[1],.T.})
AAdd( aPar, { 1, "Loja ate "	,Space( TamSx3( "A1_LOJA"	)[1] ) 	,"","",""		,"",TamSx3( "A1_LOJA"	)[1],.T.}) 
AAdd( aPar, { 1, "Emissao de " 	,Space( TamSx3( "D2_EMISSAO")[1] ) ,"","",""		,"",TamSx3( "D2_EMISSAO")[1],.T.})
AAdd( aPar, { 1, "Emissao ate "	,Space( TamSx3( "D2_EMISSAO")[1] ) ,"","",""		,"",TamSx3( "D2_EMISSAO")[1],.T.}) 
AAdd( aPar, { 2, "Tipo Custo"		,				 								,{"1-Custo Medio" , "2-Custo NF" } ,60, ,.T.}) 
AAdd( aPar, { 5, "Ao final do processamento abrir o Ms-Excel?",.F.,150,bValid,.F.})
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Tela para cooleta de informacoes                                    ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
If ParamBox( aPar, "Parâmetros de processamento", @aRet, bOk, , , , , , , .F., .F. )
	
	Processa( {|| GerarArq( aRet ) }, cCadastro, "Processando arquivo, aguarde...", .F. )
Endif

Return
/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³GeraArq   º Autor ³Cristian Gutierrerz º Data ³  07/12/18   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDescricao ³Processamento e geração do arquivo Excel					     º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³Uso Exclusivo Cliente Victorinox                            º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static Function GerarArq( aRet )
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Declaracao de Variaveis                                             ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
Local oFwMsEx 		:= NIL
Local cArq 			:= ""
Local cDir 			:= GetSrvProfString("Startpath","")
Local cWorkSheet 	:= ""
Local cDirTmp 		:= GetTempPath()
Local cQuery		:= ""                         
Local nCount		:= 0                         
Local cCFOP1        := '5910/6910' //SuperGetMV( "ES_CFOPEX1" , .F. , '5102','5104','5106','5110','5111','5112','5114','5115','5117','5119','5120','5123','5401','5402','5403','5405','5551' )
//Local cCFOP2		:= //SuperGetMV( "ES_CFOPEX2" , .F. , '6102','6104','6106','6108','6110','6111','6112','6114','6115','6117','6118','6119','6120','6123','6129','6401','6402' )
//Local cCFOP3		:= //SuperGetMV( "ES_CFOPEX3" , .F. , '6403','6404','6551','5101','6101','6107','5933','6933' )
Local cCFEN1        := '1910/2910' //SuperGetMV( "ES_CFOPEN1" , .F. , '1202', '1411', '2202', '2411' , '2204' )
Local cDesCan		:= ""                 
Local cF2Vend1		:= ""
Local cA3Nome		:= ""              
Local nCusto		:= 0
Local Ycan
Local xCan 
Local YDesCan := ""
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Declaracao de Variaveis                                             ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Monta query para selecao dos registros a serem gerados no Excel     ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
cQuery := "SELECT A1_COD, A1_LOJA, A1_NOME, A1_CGC, A1_REGIAO, A1_END, A1_MUN, A1_EST, A1_GRPTRIB, A1_NREDUZ,	"+ENTER
cQuery += "   		A1_DTCAD,'' AS C5_XSEGMEN ,'' AS ZV5_DESCC1, '' AS ZV5_DESCC2, '' AS ZV5_DESCC3,'' AS ZV5_DESCC4,'' AS ZV5_DESCC5, '' AS C5_XCANAIS,	"+ENTER
cQuery += "   		A1_DTCAD,A1_XSEGMEN, A1_DESCC1, A1_DESCC2, A1_DESCC3, A1_DESCC4, A1_DESCC5, A1_XCANAIS,	'' AS DESCAN,	"+ENTER
cQuery += "			A1_TEL, A1_EMAIL, A1_ULTCOM,																						"+ENTER
cQuery += "			D2_PEDIDO,D2_DOC, D2_SERIE, D2_EMISSAO, D2_COD, D2_QUANT, D2_PRCVEN,  D2_TOTAL,	D2_CF,					"+ENTER
cQuery += "			D2_VALBRUT,(D2_VALBRUT)-(D2_VALIPI+D2_VALICM+D2_VALIMP5+D2_VALIMP6+D2_ICMSRET)AS NET_SALLES, D2_CUSTO1, 														"+ENTER
cQuery += "			F2_VEND1, A3_NOME, BM_DESC AS D2_DSGRUPO, B2_CM1,							"+ENTER
cQuery += "			D2_VALIPI, D2_VALICM, D2_DESCON, D2_VALIMP5, D2_VALIMP6,D2_ICMSRET												"+ENTER
cQuery += "FROM "+ RetSqlName( "SA1" ) +" SA1															 							"+ENTER
cQuery += "INNER JOIN "+ RetSqlName( "SD2" ) +" SD2 																				"+ENTER
cQuery += "   ON SD2.D_E_L_E_T_	= ''			 																						"+ENTER
cQuery += "  AND SD2.D2_FILIAL	= '"+ xFilial( "SD2") +"'																		"+ENTER
cQuery += "  AND SD2.D2_CLIENTE	= SA1.A1_COD																						"+ENTER
cQuery += "  AND SD2.D2_LOJA		= SA1.A1_LOJA														 								"+ENTER
cQuery += "  AND SD2.D2_EMISSAO	BETWEEN '"+ DtoS( CtoD( aRet[5] ) )+"' AND '"+ DtoS( CtoD( aRet[6] ) )+"'	"+ENTER
cQuery += "  AND ( SD2.D2_CF		IN ( "+ cCFOP1 +" )																				"+ENTER
//cQuery += "   		 OR  SD2.D2_CF	IN ( "+ cCFOP2 +" )																				"+ENTER
//cQuery += "   		 OR  SD2.D2_CF	IN ( "+ cCFOP3 +" )																				"+ENTER
cQuery += "   		)																															"+ENTER
cQuery += "  AND SD2.D2_TIPO		= 'N'																									"+ENTER
cQuery += "INNER JOIN "+ RetSqlName( "SF2" ) +" SF2 																				"+ENTER
cQuery += "   ON SF2.D_E_L_E_T_	= ''			 																						"+ENTER
cQuery += "  AND SF2.F2_FILIAL	= '"+ xFilial( "SF2") +"'																		"+ENTER
cQuery += "  AND SF2.F2_CLIENTE	= SD2.D2_CLIENTE																					"+ENTER
cQuery += "  AND SF2.F2_LOJA		= SD2.D2_LOJA														 								"+ENTER
cQuery += "  AND SF2.F2_DOC		= SD2.D2_DOC																						"+ENTER
cQuery += "  AND SF2.F2_SERIE		= SD2.D2_SERIE														 								"+ENTER
cQuery += "INNER JOIN "+ RetSqlName( "SFT" ) +" SFT 																				"+ENTER
cQuery += "   ON SFT.D_E_L_E_T_	= ''			 																						"+ENTER
cQuery += "  AND SFT.FT_FILIAL	= '"+ xFilial( "SFT") +"'																		"+ENTER
cQuery += "  AND SFT.FT_CLIEFOR	= SD2.D2_CLIENTE																					"+ENTER
cQuery += "  AND SFT.FT_LOJA		= SD2.D2_LOJA														 								"+ENTER
cQuery += "  AND SFT.FT_NFISCAL	= SD2.D2_DOC																						"+ENTER
cQuery += "  AND SFT.FT_SERIE		= SD2.D2_SERIE														 								"+ENTER
cQuery += "  AND SFT.FT_CFOP		= SD2.D2_CF 														 								"+ENTER
cQuery += "  AND SFT.FT_ITEM		= SD2.D2_ITEM 														 								"+ENTER
cQuery += "  AND SFT.FT_PRODUTO	= SD2.D2_COD 														 								"+ENTER
cQuery += "  AND SFT.FT_DTCANC	= '' 														 											"+ENTER
cQuery += "INNER JOIN "+ RetSqlName( "SB1" ) +" SB1 																				"+ENTER
cQuery += "   ON SB1.D_E_L_E_T_	= ''                                			                                 "+ENTER
cQuery += "  AND SB1.B1_FILIAL	= '"+ xFilial( "SB1") +"'																		"+ENTER
cQuery += "  AND SB1.B1_COD		= SD2.D2_COD																						"+ENTER
cQuery += "LEFT JOIN "+ RetSqlName( "SBM" ) +" SBM 																				"+ENTER
cQuery += "   ON SBM.D_E_L_E_T_	= ''                                			                                 "+ENTER
cQuery += "  AND SBM.BM_FILIAL	= '"+ xFilial( "SBM") +"'																		"+ENTER
cQuery += "  AND SBM.BM_GRUPO		= SB1.B1_GRUPO																						"+ENTER
cQuery += "LEFT JOIN "+ RetSqlName( "SA3" ) +" SA3 																				"+ENTER
cQuery += "   ON SA3.D_E_L_E_T_	= ''                                			                                 "+ENTER
cQuery += "  AND SA3.A3_FILIAL	= '"+ xFilial( "SA3" ) +"'																		"+ENTER
cQuery += "  AND SA3.A3_COD		= SF2.F2_VEND1																						"+ENTER
cQuery += "LEFT JOIN "+ RetSqlName( "SB2" ) +" SB2 																				"+ENTER
cQuery += "   ON SB2.D_E_L_E_T_	= ''                                			                                 "+ENTER
cQuery += "  AND SB2.B2_FILIAL	= '"+ xFilial( "SB2" ) +"'																		"+ENTER
cQuery += "  AND SB2.B2_COD		= SD2.D2_COD																						"+ENTER
cQuery += "  AND SB2.B2_LOCAL		= SD2.D2_LOCAL																						"+ENTER
cQuery += "WHERE SA1.D_E_L_E_T_	= ''																									"+ENTER
cQuery += "  AND SA1.A1_FILIAL	= '"+ xFilial( "SA1" ) +"'																		"+ENTER
cQuery += "  AND SA1.A1_COD + SA1.A1_LOJA	BETWEEN '"+ aRet[1] + aRet[2] +"' AND '" + aRet[3] + aRet[4] +"'		"+ENTER
           
cQuery += "UNION ALL 																														"+ENTER

cQuery += "SELECT A1_COD, A1_LOJA, A1_NOME, A1_CGC, A1_REGIAO, A1_END, A1_MUN, A1_EST, A1_GRPTRIB, A1_NREDUZ,	"+ENTER
cQuery += "   		A1_DTCAD,'' AS C5_XSEGMEN ,'' AS ZV5_DESCC1, '' AS ZV5_DESCC2, '' AS ZV5_DESCC3,'' AS ZV5_DESCC4,'' AS ZV5_DESCC5, '' AS C5_XCANAIS,"+ENTER
cQuery += "   		A1_DTCAD,A1_XSEGMEN, A1_DESCC1, A1_DESCC2, A1_DESCC3, A1_DESCC4, A1_DESCC5, A1_XCANAIS,	'' AS DESCAN,	"+ENTER
cQuery += "			A1_TEL, A1_EMAIL, A1_ULTCOM,																						"+ENTER
cQuery += "			D1_PEDIDO AS D2_PEDIDO,D1_DOC AS D2_DOC, D1_SERIE AS D2_SERIE, D1_EMISSAO AS D2_EMISSAO, D1_COD AS D2_COD,				"+ENTER
cQuery += "			ROUND( D1_QUANT*-1 ,2 ) AS D2_QUANT, ROUND( D1_VUNIT*-1 , 2 ) AS D2_PRCVEN, 						 "+ENTER
cQuery += "			ROUND( D1_TOTAL*-1 ,2 ) AS D2_TOTAL,	                                                             "+ENTER
cQuery += "			D1_CF AS D2_CF, 																			         "+ENTER
cQuery += "			ROUND( FT_VALCONT*-1 ,2 )  AS D2_VALBRUT,ROUND( FT_VALCONT*-1 ,2 )-(D1_VALIPI+D1_VALICM+D1_VALIMP5+D1_VALIMP6+D1_ICMSRET)*-1 AS NET_SALLES,			"+ENTER
cQuery += "			( D1_CUSTO*-1  ) AS D2_CUSTO1,	F2_VEND1 AS F2_VEND1, A3_NOME, BM_DESC AS D2_DSGRUPO, ( B2_CM1*-1) AS B2_CM1,		"+ENTER
cQuery += "			D1_VALIPI*-1 AS D2_VALIPI, D1_VALICM*-1 AS D2_VALICM, D1_VALDESC*-1 AS D2_DESCON,					     "+ENTER
cQuery += "			D1_VALIMP5*-1 AS D2_VALIMP5, D1_VALIMP6*-1 AS D2_VALIMP6,D1_ICMSRET*-1 AS D2_ICMSRET 				 "+ENTER
cQuery += "FROM "+ RetSqlName( "SA1" ) +" SA1															 							"+ENTER
cQuery += "INNER JOIN "+ RetSqlName( "SD1" ) +" SD1 																				"+ENTER
cQuery += "   ON SD1.D_E_L_E_T_	= ''			 																						"+ENTER
cQuery += "  AND SD1.D1_FILIAL	= '"+ xFilial( "SD1") +"'																		"+ENTER
cQuery += "  AND SD1.D1_FORNECE	= SA1.A1_COD																						"+ENTER
cQuery += "  AND SD1.D1_LOJA		= SA1.A1_LOJA														 								"+ENTER
cQuery += "  AND SD1.D1_CF		IN ( "+ cCFEN1 +" )																					"+ENTER
cQuery += "INNER JOIN "+ RetSqlName( "SFT" ) +" SFT 																				"+ENTER
cQuery += "   ON SFT.D_E_L_E_T_	= ''			 																						"+ENTER
cQuery += "  AND SFT.FT_FILIAL	= '"+ xFilial( "SFT") +"'																		"+ENTER
cQuery += "  AND SFT.FT_TIPOMOV	= 'E'																									"+ENTER
cQuery += "  AND SFT.FT_SERIE		= SD1.D1_SERIE														 								"+ENTER
cQuery += "  AND SFT.FT_NFISCAL	= SD1.D1_DOC																						"+ENTER
cQuery += "  AND SFT.FT_CLIEFOR	= SD1.D1_FORNECE																					"+ENTER
cQuery += "  AND SFT.FT_LOJA		= SD1.D1_LOJA														 								"+ENTER
cQuery += "  AND SFT.FT_ITEM		= SD1.D1_ITEM 														 								"+ENTER
cQuery += "  AND SFT.FT_PRODUTO	= SD1.D1_COD 														 								"+ENTER
cQuery += "  AND SFT.FT_ENTRADA	BETWEEN '"+ DtoS( CtoD( aRet[5] ) )+"' AND '"+ DtoS( CtoD( aRet[6] ) )+"'	"+ENTER
cQuery += "  AND SFT.FT_DTCANC	= '' 														 											"+ENTER
cQuery += "LEFT JOIN "+ RetSqlName( "SF2" ) +" SF2 																				"+ENTER
cQuery += "   ON SF2.D_E_L_E_T_	= ''			 																						"+ENTER
cQuery += "  AND SF2.F2_FILIAL	= '"+ xFilial( "SF2") +"'																		"+ENTER
cQuery += "  AND SF2.F2_CLIENTE	= SD1.D1_FORNECE																					"+ENTER
cQuery += "  AND SF2.F2_LOJA		= SD1.D1_LOJA														 								"+ENTER
cQuery += "  AND SF2.F2_DOC		= SD1.D1_NFORI																						"+ENTER
cQuery += "  AND SF2.F2_SERIE		= SD1.D1_SERIORI													 								"+ENTER
cQuery += "INNER JOIN "+ RetSqlName( "SB1" ) +" SB1 																				"+ENTER
cQuery += "   ON SB1.D_E_L_E_T_	= ''                                			                                 "+ENTER
cQuery += "  AND SB1.B1_FILIAL	= '"+ xFilial( "SB1") +"'																		"+ENTER
cQuery += "  AND SB1.B1_COD		= SD1.D1_COD																						"+ENTER
cQuery += "LEFT JOIN "+ RetSqlName( "SBM" ) +" SBM 																				"+ENTER
cQuery += "   ON SBM.D_E_L_E_T_	= ''                                			                                 "+ENTER
cQuery += "  AND SBM.BM_FILIAL	= '"+ xFilial( "SBM") +"'																		"+ENTER
cQuery += "  AND SBM.BM_GRUPO		= SB1.B1_GRUPO																						"+ENTER
cQuery += "LEFT JOIN "+ RetSqlName( "SA3" ) +" SA3 																				"+ENTER
cQuery += "   ON SA3.D_E_L_E_T_	= ''                                			                                 "+ENTER
cQuery += "  AND SA3.A3_FILIAL	= '"+ xFilial( "SA3" ) +"'																		"+ENTER
cQuery += "  AND SA3.A3_COD		= SF2.F2_VEND1																						"+ENTER
cQuery += "LEFT JOIN "+ RetSqlName( "SB2" ) +" SB2 																				"+ENTER
cQuery += "   ON SB2.D_E_L_E_T_	= ''                                			                                 "+ENTER
cQuery += "  AND SB2.B2_FILIAL	= '"+ xFilial( "SB2" ) +"'																		"+ENTER
cQuery += "  AND SB2.B2_COD		= SD1.D1_COD																						"+ENTER
cQuery += "  AND SB2.B2_LOCAL		= SD1.D1_LOCAL																						"+ENTER
cQuery += "WHERE SA1.D_E_L_E_T_	= ''																									"+ENTER
cQuery += "  AND SA1.A1_FILIAL	= '"+ xFilial( "SA1" ) +"'																		"+ENTER
cQuery += "  AND SA1.A1_COD + SA1.A1_LOJA	BETWEEN '"+ aRet[1] + aRet[2] +"' AND '" + aRet[3] + aRet[4] +"'		"+ENTER
cQuery += "ORDER BY A1_COD, A1_LOJA, D2_COD, D2_DOC, D2_EMISSAO																"+ENTER
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Executa Query para geracao do arquivo										   ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
If Select( "TMP" ) > 0 
	dbSelectArea( "TMP" )
	dbCloseArea()
EndIf     
Memowrite( "c:\temp\Victorinox_rfatr04.txt" , cQuery )
TcQuery cQuery New Alias "TMP"
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Contagem de Registros para alimentar regua de Processamento			   ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
dbSelectArea( "TMP" )
dbGoTop()
While !TMP->( EOF() )
	nCount ++

	TMP->( dbSkip() )
EndDo
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Cria Excel																			   ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
cWorkSheet 	:= "FAT_CLI_GRP"
cTable     	:= "ITENS"
oFwMsEx 		:= FWMsExcel():New()

oFwMsEx:AddWorkSheet( cWorkSheet )
oFwMsEx:AddTable( cWorkSheet, cTable )	

oFwMsEx:AddColumn( cWorkSheet, cTable , "Codigo"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Loja"    		, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Nome"			, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Nome Reduz."	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "CNPJ" 	 		, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Regiao"  		, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Endereco"		, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Municipio"  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Estado"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Telefone"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "E-mail"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "GrpTrib"		, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Cadastro"  	, 1,1)
//oFwMsEx:AddColumn( cWorkSheet, cTable , "Segmento"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Seg C1(Pedidos)"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Seg C2(Pedidos)"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Seg C3(Pedidos)"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Seg C4(Pedidos)"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Seg C5(Pedidos)"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Canais(Pedidos)"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Desc.Canais(Pedidos)"	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Seg C1(Clientes)"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Seg C2(Clientes)"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Seg C3(Clientes)"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Seg C4(Clientes)"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Seg C5(Clientes)"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Canais(Clientes)"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Desc.Canais(Clientes)"	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Nota"	  		, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Serie"	 	 	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Emissao"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Codigo"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Quantidade" 	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Prc.Unit."  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Total"		  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Total Bruto"	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Net Salles"	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Custo Unit."  	, 1,1) 
oFwMsEx:AddColumn( cWorkSheet, cTable , "Custo"		  	, 1,1)     
oFwMsEx:AddColumn( cWorkSheet, cTable , "CFOP"		  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Grupo"		  	, 1,1)     
oFwMsEx:AddColumn( cWorkSheet, cTable , "Cod Vend."  	, 1,1)     
oFwMsEx:AddColumn( cWorkSheet, cTable , "Nome"		  	, 1,1)     
oFwMsEx:AddColumn( cWorkSheet, cTable , "IPI"		  	, 1,1)     
oFwMsEx:AddColumn( cWorkSheet, cTable , "ICMS"		  	, 1,1)     
oFwMsEx:AddColumn( cWorkSheet, cTable , "Desconto"	  	, 1,1)     
oFwMsEx:AddColumn( cWorkSheet, cTable , "Cofins"	  	, 1,1)     
oFwMsEx:AddColumn( cWorkSheet, cTable , "PIS"			, 1,1) 
oFwMsEx:AddColumn( cWorkSheet, cTable , "ST"			, 1,1)      
oFwMsEx:AddColumn( cWorkSheet, cTable , "Ult. Compra"	, 1,1)     
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Alimenta regua de Processamento												   ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
ProcRegua( nCount )
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Loop principal da rotina 															³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
dbSelectArea( "TMP" )
TMP->( dbGoTop() )
While !( TMP->( EOF() ) )

	IncProc() 
	
	          
	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³Define descricao do canal															³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ 
/*
XSEG   := Posicione("SC5",1,xFilial("SC5")+TMP->D2_PEDIDO,"C5_XSEGMEN")	                  
IF EMPTY(XSEG) 
	IF ( Select( "_SA1" ) > 0 )
		_SA1->( dbCloseArea() )
	EndIF

	_cQry := " SELECT A1_COD,A1_LOJA,A1_XSEGMEN,A1_XCANAIS,A1_DESCC1,A1_DESCC2,A1_DESCC3,A1_DESCC4,A1_DESCC5 FROM "+ RetSqlName( "SA1" ) 
	_cQry += " WHERE A1_COD='"+ TMP->A1_COD +"' AND A1_LOJA='"+TMP->A1_LOJA+"' "
	_cQry := ChangeQuery(_cQry)

	//Executando consulta e setando o total da régua
	TCQuery _cQry New Alias "_SA1"
	Count to nI
_SA1->(DBGOTOP())
*/

YSEG   := TMP->A1_XSEGMEN
YCAN   := TMP->A1_XCANAIS
YSEGC1 := TMP->A1_DESCC1	
YSEGC2 := TMP->A1_DESCC2	
YSEGC3 := TMP->A1_DESCC3	
YSEGC4 := TMP->A1_DESCC4	
YSEGC5 := TMP->A1_DESCC5

//_SA1->( dbCloseArea() )

 YDesCan := ""
	If YCAN == "1"
		YDesCan := "VX OWN RETAIL"
	ElseIf YCAN == "2"
		YDesCan := "VX E-COMMERCE"
	ElseIf YCAN == "3"
		YDesCan := "DISTRIBUTION"
	ElseIf YCAN == "4"
		YDesCan := "DIRECT SALES"   
	ElseIf YCAN == "5"
		YDesCan := "WHOLESALE" 
	ElseIf YCAN == "6"
		YDesCan := "CORPORATE BUSINESS"   
	ElseIf YCAN == "7"
		YDesCan := "TRAVEL RETAIL"   
	Else
		YDesCan := "NoName"
	EndIf	

//ELSE
XSEG   := Posicione("SC5",1,xFilial("SC5")+TMP->D2_PEDIDO,"C5_XSEGMEN")	 
XCAN   := Posicione("SC5",1,xFilial("SC5")+TMP->D2_PEDIDO,"C5_XCANAIS")
XSEGC1 := Posicione("ZV5",6,xFilial("ZV5")+XSEG,"ZV5_DESCC1")	
XSEGC2 := Posicione("ZV5",6,xFilial("ZV5")+XSEG,"ZV5_DESCC2")	
XSEGC3 := Posicione("ZV5",6,xFilial("ZV5")+XSEG,"ZV5_DESCC3")	
XSEGC4 := Posicione("ZV5",6,xFilial("ZV5")+XSEG,"ZV5_DESCC4")	
XSEGC5 := Posicione("ZV5",6,xFilial("ZV5")+XSEG,"ZV5_DESCC5")
//ENDIF

   cDesCan := ""
	If XCAN == "1"
		cDesCan := "VX OWN RETAIL"
	ElseIf XCAN == "2"
		cDesCan := "VX E-COMMERCE"
	ElseIf XCAN == "3"
		cDesCan := "DISTRIBUTION"
	ElseIf XCAN == "4"
		cDesCan := "DIRECT SALES"   
	ElseIf XCAN == "5"
		cDesCan := "WHOLESALE" 
	ElseIf XCAN == "6"
		cDesCan := "CORPORATE BUSINESS"   
	ElseIf XCAN == "7"
		cDesCan := "TRAVEL RETAIL"   
	Else
		cDesCan := "NoName"
	EndIf	
	
	 
	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³Define custo a ser impresso														³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	nCusto := 0
	If Left( aRet[7] , 1 )  == "1"
		nCusto := TMP->B2_CM1
	Else
		nCusto := TMP->D2_CUSTO1
	EndIf		         

	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³Caso nao encontre Vendedor na NF, usa o default do cliente				³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ   
	cF2Vend1 := ""
	cA3Nome	:= ""
	 
	If Empty( TMP->F2_VEND1 )
		dbSelectArea( "SA1" )
		dbSetOrder(1)
		dbSeek( xFilial( "SA1" ) + TMP->A1_COD + TMP->A1_LOJA )
			cF2Vend1 := SA1->A1_VEND
			
			If !Empty( cF2Vend1 )
				dbSelectArea( "SA3" )
				dbSetOrder(1)
				dbSeek( xFilial( "SA3" ) + cF2Vend1 )
					cA3Nome := SA3->A3_NOME
			EndIf		
	Else
		cF2Vend1 := TMP->F2_VEND1
		cA3Nome	:= TMP->A3_NOME
	EndIf               
	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³Adiciona registros nas linhas da tabela										³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	
	oFwMsEx:AddRow( cWorkSheet, cTable, { TMP->A1_COD, TMP->A1_LOJA, TMP->A1_NOME, TMP->A1_NREDUZ, TMP->A1_CGC, TMP->A1_REGIAO, TMP->A1_END, TMP->A1_MUN, TMP->A1_EST, TMP->A1_TEL, TMP->A1_EMAIL, TMP->A1_GRPTRIB,	TMP->A1_DTCAD,XSEGC1, XSEGC2, XSEGC3, XSEGC4, XSEGC5, XCAN, cDesCan,YSEGC1, YSEGC2, YSEGC3, YSEGC4, YSEGC5,YCAN,YDesCan ,TMP->D2_DOC, TMP->D2_SERIE, TMP->D2_EMISSAO, TMP->D2_COD, TMP->D2_QUANT, TMP->D2_PRCVEN, TMP->D2_TOTAL, TMP->D2_VALBRUT,TMP->NET_SALLES,nCusto, nCusto*TMP->D2_QUANT, TMP->D2_CF, TMP->D2_DSGRUPO, cF2Vend1, cA3Nome, TMP->D2_VALIPI, TMP->D2_VALICM, TMP->D2_DESCON, TMP->D2_VALIMP5, TMP->D2_VALIMP6,TMP->D2_ICMSRET, DtoC( StoD( TMP->A1_ULTCOM ) ) } )
	
	TMP->( dbSkip() )
//_SD3->( dbCloseArea() )
EndDo
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Geracao do Arquivo																	³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
oFwMsEx:Activate()

cArq := CriaTrab( NIL, .F. ) + ".xml"

LjMsgRun( "Gerando o arquivo, aguarde...", cCadastro, {|| oFwMsEx:GetXMLFile( cArq ) } )

If __CopyFile( cArq, cDirTmp + cArq )

	If aRet[8]
		oExcelApp := MsExcel():New()
		oExcelApp:WorkBooks:Open( cDirTmp + cArq )
		oExcelApp:SetVisible(.T.)
	Else
		MsgInfo( "Arquivo " + cArq + " gerado com sucesso no diretório " + cDir )
	Endif
Else
	MsgInfo( "Arquivo não copiado para temporário do usuário." )
Endif

Return

/*
Exemplo:
	Local oExcel := FWMSEXCEL():New()

	oExcel:AddworkSheet("Teste - 1")

		oExcel:AddTable ("Teste - 1","Titulo de teste 1")
			oExcel:AddColumn("Teste - 1","Titulo de teste 1","Col1",1,1)
			oExcel:AddColumn("Teste - 1","Titulo de teste 1","Col2",2,2)
			oExcel:AddColumn("Teste - 1","Titulo de teste 1","Col3",3,3)
			oExcel:AddColumn("Teste - 1","Titulo de teste 1","Col4",1,1)

			oExcel:AddRow("Teste - 1","Titulo de teste 1",{11,12,13,14})
			oExcel:AddRow("Teste - 1","Titulo de teste 1",{21,22,23,24})
			oExcel:AddRow("Teste - 1","Titulo de teste 1",{31,32,33,34})
			oExcel:AddRow("Teste - 1","Titulo de teste 1",{41,42,43,44})
	
	oExcel:AddworkSheet("Teste - 2")

		oExcel:AddTable("Teste - 2","Titulo de teste 1")
			oExcel:AddColumn("Teste - 2","Titulo de teste 1","Col1",1)
			oExcel:AddColumn("Teste - 2","Titulo de teste 1","Col2",2)
			oExcel:AddColumn("Teste - 2","Titulo de teste 1","Col3",3)
			oExcel:AddColumn("Teste - 2","Titulo de teste 1","Col4",1)

			oExcel:AddRow("Teste - 2","Titulo de teste 1",{11,12,13,stod("20121212")})
			oExcel:AddRow("Teste - 2","Titulo de teste 1",{21,22,23,stod("20121212")})
			oExcel:AddRow("Teste - 2","Titulo de teste 1",{31,32,33,stod("20121212")})
			oExcel:AddRow("Teste - 2","Titulo de teste 1",{41,42,43,stod("20121212")})
			oExcel:AddRow("Teste - 2","Titulo de teste 1",{51,52,53,stod("20121212")})

	oExcel:Activate()
	
	oExcel:GetXMLFile("teste.xml")
	MsgInfo("Arquivo texte.xml gerado com sucesso no \system\")
======================================================================================
FWMsExcel
	Classe: FWMsExcel

Descrição
	Fornece metodos para criação de arquivo XML compatível com as versões do MS Excel 2003, superiores e alternativos. 
	Fornece metodos para a criação de planilhas em formato de tabela e fornece também métodos para a configuração de cores e fonte.

Método: New
Sintaxe: FWMsExcel():New()-> NIL
Descrição: Método construtor da classe

Método: ClassName
Sintaxe: FWMsExcel():ClassName()-> NIL
Descrição: Fornece o nome da classe

Método: Activate
Sintaxe: FWMsExcel():Activate()-> NIL
Descrição: Habilita o uso da classe, indicando que esta configurada e pronto para uso

Método: DeActivate
Sintaxe: FWMsExcel():DeActivate()-> NIL
Descrição: Desabilita o uso da classe

Método: AddWorkSheet
Sintaxe: FWMsExcel():AddWorkSheet(< cWorkSheet >)-> NIL
Descrição: Adiciona uma Worksheet ( Planilha )
Parâmetros: 	cWorkSheet -> Nome da planilha que será adicionada

Método: AddTable
Sintaxe: FWMsExcel():AddTable(< cWorkSheet >, < cTable >)-> NIL
Descrição: Adiciona uma tabela na Worksheet. Uma WorkSheet pode ter mais de uma tabela
Parâmetros: 	cWorkSheet -> Nome da planilha
					cTable -----> Titulo da tabela que será adicionada

Método: AddColumn
Sintaxe: FWMsExcel():AddColumn(< cWorkSheet >, < cTable >, < cColumn >, < nAlign >, < nFormat >, < lTotal >)-> NIL
Descrição: Adiciona uma coluna a tabela de uma Worksheet.
Parâmetros: 	cWorkSheet -> Nome da planilha
					cTable -----> Título da tabela
					cColumn ----> Titulo da tabela que será adicionada
					nAlign -----> Alinhamento da coluna ( 1-Left,2-Center,3-Right )
					nFormat ----> Codigo de formatação ( 1-General,2-Number,3-Monetário,4-DateTime )
					lTotal -----> Indica se a coluna deve ser totalizada

Método: AddRow
Sintaxe: FWMsExcel():AddRow(< cWorkSheet >, < cTable >, < aRow >)-> NIL
Descrição: Adiciona uma linha a tabela de uma Worksheet.
Parâmetros: 	cWorkSheet -> Nome da planilha
					cTable -----> Título da tabela
					aRow -------> Array com as informações da linha da linha
 
Método: GetXMLFile
Sintaxe: FWMsExcel():GetXMLFile(< cFile >)-> NIL
Descrição: Cria um arquivo no formato XML do MSExcel 2003 em diante
Parâmetros: 	cFile -> Nome do arquivo fisico
 
Método: IsWorkSheet
Sintaxe: FWMsExcel():IsWorkSheet(< cWorkSheet >)-> NIL
Descrição: Verifica se um nome já foi utilizado para definir uma WorkSheet
Parâmetros: 	cWorksheet -> Nome da planilha
 
Método: SetFontSize
Sintaxe: FWMsExcel():SetFontSize(< nFontSize >)-> NIL
Descrição: Define o tamanho para a fonte para todos os estilos da planilha
Parâmetros: 	nFontSize -> Tamanho da fonte 
 
Método: SetFont
Sintaxe: FWMsExcel():SetFont(< cFont >)-> NIL
Descrição: Define a fonte que será utilizada para todos os estilos da planilha
Parâmetros: 	cFont -> Nome da fonte
 
Método: SetItalic
Sintaxe: FWMsExcel():SetItalic(< lItalic >)-> NIL
Descrição: Define se a fonte terá a configuração "Italico" para todos os estilos da planilha
Parâmetros: 	lItalic -> Se verdadeiro o formato é itálico

Método: SetBold
Sintaxe: FWMsExcel():SetBold(< lBold >)-> NIL
Descrição: Define se a fonte terá a configuração "Negrito" para todos os estilos da planilha
Parâmetros: 	lBold -> Se verdadeiro o formato é negrito
 
Método: SetUnderLine
Sintaxe: FWMsExcel():SetUnderLine(< lUnderline >)-> NIL
Descrição: Define se a fonte terá a configuração "Sublinhado" para todos os estilos da planilha
Parâmetros: 	lUnderLine -> Se verdadeiro o formato é sublinhado
 
Método: SetFrGeneralColor
Sintaxe: FWMsExcel():SetFrGeneralColor(< cColor >)-> NIL
Descrição: Define a cor de preenchimento geral para todos os estilos da planilha
Parâmetros: 	cColor -> Cor em hexadecimal
 
Método: SetBgGeneralColor
Sintaxe: FWMsExcel():SetBgGeneralColor(< cColor >)-> NIL
Descrição: Define a cor da fonte do estilo do Titulo para todos os estilos da planilha
Parâmetros: 	cColor -> Cor em hexadecimal
 
Método: SetTitleFont
Sintaxe: FWMsExcel():SetTitleFont(< cFont >)-> NIL
Descrição: Define a fonte que será utilizada no estilo do Título
Parâmetros: 	cFont -> Nome da fonte

Método: SetTitleSizeFont
Sintaxe: FWMsExcel():SetTitleSizeFont(< nFontSize >)-> NIL
Descrição: Define o tamanho para a fonte do estilo do Título
Parâmetros: nFontSize -> Tamanho da fonte
 
Método: SetTitleItalic
Sintaxe: FWMsExcel():SetTitleItalic(< lItalic >)-> NIL
Descrição: Define se a fonte terá a configuração "Italico" no estilo do Titulo
Parâmetros: 	lItalic -> Se verdadeiro define o formato itálico
 
Método: SetTitleBold
Sintaxe: FWMsExcel():SetTitleBold(< lBold >)-> NIL
Descrição: Define se a fonte terá a configuração "Negrito" no estilo do Título
Parâmetros: 	lBold -> Se verdadeiro define o fomrato negrito
 
Método: SetTitleUnderLine
Sintaxe: FWMsExcel():SetTitleUnderLine(< lUnderline >)-> NIL
Descrição: Define se a fonte terá a configuração "Sublinhado" no estilo do Titulo
Parâmetros: 	lUnderLine -> Se verdadeiro define o formato como sublinhado

Método: SetTitleFrColor
Sintaxe: FWMsExcel():SetTitleFrColor(< cColor >)-> NIL
Descrição: Define a cor de preenchimento do estilo do Titulo
Parâmetros: 	cColor -> Cor em hexadecimal
 
Método: SetTitleBgColor
Sintaxe: FWMsExcel():SetTitleBgColor(< cColor >)-> NIL
Descrição: Define a cor da fonte do estilo do Titulo
Parâmetros: 	cColor -> Cor em hexadecimal
 
Método: SetHeaderFont
Sintaxe: FWMsExcel():SetHeaderFont(< cFont >)-> NIL
Descrição: Define a fonte que será utilizada no estilo do Cabeçalho
Parâmetros: 	cFont -> Nome da fonte
 
Método: SetHeaderSizeFont
Sintaxe: FWMsExcel():SetHeaderSizeFont(< nFontSize >)-> NIL
Descrição: Define o tamanho para a fonte do estilo do Cabeçalho
Parâmetros: 	nFontSize -> Tamanho
 
Método: SetHeaderItalic
Sintaxe: FWMsExcel():SetHeaderItalic(< lItalic >)-> NIL
Descrição: Define se a fonte terá a configuração "Italico" no estilo do Cabeçalho
Parâmetros: 	lItalic -> Verdadeiro indica que sim
 
Método: SetHeaderBold
Sintaxe: FWMsExcel():SetHeaderBold(< lBold >)-> NIL
Descrição: Define se a fonte terá a configuração "Negrito" no estilo do Cabeçalho
Parâmetros: 	lBold -> Verdadeiro indica que sim
 
Método: SetHeaderUnderLine
Sintaxe: FWMsExcel():SetHeaderUnderLine(< lUnderline >)-> NIL
Descrição: Define se a fonte terá a configuração "Sublinhado" no estilo do Cabeçalho
Parâmetros: 	lUnderLine -> Verdadeiro indica que sim

Método: SetFrColorHeader
Sintaxe: FWMsExcel():SetFrColorHeader(< cColor >)-> NIL
Descrição: Define a cor de preenchimento do estilo do Cabeçalho
Parâmetros: 	cColor -> Cor em hexadecimal
 
Método: SetBgColorHeader
Sintaxe: FWMsExcel():SetBgColorHeader(< cColor >)-> NIL
Descrição: Define a cor da fonte do estilo do Cabeçalho
Parâmetros: 	cColor -> Cor em hexadecimal
 
Método: SetLineFont
Sintaxe: FWMsExcel():SetLineFont(< cFont >)-> NIL
Descrição: Define a fonte que será utilizada no estilo da Linha
Parâmetros: 	cFont -> Nome da fonte

Método: SetLineSizeFont
Sintaxe: FWMsExcel():SetLineSizeFont(< nFontSize >)-> NIL
Descrição: Define o tamanho para a fonte do estilo da Linha
Parâmetros: nFontSize -> Tamanho
 
Método: SetLineItalic
Sintaxe: FWMsExcel():SetLineItalic(< lItalic >)-> NIL
Descrição: Define se a fonte terá a configuração "Italico" no estilo da Linha
Parâmetros: 	lItalic  -> .T. indica que sim
 
Método: SetLineBold
Sintaxe: FWMsExcel():SetLineBold(< lBold >)-> NIL
Descrição: Define se a fonte terá a configuração "Negrito" no estilo da Linha
Parâmetros: 	lBold -> .T. indica que sim

Método: SetLineUnderLine
Sintaxe: FWMsExcel():SetLineUnderLine(< lUnderline >)-> NIL
Descrição: Define se a fonte terá a configuração "Sublinhado" no estilo da Linha
Parâmetros: lUnderLine -> .T. indica que sim

Método: SetLineFrColor
Sintaxe: FWMsExcel():SetLineFrColor(< cColor >)-> NIL
Descrição: Define a cor de preenchimento do estilo da Linha
Parâmetros: 	cColor -> Cor em Hexadecimal 

Método: SetLineBgColor
Sintaxe: FWMsExcel():SetLineBgColor(< cColor >)-> NIL
Descrição: Define a cor da fonte do estilo da Linha
Parâmetros: 	cColor -> Cor em Hexadecimal 
 
Método: Set2LineFont
Sintaxe: FWMsExcel():Set2LineFont(< cFont >)-> NIL
Descrição: Define a fonte que será utilizada no estilo da Linha 2
Parâmetros: 	cFont -> Nome da Fonte

Método: Set2LineSizeFont
Sintaxe: FWMsExcel():Set2LineSizeFont(< nFontSize >)-> NIL
Descrição: Define o tamanho para a fonte do estilo da Linha 2
Parâmetros: 	nFontSize -> Tamanho
 
Método: Set2LineItalic
Sintaxe: FWMsExcel():Set2LineItalic(< lItalic >)-> NIL
Descrição: Define se a fonte terá a configuração "Italico" no estilo da Linha 2
Parâmetros: 	lItalic -> .T. indica que sim

Método: Set2LineBold
Sintaxe: FWMsExcel():Set2LineBold(< lBold >)-> NIL
Descrição: Define se a fonte terá a configuração "Negrito" no estilo da Linha 2
Parâmetros: 	lBold -> .T. indica que sim

Método: Set2LineUnderLine
Sintaxe: FWMsExcel():Set2LineUnderLine(< lUnderline >)-> NIL
Descrição: Define se a fonte terá a configuração "Sublinhado" no estilo da Linha 2
Parâmetros: 	lUndeLine -> .T. indica que sim
 
Método: Set2LineFrColor
Sintaxe: FWMsExcel():Set2LineFrColor(< cColor >)-> NIL
Descrição: Define a cor de preenchimento do estilo da Linha 2
Parâmetros: 	cColor -> Cor em Hexadecimal
 
Método: Set2LineBgColor
Sintaxe: FWMsExcel():Set2LineBgColor(< cColor >)-> NIL
Descrição: Define a cor da fonte do estilo da Linha 2
Parâmetros: 	cColor -> Cor em Hexadecimal
*/