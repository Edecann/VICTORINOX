#Include "Protheus.ch"
#Include "TopConn.ch"   
#include "tbiconn.ch"
#DEFINE ENTER CHR(13) + CHR(10)
/*���������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �RFATR04   � Autor �Cristian Gutierrerz � Data �  07/12/18   ���
�������������������������������������������������������������������������͹��
���Descricao �Relatorio de itens de notas fiscais de faturamento com      ���
���          �grupo especifico do cliente                                 ���
�������������������������������������������������������������������������͹��
���Uso       �Uso Exclusivo Cliente Victorinox                            ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
���������������������������������������������������������������������������*/
User Function RFATR04()
//���������������������������������������������������������������������Ŀ
//� Declaracao de Variaveis                                             �
//�����������������������������������������������������������������������
Local aMensagem := {}
Local aBotoes   := {}

Private cCadastro := "Faturamento por Cliente em Excel"

	AAdd( aMensagem, "Geracao de itens faturados por cliente x grupos" )
	AAdd( aMensagem, "Ao final ser� aberto o Ms-Excel com este arquivo gerado." )
	
	AAdd( aBotoes, { 01, .T., { || fParam(), FechaBatch() } } )
	AAdd( aBotoes, { 22, .T., { || FechaBatch() } } )

	//FormBatch(<cTitulo>,<aMensagem>,<aBotoes>,<bValid>,nAltura,nLargura)
	FormBatch( cCadastro, aMensagem, aBotoes, , 160 )
Return
/*���������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �PARAM     � Autor �Cristian Gutierrerz � Data �  07/12/18   ���
�������������������������������������������������������������������������͹��
���Descricao �Solicitacao de Parametros ao usuarios						     ���
�������������������������������������������������������������������������͹��
���Uso       �Uso Exclusivo Cliente Victorinox                            ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
���������������������������������������������������������������������������*/
Static Function fParam()
//���������������������������������������������������������������������Ŀ
//� Declaracao de Variaveis                                             �
//�����������������������������������������������������������������������
Local aPar := {}
Local aRet := {}
Local bOk  := {|| MsgYesNo( "Confirma o processamento?", cCadastro ) }
Local bValid := {|| Iif(ApOleClient("MsExcel"),.T.,(MsgInfo("MsExcel n�o instalado"),)) }
//���������������������������������������������������������������������Ŀ
//� Parametros					                                             �
//�����������������������������������������������������������������������
AAdd( aPar, { 1, "Cliente de " 	,Space( TamSx3( "A1_COD" 	)[1] ) 	,"","","SA1"	,"",TamSx3( "A1_COD" 	)[1],.T.})
AAdd( aPar, { 1, "Loja de "		,Space( TamSx3( "A1_LOJA"	)[1] ) 	,"","",""		,"",TamSx3( "A1_LOJA"	)[1],.T.}) 
AAdd( aPar, { 1, "Cliente ate "	,Space( TamSx3( "A1_COD" 	)[1] ) 	,"","","SA1"	,"",TamSx3( "A1_COD" 	)[1],.T.})
AAdd( aPar, { 1, "Loja ate "		,Space( TamSx3( "A1_LOJA"	)[1] ) 	,"","",""		,"",TamSx3( "A1_LOJA"	)[1],.T.}) 
aAdd( aPar, { 1, "Emissao de "  ,Ctod("  /  /    "),"","","","",50,.F.}) // Tipo data
aAdd( aPar, { 1, "Emissao ate " ,Ctod("  /  /    "),"","","","",50,.F.}) // Tipo data
//AAdd( aPar, { 1, "Emissao de " 	,Space( TamSx3( "D2_EMISSAO")[1] ) ,"","",""		,"",TamSx3( "D2_EMISSAO")[1],.T.})
//AAdd( aPar, { 1, "Emissao ate "	,Space( TamSx3( "D2_EMISSAO")[1] ) ,"","",""		,"",TamSx3( "D2_EMISSAO")[1],.T.}) 
AAdd( aPar, { 2, "Tipo Custo"		,				 								,{"1-Custo Medio" , "2-Custo NF" } ,60, ,.T.}) 
AAdd( aPar, { 5, "Ao final do processamento abrir o Ms-Excel?",.F.,150,bValid,.F.})
//���������������������������������������������������������������������Ŀ
//� Tela para cooleta de informacoes                                    �
//�����������������������������������������������������������������������
If ParamBox( aPar, "Par�metros de processamento", @aRet, bOk, , , , , , , .F., .F. )
	
	Processa( {|| GerarArq( aRet ) }, cCadastro, "Processando arquivo, aguarde...", .F. )
Endif

Return
/*���������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �GeraArq   � Autor �Cristian Gutierrerz � Data �  07/12/18   ���
�������������������������������������������������������������������������͹��
���Descricao �Processamento e gera��o do arquivo Excel					     ���
�������������������������������������������������������������������������͹��
���Uso       �Uso Exclusivo Cliente Victorinox                            ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
���������������������������������������������������������������������������*/
Static Function GerarArq( aRet )
//���������������������������������������������������������������������Ŀ
//� Declaracao de Variaveis                                             �
//�����������������������������������������������������������������������
Local oFwMsEx 		:= NIL
Local cArq 			:= ""
Local cDir 			:= GetSrvProfString("Startpath","")
Local cWorkSheet 	:= ""
Local cDirTmp 		:= GetTempPath()
Local cQuery		:= ""                         
Local nCount		:= 0                         
Local cCFOP1		:= SuperGetMV( "ES_CFOPEX1" , .F. , '5102','5104','5106','5110','5111','5112','5114','5115','5117','5119','5120','5123','5401','5402','5403','5405','5551' )
Local cCFOP2		:= SuperGetMV( "ES_CFOPEX2" , .F. , '6102','6104','6106','6108','6110','6111','6112','6114','6115','6117','6118','6119','6120','6123','6129','6401','6402' )
Local cCFOP3		:= SuperGetMV( "ES_CFOPEX3" , .F. , '6403','6404','6551','5101','6101','6107','5933','6933' )
Local cCFEN1		:= SuperGetMV( "ES_CFOPEN1" , .F. , '1202', '1411', '2202', '2411' , '2204' )
Local cDesCan		:= ""                 
Local cF2Vend1		:= ""
Local cA3Nome		:= ""              
Local nCusto		:= 0
Local Ycan
Local xCan 
Local YDesCan := ""
//���������������������������������������������������������������������Ŀ
//� Declaracao de Variaveis                                             �
//�����������������������������������������������������������������������
//���������������������������������������������������������������������Ŀ
//� Monta query para selecao dos registros a serem gerados no Excel     �
//�����������������������������������������������������������������������
cQuery := "SELECT A1_COD, A1_LOJA, A1_NOME, A1_CGC, A1_REGIAO, A1_END,A1_CEP, A1_MUN, A1_EST, A1_GRPTRIB, A1_NREDUZ,"+ENTER
cQuery += "   		A1_DTCAD,'' AS C5_XSEGMEN ,'' AS ZV5_DESCC1, '' AS ZV5_DESCC2, '' AS ZV5_DESCC3,'' AS ZV5_DESCC4,'' AS ZV5_DESCC5, '' AS C5_XCANAIS,	"+ENTER
cQuery += "   		A1_DTCAD,A1_XSEGMEN, A1_DESCC1, A1_DESCC2, A1_DESCC3, A1_DESCC4, A1_DESCC5, A1_XCANAIS,	'' AS DESCAN,	"+ENTER
cQuery += "			A1_TEL, A1_EMAIL, A1_ULTCOM,																						"+ENTER
cQuery += "			D2_PEDIDO,D2_DOC, D2_SERIE, D2_EMISSAO, D2_COD, D2_QUANT, D2_PRCVEN,  D2_TOTAL,	D2_CF,					"+ENTER
cQuery += "			D2_VALBRUT,(D2_VALBRUT)-(D2_VALIPI+D2_VALICM+D2_VALIMP5+D2_VALIMP6+D2_ICMSRET+FT_DIFAL+D2_VFCPDIF+D2_VALISS)AS NET_SALLES, D2_CUSTO1, 														"+ENTER
cQuery += "			F2_VEND1, A3_NOME, BM_DESC AS D2_DSGRUPO, B2_CM1,							"+ENTER
cQuery += "			D2_VALIPI, D2_VALICM, D2_DESCON, D2_VALIMP5, D2_VALIMP6,D2_ICMSRET,FT_DIFAL,D2_VFCPDIF,D2_VALISS 	"+ENTER
cQuery += "FROM "+ RetSqlName( "SA1" ) +" SA1															 							"+ENTER
cQuery += "INNER JOIN "+ RetSqlName( "SD2" ) +" SD2 																				"+ENTER
cQuery += "   ON SD2.D_E_L_E_T_	= ''			 																						"+ENTER
cQuery += "  AND SD2.D2_FILIAL	= '"+ xFilial( "SD2") +"'																		"+ENTER
cQuery += "  AND SD2.D2_CLIENTE	= SA1.A1_COD																						"+ENTER
cQuery += "  AND SD2.D2_LOJA		= SA1.A1_LOJA														 								"+ENTER
cQuery += "  AND SD2.D2_EMISSAO	BETWEEN '"+ DtoS( aRet[5] ) +"' AND '"+ DtoS( aRet[6] )+"'	"+ENTER
cQuery += "  AND ( SD2.D2_CF		IN ( "+ cCFOP1 +" )																				"+ENTER
cQuery += "   		 OR  SD2.D2_CF	IN ( "+ cCFOP2 +" )																				"+ENTER
cQuery += "   		 OR  SD2.D2_CF	IN ( "+ cCFOP3 +" )																				"+ENTER
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

cQuery += "SELECT A1_COD, A1_LOJA, A1_NOME, A1_CGC, A1_REGIAO, A1_END,A1_CEP, A1_MUN, A1_EST, A1_GRPTRIB, A1_NREDUZ,"+ENTER
cQuery += "   		A1_DTCAD,'' AS C5_XSEGMEN ,'' AS ZV5_DESCC1, '' AS ZV5_DESCC2, '' AS ZV5_DESCC3,'' AS ZV5_DESCC4,'' AS ZV5_DESCC5, '' AS C5_XCANAIS,"+ENTER
cQuery += "   		A1_DTCAD,A1_XSEGMEN, A1_DESCC1, A1_DESCC2, A1_DESCC3, A1_DESCC4, A1_DESCC5, A1_XCANAIS,	'' AS DESCAN,	"+ENTER
cQuery += "			A1_TEL, A1_EMAIL, A1_ULTCOM,																						"+ENTER
cQuery += "			D1_PEDIDO AS D2_PEDIDO,D1_DOC AS D2_DOC, D1_SERIE AS D2_SERIE, D1_EMISSAO AS D2_EMISSAO, D1_COD AS D2_COD,				"+ENTER
cQuery += "			ROUND( D1_QUANT*-1 ,2 ) AS D2_QUANT, ROUND( D1_VUNIT*-1 , 2 ) AS D2_PRCVEN, 						 "+ENTER
cQuery += "			ROUND( D1_TOTAL*-1 ,2 ) AS D2_TOTAL,	                                                             "+ENTER
cQuery += "			D1_CF AS D2_CF, 																			         "+ENTER
cQuery += "			ROUND( FT_VALCONT*-1 ,2 )  AS D2_VALBRUT,ROUND( FT_VALCONT*-1 ,2 )-(D1_VALIPI+D1_VALICM+D1_VALIMP5+D1_VALIMP6+D1_ICMSRET+FT_DIFAL+D1_VFCPDIF+D1_VALISS)*-1 AS NET_SALLES,			"+ENTER
cQuery += "			( D1_CUSTO*-1  ) AS D2_CUSTO1,	F2_VEND1 AS F2_VEND1, A3_NOME, BM_DESC AS D2_DSGRUPO, ( B2_CM1*-1) AS B2_CM1,		"+ENTER
cQuery += "			D1_VALIPI*-1 AS D2_VALIPI, D1_VALICM*-1 AS D2_VALICM, D1_VALDESC*-1 AS D2_DESCON,					     "+ENTER
cQuery += "			D1_VALIMP5*-1 AS D2_VALIMP5, D1_VALIMP6*-1 AS D2_VALIMP6,D1_ICMSRET *-1 AS D2_ICMSRET,FT_DIFAL *-1 as FT_DIFAL,D1_VFCPDIF*-1 as D2_VFCPDIF,D1_VALISS *-1 as D2_VALISS  "+ENTER
cQuery += "FROM "+ RetSqlName( "SA1" ) +" SA1															 						"+ENTER
cQuery += "INNER JOIN "+ RetSqlName( "SD1" ) +" SD1 																			"+ENTER
cQuery += "   ON SD1.D_E_L_E_T_	= ''			 																				"+ENTER
cQuery += "  AND SD1.D1_FILIAL	= '"+ xFilial( "SD1") +"'																		"+ENTER
cQuery += "  AND SD1.D1_FORNECE	= SA1.A1_COD																					"+ENTER
cQuery += "  AND SD1.D1_LOJA		= SA1.A1_LOJA														 						"+ENTER
cQuery += "  AND SD1.D1_CF		IN ( "+ cCFEN1 +" )																				"+ENTER
cQuery += "INNER JOIN "+ RetSqlName( "SFT" ) +" SFT 																			"+ENTER
cQuery += "   ON SFT.D_E_L_E_T_	= ''			 																				"+ENTER
cQuery += "  AND SFT.FT_FILIAL	= '"+ xFilial( "SFT") +"'																		"+ENTER
cQuery += "  AND SFT.FT_TIPOMOV	= 'E'																									"+ENTER
cQuery += "  AND SFT.FT_SERIE		= SD1.D1_SERIE														 								"+ENTER
cQuery += "  AND SFT.FT_NFISCAL	= SD1.D1_DOC																						"+ENTER
cQuery += "  AND SFT.FT_CLIEFOR	= SD1.D1_FORNECE																					"+ENTER
cQuery += "  AND SFT.FT_LOJA		= SD1.D1_LOJA														 								"+ENTER
cQuery += "  AND SFT.FT_ITEM		= SD1.D1_ITEM 														 								"+ENTER
cQuery += "  AND SFT.FT_PRODUTO	= SD1.D1_COD 														 								"+ENTER
cQuery += "  AND SFT.FT_ENTRADA	BETWEEN '"+ DtoS(  aRet[5]  )+"' AND '"+ DtoS(  aRet[6] )+"'	"+ENTER
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
//���������������������������������������������������������������������Ŀ
//�Executa Query para geracao do arquivo										   �
//�����������������������������������������������������������������������
If Select( "TMP" ) > 0 
	dbSelectArea( "TMP" )
	dbCloseArea()
EndIf     
Memowrite( "c:\temp\Victorinox_rfatr04.txt" , cQuery )
TcQuery cQuery New Alias "TMP"
//���������������������������������������������������������������������Ŀ
//�Contagem de Registros para alimentar regua de Processamento			   �
//�����������������������������������������������������������������������
dbSelectArea( "TMP" )
dbGoTop()
While !TMP->( EOF() )
	nCount ++

	TMP->( dbSkip() )
EndDo
//���������������������������������������������������������������������Ŀ
//�Cria Excel																			   �
//�����������������������������������������������������������������������
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
oFwMsEx:AddColumn( cWorkSheet, cTable , "Cep"		, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Municipio"  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Estado"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Telefone"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "E-mail"	  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "GrpTrib"		, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Cadastro"  	, 1,1)
oFwMsEx:AddColumn( cWorkSheet, cTable , "Segmento"	  	, 1,1)
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
oFwMsEx:AddColumn( cWorkSheet, cTable , "Total Liquido"	, 1,1)
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
oFwMsEx:AddColumn( cWorkSheet, cTable , "DIFAL"			, 1,1)    
oFwMsEx:AddColumn( cWorkSheet, cTable , "Vlr fecp dif"	, 1,1)    
oFwMsEx:AddColumn( cWorkSheet, cTable , "ISS Servico"	, 1,1)    
oFwMsEx:AddColumn( cWorkSheet, cTable , "Ult. Compra"	, 1,1)     
//���������������������������������������������������������������������Ŀ
//�Alimenta regua de Processamento												   �
//�����������������������������������������������������������������������
ProcRegua( nCount )
//���������������������������������������������������������������������Ŀ
//�Loop principal da rotina 															�
//�����������������������������������������������������������������������
dbSelectArea( "TMP" )
TMP->( dbGoTop() )
While !( TMP->( EOF() ) )

	IncProc() 
	
	          
	//���������������������������������������������������������������������Ŀ
	//�Define descricao do canal															�
	//����������������������������������������������������������������������� 
/*
XSEG   := Posicione("SC5",1,xFilial("SC5")+TMP->D2_PEDIDO,"C5_XSEGMEN")	                  
IF EMPTY(XSEG) 
	IF ( Select( "_SA1" ) > 0 )
		_SA1->( dbCloseArea() )
	EndIF

	_cQry := " SELECT A1_COD,A1_LOJA,A1_XSEGMEN,A1_XCANAIS,A1_DESCC1,A1_DESCC2,A1_DESCC3,A1_DESCC4,A1_DESCC5 FROM "+ RetSqlName( "SA1" ) 
	_cQry += " WHERE A1_COD='"+ TMP->A1_COD +"' AND A1_LOJA='"+TMP->A1_LOJA+"' "
	_cQry := ChangeQuery(_cQry)

	//Executando consulta e setando o total da r�gua
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
	ElseIf YCAN == "8"
		YDesCan := "THIRD PARTY E-C"   
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
	ElseIf XCAN == "8"
		cDesCan := "THIRD PARTY E-C"   
	Else
		cDesCan := "NoName"
	EndIf	
	
	 
	//���������������������������������������������������������������������Ŀ
	//�Define custo a ser impresso														�
	//�����������������������������������������������������������������������
	nCusto := 0
	If Left( aRet[7] , 1 )  == "1"
		nCusto := TMP->B2_CM1
	Else
		nCusto := TMP->D2_CUSTO1
	EndIf		         

	//���������������������������������������������������������������������Ŀ
	//�Caso nao encontre Vendedor na NF, usa o default do cliente				�
	//�����������������������������������������������������������������������   
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
	//���������������������������������������������������������������������Ŀ
	//�Adiciona registros nas linhas da tabela										�
	//�����������������������������������������������������������������������
	
	oFwMsEx:AddRow( cWorkSheet, cTable, { TMP->A1_COD, TMP->A1_LOJA, TMP->A1_NOME, TMP->A1_NREDUZ, TMP->A1_CGC, TMP->A1_REGIAO, TMP->A1_END,TMP->A1_CEP, TMP->A1_MUN, TMP->A1_EST, TMP->A1_TEL, TMP->A1_EMAIL, TMP->A1_GRPTRIB,	TMP->A1_DTCAD,TMP->A1_XSEGMEN,XSEGC1, XSEGC2, XSEGC3, XSEGC4, XSEGC5, XCAN, cDesCan,YSEGC1, YSEGC2, YSEGC3, YSEGC4, YSEGC5,YCAN,YDesCan ,TMP->D2_DOC, TMP->D2_SERIE, TMP->D2_EMISSAO, TMP->D2_COD, TMP->D2_QUANT, TMP->D2_PRCVEN, TMP->D2_TOTAL, TMP->D2_VALBRUT,TMP->NET_SALLES, nCusto, TMP->D2_CF, TMP->D2_DSGRUPO, cF2Vend1, cA3Nome, TMP->D2_VALIPI, TMP->D2_VALICM, TMP->D2_DESCON, TMP->D2_VALIMP5, TMP->D2_VALIMP6,TMP->D2_ICMSRET,TMP->FT_DIFAL,TMP->D2_VFCPDIF,TMP->D2_VALISS,  DtoC( StoD( TMP->A1_ULTCOM ) ) } )
	
	TMP->( dbSkip() )
//_SD3->( dbCloseArea() )
EndDo
//���������������������������������������������������������������������Ŀ
//�Geracao do Arquivo																	�
//�����������������������������������������������������������������������
oFwMsEx:Activate()

cArq := CriaTrab( NIL, .F. ) + ".xml"

LjMsgRun( "Gerando o arquivo, aguarde...", cCadastro, {|| oFwMsEx:GetXMLFile( cArq ) } )

If __CopyFile( cArq, cDirTmp + cArq )

	If aRet[8]
		oExcelApp := MsExcel():New()
		oExcelApp:WorkBooks:Open( cDirTmp + cArq )
		oExcelApp:SetVisible(.T.)
	Else
		MsgInfo( "Arquivo " + cArq + " gerado com sucesso no diret�rio " + cDir )
	Endif
Else
	MsgInfo( "Arquivo n�o copiado para tempor�rio do usu�rio." )
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

Descri��o
	Fornece metodos para cria��o de arquivo XML compat�vel com as vers�es do MS Excel 2003, superiores e alternativos. 
	Fornece metodos para a cria��o de planilhas em formato de tabela e fornece tamb�m m�todos para a configura��o de cores e fonte.

M�todo: New
Sintaxe: FWMsExcel():New()-> NIL
Descri��o: M�todo construtor da classe

M�todo: ClassName
Sintaxe: FWMsExcel():ClassName()-> NIL
Descri��o: Fornece o nome da classe

M�todo: Activate
Sintaxe: FWMsExcel():Activate()-> NIL
Descri��o: Habilita o uso da classe, indicando que esta configurada e pronto para uso

M�todo: DeActivate
Sintaxe: FWMsExcel():DeActivate()-> NIL
Descri��o: Desabilita o uso da classe

M�todo: AddWorkSheet
Sintaxe: FWMsExcel():AddWorkSheet(< cWorkSheet >)-> NIL
Descri��o: Adiciona uma Worksheet ( Planilha )
Par�metros: 	cWorkSheet -> Nome da planilha que ser� adicionada

M�todo: AddTable
Sintaxe: FWMsExcel():AddTable(< cWorkSheet >, < cTable >)-> NIL
Descri��o: Adiciona uma tabela na Worksheet. Uma WorkSheet pode ter mais de uma tabela
Par�metros: 	cWorkSheet -> Nome da planilha
					cTable -----> Titulo da tabela que ser� adicionada

M�todo: AddColumn
Sintaxe: FWMsExcel():AddColumn(< cWorkSheet >, < cTable >, < cColumn >, < nAlign >, < nFormat >, < lTotal >)-> NIL
Descri��o: Adiciona uma coluna a tabela de uma Worksheet.
Par�metros: 	cWorkSheet -> Nome da planilha
					cTable -----> T�tulo da tabela
					cColumn ----> Titulo da tabela que ser� adicionada
					nAlign -----> Alinhamento da coluna ( 1-Left,2-Center,3-Right )
					nFormat ----> Codigo de formata��o ( 1-General,2-Number,3-Monet�rio,4-DateTime )
					lTotal -----> Indica se a coluna deve ser totalizada

M�todo: AddRow
Sintaxe: FWMsExcel():AddRow(< cWorkSheet >, < cTable >, < aRow >)-> NIL
Descri��o: Adiciona uma linha a tabela de uma Worksheet.
Par�metros: 	cWorkSheet -> Nome da planilha
					cTable -----> T�tulo da tabela
					aRow -------> Array com as informa��es da linha da linha
 
M�todo: GetXMLFile
Sintaxe: FWMsExcel():GetXMLFile(< cFile >)-> NIL
Descri��o: Cria um arquivo no formato XML do MSExcel 2003 em diante
Par�metros: 	cFile -> Nome do arquivo fisico
 
M�todo: IsWorkSheet
Sintaxe: FWMsExcel():IsWorkSheet(< cWorkSheet >)-> NIL
Descri��o: Verifica se um nome j� foi utilizado para definir uma WorkSheet
Par�metros: 	cWorksheet -> Nome da planilha
 
M�todo: SetFontSize
Sintaxe: FWMsExcel():SetFontSize(< nFontSize >)-> NIL
Descri��o: Define o tamanho para a fonte para todos os estilos da planilha
Par�metros: 	nFontSize -> Tamanho da fonte 
 
M�todo: SetFont
Sintaxe: FWMsExcel():SetFont(< cFont >)-> NIL
Descri��o: Define a fonte que ser� utilizada para todos os estilos da planilha
Par�metros: 	cFont -> Nome da fonte
 
M�todo: SetItalic
Sintaxe: FWMsExcel():SetItalic(< lItalic >)-> NIL
Descri��o: Define se a fonte ter� a configura��o "Italico" para todos os estilos da planilha
Par�metros: 	lItalic -> Se verdadeiro o formato � it�lico

M�todo: SetBold
Sintaxe: FWMsExcel():SetBold(< lBold >)-> NIL
Descri��o: Define se a fonte ter� a configura��o "Negrito" para todos os estilos da planilha
Par�metros: 	lBold -> Se verdadeiro o formato � negrito
 
M�todo: SetUnderLine
Sintaxe: FWMsExcel():SetUnderLine(< lUnderline >)-> NIL
Descri��o: Define se a fonte ter� a configura��o "Sublinhado" para todos os estilos da planilha
Par�metros: 	lUnderLine -> Se verdadeiro o formato � sublinhado
 
M�todo: SetFrGeneralColor
Sintaxe: FWMsExcel():SetFrGeneralColor(< cColor >)-> NIL
Descri��o: Define a cor de preenchimento geral para todos os estilos da planilha
Par�metros: 	cColor -> Cor em hexadecimal
 
M�todo: SetBgGeneralColor
Sintaxe: FWMsExcel():SetBgGeneralColor(< cColor >)-> NIL
Descri��o: Define a cor da fonte do estilo do Titulo para todos os estilos da planilha
Par�metros: 	cColor -> Cor em hexadecimal
 
M�todo: SetTitleFont
Sintaxe: FWMsExcel():SetTitleFont(< cFont >)-> NIL
Descri��o: Define a fonte que ser� utilizada no estilo do T�tulo
Par�metros: 	cFont -> Nome da fonte

M�todo: SetTitleSizeFont
Sintaxe: FWMsExcel():SetTitleSizeFont(< nFontSize >)-> NIL
Descri��o: Define o tamanho para a fonte do estilo do T�tulo
Par�metros: nFontSize -> Tamanho da fonte
 
M�todo: SetTitleItalic
Sintaxe: FWMsExcel():SetTitleItalic(< lItalic >)-> NIL
Descri��o: Define se a fonte ter� a configura��o "Italico" no estilo do Titulo
Par�metros: 	lItalic -> Se verdadeiro define o formato it�lico
 
M�todo: SetTitleBold
Sintaxe: FWMsExcel():SetTitleBold(< lBold >)-> NIL
Descri��o: Define se a fonte ter� a configura��o "Negrito" no estilo do T�tulo
Par�metros: 	lBold -> Se verdadeiro define o fomrato negrito
 
M�todo: SetTitleUnderLine
Sintaxe: FWMsExcel():SetTitleUnderLine(< lUnderline >)-> NIL
Descri��o: Define se a fonte ter� a configura��o "Sublinhado" no estilo do Titulo
Par�metros: 	lUnderLine -> Se verdadeiro define o formato como sublinhado

M�todo: SetTitleFrColor
Sintaxe: FWMsExcel():SetTitleFrColor(< cColor >)-> NIL
Descri��o: Define a cor de preenchimento do estilo do Titulo
Par�metros: 	cColor -> Cor em hexadecimal
 
M�todo: SetTitleBgColor
Sintaxe: FWMsExcel():SetTitleBgColor(< cColor >)-> NIL
Descri��o: Define a cor da fonte do estilo do Titulo
Par�metros: 	cColor -> Cor em hexadecimal
 
M�todo: SetHeaderFont
Sintaxe: FWMsExcel():SetHeaderFont(< cFont >)-> NIL
Descri��o: Define a fonte que ser� utilizada no estilo do Cabe�alho
Par�metros: 	cFont -> Nome da fonte
 
M�todo: SetHeaderSizeFont
Sintaxe: FWMsExcel():SetHeaderSizeFont(< nFontSize >)-> NIL
Descri��o: Define o tamanho para a fonte do estilo do Cabe�alho
Par�metros: 	nFontSize -> Tamanho
 
M�todo: SetHeaderItalic
Sintaxe: FWMsExcel():SetHeaderItalic(< lItalic >)-> NIL
Descri��o: Define se a fonte ter� a configura��o "Italico" no estilo do Cabe�alho
Par�metros: 	lItalic -> Verdadeiro indica que sim
 
M�todo: SetHeaderBold
Sintaxe: FWMsExcel():SetHeaderBold(< lBold >)-> NIL
Descri��o: Define se a fonte ter� a configura��o "Negrito" no estilo do Cabe�alho
Par�metros: 	lBold -> Verdadeiro indica que sim
 
M�todo: SetHeaderUnderLine
Sintaxe: FWMsExcel():SetHeaderUnderLine(< lUnderline >)-> NIL
Descri��o: Define se a fonte ter� a configura��o "Sublinhado" no estilo do Cabe�alho
Par�metros: 	lUnderLine -> Verdadeiro indica que sim

M�todo: SetFrColorHeader
Sintaxe: FWMsExcel():SetFrColorHeader(< cColor >)-> NIL
Descri��o: Define a cor de preenchimento do estilo do Cabe�alho
Par�metros: 	cColor -> Cor em hexadecimal
 
M�todo: SetBgColorHeader
Sintaxe: FWMsExcel():SetBgColorHeader(< cColor >)-> NIL
Descri��o: Define a cor da fonte do estilo do Cabe�alho
Par�metros: 	cColor -> Cor em hexadecimal
 
M�todo: SetLineFont
Sintaxe: FWMsExcel():SetLineFont(< cFont >)-> NIL
Descri��o: Define a fonte que ser� utilizada no estilo da Linha
Par�metros: 	cFont -> Nome da fonte

M�todo: SetLineSizeFont
Sintaxe: FWMsExcel():SetLineSizeFont(< nFontSize >)-> NIL
Descri��o: Define o tamanho para a fonte do estilo da Linha
Par�metros: nFontSize -> Tamanho
 
M�todo: SetLineItalic
Sintaxe: FWMsExcel():SetLineItalic(< lItalic >)-> NIL
Descri��o: Define se a fonte ter� a configura��o "Italico" no estilo da Linha
Par�metros: 	lItalic  -> .T. indica que sim
 
M�todo: SetLineBold
Sintaxe: FWMsExcel():SetLineBold(< lBold >)-> NIL
Descri��o: Define se a fonte ter� a configura��o "Negrito" no estilo da Linha
Par�metros: 	lBold -> .T. indica que sim

M�todo: SetLineUnderLine
Sintaxe: FWMsExcel():SetLineUnderLine(< lUnderline >)-> NIL
Descri��o: Define se a fonte ter� a configura��o "Sublinhado" no estilo da Linha
Par�metros: lUnderLine -> .T. indica que sim

M�todo: SetLineFrColor
Sintaxe: FWMsExcel():SetLineFrColor(< cColor >)-> NIL
Descri��o: Define a cor de preenchimento do estilo da Linha
Par�metros: 	cColor -> Cor em Hexadecimal 

M�todo: SetLineBgColor
Sintaxe: FWMsExcel():SetLineBgColor(< cColor >)-> NIL
Descri��o: Define a cor da fonte do estilo da Linha
Par�metros: 	cColor -> Cor em Hexadecimal 
 
M�todo: Set2LineFont
Sintaxe: FWMsExcel():Set2LineFont(< cFont >)-> NIL
Descri��o: Define a fonte que ser� utilizada no estilo da Linha 2
Par�metros: 	cFont -> Nome da Fonte

M�todo: Set2LineSizeFont
Sintaxe: FWMsExcel():Set2LineSizeFont(< nFontSize >)-> NIL
Descri��o: Define o tamanho para a fonte do estilo da Linha 2
Par�metros: 	nFontSize -> Tamanho
 
M�todo: Set2LineItalic
Sintaxe: FWMsExcel():Set2LineItalic(< lItalic >)-> NIL
Descri��o: Define se a fonte ter� a configura��o "Italico" no estilo da Linha 2
Par�metros: 	lItalic -> .T. indica que sim

M�todo: Set2LineBold
Sintaxe: FWMsExcel():Set2LineBold(< lBold >)-> NIL
Descri��o: Define se a fonte ter� a configura��o "Negrito" no estilo da Linha 2
Par�metros: 	lBold -> .T. indica que sim

M�todo: Set2LineUnderLine
Sintaxe: FWMsExcel():Set2LineUnderLine(< lUnderline >)-> NIL
Descri��o: Define se a fonte ter� a configura��o "Sublinhado" no estilo da Linha 2
Par�metros: 	lUndeLine -> .T. indica que sim
 
M�todo: Set2LineFrColor
Sintaxe: FWMsExcel():Set2LineFrColor(< cColor >)-> NIL
Descri��o: Define a cor de preenchimento do estilo da Linha 2
Par�metros: 	cColor -> Cor em Hexadecimal
 
M�todo: Set2LineBgColor
Sintaxe: FWMsExcel():Set2LineBgColor(< cColor >)-> NIL
Descri��o: Define a cor da fonte do estilo da Linha 2
Par�metros: 	cColor -> Cor em Hexadecimal
*/