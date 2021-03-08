#Include 'rwmake.ch'
#Include 'Protheus.ch'
#Include 'TbiConn.ch'
#include 'apwebsrv.ch'
#include 'apwebex.ch'
#include 'ap5mail.ch'

// NOME DO ARQUIVO NO EXCEL
User Function RELAT02()
Local cTempPath     
Local cFile			:= "Disponibilidade + Preco Venda"+Dtos(MsDate())+"_"+Substr(Time(),1,2)+Substr(Time(),4,2)+Substr(Time(),7,2)   
Local _aParamBox	:= {}
Local _aRet     	:= {}
Local _cQuery       
//

Conout("1. YAMR080 - Inicio "+Dtoc(MsDate())+" "+Time())

//CASO EXECUTADO VIA SCHEDULER

If Select("SX2") = 0                 
	Prepare Environment Empresa "01" Filial "01"   
Endif	

cTempPath     := GetTempPath(.T.)

If !IsBlind()  

	Conout("2. YAMR080 - N„o È job "+Dtoc(MsDate())+" "+Time()) 

	
//PERGUNTA
	aAdd(_aParamBox,{1,"Tabela de:",		"    ",				"@!",	"",	"DA0",	"",		0,		.F.})
	aAdd(_aParamBox,{1,"Tabela ate:",	"ZZZZ",				"@!",	"",	"DA0",	"",		0,		.T.})
	aAdd(_aParamBox,{1,"Produto de:",	"               ",	"@!",	"",	"SB1",	"",		0,		.F.})	  
	aAdd(_aParamBox,{1,"Produto ate:",	"ZZZZZZZZZZZZZZZ",	"@!",	"",	"SB1",	"",		0,		.T.})
	aAdd(_aParamBox,{1,"Filial de:",	"0103",	"@!",	"",	"SB2",	"",		0,		.F.})
	aAdd(_aParamBox,{1,"Filial ate:",	"0103",	"@!",	"",	"SB2",	"",		0,		.T.})	  
	//aAdd(_aParamBox,{1,"Local de:",		"  ",				"@!",	"",	"NNR",	"",		0,		.F.})
	//aAdd(_aParamBox,{1,"Local ate:",	"ZZ",				"@!",	"",	"NNR",	"",		0,		.T.})
	
	If ParamBox(_aParamBox,"Informe os dados:",@_aRet) 
	
		Processa( {|| Atualiza(cTempPath,cFile,_aRet)}, "Gerando Excel..." )    
		
	EndiF	
	
Else                           

//CASO EXECUTADO VIA JOB
	Conout("3. YAMR080 - È job "+Dtoc(MsDate())+" "+Time())
   //VALORES PADRAO PERGUNTA SE SCHEDULER
	AADD(_aRet,"    ")	
	AADD(_aRet,"ZZZZ")	
	AADD(_aRet,"               ")
	AADD(_aRet,"ZZZZZZZZZZZZZZZ")
	
	Atualiza(cTempPath,cFile,_aRet)	
	
Endif	

Conout("99. YAMR080 - Fim "+Dtoc(MsDate())+" "+Time())
	
Return .T.


Static Function Atualiza(cTempPath,cFile,_aRet)

Local cWorkSheet	:= OemToAnsi("Tabela 01")
Local cAliasQry1 	:= GetNextAlias()   
Local oExcel 		:= FWMsExcelEx():New()                           
Local cTable		:= "Disponibilidade x Preco de Venda"
Local cTexto		:= "GRUPO;Mod. Yamaha;CODIGO;DESCRICAO;LOCAL;ESTOQUE"+Chr(13)+Chr(10)     

Conout("4. YAMR080 - Gerando Excel "+Dtoc(MsDate())+" "+Time())
//COLUNAS DO EXCEL
oExcel:AddworkSheet(cWorkSheet)
oExcel:AddTable (cWorkSheet,cTable)
oExcel:AddColumn(cWorkSheet,cTable,"Cod. Produto",2,1)//COLUNA DO EXCEL         
oExcel:AddColumn(cWorkSheet,cTable,"Desc. Produto",1,1)//COLUNA DO EXCEL
//oExcel:AddColumn(cWorkSheet,cTable,"Saldo Disponivel",3,2,.T.)//COLUNA DO EXCEL
oExcel:AddColumn(cWorkSheet,cTable,"Saldo Disponivel",3,2)
oExcel:AddColumn(cWorkSheet,cTable,"Prc Venda",3,2)//COLUNA DO EXCEL
oExcel:AddColumn(cWorkSheet,cTable,"Cod. Tabela",2,1)//COLUNA DO EXCEL
oExcel:AddColumn(cWorkSheet,cTable,"Cod. Armazem",2,1)//COLUNA DO EXCEL
oExcel:AddColumn(cWorkSheet,cTable,"Desc. Grupo",2,1)//COLUNA DO EXCEL
//oExcel:AddColumn(cWorkSheet,cTable,"ESTOQUE",3,2,.T.)//COLUNA DO EXCEL

/*
cTexto := "<Table>"
cTexto += "<table border='1'>"
cTexto += "<tr>
cTexto += "<td align='center'><b>GRUPO</b></td>"
cTexto += "<td align='center'><b>Mod. Yamaha</b></td>"
cTexto += "<td align='center'><b>CODIGO</b></td>"
cTexto += "<td align='center'><b>DESCRICAO</b></td>"
cTexto += "<td align='center'><b>LOCAL</b></td>"      
cTexto += "<td align='center'><b>SALDO EM ESTOQUE</b></td>"
cTexto += "</tr>                                     
*/

Conout("5. YAMR080 - Antes da Query "+Dtoc(MsDate())+" "+Time())

//FUN«√O QUERY SQL

/*
BeginSql Alias cAliasQry1

	SELECT B1_GRUPO AS GRUPO, B1_MODYAM AS MODYAM, B2_COD AS CODIGO, B1_DESC AS DESCRICAO, B2_LOCAL AS LOCAL, B2_QATU AS ESTOQUE
	FROM SB2010
	INNER JOIN SB1010 ON B2_COD = B1_COD AND SB2010.D_E_L_E_T_ <> '*'
	WHERE            
	B1_GRUPO BETWEEN %Exp:_aRet[1]% AND %Exp:_aRet[2]% AND
	B1_COD BETWEEN %Exp:_aRet[3]% AND %Exp:_aRet[4]% AND
	B2_LOCAL BETWEEN %Exp:_aRet[5]% AND %Exp:_aRet[6]% AND
	SB2010.D_E_L_E_T_ <> '*' AND
	B2_QATU > 0 
	ORDER BY B1_GRUPO, B1_COD

EndSql
*/           

_cQuery := "SELECT B2_COD, B1_DESC, B2_QATU - B2_QACLASS - B2_RESERVA AS SALDO, DA1_PRCVEN, DA1_CODTAB, B2_LOCAL, BM_DESC "
_cQuery += "FROM "+RetSQLName("DA1")+" DA1 "
_cQuery += "INNER JOIN "+RetSQLName("DA0")+" AS DA0 ON DA0_CODTAB = DA1_CODTAB AND DA0.D_E_L_E_T_ <> '*' "
_cQuery += "INNER JOIN "+RetSQLName("SB2")+" AS SB2 ON B2_COD = DA1_CODPRO AND SB2.D_E_L_E_T_ <> '*' "
_cQuery += "INNER JOIN "+RetSQLName("SB1")+" AS SB1 ON B1_COD = B2_COD AND SB1.D_E_L_E_T_ <> '*' "
_cQuery += "INNER JOIN "+RetSQLName("SBM")+" AS SBM ON BM_GRUPO = B1_GRUPO AND SBM.D_E_L_E_T_ <> '*' "	   
_cQuery += "WHERE "
//_cQuery += "B2_FILIAL == ' '0103' " "' AND "
_cQuery += "B2_FILIAL BETWEEN '"+_aRet[5]+"' AND '"+_aRet[6]+"' AND "
_cQuery += "DA0_CODTAB BETWEEN '"+_aRet[1]+"' AND '"+_aRet[2]+"' AND " 
_cQuery += "DA1_CODPRO BETWEEN '"+_aRet[3]+"' AND '"+_aRet[4]+"' AND "
_cQuery += "B2_LOCAL <> '97'  AND "
_cQuery += "DA1.D_E_L_E_T_ <> '*' "


dbUseArea(.T., "TOPCONN", TCGenQry(,,_cQuery), 'QRY', .T., .T.)

dbSelectArea("QRY")

Conout("6. YAMR080 - Depois da Query "+Dtoc(MsDate())+" "+Time())                                                                                  

ProcRegua(RecCount())
DbGotop()

_n := 0

Do While !Eof() 

_n += 1

	IncProc()
//	cTexto += " "+(cAliasQry1)->GRUPO+";"+(cAliasQry1)->MODYAM+";"+(cAliasQry1)->CODIGO+";"+(cAliasQry1)->DESCRICAO+";"+" "+(cAliasQry1)->LOCAL+";"+Str((cAliasQry1)->ESTOQUE)+Chr(13)+Chr(10)
//	cTexto += "<tr>
//	cTexto += "<td>'"+(cAliasQry1)->GRUPO+"</td>"
//	cTexto += "<td>"+(cAliasQry1)->MODYAM+"</td>"
//	cTexto += "<td>"+(cAliasQry1)->CODIGO+"</td>"
//	cTexto += "<td>"+(cAliasQry1)->DESCRICAO+"</td>"
//	cTexto += "<td>'"+(cAliasQry1)->LOCAL+"</td>"
//	cTexto += "<td>"+Str((cAliasQry1)->ESTOQUE)+"</td>"
//	cTexto += "</tr>
                                                                             

//FORMATO DO EXCEL COM A QUERY
	//oExcel:AddRow( cWorkSheet, cTable, {(cAliasQry1)->GRUPO,TiraChar((cAliasQry1)->MODYAM),(cAliasQry1)->CODIGO,TiraChar((cAliasQry1)->DESCRICAO),(cAliasQry1)->LOCAL,(cAliasQry1)->ESTOQUE} )  
	oExcel:AddRow( cWorkSheet, cTable, {QRY->B2_COD, TiraChar(QRY->B1_DESC), QRY->SALDO, QRY->DA1_PRCVEN, QRY->DA1_CODTAB, QRY->B2_LOCAL, TiraChar(QRY->BM_DESC)} )

	DbSkip()

//	Exit

Enddo       

//cTexto += "</Table>      

MemoWrite(cFile,cTexto)     

if _n >= 0

oExcel:Activate()
oExcel:GetXMLFile(cFile+".xml")
	
	

//If !IsBlind()               
	
	If CpyS2T( "\SYSTEM\"+ cFile+".xml", cTempPath, .T. )
	 	Conout("7. RELAT02 - Copiando Arquivo "+cFile+" "+Dtoc(MsDate())+" "+Time())
		If !ApOleClient("MsExcel")
			MsgStop("Microsoft Excel nao instalado.") //"Microsoft Excel nao instalado."
			Return
		EndIf                                                   
		oExcelApp:= MsExcel():New()
		oExcelApp:WorkBooks:Open(AllTrim(cTempPath)+ cFile+".xml")
		oExcelApp:SetVisible(.T.)
	else
	 	Conout("8. YAMR080 - Ocorreram problemas na cÛpia do arquivo. "+cFile+" "+Dtoc(MsDate())+" "+Time())
		MsgStop("Ocorreram problemas na cÛpia do arquivo.")
	endif
endif
		                         

//ENVIAR POR EMAIL

//Else
	
	//U_enviar("workflow@yamaha.com.br",GetNewPar("MV_XEMAILPE","luiz.magalhaes@music.yamaha.com"),"PosiÁ„o de Estoques ","PosiÁ„o de Estoques "+Dtoc(MsDate())+" "+Time(),"\system\"+cFile+".xml")	
		
//Endif	
	
Return

Static Function TiraChar(_cString)
Local _x	
Local _cRet		:= ""           
Local _cChar	:= ""

For _x:=1 to Len(_cString)

	 _cChar	:= Upper(Substr(_cString,_x,1))
	 If _cChar $ "¡¿√¬ƒ∆µ‚·‡‚É†"
	 	_cChar	:= "A"
	 ElseIf _cChar $ "…» ÀÈÇà"
	 	_cChar	:= "E"	
	 ElseIf _cChar $ "ÕÕŒœÌ°"
	 	_cChar	:= "I"	
	 ElseIf _cChar $ "”“’÷¢Û"
	 	_cChar	:= "O"	                        
	 ElseIf _cChar $ "⁄Ÿ€‹£"
	 	_cChar	:= "U"	
	 ElseIf _cChar $ "«á"
	 	_cChar	:= "C"	
	 Endif	
	    
	 
	 If !_cChar $ "ABCDEFGHIJKLMNOPQRSTUVXZWY.',&/\()%¥`:1234567890-+}{[] "
	 	_cChar := " "
	 Endif	 	

	_cRet += _cChar 		

Next

//// _cRet = strtran (_cRet, chr (9), " ") // TAB

Return _cRet

/*
User function enviar(cMailDe,cMailPara,cAssunto,cCorpo,cAnexo)

local cServer := getMV("MV_RELSERV") //endereÁo SMTP
local lAutentic := getMV("MV_RELAUTH") //utilize em caso de necessidade de autenticaÁ„o
local cAccount := getmv("MV_RELACNT") //conta
local cPassword := getMV("MV_RELAPSW") //senha
LOcal cQL := CHR(13) + CHR(10)
local cRemoteip := Getclientip()
local cRemoteComputer := GetComputerName()
local lConectou := .f.       
Local cUsername := "WorkFlow"

Conout("9. YAMR080 - Antes do envio do E-mail "+cAnexo+" "+Dtoc(MsDate())+" "+Time())

// conecta com o servidor de e-mail
CONNECT SMTP SERVER cServer ACCOUNT cAccount PASSWORD cPassword Result lConectou

If lAutentic
	mailAuth(cAccount, cPassword)
Endif	

If lConectou
//  	cCorpo += cQL + cQL + "==========================================================="
//  	cCorpo += cQL + " Enviado por: " + cUsername
//  	cCorpo += cQL + " Computador: " + cRemoteComputer
//  	cCorpo += cQL + " IP: " + cRemoteip
//  	cCorpo += cQL + "===========================================================î
	
	SEND MAIL FROM cMailDe TO cMailPara SUBJECT cAssunto BODY cCorpo FORMAT TEXT RESULT lEnviado ATTACHMENT cAnexo
	if !lEnviado
		Conout("10. YAMR080 - N„o foi possivel enviar a mensagem "+Dtoc(MsDate())+" "+Time())
		alert("ALERTA: N„o foi possivel enviar a mensagem") //, pois ocorreu o seguinte erro: î + sMensagem + ì.î)
	else                                                                                                            
		Conout("11. YAMR080 - E-mail transmitido com sucesso para " + cMailPara +"! "+Dtoc(MsDate())+" "+Time())
		alert("E-mail transmitido com sucesso para " + cMailPara +"!")
	endif
else                                          
 	Conout("12. YAMR080 - N„o foi possivel executar sua solicitaÁ„o, pois n„o houve resposta do servidor de e-mail. "+Dtoc(MsDate())+" "+Time())
	alert("N„o foi possivel executar sua solicitaÁ„o, pois n„o houve resposta do servidor de e-mail."+cQL+cQL+"Informe ao Administrador do Sistema!")
return .f.
Endif

DISCONNECT SMTP SERVER Result lDisConectou

*/   *
Return                            