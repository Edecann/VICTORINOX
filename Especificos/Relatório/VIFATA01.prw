#include 'protheus.ch'
#include 'parmtype.ch'

//------------------------------------------------------------------------\\
/*/{Protheus.doc} VIFATA01
//TODO Rotina para gerar arquivo txt, a partir das ordens de separacao.
//@author Claudio Macedo
//@since 13/03/2020
//@version 1.0
//@return Nil
//@type Function
/*/
//------------------------------------------------------------------------\\
User Function VIFATA01()

Local aCampos   := {}
Local cPerg     := Padr('VIFATA01',10)

Private oBrowse   := FWMarkBrowse():New()
Private cMarca    := Getmark()
Private lMarcar   := .T.
Private cCadastro := "Exportar Ordens de Separacao"
Private aRotina   := {}
Private cAliasCB8 := GetNextAlias()
Private cArquivo  := ''

Private lOK := .T.
Private cArq := "TRB"
aadd(aRotina,{'Exportar', 'U_Processa()',0,4})

If Pergunte(cPerg,.T.)

	Processa({|| GetCB8(cArq)}, 'Selecionando Registros ...')
IF TRB->( !Eof() )	
	TRB->(DbGoTop())
ENDIF	
	If TRB->(EOF())
		Help(NIL, NIL, 'Ordens de Separacao',;
				  NIL, 'Ordens de separacao nao encontradas', 1, 0, NIL, NIL, NIL, NIL, NIL,;
	                  {'Verifique os parametros informados.'})
//		oTempTable:Delete() 
		Return Nil
	Endif

    oBrowse:SetDescription(cCadastro)
    oBrowse:SetAlias('TRB') 
    oBrowse:SetFieldMark('OK')
    oBrowse:SetTemporary() 
    oBrowse:DisableDetails()
 
    oBrowse:SetColumns(AaddColuna('CB8_ORDSEP', 'Ordem'    , 02, '@!', 1, 6, 0))
    oBrowse:SetColumns(AaddColuna('CB8_PEDIDO', 'Pedido'   , 03, '@!', 1, 6, 0))
    oBrowse:SetColumns(AaddColuna('C5_XCANAIS', 'Canal'    , 04, '@!', 1, 1, 0))
    oBrowse:SetColumns(AaddColuna('C5_EMISSAO', 'Emiss�o'  , 05, '@!', 1, 8, 0))
    oBrowse:SetColumns(AaddColuna('C5_CLIENTE', 'Cliente'  , 06, '@!', 1, 6, 0))
    oBrowse:SetColumns(AaddColuna('C5_LOJACLI', 'Loja'     , 07, '@!', 1, 2, 0))
    oBrowse:SetColumns(AaddColuna('A1_NOME'   , 'Nome'     , 08, '@!', 1, 40, 0))
    oBrowse:SetColumns(AaddColuna('A1_MUN'    , 'Munic�pio', 09, '@!', 1, 60, 0))
    oBrowse:SetColumns(AaddColuna('A1_EST'    , 'Estado'   , 10, '@!', 1,  2, 0))
    oBrowse:SetColumns(AaddColuna('C5_TRANSP' , 'Transportadora', 11, '@!', 1,  6, 0))
    oBrowse:SetColumns(AaddColuna('A4_NOME'   , 'Nome', 12, '@!', 1, 40, 0))
    
    oBrowse:bAllMark := { || MarkAll(oBrowse:Mark(),lMarcar := !lMarcar ), oBrowse:Refresh(.T.)  }

    oBrowse:Activate()
    
    oBrowse:oBrowse:Setfocus()

	//---------------------------------
	//Exclui a tabela 
	//---------------------------------
//	oTempTable:Delete() 
	
Endif

Return Nil

//------------------------------------------------------------------------\\
/*/{Protheus.doc} GetCB8
//TODO Rotina para selecionar as ordens de separa��o, conforme os par�metros
	   informados.
@author Claudio Macedo
@since 13/03/2020
@version 1.0
@return Nil
@type Function
/*/
//------------------------------------------------------------------------\\
Static Function GetCB8(cArqCB8)


Local aStru   := {}
Local oTempTable
Local cAlias := "TRB"      

Private cArq    := cArqCB8

BeginSQL Alias cAliasCB8
		
	SELECT 	CB8_ORDSEP, CB8_PEDIDO, CB8_ITEM, C5_CLIENTE, C5_LOJACLI, C5_TRANSP, C5_EMISSAO, 
	   		C5_XCANAIS, A1_NOME, A1_EST, A1_MUN, A4_NOME
	FROM CB8010 CB8 INNER JOIN CB7010 CB7 ON
			CB7_FILIAL = CB8_FILIAL
		AND CB7_ORDSEP = CB8_ORDSEP
		AND CB7_STATUS = '0'
		AND CB7_XARQTX <> 'S'
		AND CB7.%notdel% INNER JOIN SB1010 SB1 ON
			B1_FILIAL = %xFilial:SB1%
		AND B1_COD = CB8_PROD
		AND B1_MSBLQL <> '1'
		AND SB1.%notdel%INNER JOIN SC5010 SC5 ON
			C5_FILIAL = CB8_FILIAL
		AND C5_NUM = CB8_PEDIDO
		AND SC5.%notdel% INNER JOIN SC6010 SC6 ON
			C6_FILIAL = CB8_FILIAL
		AND C6_NUM = CB8_PEDIDO
		AND C6_PRODUTO = CB8_PROD
		AND C6_ITEM = CB8_ITEM
		AND SC6.%notdel% INNER JOIN SA1010 SA1 ON
			A1_FILIAL = %xFilial:SA1%
		AND A1_COD = C5_CLIENTE
		AND A1_LOJA = C5_LOJACLI 
		AND SA1.%notdel% INNER JOIN SA4010 SA4 ON
			A4_FILIAL = %xFilial:SA4%
		AND C5_TRANSP = A4_COD
		AND SA4.%notdel%
	WHERE CB8_FILIAL = %xFilial:CB8%
		AND CB8_ORDSEP BETWEEN %Exp:mv_par01% AND %Exp:mv_par02% 
		AND CB8.%notdel%
	ORDER BY CB8_ORDSEP, CB8_PEDIDO
		
EndSQL

//aadd(aStru,{'MARK'      , 'C',  2, 0})
aadd(aStru,{'OK'      , 'C',  2, 0})
aadd(aStru,{'CB8_ORDSEP', 'C',  6, 0})
aadd(aStru,{'CB8_PEDIDO', 'C',  6, 0})
aadd(aStru,{'CB8_ITEM'  , 'C',  2, 0})
aadd(aStru,{'C5_XCANAIS', 'C',  1, 0})
aadd(aStru,{'C5_EMISSAO', 'D',  8, 0})
aadd(aStru,{'C5_CLIENTE', 'C',  6, 0})
aadd(aStru,{'C5_LOJACLI', 'C',  2, 0})
aadd(aStru,{'A1_NOME'   , 'C', 40, 0})
aadd(aStru,{'A1_MUN'    , 'C', 60, 0})
aadd(aStru,{'A1_EST'    , 'C',  2, 0})
aadd(aStru,{'C5_TRANSP' , 'C',  6, 0})
aadd(aStru,{'A4_NOME'   , 'C', 40, 0})

//-------------------
//Criação do objeto
//-------------------
oTempTable := FWTemporaryTable():New( cAlias , aStru )
//------------------
//Criação da tabela
//------------------
oTempTable:Create()

(cAliasCB8)->(DbGoTop())
 
While !(cAliasCB8)->(EOF()) 

	TRB->(reclock('TRB',.T.))
	
	TRB->CB8_ORDSEP := (cAliasCB8)->CB8_ORDSEP
	TRB->CB8_PEDIDO := (cAliasCB8)->CB8_PEDIDO
	TRB->CB8_ITEM   := (cAliasCB8)->CB8_ITEM
	TRB->C5_XCANAIS := (cAliasCB8)->C5_XCANAIS
	TRB->C5_EMISSAO := STOD((cAliasCB8)->C5_EMISSAO)
	TRB->C5_CLIENTE := (cAliasCB8)->C5_CLIENTE
	TRB->C5_LOJACLI := (cAliasCB8)->C5_LOJACLI
	TRB->A1_NOME    := (cAliasCB8)->A1_NOME
	TRB->A1_MUN     := (cAliasCB8)->A1_MUN
	TRB->A1_EST     := (cAliasCB8)->A1_EST
	TRB->C5_TRANSP  := (cAliasCB8)->C5_TRANSP
	TRB->A4_NOME    := (cAliasCB8)->A4_NOME
	
	TRB->(MsUnlock())

	(cAliasCB8)->(DbSkip())

	While !(cAliasCB8)->(EOF()) .And. (cAliasCB8)->CB8_ORDSEP = TRB->CB8_ORDSEP .And. (cAliasCB8)->CB8_PEDIDO = TRB->CB8_PEDIDO

		(cAliasCB8)->(DbSkip())

	Enddo
	
Enddo

Return Nil

// ------------------------------------------------------- \\
/*/{Protheus.doc} Marcar
// Funcao executada ao marcar/desmarcar um registro.   
@author Claudio Macedo
@since 15/11/2018
@version 1.0
@return Nil
@type function
/*/
// ------------------------------------------------------- \\
Static Function Marcar()

TRB->(RecLock('TRB',.F.))

If Marked('OK')
	TRB->OK := cMarca
Else	
	TRB->OK := ''
Endif   
          
TRB->(MsUnlock())

Return Nil	

// ------------------------------------------------------- \\
/*/{Protheus.doc} Processa
//TODO Processa o markbrow para gerar o arquivo texto.
@author Claudio Macedo
@since 16/03/2020 
@version 1.0
@return NIl
@type Function
/*/
// ------------------------------------------------------- \\
User Function Processa()

Local lMarcados := .F.

TRB->(DbGoTop())

While !TRB->(EOF()) .And. !lMarcados
	If !Empty(TRB->OK)
		lMarcados := .T.
	Endif
	TRB->(DbSkip())
Enddo

If !lMarcados
	Help(NIL, NIL, 'EXPORTAR ORDEM DE SEPARA��O',;
			  NIL, 'Ordens de separa��o n�o selecionadas', 1, 0, NIL, NIL, NIL, NIL, NIL,;
                  {'Selecione pelo menos uma ordem de separa��o para ser exportada.'})
	Return Nil	
Endif

If !GetMV('VI_SEQARQ', .T.)

		Help(NIL, NIL, 'PAR�METRO',;
				  NIL, 'Par�metro VI_ARQSEQ n�o encontrado', 1, 0, NIL, NIL, NIL, NIL, NIL,;
	                  {'Solicite ao administrador do sistema a inclus�o do par�metro acima.'})
		Return Nil

Endif

If MsgYesNo('Confirma a exporta��o das ordens de separa��o ?', 'Gerar arquivo texto')
	Processa({|| U_Exportar()}, 'Criando o arquivo texto','Aguarde ...')
	If lOK	
		MsgInfo('Arquivo texto ' + cArquivo + ' gerado com sucesso.')
	Else
		MsgInfo('Erro ao gerar o arquivo texto.')
	Endif
Endif

Return Nil

// ------------------------------------------------------- \\
/*/{Protheus.doc} Exportar
//TODO Fun��o para criar um arquivo texto, a partir das ordens
       de separa��o que foram marcadas no browse.
@author Claudio Macedo
@since 13/03/2020 
@version 1.0
@return NIl
@type Function
/*/
// ------------------------------------------------------- \\
User Function Exportar()

Local cPath     := GetMV('VI_LOCALOS',.F.,'C:\')
Local cNumSeq   := ''
Local nHandle   := NIl
Local cString   := ''
Local cOrdSep   := '' 

lOk := .T.

TRB->(DbGoTop())

cNumSeq  := GetMV('VI_SEQARQ')
PutMV('VI_SEQARQ',StrZero(cNumSeq+1,6))
cArquivo := Alltrim(cPath)+StrZero(cNumSeq,6)+'.txt'
nHandle  := FCREATE(cArquivo, 0)

If nHandle == -1
    MsgAlert('O Arquivo nao pode ser criado. Erro: ' + STR(FERROR()))
    lOk := .F.
    Return lOk
Endif

TRB->(DbGoTop())

cString += 'CB8_ORDSEP;CB8_PEDIDO;CB8_PROD;B1_DESC;CB8_SALDOE;CB8_LCALIZ;B1_GRUPO;B1_PESO;B1_PESBRU;'
cString += 'B1_CODBAR;C5_CLIENTE;C5_LOJACLI;C5_TRANSP;C5_TABELA;C5_VEND1;C5_EMISSAO;C5_TPFRETE;C5_VOLUME1;'
cString += 'C5_PEDECOM;C5_XCANAIS;A1_NOME;A1_PESSOA;A1_END;A1_COMPLEM;A1_BAIRRO;A1_EST;A1_CEP;A1_MUN;A1_TEL;'
cString += 'A1_CGC;A1_DTCAD;A4_NOME;CB8_ITEM;C6_PRCVEN;C6_VALOR;C6_TES;C6_CF;C6_DESCONT;C6_VALDESC;C6_PRUNIT' + CRLF

FWrite(nHandle, cString)

While !TRB->(EOF())

	If Empty(TRB->OK) 
		TRB->(DbSkip())
		Loop
	Endif

	cOrdSep := TRB->CB8_ORDSEP
	
	SC5->(DbSetOrder(1))
	If !SC5->(DbSeek(xFilial('SC5') + TRB->CB8_PEDIDO))
		TRB->(DbSkip())
		Loop	
	Endif
	
	SA1->(DbSetOrder(1))
	SA1->(DbSeek(xFilial('SA1') + SC5->C5_CLIENTE + SC5->C5_LOJACLI))

	CB8->(DbSetOrder(9))
	CB8->(DbSeek(xFilial('CB8') + TRB->CB8_ORDSEP + TRB->CB8_PEDIDO))
	
	While !CB8->(EOF()) .And. CB8->CB8_ORDSEP = TRB->CB8_ORDSEP .And. CB8->CB8_PEDIDO = TRB->CB8_PEDIDO

		SC6->(DbSetOrder(1))
		If !SC6->(DbSeek(xFilial('SC6') + CB8->CB8_PEDIDO + CB8->CB8_ITEM))
			CB8->(DbSkip())
			Loop
		Endif

		SB1->(DbSetOrder(1))
		SB1->(DbSeek(xFilial('SB1') + CB8->CB8_PROD))
	
		cString := TRB->CB8_ORDSEP + ';' +;
		           TRB->CB8_PEDIDO + ';' +;
		           CB8->CB8_PROD   + ';' +;
		           SB1->B1_DESC    + ';' +;
		           PADL(CB8->CB8_SALDOE,12) + ';' +;
		           CB8->CB8_LCALIZ + ';' +;
		           SB1->B1_GRUPO   + ';' +;
		           PADL(SB1->B1_PESO * CB8->CB8_SALDOE,11) + ';' +;
		           PADL(SB1->B1_PESBRU * CB8->CB8_SALDOE,11) + ';' +;
		           SB1->B1_CODBAR  + ';' +;
		           SC5->C5_CLIENTE + ';' +;
		           SC5->C5_LOJACLI + ';' +;
		           SC5->C5_TRANSP  + ';' +;
		           SC5->C5_TABELA  + ';' +;
		           SC5->C5_VEND1   + ';' +;
		           DTOS(SC5->C5_EMISSAO) + ';' +;
		           SC5->C5_TPFRETE + ';' +;
		           PADL(SC5->C5_VOLUME1,5) + ';' +;
		           SC5->C5_PEDECOM + ';' +;
		           SC5->C5_XCANAIS + ';' +;
		           SA1->A1_NOME    + ';' +;
		           SA1->A1_PESSOA  + ';' +;
		           SA1->A1_END     + ';' +;
		           SA1->A1_COMPLEM + ';' +;
		           SA1->A1_BAIRRO  + ';' +;
		           SA1->A1_EST     + ';' +;
		           SA1->A1_CEP     + ';' +;
		           SA1->A1_MUN     + ';' +;
		           SA1->A1_TEL     + ';' +;
		           SA1->A1_CGC     + ';' +;
		           DTOS(SA1->A1_DTCAD)   + ';' +;
		           TRB->A4_NOME    + ';' +;
		           CB8->CB8_ITEM   + ';' +;
		           PADL(SC6->C6_PRCVEN,11) + ';' +;
		           PADL(SC6->C6_VALOR,12)  + ';' +;
		           SC6->C6_TES     + ';' +;
		           SC6->C6_CF      + ';' +;
		           PADL(SC6->C6_DESCONT,5) + ';' +;
		           PADL(SC6->C6_VALDESC,14) + ';' +;
		           PADL(SC6->C6_PRUNIT,11)  + CRLF
	
	    FWrite(nHandle, cString)

	    CB8->(DbSkip())
    
    Enddo

	TRB->(DbSkip())
	
	If cOrdSep <> TRB->CB8_ORDSEP
		CB7->(DbSetOrder(1))
		If CB7->(DbSeek(xFilial('CB8') + cOrdSep))
			CB7->(reclock('CB7',.F.))
			CB7->CB7_XARQTX := 'S'
			CB7->(MsUnlock())
		Endif
		cOrdSep := TRB->CB8_ORDSEP
	Endif

Enddo

FClose(nHandle)

TRB->(DbGoTop())

While !TRB->(EOF())
	
	If !Empty(TRB->OK)
		TRB->(RecLock('TRB',.F.))
		TRB->(DbDelete())
		TRB->(MsUnlock())
	Endif   
          
	TRB->(DbSkip())
Enddo

TRB->(DbGoTop())

Return lOk

// ------------------------------------------------------- \\
/*/{Protheus.doc} AaddColuna
//TODO Fun��o para criar as colunas do grid.
@author Claudio Macedo
@since 17/03/2020 
@version 1.0
@return NIl
@type Function
/*/
// ------------------------------------------------------- \\
Static Function AaddColuna(cCampo,cTitulo,nArrData,cPicture,nAlign,nSize,nDecimal)
    
Local aColumn
Local bData   := {||}

Default nAlign   := 1
Default nSize    := 20
Default nDecimal := 0
Default nArrData := 0  
    
If nArrData > 0
    bData := &("{||" + cCampo +"}")
EndIf

/* Array da coluna
[n][01] T�tulo da coluna
[n][02] Code-Block de carga dos dados
[n][03] Tipo de dados
[n][04] M�scara
[n][05] Alinhamento (0=Centralizado, 1=Esquerda ou 2=Direita)
[n][06] Tamanho
[n][07] Decimal
[n][08] Indica se permite a edi��o
[n][09] Code-Block de valida��o da coluna ap�s a edi��o
[n][10] Indica se exibe imagem
[n][11] Code-Block de execu��o do duplo clique
[n][12] Vari�vel a ser utilizada na edi��o (ReadVar)
[n][13] Code-Block de execu��o do clique no header
[n][14] Indica se a coluna est� deletada
[n][15] Indica se a coluna ser� exibida nos detalhes do Browse
[n][16] Op��es de carga dos dados (Ex: 1=Sim, 2=N�o)
*/

aColumn := {cTitulo,bData,,cPicture,nAlign,nSize,nDecimal,.F.,{||.T.},.F.,{||.T.},NIL,{||.T.},.F.,.F.,{}}

Return {aColumn}

// ------------------------------------------------------- \\
/*/{Protheus.doc} MarkAll
//TODO Fun��o para marcar/desmarcar todos os registros do grid.
@author Claudio Macedo
@since 17/03/2020 
@version 1.0
@return NIl
@type Function
/*/
// ------------------------------------------------------- \\
Static Function MarkAll(cMarca,lMarcar)

TRB->(DbGoTop())
While !TRB->(Eof())
    TRB->(RecLock('TRB', .F.))
    TRB->OK := IIf(lMarcar, cMarca, '  ')
    TRB->(MsUnlock())
    TRB->(DbSkip())
EndDo

Return .T.
