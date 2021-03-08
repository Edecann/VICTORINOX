#include 'totvs.ch'
/*/{Protheus.doc} MTI030NOM
Ponto de entrada no fonte MATI030O - Integracao mensagem unica do Protheus
@author TOTVS
@since 19/10/2020
@version 1.0
@return aRet, Retorna aCliente e aItem do pedido.
@type function
/*/
User Function MTI030NOM()
    Local aCliente:= {}

    aCliente := F030GetCli(aClone(PARAMIXB))
Return aCliente

Static Function F030GetCli(PARAMIXB)

    Local cJson   := ''
    Local oJson   := JsonObject():New()
    Local aRet    := {}
    Local nPos    := 0
    Local ni      := 0
    Local cCpf    := ''
    Local aCli    := {}
    Local aCliente:= {}

    If Len(PARAMIXB) < 4 // se o fonte padrão MATI030O não for a versão correta retorna o cliente.
        Return PARAMIXB[1]
    Endif

    If  Valtype(PARAMIXB[3]) <> 'C' // se não vier um json return
        Return PARAMIXB[1]
    Endif

    If PARAMIXB[4] <> 3 // se for diferente de inclusão tenho que retornar o proprio array vindo do Padrão.
        Return PARAMIXB[1]
    Endif

    aCliente:= PARAMIXB[1]
    cJson   := PARAMIXB[3]

    oJson:FromJson(cJson)
    
    nPos := AT( Upper('Content'), Upper(cJson) )
    If nPos > 0 .and.;
            Valtype(oJson['Content']) $ 'C J'.and.;
            Valtype(oJson['Content']['GovernmentalInformation']) = 'A' .and.;
            Valtype(oJson['Content']['StoreId']) = 'C'

        aRet := oJson['Content']['GovernmentalInformation']
        For nI := 1 to Len(aRet)
            iF Upper(aRet[ni]["Name"]) = 'CPF' .or. Upper(aRet[ni]["Name"]) = 'CNPJ' .OR.  Upper(aRet[ni]["Name"]) = 'CGC'
                cCpf := aRet[ni]["Id"]
            Endif
        Next nI

        //cStoreId := oJson['Content']['StoreId'] // Loja esta sendo inclementada pela função GetCodCli
        aCli     := GetCodCli(cCpf)

        If !Empty(cCpf) .and. !Empty(aCli)
            nPos := aScan(aCliente, {|X| Alltrim(x[1]) == 'A1_COD'})
            If nPos > 0
                aCliente[nPos][2] := aCli[1] // Codigo do Cliente no SA1 pelo CPF
            Endif

            nPos := aScan(aCliente, {|X| Alltrim(x[1]) == 'A1_LOJA'})
            If nPos > 0
                aCliente[nPos][2] := aCli[2] // Loja Atual  + 1
            else
                aAdd(aCliente, {"A1_LOJA", aCli[2], Nil})
            Endif
			
			nPos := aScan(aCliente, {|X| Alltrim(x[1]) == 'A1_XCANAIS'})
            If nPos > 0
                aCliente[nPos][2] := '2' // VX Canal
            else
                aAdd(aCliente, {"A1_XCANAIS", '2', Nil})
            Endif

        Endif
    Endif


    FreeObj(oJson)

Return aCliente

Static Function  GetCodCli(cCpf)
    Local aRet   := {}
    Local cQuery := ''

    If Empty(cCpf)
        Return aRet
    Endif

    cAliasFK1   := GetNextAlias()
    cQuery += " SELECT  A1_COD, MAX(A1_LOJA) LOJA FROM " + RetSqlName("SA1") + " SA1 "
    cQuery += " where "
    cQuery += "     A1_FILIAL = '" + xFilial("SA1") + "' "
    cQuery += " AND A1_CGC = '" + Alltrim(cCpf) + "' "
    cQuery += " AND D_E_L_E_T_ = '' "
    cQuery += " GROUP BY A1_COD "

    MPSysOpenQuery(cQuery, cAliasFK1)
    If (cAliasFK1)->(!(EoF()))
        aRet := { (cAliasFK1)->A1_COD, Soma1((cAliasFK1)->LOJA) }
	Else
		aRet := { GETSXENUM("SA1","A1_COD") , StrZero(1, (TamSX3( 'A1_LOJA' )[1]))}
    Endif

    (cAliasFK1)->(dbCloseArea())

Return aRet
