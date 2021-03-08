//Bibliotecas
#INCLUDE "TOTVS.CH"
#INCLUDE "PROTHEUS.CH"
#include "RWMAKE.CH"
#INCLUDE "AP5MAIL.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "FWMVCDEF.CH"

/*{Protheus.doc} VIFIMP01()
Programa para chamada principal da rotina do EDI Financeiro Victor-Inox
@type    User Function
@author  Carlos H. Fernandes
@since   14/07/2020
@version 1.0
@param   none
@return  Nil
@history Mário Augusto Cavenaghi - EthosX
@        05/10/2020: Error.log com novo layout do .CSV
@        09/10/2020: Não está pegando o banco correto
@        14/10/2020: Baixando título com banco C03
*/
User Function VIFIMP01()
	//Declarar variáveis locais
	Local oDlg	    := Nil
	Local oSay01	:= Nil
	Local oSay02	:= Nil
	Local oSay03	:= Nil
	Local oSay04	:= Nil
	Local oBut01	:= Nil
	Local oBut02	:= Nil
	Local oFont01a	:= TFont():New("Arial",,18,.T.,.F.)
	Local oFont02a	:= TFont():New("Arial",,16,.T.,.F.)

	//Caixa de dialogo apresentação da rotina
	oDlg    := MSDialog():New( 000,000,260,600," EDI FINANCEIRO VICTORINOX ",,,.F.,,,,,,.T.,,,.T. )
	oSay01  := TSay():New( 015,010,{|| "Processos afetados: Baixa de titulos financeiros" },oDlg,,oFont01a,.F.,.F.,.F.,.T.,CLR_BLACK,,550,008)
	oSay02  := TSay():New( 035,010,{|| "Esta rotina tem como função realizar a conciliação dos titulos financeiros a partir dos arquivos" },oDlg,,oFont02a,.F.,.F.,.F.,.T.,,,550,008)
	oSay03  := TSay():New( 045,010,{|| "integrados conforme leiaut pré-configurado. " },oDlg,,oFont02a,.F.,.F.,.F.,.T.,,,550,008)
	oSay03  := TSay():New( 055,010,{|| "Este processo irá  alterar os status dos  registros de titulos  financeiros na tabela de titulos a" },oDlg,,oFont02a,.F.,.F.,.F.,.T.,,,550,008)
	oSay04  := TSay():New( 065,010,{|| "receber." } ,oDlg,,oFont02a,.F.,.F.,.F.,.T.,,,550,008)
	//Botões
	oBut01  := TButton():New( 100,175, "Avançar", oDlg, {|| VIPROCROT(.T.), oDlg:End() }, 050, 015, , , .F., .T., .F., , .F., , , .F. )
	oBut02  := TButton():New( 100,235, "Fechar"	, oDlg, {|| oDlg:End() }, 050, 015, , , .F., .T., .F., , .F., , , .F. )

	oDlg:Activate(,,,.T.)

Return NIL

/*/{Protheus.doc} VIPROCROT
Rotina para realização dos processos de seleção, leitura, analise e baixa dos titulos financeiros
@type  Static Function
@param none
@return return_var, return_type, return_description
/*/
Static Function VIPROCROT(lProcRot)
	//Declara variaveis locais
	Local nA
	Local lErro     := .F.
	Local cTit      := "Atenção"
	Local aArquivos := {}
	//Declara variaveis privadas
	Private oProcess
	Private nSize   := 1
	Private cMsg    := " "
	Private aOp     := {"OK"}

	If lProcRot
		//Rotina de seleção de arquivos
		MsgRun("Executando a rotina de seleção...", "Seleção de arquivos", {|| VISELARQ(@aArquivos,@lErro) })

		//Rotina de leitura, processamento e baixa dos arquivos
		If aArquivos != NIL
			oProcess := MsNewProcess():New( { |lEnd| VIPROCARQ(aArquivos,@lErro,@lEnd,oProcess) }, "Leitura de arquivos", "Aguarde, executando a leitura do(s) arquivo(s)...", .F. )
			oProcess:Activate()
		Else
			cMsg += "Não foram selecionados arquivos para importação! Processamento abortado."
		Endif
	Else
		//Mensagem em tela
		lErro := .T.
		cMsg  += "Esta rotina não foi chamada corretamente e este processo será"+CRLF
		cMsg  += "abortado. Por  favor  entre em contato com o responsável pela"+CRLF
		cMsg  += "rotina/e ou sistema."+CRLF
	Endif

	If lErro
		If Len(cMsg) > 48
			nSize := 2
		Endif
		Aviso(cTit,cMsg,aOp,nSize)
	Else
		cMsg  += ""+CRLF
		cMsg  += "A rotina foi processada corretamente."+CRLF
		cMsg  += "Analise o log de processamento para obter mais informações."
		Aviso("Concluido ",cMsg,aOp,2)
	Endif

Return

/*/{Protheus.doc} VISELARQ
Rotina para seleção dos arquivos a importar
@type  Static Function
@param aArquivos, Recebe array para ser preenchido com os arquivos selecionados
@param lErro, Recebe logico para controle de erro
@return none
/*/
Static Function VISELARQ(aArquivos,lErro)
	//Declarar variáveis locais
	Local nMascpad  := nil
	Local cDirIni   := nil
	Local lSalvar   := .F.                  /*.T. = Salva || .F. = Abre*/
	Local cTitulo   := iif(lSalvar, "Salvar arquivo","Abrir arquivo")
	Local nOpcoes   := GETF_NETWORKDRIVE+GETF_LOCALHARD+GETF_MULTISELECT
	Local lArvore   := .F.                  /*.T. = apresenta o árvore do servidor || .F. = não apresenta*/
	Local cTargetDir := ""
	local cMascara  := "Arquivos de CSV|*.csv"

	cTargetDir := cGetFile( cMascara, cTitulo, nMascpad, cDirIni, lSalvar, nOpcoes, lArvore)
	aArquivos  := Separa(cTargetDir,'|',.f.)

	If Len(aArquivos) = 0
		lErro := .T.
		aArquivos := NIL
	Endif
Return

/*/{Protheus.doc} VIPROCARQ
Rotina para leitura, processamento e baixa dos registros
@type  Static Function
@param none
@return none
/*/
Static Function VIPROCARQ(xArquivo,lErro,lEnd,oProcess)
	//Declarar variaveis locais
	Local oFile
	local nX        := 0
	Local n         := 0
	Local nY        := 0
	Local nLin      := 0
	Local nPosCod   := 0
	Local nValBaixa := 0
	Local nTxOper   := 0
	Local nValAnt   := 0
	Local nValMDR   := 0
	Local nValTitL  := 0
	Local nValTitB  := 0
	Local lParcial 	:= .F.
	Local lGeraNDF 	:= .F.
	Local lSemParc  := .F.
	Local lAbort    := .F.
	Local aCabec    := {}
	Local aLinha    := {}
	Local aTitSE1   := {}
	Local aTitNDF   := {}
	Local cCpo      := ""
	Local cCategor  := ""
	Local cParcela  := ""
	Local cTipo     := ""
	Local cMsg2     := ""
	Local xAux      := ""
	Local dBaixa    := CtoD(" ")
	Local nTamID    := SuperGetMV("VI_LTTAMCP" , .T., 12)                  //Parametro com o tamanho do campo ID na tabela de titulos a receber
	Local cArqLog   := GetTempPath()+"LOGPROC_BAIXA_"+StrTran(DtoC(dDataBase),"/","")+"_"+StrTran(Time(),":","")+".txt"
	Local oFileLog  := FwFileWriter():New(cArqLog,.T.)

	//Realiza a junção das tabelas
	VIDEFLAY(xArquivo,cMsg,cMsg2,@oFile,@oFileLog,@aCabec,@aLinha,@oProcess,cArqLog)

	//Abertura da tabela de titulos
	dbSelectArea("SE1")
	SE1->(DbSetOrder(1))

	//Abertura do log de gravação
	cMsg := "["+DtoC(Date())+"]*** Arquivo de Log para Processamento de baixa financeira ***"+CRLF
	VILOGPROC(nLin,nPosCod,"",cArqLog,@cMsg,@cMsg2,cCpo,cCategor,@oFileLog)

	If Len(aLinha) > 0
		//***********************************************
		//Inicio do bloco para processamento da baixa
		//***********************************************
		nPosCod  := aScan(aCabec,{|x| UPPER(NoAcento(Alltrim(x))) == 'COD_ID'})    //Campo contendo o numero do titulo Protheus
		nPosCat  := aScan(aCabec,{|x| UPPER(NoAcento(Alltrim(x))) == "CATEGORIA"})
		nPosParc := aScan(aCabec,{|x| UPPER(NoAcento(Alltrim(x))) == "NDAPARCELA"})
		nPosValL := aScan(aCabec,{|x| UPPER(NoAcento(Alltrim(x))) == "VALORLIQUIDO"})
		nPosValB := aScan(aCabec,{|x| UPPER(NoAcento(Alltrim(x))) == "VALORBRUTO"})
		nPosDtBx := aScan(aCabec,{|x| UPPER(NoAcento(Alltrim(x))) == "DATADOULTIMOSTATUS"})
		nPosAntec:= aScan(aCabec,{|x| UPPER(NoAcento(Alltrim(x))) == "DESCONTODEANTECIPACAO"})
		nPosMDR  := aScan(aCabec,{|x| UPPER(NoAcento(Alltrim(x))) == "DESCONTODEMDR"})
		nPosTipo := aScan(aCabec,{|x| UPPER(NoAcento(Alltrim(x))) == "TIPO"})

		oProcess:SetRegua1(Len(aLinha))
		oProcess:SetRegua2(2)
		//Localiza os registro de venda e cancelamento e marca para não baixar
		xCont:= 0
		aAux := {}
		aNBaixa := {}
		aAux := aClone(aLinha)
		oProcess:IncRegua2("Aguarde... ")
		For nX := 1 To Len(aLinha)
			oProcess:IncRegua1("Procurando titulos cancelados... " + cValToChar(nX)+ " de " + cValToChar(Len(aLinha)))
			xCont := 0
			For nY := 1 To Len(aAux)
				If (aAux[nY][nPosCod] == aLinha[nX][nPosCod] .And. aAux[nY][nPosParc] == aLinha[nX][nPosParc])
					If xCont < 1
						aAux[nY][nPosCod] := "X"+aAux[nX][nPosCod]
						xCont++
					Else
						aAdd(aNBaixa,{aLinha[nX][nPosCat],aLinha[nX][nPosCod],aLinha[nX][nPosParc],aLinha[nX][nPosValL],nX})
						aAdd(aNBaixa,{aAux[nY][nPosCat],aAux[nY][nPosCod],aAux[nY][nPosParc],aAux[nY][nPosValL],nY})
					Endif
				Endif
			Next nY
		Next nX
		oProcess:IncRegua1()
		oProcess:IncRegua2()

		//Abertura do log de gravação
		cMsg := ""+CRLF
		cMsg += "["+DtoC(Date())+"]*** Arquivo de Log para Processamento dos Cancelamentos ***"+CRLF
		VILOGPROC(0,0,"",cArqLog,@cMsg,@cMsg2,"","",@oFileLog)

		oProcess:SetRegua2(Len(aNBaixa))
		For nZ := 1 To Len(aNBaixa)
			oProcess:IncRegua2("Gerando Log de titulos Cancelados...")
			cMsg  := "Não Baixado"
			cMsg2 := "O Titulo não sera baixado - Categoria "+aNBaixa[nZ][1]
			VILOGPROC(aNBaixa[nZ][5],2,aNBaixa[nZ],cArqLog,@cMsg,@cMsg2,"","",@oFileLog)
		Next nZ
		oProcess:IncRegua2()


		//Inicia processo de gravação
		cMsg := ""+CRLF
		cMsg += "["+DtoC(Date())+"]*** Arquivo de Log para Processamento das Baixas ***"+CRLF
		VILOGPROC(0,0,"",cArqLog,@cMsg,@cMsg2,"","",@oFileLog)

		oProcess:SetRegua1(Len(aLinha))
		For nW := 1 To Len(aLinha)
			If (aScan(aNBaixa,{|x| (x[2]) == aLinha[nW][nPosCod]}) = 0) //Para não processar registros do aNBaixa
				nLin++
				oProcess:IncRegua1("Processando titulos para baixa " + cValToChar(nLin)+ " de " + cValToChar(Len(aLinha)))
				oProcess:IncRegua2("Baixando...")
				//Prepara variaveis
				cId      := Alltrim(NoAcento(SubStr(aLinha[nW][nPosCod],1,nTamID)))
				cCategor := Alltrim(NoAcento(aLinha[nW][nPosCat]))
				If nPosParc = 0
					lSemParc := .T. //Verificar com usuário se tem parcela
				Else
					cParcela := NoAcento(aLinha[nW][nPosParc])
				Endif
				nValTitL := Round(Val(StrTran(StrTran(aLinha[nW][nPosValL],'"',""),",",".")),2)
				nValTitB := Round(Val(StrTran(StrTran(aLinha[nW][nPosValB],'"',""),",",".")),2)
				cTipo    := aLinha[nW][nPosTipo]
				dBaixa   := CtoD(SubStr(StrTran(aLinha[nW][nPosDtBx],'"',""),1,10))

				IF !(UPPER(cCategor) $ 'CHARGEBACK|CANCELAMENTOPARCIAL|CANCELAMENTO|BALANCEAMENTODESALDO')
					//Rotina para busca do titulo na SE1
					IF VIGETTIT(cId,cParcela,lSemParc,cTipo,@aTitSE1,@oProcess)
						//Se encontrou o titulo classifica o tipo da baixa
						DbSelectArea("SE1")
						SE1->(DbsetOrder(1))
						For nQ := 1 To Len(aTitSE1) //Se houver mais de um titulo a baixar

							If SE1->(Dbseek(aTitSE1[nQ,1]))
								//Classificação da baixa
								If ( nValTitL > SE1->E1_SALDO ) //Valor do arquivo maior que o titulo
									nValBaixa   := SE1->E1_SALDO                //valor a ser baixado
									nValTitL    := nValTitL - SE1->E1_SALDO      //caso exista residuo para RA
								ElseIf ( SE1->E1_SALDO > nValTitL ) //valor do titulo maior que o arquivo
									nValBaixa   := SE1->E1_SALDO//valor a ser baixado
									nValTitL 	:= 0
									lParcial    := .T.
									lGeraNDF    := .T.
								ElseIf ( SE1->E1_SALDO == nValTitL )
									nValBaixa   := SE1->E1_SALDO
									nValTitL    := 0
								Endif

								Begin Transaction
									If nValBaixa > 0
										//Rotina para Baixa financeira
										If VIBAIXATIT(dBaixa,nValBaixa,@lParcial,@lAbort)
											cMsg  := "Baixado"
											cMsg2 := "Titulo "+Iif(lParcial,"baixado parcialmente","baixado")+" na tabela SE1 - Contas a receber"
											VILOGPROC(nLin,nPosCod,aLinha[nW],cArqLog,@cMsg,@cMsg2,cCpo,cCategor,@oFileLog)
											oProcess:IncRegua2("Baixando...")
										Else
											cMsg  := "Erro na baixa"
											cMsg2 := "O Titulo não foi baixado devido erro na execauto de baixa"
											VILOGPROC(nLin,nPosCod,aLinha[nW],cArqLog,@cMsg,@cMsg2,cCpo,cCategor,@oFileLog)
											oProcess:IncRegua2("Baixando...")
										Endif

										//Totalizador para geração de NDF desmembrada no final do processo
										If lParcial
											nValAnt += Val(StrTran(StrTran(aLinha[nW][nPosAntec],'"',""),",","."))
											nValMDR += Val(StrTran(StrTran(aLinha[nW][nPosMDR],'"',""),",","."))
											lParcial := .F.
										Endif
									Else
										cMsg  := "Não Baixado"
										cMsg2 := "Titulo foi baixado anteriormente"
										VILOGPROC(nLin,nPosCod,aLinha[nW],cArqLog,@cMsg,@cMsg2,cCpo,cCategor,@oFileLog)
										oProcess:IncRegua2("Baixando...")
									Endif

								End Transaction
							Endif
							aTitSE1 := {}
							//Tratamento de erro para abortar operação
							If lAbort
								cMsg := " "+CRLF
								cMsg += "Processo Abortado. Desculpe!"
								lGeraNDF := .F.
								lErro := .T.
								oFile:Close()
								Return .F.
							Endif

						Next nQ
					Else
						//Rotina para gravação do Log de Processamento
						cMsg  := "Não Localizado"
						cMsg2 := "Titulo não localizado na tabela SE1 - Contas a receber"
						VILOGPROC(nLin,nPosCod,aLinha[nW],cArqLog,@cMsg,@cMsg2,cCpo,cCategor,@oFileLog)
						oProcess:IncRegua2("Baixando...")
					Endif
					//**********************************************
					//Fim do bloco para processamento da baixa
					//**********************************************
				Else
					//Rotina para gravação do Log de Processamento
					cMsg  := cCategor
					cMsg2 := "Titulo não sera baixado neste momento"
					VILOGPROC(nLin,nPosCod,aLinha[nW],cArqLog,@cMsg,@cMsg2,cCpo,cCategor,@oFileLog)
					oProcess:IncRegua2("Baixando...")
				Endif
			Endif
		Next nW
		//Finalizou todo processo gera NDF desmembrada
		If lGeraNDF
			If nValAnt > 0 .Or. nValMDR > 0
				aTitNDF := {{"Ant",nValAnt},{"MDR",nValMDR}}
				For n := 1 To Len(aTitNDF)
					If VIGERATIT(aTitNDF[n])
						cMsg  := "Titulo "+aTitNDF[n][1]
						cMsg2 := "Titulo gerado no valor de R$ "+cValToChar(aTitNDF[n][2])+" a partir da baixa parcial na tabela SE2 - Contas a Pagar"
						VILOGPROC(nLin,nPosCod,aLinha[n],cArqLog,@cMsg,@cMsg2,cCpo,cCategor,@oFileLog)
						oProcess:IncRegua2("Baixando...")
					Else
						cMsg  := "Erro na geração"
						cMsg2 := "O Titulo não foi gerado devido a erro na execauto de titulo"
						VILOGPROC(nLin,nPosCod,aLinha[n],cArqLog,@cMsg,@cMsg2,cCpo,cCategor,@oFileLog)
						oProcess:IncRegua2("Baixando...")
					Endif
				Next n
			Endif
		Endif

		oFile:Close()

	Else
		lErro := .T.
		cMsg  += "Não existem registros para processar a baixa."
		VILOGPROC(nLin,nPosCod,,cArqLog,@cMsg,@cMsg2,cCpo,cCategor,@oFileLog)
		Return .F.
	Endif
*/
	//Fecha o arquivo de Log e envia por email
	oFileLog:Close()
	VIENVMAIL(cArqLog)

Return

/*/{Protheus.doc} VIDEFLAY
Rotina para definição dos layouts e ajustes nos arrays quando MundiPag+Stones/ ou apenas Stones
@type  Static Function
@param none
@return none
/*/
Static Function VIDEFLAY(xArquivo,cMsg,cMsg2,oFile,oFileLog,aCabec,aLinha,oProcess,cArqLog)
	//Declarar variaveis locais
	Local nA       := 0
	Local nB       := 0
	Local nC       := 0
	Local nD       := 0
	Local nE       := 0
	Local nF       := 0
	Local nPosStone:= 0  //Posição do campo chave Stone
	Local nPosMunPg:= 0  //Posição do campo chave Mundipag
	Local cCpo     := "" //String de campo para pesquisa no aScan
	Local xArray   := "" //String de nome do array para macro substituição
	Local xAux     := {} //Array Simples auxiliar
	Local aAux1    := {} //Array Simples auxiliar
	Local aAux2    := {} //Array Simples auxiliar
	Local aCabAux1 := {} //Array Simples auxiliar
	Local aCabAux2 := {} //Array Simples auxiliar
	Local aAuxMunPg:= {} //Array Simples
	Local aCabMunPg:= {} //Array Simples
	Local aLinMunPg:= {} //Array Multi-dimensional
	Local aAuxStone:= {} //Array Simples
	Local aCabStone:= {} //Array simples
	Local aLinStone:= {} //Array Multi-dimensional
	Local aNAchou  := {} //Array Multi-dimensional
	Local lErro    := .F.

	//Abertura do log de gravação
	cMsg := "["+DtoC(Date())+"]*** Arquivo de Log para Processamento do ProcV entre tabelas ***"+CRLF
	VILOGPROC(0,0,"",cArqLog,@cMsg,@cMsg2,"","",@oFileLog)

	oProcess:SetRegua1(4)
	oProcess:IncRegua1("Processando registros...")

	//Efetua a leitura do(s) arquivo(s)
	oProcess:SetRegua2(Len(xArquivo))
	For nA := 1 To Len(xArquivo)
		oProcess:IncRegua2("Efetuando a leitura dos registros...")
		//Realiza a leitura do arquivo
		xArray := ("aAux"+cValToChar(nA)) //Faz a atribuição pela ordem dos arquivos retornados e pela quantidade existente no array
		oFile  := FwFileReader():New(Alltrim(xArquivo[nA]))
		If oFile:Open()
			oFile:SetBufferSize(4096)
			&xArray:= oFile:GetAllLine()
			oFile:Close()
		Else
			lErro := .T.
			cMsg  += "Erro na abertura do arquivo!" +CRLF
			cMsg  += "Não foi possivel realizar a abertura do arquivo de importação." +CRLF
			cMsg  += "Operação abortada."
			VILOGPROC("","","",cArqLog,@cMsg,@cMsg2,"",@oFileLog)
			Return .F.
		Endif
	Next nA

	//Gravação dos arquivos em array, quando houver layout da Mundipag, fara a intersecção
	//entre os arquivos para localizar e gravar o codigo NSU que existe apenas no arquivo da Stone.
	oProcess:SetRegua2(2)
	If Len(xArquivo) > 1 .And. Len(xArquivo) <= 2
		oProcess:IncRegua1()
		//Seta arrays localizando onde estão os arquivos de layout
		If (Len(aAux1) > 0 .And. Len(aAux2) > 0 )
			oProcess:IncRegua2("Analisando...")
			cCpo    := NoAcento(SuperGetMV("VI_LTBOCPO"  , .T., 'BOLETOTID')) //Layout Mundipag
			//Primeira linha devera ser o cabeçalho obrigatoriamente, arrays utilizados para melhora da performance do aScan
			aCabAux1:= Separa(aAux1[1],";",.T.)
			aCabAux2:= Separa(aAux2[1],";",.T.)
			If (aScan(aCabAux1,{|x| UPPER(NoAcento(Alltrim(x))) == NoAcento(cCpo)})) = 0
				If (aScan(aCabAux2,{|x| UPPER(NoAcento(Alltrim(x))) == NoAcento(cCpo)})) = 0
					lErro := .T.
				Else
					//Layout MundiPag
					aAuxMunPg  := aclone(aDel(aAux2,1))
					nPosMunPg  := aScan(aCabAux2,{|x| UPPER(NoAcento(Alltrim(x))) == NoAcento(cCpo)})    //Campo contendo o numero do titulo Protheus
					aCabMunPg  := aClone(aCabAux2)
					//Layout Stone
					cCpo := NoAcento(SuperGetMV("VI_LTCACPO"  , .T., 'STONEID')) //Layout Stone
					If (aScan(aCabAux1,{|x| UPPER(NoAcento(Alltrim(x))) == NoAcento(cCpo)})) = 0
						lErro := .T.
					Else
						aAuxStone := aclone(aDel(aAux1,1))
						nPosStone := aScan(aCabAux1,{|x| UPPER(NoAcento(Alltrim(x))) == NoAcento(cCpo)}) //Campo contendo o numero do titulo Protheus
						aCabStone := aClone(aCabAux1)
					Endif
				Endif
			Else
				//Layout MundiPag
				aAuxMunPg := aclone(aDel(aAux1,1))
				nPosMunPg := aScan(aCabAux1,{|x| UPPER(NoAcento(Alltrim(x))) == NoAcento(cCpo)})    //Campo contendo o numero do titulo Protheus
				aCabMunpg := aClone(aCabAux1)
				//Layout Stone
				cCpo := NoAcento(SuperGetMV("VI_LTCACPO"  , .T., 'STONEID')) //Layout Stone
				If (aScan(aCabAux2,{|x| UPPER(NoAcento(Alltrim(x))) == NoAcento(cCpo)})) = 0
					lErro := .T.
				Else
					aAuxStone := aclone(aDel(aAux2,1))
					nPosStone := aScan(aCabAux2,{|x| UPPER(NoAcento(Alltrim(x))) == NoAcento(cCpo)}) //Campo contendo o numero do titulo Protheus
					aCabStone := aClone(aCabAux2)
				Endif
			Endif
		Endif

	ElseIf Len(xArquivo) = 1
		oProcess:IncRegua1()
		oProcess:IncRegua2("Analisando...")

		cCpo := NoAcento(SuperGetMV("VI_LTCACPO"  , .T., 'STONEID')) //Layout Stone
		aCabAux1:= Separa(aAux1[1],";",.T.)
		If (aScan(aCabAux1,{|x| UPPER(NoAcento(Alltrim(x))) == NoAcento(cCpo)})) = 0
			lErro := .T.
		Else
			//Layout Stone
			aAuxStone := aclone(aDel(aAux1,1))
			nPosStone := aScan(aCabAux1,{|x| UPPER(NoAcento(Alltrim(x))) == NoAcento(cCpo)})    //Campo contendo o numero do titulo Protheus
			aCabStone := aClone(aCabAux1)
		Endif
	Endif
	oProcess:IncRegua2("Analisando...")

	//Geração do Array para manipulação
	If !lErro
		If Len(aAuxStone) > 0
			oProcess:IncRegua1()
			oProcess:SetRegua2(Len(aAuxStone)-1)

			nPos1  := aScan(aCabStone,{|x| UPPER(NoAcento(Alltrim(x))) == 'STONEID'})
			nPos2  := aScan(aCabStone,{|x| UPPER(NoAcento(Alltrim(x))) == 'CATEGORIA'})
			nPos3  := aScan(aCabStone,{|x| UPPER(NoAcento(Alltrim(x))) == 'BANDEIRA'})
			nPos4  := aScan(aCabStone,{|x| UPPER(NoAcento(Alltrim(x))) == 'NDAPARCELA'})
			nPos5  := aScan(aCabStone,{|x| UPPER(NoAcento(Alltrim(x))) == 'VALORBRUTO'})
			nPos6  := aScan(aCabStone,{|x| UPPER(NoAcento(Alltrim(x))) == 'VALORLIQUIDO'})
			nPos7  := aScan(aCabStone,{|x| UPPER(NoAcento(Alltrim(x))) == 'DESCONTODEANTECIPACAO'})
			nPos8  := aScan(aCabStone,{|x| UPPER(NoAcento(Alltrim(x))) == 'DESCONTODEMDR'})
			nPos9  := aScan(aCabStone,{|x| UPPER(NoAcento(Alltrim(x))) == 'ULTIMOSTATUS'})
			nPos10 := aScan(aCabStone,{|x| UPPER(NoAcento(Alltrim(x))) == 'DATADOULTIMOSTATUS'})
			nPos11 := aScan(aCabStone,{|x| UPPER(NoAcento(Alltrim(x))) == 'STONECODE'})
			aSize(aAuxStone,Len(aAuxStone)-1)
			For nB := 1 To Len(aAuxStone)
				oProcess:IncRegua2("Processando layout Stone...")
				xAux := Separa(StrTran(aAuxStone[nB],'"',""),";",.T.)
				If !Empty(xAux[nPos1]) //Para eliminar registro de "balanceamento de carga"
					aAdd(aLinStone,{xAux[nPos1],{xAux[nPos2],xAux[nPos3],xAux[nPos4],xAux[nPos5],xAux[nPos6],;
						xAux[nPos7],xAux[nPos8],xAux[nPos9],xAux[nPos10],xAux[nPos11],xAux[nPos1],"CC"}})
				Endif
			Next nB
			xAux := {UPPER(NoAcento(aCabStone[nPos2])) ,UPPER(NoAcento(aCabStone[nPos3])),UPPER(NoAcento(aCabStone[nPos4])) ,;
				UPPER(NoAcento(aCabStone[nPos5])) ,UPPER(NoAcento(aCabStone[nPos6])),UPPER(NoAcento(aCabStone[nPos7])) ,;
				UPPER(NoAcento(aCabStone[nPos8])) ,UPPER(NoAcento(aCabStone[nPos9])),UPPER(NoAcento(aCabStone[nPos10])),;
				UPPER(NoAcento(aCabStone[nPos11])),"COD_ID","TIPO"}
			aCabStone := aClone(xAux)

			If Len(aAuxMunPg) > 0
				oProcess:SetRegua2(Len(aAuxStone)-1)

				nPos1  := aScan(aCabMunPg,{|x| UPPER(NoAcento(Alltrim(x))) == 'BOLETOTID'})
				nPos2  := aScan(aCabMunPg,{|x| UPPER(NoAcento(Alltrim(x))) == 'TRANSACTIONTYPE'})
				nPos3  := aScan(aCabMunPg,{|x| UPPER(NoAcento(Alltrim(x))) == 'PROCESSOR'})
				nPos4  := aScan(aCabMunPg,{|x| UPPER(NoAcento(Alltrim(x))) == 'INSTALLMENTCOUNT'})
				nPos5  := aScan(aCabMunPg,{|x| UPPER(NoAcento(Alltrim(x))) == 'BUYERNAME'})
				nPos6  := aScan(aCabMunPg,{|x| UPPER(NoAcento(Alltrim(x))) == 'BUYERDOCUMENTNUMBER'})
				nPos7  := aScan(aCabMunPg,{|x| UPPER(NoAcento(Alltrim(x))) == 'NOSSONUMERO'})
				nPos8  := aScan(aCabMunPg,{|x| UPPER(NoAcento(Alltrim(x))) == 'NSU'})
				nPos9  := aScan(aCabMunPg,{|x| UPPER(NoAcento(Alltrim(x))) == 'BOLETOPAYMENTDATE'})
				nPos10 := aScan(aCabMunPg,{|x| UPPER(NoAcento(Alltrim(x))) == 'BOLETOCREDITDATE'})
				nPos11 := aScan(aCabMunPg,{|x| UPPER(NoAcento(Alltrim(x))) == 'MERCHANTNAME'})
				aSize(aAuxMunPg,Len(aAuxMunPg)-1)
				For nC := 1 To Len(aAuxMunPg)
					oProcess:IncRegua2("Processando layout MundiPag...")
					xAux := Separa(StrTran(aAuxMunPg[nC],'"',""),";",.T.)
					aAdd(aLinMunPg,{xAux[nPos1],{xAux[nPos2],xAux[nPos3],xAux[nPos4],xAux[nPos5],xAux[nPos6],;
						xAux[nPos7],xAux[nPos8],xAux[nPos9],xAux[nPos10],xAux[nPos11],"FI"}})
				Next nC
				xAux := {NoAcento(aCabMunPg[nPos1]) ,NoAcento(aCabMunPg[nPos2]) ,NoAcento(aCabMunPg[nPos3]),;
					NoAcento(aCabMunPg[nPos4]) ,NoAcento(aCabMunPg[nPos5]) ,NoAcento(aCabMunPg[nPos6]),;
					NoAcento(aCabMunPg[nPos7]) ,NoAcento(aCabMunPg[nPos8]) ,NoAcento(aCabMunPg[nPos9]),;
					NoAcento(aCabMunPg[nPos10]),NoAcento(aCabMunPg[nPos11])}
				aCabMunPg := aClone(xAux)
			Endif
		Endif
		//Reseta arrays não mais utilizados
		xAux     := {}
		aAux1    := {}
		aAux2    := {}
		aCabAux1 := {}
		aCabAux2 := {}
		aAuxMunPg:= {}
		aAuxStone:= {}

		//Realiza o "ProcV" entre as duas tabelas gerando apenas 1 unico arquivo para baixas já contendo boletos e cartão
		//despresando tudo que for Chargeback, Cancelamento parcial, cancelamento total
		oProcess:IncRegua1()
		oProcess:SetRegua2(Len(aLinMunPg)-1)
		For nD := 1 To Len(aLinMunPg)
			oProcess:IncRegua2("Realizando ProcV entre layouts...")
			If !Empty(aLinMunPg[nD][1])
				nPosTit := aScan(aLinStone,{|x| (x[1]) == aLinMunPg[nD][1]})
				If nPosTit > 0
					aLinStone[nPosTit][2][11] := aLinMunPg[nD][2][6]
					aLinStone[nPosTit][2][12] := aLinMunPg[nD][2][11]
				Else
					aAdd(aNAchou,{aLinMunPg[nD][1],nD,aLinMunPg[nD][2][1],aLinMunPg[nD][2][4],aLinMunPg[nD][2][5]})
				Endif
			Endif
		Next nD
		//Ordena o Array por STONEID (Array multidimensional) - Crescente
		aSort(aLinStone,,, {|x, y| x[1] < y[1]})

		//Grava Log de Boletos não localizados no ProcV
		oProcess:SetRegua2(Len(aNAchou)-1)
		If Len(aNAchou) > 0
			For nE := 1 To Len(aNAchou)
				oProcess:IncRegua2("Gerando log dos titulos não encontrados...")
				cMsg  := "Não Localizado"
				cMsg2 := "O boleto não foi localizado na tabela Stone"
				VILOGPROC(aNAchou[nE][2],1,aNAchou[nE],cArqLog,@cMsg,@cMsg2,"BOLETOTID","Boleto",@oFileLog)
			Next nE
		Endif

		//Devolve array aCabec e aLinha para baixa
		aCabec := aClone(aCabStone)
		For nF := 1 To Len(aLinStone)
			aAdd(aLinha,aClone(aLinStone[nF][2]))
		Next nF

	Endif

	//Gera log de erro
	If lErro
		cMsg  += "Erro nos arquivos!" +CRLF
		cMsg  += "Os arquivos selecionados não correspondem aos arquivos corretos para baixa." +CRLF
		cMsg  += "Operação abortada."
		VILOGPROC("","","",cArqLog,@cMsg,@cMsg2,"","",@oFileLog)
		Return .F.
	Endif

Return

/*/{Protheus.doc} VIGETTIT
Rotina para busca dos titulos financeiros
@type  Static Function
@param none
@return none
/*/
Static Function VIGETTIT(cId,cParcela,lSemParc,cTipo,aTitSE1,oProcess)
	//Declarar variaveis locais
	Local cQry      := ""
	Local lRet      := .F.
	Local cSE1      := "TMPX"

	cQry := " SELECT  " + CRLF
	cQry += " E1_FILIAL, E1_PREFIXO, E1_NUM, E1_PARCELA, E1_TIPO" + CRLF
	cQry += " FROM " + RetSqlName("SE1") +  CRLF
	cQry += " WHERE E1_FILIAL = '" + xFilial("SE1") + "' " + CRLF
	cQry += " AND D_E_L_E_T_ = ' '"
	cQry += " AND E1_TIPO = '"+cTipo+"'"
	cQry += " AND E1_NSUTEF = '"+cId+"'"
	cQry += " AND (E1_PARCELA = '"+If(Len(cParcela)=1,'0'+cParcela,cParcela)+"'"
	If lSemParc
		cQry += " OR E1_PARCELA = ''"
	Endif
	cQry += " )"
	cQry := ChangeQuery(cQry)
	If Select(cSE1)<>0
		(cSE1)->(dbCloseArea())
	EndIf
	dbUseArea(.T.,'TOPCONN', TCGenQry(,,cQry), cSE1,.F.,.T.)

	While !((cSE1)->(EOF()))
		lRet := .T.
		AADD(aTitSE1,{(cSE1)->(E1_FILIAL + E1_PREFIXO + E1_NUM + E1_PARCELA + E1_TIPO)})
		(cSE1)->(dbSkip())
	EndDo
	(cSE1)->(DbCloseArea())
	oProcess:IncRegua2("Baixando...")

Return lRet


/*/{Protheus.doc} VIBAIXATIT
Rotina para baixa dos titulos
@type  Static Function
@param none
@return none
/*/
Static Function VIBAIXATIT(cDat,nValBaixa,lParcial,lAbort)

	Local aBaixa   := {}
	Local aAreaSA6 := SA6->(GetArea())
	Local _cBanco  := PadR(SuperGetMV("VI_PAGBANC"    , .T., 'CX1' ), Len(SA6->A6_COD))	//	NoAcento(SuperGetMV("VI_PAGBANC"  , .T., 'CX1'))
	Local _cAgenc  := PadR(SuperGetMV("VI_PAGAGEN"    , .T., '2139'), Len(SA6->A6_AGENCIA))	//	NoAcento(SuperGetMV("VI_PAGAGEN"  , .T., '2139')) + Space(5 - Len(SuperGetMV("VI_PAGAGEN"  , .T., '00000')))
	Local _cCtaTr  := Alltrim(SuperGetMV("VI_PAGCONT" , .T., '45004'))	//	NoAcento(SuperGetMV("VI_PAGCONT"  , .T., '45004'))
	Local _cMotBx  := NoAcento(SuperGetMV("VI_PAGMOTB", .T., 'AUT'))
	Local _dData   := dDataBase
	Local _dDtUtil := DataValida(_dData,.t.)
	Local _cHist   := 'BAIXA AUTOMATICA ROTINA VIFIMP01'
	Local _nJuros  := 0
	Local _nVlBru  := (nValBaixa) + _nJuros
	Local lRet     := "ERR"
	Local nLoop
	Local cMsgErro	:= ""

	Private lMsErroAuto		:= .F. //Determina se houve algum tipo de erro durante a execucao do ExecAuto
	Private lMsHelpAuto		:= .T. //Define se mostra ou não os erros na tela (T= Nao mostra; F=Mostra)
	Private lAutoErrNoFile	:= .T. //Habilita a gravacao de erro da rotina automatica

	lMsErroAuto := .F.

	SA6->(DbSetOrder(1))
	If SA6->(DbSeek(xFilial("SA6") + _cBanco + _cAgenc + _cCtaTr))
		_cCtaTr := Left(_cCtaTr, Len(_cCtaTr) - 1)	//	Tira o DV da conta
		SA6->(DbSeek(xFilial("SA6") + _cBanco + _cAgenc + _cCtaTr))
	Endif
	If ! SA6->(Eof())
		aBaixa :={  ;
			{"E1_FILIAL"   , xFilial("SE1")       , Nil}, ;
			{"E1_PREFIXO"  , SE1->E1_PREFIXO      , Nil}, ;
			{"E1_NUM"      , SE1->E1_NUM          , Nil}, ;
			{"E1_PARCELA"  , SE1->E1_PARCELA      , Nil}, ;
			{"E1_TIPO"     , Alltrim(SE1->E1_TIPO), Nil}, ;
			{"E1_CLIENTE"  , SE1->E1_CLIENTE      , Nil}, ;
			{"E1_LOJA"     , SE1->E1_LOJA         , Nil}, ;
			{"AUTJUROS"    , 0                    , Nil}, ;
			{"AUTMULTA"    , _nJuros              , Nil}, ;
			{"AUTVALREC"   , _nVlBru              , Nil}, ;
			{"AUTMOTBX"    , _cMotBx              , Nil}, ;
			{"AUTDTBAIXA"  , _dData               , Nil}, ;
			{"AUTHIST"     , _cHist               , Nil}, ;
			{"AUTDTCREDITO", _dDtUtil             , Nil}, ;
			{"AUTBANCO"    , SA6->A6_COD          , Nil}, ;
			{"AUTAGENCIA"  , SA6->A6_AGENCIA      , Nil}, ;
			{"AUTCONTA"    , SA6->A6_NUMCON       , Nil} } 
		MSExecAuto({|x,y| fina070(x,y)},aBaixa,3) //Inclusao
		If lMsErroAuto
			aErrPCAuto	:= GETAUTOGRLOG()
			cMsgErro	:= ""
			For nLoop := 1 To Len(aErrPCAuto)
				cMsgErro += aErrPCAuto[nLoop]+CRLF
			Next
			Aviso("ERRO",cMsgErro,{'Fechar'},2)
			MSGINFO("Foram encontradas inconsistencias que não permitem o seguimento do processo." ,"Atenção")
			lParcial:= .F.
			lAbort  := .T.
			lRet    := .F.
		Else
			SE1->(Reclock("SE1"))
			SE1->E1_PORTADO := SA6->A6_COD
			SE1->E1_AGEDEP  := SA6->A6_AGENCIA
			SE1->E1_CONTA   := SA6->A6_NUMCON
			SE1->(msUnLock())
			lRet := .T.
		Endif
	Else
		Alert("Dados bancarios incorretos. O digito verificador da conta no parametro, não confere com os dados da conta localizada.")
		MSGINFO("O digito verificador deve ser informado junto ao numero da conta, caso não o tenha feito." ,"Atenção")
		lParcial:= .F.
		lAbort  := .T.
		lRet    := .F.
	Endif
	RestArea(aAreaSA6)

Return lRet

/*/{Protheus.doc} VIGERATIT
Rotina para gerar novo titulo financeiro a partir da baixa parcial
@type  Static Function
@param none
@return none
/*/
Static Function VIGERATIT(aTxOper)
	//Declarar variaveis locais
	Local cFornec	:= getmv('VI_NDFFORN',.T.,'000309')
	Local cLoja		:= getmv('VI_NDFLOJA',.T.,'01')
	Local cPrefix   := UPPER(aTxOper[1])
	Local cTipo		:= getmv('VI_NDFTIPO',.T.,'NDF')
	Local cNaturez	:= getmv('VI_NDFNATU',.T.,'201006')
	Local nTxOper   := aTxOper[2]
	Local lRet      := .F.
	Local cMsgErro  := ""
	Local nLoop

	Private lMsErroAuto		:= .F. //Determina se houve algum tipo de erro durante a execucao do ExecAuto
	Private lMsHelpAuto		:= .T. //Define se mostra ou não os erros na tela (T= Nao mostra; F=Mostra)
	Private lAutoErrNoFile	:= .T. //Habilita a gravacao de erro da rotina automatica

	If nTxOper > 0
		aVetor :=   {	;
			{"E2_PREFIXO" 	,cPrefix								,Nil},;
			{"E2_NUM"	  	,GetSxeNum("SE2","E2_NUM")  			,Nil},;
			{"E2_PARCELA" 	,''										,Nil},;
			{"E2_TIPO"	  	,cTipo									,Nil},;
			{"E2_NATUREZ" 	,cNaturez								,Nil},;
			{"E2_FORNECE" 	,cFornec								,Nil},;
			{"E2_LOJA"	  	,cLoja									,Nil},;
			{"E2_EMISSAO" 	,dDataBase								,Nil},;
			{"E2_VENCTO"  	,dDataBase								,Nil},;
			{"E2_VENCREA" 	,dDataBase								,Nil},;
			{"E2_VALOR"		,nTxOper		 						,Nil}}

		MSExecAuto({|x,y| Fina050(x,y)},aVetor,3) //Inclusao

		If lMsErroAuto
			aErrPCAuto	:= GETAUTOGRLOG()
			cMsgErro	:= ""
			For nLoop := 1 To Len(aErrPCAuto)
				cMsgErro += aErrPCAuto[nLoop]+ "<br>"
			Next
		Else
			ConfirmSX8()
			lRet := .T.
		EndIf
	Else
		lRet := .T.
	Endif

Return lRet

/*/{Protheus.doc} VILOGPROC
Rotina para leitura e processamento dos registros
@type  Static Function
@param none
@return none
/*/
Static Function VILOGPROC(xLin,xPos,xArray,cArqLog,cMsg,cMsg2,xCpo,xCategor,oFileLog)

	//Processa a gravação do Log
	If Valtype(oFileLog) == 'O'
		If oFileLog:Exists()
			If !Empty(xArray) .And. !Empty(cMsg)
				If "EM BRANCO" $ cMsg
					oFileLog:Write("["+cMsg+"] Registro: "+xArray[xPos]+" [Linha:"+cValToChar(xLin)+"] - NÃO ENCONTRADO ID - ["+xCpo+"]"+CRLF)
				Else
					oFileLog:Write("["+cMsg+"] Registro: "+xArray[xPos]+" [Linha:"+Iif(ValType(xLin)='N',cValToChar(xLin),xLin)+"] - "+cMsg2+CRLF)
				Endif
			ElseIf !Empty(cMsg)
				oFileLog:Write(cMsg)
			Endif
		Else
			If (oFileLog:Create())
				If !Empty(xArray) .And. !Empty(cMsg)
					If "EM BRANCO" $ cMsg
						oFileLog:Write("["+cMsg+"] Registro: "+xArray[xPos]+" [Linha:"+cValToChar(xLin)+"] - NÃO ENCONTRADO ID"+CRLF)
					Else
						oFileLog:Write("["+cMsg+"] Registro: "+xArray[xPos]+" [Linha:"+Iif(ValType(xLin)='N',cValToChar(xLin),xLin)+"] - "+cMsg2+CRLF)
					Endif
				ElseIf !Empty(cMsg)
					oFileLog:Write(cMsg)
				Endif
			Else
				MsgStop("Houve um erro ao gerar o arquivo: " + CRLF + oFileLog:Error():Message, "Atenção")
			Endif
		EndIf
	Else
		oFileLog  := FWFILEWRITER():New(cArqLog,.T.)
		If (oFileLog:Create())
			If !Empty(xArray) .And. !Empty(cMsg)
				If "EM BRANCO" $ cMsg
					oFileLog:Write("["+cMsg+"] Registro: "+xArray[xPos]+" [Linha:"+cValToChar(xLin)+"] - NÃO ENCONTRADO ID"+CRLF)
				Else
					oFileLog:Write("["+cMsg+"] Registro: "+xArray[xPos]+" [Linha:"+Iif(ValType(xLin)='N',cValToChar(xLin),xLin)+"] - "+cMsg2+CRLF)
				Endif
			ElseIf !Empty(cMsg)
				oFileLog:Write(cMsg)
			Endif
		Else
			MsgAlert("Não foi possível criar o arquivo de Log","Atenção")
		Endif
	Endif
	cMsg  := ''
	cMsg2 := ''

Return

/*/{Protheus.doc} VIENVMAIL
Função para envio do arquivo de log por email
@type  Static Function
@author user
@since 16/07/2020
@version version    
/*/
Static Function VIENVMAIL(cPath)
	//Declarar variaveis locais
	Local cServer    := GetMV("MV_RElseRV")     // Nome do servidor de e-mail
	Local cConta     := GetMV("MV_RELACNT")     // Nome da conta a ser usada no e-mail
	Local cPaswd     := GetMV("MV_RELPSW")      // Senha
	Local lAutentica := GetMv("MV_RELAUTH",,.F.)// Determina se o Servidor de Email necessita de Autenticação
	Local cCCEXP     := GetMv("VI_CCOCUL",, "")// Determina os e-mails que receberão cópia oculta
	Local cAnexo     := cPath
	Local cAssunto 	 := "Log de processamento de baixa - Rotina VIFIMP01"
	Local cEmails    := SuperGetMV("VI_LOGMAIL" , .T., 'isabel.cardoso@victorinox.com;carllos.hfernandes@gmail.com')
	Local cMensagem  := ""
	Local cBody      := ""

	//Monta HTML para o corpo do email
	cBody := '<html>'
	cBody := '<head><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>'
	cBody += '<body>'
	cBody += '<style type="text/css">'
	cBody += '.table  {border-bottom:1px solid #999; border-right:1px solid #999; border-left:1px solid #999; border-top:1px solid #999; margin:1em auto}'
	cBody += '.top_bg {background-color:#F0FFF0; center repeat-x; margin:0; padding:0; font-family: Verdana,Arial; font-size: 17px}'
	cBody += '</style>'
	cBody += '<table width="100%" class="table">'
	cBody += '<tr><td align="left" height="30px"><b>  Prezado(a) </b></td></tr>'
	cBody += '<tr><td height="100px"></td></tr>'
	cBody += '<tr><td align="left" height="30px">Informo que foi executada a rotina de conciliação financeira em '+DtoC(Date())+' às '+Time()+'hrs.</td></tr>'
	cBody += '<tr><td align="left" height="30px">Segue anexo o arquivo de Log de processamento da rotina para analise.</td></tr>'
	cBody += '<tr><td height="100px"></td></tr>'
	cBody += '<tr><td align="left" height="30px">Obrigado.</td></tr>'
	cBody += '<tr><td height="100px"></td></tr>'
	cBody += '<tr><td align="left" height="30px">( Mensagem automatica - Favor nao responder )</td></tr>'
	cBody += '<tr><td align="center" ><em>*** E-mail enviado por ERP Totvs Protheus ***</em></td></tr>'
	cBody += '</table>'
	cBody += '</body>'
	cBody += '</html>'

	cFrom :=cConta
	cMensagem := cBody
	If !Empty(cServer) .And. !Empty(cConta)
		CONNECT SMTP  SERVER   cServer  ACCOUNT  cConta  PASSWORD  cPaswd  RESULT  lOk

		If lOk
			If lAutentica
				lOk := Mailauth(cConta, cPaswd)
			Else
				lOk := .T.
			Endif
			IF lOk
				IF EMPTY(cAnexo)
					IF EMPTY(cCCEXP)
						SEND MAIL FROM  cFrom TO cEmails  SUBJECT cAssunto  BODY  cMensagem  RESULT lOk
					Else
						SEND MAIL FROM  cFrom TO cEmails BCC cCCEXP SUBJECT cAssunto  BODY  cMensagem  RESULT lOk
					ENDIF
				Else
					SEND MAIL FROM  cFrom TO cEmails  SUBJECT cAssunto  BODY  cMensagem ATTACHMENT cAnexo  RESULT lOk
				ENDIF
				If !lOk
					GET MAIL ERROR cError
					Aviso("Erro no envio do e-Mail",cError,{"Fechar"},2)
				Else
					Conout("*** Envio de e-mail: " + cAssunto + " para: " + cEmails + " Filial: " + cFilAnt + " ***")
				EndIf
			Else
				GET MAIL ERROR cError
				Aviso("Erro na autenticação do servidor SMTP",cError,{"Fechar"},2)
			EndIf
		Else
			GET MAIL ERROR cError
			Alert("Erro na conexão do servidor  SMTP.")

		Endif
		DISCONNECT SMTP SERVER
	Else
		Alert("Paramêtros para envio de e-mail não configurados (MV_RElseRV,MV_RELACNT).")
	EndIf

Return


/*/{Protheus.doc} NoAcento
Função para retirada de acento da string
@type  Static Function
@author NFESEFAZ
@since 16/07/2020
@version version    
/*/
Static Function NoAcento(cString)
	Local cChar  := ""
	Local nX     := 0
	Local nY     := 0
	Local cVogal := "aeiouAEIOU"
	Local cAgudo := "áéíóú"+"ÁÉÍÓÚ"
	Local cCircu := "âêîôû"+"ÂÊÎÔÛ"
	Local cTrema := "äëïöü"+"ÄËÏÖÜ"
	Local cCrase := "àèìòù"+"ÀÈÌÒÙ"
	Local cTio   := "ãõÃÕ"
	Local cCecid := "çÇ"
	Local cMaior := "&lt;"
	Local cMenor := "&gt;"
	Local cAspasD:= '"'
	Local cAspasS:= "'"
	Local cPonto := "."
	Local cVirg  := ","
	Local cIfen  := "-"
	Local cBarrIn:= "/"
	Local c2Ponto:= ":"
	Local cNumO  := "º"
	Local cNumA  := "ª"
	Local cGrau  := "°"

	cString := StrTran( cString, "ANTECIPA€ÇO", "ANTECIPACAO" )
	cString := StrTran( cString, "LÖQUIDO", "LÍQUIDO" )
	cString := StrTran( cString, "éLTIMO" , "ÚLTIMO" )
	cString := StrTran( cString, "Cr‚dito", "Crédito" )
	cString := StrTran( cString, "§", "" )
	For nX := 1 To Len(cString)
		cChar:=SubStr(cString, nX, 1)
		IF cChar$cAgudo+cCircu+cTrema+cCecid+cTio+cCrase
			nY := At(cChar,cAgudo)
			If nY > 0
				cString := StrTran(cString,cChar,SubStr(cVogal,nY,1))
			EndIf
			nY := At(cChar,cCircu)
			If nY > 0
				cString := StrTran(cString,cChar,SubStr(cVogal,nY,1))
			EndIf
			nY := At(cChar,cTrema)
			If nY > 0
				cString := StrTran(cString,cChar,SubStr(cVogal,nY,1))
			EndIf
			nY := At(cChar,cCrase)
			If nY > 0
				cString := StrTran(cString,cChar,SubStr(cVogal,nY,1))
			EndIf
			nY := At(cChar,cTio)
			If nY > 0
				cString := StrTran(cString,cChar,SubStr("aoAO",nY,1))
			EndIf
			nY := At(cChar,cCecid)
			If nY > 0
				cString := StrTran(cString,cChar,SubStr("cC",nY,1))
			EndIf
		Endif
	Next

	If cMaior$ cString
		cString := strTran( cString, cMaior, "" )
	EndIf
	If cMenor$ cString
		cString := strTran( cString, cMenor, "" )
	EndIf

	cString := StrTran( cString, CRLF, " " )
	cString := StrTran( cString, cAspasD, "" )
	cString := StrTran( cString, cAspasS, "" )
	cString := StrTran( cString, cPonto, "" )
	cString := StrTran( cString, cVirg, "" )
	cString := StrTran( cString, cIfen, "" )
	cString := StrTran( cString, cBarrIn, "" )
	cString := StrTran( cString, c2Ponto, "" )
	cString := StrTran( cString, cNumO, "" )
	cString := StrTran( cString, cNumA, "" )
	cString := StrTran( cString, cGrau, "" )
	cString := StrTran( cString, " ", "" )

	For nX:=1 To Len(cString)
		cChar:=SubStr(cString, nX, 1)
		If (Asc(cChar) < 32 .Or. Asc(cChar) > 123) .and. !cChar $ '|'
			cString:=StrTran(cString,cChar,".")
		Endif
	Next nX

Return cString
