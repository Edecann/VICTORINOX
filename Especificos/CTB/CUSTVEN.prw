#include "rwmake.ch"
#Include "protheus.ch"

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  � CUSTVEN   �Autor  �Edelcio Cano        � Data �   22/06/20 ���
�������������������������������������������������������������������������͹��
���Desc.     � Captura Centro de Custo do Cad. Vendedores - LP 678        ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � VICTORINOX                                                 ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
User Function CUSTVEN()

Local _aAreaAt	:=	GetArea()
Local _cCusto   := ""    
Local _cVend1   := ""

_cVend1 := Posicione("SF2",1,xFilial("SF2")+SD2->D2_DOC+SD2->D2_SERIE+SD2->D2_CLIENTE+SD2->D2_LOJA+SD2->D2_FORMUL+SD2->D2_TIPO,"F2_VEND1")

If _cVend1 <> ""

	dbSelectArea("SA3")
	_aAreaA3	:= GetArea()
	dbSetOrder(1)
	dbGotop()

	If dbSeek(xFilial("SA3")+_cVend1)
	
		_cCusto	:= SA3->A3_XCC

	Endif
	
	RestArea(_aAreaA3)
	
Endif

RestArea(_aAreaAt)

Return(_cCusto)