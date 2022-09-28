#Include "Protheus.ch"
#Include "TopConn.ch"
#Include "RPTDef.ch"
#Include "FWPrintSetup.ch"

Static _cRepDb	:= GetSrvProfString("RepositInDataBase","")
Static _cRep	:= SuperGetMv("MV_REPOSIT",.F.,"1")
Static _lRepDb	:= ( _cRepDb == "1" .And. _cRep == "2" )

//Alinhamentos
#Define PAD_LEFT            0
#Define PAD_RIGHT           1
#Define PAD_CENTER          2
#Define ENTER          		chr(13)+chr(10) 

//Cores
#Define COR_CINZA   RGB(180, 180, 180)
#Define COR_PRETO   RGB(000, 000, 000)
 
//Colunas
#Define COL_LOGO        0020
#Define COL_N3_TIPO     0150 
#Define COL_N1_FILIAL   0175
#Define COL_NG_DESCRIC  0250
#Define COL_N1_CBASE    0350
#Define COL_N1_ITEM     0400
#Define COL_N1_DESCRIC  0450
#Define COL_N3_CCUSTO   0600
#Define COL_N1_CHAPA    0650
#Define COL_N1_NFISCAL  0700
#Define COL_N1_AQUISIC  0450
#Define COL_N3_TXDEPR1  0600
#Define COL_N3_VORIG1   0650
#Define COL_N3_RESIDUAL 0700
#Define COL_N3_VRDACM1  0750
/*
‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
±±∫Programa  ≥IBATF002∫  Autor ≥Denys Dias           ∫ Data ≥  31/05/2021 ∫±±
±±ÃÕÕÕÕÕÕÕÕÕÕÿÕÕÕÕÕÕÕÕÕÕ ÕÕÕÕÕÕÕœÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕ ÕÕÕÕÕÕœÕÕÕÕÕÕÕÕÕÕÕÕÕπ±±
±±∫Desc.     ≥ Imprime relatÛrio de ativos em spool ou Excel              ∫±±
±±ÃÕÕÕÕÕÕÕÕÕÕÿÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕπ±±
±±∫Uso       ≥ IBRATEC GR¡FICA                                            ∫±±
ﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂ
*/
 
User Function IBATF002()
Local aArea := GetArea()
Local cCadastro 	:= OemToAnsi("RELAT”RIO DE ATIVOS")
Local aSays			:={}
Local aButtons		:={}
Local nOpca 		:= 0

//Linhas e colunas
Private cPerg     := "IATF001"
Private nLinAtu   := 000
Private nLinLogo  := 000
Private nTamLin   := 028
Private nLinFin   := 550
Private nColIni   := 010
Private nColFin   := 820
Private nColMeio  := (nColFin-nColIni)/2
//Objeto de Impress„o
Private oPrintPvt
Private oBrushCinza
//armazena nome  foto
Private aFotos		:= {}			 
//Vari·veis auxiliares
Private dDataGer  := Date()
Private cHoraGer  := Time()
Private nPagAtu   := 1
Private cNomeUsr  := UsrRetName(RetCodUsr())
//Fontes
Private cNomeFont := "Arial"
Private oFontDet  := TFont():New(cNomeFont, 9, -10, .T., .F., 5, .T., 5, .T., .F.)
Private oFontDetN := TFont():New(cNomeFont, 9, -10, .T., .T., 5, .T., 5, .T., .F.)
Private oFontRod  := TFont():New(cNomeFont, 9, -08, .T., .F., 5, .T., 5, .T., .F.)
Private oFontTit  := TFont():New(cNomeFont, 9, -13, .T., .T., 5, .T., 5, .T., .F.) 

PergAtf01()
Pergunte(cPerg,.T.)

AADD(aSays,OemToAnsi( "  Este programa ira imprimir os dados referente aos "))
AADD(aSays,OemToAnsi( "  dados do Ativo Fixo, conforme parametros do usuario. "))

AADD(aButtons, { 1,.T.,{|| nOpca := 1,FechaBatch() }} )
AADD(aButtons, { 2,.T.,{|| nOpca := 0,FechaBatch() }} )

FormBatch( cCadastro, aSays, aButtons )

If nOpca == 1
	If Mv_Par09 == 1 
		//Processa( { |lEnd| 	Impress(oPrintPvt) })
        Processa({|lEnd| fMontaRel(oPrintPvt)}, "Processando...")
	Else 
		//Processa( { |lEnd| 	ImprExc(oPrintPvt) })
        Processa({|lEnd| fRelExc(oPrintPvt)}, "Processando...")
	EndIf
EndIf

//Se a pergunta for confirmada
//If MsgYesNo("Deseja gerar o relatÛrio de grupos de produtos?", "AtenÁ„o")
//    Processa({|| fMontaRel()}, "Processando...")
//EndIf
     
RestArea(aArea)

Return
 
/*---------------------------------------------------------------------*
 | Func:  fMontaRel                                                    |
 | Desc:  FunÁ„o que monta o relatÛrio spool                           |
 *---------------------------------------------------------------------*/
 
Static Function fMontaRel()
Local cNamefil    := ""
Local cCaminho    := ""
Local cArquivo    := ""
Local cQuery      := ""
Local nAtual      := 0
Local nTotal      := 0
Local cLogo		  := "C:\temp\TESTE123.bmp"
Local nTVlrBem	:= nTVlrRes	:= nTVlrDepA:= 0  
Local cChvIGM	:= ""
Local cLogoAux	:= ""  
Local aFiles 	:= {}
Local nX, nPos	:= 0
Local aImgsAtf 	:= {}	
//Definindo o diretÛrio como a tempor·ria do S.O. e o nome do arquivo com a data e hora (sem dois pontos)
cCaminho  := GetTempPath()
cArquivo  := "RATF002" + dToS(dDataGer) + "_" + StrTran(cHoraGer, ':', '-')
    
//Criando o objeto do FMSPrinter
oPrintPvt := FWMSPrinter():New(cArquivo, IMP_PDF, .F., "", .T., , @oPrintPvt, "", , , , .T.)
    
//Setando os atributos necess·rios do relatÛrio
oPrintPvt:SetResolution(72)
oPrintPvt:SetLandscape()
oPrintPvt:SetPaperSize(DMPAPER_A4)
oPrintPvt:SetMargin(60, 60, 60, 60)

If !(ExistDir("C:\temp\imgs" ))
	If (Makedir("C:\temp\imgs")) <> 0
		MSGSTOP("N„o foi possÌvel criar o diretÛrio C:\temp\imgs, para salvar as imagens tempor·rias!"+chr(13)+chr(10)+;
				"PROGRAMA SER¡ FINALIZADO!!!","ATEN«√O")
		Return()
	Endif
EndIf

aFiles := Directory("C:\temp\imgs\*.*", "D")
For nX := 1 to Len( aFiles )
	IF File("C:\temp\imgs\"+aFiles[nX,1])
		fErase("C:\temp\imgs\"+aFiles[nX,1])
	Endif
Next nX
    
//Imprime o cabeÁalho
fImpCab()
    
//Montando a consulta
cQuery :=""
cQuery += " SELECT N3_TIPO, N1_FILIAL, N1_GRUPO, NG_DESCRIC,N1_CBASE, N1_ITEM,N1_DESCRIC, N3_CCUSTO, N1_CHAPA, N3_TXDEPR1, N1_NFISCAL, N1_AQUISIC, "+ENTER
cQuery += " N3_AQUISIC, N3_VORIG1, N3_VORIG1-N3_VRDACM1 AS N3_RESIDUAL, N3_VRDACM1, N1_BITMAP "+ENTER
cQuery += " FROM "+RetSqlName("SN1")+" SN1  "+ENTER
cQuery += " INNER JOIN "+RetSqlName("SN3")+" SN3 ON N1_FILIAL=N3_FILIAL "+ENTER
cQuery += " AND N1_CBASE=N3_CBASE "+ENTER
cQuery += " AND N1_ITEM=N3_ITEM "+ENTER
cQuery += " AND N3_CCUSTO BETWEEN  '"+MV_PAR12+"' AND '"+MV_PAR13+"'"+ENTER
cQuery += " AND N3_BAIXA <> '1' "+ENTER
If AllTrim(MV_PAR17) != ""
	cQuery += " AND N3_TIPO BETWEEN '"+AllTrim(MV_PAR17)+"' AND '"+AllTrim(MV_PAR17)+"'  "+ENTER
ElseIf AllTrim(MV_PAR17) == ""
	cQuery += " AND N3_TIPO BETWEEN '  ' AND 'ZZ'  "+ENTER
EndIf
cQuery += " AND SN3.D_E_L_E_T_='' "+ENTER
cQuery += " LEFT JOIN "+RetSqlName("SNG")+" SNG ON N1_FILIAL=NG_FILIAL "+ENTER
cQuery += " AND N1_GRUPO=NG_GRUPO "+ENTER
cQuery += " AND SNG.D_E_L_E_T_='' "+ENTER
cQuery += " WHERE SN1.D_E_L_E_T_='' "+ENTER
cQuery += " AND N1_FILIAL BETWEEN '"+MV_PAR01+"' AND '"+MV_PAR02+"'"+ENTER
cQuery += " AND N1_ITEM BETWEEN  '"+MV_PAR03+"' AND '"+MV_PAR04+"'"+ENTER
cQuery += " AND N1_CBASE BETWEEN  '"+MV_PAR05+"' AND '"+MV_PAR06+"'"+ENTER
cQuery += " AND N1_AQUISIC BETWEEN '"+DTOS(MV_PAR07)+"' AND '"+DTOS(MV_PAR08)+"' "+ENTER
cQuery += " AND N1_GRUPO BETWEEN  '"+MV_PAR10+"' AND '"+MV_PAR11+"'"+ENTER
cQuery += " AND N1_CHAPA BETWEEN  '"+MV_PAR14+"' AND '"+MV_PAR15+"'"+ENTER
If MV_PAR16 == 1
	cQuery += " AND N1_BITMAP <> '' "
ElseIf MV_PAR16 == 2
	cQuery += " AND N1_BITMAP = '' "
EndIf

cQuery += " ORDER BY N1_FILIAL, N1_CBASE, N1_ITEM  "+ENTER
// N3_CUSTBEM, N3_CDEPREC, N3_CCUSTO, N3_CCDEPR, N3_CDESP
cQuery  := ChangeQuery(cQuery)

If Select("TATF")> 0
	TATF->(dbCloseArea())
EndIf

DbUseArea(.T.,"TopConn",TcGenQry(,,cQuery),"TATF",.T.,.T.)                       

//Conta o total de registros, seta o tamanho da rÈgua, e volta pro topo
Count To nTotal
ProcRegua(nTotal)
TATF->(DBGoTop())
nAtual := 0

While TATF->(!Eof()) 
    nAtual++
    IncProc("Imprimindo Bem " + TATF->N1_CBASE + " (" + cValToChar(nAtual) + " de " + cValToChar(nTotal) + ")...")
        
    //Se a linha atual mais o espaÁo que ser· utilizado forem maior que a linha final, imprime rodapÈ e cabeÁalho
    If nLinAtu + nTamLin > nLinFin
        fImpRod()
        fImpCab()
    EndIf
	//N3_TIPO, N1_FILIAL, NG_DESCRIC,N1_CBASE, N1_ITEM,N1_DESCRIC, N3_CCUSTO, N1_CHAPA, N1_NFISCAL, 
	//N1_AQUISIC, N3_TXDEPR1, N3_VORIG1, N3_RESIDUAL, N3_VRDACM1        
    //Imprimindo a linha atual
    
	nLinLogo:= nLinAtu
	nPos:= aScan(aImgsAtf, 	{|x| x[1] = Alltrim(TATF->N1_BITMAP)})
			//aScan( aGer,	{|x| x[1] = QRY->A3_GEREN})
	If (cChvIGM== "" .AND. Alltrim(TATF->N1_BITMAP)!= "") .OR. (nPos = 0 .AND. Alltrim(TATF->N1_BITMAP)!= "")  //Alltrim(cChvIGM) != Alltrim(TATF->N1_BITMAP)
    	cLogo:= IMGATF(oPrintPvt)
		cChvIGM:= Alltrim(TATF->N1_BITMAP)
		cLogoAux:=cLogo
		aAdd(aImgsAtf,{Alltrim(TATF->N1_BITMAP)})
	Else 
		If Alltrim(TATF->N1_BITMAP)!= ""
			cLogo:="C:\temp\imgs\"+aImgsAtf[nPos,1]+".jpg"
		Else
			cLogo:=""
		EndIf
	EndIf
	//cLogo:= Iif(MV_PAR16 == 1, cLogo, "")
    oPrintPvt:SayBitmap(nLinLogo+2,COL_LOGO,cLogo,100,050)
    oPrintPvt:SayAlign(nLinAtu, COL_N3_TIPO     , TATF->N3_TIPO     ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
    DBSELECTAREA("SM0")          
	DBSETORDER(1)
	If SM0->(DbSeek(cEmpAnt+TATF->N1_FILIAL))
		cNamefil:= SM0->M0_FILIAL
	EndIf 
    oPrintPvt:SayAlign(nLinAtu, COL_N1_FILIAL   , cNamefil   ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
    oPrintPvt:SayAlign(nLinAtu, COL_NG_DESCRIC  , TATF->NG_DESCRIC  ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
    oPrintPvt:SayAlign(nLinAtu, COL_N1_CBASE    , TATF->N1_CBASE    ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
    oPrintPvt:SayAlign(nLinAtu, COL_N1_ITEM     , TATF->N1_ITEM     ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
    oPrintPvt:SayAlign(nLinAtu, COL_N1_DESCRIC  , TATF->N1_DESCRIC  ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
    oPrintPvt:SayAlign(nLinAtu, COL_N3_CCUSTO   , TATF->N3_CCUSTO   ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
    oPrintPvt:SayAlign(nLinAtu, COL_N1_CHAPA    , TATF->N1_CHAPA    ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
    oPrintPvt:SayAlign(nLinAtu, COL_N1_NFISCAL  , TATF->N1_NFISCAL  ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
    nLinAtu+= nTamLin
    oPrintPvt:SayAlign(nLinAtu, COL_N1_AQUISIC  , "Data: "+DTOC(STOD(TATF->N1_AQUISIC))  ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
    oPrintPvt:SayAlign(nLinAtu, COL_N3_TXDEPR1  , "Tx.: "+Transform(TATF->N3_TXDEPR1 ,"@E 99")+"%"  ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
    oPrintPvt:SayAlign(nLinAtu, COL_N3_VORIG1   , Alltrim(Transform(TATF->N3_VORIG1	,"@E 99,999,999.99"))  ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
    oPrintPvt:SayAlign(nLinAtu, COL_N3_RESIDUAL , Alltrim(Transform(TATF->N3_RESIDUAL,"@E 99,999,999.99")) ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
    oPrintPvt:SayAlign(nLinAtu, COL_N3_VRDACM1  , Alltrim(Transform(TATF->N3_VRDACM1 ,"@E 99,999,999.99")) ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
    nLinAtu += nTamLin
	oPrintPvt:Line(nLinAtu, nColIni, nLinAtu, nColFin, COR_PRETO)
	nTVlrBem	+= TATF->N3_VORIG1
	nTVlrRes	+= TATF->N3_RESIDUAL
	nTVlrDepA	+= TATF->N3_VRDACM1        
    TATF->(dbskip())
EndDo
TATF->(DbCloseArea())
    
//oBrushCinza := TBrush():New(,Rgb(214,214,214))
//oPrint:FillRect({_nLin,0100,_nLin+50,3000}, oBrushCinza)
//oPrint:Box (_nLin,0100,_nLin+50,3000)
oPrintPvt:SayAlign(nLinAtu, COL_N1_AQUISIC  , "Total Geral" ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinAtu, COL_N3_VORIG1   , Alltrim(Transform(nTVlrBem	,"@E 99,999,999.99")) ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinAtu, COL_N3_RESIDUAL , Alltrim(Transform(nTVlrRes    ,"@E 99,999,999.99")) ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinAtu, COL_N3_VRDACM1  , Alltrim(Transform(nTVlrDepA   ,"@E 99,999,999.99")) ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)

//Se ainda tiver linhas sobrando na p·gina, imprime o rodapÈ final
If nLinAtu <= nLinFin
    fImpRod()
EndIf

//Mostrando o relatÛrio
oPrintPvt:Preview()

aFiles := Directory("C:\temp\imgs\*.*", "D")
For nX := 1 to Len( aFiles )
	IF File("C:\temp\imgs\"+aFiles[nX,1])
		fErase("C:\temp\imgs\"+aFiles[nX,1])
	Endif
Next nX

Return
 
/*---------------------------------------------------------------------*
 | Func:  fImpCab                                                      |
 | Desc:  FunÁ„o que imprime o cabeÁalho                               |
 *---------------------------------------------------------------------*/
 
Static Function fImpCab()
Local cTexto   := ""
Local nLinCab  := 030

//Iniciando P·gina
oPrintPvt:StartPage()
    
//CabeÁalho
cTexto := UPPER("RelatÛrio de Ativos - "+MesExtenso(Month(date()))+" de "+ALLTRIM(STR(YEAR(date()))))
oPrintPvt:SayAlign(nLinCab, nColMeio - 120, cTexto, oFontTit, 240, 20, COR_CINZA, PAD_CENTER, 0)
    
//Linha SeparatÛria
nLinCab += nTamLin + 10 
oPrintPvt:Line(nLinCab, nColIni, nLinCab, nColFin, COR_CINZA)
    
//CabeÁalho das colunas
//nLinCab += nTamLin
//oPrintPvt:SayAlign(nLinCab, COL_GRUPO, "Grupo",     oFontDetN, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
//oPrintPvt:SayAlign(nLinCab, COL_DESCR, "DescriÁ„o", oFontDetN, 0200, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinCab, COL_N3_TIPO     , "Tipo"        ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinCab, COL_N1_FILIAL   , "Filial"      ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinCab, COL_NG_DESCRIC  , "Grupo"       ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinCab, COL_N1_CBASE    , "CÛd. Ativo"  ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinCab, COL_N1_ITEM     , "Item"        ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinCab, COL_N1_DESCRIC  , "DescriÁ„o"   ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinCab, COL_N3_CCUSTO   , "C. Custo"    ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinCab, COL_N1_CHAPA    , "Placa"       ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinCab, COL_N1_NFISCAL  , "Documento"   ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
nLinCab += 14
oPrintPvt:SayAlign(nLinCab, COL_N1_AQUISIC  , "AquisÁ„o"    ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinCab, COL_N3_TXDEPR1  , "Taxa (%): "  ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinCab, COL_N3_VORIG1   , "Vlr do Bem"	,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinCab, COL_N3_RESIDUAL , "Vlr Residual",oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
oPrintPvt:SayAlign(nLinCab, COL_N3_VRDACM1  , "Dep. Acumul." ,oFontDet, 0080, nTamLin, COR_PRETO, PAD_LEFT, 0)
nLinCab += 14
oPrintPvt:Line(nLinCab, nColIni, nLinCab, nColFin, COR_CINZA)    
//Atualizando a linha inicial do relatÛrio
nLinAtu := nLinCab + 10
Return()
 
/*---------------------------------------------------------------------*
 | Func:  fImpRod                                                      |
 | Desc:  FunÁ„o que imprime o rodapÈ                                  |
 *---------------------------------------------------------------------*/
 
Static Function fImpRod()
Local nLinRod   := nLinFin + nTamLin
Local cTextoEsq := ''
Local cTextoDir := ''

//Linha SeparatÛria
oPrintPvt:Line(nLinRod, nColIni, nLinRod, nColFin, COR_CINZA)
nLinRod += 3
    
//Dados da Esquerda e Direita
cTextoEsq := dToC(dDataGer) + "    " + cHoraGer + "    " + "IBATF002" + "    " + cNomeUsr
cTextoDir := "P·gina " + cValToChar(nPagAtu)
    
//Imprimindo os textos
oPrintPvt:SayAlign(nLinRod, nColIni,    cTextoEsq, oFontRod, 200, 05, COR_CINZA, PAD_LEFT,  0)
oPrintPvt:SayAlign(nLinRod, nColFin-40, cTextoDir, oFontRod, 040, 05, COR_CINZA, PAD_RIGHT, 0)
    
//Finalizando a p·gina e somando mais um
oPrintPvt:EndPage()
nPagAtu++

Return()
/*
‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
±±∫Programa  ≥fRelExc∫    Autor ≥Denys Dias           ∫ Data ≥ 31/05/2021 ∫±±
±±ÃÕÕÕÕÕÕÕÕÕÕÿÕÕÕÕÕÕÕÕÕÕ ÕÕÕÕÕÕÕœÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕ ÕÕÕÕÕÕœÕÕÕÕÕÕÕÕÕÕÕÕÕπ±±
±±∫Desc.     ≥ Imprime relatÛrio de ativos em Excel                       ∫±±
±±ÃÕÕÕÕÕÕÕÕÕÕÿÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕπ±±
±±∫Uso       ≥ IBRATEC GR¡FICA                                            ∫±±
ﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂ
*/

STATIC Function fRelExc(oPrint)
Local cQueExc 	  :=""
Local cQueGrp 	  :=""
Local aArea       := GetArea()
Local aAreaX3     := SX3->(GetArea())
Local oFWMsExcel
Local cDiretorio  := GetTempPath()
Local cArquivo    := 'Qry2021Exc.xml'
Local cArqFull    := cDiretorio + cArquivo
Local cPlaSint 	  := "Ativo Fixo" 
Local cTitulo	  := "RelatÛrio de Ativo Fixo"
Local cTable      := ""+UPPER(cTitulo+"-"+MesExtenso(Month(date()))+" de "+ALLTRIM(STR(YEAR(date()))))
Local cNomefil	  := ""
Local nTVlrBemE	:= nTVlrResE	:= nTVlrDepAE := 0         

DBSELECTAREA("SM0")          
DBSETORDER(1)
SM0->(DbSeek(cEmpAnt+cFilAnt))     

//Criando o objeto que ir· gerar o conte˙do do Excel
oFWMsExcel:=FWMSExcel():New()       

cQueGrp :=""
cQueGrp += " SELECT DISTINCT NG_FILIAL, NG_DESCRIC, NG_GRUPO "+ENTER
cQueGrp += " FROM "+RetSqlName("SN1")+" SN1  "+ENTER
cQueGrp += " INNER JOIN "+RetSqlName("SN3")+" SN3 ON N1_FILIAL=N3_FILIAL "+ENTER
cQueGrp += " AND N1_CBASE=N3_CBASE "+ENTER
cQueGrp += " AND N1_ITEM=N3_ITEM "+ENTER
cQueGrp += " AND N3_CCUSTO BETWEEN  '"+MV_PAR12+"' AND '"+MV_PAR13+"'"+ENTER
cQueGrp += " AND SN3.D_E_L_E_T_='' "+ENTER
cQueGrp += " LEFT JOIN "+RetSqlName("SNG")+" SNG ON N1_FILIAL=NG_FILIAL "+ENTER
cQueGrp += " AND N3_BAIXA <> '1' "+ENTER
cQueGrp += " AND N1_GRUPO=NG_GRUPO "+ENTER
cQueGrp += " AND SNG.D_E_L_E_T_='' "+ENTER
cQueGrp += " WHERE SN1.D_E_L_E_T_='' "+ENTER
cQueGrp += " AND N1_FILIAL BETWEEN '"+MV_PAR01+"' AND '"+MV_PAR02+"'"+ENTER
cQueGrp += " AND N1_ITEM BETWEEN  '"+MV_PAR03+"' AND '"+MV_PAR04+"'"+ENTER
cQueGrp += " AND N1_CBASE BETWEEN  '"+MV_PAR05+"' AND '"+MV_PAR06+"'"+ENTER
cQueGrp += " AND N1_AQUISIC BETWEEN '"+DTOS(MV_PAR07)+"' AND '"+DTOS(MV_PAR08)+"' "+ENTER
cQueGrp += " AND N1_GRUPO BETWEEN  '"+MV_PAR10+"' AND '"+MV_PAR11+"'"+ENTER
cQueGrp += " AND N1_CHAPA BETWEEN  '"+MV_PAR14+"' AND '"+MV_PAR15+"'"+ENTER
If MV_PAR16 == 1
	cQueGrp += " AND N1_BITMAP <> '' "+ENTER
ElseIf MV_PAR16 == 2
	cQueGrp += " AND N1_BITMAP = '' "+ENTER
EndIf
cQueGrp += " ORDER BY NG_FILIAL, NG_DESCRIC, NG_GRUPO  "+ENTER

cQueGrp  := ChangeQuery(cQueGrp)

If Select("GRPATF")> 0
	GRPATF->(dbCloseArea())
EndIf
DbUseArea(.T.,"TopConn",TcGenQry(,,cQueGrp),"GRPATF",.T.,.T.)                       

DbSelectArea("GRPATF")
GRPATF->(DBGoTop())

While GRPATF->(!Eof()) 
	
	// Planilha SintÈtico 
	cPlaSint :=  Alltrim(GRPATF->NG_FILIAL+GRPATF->NG_DESCRIC)
	oFWMsExcel:AddworkSheet(cPlaSint)
	oFWMsExcel:AddTable(cPlaSint, cTable)
	//nAlign	NumÈrico	Alinhamento da coluna ( 1-Left,2-Center,3-Right )	
	//nFormat	NumÈrico	Codigo de formataÁ„o ( 1-General,2-Number,3-Monet·rio,4-DateTime )
	//N3_TIPO, N1_FILIAL, NG_DESCRIC,N1_CBASE, N1_ITEM,N1_DESCRIC, N3_CCUSTO, N1_CHAPA, N3_TXDEPR1, N1_NFISCAL,N3_AQUISIC, N3_VORIG1, N3_RESIDUAL, N3_VRDACM1
	oFWMsExcel:AddColumn(cPlaSint, cTable, "", 1, 1)       
	oFWMsExcel:AddColumn(cPlaSint, cTable, "", 1, 1)        
	oFWMsExcel:AddColumn(cPlaSint, cTable, "", 1, 1)       
	oFWMsExcel:AddColumn(cPlaSint, cTable, "", 1, 1)       
	oFWMsExcel:AddColumn(cPlaSint, cTable, "", 1, 1)       
	oFWMsExcel:AddColumn(cPlaSint, cTable, "", 1, 1)       
	oFWMsExcel:AddColumn(cPlaSint, cTable, "", 1, 1)       
	oFWMsExcel:AddColumn(cPlaSint, cTable, "", 1, 1)       
	oFWMsExcel:AddColumn(cPlaSint, cTable, "", 3, 2)       
	oFWMsExcel:AddColumn(cPlaSint, cTable, "", 1, 1)       
	oFWMsExcel:AddColumn(cPlaSint, cTable, "", 1, 4)       
	oFWMsExcel:AddColumn(cPlaSint, cTable, "", 3, 2)       
	oFWMsExcel:AddColumn(cPlaSint, cTable, "", 3, 2)       
	oFWMsExcel:AddColumn(cPlaSint, cTable, "", 3, 2)       

	// CabeÁalho SintÈtico
	oFWMsExcel:AddRow(cPlaSint, cTable, {"","","","","","","","","","","","","",""} )  
	oFWMsExcel:AddRow(cPlaSint, cTable, {"","","Filial: "+cEmpAnt+"-"+SM0->M0_FILIAL,"","","","","","","","","Emiss„o:"+Alltrim(Strzero(Month(date()),2))+"/"+Alltrim(Str(Year(date()))),"",""	})  
	oFWMsExcel:AddRow(cPlaSint, cTable, {"","","Hor·rio:"+Time()					,"","","","","","","","Periodo de ","AquisiÁ„o: "+DTOC(MV_PAR07)				," atÈ "+DTOC(MV_PAR08),""	})
	oFWMsExcel:AddRow(cPlaSint, cTable, {"","",""									,"","","","","","","","",""																			,"",""	}) 
	oFWMsExcel:AddRow(cPlaSint, cTable, {"Tipo","Filial","Grupo","CÛdigo Bem","Item","DescriÁ„o","C. Custo","Placa","Taxa","Documento","AquisiÁ„o","Valor","ResÌdual","DepreciaÁ„o Acum."		})

	//¿ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒŸ
	//≥ QUERY DE DADOS≥
	//¿ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒŸ
		
	cQueExc :=""
	cQueExc += " SELECT N3_TIPO, N1_FILIAL, N1_GRUPO, NG_DESCRIC,N1_CBASE, N1_ITEM,N1_DESCRIC, N3_CCUSTO, N1_CHAPA, N3_TXDEPR1, N1_NFISCAL, N1_AQUISIC, "+ENTER
	cQueExc += " N3_AQUISIC, N3_VORIG1, N3_VORIG1-N3_VRDACM1 AS N3_RESIDUAL, N3_VRDACM1, N1_BITMAP "+ENTER
	cQueExc += " FROM "+RetSqlName("SN1")+" SN1  "+ENTER
	cQueExc += " INNER JOIN "+RetSqlName("SN3")+" SN3 ON N1_FILIAL=N3_FILIAL "+ENTER
	cQueExc += " AND N1_CBASE=N3_CBASE "+ENTER
	cQueExc += " AND N1_ITEM=N3_ITEM "+ENTER
	cQueExc += " AND N3_CCUSTO BETWEEN  '"+MV_PAR12+"' AND '"+MV_PAR13+"'"+ENTER
	cQueExc += " AND N3_BAIXA <> '1' "+ENTER
	cQueExc += " AND SN3.D_E_L_E_T_='' "+ENTER
	cQueExc += " LEFT JOIN "+RetSqlName("SNG")+" SNG ON N1_FILIAL=NG_FILIAL "+ENTER
	cQueExc += " AND N1_GRUPO=NG_GRUPO "+ENTER
	cQueExc += " AND SNG.D_E_L_E_T_='' "+ENTER
	cQueExc += " WHERE SN1.D_E_L_E_T_='' "+ENTER
	cQueExc += " AND N1_FILIAL BETWEEN '"+GRPATF->NG_FILIAL+"' AND '"+GRPATF->NG_FILIAL+"'"+ENTER
	cQueExc += " AND N1_ITEM BETWEEN  '"+MV_PAR03+"' AND '"+MV_PAR04+"'"+ENTER
	cQueExc += " AND N1_CBASE BETWEEN  '"+MV_PAR05+"' AND '"+MV_PAR06+"'"+ENTER
	cQueExc += " AND N1_AQUISIC BETWEEN '"+DTOS(MV_PAR07)+"' AND '"+DTOS(MV_PAR08)+"' "+ENTER
	cQueExc += " AND N1_GRUPO BETWEEN  '"+GRPATF->NG_GRUPO+"' AND '"+GRPATF->NG_GRUPO+"'"+ENTER
	cQueExc += " AND N1_CHAPA BETWEEN  '"+MV_PAR14+"' AND '"+MV_PAR15+"'"+ENTER
	If MV_PAR16 == 1
		cQueExc += " AND N1_BITMAP <> '' "+ENTER
	ElseIf MV_PAR16 == 2
		cQueExc += " AND N1_BITMAP = '' "+ENTER
	EndIf
	cQueExc += " ORDER BY N1_FILIAL, N1_CBASE, N1_ITEM  "+ENTER

	cQueExc  := ChangeQuery(cQueExc)

	If Select("EATF")> 0
		EATF->(dbCloseArea())
	EndIf
	DbUseArea(.T.,"TopConn",TcGenQry(,,cQueExc),"EATF",.T.,.T.)                       

	DbSelectArea("EATF")
	EATF->(DBGoTop())

	While EATF->(!Eof())        
		DBSELECTAREA("SM0")          
		DBSETORDER(1)
		If SM0->(DbSeek(cEmpAnt+EATF->N1_FILIAL))
			cNomefil:= SM0->M0_FILIAL
		EndIf                      
		//N3_TIPO, N1_FILIAL, N1_GRUPO, NG_DESCRIC,N1_CBASE, N1_ITEM,N1_DESCRIC, N3_CCUSTO, N1_CHAPA, N3_TXDEPR1, N1_NFISCAL, N1_AQUISIC, "+ENTER
		//N3_AQUISIC, N3_VORIG1, N3_VORIG1-N3_VRDACM1 AS N3_RESIDUAL, N3_VRDACM1
		oFWMsExcel:AddRow(cPlaSint, cTable, {	EATF->N3_TIPO									,;
												cNomefil										,;
												EATF->NG_DESCRIC								,;
												EATF->N1_CBASE									,;
												EATF->N1_ITEM									,;
												EATF->N1_DESCRIC								,;
												EATF->N3_CCUSTO									,;
												EATF->N1_CHAPA									,;
												Transform(EATF->N3_TXDEPR1	,"@E 99")+"%"		,;
												EATF->N1_NFISCAL								,;
												DTOC(STOD(EATF->N1_AQUISIC))					,;
												EATF->N3_VORIG1									,;
												EATF->N3_RESIDUAL								,;
												EATF->N3_VRDACM1								})    
		nTVlrBemE	+= EATF->N3_VORIG1
		nTVlrResE	+= EATF->N3_RESIDUAL
		nTVlrDepAE	+= EATF->N3_VRDACM1
		
		EATF->(dbskip())
	Enddo
	// IMPRIME OS VALORES TOTAIS                                                                                          
	oFWMsExcel:AddRow(cPlaSint, cTable, {"","","","","","","","","","","","","",""})    
	oFWMsExcel:AddRow(cPlaSint, cTable, {"","","","","","","","","","","Total Geral:",nTVlrBemE,nTVlrResE,nTVlrDepAE})    

	GRPATF->(dbskip())
	nTVlrBemE	:= 0
	nTVlrResE	:= 0
	nTVlrDepAE	:= 0
Enddo

//Ativando o arquivo e gerando o xml
oFWMsExcel:Activate()
oFWMsExcel:GetXMLFile(cArqFull)
 
//Se tiver o excel instalado
If ApOleClient("msexcel")
	oExcel := MsExcel():New()
	oExcel:WorkBooks:Open(cArqFull)
	oExcel:SetVisible(.T.)
	oExcel:Destroy()
Else
	//Se existir a pasta do LibreOffice 5
	If ExistDir("C:\Program Files (x86)\LibreOffice 5")
		WaitRun('C:\Program Files (x86)\LibreOffice 5\program\scalc.exe "'+cDiretorio+cArquivo+'"', 1)
	//Sen„o, abre o XML pelo programa padr„o
	Else
		ShellExecute("open", cArquivo, "", cDiretorio, 1)
	EndIf
EndIf

EATF->(DbCloseArea()) 

RestArea(aAreaX3)
RestArea(aArea)

RETURN()

/*
‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
±±∫Programa  FotoATF     Autor ≥Denys Dias           ∫ Data ≥  31/05/2021 ∫±±
±±ÃÕÕÕÕÕÕÕÕÕÕÿÕÕÕÕÕÕÕÕÕÕ ÕÕÕÕÕÕÕœÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕ ÕÕÕÕÕÕœÕÕÕÕÕÕÕÕÕÕÕÕÕπ±±
±±∫Uso       ≥ IBRATEC GR¡FICA                                            ∫±±
ﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂ
*/
Static Function IMGATF(oPrintPvt)
//Local aArea		:= GetArea()
//Local cAlias	:= "PROTHEUS_REPOSIT"
Local cBmpPict	:= ""
Local cPath		:= "C:\temp\imgs\"  //GetSrvProfString("Startpath","")
Local cPathPict	:= ""
Local cRetImg	:= ""
Local lFile	
Local oDlg8
Local oBmp
Local cSAlias := Alias()
Local nSRecno := RecNo()
Local nSOrdem := IndexOrd()

//⁄ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒ
//≥ Carrega a Foto do ATIVO FIXO 								≥
//¿ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒ
cBmpPict := Upper( AllTrim(TATF->N1_BITMAP))
cPathPict 	:= (cPath + cBmpPict)
/*
⁄ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒø
≥ Para impressao da foto eh necessario abrir um dialogo para   ≥
≥ extracao da foto do repositorio.No entanto na impressao,nao  |
≥ ha a necessidade de visualiza-lo( o dialogo).Por esta razao  ≥
≥ ele sera montado nestas coordenadas fora da Tela             ≥
¿ƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒƒŸ
*/          
DEFINE MSDIALOG oDlg8   FROM -1000000,-4000000 TO -10000000,-8000000  PIXEL 
@ -10000000, -1000000000000 REPOSITORY oBmp SIZE -6000000000, -7000000000 OF oDlg8  

// Verifica se a imagem existe no repositorio
If !Empty(cBmpPict) .AND. oBMP:ExistBMP(cBmpPict) 
	If !_lRepDb
		oBmp:LoadBmp(cBmpPict)
	Else
		RepExtract(cBmpPict,cPathPict)
	EndIf
	//cPathPict:= "\system\TESTE123"
	//_nLin += 20		
	IF !Empty( cBmpPict := Upper( AllTrim(TATF->N1_BITMAP) ) )
		lFile:=Iif(_lRepDb, .T., oBmp:Extract(cBmpPict, cPathPict))
		If lFile 
			If File(cPathPict+".bmp") 

				aAdd(aFotos,cPathPict + ".bmp")
				cRetImg:= cPathPict + ".bmp" 
			ElseIf File(cPathPict+".jpg")
				aAdd(aFotos,cPathPict + ".jpg")
				cRetImg:= cPathPict + ".jpg"
			EndIf
		EndIf	
	EndIf	
EndIf
ACTIVATE MSDIALOG oDlg8 ON INIT (oBmp:lStretch := .T., oDlg8:End())

dbselectarea(cSAlias)
dbsetorder(nSOrdem)
dbgoto(nSRecno)
        
Return(cRetImg)
/*
‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹
±±∫Programa  PergAtf01   Autor ≥Denys Dias           ∫ Data ≥  31/05/2021 ∫±±
±±ÃÕÕÕÕÕÕÕÕÕÕÿÕÕÕÕÕÕÕÕÕÕ ÕÕÕÕÕÕÕœÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕÕ ÕÕÕÕÕÕœÕÕÕÕÕÕÕÕÕÕÕÕÕπ±±
±±∫Uso       ≥ IBRATEC GR¡FICA                                            ∫±±
ﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂﬂ
*/
Static Function PergAtf01()
Local sAlias := Alias()                          // Variaveis Auxiliares
Local aRegs  := {}
Local i, j

SX1->(DbSetOrder(1))                             // Perguntas do Sistema

Aadd(aRegs,{cPerg,"01","Filial de? 		","","","mv_cha","C",02,0,0,"G",""			,"Mv_Par01","","","","  "		,"","","","","","","","","","","","","","","","","","","","","XM0"	,"","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"02","Filial atÈ	?	","","","mv_chb","C",02,0,0,"G","NaoVazio()","Mv_Par02","","","","ZZ"		,"","","","","","","","","","","","","","","","","","","","","XM0"	,"","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"03","Item de?  		","","","mv_chc","C",04,0,0,"G",""			,"Mv_Par03","","","","  "		,"","","","","","","","","","","","","","","","","","","","",""		,"","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"04","Item atÈ? 		","","","mv_chd","C",04,0,0,"G","NaoVazio()","Mv_Par04","","","","ZZ"		,"","","","","","","","","","","","","","","","","","","","",""		,"","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"05","Do Bem ?   		","","","mv_che","C",10,0,0,"G",""			,"Mv_Par05","","","","  "		,"","","","","","","","","","","","","","","","","","","","",""		,"","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"06","AtÈ o Bem? 		","","","mv_chf","C",10,0,0,"G","NaoVazio()","Mv_Par06","","","","ZZ"		,"","","","","","","","","","","","","","","","","","","","",""		,"","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"07","AquisiÁ„o de?	","","","mv_chg","D",08,0,0,"G","NaoVazio()","Mv_Par07","","","","01/01/20"	,"","","","","","","","","","","","","","","","","","","","",""		,"","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"08","AquisiÁ„o AtÈ?	","","","mv_chh","D",08,0,0,"G","NaoVazio()","Mv_Par08","","","","31/01/20"	,"","","","","","","","","","","","","","","","","","","","",""		,"","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"09","Tipo RelatÛrio?	","","","mv_chi","C",01,0,0,"C",""			,"Mv_Par09","1-PDF","1-PDF","1-PDF","","","2-Excel","2-Excel","2-Excel","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"10","Grupo de?  		","","","mv_chj","C",04,0,0,"G",""			,"Mv_Par10","","","","  "		,"","","","","","","","","","","","","","","","","","","","",""		,"","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"11","Grupo atÈ? 		","","","mv_chk","C",04,0,0,"G","NaoVazio()","Mv_Par11","","","","ZZ"		,"","","","","","","","","","","","","","","","","","","","",""		,"","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"12","C. Custo de?	","","","mv_chl","C",09,0,0,"G",""			,"Mv_Par12","","","","  "		,"","","","","","","","","","","","","","","","","","","","","CTT"	,"","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"13","C. Custo atÈ?	","","","mv_chm","C",09,0,0,"G","NaoVazio()","Mv_Par13","","","","ZZ"		,"","","","","","","","","","","","","","","","","","","","","CTT"	,"","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"14","Plaqueta de?	","","","mv_chn","C",20,0,0,"G",""			,"Mv_Par14","","","","  "		,"","","","","","","","","","","","","","","","","","","","",""		,"","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"15","Plaqueta atÈ?	","","","mv_cho","C",20,0,0,"G","NaoVazio()","Mv_Par15","","","","ZZ"		,"","","","","","","","","","","","","","","","","","","","",""		,"","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"16","Ativo com Fotos?","","","mv_chp","C",01,0,0,"C",""			,"Mv_Par16","1-Sim","1-Sim","1-Sim","","","2-N„o","2-N„o","2-N„o","","","Ambos","Ambos","Ambos","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
Aadd(aRegs,{cPerg,"17","Tipo?           ","","","mv_chq","C",02,0,0,"G",""			,"Mv_Par17","","","",""         ,"","","","","","","","","","","","","","","","","","","","","G1"   ,"","","","","","","","","","","","","","",""})
DBSELECTAREA("SX1")
DBSETORDER(1)
cPerg:= cPerg+REPLICATE(" ",(LEN(SX1->X1_GRUPO)-LEN(cPerg)))

For i := 1 To Len(aRegs)                         // Gravar as Perguntas
    SX1->(DbSeek(cPerg + aRegs[i, 2]))
    If SX1->(!Found())
       DbSelectArea("SX1")
       If SX1->(Reclock("SX1", .T.))
          For j := 1 To FCount()
              If j <= Len(aRegs[i])
                 FieldPut(j, aRegs[i, j])
              Endif
          Next j
          SX1->(MsUnlock())
       Endif
    Endif
Next i

DbSelectArea(sAlias)

Return (.T.)
