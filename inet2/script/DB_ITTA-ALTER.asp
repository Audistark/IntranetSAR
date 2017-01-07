<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!-- #include virtual = "/inet2/class/cAAA.asp" -->
<%
' Init AAA - Authentication, Authorization and Accounting
Dim oAAA : Set oAAA = new cAAA
Dim ret : ret = oAAA.WinAuthenticate(True)
If ret < 0 Then
	oAAA.Print()
End If

' S� MASTER
If oAAA.AuthorWinMaster() <> True Then
	Response.Status = "403 Forbidden"
	Response.End
End If

Dim querySQL, rsDiv
Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
If oDbFDH.ErrorNumber < 0 then
	oDbFDH.Print()
End If

Response.Status = "200 OK"

'''''''''''''''''''' Create Table ITTA ''''''''''''''''''''''''''''''''''
'

Set ret = oDbFDH.Execute( "DROP TABLE ITTA;" )
Response.Write "Database ITTA was sucessfully droped<br>"

On Error Resume Next
Set ret = oDbFDH.Execute( "CREATE TABLE ITTA " & _
						  "	(ITTA_REG INTEGER NOT NULL, " & _
						  "  ITTA_SEQ INTEGER NOT NULL, " & _
						  "  ITTA_REV TEXT(4) NOT NULL, " & _
						  "  ITTA_GER TEXT(10) NOT NULL, " & _ 
						  "	 ITTA_TYPE TEXT(1) NOT NULL, " & _
						  "	 ITTA_TITLE TEXT NOT NULL, " & _
						  "	 ITTA_DATE DATE NOT NULL, " & _
						  "	 ITTA_VALID TEXT(1) NOT NULL, " & _
						  "	 CONSTRAINT AIRStatistics_PK PRIMARY KEY(ITTA_REG,ITTA_SEQ,ITTA_REV,ITTA_GER));" )

If ret Is Nothing Then
	Response.Write "The TABLE ITTA already exists in Database.<br>"
Else
	Response.Write "The TABLE ITTA was created sucessfully in Database!<br>"
End If
On Error GoTo 0

'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Arq: ITTAFile  = "ITTA-" & Reg & "-" & Seq & Rev & "-" & Date & ".pdf"
' Exemplos: ITTA-091-001A-20150913.pdf ITTA-119-002-20150724 ITTA-183-001-20150911.pdf
'

querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 21,	1,	'',	'GTAI',	'I',	'Meios de Cumprimento para Certifica��o de Organiza��o de Produ��o',	#05/16/2016#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 43,	3,	'',	'GCVC',	'B',	'Orienta��es para certifica��o expedita conforme previs�o regulamentar expressa no requisito 43.1(e)-I do RBAC 43',	#08/14/2015#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 91,	1,	'',	'GCVC',	'B',	'Tratamento de inconsist�ncias no Export C of A emitido pela FAA',	#07/14/2015#,	'R' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 91,	1, 'A',	'GCVC',	'B',	'Tratamento de inconsist�ncias no Export C of A emitido por autoridades estrangeiras - Revisado',	#09/23/2015#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 91,	7,	'',	'GCVC',	'I',	'Lanterna port�til para atendimento ao 135.159(f)(3) e 91.503(a)(1)',	#11/23/2015#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 91,	10,	'',	'GCVC',	'I',	'Orienta��o sobre a opera��o em �reas restritas com o uso da tecnologia ADS-B',	#06/08/2016#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 91,	14,	'',	'GCVC',	'I',	'Emiss�o de NCIA � Altera��o na numera��o da NCIA',	#07/19/2016#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 119,	2,	'',	'GCVC',	'B',	'Certifica��o dos operadores RBAC 121 e 135 para uso do Eletronic Flight Bag (EFB)',	#07/24/2015#,	'R' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 119,	2, 'A',	'GCVC',	'B',	'Certifica��o dos operadores RBAC 121 e 135 para uso do Eletronic Flight Bag (EFB) - Revisado',	#05/10/2016#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 119,	5,	'',	'GCVC',	'B',	'Aprova��o de Lista de Equipamentos M�nimos (MEL) contendo menos itens que a Lista Mestra de Equipamentos M�nimos (MMEL)',	#09/30/2015#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 119,	6,	'',	'GCVC',	'B',	'Procedimentos de uso de �toler�ncia� nos intervalos das tarefas dos Programas de Manuten��o Aprovados',	#09/30/2015#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 119,	8,	'',	'GCVC',	'I',	'Programa e Lista de equipamentos e acess�rios da aeronave considerados como n�o essenciais (NEF)',	#11/26/2015#,	'X' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 119,	9,	'',	'GCVC',	'I',	'Per�odo de vac�ncia do cargo de diretor de manuten��o requerido pelo RBAC 119',	#11/30/2015#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 119,	12,	'',	'GCVC',	'I',	'Orienta��o sobre a an�lise de processos de autoriza��o ILS CAT II e III',	#07/14/2016#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 119,	13,	'',	'GCVC',	'I',	'Extens�o de prazo para itens categoria �B� ou �C� da Lista de Equipamentos M�nimos (MEL)',	#07/14/2016#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 121,	11,	'',	'GCVC',	'I',	'Concess�o da etiqueta dimensional e selo ANAC - Programa de Avalia��o Dimensional',	#06/08/2016#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 183,	1,	'',	'GTAS',	'I',	'Estabelecer compet�ncia ao Gerente do �rg�o respons�vel pela coordena��o do Profissional Credenciado, a determina��o do n�mero de Orientados por Orientador',	#09/11/2015#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 183,	2,	'',	'GTAS',	'I',	'Procedimento alternativo ao processo de treinamento pr�tico aos candidatos a PCA e PCF',	#10/16/2015#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 183,	3,	'',	'GTAS',	'I',	'Adequa��o de exig�ncias relacionadas � documenta��o para credenciamento de Profissional Credenciado',	#12/14/2015#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 183,	4,	'',	'GCEN',	'I',	'Remo��o da obrigatoriedade de 1 ano de contato direto com a ANAC para o credenciamento de pessoas f�sicas',	#11/15/2015#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 183,	5,	'',	'GTAS',	'I',	'Procedimento relacionado ao treinamento de Profissional Credenciado',	#06/07/2016#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 183,	6,	'',	'GTAS',	'I',	'Orienta��es sobre comprova��o de atribui��o no CREA para PCFs',	#06/22/2016#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 293,	1,	'',	'GTRAB',	'B',	'Orienta��es relativas � cobran�a do seguro de cargas conforme Resolu��o n� 293/2013',	#01/07/2016#,	'Y' )"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 183, 2, '', 'GTAI', 'I', 'PCF para inspe��o de prot�tipos de artigos.', #06/24/2016#, 'Y' )"
oDbFDH.Execute( querySQL )

Response.Write "Records sucessfully inserted<br>"
Response.End

%>