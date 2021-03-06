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



querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
		"( 91, 14, '', 'GCVC', 'I', 'Emiss�o de NCIA � Altera��o na numera��o da NCIA', #07/19/2016#, 'Y')"
oDbFDH.Execute( querySQL )


Response.Write "Records sucessfully inserted<br>"
Response.End




'querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
'			"( 119, 2, 'A', 'GCVC', 'B', 'Certifica��o dos operadores RBAC 121 e 135 para uso do Eletronic Flight Bag (EFB) - Revisado', #05/10/2016#, 'Y')"
'oDbFDH.Execute( querySQL )


querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 91, 7, '', 'GCVC', 'I', 'Lanterna port�til para atendimento ao 135.159(f)(3) e 91.503(a)(1)', #11/23/2015#, 'Y')"
oDbFDH.Execute( querySQL )



Response.Write "Records sucessfully inserted<br>"
Response.End


querySQL = "Update ITTA SET ITTA_DATE = #10/16/2015# WHERE ITTA_REG=183 AND ITTA_SEQ = 2 AND ITTA_REV = ''"
oDbFDH.Execute( querySQL )

querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 183, 2, '', 'GTAS', 'I', 'Procedimento alternativo ao processo de treinamento pr�tico aos candidatos a PCA e PCF', #10/16/2015#, 'Y')"
oDbFDH.Execute( querySQL )


Dim i
For i=1 To 3

	Select Case i
		Case 1
			Sec = "145"
		Case 2
			Sec = "135"
		Case 3
			Sec = "121"
	End Select

''  Set ret = oDbFDH.Execute( "DROP TABLE TMP_" & Sec & "ORGS;" )
''  Response.Write "Database TMP_" & Sec & "ORGS was sucessfully droped<br>"

	'''''''''''''''''''' Create Table TMP_145ORGS ''''''''''''''''''''''''''''''''''
	'
	On Error Resume Next
	Set ret = oDbFDH.Execute( "CREATE TABLE TMP_" & Sec & "ORGS " & _
							  "	(ITEM INTEGER NOT NULL, " & _
							  "	 GTAR			TEXT(10) NOT NULL, " & _
							  "	 ORGP_CODI		TEXT(10) NOT NULL, " & _
							  "	 BSEC_CODI		TEXT(10) NOT NULL, " & _
							  "	 RAZAOSOCIAL	TEXT, " & _
							  "	 NOMEABREV		TEXT NOT NULL, " & _
							  "	 ENDERECO		TEXT, " & _
							  "	 CIDADE			TEXT, " & _
							  "	 UF				TEXT(5), " & _
							  "	 PAIS			TEXT, " & _
							  "	 ADM			TEXT(1), " & _
							  "	 TIPO			TEXT(1), " & _
							  "	 EMAIL			TEXT, " & _
							  "	 CHE			TEXT(20), " & _
							  "	 STATUSCHE		TEXT(1), " & _
							  "	 RCA			TEXT(4), " & _
							  "  AGREEMENTS     TEXT, " & _
							  "	 PADROES		TEXT, " & _
							  "	 STATUSBASE		TEXT(1), " & _
							  "	 ULTAUDIT		TEXT(12), " & _
							  "	 RTNOME			TEXT, " & _
							  "  TSTAMP			DATETIME, " & _
							  "	 CONSTRAINT TMP_" & Sec & "ORGS_PK PRIMARY KEY(ITEM));" )

	If ret Is Nothing Then
		Response.Write "The TABLE TMP_" & Sec & "ORGS already exists in Database.<br>"
	Else
		Response.Write "The TABLE TMP_" & Sec & "ORGS was created sucessfully in Database!<br>"
	End If
	On Error GoTo 0

	'''''''''''''''''''' Create Table TMP_145PROCS ''''''''''''''''''''''''''''''''''
	'
	On Error Resume Next
	Set ret = oDbFDH.Execute( "CREATE TABLE TMP_" & Sec & "PROCS " & _
							  "	(ITEM INTEGER NOT NULL, " & _
							  "	 GTAR			TEXT(10) NOT NULL, " & _
							  "	 NOMEABREV		TEXT NOT NULL, " & _
							  "	 CHE			TEXT(20), " & _
							  "	 PROCESSO		TEXT(16), " & _
							  "	 PROCESSODATA	TEXT(16), " & _
							  "	 SOLICSEQ		INTEGER, " & _
							  "	 TSOLCODI		TEXT(3), " & _
							  "	 SOLICDESCR		TEXT(90), " & _
							  "	 SOLICSTATUS	TEXT(40), " & _
							  "	 ANALISTA		TEXT(35), " & _
							  "	 TASKSEQ		INTEGER, " & _
							  "	 TASKDESCR		TEXT(40), " & _
							  "	 TASKSTATUS		TEXT(40), " & _
							  "	 TASKSTATUSDATA	TEXT(16), " & _
							  "	 TASKDATA		TEXT(16), " & _
							  "	 TASKSPENDDATA	TEXT(16), " & _
							  "	 TASKOBSERV		TEXT, " & _
							  "	 CALC			TEXT(1), " & _
							  "	 TSTAMP			DATETIME, " & _
							  "	 CONSTRAINT TMP_" & Sec & "PROCS_PK PRIMARY KEY(ITEM));" )
	If ret Is Nothing Then
		Response.Write "The TABLE TMP_" & Sec & "PROCS already exists in Database.<br>"
	Else
		Response.Write "The TABLE TMP_" & Sec & "PROCS was created sucessfully in Database!<br>"
	End If

	On Error GoTo 0

Next


'''''''''''''''''''' Create Table ITTA ''''''''''''''''''''''''''''''''''
'
On Error Resume Next
Set ret = oDbFDH.Execute( "CREATE TABLE ITTA " & _
						  "	(ITTA_REG INTEGER NOT NULL, " & _
						  "  ITTA_SEQ INTEGER NOT NULL, " & _
						  "  ITTA_REV TEXT(4) NOT NULL, " & _
						  "  ITTA_GER TEXT(10) NOT NULL, " & _ 
						  "	 ITTA_TYPE TEXT(1) NOT NULL, " & _
						  "	 ITTA_TITLE TEXT NOT NULL, " & _
						  "	 ITTA_GUIDE_HREF TEXT NOT NULL, " & _
						  "	 ITTA_GUIDE_PAGE INTEGER NOT NULL, " & _
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

Response.Write "Nothing to do<br>"
Response.End



querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 119, 7, '', 'GCVC', 'I', 'Per�odo de vac�ncia dos cargos de administra��o requeridos pelo RBAC 119', #09/23/2015#, 'Y')"
oDbFDH.Execute( querySQL )

querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 91, 8, '', 'GCVC', 'B', 'Lanterna', #09/23/2015#, 'Y')"
oDbFDH.Execute( querySQL )


Response.Write "Records sucessfully inserted<br>"
Response.End







'querySQL = "Update ITTA SET ITTA_DATE = #09/11/2015# WHERE ITTA_REG=183 AND ITTA_SEQ = 1 AND ITTA_REV = ''"
'oDbFDH.Execute( querySQL )
'Response.End

querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 91, 1, '', 'GCVC', 'B', 'Tratamento de inconsist�ncias no Export C of A emitido pela FAA', #07/14/2015#, 'R')"
oDbFDH.Execute( querySQL )

querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 91, 1, 'A', 'GCVC', 'B', 'Tratamento de inconsist�ncias no Export C of A emitido por autoridades estrangeiras - Revisado', #09/23/2015#, 'Y')"
oDbFDH.Execute( querySQL )

querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 119, 2, '', 'GCVC', 'B', 'Certifica��o dos operadores RBAC 121 e 135 para uso do Eletronic Flight Bag (EFB)', #07/24/2015#, 'Y')"
oDbFDH.Execute( querySQL )

querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 43, 3, '','GCVC',  'B', 'Orienta��es para certifica��o expedita conforme previs�o regulamentar expressa no requisito 43.1(e)-I do RBAC 43', #08/14/2015#, 'Y')"
oDbFDH.Execute( querySQL )

querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 183, 1, '', 'GTAS', 'I', 'Estabelecer compet�ncia ao Gerente do �rg�o respons�vel pela coordena��o do Profissional Credenciado, a determina��o do n�mero de Orientados por Orientador', #09/11/2015#, 'Y')"
oDbFDH.Execute( querySQL )

querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_NUM, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 119, 5, '', 'GCVC', 'B', '119-005/GCVC', 'Aprova��o de Lista de Equipamentos M�nimos (MEL) contendo menos itens que a Lista Mestra de Equipamentos M�nimos (MMEL)', #09/30/2015#, 'Y')"
oDbFDH.Execute( querySQL )

querySQL = "INSERT INTO ITTA (ITTA_REG, ITTA_SEQ, ITTA_REV, ITTA_GER, ITTA_TYPE, ITTA_TITLE, ITTA_DATE, ITTA_VALID) VALUES " & _
			"( 119, 6, '', 'GCVC', 'B', 'Procedimentos de uso de �toler�ncia� nos intervalos das tarefas dos Programas de Manuten��o Aprovados', #09/30/2015#, 'Y')"
oDbFDH.Execute( querySQL )


Response.Write "Records sucessfully inserted<br>"

Response.End


'''querySQL = "DELETE FROM DTT WHERE DTT_NUM = '91-001-15'"
'''oDbFDH.Execute( querySQL )

'querySQL = "INSERT INTO DTT (DTT_NUM, DTT_TYPE, DTT_TITLE, DTT_DATE, DTT_VALID) VALUES ( '91-001-15', 'B', 'Procedimento para quando identificadas inconsist�ncias no Export C of A emitido pela FAA', '14/07/2015', 'Y')"
'oDbFDH.Execute( querySQL )

'querySQL = "Update DTT SET DTT_TITLE = 'Procedimento para quando identificadas inconsist�ncias no Export C of A utilizado para nacionaliza��o de aeronaves' WHERE DTT_NUM = '91-001-15'"
'querySQL = "Update DTT SET DTT_DATE = '14/07/2015' WHERE DTT_NUM = '91-001-15'"
'oDbFDH.Execute( querySQL )
'querySQL = "Update DTT SET DTT_VALID = 'Y' WHERE DTT_NUM = '91-001-15'"
'oDbFDH.Execute( querySQL )


Response.Write "Record sucessfully inserted<br>"

querySQL = "SELECT * FROM Acesso WHERE AREA = 'ITTA'"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv.Eof Then
	querySQL = "INSERT INTO Acesso (AREA, DTACESSO) VALUES ( 'ITTA', #" & Month(Date()) & "/" & Day(Date()) & "/" & Year(Date) & "#)"
	oDbFDH.Execute( querySQL )
End If
querySQL = "UPDATE Acesso SET DTACESSO = #" & Month(Date()) & "/" & Day(Date()) & "/" & Year(Date) & "# WHERE AREA = 'ITTA'"
oDbFDH.Execute( querySQL )

Response.Status = "200 OK"
Response.End

%>