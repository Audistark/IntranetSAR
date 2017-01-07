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

' Só MASTER
If oAAA.AuthorWinMaster() <> True Then
	Response.Status = "403 Forbidden"
	Response.End
End If

Dim querySQL, rsDiv, col
Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
If oDbFDH.ErrorNumber < 0 then
	oDbFDH.Print()
End If

Response.Status = "200 OK"

Set ret = oDbFDH.Execute( "DROP TABLE A145_TabSolicGroup;" )
Set ret = oDbFDH.Execute( "DROP TABLE A135_TabSolicGroup;" )
Set ret = oDbFDH.Execute( "DROP TABLE A121_TabSolicGroup;" )

Set ret = oDbFDH.Execute( "DROP TABLE A145_TabStatsGroup;" )
Set ret = oDbFDH.Execute( "DROP TABLE A135_TabStatsGroup;" )
Set ret = oDbFDH.Execute( "DROP TABLE A121_TabStatsGroup;" )


'''''''''''''''''''' Novo campo A145_TabSolic '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM A145_TabSolic"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("TSOL_GROUP")
If Err.Number = 0 Then
	Response.Write "The Column TSOL_GROUP already exists in Database.<br>"
Else
	oDbFDH.Execute( "ALTER TABLE A145_TabSolic ADD TSOL_GROUP INTEGER" )
	Response.Write "Column TSOL_GROUP created sucessfully in Database!<br>"
End If
col = rsDiv("TSOL_ENABLED")
If Err.Number = 0 Then
	Response.Write "The Column TSOL_ENABLED already exists in Database.<br>"
Else
	oDbFDH.Execute( "ALTER TABLE A145_TabSolic ADD TSOL_ENABLED TEXT(1)" )
	Response.Write "Column TSOL_ENABLED created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()

'''''''''''''''''''' Novo campo A145_TabTarefa '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM A145_TabTarefa"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("TSK_GROUP")
If Err.Number = 0 Then
	Response.Write "The Column TSK_GROUP already exists in Database.<br>"
Else
	oDbFDH.Execute( "ALTER TABLE A145_TabTarefa ADD TSK_GROUP INTEGER" )
	Response.Write "Column TSK_GROUP created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()

''''''''''''''''''''''''''' A145_TabStatsGroup '''''''''''''''''''''''''''''''''
On Error Resume Next
Set ret = oDbFDH.Execute( "CREATE TABLE A145_TabStatsGroup " & _
						  "	(TStatsGr_Id INTEGER NOT NULL, " & _
						  "	 TStatsGr_Name TEXT NOT NULL, " & _
						  "	 TStatsGr_Goal INTEGER, " & _
						  "	 TStatsGr_ByTask CHAR(1), " & _
						  "	 CONSTRAINT A145_TabStatsGroup_PK PRIMARY KEY(TStatsGr_Id));" )
If ret Is Nothing Then
	Response.Write "The TABLE A145_TabStatsGroup already exists in Database.<br>"
Else

	'	1- Inclusão de Capacidade na EO
'		Incl. de Srv no COM/EO + Audit. Fiscal.
'Remover esse daqui!!!!! E incluir a tarefa auditoria quando tiver.
	'		Inclusão de Serviço no COM/EO
	'		Aceitação de Lista de Capacidade
	'
	'	2- Certificação Inicial e Outros Processos
	'		 Certificação Inicial
	'		 Cadastramento RT e GR
	'		 Execução de Serviço Excepcional
	'		 MNT Fora de Sede
	'
	'	3- Supervisão Continuada - Auditorias
	'	 Auditoria de Supervisão
	'	 Task contendo Auditoria Técnica
	'
	Response.Write "The TABLE A145_TabStatsGroup was created sucessfully in Database!<br>"
	querySQL = "INSERT INTO A145_TabStatsGroup (TStatsGr_Id, TStatsGr_Name, TStatsGr_Goal, TStatsGr_ByTask) VALUES ( 1, 'Inclusão de Capacidade na EO', 25, 'N' )"
	oDbFDH.Execute( querySQL )
	querySQL = "INSERT INTO A145_TabStatsGroup (TStatsGr_Id, TStatsGr_Name, TStatsGr_Goal, TStatsGr_ByTask) VALUES ( 2, 'Supervisão Continuada - Auditorias', 0, 'N' )"
	oDbFDH.Execute( querySQL )
	querySQL = "INSERT INTO A145_TabStatsGroup (TStatsGr_Id, TStatsGr_Name, TStatsGr_Goal, TStatsGr_ByTask) VALUES ( 3, 'Certificação Inicial e Outros Processos', 0, 'S' )"
	oDbFDH.Execute( querySQL )
End If
On Error GoTo 0

'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''' Novo campo A135_TabSolic '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM A135_TabSolic"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("TSOL_GROUP")
If Err.Number = 0 Then
	Response.Write "The Column TSOL_GROUP already exists in Database.<br>"
Else
	oDbFDH.Execute( "ALTER TABLE A135_TabSolic ADD TSOL_GROUP INTEGER" )
	Response.Write "Column TSOL_GROUP created sucessfully in Database!<br>"
End If
col = rsDiv("TSOL_ENABLED")
If Err.Number = 0 Then
	Response.Write "The Column TSOL_ENABLED already exists in Database.<br>"
Else
	oDbFDH.Execute( "ALTER TABLE A135_TabSolic ADD TSOL_ENABLED TEXT(1)" )
	Response.Write "Column TSOL_ENABLED created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()

'''''''''''''''''''' Novo campo A135_TabTarefa '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM A135_TabTarefa"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("TSK_GROUP")
If Err.Number = 0 Then
	Response.Write "The Column TSK_GROUP already exists in Database.<br>"
Else
	oDbFDH.Execute( "ALTER TABLE A135_TabTarefa ADD TSK_GROUP INTEGER" )
	Response.Write "Column TSK_GROUP created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()


''''''''''''''''''''''''''' A135_TabStatsGroup '''''''''''''''''''''''''''''''''
On Error Resume Next
Set ret = oDbFDH.Execute( "CREATE TABLE A135_TabStatsGroup " & _
						  "	(TStatsGr_Id INTEGER NOT NULL, " & _
						  "	 TStatsGr_Name TEXT NOT NULL, " & _
						  "	 TStatsGr_Goal INTEGER, " & _
						  "	 TStatsGr_ByTask CHAR(1), " & _
						  "	 CONSTRAINT A135_TabStatsGroup_PK PRIMARY KEY(TStatsGr_Id));" )
If ret Is Nothing Then
	Response.Write "The TABLE A135_TabStatsGroup already exists in Database.<br>"
Else
 
	'	1- Inclusão de Capacidade na EO
	'	   Inclusão de Aeronave na EO
	'	   Inclusão de Nova Operação na EO
	'	   Inclusão de Novo Modelo de Aeronave
	'
	'	2- Supervisão Continuada - Auditorias
	'	   Auditoria PTA
	'	   Auditoria por Demanda
	'
	Response.Write "The TABLE A135_TabStatsGroup was created sucessfully in Database!<br>"
	querySQL = "INSERT INTO A135_TabStatsGroup (TStatsGr_Id, TStatsGr_Name, TStatsGr_Goal, TStatsGr_ByTask) VALUES ( 1, 'Inclusão de Capacidade na EO', 25, 'N' )"
	oDbFDH.Execute( querySQL )
	querySQL = "INSERT INTO A135_TabStatsGroup (TStatsGr_Id, TStatsGr_Name, TStatsGr_Goal, TStatsGr_ByTask) VALUES ( 2, 'Supervisão Continuada - Auditorias', 0, 'N' )"
	oDbFDH.Execute( querySQL )
End If
On Error GoTo 0

'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''' Novo campo A121_TabSolic '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM A121_TabSolic"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("TSOL_GROUP")
If Err.Number = 0 Then
	Response.Write "The Column TSOL_GROUP already exists in Database.<br>"
Else
	oDbFDH.Execute( "ALTER TABLE A121_TabSolic ADD TSOL_GROUP INTEGER" )
	Response.Write "Column TSOL_GROUP created sucessfully in Database!<br>"
End If
col = rsDiv("TSOL_ENABLED")
If Err.Number = 0 Then
	Response.Write "The Column TSOL_ENABLED already exists in Database.<br>"
Else
	oDbFDH.Execute( "ALTER TABLE A121_TabSolic ADD TSOL_ENABLED TEXT(1)" )
	Response.Write "Column TSOL_ENABLED created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()

'''''''''''''''''''' Novo campo A121_TabTarefa '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM A121_TabTarefa"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("TSK_GROUP")
If Err.Number = 0 Then
	Response.Write "The Column TSK_GROUP already exists in Database.<br>"
Else
	oDbFDH.Execute( "ALTER TABLE A121_TabTarefa ADD TSK_GROUP INTEGER" )
	Response.Write "Column TSK_GROUP created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()


''''''''''''''''''''''''''' A121_TabStatsGroup '''''''''''''''''''''''''''''''''
On Error Resume Next
Set ret = oDbFDH.Execute( "CREATE TABLE A121_TabStatsGroup " & _
						  "	(TStatsGr_Id INTEGER NOT NULL, " & _
						  "	 TStatsGr_Name TEXT NOT NULL, " & _
						  "	 TStatsGr_Goal INTEGER, " & _
						  "	 TStatsGr_ByTask CHAR(1), " & _
						  "	 CONSTRAINT A121_TabStatsGroup_PK PRIMARY KEY(TStatsGr_Id));" )
If ret Is Nothing Then
	Response.Write "The TABLE A121_TabStatsGroup already exists in Database.<br>"
Else

	'	1- Inclusão de Capacidade na EO
	'	 Inclusão de Aeronave na EO
	'	 Inclusão de Nova Operação na EO
	'	 Inclusão de Novo Modelo de Aeronave
	'
	'	2- Supervisão Continuada
	'	 Auditoria PTA
	'	 Auditoria por Demanda
	'
	Response.Write "The TABLE A121_TabStatsGroup was created sucessfully in Database!<br>"
	querySQL = "INSERT INTO A121_TabStatsGroup (TStatsGr_Id, TStatsGr_Name, TStatsGr_Goal, TStatsGr_ByTask) VALUES ( 1, 'Inclusão de Capacidade na EO', 25, 'N' )"
	oDbFDH.Execute( querySQL )
	querySQL = "INSERT INTO A121_TabStatsGroup (TStatsGr_Id, TStatsGr_Name, TStatsGr_Goal, TStatsGr_ByTask) VALUES ( 2, 'Supervisão Continuada', 0, 'N' )"
	oDbFDH.Execute( querySQL )
End If
On Error GoTo 0

'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Set ret = oDbFDH.Execute( "DROP TABLE AIRStats;" )
Set ret = oDbFDH.Execute( "DROP TABLE AIRStatistics;" )

'''''''''''''''''''' Create Table AIRStats ''''''''''''''''''''''''''''''''''
'
On Error Resume Next
Set ret = oDbFDH.Execute( "CREATE TABLE AIRStats " & _
						  "	(Stats_DATE DATE NOT NULL, " & _
						  "	 Stats_GTAR TEXT(10) NOT NULL, " & _
						  "	 Stats_RBAC TEXT(10) NOT NULL, " & _
						  "	 Stats_TYPE TEXT(1) NOT NULL, " & _
						  "	 Stats_CODI TEXT(3) NOT NULL, " & _
						  "	 Stats_ANAC INTEGER, " & _
						  "  Stats_GOAL INTEGER, " & _
						  "  Stats_DELAY_ANAC INTEGER, " & _
						  "  Stats_MAX_ANAC INTEGER, " & _
						  "  Stats_DAYS_ANAC INTEGER, " & _
						  "  Stats_30D_CLOSED INTEGER, " & _
						  "  Stats_ITR_CLOSED INTEGER, " & _
						  "  Stats_30D_DOCS INTEGER, " & _
						  "  Stats_CLIENT INTEGER, " & _
						  "  Stats_DELAYED_CLIENT INTEGER, " & _
						  "  Stats_MAX_CLIENT INTEGER, " & _
						  "  Stats_TIMESTAMP DATETIME, " & _
						  "	 CONSTRAINT AIRStats_PK PRIMARY KEY(Stats_DATE,Stats_GTAR,Stats_RBAC,Stats_TYPE,Stats_CODI));" )
If ret Is Nothing Then
	Response.Write "The TABLE AIRStats already exists in Database.<br>"
Else
	Response.Write "The TABLE AIRStats was created sucessfully in Database!<br>"
End If
On Error GoTo 0

'querySQL = "UPDATE AIRStats SET TStatsGr_Goal=25 WHERE TStatsGr_Goal=15"
'oDbFDH.Execute( querySQL )




'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Response.End
%>