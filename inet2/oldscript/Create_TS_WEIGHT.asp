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


'''''''''''''''''''' Novo campo TS_WEIGHT '''''''''''''''''''''''''''''''''''
'
' http://sar/inet2/script/Create_TS_WEIGHT.asp
'
' Esse campo é utilizado para o peso de cada solicitação em Homem.Dia
'
Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
If oDbFDH.ErrorNumber < 0 then
	oDbFDH.Print()
End If

Dim querySQL
Response.Status = "200 OK"

' lixos
oDbFDH.Execute( "DELETE * FROM A145_TabTarefa WHERE TSK_CODI = '007'" )
oDbFDH.Execute( "DELETE * FROM A145_TabTarefa WHERE TSK_CODI = '008'" )
oDbFDH.Execute( "DELETE * FROM A145_TabTarefa WHERE TSK_CODI = '009'" )

oDbFDH.Execute( "ALTER TABLE A145_TabSolic DROP COLUMN TSOL_AUDIT" )
oDbFDH.Execute( "ALTER TABLE A135_TabSolic DROP COLUMN TSOL_AUDIT" )
oDbFDH.Execute( "ALTER TABLE A121_TabSolic DROP COLUMN TSOL_AUDIT" )


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 145
querySQL = "SELECT * FROM A145_TarefaSolic"
Dim rsDiv : Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("TS_WEIGHT")
If Err.Number = 0 Then
	Response.Write "The Column TS_WEIGHT in A145_TarefaSolic already exists in Database.<BR>"
Else
	oDbFDH.Execute( "ALTER TABLE A145_TarefaSolic ADD TS_WEIGHT FLOAT" )
	Response.Write "Column TS_WEIGHT created sucessfully in Database!<BR>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 135
querySQL = "SELECT * FROM A135_TarefaSolic"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("TS_WEIGHT")
If Err.Number = 0 Then
	Response.Write "The Column TS_WEIGHT in A135_TarefaSolic already exists in Database.<BR>"
Else
	oDbFDH.Execute( "ALTER TABLE A135_TarefaSolic ADD TS_WEIGHT FLOAT" )
	Response.Write "Column TS_WEIGHT created sucessfully in Database!<BR>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 121
querySQL = "SELECT * FROM A121_TarefaSolic"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("TS_WEIGHT")
If Err.Number = 0 Then
	Response.Write "The Column TS_WEIGHT in A121_TarefaSolic already exists in Database.<BR>"
Else
	oDbFDH.Execute( "ALTER TABLE A121_TarefaSolic ADD TS_WEIGHT FLOAT" )
	Response.Write "Column TS_WEIGHT created sucessfully in Database!<BR>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Response.End

%>