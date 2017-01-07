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

'''''''''''''''''''' Novo campo T145_EMAIL '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM A145_Tarefas"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("T145_EMAIL")
If Err.Number = 0 Then
	Response.Write "The Column T145_EMAIL already exists in Database.<br>"
Else
	oDbFDH.Execute( "ALTER TABLE A145_Tarefas ADD T145_EMAIL TEXT(1)" )
	Response.Write "Column T145_EMAIL created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''' Novo campo T135_EMAIL '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM A135_Tarefas"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("T135_EMAIL")
If Err.Number = 0 Then
	Response.Write "The Column T135_EMAIL already exists in Database.<br>"
Else
	oDbFDH.Execute( "ALTER TABLE A135_Tarefas ADD T135_EMAIL TEXT(1)" )
	Response.Write "Column T135_EMAIL created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''' Novo campo T121_EMAIL '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM A121_Tarefas"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("T121_EMAIL")
If Err.Number = 0 Then
	Response.Write "The Column T121_EMAIL already exists in Database.<br>"
Else
	oDbFDH.Execute( "ALTER TABLE A121_Tarefas ADD T121_EMAIL TEXT(1)" )
	Response.Write "Column T121_EMAIL created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Response.End
%>