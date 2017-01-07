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


'''''''''''''''''''' Novo campo TST_RESTRICT '''''''''''''''''''''''''''''''''''
'
' http://sar/inet2/script/Create_TST_RESTRICT.asp
'
' Esse campo é utilizado para status que só podem ser incluídos pelo Líder
' ou Master, ou Administrativo, seguindo regras específicas de negócio.
'
Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
If oDbFDH.ErrorNumber < 0 then
	oDbFDH.Print()
End If

Dim querySQL
Response.Status = "200 OK"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 145
querySQL = "SELECT * FROM A145_TabStatus"
Dim rsDiv : Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("TST_RESTRICT")
If Err.Number = 0 Then
	Response.Write "The Column TST_RESTRICT in A145_TabStatus already exists in Database.<BR>"
Else
	oDbFDH.Execute( "ALTER TABLE A145_TabStatus ADD TST_RESTRICT TEXT(1)" )
	Response.Write "Column TST_RESTRICT created sucessfully in Database!<BR>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 135
querySQL = "SELECT * FROM A135_TabStatus"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("TST_RESTRICT")
If Err.Number = 0 Then
	Response.Write "The Column TST_RESTRICT in A135_TabStatus already exists in Database.<BR>"
Else
	oDbFDH.Execute( "ALTER TABLE A135_TabStatus ADD TST_RESTRICT TEXT(1)" )
	Response.Write "Column TST_RESTRICT created sucessfully in Database!<BR>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 121
querySQL = "SELECT * FROM A121_TabStatus"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("TST_RESTRICT")
If Err.Number = 0 Then
	Response.Write "The Column TST_RESTRICT in A121_TabStatus already exists in Database.<BR>"
Else
	oDbFDH.Execute( "ALTER TABLE A121_TabStatus ADD TST_RESTRICT TEXT(1)" )
	Response.Write "Column TST_RESTRICT created sucessfully in Database!<BR>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Response.End

%>