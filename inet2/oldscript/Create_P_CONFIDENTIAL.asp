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

'''''''''''''''''''' Novo campo CONFIDENTIAL '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM A145_Processos"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("CONFIDENTIAL")
If Err.Number = 0 Then
	Response.Write "The Column CONFIDENTIAL already exists in Database.<br>"
Else
	'--------------------------
	' "T" -> Tramitado 
	' "D" -> Distribuido
	' 
	oDbFDH.Execute( "ALTER TABLE A145_Processos ADD CONFIDENTIAL TEXT(1)" )
	Response.Write "Column CONFIDENTIAL created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''' Novo campo CONFIDENTIAL '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM A135_Processos"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("CONFIDENTIAL")
If Err.Number = 0 Then
	Response.Write "The Column CONFIDENTIAL already exists in Database.<br>"
Else
	'--------------------------
	' "T" -> Tramitado 
	' "D" -> Distribuido
	' 
	oDbFDH.Execute( "ALTER TABLE A135_Processos ADD CONFIDENTIAL TEXT(1)" )
	Response.Write "Column CONFIDENTIAL created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''' Novo campo CONFIDENTIAL '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM A121_Processos"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("CONFIDENTIAL")
If Err.Number = 0 Then
	Response.Write "The Column CONFIDENTIAL already exists in Database.<br>"
Else
	'--------------------------
	' "T" -> Tramitado 
	' "D" -> Distribuido
	' 
	oDbFDH.Execute( "ALTER TABLE A121_Processos ADD CONFIDENTIAL TEXT(1)" )
	Response.Write "Column CONFIDENTIAL created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Response.End

%>