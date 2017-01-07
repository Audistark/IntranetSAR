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

'''''''''''''''''''' Novos campos SDIV_EMAIL e SDIV_PHONE '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM Tab_Subdivisao"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("SDIV_EMAIL")
If Err.Number = 0 Then
	Response.Write "The Column SDIV_EMAIL already exists in Database.<br>"
Else
	oDbFDH.Execute( "ALTER TABLE Tab_Subdivisao ADD SDIV_EMAIL TEXT(32)" )
	Response.Write "Column SDIV_EMAIL created sucessfully in Database!<br>"
End If
col = rsDiv("SDIV_PHONE")
If Err.Number = 0 Then
	Response.Write "The Column SDIV_PHONE already exists in Database.<br>"
Else
	oDbFDH.Execute( "ALTER TABLE Tab_Subdivisao ADD SDIV_PHONE TEXT(32)" )
	Response.Write "Column SDIV_PHONE created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Response.End

%>