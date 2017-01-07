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

Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
If oDbFDH.ErrorNumber < 0 then
	oDbFDH.Print()
End If

Dim tSec(3) : tSec(0) = "121" : tSec(1) = "135" : tSec(2) = "145"
Dim i

'''''''''''''''''''' Novo campo D145_DTMOV '''''''''''''''''''''''''''''''''''
'
For i=0 To 2
	Dim querySQL : querySQL = "SELECT * FROM A" & tSec(i) & "_Documentos"
	Dim rsDiv : Set rsDiv = oDbFDH.getRecSetRd(querySQL)
	If rsDiv Is Nothing then
		oDbFDH.Print()
	End If
	On Error Resume Next
	Dim DtMov : DtMov = rsDiv("D" & tSec(i) & "_DTMOV")
	If Err.Number = 0 Then
		Response.Status = "200 OK"
		Response.Write "The Column D" & tSec(i) & "_DTMOV already exists in Database.<br>"
	Else
		oDbFDH.Execute( "ALTER TABLE A" & tSec(i) & "_Documentos ADD D" & tSec(i) & "_DTMOV DATE" )
		Response.Status = "200 OK"
		Response.Write "Column D" & tSec(i) & "_DTMOV created sucessfully in Database!<br>"
	End If
	On Error GoTo 0
	rsDiv.Close()
Next
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Response.End
%>