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

Dim Sec : Sec = Request.QueryString( "Sec" )
If Sec <> "145" And Sec <> "121" And Sec <> "135" And Sec <> "91" Then
	Response.Status = "400 Bad Request"
	Response.Write "Argumentos inválidos. Favor informar o 'Sec'"
	Response.End
End If

'''''''''''''''''''' Novo campo D_145_TPDOC '''''''''''''''''''''''''''''''''''
'
Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
If oDbFDH.ErrorNumber < 0 then
	oDbFDH.Print()
End If
Dim querySQL : querySQL = "SELECT * FROM A" & Sec & "_Documentos"
Dim rsDiv : Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
TpDoc = rsDiv("D" & Sec & "_TPDOC")
If Err.Number = 0 Then
	Response.Status = "200 OK"
	Response.Write "The Column D" & Sec & "_TPDOC already exists in Database."
Else
	oDbFDH.Execute( "ALTER TABLE A" & Sec & "_Documentos ADD D" & Sec & "_TPDOC TEXT" )
	Response.Status = "200 OK"
	Response.Write "Column D" & Sec & "_TPDOC created sucessfully in Database!"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Response.End
%>