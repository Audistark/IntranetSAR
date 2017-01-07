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


'''''''''''''''''''' Novo campo D145_FILE '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM A145_Documentos"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("D145_FILE")
If Err.Number = 0 Then
	Response.Write "The Column D145_FILE already exists in Database.<br>"
    querySQL = "UPDATE ( ( A145_Bases INNER JOIN (A145_Processos INNER JOIN A145_Documentos" & _
               "     ON A145_Processos.P145_CODI = A145_Documentos.P145_CODI)" & _
               "     ON A145_Bases.B145_CODI = A145_Processos.B145_CODI) INNER JOIN Pessoal" & _
               "     ON A145_Bases.PES_CODI = Pessoal.PES_CODI) INNER JOIN Tab_Subdivisao" & _
               "     ON Pessoal.SDIV_CODI = Tab_Subdivisao.SDIV_CODI" & _
               " SET A145_Documentos.D145_FILE='1'" & _
               " WHERE Tab_Subdivisao.SDIV_SIGLA='GTAR-DF'"
    oDbFDH.Execute(querySQL)
Else
	oDbFDH.Execute( "ALTER TABLE A145_Documentos ADD D145_FILE TEXT(1)" )
	Response.Write "Column D145_FILE created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''' Novo campo D135_FILE '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM A135_Documentos"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("D135_FILE")
If Err.Number = 0 Then
	Response.Write "The Column D135_FILE already exists in Database.<br>"
    querySQL = "UPDATE ( ( A135_Bases INNER JOIN (A135_Processos INNER JOIN A135_Documentos" & _
               "     ON A135_Processos.P135_CODI = A135_Documentos.P135_CODI)" & _
               "     ON A135_Bases.B135_CODI = A135_Processos.B135_CODI) INNER JOIN Pessoal" & _
               "     ON A135_Bases.PES_CODI = Pessoal.PES_CODI) INNER JOIN Tab_Subdivisao" & _
               "     ON Pessoal.SDIV_CODI = Tab_Subdivisao.SDIV_CODI" & _
               " SET A135_Documentos.D135_FILE='1'" & _
               " WHERE Tab_Subdivisao.SDIV_SIGLA='GTAR-DF'"
    oDbFDH.Execute(querySQL)
Else
	oDbFDH.Execute( "ALTER TABLE A135_Documentos ADD D135_FILE TEXT(1)" )
	Response.Write "Column D135_FILE created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''' Novo campo D121_FILE '''''''''''''''''''''''''''''''''''
'
querySQL = "SELECT * FROM A121_Documentos"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
col = rsDiv("D121_FILE")
If Err.Number = 0 Then
	Response.Write "The Column D121_FILE already exists in Database.<br>"
Else
	oDbFDH.Execute( "ALTER TABLE A121_Documentos ADD D121_FILE TEXT(1)" )
	Response.Write "Column D121_FILE created sucessfully in Database!<br>"
End If
On Error GoTo 0
rsDiv.Close()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Response.End
%>
