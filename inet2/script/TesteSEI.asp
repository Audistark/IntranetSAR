<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!-- #include virtual = "/inet2/class/cAAA.asp" -->
<!-- #include virtual = "/inet2/class/cSEI.asp" -->
<%
'
' Init AAA - Authentication, Authorization and Accounting
Dim oAAA : Set oAAA = new cAAA
Dim SDiv : SDiv = oAAA.AuthentWinUserSDiv

Dim oSEI : Set oSEI = new cSEI
Dim ret

Dim args : args = Request.QueryString( "ProtoDoc" )
Dim ProtoDoc
If args <> "" Then
	ProtoDoc = CInt(args)
Else
	ProtoDoc = 97
End If

Response.Status = "200 OK"


oSEI.ConsultarDocumento(ProtoDoc)
ret = oSEI.numResult
Response.Write "ret = " & ret & "<br>"
If ret > 0 Then
'	Response.Clear()
	Response.Write "IdProcedimento: " & oSEI.IdProcedimento() & "<br>"
	Response.Write "ProcedimentoFormatado: " & oSEI.ProcedimentoFormatado() & "<br>"
	Response.Write "IdDocumento: " & oSEI.IdDocumento() & "<br>"
	Response.Write "DocumentoFormatado: " & oSEI.DocumentoFormatado() & "<br>"
	Response.Write "LinkAcesso: " & oSEI.LinkAcesso() & "<br>"
	Response.Write "SerieId: " & oSEI.SerieId() & "<br>"
	Response.Write "SerieNome: " & oSEI.SerieNome() & "<br>"
	Response.Write "Numero: " & oSEI.Numero() & "<br>"
	Response.Write "Data: " & oSEI.Data() & "<br>"
Else
	Response.Write "ProtoDoc " & ProtoDoc & " not found!<br>"
End If

Response.End


'-------------------------------------------------------------------
' Le as unidades no SEI
'
ret = oSEI.ListarUnidades()
If ret = 0 Then
	Response.Write "Não foi possível ler as unidades no SEI.<br>"
	Response.End
ElseIf ret < 0 Then
	Response.Write "Não foi possível ler as unidades no SEI.<br>"
	Response.Write "Erro de acesso ao SEI.\n"
	Response.Write "Aguarde alguns instantes e tente novamente."
	Response.End
End If

Dim m_Unidades
m_Unidades = oSEI.GetUnidades
Response.Write "Unidades = '" & m_Unidades & "'<br>"

'
'-------------------------------------------------------------------

Response.End

%>
