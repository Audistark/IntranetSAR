<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!-- #include virtual = "/inet2/class/cAAA.asp" -->
<!-- #include virtual = "/inet2/class/cSIGAD.asp" -->
<%
'
' Init AAA - Authentication, Authorization and Accounting
Dim oAAA : Set oAAA = new cAAA
Dim SDiv : SDiv = oAAA.AuthentWinUserSDiv

Protocolo  = Request.QueryString( "Proto" )
Response.Write "Protocolo = '" & Protocolo & "'<br>"

Dim DiaMov : DiaMov = ""
Dim MesMov : MesMov = ""
Dim AnoMov : AnoMov = ""

Dim ProtoOrgaoCode : ProtoOrgaoCode = ""
Dim ProtoCode : ProtoCode   = ""

Dim oSIGAD : Set oSIGAD = new cSIGAD
Dim ret

Response.Status = "200 OK"

'-------------------------------------------------------------------
' Le o Protocolo
'
ret = oSIGAD.GetProtocol(Protocolo)
If ret = 0 Then
	Response.Write "Não foi possível ler o protocolo '" & Protocolo & "' no SIGAD.<br>"
	Response.Write "Protocolo não cadastrado."
	Response.End
ElseIf ret < 0 Then
	Response.Write "Não foi possível ler o protocolo '" & Protocolo & "' no SIGAD.<br>"
	Response.Write "Erro de acesso ao SIGAD.\n"
	Response.Write "Aguarde alguns instantes e tente novamente."
	Response.End
End If

ProtoOrgaoCode = oSIGAD.OrgaoCode
Response.Write "ProtoOrgaoCode = '" & ProtoOrgaoCode & "'<br>"

ProtoCode = oSIGAD.ProtoCode
Response.Write "ProtoCode = '" & ProtoCode & "'<br>"


'
'-------------------------------------------------------------------


' Data do movimento
Dim dtMov : dtMov = oSIGAD.GetLastMoveDate( ProtoCode, ProtoOrgaoCode )

DiaMov = Day( dtMov )
MesMov = Month( dtMov )
AnoMov = Year( dtMov )

Response.Write "Data Movimento = '" & dtMov & "'.<br>"
Response.End

%>
