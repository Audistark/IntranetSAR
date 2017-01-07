<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!-- #include virtual = "/inet2/class/cAAA.asp" -->
<%

Private Sub alert(msg)
%>
<html><head>
<meta http-equiv="content-type" content="text/html; charset=iso-8859-1"/>
<script language="JavaScript" type="text/javascript">
  	alert('<%=msg %>');
</script>
</head><body></body></html>
<%
End Sub

Private Sub redir(href)
%>
<script language="JavaScript" type="text/javascript">
	<% If href <> "" Then %>
    location.href = "<%=href %>";
	<% End If %>
</script>
<%
	Response.End
End Sub


'------------------------------------------------------
'
'	SendMail
'
'

' Init AAA - Authentication, Authorization and Accounting
Dim oAAA : Set oAAA = new cAAA
' Windows Authentication
Dim ret : ret = oAAA.WinAuthenticate(True)
If ret < 0 Then
	oAAA.Print()
End If
' Allow only from Intranet SAR
oAAA.AllowContentFromHost = "sar"
oAAA.AllowContentFromHost = "sar.anac.gov.br"
oAAA.AllowContentFromHost = "sar-dev"
If oAAA.ContentTheft = True Then
	oAAA.Print()
End If

Dim Back : Back = Request.ServerVariables("HTTP_REFERER")

'----------------------------------------------------------
' Only POST method is allowed
If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
	Response.Status = "400 Bad Request"
	Response.End
End If

'----------------------------------------------------------
' Get Arguments
Dim mailSubject : mailSubject = Request.Form("subject")
Dim mailFrom : mailFrom = Request.Form("from")
Dim mailTo : mailTo = Request.Form("to")
Dim mailCc : mailCc = Request.Form("cc")
Dim mailBcc : mailBcc = Request.Form("bcc")
Dim mailText : mailText = Request.Form("text")

'  Mail From using windows user
If mailFrom = "" Then
	mailFrom = oAAA.AuthentUser & "@" & oAAA.AuthentInetDomain(oAAA.AuthentDomain)
End If

'----------------------------------------------------------
' Verify Arguments
If mailText = "" Then
	alert("O Campo texto deve ser preenchido!")
	redir(Back)
End If
If mailFrom = "" Or mailTo = "" Or mailSubject = "" Then
	Response.Status = "400 Bad Request"
	Response.End
End If

'-----------------------------------------------------------------
' CDO (Collaboration Data Objects) is a Microsoft technology that
' is designed to simplify the creation of messaging applications.
Dim oMailConf : Set oMailConf = Server.CreateObject ("CDO.Configuration")
oMailConf.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "relay.anac.gov.br"
oMailConf.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
oMailConf.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'oMailConf.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'oMailConf.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = mailFrom
'oMailConf.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
oMailConf.Fields.Update

Dim oMail : Set oMail = CreateObject("CDO.Message")
Set oMail.Configuration = oMailConf

oMail.Subject	= mailSubject
oMail.From		= mailFrom
oMail.To		= mailTo
oMail.Bcc		= mailCc
oMail.Cc		= mailBcc
oMail.TextBody	= mailText

'''''''''''''''''''''''''oMail.AddAttachment "C:\Scripts\Output.txt"

On Error Resume Next
oMail.Send
If Err.Number <> 0 Then
	Dim msg : msg = Err.Description
	' convert all CRLFs to spaces
	msg = Replace(msg, vbCrLf, " ")
	msg = "Erro Fatal!\n\nO servidor reportou um erro ao tentar enviar o e-mail.\n\n" & _
		  "Err: " & msg & "\n\n" & _
		  "Por favor, contate o administrador do sistema."
	alert(msg)
	redir(Back)
End If
On Error GoTo 0

' Mail is sent - tidy up and delete the AspEmail message object
Set oMail = Nothing
Set oMailConf = Nothing

alert("Email enviado com sucesso!")
redir(Back)

%>
