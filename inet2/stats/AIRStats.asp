<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<% Response.CodePage = 1252 %>
<% session.LCID = 1046 'BRASIL %>

<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!--#include virtual = "/inet2/class/cAAA.asp" -->

<!-- #include virtual = "/inet2/stats/cStatistics.asp" -->
<!-- #include virtual = "/inet2/stats/cStCalc.asp" -->
<!-- #include virtual = "/inet2/stats/cStData.asp" -->

<%

' http://sar/inet2/stats/AIRStats.asp?gtar=GTAR-DF&rbac=145&silent=1&send=1

Dim rbac : rbac = Request.QueryString("rbac")
If rbac <> "121" And rbac <> "135" And rbac <> "145" Then
	rbac = "121"
End If

Dim strGtar : strGtar = ""
Dim gtar : gtar = Request.QueryString("gtar")
If gtar <> "GTAR-DF" And gtar <> "GTAR-SP" And gtar <> "GTAR-RJ" Then
	gtar = "SAR"
Else
	strGtar =  "/" & gtar
End If

Dim silent : silent = False
If Request.QueryString("silent") = "1" Then
	silent = True
End If

' actions
Const cStNone   = 0
Const cStSend   = 1
Const cStMail   = 2
Dim bSend : bSend = False
Dim action : action = cStNone
If Request.QueryString("send") = "1" Then
	action = cStSend
	bSend = True
ElseIf Request.QueryString("send") = "2" Then
	action = cStMail
	bSend = True
End If

'----------------------------------------------------------
' Init AAA - Authentication, Authorization and Accounting
Dim oAAA : Set oAAA = new cAAA
' Windows Authentication
Dim ret : ret = oAAA.WinAuthenticate(True)
If ret < 0 Then
	oAAA.Print()
End If

Dim optSendMail : optSendMail = False
Dim allowSendMail : allowSendMail = False
If oAAA.AuthorWinMasterSec(rbac) = True And gtar = "SAR" Then
	optSendMail = True
	Application.Lock
	Dim dtMail : dtMail = CDate( Application("AIRStats" & rbac & "TStamp") )
	Dim dDiff : dDiff = DateDiff("d", dtMail, Now)
	If dDiff >= 1 Then
		allowSendMail = True
	End If
	Application.UnLock
End If

Dim querySQL, rsDiv
Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
If oDbFDH.ErrorNumber < 0 then
	oDbFDH.Print()
End If

Dim oStats(10)
Dim nGroups : nGroups = 0
querySQL =	"SELECT * FROM A" & rbac & "_TabStatsGroup ORDER By TStatsGr_Id"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
Do While Not rsDiv.Eof
	Dim iGroup : iGroup = rsDiv( "TStatsGr_Id" )
	Set oStats(nGroups) = (new cStatistics)( Array(rbac, gtar, iGroup, bSend) )
	If oStats(nGroups).Ret() > 0 Then
		If nGroups < 10 Then
			nGroups = nGroups + 1
		Else
			Exit Do
		End If
	End If
	rsDiv.MoveNext
Loop
rsDiv.Close()

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

Private Sub closeWin()
%>
<script language="JavaScript" type="text/javascript">
	window.close();
</script>
<%
	Response.End
End Sub

If action <> cStMail Then %>

<!DOCTYPE html>
<html>
<head>

  <title>Statistics<%=strGtar %> - RBAC<%=rbac %></title>

  <style type="text/css">
		.linha
		{
			border-bottom: 1px solid black;
			margin: 2px 12px 7px 12px;
		}
		.linha2
		{
			border-bottom: 1px dashed #c9cacb;
			margin: 7px 12px 7px 12px;
			width: 95%;
		}
		a:link {
			border-style: none;
			color: blue;
			text-decoration: none;
			font-style: normal;
			font-weight: normal;
		}
		a:visited {
			color: blue;
			text-decoration: none;
			font-style: normal;
			font-weight: normal;
		}
		a:hover {
			color: red;
			text-decoration: none;
			font-style: normal;
			font-weight: normal;
		}
		a:active {
			color: blue;
			text-decoration: none;
			font-style: normal;
			font-weight: normal;
		}
		.SarFont {
			font-family: Calibri;
			font-size: 16px;
			color: red;
			font-style: normal;
			font-weight: normal;
		}
		.SarText 
		{
			font-family: Calibri;
			font-size: 12px;
			color: black;
			font-style: normal;
			font-weight: normal;
		}
  </style>
  
  <!--Load the AJAX API-->
  <script type="text/javascript" src="https://www.google.com/jsapi"></script>

  <script type="text/javascript">
	function refreshEvery3Hours() {
		var now = new Date();
		if (now.getHours() % 3 == 0 && now.getMinutes() < 2 ) {
			window.location.reload(true);
		}
	}
	var myCheck = setInterval(function () { refreshEvery3Hours() }, 60000);<%

	Dim i
	If action = cStSend Then
	%>
	var myCkFile = setInterval(function () { checkFiles() }, 1000);
	function checkFiles() {<%
		For i=0 To nGroups-1
			If oStats(i).IsPriority() Then %>
		if( isImgSaved<%=oStats(i).GetObjId() %>() == false ) {
			return false;
		}<%
			End If
		Next %>
		clearInterval(myCheck);
		clearInterval(myCkFile);
		document.getElementById("IconSendMail").src = "../img/icons/glyphicons_128_message_lock.png";
		document.getElementById("IconSendMail").title = "Email j� enviado";
		document.getElementById("IconSendMail").width = 28;
		document.getElementById("IconSendMail").height = 24;
		window.open('AIRStats.asp?rbac=<%=rbac %>&gtar=<%=gtar %>&send=<%=cStMail %>', '_blank');
		return true;
	}
	<%
	End If %>

  	// zoom
  	var fs = window.top.document.getElementsByTagName("frameset");
  	fs[1].cols = "8,*"
  	function clearBkdFrame(id) {
  		var x = document.getElementById(id);
  		var y = (x.contentWindow || x.contentDocument);
  		if (y.document) y = y.document;
  		y.body.style.background = "white none repeat";
  	}

	function sendmail() {
		if (!confirm('Deseja enviar e-mail das estat�sticas aos gestores?\n\n')) {
			return;
		}
		else {
			window.location.href = "AIRStats.asp?rbac=<%=rbac %>&gtar=<%=gtar %>&send=<%=cStSend %>"
		}
	}

  </script>

<%
	For i=0 To nGroups-1
		oStats(i).JavaScripts()
		oStats(i).Gauge()
		oStats(i).Charts()
		oStats(i).PieChart()
	Next
 %>

</head>
<body>

	<%
	If Not silent Then %>

  <div class="linha"></div>

  <table border="0" cellpadding="0" cellspacing="0" width="98%">
    <tr align="center" valign="middle">
        <td><img src=/inet2/img/statpic.jpg border=0 width=82 height=82></td>
        <td align="center"><font size="8" face="Calibri">Statistics - SAR<%=strGtar %> - RBAC <%=rbac %></font></td>
		<td align="center" valign="bottom"><table border="0" cellpadding="1" cellspacing="0">
			<tr>
				<td align="center" class="SarFont"><%
				Dim tRbac : tRbac = Array("121", "135", "145")
				For i=0 To 2
					Response.Write( " " )
					If tRbac(i) <> rbac Or action = cStSend Then Response.Write("<a href='AIRStats.asp?rbac=" & tRbac(i) & "&gtar=" & gtar & "'>")
					Response.Write( "RBAC" & tRbac(i) )
					If tRbac(i) <> rbac Or action = cStSend Then Response.Write("</a>")
					Response.Write( " " )
				Next %>
				</td>
			</tr>
			<tr>
				<td align="center" class="SarFont"><%
				Dim tGtar : tGtar = Array("GTAR-DF", "GTAR-SP", "GTAR-RJ", "SAR")
				For i=0 To 3
					Response.Write( " " )
					If tGtar(i) <> gtar Or action = cStSend Then Response.Write("<a href='AIRStats.asp?rbac=" & rbac & "&gtar=" & tGtar(i) & "'>")
					Response.Write( tGtar(i) )
					If tGtar(i) <> gtar Or action = cStSend Then Response.Write("</a>")
					Response.Write( " " )
				Next %>
				</td>
			</tr>
		</table>
		<td align="center" valign="bottom"><table border="0" cellpadding="1" cellspacing="0">
			<tr>
				<td align="center"><%
				If optSendMail = True Then
					If allowSendMail = False Then %>
					<img src="../img/icons/glyphicons_128_message_lock.png" width="28" height="24" border="0" id="IconSendMail" title="Email j� enviado"><%
					ElseIf action = cStNone Then %>
					<img src="../img/icons/glyphicons_129_message_new.png" width="28" height="22" border="0" id="IconSendMail" title="Send Email" OnClick="sendmail();"><%
					ElseIf action = cStSend Then %>
					<img src="../img/icons/glyphicons_041_charts.png" width="27" height="24" border="0" id="IconSendMail" title="Send Email"><%
					End If
				End If %>
				</td>
			</tr>
		</table>
		<td align="right" valign="bottom"><font size="3" face="Calibri">Date: <%=Now() %><br>Version 2.0<br>GCVC/GGAC/SAR</font></td>
    </tr>
  </table>

  <div class="linha"></div>

	<%
	End If

	For i=0 To nGroups-1

		oStats(i).PrintHtml(True) %>

  <center><div class="linha2"></div></center>

		<%

	Next
	
	%>

  <table border="0" cellpadding="0" cellspacing="0" width="98%">
    <tr>
	  <td align="left"><font size="2" face="Calibri">
	Observa��es:<br>
	1. Atrasos da ANAC s�o computados a partir da meta estabelecida ou depois de 60 dias se n�o houver meta para o grupo estat�stico.<br>
	2. Atrasos do Cliente s�o computados a partir do prazo estabelecido pela ANAC + uma toler�ncia de 30 dias;<br>
	3. Os cursores dos gauges indicam a posi��o dos processos em rela��o aos atrasos, sejam da ANAC ou do Cliente;<br>
	4. O 'gr�fico de tempos em dias' � mostrado apenas quando h� meta estabelecida para o grupo estat�stico;<br>
	5. O alarme dispara se houver algum processo que n�o esteja atendendo a meta estabelecida;<br>
	6. Os '30d' Docs e Closed mostram o n�mero de documentos emitidos e processos conclu�dos pela ANAC nos �ltimso 30 dias;<br>
	  </font></td>
	</tr>
  </table>

</body>
</html>

<%
Else ' Send mail (step 2)

	Dim mailSubject : mailSubject = "[SAR-GGAC/GCVC] Estat�sticas dos processos RBAC " & rbac & " priorit�rios da GGAC"
	Dim hs : hs = Hour(Now())
	Dim Greetings
	If hs < 13 Then
		Greetings = "bom dia"
	ElseIf hs < 18 Then
		Greetings = "boa tarde"
	Else
		Greetings = "boa noite"
	End If
	Dim mailBody : mailBody = "<html>" & vbCrLf & _
		"<p><font size=""3"" face=""Calibri"">Senhores(as), " & Greetings & ".</font></p>" & vbCrLf & _
		"<p><font size=""3"" face=""Calibri"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;A seguir as estat�sticas atualizadas dos processos priorit�rios da GGAC: </font></p>" & vbCrLf

	For i=0 To nGroups-1

		oStats(i).PrintHtml(False)

		mailBody = mailBody & oStats(i).GetMail

	Next

	mailBody = mailBody & "<br>" & vbCrLf & _
			"  <table border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf & _
			"    <tr>" & vbCrLf & _
			"	  <td align=""left""><font size=""2"" face=""Calibri"">" & vbCrLf & _
			"	Observa��es:<br>" & vbCrLf & _
			"	1. Atrasos da ANAC s�o computados a partir da meta estabelecida.<br>" & vbCrLf & _
			"	2. Atrasos do Cliente s�o computados a partir do prazo estabelecido pela ANAC + uma toler�ncia de 30 dias;<br>" & vbCrLf & _
			"	3. Os '30d' Docs e Closed mostram o n�mero de documentos emitidos e processos conclu�dos pela ANAC nos �ltimos 30 dias.<br>" & vbCrLf & _
			"   Maiores informa��es acerca dos dados estat�sticos podem ser obtidas em http://sar/inet2/stats/AIRStats.asp<br><br>" & vbCrLf & _
			"	Atenciosamente,<br>" & vbCrLf & _
			"	Ger�ncia de Coordena��o da Vigil�ncia Continuada - GCVC<br>" & vbCrLf & _
			"	  </font></td>" & vbCrLf & _
			"	</tr>" & vbCrLf & _
			"  </table>" & vbCrLf & _
			"</html>" & vbCrLf

	Dim mailFrom : mailFrom = "gcvc" & rbac & "@anac.gov.br"
	
	Dim mailTo : mailTo = "sergio.dellamora@anac.gov.br; fabiano.nascimento@anac.gov.br; robson.ribeiro@anac.gov.br"
	querySQL =	"SELECT DISTINCT Pes.PES_LOGIN, TabSDiv.SDIV_SIGLA" & _
				"  FROM (Pessoal AS Pes LEFT JOIN Permissoes AS Per ON Pes.PES_CODI = Per.PES_CODI)" & _
				"       INNER JOIN Tab_Subdivisao AS TabSDiv ON Pes.SDIV_CODI = TabSDiv.SDIV_CODI" & _
				"  WHERE ( TabSDiv.SDIV_SIGLA='GTAR-DF' Or TabSDiv.SDIV_SIGLA='GTAR-RJ' Or TabSDiv.SDIV_SIGLA='GTAR-SP' )" & _
				"        And Per.PER_AREA = '" & rbac & "_LDR'"
	Set rsDiv = oDbFDH.getRecSetRd(querySQL)
	If Not rsDiv.Eof Then
		Do While Not rsDiv.Eof
			If InStr(mailTo,Trim(rsDiv("PES_LOGIN"))) <= 0 Then
				mailTo = mailTo & "; " & Trim(rsDiv("PES_LOGIN")) & "@anac.gov.br"
			End If
			rsDiv.MoveNext 
		Loop
	End If

	Select Case rbac
		Case "121"
			If InStr(mailTo,"ricardo.carone") <= 0 Then
				mailTo = mailTo & "; ricardo.carone@anac.gov.br"
			End If
			If InStr(mailTo,"vitor.nascimento") <= 0 Then
				mailTo = mailTo & "; vitor.nascimento@anac.gov.br"
			End If
		Case "135"
			If InStr(mailTo,"ricardo.carone") <= 0 Then
				mailTo = mailTo & "; ricardo.carone@anac.gov.br"
			End If
			If InStr(mailTo,"vitor.nascimento") <= 0 Then
				mailTo = mailTo & "; vitor.nascimento@anac.gov.br"
			End If
			If InStr(mailTo,"rodrigo.torres") <= 0 Then
				mailTo = mailTo & "; rodrigo.torres@anac.gov.br"
			End If
		Case "145"
			If InStr(mailTo,"wenderson.pires") <= 0 Then
				mailTo = mailTo & "; wenderson.pires@anac.gov.br"
			End If
			If InStr(mailTo,"marcos.aduar") <= 0 Then
				mailTo = mailTo & "; marcos.aduar@anac.gov.br"
			End If
			If InStr(mailTo,"stenio.neves") <= 0 Then
				mailTo = mailTo & "; stenio.neves@anac.gov.br"
			End If
	End Select
	
	Dim mailCc : mailCc = "gcvc" & rbac & "@anac.gov.br"
	mailCc = mailCc & "; roberto.honorato@anac.gov.br; helio.tarquinio@anac.gov.br; henri.bigatti@anac.gov.br; eduardo.campos@anac.gov.br; LD.SAR.GTPA@anac.gov.br"
	Select Case rbac
		Case "145"
			If InStr(mailCc,"carlos.almeida") <= 0 Then
				mailCc = mailCc & "; carlos.almeida@anac.gov.br"
			End If
		Case "121"
			If InStr(mailCc,"alexandre.henriques") <= 0 Then
				mailCc = mailCc & "; alexandre.henriques@anac.gov.br"
			End If
		Case "135"
			If InStr(mailCc,"gustavo.carneiro") <= 0 Then
				mailCc = mailCc & "; gustavo.carneiro@anac.gov.br"
			End If
	End Select

	Dim mailBcc : mailBcc = ""

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

	Const CdoReferenceTypeName = 1
	Dim oMail : Set oMail = CreateObject("CDO.Message")
	Set oMail.Configuration = oMailConf
	oMail.MimeFormatted = True
	oMail.Subject	= mailSubject
	oMail.From		= mailFrom
	oMail.To		= mailTo
	oMail.Cc		= mailCc
	oMail.Bcc		= mailBcc
	'oMail.TextBody	= mailBody
	oMail.HTMLBody	= mailBody

	Dim objBP, fname
	Dim posi : posi = 1
	Dim pose
	Do

		' <html>Check this out: <img src=""cid:myimage.gif""></html>
		' <img src=""cid:img_2015627_1SAR145.png""></td>" & _

		posi = InStr(posi,mailBody,"cid:img_")
			
		If posi > 0 Then

			pose = InStr(posi,mailBody,".png")

			posi = posi + 4

			If pose > posi And (pose-posi) < 60 Then

				fname = Mid(mailBody,posi,pose-posi+4)

				' Here's the good part, thanks to some little-known members.
				' This is a BodyPart object, which represents a new part of the multipart MIME-formatted message.
				' Note you can provide an image of ANY name as the source, and the second parameter essentially
				' renames it to anything you want.  Great for giving sensible names to dynamically-generated images.
				Set objBP = oMail.AddRelatedBodyPart(Server.MapPath("/Public/" & fname), fname, CdoReferenceTypeName)

				' Now assign a MIME Content ID to the image body part.
				' This is the key that was so hard to find, which makes it 
				' work in mail readers like Yahoo webmail & others that don't
				' recognise the default way Microsoft adds it's part id's,
				' leading to "broken" images in those readers.  Note the
				' < and > surrounding the arbitrary id string.  This is what
				' lets you have SRC="cid:myimage.gif" in the IMG tag.
				objBP.Fields.Item("urn:schemas:mailheader:Content-ID") = "<" & fname & ">"
				objBP.Fields.Update

			Else

				Exit Do

			End If

		End If

	Loop While posi > 0 And pose > 0

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
		closeWin()
	End If
	On Error GoTo 0

	' Mail is sent - tidy up and delete the AspEmail message object
	Set oMail = Nothing
	Set oMailConf = Nothing

	alert("Email enviado com sucesso!")

	Application.Lock
	Application("AIRStats" & rbac & "TStamp") = Now()
	Application.UnLock

	closeWin()

End If
%>

