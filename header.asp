<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.CodePage = 1252 %> 
<!-- #include virtual = "/inet2/class/cCtrlErr.asp" -->
<!-- #include virtual = "/inet2/class/cLog.asp" -->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!-- #include virtual = "/inet2/class/cAAA.asp" -->

<%
' Init AAA - Authentication, Authorization and Accounting
Dim oAAA : Set oAAA = new cAAA
Dim nick : nick = ""
Dim sdiv : sdiv = ""
If oAAA.IsUsrAuthenticated() Then
	nick = oAAA.AuthentUserNick
	sdiv = oAAA.AuthentUserSDiv
ElseIf oAAA.IsWinAuthenticated() Then
	nick = oAAA.AuthentWinUserNick
	sdiv = oAAA.AuthentWinUserSDiv
Else
	Dim ret : ret = oAAA.WinAuthenticate(False)
	If ret >= 0 Then
		nick = oAAA.AuthentWinUserNick
		sdiv = oAAA.AuthentWinUserSDiv
	End If
End If
%>

<!DOCTYPE html>
<html>
<head>
  <style type="text/css">
	a:link {
	  border-style: none;
	  color: red;
	  text-decoration: none;
	}
	a:visited {
	  color: red;
	  text-decoration: none;
	}
	a:hover {
	  color: red;
	  text-decoration: none;
	}
	a:active {
	  color: red;
	  text-decoration: none;
	}
	body { 
		margin:0px;
		margin-top:1px;
		padding:0px;
	}
	.SarFont {
		font-family: Calibri;
		font-size: 18px;
		color: black;
		font-style: normal;
	}
  </style>
  <title></title>
</head>
<body>
<table border="0" cellspacing="0" width="100%">
  <tbody>
    <tr style="border:0px;">
      <td style="line-height:0px"><a href="/" target="_top"><img alt="Intranet SAR"
 src="inet2/img/Anac-Title.png" border="0" height="60" hspace="18"
 vspace="2" width="314"></a></td>
      <td class="SarFont" align="right"><%=nick %><br><%=sdiv %></td>
      <td width="16"></td>
      <td style="line-height:0px" align="center" width="64"><a href="Pessoal/Pessoal.asp" target="mainFrame"><img alt="User"
 src="inet2/img/User-Title.png" border="0" height="60" width="42"></a></td>
      <td style="line-height:0px" align="center" width="64"><a href="linkutil.asp" target="mainFrame"><img alt=""
 src="inet2/img/Links-Title.png" border="0" height="60" width="42"></a></td>
      <td style="line-height:0px" align="center" width="64"><a href="Senha.asp?Arq=_Mnt/Manutencao.asp" target="mainFrame"><img alt=""
 src="inet2/img/Menu-Title.png" border="0" height="60" width="42"></a></td>
    </tr>
    <tr style="border:0px; font-size:0px; line-height:0px">
      <td height="1" colspan="6" valign="top"><img
 src="inet2/img/pixel.png" height="1" width="100%"></td>
    </tr>
  </tbody>
</table>
<iframe marginwidth="0" marginheight="0" id="iFrameStatistics" src="/inet2/wrAIRStatistics.asp" frameborder="0"></iframe>
</body>
</html>
