<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<% Option Explicit %>
<% Response.CodePage = 1252 %>
<% Response.Buffer  = True %>
<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!-- #include virtual = "/inet2/class/cAAA.asp" -->
<!-- #include virtual = "/inet2/lib/libFuncDiv.asp" -->
<%

' web service ?Srv=web.service
Dim WebSrv : WebSrv = Request.QueryString( "Srv" )
If WebSrv <> "web.service" Then
	Response.Status = "400 Bad Request"
	Response.Write "Argumentos inválidos."
	Response.End
End If

Dim Value : Value = Request.QueryString( "Value" )

' command ?Command=get
Dim Command : Command = Request.QueryString( "Command" )
If Command = "get" Then
	If Value = "" Or CInt(Value) <= 0 Then
		Response.Status = "400 Bad Request"
		Response.Write "Argumentos inválidos."
		Response.End
	End If
Else
	Response.Status = "400 Bad Request"
	Response.Write "Argumentos inválidos."
	Response.End
End If

'-----------------------------------------------------
'' Key
'Dim SecKey : SecKey = Request.QueryString("SEC_KEY")
'If Len(SecKey) > 5 Then
'	Response.Status = "400 Bad Request"
'	Response.End
'End If
'If Abs( CLng(SecKey) - Timer() ) > 60 Then
'	Response.Status = "400 Bad Request"
'	Response.End
'End If
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Init AAA - Authentication, Authorization and Accounting
Dim oAAA : Set oAAA = new cAAA
Dim ret : ret = oAAA.WinAuthenticate(True)
If ret < 0 Then
	Response.Status = "403 Forbidden"
	Response.End
End If

' Database access
Dim oDbPostgreSQL : Set oDbPostgreSQL = (new cDBAccess)( "POSTGRESQL" )
If oDbPostgreSQL.ErrorNumber < 0 then
	oDbPostgreSQL.Print()
End If

Sub PrintErr( err, descr )
%>
		<tr>
			<td>RET</td>
			<td><%=err %></td>
			<td><%=descr %></td>
		</tr>
	</table>
</body>
</html>
<%
	Response.End
End Sub

%>
<!DOCTYPE html>
<html>
<head>
  <title>AvGeralIN81WebService</title>
</head>
<body>
	<table border="1">
<%

Dim querySQL, rsDiv, res
Dim id, status, dt_init

'------------------------------------------------------
'
'	Get inspection
'
If Command = "get" Then

	' Le todos em busca do cara
	querySQL = "SELECT" & _
			   "  fiscalizacao.id_fiscalizacao AS id," & _
			   "  fiscalizacao.sn_encerrada AS status," & _
			   "  fiscalizacao.dt_inicio_fiscalizacao AS dt_init" & _
			   " FROM" & _
			   "  public.fiscalizacao" & _
			   " WHERE" & _
			   "  fiscalizacao.id_fiscalizacao = " & Value
	Set rsDiv = oDbPostgreSQL.getRecSetRd(querySQL)
	If rsDiv Is Nothing then
		Call PrintErr(-3, "Internal Error")
	End If

	If rsDiv.Eof Then
		Call PrintErr(0, "Not Found")
	End If

	id = rsDiv( "id" )
	status = rsDiv( "status" )
	dt_init = rsDiv( "dt_init" )

Else

	Call PrintErr(-1, "Argumentos inválidos.")

End If

%>
		<tr>
			<td>RET</td>
			<td>1</td>
		</tr>
		<tr>
			<td>ID</td>
			<td><%=id %></td>
		</tr>
		<tr>
			<td>STATUS</td>
			<td><%=status %></td>
		</tr>
		<tr>
			<td>DATE_INIT</td>
			<td><%=dt_init %></td>
		</tr>
	</table>
</body>
</html>
<%

rsDiv.Close
oDbPostgreSQL.Close

 %>
