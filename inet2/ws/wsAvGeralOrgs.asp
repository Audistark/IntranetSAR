<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<% Option Explicit %>
<% Response.CodePage = 1252 %>
<% Response.Buffer  = False %>
<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!-- #include virtual = "/inet2/class/cAAA.asp" -->
<!-- #include virtual = "/inet2/lib/libFuncDiv.asp" -->
<%

' .Pattern = "[0-9][0-9][0-9][0-9]-[0-9][0-9]"
Public Function ReCHE(strData)
    Dim RE, REMatches
    On Error Resume Next ' para não dar erro na linha
    Set RE = CreateObject("vbscript.regexp")
    With RE
        .MultiLine = False
        .Global = False
        .IgnoreCase = True
        .Pattern = "[0-9][0-9][0-9][0-9]-[0-9][0-9]"
    End With
    ' 0501-02/ANAC
    Set REMatches = RE.Execute(strData)
    If REMatches.Count <> 0 Then
        ReCHE = REMatches(0)
    Else
        ReCHE = ""
    End If
End Function

Dim Sec : Sec = Request.QueryString( "Sec" )
Dim StatCodi
Select Case Sec
	Case "145"
		StatCodi = "02"
	Case "121"
		StatCodi = "09"
	Case "135"
		StatCodi = "16"
	Case Else
		Response.Status = "400 Bad Request"
		Response.Write "Argumentos inválidos."
		Response.End
End Select

' web service ?Srv=web.service
Dim WebSrv : WebSrv = Request.QueryString( "Srv" )
If WebSrv <> "web.service" Then
	Response.Status = "400 Bad Request"
	Response.Write "Argumentos inválidos."
	Response.End
End If

Dim Row
Dim reset : reset = False

' command ?Command=fetch (/search)
Dim Command : Command = Request.QueryString( "Command" )
If Command = "fetch" Then
	Row = CLng("0" & Request.QueryString( "Row" ))
	If Row <= 0 Then
		Response.Status = "400 Bad Request"
		Response.Write "Argumentos inválidos."
		Response.End
	End If
ElseIf Command = "search" Then
	Row = 1 ' para forçar cálculo se estiver vazio
ElseIf Command = "code" Then
	Row = 1 ' para forçar cálculo se estiver vazio
ElseIf Command = "reset" Then
	reset = True
	Row = 1
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
Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
If oDbFDH.ErrorNumber < 0 then
	oDbFDH.Print()
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
  <title>AvGeralOrgsWebService</title>
</head>
<body>
	<table border="1">
<%

Dim Gtar, OrgPCodi, BSecCodi, OrgPNome, OrgNAbrev, Address, Complemento, City, State, Country
Dim Email, CHE, CHEStatus, RCA, StdPadrao, BaseStatus, UltAudit, RtNome, Adm, Tipo, Agreement
Dim querySQL, rsDiv, rsDiv2, res
Dim bCalculate : bCalculate = False

If Row = 1 Then
	If reset = True Then
		bCalculate = True
	Else
		' Verifica se tem que calcular mesmo
		querySQL = "SELECT TSTAMP FROM TMP_" & Sec & "ORGS WHERE ITEM = " & Row
		Set rsDiv = oDbFDH.getRecSetRd(querySQL)
		If rsDiv Is Nothing then
			Call PrintErr(-3, "Internal Error")
		End If
		If Not rsDiv.Eof Then
			' timestamp
			Dim tStamp : tStamp = rsDiv( "TSTAMP" )
			Dim minLast : minLast = DateDiff("n", tStamp, Now())
			If minLast < 120 Then ' não recalcula!!!
				bCalculate = False
			Else
				bCalculate = True
			End If
		Else
			bCalculate = True
		End If
	End If
End If

If bCalculate Then ' recalculate

	querySQL = "DELETE * FROM TMP_" & Sec & "ORGS"
	ret = oDbFDH.Execute( querySQL )

	querySQL = "  SELECT DISTINCT OP.ORGP_CODI, OP.ORGP_SIGLA, OP.ORGP_NOME, O.ORG_CODI, O.ORG_NABREV, " & _
			   "         O.ORG_ENDER, O.ORG_COMPL, O.ORG_CIDADE, O.ORG_ESTADO, O.ORG_CEP, O.ORG_FONEAREA, " & _
			   "         O.ORG_FONE, O.ORG_FAX, O.ORG_SITE, O.ORG_EMAIL, Pa.PAIS_NOME, CHE.CHE_CODI, " & _
			   "         CHE.CHE_DATA, CHE.CHE_STATUS, CHE.CHE_VALID, CHE.CHE_DTEXT, CHE.CHE_RCA, B.B" & Sec & "_CODI, " & _
			   "         B.B" & Sec & "_ADM, B.B" & Sec & "_TIPO, B.B" & Sec & "_TAMANHO, B.PES_CODI,  " & _
			   "         TP.STD_PADRAO,B.B" & Sec & "_STATUS, B.B" & Sec & "_DTSTAT, " & _
			   "         B.B" & Sec & "_ULTAUDIT, Pes.PES_NGUERRA, Ger.SDIV_SIGLA, B.B" & Sec & "_CNPJ, " & _
			   "         RT.RESP_TIPO, RT.RESP_DATA, OC.CONT_NOME, OC.CONT_EMAIL, OC.CONT_CPF " & _
			   "    FROM ( ( Paises AS Pa INNER JOIN ( ( ( ( ( ( ( ( Organizacao AS O " & _
			   "         LEFT JOIN A" & Sec & "_Bases AS B ON O.ORG_CODI = B.ORG_CODI ) " & _
			   "         LEFT JOIN A" & Sec & "_CHE AS CHE ON B.CHE_CODI = CHE.CHE_CODI ) " & _
			   "         INNER JOIN OrganizacaoP AS OP ON O.ORGP_CODI = OP.ORGP_CODI ) " & _
			   "         INNER JOIN A" & Sec & "_Empr AS E ON OP.ORGP_CODI = E.ORGP_CODI ) " & _
			   "         LEFT JOIN Pessoal AS Pes ON B.PES_CODI = Pes.PES_CODI ) " & _
			   "         LEFT JOIN Tab_Subdivisao AS Ger ON Pes.SDIV_CODI = Ger.SDIV_CODI ) " & _
			   "         LEFT JOIN ( A" & Sec & "_BasePadroes AS BP LEFT JOIN A" & Sec & "_TabPadroes AS TP " & _
			   "         ON BP.STD_CODI = TP.STD_CODI ) ON B.B" & Sec & "_CODI = BP.B" & Sec & "_CODI ) " & _
			   "         INNER JOIN OrgTipo AS OT ON O.ORG_CODI = OT.ORG_CODI ) " & _
			   "         ON Pa.PAIS_CODI = O.PAIS_CODI ) LEFT JOIN A" & Sec & "_RespTec AS RT " & _
			   "         ON B.CHE_CODI = RT.CHE_CODI ) LEFT JOIN Org_Contatos AS OC " & _
			   "         ON RT.CONT_CODI = OC.CONT_CODI " & _
			   "   WHERE OT.STAT_CODI = '" & StatCodi & "' " & _
			   "ORDER BY Ger.SDIV_SIGLA, OP.ORGP_SIGLA, O.ORG_NABREV, TP.STD_PADRAO, RT.RESP_TIPO, OC.CONT_NOME"

	Dim i : i = 1

	Set rsDiv = oDbFDH.getRecSetRd(querySQL)
	If rsDiv Is Nothing then
		Call PrintErr(-3, "Internal Error")
	End If
	If rsDiv.Eof Then
		Call PrintErr(-2, "Table Empty")
	End If

	Dim bSaveValues

	Do While True

		Gtar = rsDiv( "SDIV_SIGLA" )
		OrgPCodi = rsDiv( "ORGP_CODI" )
		BSecCodi = rsDiv( "B" & Sec & "_CODI" )
		OrgPNome = rsDiv( "ORGP_NOME" )
		If IsNull(OrgPNome) Then OrgPNome = ""
		OrgPNome = Trim(OrgPNome)
		OrgNAbrev = rsDiv( "ORG_NABREV" )
		If IsNull(OrgNAbrev) Then OrgNAbrev = ""
		OrgNAbrev = Trim(OrgNAbrev)
		Address = rsDiv( "ORG_ENDER" )
		If IsNull(Address) Then Address = ""
		Complemento = rsDiv( "ORG_COMPL" )
		If IsNull(Complemento) Then Complemento = ""
		If Complemento <> "" Then Address = Address & " " & Complemento
		City = rsDiv( "ORG_CIDADE" )
		If IsNull(City) Then City = ""
		State = rsDiv( "ORG_ESTADO" )
		If IsNull(State) Then State = ""
		Country = rsDiv( "PAIS_NOME" )
		If IsNull(Country) Then Country = ""
		Adm = rsDiv( "B" & Sec & "_ADM" )
		If IsNull(Adm) Then Adm = ""
		Tipo = rsDiv( "B" & Sec & "_TIPO" )
		If IsNull(Tipo) Then Tipo = ""
		Email = rsDiv( "ORG_EMAIL" )
		If IsNull(Email) Then Email = ""
        CHE = rsDiv( "CHE_CODI" )
		If IsNull(CHE) Then CHE = ""
        CHEStatus = rsDiv( "CHE_STATUS" )
		RCA = rsDiv("CHE_RCA")
		If IsNull(RCA) Then RCA = ""
		StdPadrao = rsDiv( "STD_PADRAO" )
        BaseStatus = rsDiv( "B" & Sec & "_STATUS" )
		UltAudit = rsDiv( "B" & Sec & "_ULTAUDIT" )
		If IsNull(UltAudit) Then UltAudit = ""
		RtNome = rsDiv( "CONT_NOME" )
		If IsNull(RtNome) Then RtNome = ""
		Agreement = ""

		Do While True

			rsDiv.MoveNext

			bSaveValues = False

			If rsDiv.Eof Then

				bSaveValues = True
			
			ElseIf OrgPCodi <> rsDiv( "ORGP_CODI" ) Or BSecCodi <> rsDiv( "B" & Sec & "_CODI" ) Then

				bSaveValues = True

			Else

				' StdPadrao
				res = Trim(rsDiv( "STD_PADRAO" ))
				If res <> "" Then
					If StdPadrao = "" Then
						StdPadrao = res
					ElseIf InStr(StdPadrao, res) = 0 Then
						StdPadrao = StdPadrao & ", " & res
					End If
				End If

				' RtNome
				res = Trim(rsDiv( "CONT_NOME" ))
				If IsNull(res) Then res = ""
				If res <> "" Then
					If Sec <> "145" Or rsDiv( "RESP_TIPO" ) = "R" Then
						If InStr(RtNome, res) = 0 Then
							If RtNome <> "" Then res = ", " & res
							RtNome = RtNome & res
						End If
					End If
				End If

			End If

			If Sec = "145" Then
				querySQL = "SELECT ORGP_CODI FROM A" & Sec & "_CertifEstrg " & _
							"   WHERE B" & Sec & "_CODI = '" & BSecCodi & "'"
				Set rsDiv2 = oDbFDH.getRecSetRd(querySQL)
				If rsDiv2 Is Nothing then
					Call PrintErr(-4, "Internal Error")
				End If
				While Not rsDiv2.eof
					If rsDiv2( "ORGP_CODI" ) = "0296" Then ' TCCA
						If Agreement <> "" Then Agreement = Agreement & ";" 
						Agreement = Agreement & "TCCA"
					ElseIf rsDiv2( "ORGP_CODI" ) = "0291" Then ' FAA
						If Agreement <> "" Then Agreement = Agreement & ";" 
						Agreement = Agreement & "FAA"
					ElseIf rsDiv2( "ORGP_CODI" ) = "0091" Then ' EASA
						If Agreement <> "" Then Agreement = Agreement & ";" 
						Agreement = Agreement & "EASA"
					End If
					rsDiv2.MoveNext
			   Wend
			   rsDiv2.Close()
			End If

			If bSaveValues = True Then

				On Error Resume Next
				querySQL =  "INSERT INTO TMP_" & Sec & "ORGS (ITEM, GTAR, ORGP_CODI, BSEC_CODI, RAZAOSOCIAL, NOMEABREV, " & _
							"ENDERECO, CIDADE, UF, PAIS, ADM, TIPO, EMAIL, CHE, STATUSCHE, RCA, AGREEMENTS, PADROES, STATUSBASE, ULTAUDIT, RTNOME, TSTAMP) " & _
							"VALUES (" & i & ", '" & Gtar & "', '" & OrgPCodi & "', '" & BSecCodi & "', '" & Replace(OrgPNome,"'","''") & "', '" & Replace(OrgNAbrev,"'","''") & "', '" & _
							Replace(Address,"'","''") & "', '" & Replace(City,"'","''") & "', '" & State & "', '" & Replace(Country,"'","''") & "', '" & Adm & _
							"', '" & Tipo  & "', '" & Email & "', '" & CHE & "', '" & CHEStatus & "', '" & RCA & "', '" & _
							Agreement & "', '" & StdPadrao & "', '" & BaseStatus & "', '" & UltAudit & "', '" & Replace(RtNome,"'","''") & "', '" & Now() & "');"
				ret = oDbFDH.Execute( querySQL )
				If Err.Number = 0 Then
					i = i + 1
				Else
					' Erase
					querySQL = "DELETE * FROM TMP_" & Sec & "ORGS"
					ret = oDbFDH.Execute( querySQL )
					Call PrintErr(-Err.Number, Err.Description)
				End If
				On Error GoTo 0

				Exit Do

			End If

		Loop

		If rsDiv.Eof Then Exit Do

	Loop

End If
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'------------------------------------------------------
'
'	Get org by COM and Name
'
If Command = "search" Then

	Dim Com : Com	= Request.QueryString( "COM" )
	Dim Base : Base	= Trim(Request.QueryString( "Base" ))
	If Base = "" Then
		Call PrintErr(0, "Not Found")
	End If

	' Le todos em busca do cara
	querySQL = "SELECT * FROM TMP_" & Sec & "ORGS"
	Set rsDiv = oDbFDH.getRecSetRd(querySQL)
	If rsDiv Is Nothing then
		Call PrintErr(-3, "Internal Error")
	End If

	If rsDiv.Eof Then
		Call PrintErr(0, "Not Found")
	End If

	Com = ReCHE(Com)

	Do While True

		OrgNAbrev = rsDiv( "NOMEABREV" )
		CHE = ReCHE(rsDiv( "CHE" ))

		' che e base iguais
		If Base = OrgNAbrev Then
			If Com = "" Or CHE = "" Or Com = CHE Then
				Exit Do
			End If
		End If

		rsDiv.MoveNext

		If rsDiv.Eof Then
			Call PrintErr(0, "Not Found")
		End If

	Loop

'------------------------------------------------------
'
'	Fetch orgs
'
ElseIf Command = "fetch" Or Command = "reset" Then

	querySQL = "SELECT * FROM TMP_" & Sec & "ORGS WHERE ITEM = " & Row & " ORDER BY ITEM"
	Set rsDiv = oDbFDH.getRecSetRd(querySQL)
	If rsDiv Is Nothing then
		Call PrintErr(-3, "Internal Error")
	End If
	If rsDiv.Eof Then
		Call PrintErr(0, "Not Found")
	End If

'------------------------------------------------------
'
'	Get Code reg
'
ElseIf Command = "code" Then

	Dim Code : Code = Request.QueryString( "Code" )
	If Code = "" Then
		Call PrintErr(-1, "Argumentos inválidos.")
	End If

	' Le em busca do cara
	If Len(Code) < 10 Then Code = Right("000000000000",10-Len(Code)) & Code
	querySQL = "SELECT * FROM TMP_" & Sec & "ORGS WHERE BSEC_CODI = '" & Code & "'"
	Set rsDiv = oDbFDH.getRecSetRd(querySQL)
	If rsDiv Is Nothing then
		Call PrintErr(-3, "Internal Error")
	End If

	If rsDiv.Eof Then
		Call PrintErr(0, "Not Found")
	End If

Else

	Call PrintErr(-1, "Argumentos inválidos.")

End If

%>
		<tr>
			<td>RET</td>
			<td>1</td>
		</tr>
		<tr>
			<td>ITEM</td>
			<td><%=rsDiv("ITEM") %></td>
		</tr>
		<tr>
			<td>GTAR</td>
			<td><%=rsDiv("GTAR") %></td>
		</tr>
		<tr>
			<td>ORGP_CODI</td>
			<td><%=rsDiv("ORGP_CODI") %></td>
		</tr>
		<tr>
			<td>BSEC_CODI</td>
			<td><%=rsDiv("BSEC_CODI") %></td>
		</tr>
		<tr>
			<td>RAZAOSOCIAL</td>
			<td><%=rsDiv("RAZAOSOCIAL") %></td>
		</tr>
		<tr>
			<td>NOMEABREV</td>
			<td><%=rsDiv("NOMEABREV") %></td>
		</tr>
		<tr>
			<td>ENDERECO</td>
			<td><%=rsDiv("ENDERECO") %></td>
		</tr>
		<tr>
			<td>CIDADE</td>
			<td><%=rsDiv("CIDADE") %></td>
		</tr>
		<tr>
			<td>UF</td>
			<td><%=rsDiv("UF") %></td>
		</tr>
		<tr>
			<td>PAIS</td>
			<td><%=rsDiv("PAIS") %></td>
		</tr>
		<tr>
			<td>ADM</td>
			<td><%=rsDiv("ADM") %></td>
		</tr>
		<tr>
			<td>TIPO</td>
			<td><%=rsDiv("TIPO") %></td>
		</tr>
		<tr>
			<td>EMAIL</td>
			<td><%=rsDiv("EMAIL") %></td>
		</tr>
		<tr>
			<td>CHE</td>
			<td><%=rsDiv("CHE") %></td>
		</tr>
		<tr>
			<td>STATUSCHE</td>
			<td><%=rsDiv("STATUSCHE") %></td>
		</tr>
		<tr>
			<td>RCA</td>
			<td><%=rsDiv("RCA") %></td>
		</tr>
		<tr>
			<td>AGREEMENTS</td>
			<td><%=rsDiv("AGREEMENTS") %></td>
		</tr>
		<tr>
			<td>PADROES</td>
			<td><%=rsDiv("PADROES") %></td>
		</tr>
		<tr>
			<td>STATUSBASE</td>
			<td><%=rsDiv("STATUSBASE") %></td>
		</tr>
		<tr>
			<td>ULTAUDIT</td>
			<td><%=rsDiv("ULTAUDIT") %></td>
		</tr>
		<tr>
			<td>RTNOME</td>
			<td><%=rsDiv("RTNOME") %></td>
		</tr>
		<tr>
			<td>TIMESTAMP</td>
			<td><%=rsDiv( "TSTAMP" ) %></td>
		</tr>
	</table>
</body>
</html>
<%

rsDiv.Close
oDbFDH.Close

 %>
