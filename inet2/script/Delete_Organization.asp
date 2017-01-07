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

Dim querySQL, rsDiv, rsDiv2
Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
If oDbFDH.ErrorNumber < 0 then
	oDbFDH.Print()
End If

Response.Buffer = true
Response.Status = "200 OK"


'''''''''''''''''''' Search duplicate ones ''''''''''''''''''''''''''''''''
'
Response.Write "Searching duplicated Bases...<br>"
Response.Flush

Dim rbac(3) : rbac(0) = "121": rbac(1) = "135": rbac(2) = "145"
Dim i
For i=0 to 2

	Response.Write "RBAC " & rbac(i) & "<br>"

	querySQL = "SELECT A" & rbac(i) & "_Bases.ORG_CODI, Count(*) FROM A" & rbac(i) & "_Bases GROUP BY A" & rbac(i) & "_Bases.ORG_CODI HAVING Count(*)>1"
	Set rsDiv = oDbFDH.getRecSetRd(querySQL)
	If rsDiv Is Nothing Or Err.Number <> 0 Then
		oDbFDH.Print()
	End If
	On Error Resume Next
	If rsDiv.Eof Then
		Response.Write "No duplicated records found on A" & rbac(i) & "_Bases.<br>"
	Else
		Do While not rsDiv.Eof
			OrgCodi = rsDiv("ORG_CODI")
			Response.Write "Duplicated ORG_CODI : '" & OrgCodi & "'<br>"

			querySQL =  "SELECT B.B" & rbac(i) & "_CODI, B.CHE_CODI, B.B" & rbac(i) & "_TIPO, B.B" & rbac(i) & "_TAMANHO, O.ORG_NABREV " & _
						" FROM A" & rbac(i) & "_Bases AS B INNER JOIN Organizacao AS O ON B.ORG_CODI = O.ORG_CODI " & _
						" WHERE B.ORG_CODI='" & OrgCodi & "'"
			Set rsDiv2 = oDbFDH.getRecSetRd(querySQL)
			If rsDiv2 Is Nothing Or Err.Number <> 0 Then
				oDbFDH.Print()
			End If
			If rsDiv2.Eof = True Then
				Response.Write "Record not found on A" & rbac(i) & "_Bases.<br>"
				On Error GoTo 0
				rsDiv.Close()
				Response.End
			End If
			Do While not rsDiv2.Eof
				Dim CheCodi : CheCodi = rsDiv2("CHE_CODI")
				Dim BTipo : BTipo = rsDiv2("B" & rbac(i) & "_TIPO")
				Dim BTam : BTam = rsDiv2("B" & rbac(i) & "_TAMANHO")
				Dim BCodi : BCodi = rsDiv2("B" & rbac(i) & "_CODI")
				Dim Name : Name = rsDiv2("ORG_NABREV")
				If (IsNull(CheCodi) Or CheCodi = "") And _
					(IsNull(BTipo) Or BTipo = "") And _
					 (IsNull(BTam) Or BTam = "") Then
					Response.Write "B" & rbac(i) & "_CODI: " & BCodi & "<br>"
					Response.Write "NAME : " & Name & " <br>"

					querySQL = "DELETE * FROM A" & rbac(i) & "_Bases WHERE ORG_CODI = '" & OrgCodi & "' AND B" & rbac(i) & "_CODI = '" & BCodi & "'"
					res = oDbFDH.Execute( querySQL )

					Response.Write "Cleaned duplicated B" & rbac(i) & "_CODI = '" & BCodi & "' from A" & rbac(i) & "_Bases<br>"
					Response.Flush

					Exit Do
				End If
				rsDiv2.MoveNext
			Loop

			On Error GoTo 0
			rsDiv2.Close()

			rsDiv.MoveNext

		Loop
	
	End If

	rsDiv.Close()

Next


'''''''''''''''''''' Delete Base com nome "XXXXXXXXXXXXXXXXXXXX" ''''''''''''''''''''''''''''''''
'
Response.Write "Searching deleted Bases...<br>"
Response.Flush

' Select "XXXXXXXXXXXXXXXXXXXX"
querySQL = "SELECT * FROM Organizacao WHERE ORG_NOME='XXXXXXXXXXXXXXXXXXXX' AND ORG_NABREV='XXXXXXXXXXXXXXXXXXXX'"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
If Err.Number <> 0 Then
	oDbFDH.Print()
End If
If rsDiv.Eof = True Then
	Response.Write "No deleted records found on Organizacao.<br>"
	On Error GoTo 0
	rsDiv.Close()
	Response.End
End If

Response.Write "Record found!<br>"
Dim OrgCodi
OrgCodi = rsDiv("ORG_CODI")
Response.Write "ORG_CODI = " & OrgCodi & "<br>"
On Error GoTo 0
rsDiv.Close()

' Get Org Type
Response.Write "Verify Organization Type.<br>"
querySQL = "SELECT * FROM OrgTipo WHERE ORG_CODI = '" & OrgCodi & "'"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
If Err.Number = 0 Then
	Dim StatCodi
	Do While not rsDiv.eof
		StatCodi = rsDiv("STAT_CODI")
		Response.Write "STAT_CODI = " & StatCodi & "<br>"
		'	STAT_CODI	STAT_DESCR
		'	02	Empresa de Manutenção RBAC 145
		'	09	Empresa Aérea RBAC 121
		'	16	Empresa de Táxi Aéreo RBAC 135
		If StatCodi <> "02" And StatCodi <> "09" And StatCodi <> "16" Then
			Response.Write "Detected invalid StatCodi in OrgType to perform deletion!<br>"
			On Error GoTo 0
			rsDiv.Close()
			Response.End
		End If
		rsDiv.MoveNext
	Loop
End If
On Error GoTo 0
rsDiv.Close()

' Processos
Response.Write "Verifying 121 Processes.<br>"
querySQL = "SELECT * FROM A121_Processos INNER JOIN A121_Bases ON A121_Processos.B121_CODI = A121_Bases.B121_CODI WHERE A121_Bases.ORG_CODI = '" & OrgCodi & "'"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
If Err.Number = 0 Then
	If Not rsDiv.eof Then
		Response.Write "Detected processes to 121 certification!<br>"
		On Error GoTo 0
		rsDiv.Close()
		Response.End
	End If
End If
On Error GoTo 0
rsDiv.Close()

Response.Write "Verifying 135 Processes.<br>"
querySQL = "SELECT * FROM A135_Processos INNER JOIN A135_Bases ON A135_Processos.B135_CODI = A135_Bases.B135_CODI WHERE A135_Bases.ORG_CODI = '" & OrgCodi & "'"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
If Err.Number = 0 Then
	If Not rsDiv.eof Then
		Response.Write "Detected processes to 135 certification!<br>"
		On Error GoTo 0
		rsDiv.Close()
		Response.End
	End If
End If
On Error GoTo 0
rsDiv.Close()

Response.Write "Verifying 145 Processes.<br>"
querySQL = "SELECT * FROM A145_Processos INNER JOIN A145_Bases ON A145_Processos.B145_CODI = A145_Bases.B145_CODI WHERE A145_Bases.ORG_CODI = '" & OrgCodi & "'"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv Is Nothing then
	oDbFDH.Print()
End If
On Error Resume Next
If Err.Number = 0 Then
	If Not rsDiv.eof Then
		Response.Write "Detected processes to 145 certification!<br>"
		On Error GoTo 0
		rsDiv.Close()
		Response.End
	End If
End If
On Error GoTo 0
rsDiv.Close()

querySQL = "DELETE * FROM A121_Bases WHERE ORG_CODI = '" & OrgCodi & "'"
res = oDbFDH.Execute( querySQL )
Response.Write "Cleaned ORG_CODI from A121_Bases<br>"

querySQL = "DELETE * FROM A135_Bases WHERE ORG_CODI = '" & OrgCodi & "'"
res = oDbFDH.Execute( querySQL )
Response.Write "Cleaned ORG_CODI from A135_Bases<br>"

querySQL = "DELETE * FROM A145_Bases WHERE ORG_CODI = '" & OrgCodi & "'"
res = oDbFDH.Execute( querySQL )
Response.Write "Cleaned ORG_CODI from A145_Bases<br>"

querySQL = "DELETE * FROM OrgTipo WHERE ORG_CODI = '" & OrgCodi & "'"
res = oDbFDH.Execute( querySQL )
Response.Write "Deleted ORG_CODI from OrgTipo<br>"

querySQL = "DELETE * FROM Organizacao WHERE ORG_CODI = '" & OrgCodi & "'"
res = oDbFDH.Execute( querySQL )
Response.Write "Deleted ORG_CODI from Organizacao<br>"

'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Response.End

%>