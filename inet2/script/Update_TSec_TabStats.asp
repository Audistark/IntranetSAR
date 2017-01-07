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

'TStatsGr_Id	TStatsGr_Name									TStatsGr_Goal	TStatsGr_ByTask
'	1			Inclusão de Capacidade na EO - Nacional				25				N
'	2			Supervisão Continuada - Auditorias - Nacional		 0				N
'	3			Certificação Inicial								25				N

querySQL = "UPDATE A145_TabStatsGroup SET TStatsGr_Goal = 0 WHERE TStatsGr_Id=3"
oDbFDH.Execute( querySQL )
Response.Write "The TABLE A145_TabStatsGroup was updated sucessfully in Database!<br>"

querySQL = "UPDATE A135_TabStatsGroup SET TStatsGr_Goal = 0 WHERE TStatsGr_Id=3"
oDbFDH.Execute( querySQL )
Response.Write "The TABLE A135_TabStatsGroup was updated sucessfully in Database!<br>"

querySQL = "UPDATE A121_TabStatsGroup SET TStatsGr_Goal = 0 WHERE TStatsGr_Id=3"
oDbFDH.Execute( querySQL )
Response.Write "The TABLE A121_TabStatsGroup was updated sucessfully in Database!<br>"

Response.End


'Response.Write "The TABLE A145_TabStatsGroup was updated sucessfully in Database!<br>"
'querySQL = "UPDATE A145_TabStatsGroup SET TStatsGr_Name='Supervisão Continuada - Auditorias - Nacional' WHERE TStatsGr_Id=2"
'oDbFDH.Execute( querySQL )


querySQL = "SELECT * FROM Paises WHERE PAIS_CODI='0117'"
Set rsDiv = oDbFDH.getRecSetRd(querySQL)
If rsDiv.Eof Then
	querySQL = "INSERT INTO Paises (PAIS_CODI, PAIS_NOME, PAIS_NAME, PAIS_FONEAREA, IDM_CODI, BRASIL)" & _
				"    VALUES ( '0117', 'Guiana Francesa', 'French Guiana', '594', '006', 'N')"
	oDbFDH.Execute( querySQL )
	Response.Write "Ok, inserted<br>"
Else
	Response.Write "None, record already exists<br>"
End If

querySQL = "UPDATE Tab_Subdivisao SET SDIV_EMAIL='gtar.sp@anac.gov.br', SDIV_PHONE='(11) 3636-8686' WHERE SDIV_CODI=24"
oDbFDH.Execute( querySQL )

querySQL = "UPDATE Tab_Subdivisao SET SDIV_EMAIL='gtar.rj@anac.gov.br', SDIV_PHONE='(21) 3501-5348' WHERE SDIV_CODI=24"
oDbFDH.Execute( querySQL )


'''''''''''''''''''' Table AIRStats ''''''''''''''''''''''''''''''''''
'
'On Error Resume Next
'

querySQL = "DELETE * FROM AIRStats WHERE AIRStats.Stats_DATE < #6/19/2015#"
oDbFDH.Execute( querySQL )

'
'querySQL = "DELETE * FROM AIRStats WHERE AIRStats.Stats_DATE < #6/7/2015# AND AIRStats.Stats_RBAC = '145' AND " & _
'			"AIRStats.Stats_CODI <> '001' AND AIRStats.Stats_CODI <> '006' AND AIRStats.Stats_CODI <> '015'"
'oDbFDH.Execute( querySQL )
'
'Response.Write "DELETE * FROM AIRStats WHERE AIRStats.Stats_DATE < #5/19/2015#<br>"
'
'
'On Error GoTo 0
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''' Novo campo Stats_TOPTEN em AIRStats '''''''''''''''''''''''''''''''''''
'
'querySQL = "SELECT * FROM AIRStats"
'Set rsDiv = oDbFDH.getRecSetRd(querySQL)
'If rsDiv Is Nothing then
'	oDbFDH.Print()
'End If
'On Error Resume Next
'col = rsDiv("Stats_TOPTEN")
'If Err.Number = 0 Then
'	Response.Write "The Column Stats_TOPTEN already exists in Database.<br>"
'Else
'	oDbFDH.Execute( "ALTER TABLE AIRStats ADD Stats_TOPTEN TEXT(255)" )
'	Response.Write "Column Stats_TOPTEN created sucessfully in Database!<br>"
'End If
'On Error GoTo 0
'rsDiv.Close()

Response.End
%>