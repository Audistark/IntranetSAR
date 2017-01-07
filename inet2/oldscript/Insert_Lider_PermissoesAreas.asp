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

' DB
Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
If oDbFDH.ErrorNumber < 0 then
	oDbFDH.Print()
End If


Dim querySQL

'------------------------------------------------------------------
'
' Deleta as Permissões 999_LDR utilizada pelo pessoal administrativo
'
' tenho que limpar o lixo anterior :(
querySQL = "DELETE * FROM PermissoesAreas WHERE PER_AREA = '145_LDR'"
oDbFDH.Execute( querySQL )
querySQL = "DELETE * FROM PermissoesAreas WHERE PER_AREA = '135_LDR'"
oDbFDH.Execute( querySQL )
querySQL = "DELETE * FROM PermissoesAreas WHERE PER_AREA = '121_LDR'"
oDbFDH.Execute( querySQL )

'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'------------------------------------------------------------------
'
' Insere Permissão 999_LDR utilizada pelo pessoal administrativo
'
' http://sar/inet2/script/Insert_Lider_PermissoesAreas.asp
'
querySQL = "INSERT INTO PermissoesAreas (PER_AREA, SDIV_CODI, SEC_CODI, AREA, PER_DESC, PER_MST)" & _
			"    VALUES ( '145_LDR', '18', '031', 'AIR-145', 'Usuário Líder do Grupo', 'N')"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO PermissoesAreas (PER_AREA, SDIV_CODI, SEC_CODI, AREA, PER_DESC, PER_MST)" & _
			"    VALUES ( '135_LDR', '17', '029', 'AIR-135', 'Usuário Líder do Grupo', 'N')"
oDbFDH.Execute( querySQL )
querySQL = "INSERT INTO PermissoesAreas (PER_AREA, SDIV_CODI, SEC_CODI, AREA, PER_DESC, PER_MST)" & _
			"    VALUES ( '121_LDR', '17', '028', 'AIR-121', 'Usuário Líder do Grupo', 'N')"
oDbFDH.Execute( querySQL )

Response.Status = "200 OK"
Response.Write "Data was inserted sucessfully in Database!"

'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Response.End

%>