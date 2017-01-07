<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<% Option Explicit %>
<% Response.CodePage = 1252 %>
<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!-- #include virtual = "/inet2/class/cAAA.asp" -->
<!-- #include virtual = "/inet2/lib/libFuncDiv.asp" -->

<!DOCTYPE html>
<html>
<head>
  <title>Intranet SAR - PortalRBAC145WebService</title>
	</head>
<body>

<%

' Init AAA - Authentication, Authorization and Accounting
Dim oAAA : Set oAAA = new cAAA
Dim ret : ret = oAAA.WinAuthenticate(True)
If ret < 0 Then
	Response.Status = "403 Forbidden"
	Response.End
End If

' Só MASTERs
If oAAA.AuthorWinMasterSec("145") <> True And _
	oAAA.AuthorWinMasterSec("135") <> True And _
	 oAAA.AuthorWinMasterSec("121") <> True Then
	Response.Status = "403 Forbidden"
	Response.End
End If

Dim Path : Path = Request.ServerVariables( "APPL_PHYSICAL_PATH" ) & "FDH\AvGeral\Tmp"

Dim oFSO : Set oFSO = Server.CreateObject( "Scripting.FileSystemObject" )

'' Tirar depois
On Error Resume Next
Dim f : Set f = oFSO.CreateFolder(Path)
On Error GoTo 0
Set f = Nothing

' Database access
Dim oDbFDH : Set oDbFDH = (new cDBAccess)( "FDH" )
If oDbFDH.ErrorNumber < 0 then
	oDbFDH.Print()
End If
Dim querySQL, rsDiv

Dim oUpload : Set oUpload = Server.CreateObject( "AspSmartUpLoad.SmartUpLoad" )
On Error Resume Next
oUpload.UpLoad
If Err.Number <> 0 Then
%>
<form action="UploadRCA.asp" method="post" enctype="multipart/form-data">
<table border=0>
  <tr>
	<td>Upload do arquivo em formato Excel contendo as informações sobre RCA realizados 
		no ano anterior.</td>
  </tr>
  <tr>
	<td>Nos dados devem estar todos os RCA do ano anterior, já que ao atualizar o 
		sistema apaga os dados anteriores.</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
  </tr>
  <tr>
	<td>Informar o arquivo excel (.xls) contendo os dados dos RCA realizados</td>
  </tr>
  <tr>
	<td><input type="file" size="120" name="ArquivoDigital"></td>
  </tr>
  <tr>
	<td><input type="submit" value="Submit"></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
  </tr>
  <tr>
	<td>&nbsp;</td>
  </tr>
  <tr>
	<td>Obs.: O arquivo deve conter as seguintes colunas, sendo que na primeira linha deve estar o cabeçalho contendo as lables:<br />
&nbsp;&nbsp;&nbsp;&nbsp; - 'DATA' - Data de realização do RCA;<br />
&nbsp;&nbsp;&nbsp;&nbsp; - 'MARCA' - Marcas da aeronave;<br />
&nbsp;&nbsp;&nbsp;&nbsp; - 'CERTIFICADO' - Certificado COM ou Certificado ETA;<br />
&nbsp;&nbsp;&nbsp;&nbsp; - 'RESULTADO' - Aeronavegável, Sistema de Amostragem, etc..<br />
		Não precisa ser só essas colunas nem estarem nesta ordem.</td>
  </tr>
</table>
</form>
<%
End If
On Error GoTo 0
If oUpload.Files.Count = 1 Then
	If Not oUpload.Files.Item(1).IsMissing Then
		Dim Filename : Filename  = oUpload.Files.Item( 1 ).FileName
		Dim Extension : Extension = oFSO.GetExtensionName( Filename )
		oUpload.Files.Item(1).SaveAs Path & "\" & "RCA." & Extension

		Dim ExcelFile : ExcelFile = Path & "\" & "RCA." & Extension

		Dim ExcelConnection : Set ExcelConnection = Server.createobject("ADODB.Connection")
		ExcelConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ExcelFile & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
		Dim strSQL : strSQL = "SELECT * FROM A1:J10000"
		Dim rsExcel : Set rsExcel = Server.CreateObject("ADODB.Recordset")
		rsExcel.Open strSQL, ExcelConnection

		Response.Write "<table border=""1""><thead><tr>"
		Dim Column, Field
		Dim colData : colData = -1
		Dim colMarca : colMarca = -1
		Dim colCert : colCert = -1
		Dim colResult : colResult = -1
		Dim i : i = 0
		Response.Write "<th>ITEM</th>"
		For Each Column In rsExcel.Fields
			If UCase(Trim(Column.Name)) = "DATA" Then
				colData = i
				Response.Write "<th>" & Column.Name & "</th>"
			ElseIf UCase(Trim(Column.Name)) = "MARCA" Then
				colMarca = i
				Response.Write "<th>" & Column.Name & "</th>"
			ElseIf UCase(Trim(Column.Name)) = "CERTIFICADO" Then
				colCert = i
				Response.Write "<th>" & Column.Name & "</th>"
			ElseIf UCase(Trim(Column.Name)) = "RESULTADO" Then
				colResult = i
				Response.Write "<th>" & Column.Name & "</th>"
			End If
			i = i + 1
		Next
		Response.Write "<th>OBSERVAÇÃO</th>"
		Response.Write "</tr></thead><tbody>"
		If colData < 0 Or colMarca < 0 Or colCert < 0 Or colResult < 0 Then
			Response.Write "Error: Invalid parameters. Not Found expected header data on uploaded file."
		End If
		Dim j : j = 1
		Dim iRes : iRes = 1
		Dim sObs : sObs = ""
		Dim nSuccess : nSuccess = 0
		Dim keyCert, rbac
		If Not rsExcel.Eof Then

			While Not rsExcel.Eof

				Response.Write "<tr>"
				Response.Write "<th>" & j & "</th>"

				j = j + 1
				i = 0
				iRes = 1

				For Each Field In rsExcel.Fields

					If i = colData And iRes > 0 Then
						If Field.value <> "" Then
							If Year(CDate(Field.value)) <> (Year(Date())-1) Then
								iRes = -1
								sObs = "<font color=BlueViolet>Data " & Field.value & " não considerada</font>"
							End If
						Else
							iRes = -2
							sObs = "<font color=Red>Data inválida</font>"
						End If

					ElseIf i = colCert And iRes > 0 Then
						' colCert: 0811-52/ANAC ou 2004-03-OCDD-03-01
						Dim pos, shortCert, anacCert, dacCert
						If IsNull(Field.value) Or Field.value = "" Then
							iRes = -3
							sObs = "<font color=Red>Certificado inválido</font>"
						ElseIf Len(Trim(Field.value)) < 13 Then ' COM
							Dim val : val = Trim(Field.value)
							pos = InStr((val),"/")
							If pos <= 0 Then
								shortCert = val
								anacCert = val & "/ANAC"
								dacCert = val & "/DAC"
							Else
								shortCert = Left(val,pos-1)
								anacCert = Left(val,pos-1) & "/ANAC"
								dacCert = Left(val,pos-1) & "/DAC"
							End If
							querySQL = "SELECT CHE.CHE_CODI, CHE.CHE_RCA, OP.ORGP_SIGLA" & _
										" FROM ( ( A145_CHE AS CHE INNER JOIN A145_Bases AS B ON CHE.CHE_CODI = B.CHE_CODI) " & _
										" INNER JOIN Organizacao AS O ON B.ORG_CODI = O.ORG_CODI) " & _
										" INNER JOIN OrganizacaoP AS OP ON O.ORGP_CODI = OP.ORGP_CODI " & _
										" WHERE CHE.CHE_CODI='" & shortCert & "' OR CHE.CHE_CODI='" & anacCert & "' OR CHE.CHE_CODI='" & dacCert & "'"
							Set rsDiv = oDbFDH.getRecSetRd(querySQL)
							If rsDiv.Eof Then
								iRes = 0
								sObs = "<font color=red>Certificado não identificado</font>"
							Else
								sObs = rsDiv("ORGP_SIGLA")
								keyCert = rsDiv("CHE_CODI")
								rbac = 145
							End If
							rsDiv.Close
						Else ' Cert ETA
							querySQL = "SELECT CHE.CHE_CODI, CHE.     CHE_RCA, OP.ORGP_SIGLA" & _
										" FROM ( ( A135_CHE AS CHE INNER JOIN A135_Bases AS B ON CHE.CHE_CODI = B.CHE_CODI) " & _
										" INNER JOIN Organizacao AS O ON B.ORG_CODI = O.ORG_CODI) " & _
										" INNER JOIN OrganizacaoP AS OP ON O.ORGP_CODI = OP.ORGP_CODI " & _
										" WHERE CHE.CHE_CODI='" & Field.value & "'"
							Set rsDiv = oDbFDH.getRecSetRd(querySQL)
							If rsDiv.Eof Then
								iRes = 0
								sObs = "<font color=red>Certificado não identificado</font>"
							Else
								sObs = rsDiv("ORGP_SIGLA")
								keyCert = rsDiv("CHE_CODI")
								rbac = 135
							End If
							rsDiv.Close
						End If

					ElseIf i = colResult And iRes > 0 Then
						If InStr(UCase(Field.value),"AERONAVEGÁVEL") <= 0 Then
							iRes = -4
							sObs = "<font color=DarkMagenta>Não identificado status 'AERONAVEGÁVEL'</font>"
						End If
					End If

					If i = colData Or i = colMarca Or i = colCert Or i = colResult Then
						Response.Write "<td>" & Field.value & "</td>"
					End If

					i = i + 1

				Next

				Response.Write "<td>" & sObs & "</td>"
				Response.Write "</tr>"

				If iRes > 0 Then
					If nSuccess = 0 Then ' first
						querySQL = "UPDATE A145_CHE AS CHE SET CHE.CHE_RCA = 0"
						oDbFDH.Execute( querySQL )
						querySQL = "UPDATE A135_CHE AS CHE SET CHE.CHE_RCA = 0"
						oDbFDH.Execute( querySQL )
					End If
					querySQL = "UPDATE A" & rbac & "_CHE AS CHE SET CHE.CHE_RCA = CHE.CHE_RCA + 1 WHERE CHE.CHE_CODI='" & keyCert & "'"
					oDbFDH.Execute( querySQL )
					nSuccess = nSuccess + 1
				End If

				rsExcel.MoveNext
			WEnd
		End If

		Response.Write "</tbody></table>"

		Response.Write "<br><br>" & nSuccess & " records were updated successfully!<br><br>"

		rsExcel.Close
		ExcelConnection.Close

	End If
End If

oDbFDH.Close()
Set oFSO = Nothing

%>
</body>
</html>
