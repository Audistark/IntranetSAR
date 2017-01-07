<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Response.CodePage = 1252 %> 
<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!-- #include virtual = "/inet2/class/cAAA.asp" -->
<!-- #include virtual = "/inet2/lib/libFuncDiv.asp" -->
<%
'------------------------------------------------------
'
'	Grava na Pasta FDH\AvGeral\AIR999
'
'	Usado pelo Excel pra upload de arquivos
'

'Request
On Error Resume Next
Dim oRequest : Set oRequest = GetUpload()
If Err.Number <> 0 Then
	Response.Status = "400 Bad Request"
	Response.Write "Argumentos inválidos."
	Response.End
End If
On Error GoTo 0
Dim rSec : rSec = Request.QueryString("Sec")
Dim Tipo : rTipo = Request.QueryString("Tipo")
If rSec <> "145" And rSec <> "121" And _
	rSec <> "135" And rSec <> "91" Then
	Response.Status = "400 Bad Request"
	Response.Write "Argumentos inválidos."
	Response.End
End If
Dim webuser : webuser = ""
Dim webpass : webpass = ""
Dim res : res = False

' Init AAA - Authentication, Authorization and Accounting
Dim oAAA : Set oAAA = new cAAA
Dim ret : ret = oAAA.WinAuthenticate(True)

If ret > 0 Then
	res = oAAA.AuthorWinMasterSec( rSec )
Else
	webuser = Request.QueryString("user")
	webpass = Request.QueryString("pass")
	ret = oAAA.Authenticate(webuser,webpass)
	If ret < 0 Then
		oAAA.Print()
	End If
	res = oAAA.AuthorMasterSec( rSec )
End If

' Authorization
If res <> True Then
' In order to publish TI Reports, user does not to be MASTER
        if rSec <> "145" or rTipo <> "TI_Report" Then
	Response.Status = "403 Forbidden"
	Response.End
        End If
End If

' Save file
Dim oFile : Set oFile = oRequest("File")
If Not oFile Is Nothing Then
	If oFile.Length > 0 Then
		Dim FileName : FileName = oFile.filename
		Dim FileContent : FileContent = MultiByteToBinary(oFile.value)
		vbsSaveAs Request.ServerVariables("APPL_PHYSICAL_PATH") & "FDH\AvGeral\AIR" & rSec & "\" & FileName, FileContent
    End If
End If

Response.Status = "200 OK"
Response.Write "Successful file upload execution." & "<br>"
Response.Write "File: oFile.filename" & "<br>"
Response.End

 %>
