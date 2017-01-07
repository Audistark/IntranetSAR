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
	Response.Write "Você não está autorizado a acessar esse recurso."
	Response.End
End If

Dim Dir : Dir = Request.QueryString( "Dir" )
If Dir = "" Then
	Response.Status = "400 Bad Request"
	Response.End
End If

Dim fs : Set fs = Server.CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Dim f : Set f = fs.CreateFolder(Request.ServerVariables("APPL_PHYSICAL_PATH") & "FDH\" & Dir)
If Err.Number <> 0 Then
	Response.Status = "200 OK"
	Response.Write "Folder already exists." & "<br>"
Else
	Response.Status = "200 OK"
	Response.Write "Successful folder creation." & "<br>"
End If
On Error GoTo 0
Set f = Nothing
Set fs = Nothing
Response.Write "Folder: " & Request.ServerVariables("APPL_PHYSICAL_PATH") & "FDH\" & Dir
Response.End

%>