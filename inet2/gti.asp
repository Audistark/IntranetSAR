<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.CodePage = 1252 %>
<!-- #include virtual = "/inet2/class/cCtrlErr.asp"-->
<!-- #include virtual = "/inet2/class/cLog.asp"-->
<!-- #include virtual = "/inet2/class/cDBAccess.asp" -->
<!-- #include virtual = "/inet2/class/cAAA.asp" -->
<%
' Init AAA - Authentication, Authorization and Accounting
Dim oAAA : Set oAAA = new cAAA
' Windows Authentication
Dim ret : ret = oAAA.WinAuthenticate(False)
If ret < 0 Then
	oAAA.Print()
End If
Dim user : user = oAAA.AuthentWinUser
user = LCase(user)

Dim updated : updated = "[Atualizando]"
Dim fso : Set fso = Server.CreateObject("Scripting.FileSystemObject")
Dim filename : filename = Request.ServerVariables("APPL_PHYSICAL_PATH") & "FDH\AvGeral\AIR145\ti-update.txt"
On Error Resume Next
Dim fle : Set fle = fso.OpenTextFile(filename)
If Err.Number = 0 Then
	'Do Until fle.AtEndOfStream
	' Last Updated Date: 25/10/2013 13:49:23
	updated = Right( fle.ReadLine, 19)
	'Loop
	fle.Close()
End If
On Error GoTo 0
Set fle = Nothing
Set fso = Nothing

Dim Title : Title = "Gestão de TI-SAR"

 %>
<!-- #include file="gti.html" -->
