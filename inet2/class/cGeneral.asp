<%
'Option Explicit

'----------------------------------------------------------------
'
'	Class cGeneral
'
'----------------------------------------------------------------
Class cGeneral

	'Declarations

	'Class Initialization
	Private Sub Class_Initialize()
	End Sub

	'Terminate Class
	Private Sub Class_Terminate()
		' Empty
	End Sub


	Public Function Sleep(seconds)
		Set oShell = CreateObject("Wscript.Shell")
		cmd = "%COMSPEC% /c timeout " & seconds & " /nobreak"
		oShell.Run cmd,0,1
	End Function

End Class

%>


