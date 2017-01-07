<%
'Option Explicit

'----------------------------------------------------------------
'
'	Class cTrx
'
'----------------------------------------------------------------
Class cTrx

	'Declarations
Private m_logFile


	Private oCtrlErr			' Error Object
	Private oLog				' Log


	' DB Connection
	Private oDbTrx				' Database Object


	' Data
	Private m_sProcess			' Processo 
	Private	m_sSolic
	Private	m_sSolicStatus
	Private	m_sTask
	Private	m_sTaskStatus

	Private m_sProcUser

	Private m_sUserTrx






	'Class Initialization
	Private Sub Class_Initialize()

		Set oDbTrx = (new cDBAccess)("TRX") ' Database Object
		If oDbTrx.ErrorNumber < 0 then




		End If

		m_logFile = Request.ServerVariables( "APPL_PHYSICAL_PATH" ) & "Public/log.txt"

	End Sub

	'Terminate Class
	Private Sub Class_Terminate()
		' Empty
	End Sub

	Private Sub write(txt)
		Dim fs, f
		Set fs=Server.CreateObject("Scripting.FileSystemObject") 
		If fs.FileExists(m_logFile) Then
			Set f = fs.OpenTextFile(m_logFile, 8)
		Else
			Set f = fs.CreateTextFile(m_logFile, True)
		End If
		f.write(txt & vbCrLf)
		f.close()
		Set f=nothing
		Set fs=nothing
	End Sub

	' Error
	Public Sub Error( txt )
		Dim log
		log = "[ERR][" & FormatDateTime(Now()) & "] " & Request.ServerVariables("AUTH_USER") & " [" & _
			   Request.ServerVariables("REQUEST_METHOD") & " " & Request.ServerVariables("URL") & " " & _
				Request.ServerVariables("SERVER_PROTOCOL") & "] '" & txt & "'"
		write(log)
	End Sub

	' Msg
	Public Sub Msg( txt )
		Dim log
		log = "[MSG][" & FormatDateTime(Now()) & "] " & Request.ServerVariables("AUTH_USER") & " [" & _
			   Request.ServerVariables("REQUEST_METHOD") & " " & Request.ServerVariables("URL") & " " & _
				Request.ServerVariables("SERVER_PROTOCOL") & "] '" & txt & "'"
		write(log)
	End Sub

End Class

%>
