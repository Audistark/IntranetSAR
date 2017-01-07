<%
'Option Explicit

'----------------------------------------------------------------
'
'	Class cCtrlErr
'
'----------------------------------------------------------------
Class cCtrlErr

	'Declarations
	Private m_nErrNumber	' Error
	Private m_sErrDescr		' Error

	'Class Initialization
	Private Sub Class_Initialize()
		m_nErrNumber = 0
		m_sErrDescr = ""
	End Sub

	'Terminate Class
	Private Sub Class_Terminate()
		' Empty
	End Sub

	Public Sub Clear()
		m_nErrNumber = 0
		m_sErrDescr = ""
	End Sub

	'  Get Error
	Public Property Get ErrorNumber()
		ErrorNumber = m_nErrNumber
	End Property
	Public Property Get ErrorDescr()
		ErrorDescr = m_sErrDescr
	End Property

	'  Let Error
	Public Property Let ErrorNumber( number )
		m_nErrNumber = number
	End Property
	Public Property Let ErrorDescr( descr )
		m_sErrDescr = descr
	End Property
	Public Property Let Error( descr )
		m_nErrNumber = -1
		m_sErrDescr = descr
	End Property

	'  Set Error
	Public Function SetError(nbr, str)
		m_nErrNumber = nbr
		m_sErrDescr = str
	End Function

	'  Import Error
	Public Function Import(obj)
		Dim oErr
		Set oErr = obj
		m_nErrNumber = oErr.ErrorNumber
		m_sErrDescr = oErr.ErrorDescr
	End Function

	'  Print
	Public Sub Print
		If m_nErrNumber <> 0 then
			Response.Clear
			Response.Write( "Error: 0x" & Hex(m_nErrNumber) & "<br>" & "Descr: " & m_sErrDescr )
			Response.End
		End If
	End Sub

End Class

%>


