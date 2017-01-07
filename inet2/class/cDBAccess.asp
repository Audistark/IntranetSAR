<%
'Option Explicit

'----------------------------------------------------------------
'
'	Class cDBAccess( DatabaseName )
'
'	Date: 06/10/2013
'	Author: Henri Bigatti <henri.bigatti@anac.gov.br>
'
'----------------------------------------------------------------
'
' Usage:
'
'	Dim oDBEventos
'	Set oDBEventos = (new cDBAccess)( "Eventos" )
'	If oDBEventos.ErrorNumber < 0 then
'		Response.write(oDBEventos.ErrorDescr)
'		Response.End
'	End If
'
'	' With Record Set Cursor (can be used until 20 cursors simultaneous)
'	Dim querySQL
'	querySQL = "Select * from Acessos where Usuarios = '" & usuario & "' order by usuarios asc"
'	Dim rsDiv
'	Set rsDiv = oDBEventos.getRecSetRd(querySQL)
'	If rsDiv Is Nothing then
'		Response.write(oDBEventos.ErrorDescr)
'		Response.End
'	End If
'	If Not rsDiv.EOF Then
'		nome = rsDiv("usuarios")
'		area = rsDiv("area")
'		setor = rsDiv("setor")
'		permissao = rsDiv("permissao")
'	End If
'	rsDiv.Close()
'
'	' Using Execute()
'	Set TabUsers = oDBEventos.Execute( "Select * from Acessos where Usuarios = '" & usuario & "' order by usuarios asc" )
'
'	If Not TabUsers Is Nothing then
'		If Not TabUsers.EOF Then
'			nome = TabUsers.fields("usuarios")
'			area = TabUsers.fields("area")
'			setor = TabUsers.fields("setor")
'			permissao = TabUsers.fields("permissao")
'		End If
'	Else 
'		Response.write(oDBEventos.ErrorDescr)
'		Response.End
'	End If
'	oDBEventos.Close()
'

' FINANCEIRO
Const constDBAccessFinanceiro   = 1

' FINANCEIRO_RJ
Const constDBAccessFinanceiroRJ = 2

' EVENTOS (MS SQL)
Const constDBAccessEventos      = 3

' FDH
Const constDBAccessFDH          = 4

' Transactions (MS SQL)
Const constDBTrxProc			= 5

' IN81 (PostgreSQL)
Const constDBAccessPostgreSQL   = 6

' Error
Const constDBAccessError        = -1


Const MAX_RS					= 20

'----------------------------------------------------------------
'
'	Class cDBAccess( DatabaseName )
'

Class cDBAccess

	'Declarations
	Private m_oDBConn		' Connection
	Private m_oDBRecSet(20)	' Record Set
	Private m_iDB
	Private oCtrlErr		' Error Object

	'Class Initialization
	Private Sub Class_Initialize()
		Set m_oDBConn = Nothing
		Dim i
		For i = 1 to MAX_RS
			Set m_oDBRecSet(i) = Nothing
		Next
		Set oCtrlErr = new cCtrlErr
	End Sub
	Public Default Function construct( pPars )
		Select Case UCase(pPars)
			Case "FINANCEIRO"
				m_iDB = constDBAccessFinanceiro
			Case "FINANCEIRO_RJ"
				m_iDB = constDBAccessFinanceiroRJ
			Case "EVENTOS"
				m_iDB = constDBAccessEventos
			Case "FDH"
				m_iDB = constDBAccessFdh
			Case "TRX"
				m_iDB = constDBTrxProc
			Case "POSTGRESQL"
				m_iDB = constDBAccessPostgreSQL
			Case Else
				m_iDB = constDBAccessError ' -1
		End Select
		Call Init()
        set construct = me
    end function

	'Terminate Class
	Private Sub Class_Terminate()
		Call Close()
		Set oCtrlErr = Nothing
	End Sub


	' Open Connection
	Private Function Init()

		' server sql
		Dim sqlServerSAR
		sqlServerSAR = "svarj1220.anac.gov.br"

		Dim strConnection
		Select Case m_iDB

			Case constDBAccessFinanceiro
				strConnection = "Dsn=Financeiro;uid=ggcp;pwd=ggcpb;"

			Case constDBAccessFinanceiroRJ
				strConnection = "Dsn=Financeiro_RJ;uid=ggcp;pwd=ggcpb;"

			Case constDBAccessEventos
				' MS SQL Database Eventos (Acessos, Eventos, etc..)
				strConnection = "Driver={SQL Server};server=" & sqlServerSAR & ";" +_
								"uid=ggcp;pwd=ggcpb;database=Eventos"

			Case constDBAccessFdh
				strConnection = "Dsn=FDH;uid=fdh;pwd=#CcB&_25Chi38;"

			Case constDBTrxProc
				' MS SQL Database
				strConnection = "Driver={SQL Server};server=" & sqlServerSAR & ";" +_
								"uid=ggcp;pwd=ggcpb;database=Eventos"

			Case constDBAccessPostgreSQL
				' PostgreSQL
				strConnection = "Driver={PostgreSQL};server=svarj1210.anac.gov.br;port=5432;" +_
								"uid=sar_select;pwd=sar123;database=fiscalizacao"

			Case Else
				strConnection = "Null"

		End Select

		Set m_oDBConn = (new cDBConn)(strConnection)

		oCtrlErr.Import( m_oDBConn.getObjErr() )

		If oCtrlErr.ErrorNumber < 0 then
			Init = -1 ' Bad
		Else
			Init = 1 ' OK
		End If

	End Function

	'  Get Error
	Public Property Get ErrorNumber()
		ErrorNumber = oCtrlErr.ErrorNumber
	End Property
	Public Property Get ErrorDescr()
		ErrorDescr = oCtrlErr.ErrorDescr
	End Property


	Public Function getRecSetRd( querySql )

		' Pega o primeiro que nao estiver em uso
		Dim i
		Dim found : found = False
		For i = 1 to MAX_RS
			If Not m_oDBRecSet(i) Is Nothing then
				If Not m_oDBRecSet(i).getRecSet() Is Nothing then
					found = True
					Exit For
				End If
			End If
		Next
		If Not found then
			For i = 1 to MAX_RS
				If m_oDBRecSet(i) Is Nothing then
					Set m_oDBRecSet(i) = (new cDBRecSet)( m_oDBConn )
					found = True
					Exit For
				Else
				End If
			Next
		End If
		If Not found then
			oCtrlErr.Error = "Error: reached maximum number of RecSet allowed"
			Set getRecSetRd = Nothing
			Exit Function
		End If
		m_oDBRecSet(i).setRecToRead()
		Set getRecSetRd = m_oDBRecSet(i).getRecSetSQL( querySql )
		oCtrlErr.Import( m_oDBRecSet(i).getObjErr() )
	End Function

	Public Function getRecSetWr( querySql )
		' Pega o primeiro que nao estiver em uso
		Dim i
		Dim found : found = False
		For i = 1 to MAX_RS
			If Not m_oDBRecSet(i) Is Nothing then
				If Not m_oDBRecSet(i).getRecSet() Is Nothing then
					found = True
					Exit For
				End If
			End If
		Next
		If Not found then
			For i = 1 to MAX_RS
				If m_oDBRecSet(i) Is Nothing then
					Set m_oDBRecSet(i) = (new cDBRecSet)( m_oDBConn )
					found = True
					Exit For
				Else
				End If
			Next
		End If
		If Not found then
			oCtrlErr.Error = "Error: reached maximum number of RecSet allowed"
			Set getRecSetWr = Nothing
			Exit Function
		End If
		m_oDBRecSet(i).setRecToWrite()
		Set getRecSetWr = m_oDBRecSet(i).getRecSetSQL( querySql )
		oCtrlErr.Import( m_oDBRecSet(i).getObjErr() )
	End Function

	' Execute
	Public Function Execute( querySql )
		Set Execute = m_oDBConn.Execute( querySql )
		oCtrlErr.Import( m_oDBConn.getObjErr() )
	End Function

	'  Get Error Object
	Public Function getObjErr()
		Set getObjErr = oCtrlErr
	End Function

	'  Print Error
	Public Sub Print
		oCtrlErr.Print()
	End Sub

	' HTML Encode
	Public Function HTMLEncode(text)
		HTMLEncode = Server.HTMLEncode(text)
	End Function
	
	' SQL Encode (Remove Quotes)
	Public Function SQLEncode(texto)
		'-- Tratar aspas simples
		texto = Replace("" & texto,"'", "''")
		'-- Tratar aspas duplas (especial)
		texto = Replace(texto,"“", chr(34))
		texto = Replace(texto,"”", chr(34))
		if len(texto)=0 then
			SQLEncode = "NULL"
		else
			SQLEncode = "'" & texto & "'"
		end if
	End Function

	' Close
	Public Sub Close()
		Dim i
		For i = 1 to MAX_RS
			If Not m_oDBRecSet(i) Is Nothing then
				m_oDBRecSet(i).Close()
				Set m_oDBRecSet(i) = Nothing
			End If
		Next
		If Not m_oDBConn Is Nothing then
			m_oDBConn.Close()
			Set m_oDBConn = Nothing
		End If
	End Sub

End Class


'----------------------------------------------------------------
'
'	Class cDBRecSet(oConn)
'
'----------------------------------------------------------------

Class cDBRecSet

	'Declarations
	Private m_oConn		' connection object
	Private m_rsDiv		' record set
	Private oCtrlErr	' Error Object

	'Class Initialization
	Private Sub Class_Initialize()
		Set m_oConn = Nothing
		Set m_rsDiv = Server.CreateObject("ADODB.RecordSet")
		Set oCtrlErr = new cCtrlErr
	End Sub
	Public Default Function construct( pPars )
		Set m_oConn = pPars	' Connection object
		setRecToRead()
        set construct = me
    end function

	'Terminate Class
	Private Sub Class_Terminate()
		if Not m_rsDiv Is Nothing then
			Close()
			Set m_rsDiv = Nothing
		End If
		Set oCtrlErr = Nothing
	End Sub

	'  Get Error
	Public Property Get ErrorNumber()
		ErrorNumber = oCtrlErr.ErrorNumber
	End Property
	Public Property Get ErrorDescr()
		ErrorDescr = oCtrlErr.ErrorDescr
	End Property

	'  Get Error Object
	Public Function getObjErr()
		Set getObjErr = oCtrlErr
	End Function

	' Get Record Set
	Public Function getRecSetSQL( querySql )
		If m_oConn is Nothing Or m_rsDiv Is Nothing then
			Set getRecSetSQL = Nothing
			oCtrlErr.Error = "Error: Unknown handles"
		Else
			oCtrlErr.Clear()
			On Error Resume Next
			Err.Clear
			' Se aberta fecha primeiro
			If m_rsDiv.State = 1 then ' adStateOpen
				m_rsDiv.Close
			End If
			Call m_rsDiv.Open( querySql, m_oConn.getHandle() )
			If Err.Number <> 0 then
				oCtrlErr.Error = "0x" & Hex(Err.Number) & "-" & Err.Description & _
								 "<br>" & "Function: " & Err.Source
				Set getRecSetSQL = Nothing
			Else
				Set getRecSetSQL = m_rsDiv
			End If
			On Error Goto 0
		End If
	End Function


	Public Function getRecSet()
		If m_oConn is Nothing Or m_rsDiv Is Nothing then
			oCtrlErr.Error = "Error: Unknown handles"
			Set getRecSet = Nothing
		Else
			oCtrlErr.Clear()
			' Se fechada então devolve o handle do Record Set para poder ser reaproveitado
			If m_rsDiv.State <> 1 then ' adStateOpen
				Set getRecSet = m_rsDiv
			Else
				Set getRecSet = Nothing
			End If
		End If
	End Function


	' set record set options to read only
	Public Sub setRecToRead()

		' Se aberto fecha primeiro
		If m_rsDiv.State = 1 then ' adStateOpen
			m_rsDiv.Close
		End If

		' read only cursor

		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Constant	  Value	Description
		' adUseNone		1	OBSOLETE (appears only for backward compatibility).
		' adUseServer	2	Default. Uses a server-side cursor.
		' adUseClient	3	Uses a client-side cursor supplied by a local cursor library. For backward
		'					compatibility, the synonym adUseClientBatch is also supported
		'
		m_rsDiv.CursorLocation = 3 ' adUseClient

		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Constant		  Value	Description
		' adOpenUnspecified	-1	Does not specify the type of cursor.
		' adOpenForwardOnly	0	Default. Uses a forward-only cursor. Identical to a static cursor, except
		'						that you can only scroll forward through records. This improves performance
		'						when you need to make only one pass through a Recordset.
		' adOpenKeyset		1	Uses a keyset cursor. Like a dynamic cursor, except that you can't see
		'						records that other users add, although records that other users delete are
		'						inaccessible from your Recordset. Data changes by other users are still visible.
		' adOpenDynamic		2	Uses a dynamic cursor. Additions, changes, and deletions by other users are
		'						visible, and all types of movement through the Recordset are allowed, except
		'						for bookmarks, if the provider doesn't support them.
		' adOpenStatic		3	Uses a static cursor. A static copy of a set of records that you can use to
		'						find data or generate reports. Additions, changes, or deletions by other users
		'						are not visible.
		m_rsDiv.CursorType = 3 ' adOpenStatic

		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Constant		  Value	Description
		' adLockUnspecified	-1	Unspecified type of lock. Clones inherits lock type from the original Recordset.
		' adLockReadOnly	1	Read-only records
		' adLockPessimistic	2	Pessimistic locking, record by record. The provider lock records immediately after editing
		' adLockOptimistic	3	Optimistic locking, record by record. The provider lock records only when calling update
		' adLockBatchOptimistic	4	Optimistic batch updates. Required for batch update mode
		m_rsDiv.LockType = 1 ' adLockReadOnly

	End Sub


	' set record set options to write
	Public Sub setRecToWrite()

		' Se aberto fecha primeiro
		If m_rsDiv.State = 1 then ' adStateOpen
			m_rsDiv.Close
		End If

		' write cursor
		m_rsDiv.CursorLocation = 2 ' adUseServer
		m_rsDiv.CursorType = 0 ' adOpenForwardOnly
		m_rsDiv.LockType = 3 ' adLockOptimistic

	End Sub


	'Close RecordSet
	Public Sub Close()
		If m_rsDiv.State = 1 then ' adStateOpen
			m_rsDiv.Close
		End If
	End Sub

End Class

'----------------------------------------------------------------
'
'	Class cDBConn
'
'----------------------------------------------------------------

Class cDBConn

	'Declarations
	Private m_hdConn	' connection handle
	Private m_strConn	' string connection
	Private oCtrlErr	' Error Object

	'Class Initialization
	Private Sub Class_Initialize()
		Set m_hdConn = Nothing
		m_strConn = ""
		Set oCtrlErr = new cCtrlErr
	End Sub
	Public Default Function construct( pPars )
		m_strConn = pPars
		Call Init()
        set construct = me
    End Function

	'Terminate Class
	Private Sub Class_Terminate()
		if Not m_hdConn Is Nothing then
			m_hdConn.Close
			Set m_hdConn = Nothing
		End If
		Set oCtrlErr = Nothing
	End Sub

	'  Get Error
	Public Property Get ErrorNumber()
		ErrorNumber = oCtrlErr.ErrorNumber
	End Property
	Public Property Get ErrorDescr()
		ErrorDescr = oCtrlErr.ErrorDescr
	End Property

	'  Get Error Object
	Public Function getObjErr()
		Set getObjErr = oCtrlErr
	End Function

	'Accessor method, controlled aliased
	'READ access to your class variable
	Public Property Get strConnection()
		StrConnection = m_strConn
	End Property
	
	'  Get Connection handle
	Public Property Get getHandle()
		Set getHandle = m_hdConn
	End Property

	'Connect Database
    Private Function Init()
		Set m_hdConn = server.CreateObject("ADODB.Connection")
        m_hdConn.ConnectionString = m_strConn
		oCtrlErr.Clear()
		On Error Resume Next
		m_hdConn.Open
		If Err.Number <> 0 then
			oCtrlErr.ErrorNumber = -1
			If m_hdConn.Errors.Count > 0 Then
				' Enumerate Errors collection and display
				' properties of each Error object.
				Dim errLoop
				For Each errLoop In m_hdConn.Errors
					Dim sErrDescr
					sErrDescr = "DB Error: 0x" & Hex(errLoop.Number) & "<br>" & _
								"   " & errLoop.Description & "<br>" & _
								"   (Source: " & errLoop.Source & ")" & "<br>" & _
								"   (SQL State: " & errLoop.SQLState & ")" & "<br>" & _
								"   (NativeError: " & errLoop.NativeError & ")" & "<br>"
					If errLoop.HelpFile = "" Then
						sErrDescr =	sErrDescr & _
									"   No Help file available" & _
									"<br><br>"
					Else
						sErrDescr =	sErrDescr & _
									"   (HelpFile: " & errLoop.HelpFile & ")" & "<br>" & _
									"   (HelpContext: " & errLoop.HelpContext & ")" & _
									"<br><br>"
					End If
					oCtrlErr.ErrorDescr = sErrDescr
				Next
			Else
				oCtrlErr.ErrorDescr = "Open Connection with database failed"
			End If
			Init = -1
			Set m_hdConn = Nothing
		Else
			Init = 1
		End If
		On Error Goto 0
	End Function


	' Execute
	Public Function Execute( querySql )
		If m_hdConn is Nothing then
			Set Execute = Nothing
		Else
			On Error Resume Next
			Err.Clear
			Set Execute = m_hdConn.Execute( querySql )
			If Err.Number <> 0 then
				Set Execute = Nothing
				oCtrlErr.ErrorNumber = -1
				If m_hdConn.Errors.Count > 0 Then
					' Enumerate Errors collection and display
					' properties of each Error object.
					Dim errLoop
					For Each errLoop In m_hdConn.Errors
						Dim sErrDescr
						sErrDescr = "Error: #" & errLoop.Number & "<br>" & _
									"   " & errLoop.Description & "<br>" & _
									"   (Source: " & errLoop.Source & ")" & "<br>" & _
									"   (SQL State: " & errLoop.SQLState & ")" & "<br>" & _
									"   (NativeError: " & errLoop.NativeError & ")" & "<br>"
						If errLoop.HelpFile = "" Then
							sErrDescr =	sErrDescr & _
										"   No Help file available" & _
										"<br><br>"
						Else
							sErrDescr =	sErrDescr & _
										"   (HelpFile: " & errLoop.HelpFile & ")" & "<br>" & _
										"   (HelpContext: " & errLoop.HelpContext & ")" & _
										"<br><br>"
						End If
						oCtrlErr.ErrorDescr = sErrDescr
					Next
				Else
					oCtrlErr.ErrorDescr = "Execute failed"
				End If
			End If
			On Error Goto 0
		End If
	End Function


	'Close Database
	Public Sub Close()
		if not m_hdConn is Nothing then
			m_hdConn.Close
			Set m_hdConn = Nothing
		End If
	End Sub


End Class

%>
