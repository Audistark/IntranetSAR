<%

'Option Explicit

' AAA - Authentication, Authorization and Accounting
'
' Authentication
'   Authentication refers to the process where an entity's identity is authenticated, typically by providing evidence that it holds
'   a specific digital identity such as an identifier and the corresponding credentials. Examples of types of credentials are passwords,
'   one-time tokens, digital certificates, digital signatures and phone numbers (calling/called).
'
' Authorization
'   The authorization function determines whether a particular entity is authorized to perform a given activity, typically inherited
'   from authentication when logging on to an application or service. Authorization may be determined based on a range of restrictions,
'   for example time-of-day restrictions, or physical location restrictions, or restrictions against multiple access by the same entity
'   or user. Typical authorization in everyday computer life is for example granting read access to a specific file for authenticated 
'   user.
'
' Accounting
'   Accounting refers to the tracking of network resource consumption by users for the purpose of capacity and trend analysis,
'   cost allocation, billing.[4] In addition, it may record events such as authentication and authorization failures, and include
'   auditing functionality, which permits verifying the correctness of procedures carried out based on accounting data.
'
'

'----------------------------------------------------------------
'
'	Class cAAA
'

Class cAAA

	'Declarations

	Private m_bWinAuthenticated	' User already authenticated by windows
	Private m_bWinAuthorized	' User already authorized

	Private	m_sWinUser			' Windows User
	Private	m_sWinDomain		' Windows Domain

	Private m_bUsrAuthentAuthor	' Both (authenticated and authorizated)
	Private	m_sUsrUser			' User
	Private	m_sUsrDomain		' User Domain

	Private oCtrlErr			' Error Object
	Private oLog				' Log

	Private m_sHostTheft(20)
	Private m_nHostTheft

	'Class Initialization
	Private Sub Class_Initialize()

		Set oCtrlErr = new cCtrlErr
		Set oLog = new cLog

		m_bWinAuthenticated = False
		m_bWinAuthorized = False
		m_sWinDomain = ""
		m_sWinUser = ""

		m_bUsrAuthentAuthor = False
		m_sUsrUser = ""
		m_sUsrDomain = ""

		m_nHostTheft = 0

		Call InitObject()

	End Sub

	'Terminate Class
	Private Sub Class_Terminate()
		Set oCtrlErr = Nothing
		Set oLog = Nothing
	End Sub


	'  Get Error Object
	Public Function getObjErr()
		Set getObjErr = oCtrlErr
	End Function


	'  Print Error
	Public Sub Print
		oCtrlErr.Print()
	End Sub


	' Roda apenas na inicialização do objeto
	Private Sub InitObject()

		'-------------------------------------------'
		'											'
		' Verifica se não está na sessão ativa		'
		'											'
		'-------------------------------------------'

		' Win Authentication and Authorization
		If SessionId <> "" And _
			SessionId = Session.SessionID And _
			 AuthentWinUser <> "" And _
			  AuthentWinDomain <> "" Then

			m_bWinAuthenticated = True

			m_sWinUser = AuthentWinUser
			m_sWinDomain = AuthentWinDomain

			If AuthentWinUserCodi <> "" And _
			    AuthentWinUserName <> "" And _
				 AuthentWinUserNick <> "" And _
			      AuthentWinUserSDiv <> "" And _
				   SessionWinPerm <> "" Then

				m_bWinAuthorized = True

			End If

		End If


		' User
		If SessionId <> "" And _
			SessionId = Session.SessionID And _
			 AuthentUser <> "" And _
			  AuthentDomain <> "" And _
			   AuthentUserCodi <> "" And _
			    AuthentUserName <> "" And _
				 AuthentUserNick <> "" And _
			      AuthentUserSDiv <> "" And _
				   SessionPerm <> "" Then

				m_bUsrAuthentAuthor = True
				m_sUsrUser = AuthentUser
				m_sUsrDomain = AuthentDomain

		Else

			'
			' Vamos tentar ver se ele não está autenticado
			' pelo sistema antigo (dsv pelo Chiessi)
			'
			If Session( "Logado" ) = True And _
				Session( "User" ) <> "" And _
				 Session( "Permissoes" ) <> "" Then

				' Sim !!! estava :(

				' Então converto para o atual

				Dim user : user = Session( "User" )

				' session timeout
				Session.Timeout = 30	' 30 minutes

				' Latin
				Session.Codepage = 1252

				' init Session Id
				InitSessionId()

				' grava o usuario do banco na Sessao
				AuthentUser = user
				AuthentDomain = "ANAC"

				' User Name
				AuthentUserName = Session( "Name" )

				' Nick Name
				AuthentUserNick = UCase(Left(AuthentUserName,InStr(AuthentUserName," ")-1) & " " & Right(AuthentUserName,Len(AuthentUserName)-InStrRev(AuthentUserName," ")))

				' User Codi
				AuthentUserCodi = Session( "PesCodi" )

				' Divisão
				AuthentUserSDiv = Session( "SDiv" )

				' Permissões
				SessionPerm = Session( "Permissoes" )

				m_bUsrAuthentAuthor = True ' Yes!!! authenticated

				m_sUsrUser = AuthentUser
				m_sUsrDomain = AuthentDomain

			End If

		End If

	End Sub


	'--------------------------------------------------------
	' Login Authentication user/pass
	'
	' Authenticate and authorizate
	Public Function Authenticate( user, pass )

		Dim i

		If m_bUsrAuthentAuthor = True And _
			m_sUsrUser = user Then
			Authenticate = 1 ' authenticated
			Exit Function
		End If

		Dim crypt : crypt = ""
		For i = 1 to Len( Trim( pass ) )
			crypt = crypt & Chr( Asc( Mid( Trim( pass ), i, 1 ) ) + 30 + i )
		Next


		'----------------------------------------------------------
		'	Aqui é o seguinte, se o usuário está autenticado pelo
		'	windows no IIS então ele só consegue acesso ao banco de
		'	dados se estiver nos grupos do NTFS para a pasta onde
		'	estiver o arquivo do banco (se for access), ou se o
		'	usuário tem permissão no SQL (se for SQL).
		'	Então pode dar erro aqui, e deve ser tratado como falta
		'	de autorização!
		'

		' DB Connection
		Dim oDbFDH	' Database Object To DB FDH (MSAccess)
		Set oDbFDH = (new cDBAccess)("FDH")
		If oDbFDH.ErrorNumber < 0 then
			m_bUsrAuthentAuthor = False
			Dim sLogonUser : sLogonUser = Request.ServerVariables("LOGON_USER")
			If IsEmpty(sLogonUser) Or IsNull(sLogonUser) Or sLogonUser = "" Then
				oCtrlErr.Import( oDbFDH.getObjErr() )
				oLog.Error("cAAA.Authenticate(): anonymous user cannot access DB FDH" )

			Else
				oCtrlErr.Error = "Authorization Error.<br><br>" & _
								 "O seu usuário da rede ANAC '" & LCase(sLogonUser) & _
								 "' não tem autorização para acessar esta funcionalidade.<br><br>" & _
								 "Atenciosamente,<br>" & _
								 "Equipe de Gestão de TI da SAR"
				oLog.Error("cAAA.Authenticate(): user '" & sLogonUser & "' cannot access DB FDH" )
			End If
			Authenticate = oDbFDH.ErrorNumber
 			Exit Function
		End If


		Dim querySQL
		querySQL =	"SELECT DISTINCT Pes.PES_CODI, Pes.PES_NOME, Pes.PES_NGUERRA, Per.AREA, Per.PER_AREA, TabSDiv.SDIV_SIGLA" & _
					"  FROM (Pessoal AS Pes LEFT JOIN Permissoes AS Per ON Pes.PES_CODI = Per.PES_CODI)" & _
					"		INNER JOIN Tab_Subdivisao AS TabSDiv ON Pes.SDIV_CODI = TabSDiv.SDIV_CODI" & _
					"  WHERE Pes.PES_LOGIN = '" & user & "' AND Pes.PES_SENHA = '" & crypt & "'"
		Dim rsDiv
		Set rsDiv = oDbFDH.getRecSetRd(querySQL)
		If rsDiv Is Nothing then
			oCtrlErr.Import( oDbFDH.getObjErr() )
			m_bUsrAuthentAuthor = False
			Authenticate = oDbFDH.ErrorNumber
			Exit Function
		End If

		' Testa se achou o user com essa senha
		If Not rsDiv.Eof then

			' session timeout
			Session.Timeout = 30	' 30 minutes

			' Latin
			Session.Codepage = 1252

			' init Session Id
			InitSessionId()

			' grava o usuario do banco na Sessao
			AuthentUser = user
			AuthentDomain = "ANAC"

			' User Name
			AuthentUserName = rsDiv( "PES_NOME" )

			' User Nick Name
			AuthentUserNick = rsDiv( "PES_NGUERRA" )

			' User Codi
			Dim PesCodi : PesCodi = rsDiv( "PES_CODI" )
			AuthentUserCodi = PesCodi

			' Divisão
			Dim sDivSigla : sDivSigla = UCase(rsDiv( "SDIV_SIGLA" ))
			AuthentUserSDiv = sDivSigla

			' Permissões
			Dim Permission : Permission = ""
			Dim PerArea
			Do While Not rsDiv.Eof
				if rsDiv( "AREA" ) <> "" then
					PerArea = rsDiv( "PER_AREA" )
				End If
				If PerArea <> "" Then
					Permission = Permission & "[" & PerArea & "]"
				End If
				rsDiv.MoveNext 
			Loop

			' Permissions
			SessionPerm = SessionPerm

			' legado (no futuro pode-se remover isso)
			Session("Nome") = AuthentUser
			Session("User") = AuthentUser
			Session("Login") = AuthentUser
			Session("UserBanco") = AuthentUser
			Session("Dominio") = AuthentDomain
			Session("PesCodi") = AuthentUserCodi
			Session("Permissoes") = SessionPerm

			m_bUsrAuthentAuthor = True
			m_sUsrUser = AuthentUser
			m_sUsrDomain = AuthentDomain

			Authenticate = 1 ' authenticated

		Else
		
			m_bUsrAuthentAuthor = False

			Authenticate = 0 ' not authenticated

		End If

		rsDiv.Close()
		oDbFDH.Close()

	End Function


	'--------------------------------------------------------
	' Windows Authentication
	'
	' 1. Quando entra uma primeira vez ele entra como anonymous
	' 2. Depois se a Response.Status = "401 Acesso Negado" então
	'    a próxima vez ele entra como Windows Authentication
	' 3. E nas vezes seguintes entrará sempre como Windows Authentication!!!
	'
	' Authenticate and authorizate
	Public Function WinAuthenticate( bNeedAuthorization )

		Dim sLogonUser : sLogonUser = Request.ServerVariables("LOGON_USER")
		Dim bAnonymous : bAnonymous = False
		If IsEmpty(sLogonUser) Or IsNull(sLogonUser) Or sLogonUser = "" Then
			bAnonymous = True
		End If

		If m_bWinAuthenticated = True Then	' Autenticado já está

			' Agora vejo se a autorização está ok e se é necessária
			If m_bWinAuthorized = True Or _
			   Not bNeedAuthorization Then ' Ok
				WinAuthenticate = 1 ' authenticated and authorized
				Exit Function
			End If

		End If


		'-----------------------------------------------------
		' Get windows user and domain
		'
		Dim sParse
		' na primeira vez que acessa entra como anonymous
		' e na segunda como windows authenticate
		If bAnonymous = True Then
			Response.Status = "401 Acesso Negado" ' isso força a requisição pedir o user do windows
			Response.End
		End If

		sLogonUser = UCase(sLogonUser)

		sLogonUser = Replace(sLogonUser, "\", "/")
		If InStr(sLogonUser, "/") < 1 Then
			sLogonUser = "UNKNOWN/" & sLogonUser
		End If

		sParse = Split(sLogonUser, "/")

		m_sWinDomain = sParse(0)
		m_sWinUser = sParse(1)
		'
		'--------------------------------------------------------

		m_bWinAuthenticated = True ' Yes!!! authenticated

		m_bWinAuthorized = False ' Yet

		' security domain
		If bNeedAuthorization = True And _
			m_sWinDomain <> "ANAC" And _
			 m_sWinDomain <> "SAR-DEV" And _
			  m_sWinDomain <> "AUDISTARK-X86" Then
			oCtrlErr.Error = "Authorization fail. The domain '" & m_sWinDomain & "' was not recognized as a valid domain."
			WinAuthenticate = oCtrlErr.ErrorNumber
			Exit Function
		End If


		'----------------------------------------------------------
		'	Aqui é o seguinte, se o usuário está autenticado pelo
		'	windows no IIS (segunda vez que entra e consegue ler o
		'	USER_LOGON) então ele só consegue acesso ao banco de
		'	dados se estiver nos grupos do NTFS para a pasta onde
		'	estiver o arquivo do banco (se for access), ou se o
		'	usuário tem permissão no SQL (se for SQL).
		'	Então pode dar erro aqui, e deve ser tratado como falta
		'	de autorização!!!
		'

		' DB Connection
		Dim oDbFDH	' Database Object To DB FDH (MSAccess)
		Set oDbFDH = (new cDBAccess)("FDH")
		If oDbFDH.ErrorNumber < 0 then

			If bNeedAuthorization = True Then

				oCtrlErr.Error = "Authorization Error.<br><br>" & _
								 "O seu usuário da rede ANAC '" & LCase(m_sWinUser) & _
								 "' não tem autorização para acessar esta funcionalidade.<br><br>" & _
								 "Atenciosamente,<br>" & _
								 "Equipe de Gestão de TI da SAR"
				oLog.Error("cAAA.WinAuthenticate(): user '" & m_sWinUser & "' cannot access DB FDH" )
				
				WinAuthenticate = oDbFDH.ErrorNumber
				Exit Function

			Else

				' veja bem.. se não precisa ter usuário válido, então..
				oDbFDH.Close()

				' session timeout
				Session.Timeout = 60 ' 60 minutes

				' Latin
				Session.Codepage = 1252

				' init Session Id
				InitSessionId()

				' grava o usuario do banco na Sessao
				AuthentWinUser = m_sWinUser
				AuthentWinDomain = m_sWinDomain

				AuthentWinUserName = ""
				AuthentWinUserNick = UCase(Left(m_sWinUser,InStr(m_sWinUser,".")-1) & " " & Right(m_sWinUser,Len(m_sWinUser)-InStrRev(m_sWinUser,".")))
				AuthentWinUserCodi = ""
				AuthentWinUserSDiv = "ANAC"
				SessionWinPerm = ""

				WinAuthenticate = 1 ' authenticated and authorized
				Exit Function

			End If

		End If

		Dim querySQL
		querySQL =	"SELECT DISTINCT Pes.PES_CODI, Pes.PES_NOME, Pes.PES_NGUERRA, Per.AREA, Per.PER_AREA, TabSDiv.SDIV_SIGLA" & _
					"  FROM (Pessoal AS Pes LEFT JOIN Permissoes AS Per ON Pes.PES_CODI = Per.PES_CODI)" & _
					"		INNER JOIN Tab_Subdivisao AS TabSDiv ON Pes.SDIV_CODI = TabSDiv.SDIV_CODI" & _
					"  WHERE Pes.PES_LOGIN = '" & m_sWinUser & "'"
		Dim rsDiv
		Set rsDiv = oDbFDH.getRecSetRd(querySQL)
		If rsDiv Is Nothing then
			m_bWinAuthorized = False
			oCtrlErr.Import( oDbFDH.getObjErr() )
			WinAuthenticate = oDbFDH.ErrorNumber
			Exit Function
		End If

		' Testa se achou o user com essa senha
		If Not rsDiv.Eof then

			' session timeout
			Session.Timeout = 60 ' 60 minutes

			' Latin
			Session.Codepage = 1252

			' init Session Id
			InitSessionId()

			' grava o usuario do banco na Sessao
			AuthentWinUser = m_sWinUser
			AuthentWinDomain = m_sWinDomain

			' User Name
			AuthentWinUserName = rsDiv( "PES_NOME" )

			' Nick Name
			AuthentWinUserNick = rsDiv( "PES_NGUERRA" )

			' User Codi
			Dim PesCodi : PesCodi = rsDiv( "PES_CODI" )
			AuthentWinUserCodi = PesCodi

			' Divisão
			Dim sDivSigla : sDivSigla = UCase(rsDiv( "SDIV_SIGLA" ))
			AuthentWinUserSDiv = sDivSigla

			' Permissões
			Dim Permission : Permission = ""
			Dim PerArea
			Do While Not rsDiv.Eof
				if rsDiv( "AREA" ) <> "" then
					PerArea = rsDiv( "PER_AREA" )
				End If
				If PerArea <> "" Then
					Permission = Permission & "[" & PerArea & "]"
				End If
				rsDiv.MoveNext 
			Loop

			' Permissions
			SessionWinPerm = Permission

			m_bWinAuthorized = True

			WinAuthenticate = 1 ' authenticated and authorized

		Else
		
			m_bWinAuthorized = False

			WinAuthenticate = 0 ' not authorized
		
		End If

		rsDiv.Close()
		oDbFDH.Close()

	End Function
	

	' Esse é usado para verificar se houve autenticação de usuário
	Public Function IsUsrAuthenticated()
		If m_bUsrAuthentAuthor = True Then
			IsUsrAuthenticated = True ' authenticated
		Else
			IsUsrAuthenticated = False
		End If
	End Function


	' Esse é usado para verificar se houve autenticação de usuário
	Public Function IsWinAuthenticated()
		If m_bWinAuthenticated = True Then
			IsWinAuthenticated = True ' authenticated
		Else
			IsWinAuthenticated = False
		End If
	End Function


	' Esse é usado para verificar se houve autorização do usuário
	Public Function IsWinAuthorizated()
		If m_bWinAuthorized = True Then
			IsWinAuthorizated = True ' authorizated
		Else
			IsWinAuthorizated = False
		End If
	End Function


	'  Session Id
	Public Property Get SessionId()
		SessionId = Session("Sar.SessionId")
	End Property
	Private Property Let SessionId( id )
		Session("Sar.SessionId") = id
	End Property
	Private Sub InitSessionId()
		Session("Sar.SessionId") = Session.SessionID
	End Sub


	'  Session User
	Public Property Get AuthentUser()
		AuthentUser = Session("Sar.User")
	End Property
	Private Property Let AuthentUser( user )
		Session("Sar.User") = user
	End Property

	'  Session Domain
	Public Property Get AuthentDomain()
		AuthentDomain = Session("Sar.Domain")
	End Property
	Private Property Let AuthentDomain( domain )
		Session("Sar.Domain") = domain
	End Property

	'  Session Internet Domain
	Public Property Get AuthentInetDomain(domain)
		Select Case domain
			Case "ANAC"
				AuthentInetDomain = "anac.gov.br"
			Case "SAR-DEV"
				AuthentInetDomain = "anac.gov.br"
			Case "AUDISTARK-X86"
				AuthentInetDomain = "anac.gov.br"
			Case Else
				AuthentInetDomain = domain
		End Select
	End Property

	'  Session UserCodi (=PES_CODI)
	Public Property Get AuthentUserCodi()
		AuthentUserCodi = Session("Sar.UserCodi")
	End Property
	Private Property Let AuthentUserCodi( PesCodi )
		Session("Sar.UserCodi") = PesCodi
	End Property

	'  Session UserName (=PES_NOME)
	Public Property Get AuthentUserName()
		AuthentUserName = Session("Sar.UserName")
	End Property
	Private Property Let AuthentUserName( PesNome )
		Session("Sar.UserName") = PesNome
	End Property

	'  Session UserName (=PES_NGUERRA)
	Public Property Get AuthentUserNick()
		AuthentUserNick = Session("Sar.UserNick")
	End Property
	Private Property Let AuthentUserNick( PesNick )
		Session("Sar.UserNick") = PesNick
	End Property

	'  Session UserSDiv (=SDIV_SIGLA)
	Public Property Get AuthentUserSDiv()
		AuthentUserSDiv = Session("Sar.UserSDiv")
	End Property
	Private Property Let AuthentUserSDiv( UserSDiv )
		Session("Sar.UserSDiv") = UserSDiv
	End Property

	'  Session Permission
	Public Property Get SessionPerm()
		SessionPerm = Session("Sar.Permission")
	End Property
	Private Property Let SessionPerm( Permission )
		Session("Sar.Permission") = Permission
	End Property

	'  Session Windows User
	Public Property Get AuthentWinUser()
		AuthentWinUser = Session("Sar.WinUser")
	End Property
	Private Property Let AuthentWinUser( user )
		Session("Sar.WinUser") = user
	End Property

	'  Session Windows Domain
	Public Property Get AuthentWinDomain()
		AuthentWinDomain = Session("Sar.WinDomain")
	End Property
	Private Property Let AuthentWinDomain( domain )
		Session("Sar.WinDomain") = domain
	End Property

	'  Session UserCodi (=PES_CODI)
	Public Property Get AuthentWinUserCodi()
		AuthentWinUserCodi = Session("Sar.WinUserCodi")
	End Property
	Private Property Let AuthentWinUserCodi( PesCodi )
		Session("Sar.WinUserCodi") = PesCodi
	End Property

	'  Session UserName (=PES_NOME)
	Public Property Get AuthentWinUserName()
		AuthentWinUserName = Session("Sar.WinUserName")
	End Property
	Private Property Let AuthentWinUserName( PesNome )
		Session("Sar.WinUserName") = PesNome
	End Property

	'  Session UserNick (=PES_NGUERRA)
	Public Property Get AuthentWinUserNick()
		AuthentWinUserNick = Session("Sar.WinUserNick")
	End Property
	Private Property Let AuthentWinUserNick( PesNome )
		Session("Sar.WinUserNick") = PesNome
	End Property

	'  Session UserSDiv (=SDIV_SIGLA)
	Public Property Get AuthentWinUserSDiv()
		AuthentWinUserSDiv = Session("Sar.WinUserSDiv")
	End Property
	Private Property Let AuthentWinUserSDiv( UserSDiv )
		Session("Sar.WinUserSDiv") = UserSDiv
	End Property

	'  Session Permission
	Public Property Get SessionWinPerm()
		SessionWinPerm = Session("Sar.WinPermission")
	End Property
	Private Property Let SessionWinPerm( Permission )
		Session("Sar.WinPermission") = Permission
	End Property


	'--------------------------------------------------------
	'
	' Cookies
	'
	' Note: The Response.Cookies command must appear BEFORE the <html> tag
	'
	Public Sub CleanCookie( name )
		Response.Cookies("SAR" & name).Expires = DateAdd("d",-1,Date())
	End Sub

	' Se expireDays for <= 0 ele morre depois de fechar o browser
	Public Sub SetCookie( name, value, expireDays )
		%>
		<script language="JavaScript" type="text/javascript">
		var d = new Date();
		var u = new Date(); <%
		If expireDays > 0 Then
		%>
		u.setDate(d.getDate() + <%=expireDays %>);
		document.cookie = "SAR<%=name %>=<%=value %>; expires=" + u.toUTCString() + "; path=/"; <%
		Else
		%>
		document.cookie = "SAR<%=name %>=<%=value %>; path=/";<%
		End If
		%>
		</script>
		<%
	End Sub

	Public Function GetCookie( name )
		Dim cookie : cookie = Request.Cookies("SAR" & name)
		GetCookie = cookie
	End Function
	'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	'--------------------------------------------------------
	' Prevent content theft
	'
	Public Property Let AllowContentFromHost( hostname )
		If m_nHostTheft < 20 Then
			m_sHostTheft(m_nHostTheft) = LCase(hostname)
			m_nHostTheft = m_nHostTheft + 1
		End If
	End Property
	Public Property Get ContentTheft()
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		Dim referer : referer = LCase(Request.ServerVariables("HTTP_REFERER"))
		oCtrlErr.Clear()
		If referer = "" Then
			ContentTheft = True
			oCtrlErr.Error = "Authorization Error.<br>The hostname is not allowed to show that content."
			Exit Property
		End If

		Dim found : found = False
		Dim i, pos
		' http://sar/lado.asp
		If UCase(Left(referer,7)) = "HTTP://" Then
			referer = Right(referer,Len(referer)-7)
			pos = InStr(referer,"/")
			If pos > 0 Then
				referer = Left(referer,pos-1)
			End If
		End If
		For i=0 to m_nHostTheft - 1
			If m_sHostTheft(i) = referer Then
				found = True
				Exit For
			End If
		Next
		If Not found Then
			ContentTheft = True
			oCtrlErr.Error = "Authorization Error.<br>The hostname '" & referer & "' is not allowed to show that content."
			Dim fs, f
			Dim logFile : logFile = Request.ServerVariables( "APPL_PHYSICAL_PATH" ) & "Arquivos/PreventContentTheft.txt"
			Set fs=Server.CreateObject("Scripting.FileSystemObject") 
			If fs.FileExists(logFile) Then
				Set f = fs.OpenTextFile(logFile, 8)
			Else
				Set f = fs.CreateTextFile(logFile, True)
			End If
			f.write(Now() & ": " & referer & vbCrLf)
			f.close()
			Set f=nothing
			Set fs=nothing
			Exit Property
		End If
		ContentTheft = False
	End Property
	'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	'--------------------------------------------------------
	' Authorization
	'
	'		PER_AREA	PER_DESC
	'		-----------------------------------------------------------------------
	'		POLICY		Manutenção das interpretações de Requisitos (Policy Files)
	'		MST_GTRAB	Usuário Master do Módulo de Desenvolvimento Regulatório
	'		MST_PHT		Líder dos Coordenadores dos Programas de Certificação
	'		PHT			Coordenador de Programa de Certificação
	'		PROD		Manutenção no cadastro de Produtos Classe I certificados
	'		ISENTO		Manutenção no cadastro de Produtos isentos de certificação
	'		PHT_ORG		Manutenção dos dados de Empresas e Contatos
	'		MST_GR		Usuário Master da GTPN
	'		WEBSITE		Manutenção da Área do Website
	'		PAIS		Manutenção da Área Internacional (Países e Idiomas)
	'		ORG			Manutenção da Área de Organizações
	'		CLASSIFDOC	Classificação e Tipos de Documentos / Processos
	'		SIGLAS		Abreviaturas e Siglas
	'		ACERVO		Acervo Técnico
	'		ACORDO		Acordo Internacional
	'		CI_IS		Circular de Informação / Instrução Suplementar
	'		CONSINT		Consulta Interna
	'		DICTEC		Dicionário de Termos Técnicos
	'		DIRETINT	Diretiva Interna
	'		DA_AD		Diretriz de Aeronavegabilidade e NPR-DA
	'		TCDS		Especificação de Tipo
	'		DOU			Federal Register
	'		FORMULARIO	Formulário Padronizado
	'		GTRAB		Grupos de Trabalho
	'		INFOGR		Informativo Interno da GTPN
	'		MPH_MPR		Manual de Procedimentos
	'		PA/PT		Procedimentos Internos
	'		RBAC		Regulamento Brasileiro da Aviação Civil
	'		CAI			CAI
	'		FCAR		FCAR
	'		ECJ-190		Programa ECJ 190
	'		ERJ-190		Programa ERJ 190
	'		EMB-500		Programa EMB-500
	'		EMB-505		Programa EMB-505
	'		NTEC		Embraer - Programa Novas Tecnologias
	'		MST_PAP		Usuário Master do Grupo de Aprovação de Peças
	'		PAP11		Usuário do sistema
	'		PAP12		Impressão de Relatórios Gerenciais
	'		PAP13		Impressão de Relatórios H.09
	'		PAP14		Indicadores de Processos
	'		PAP15		Impressão de Relatórios de Apoio
	'		MST_GE		Usuário Master da Gerência de Engenharia
	'		MST_GP		Usuário Master da Gerência de Programas
	'		MST_145		Usuário Master do Grupo AIR-145
	'		145_REG		Inclusão e edição dos dados da Empresa 145 e de RPQS
	'		145_BASE	Inclusão e edição dos dados e de registro das Bases
	'		145_PROC	Inclusão e edição de Processos
	'		145_HISTO	Histórico das Operações em Solicitações e Tarefas
	'		145_CONS	Consulta de Processos
	'		MST_91		Usuário Master do Grupo AIR-91
	'		91_H03APRV	Aprovação dos Processos H.03
	'		91_H03FABR	Inclusão e edição dos dados de Requerentes e Produtos
	'		ENGAER		Inclusão e edição dos dados dos Engenheiros Aeronáuticos
	'		91_H03REG	Inclusão e edição dos Processos H.03
	'		91_H03CONS	Consulta dos Processos H.03
	'		MST_GI		Usuário Master do Grupo de Inspeção e Produção
	'		INSP_IHE	Usuário do sistema
	'		PES			Usuário do módulo de Pessoal
	'		MST_TC		Usuário Master dos módulos de Pessoal
	'		COM_AFM		Comentário de AFM
	'		COM_DCA		Comentário de DCA
	'		RCERCF		Manutenção dos dados do RCE/RCF
	'		MST_RC		Usuário Master do Módulo de RCE/RCF
	'		MST_DCA		Usuário Master do Módulo de Comentários de DCA
	'		COM_DCA		Inclusão de DCA para Comentários
	'		MST_AFM		Usuário Master do Módulo de Comentários de AFM
	'		COM_AFM		Inclusão de AFM para Comentários
	'		INSP_EMP	Manutenção do cadastro de Empresas
	'		MST_MEL		Usuário Master do Módulo de Comentários de MMEL
	'		COM_MMEL	Inclusão de MMEL para Comentários
	'		145_DEL		Exclusão de Processos, documentos, solicitações e tarefas
	'		145_ALL		Acesso a todas as DAR
	'		145_CONT	Inclusão e edição de Contatos (RPQS) das Empresas
	'		91_ASCREG	Inclusão e edição da lista de aeronaves sob cuidados
	'		MST_ASC		Usuário Master do módulo de Aeronaves sob Cuidados Especiais
	'		NORMAS		Inclusão e Edição de Outras Normas
	'		121_BASE	Inclusão e edição dos dados e de registro das Bases
	'		121_CONS	Consulta de Processos
	'		121_CONT	Inclusão e edição de Contatos (Dir./Insp.) das Empresas
	'		121_DEL		Exclusão de Processos, documentos, solicitações e tarefas
	'		121_HISTO	Histórico das Operações em Solicitações e Tarefas
	'		121_ALL		Acesso a todas as DAR
	'		121_PROC	Inclusão e edição de Processos
	'		121_REG		Inclusão e edição dos dados da Empresa 121 e de Dir./Insp.
	'		135_BASE	Inclusão e edição dos dados e de registro das Bases
	'		135_CONS	Consulta de Processos
	'		135_CONT	Inclusão e edição de Contatos (Dir./Insp.) das Empresas
	'		135_DEL		Exclusão de Processos, documentos, solicitações e tarefas
	'		135_HISTO	Histórico das Operações em Solicitações e Tarefas
	'		135_ALL		Acesso a todas as DAR
	'		135_PROC	Inclusão e edição de Processos
	'		135_REG		Inclusão e edição dos dados da Empresa 135 e de Dir./Insp.
	'		MST_121		Usuário Master do Grupo AIR-121
	'		MST_135		Usuário Master do Grupo AIR-135
	'		MST_H03		Usuário Master do módulo de Processos H.03
	'		91_AL01APR	Aprovação dos Processos AL.01
	'		91_AL01CON	Consulta dos Processos AL.01
	'		91_AL01REG	Inclusão e edição dos Processos AL.01
	'		MST_AL01	Usuário Master do módulo de Processos AL.01
	'		MMEL		Permissão para manutenção dos dados de MMEL
	'		GTPN		Fatima Siqueira(Fátima Aparecida Fabricio Siqueira)
	'

	' Master
	Public Property Get AuthorMaster()
		If InStr( SessionPerm, "[MASTER]" ) <> 0 Then
			 AuthorMaster = True
		Else
			 AuthorMaster = False
		End If
	End Property
	Public Property Get AuthorWinMaster()
		If InStr( SessionWinPerm, "[MASTER]" ) <> 0 Then
			 AuthorWinMaster = True
		Else
			 AuthorWinMaster = False
		End If
	End Property

	'--------------------------------------------------------
	' Master Sec
	'
	'	MST_GTRAB	Usuário Master do Módulo de Desenvolvimento Regulatório
	'	MST_PHT		Líder dos Coordenadores dos Programas de Certificação
	'	MST_GR		Usuário Master da GTPN
	'	MST_PAP		Usuário Master do Grupo de Aprovação de Peças
	'	MST_GE		Usuário Master da Gerência de Engenharia
	'	MST_GP		Usuário Master da Gerência de Programas
	'	MST_91		Usuário Master do Grupo AIR-91
	'	MST_121		Usuário Master do Grupo AIR-121
	'	MST_135		Usuário Master do Grupo AIR-135
	'	MST_145		Usuário Master do Grupo AIR-145
	'	MST_GI		Usuário Master do Grupo de Inspeção e Produção
	'	MST_TC		Usuário Master dos módulos de Pessoal
	'	MST_RC		Usuário Master do Módulo de RCE/RCF
	'	MST_DCA		Usuário Master do Módulo de Comentários de DCA
	'	MST_AFM		Usuário Master do Módulo de Comentários de AFM
	'	MST_MEL		Usuário Master do Módulo de Comentários de MMEL
	'	MST_ASC		Usuário Master do módulo de Aeronaves sob Cuidados Especiais
	'	MST_H03		Usuário Master do módulo de Processos H.03
	Public Property Get AuthorMasterSec( Sec )
		If AuthorMaster = True Then
			AuthorMasterSec = True
			Exit Property
		End If
		If InStr( SessionPerm, "[MST_" & Sec & "]" ) <> 0 Then
			 AuthorMasterSec = True
		Else
			 AuthorMasterSec = False
		End If
	End Property
	Public Property Get AuthorWinMasterSec( Sec )
		If AuthorWinMaster = True Then
			AuthorWinMasterSec = True
			Exit Property
		End If
		If InStr( SessionWinPerm, "[MST_" & Sec & "]" ) <> 0 Then
			 AuthorWinMasterSec = True
		Else
			 AuthorWinMasterSec = False
		End If
	End Property

	' ADM
	Public Property Get AuthorAdminSec( Sec )
		If AuthorMaster = True Then
			AuthorAdminSec = True
			Exit Property
		End If
		If InStr( SessionPerm, "[" & Sec & "_ADM]" ) <> 0 Then
			 AuthorAdminSec = True
		Else
			 AuthorAdminSec = False
		End If
	End Property
	Public Property Get AuthorWinAdminSec( Sec )
		If AuthorWinMaster = True Then
			AuthorWinAdminSec = True
			Exit Property
		End If
		If InStr( SessionWinPerm, "[" & Sec & "_ADM]" ) <> 0 Then
			 AuthorWinAdminSec = True
		Else
			 AuthorWinAdminSec = False
		End If
	End Property

	' LDR
	Public Property Get AuthorLiderSec( Sec )
		If AuthorMaster = True Then
			AuthorLiderSec = True
			Exit Property
		End If
		If InStr( SessionPerm, "[" & Sec & "_LDR]" ) <> 0 Then
			 AuthorLiderSec = True
		Else
			 AuthorLiderSec = False
		End If
	End Property
	Public Property Get AuthorWinLiderSec( Sec )
		If AuthorWinMaster = True Then
			AuthorWinLiderSec = True
			Exit Property
		End If
		If InStr( SessionWinPerm, "[" & Sec & "_LDR]" ) <> 0 Then
			 AuthorWinLiderSec = True
		Else
			 AuthorWinLiderSec = False
		End If
	End Property

	' Upload (Não utilizado por conta alguma por enquanto!!!)
	Public Property Get AuthorUpld( Sec )
		' 91_UPLOAD 145_UPLOAD ..
		If InStr( SessionPerm, "[" & Sec & "_UPLOAD]" ) <> 0 Then
			 AuthorUpld = True
		Else
			 AuthorUpld = False
		End If
	End Property

	' 91 - Aeronaves

	' H03 - Experimental
	Public Property Get AuthorH03
		If InStr( SessionPerm, "[91_H03REG]" ) <> 0 Then
			 AuthorH03 = True
		Else
			 AuthorH03 = False
		End If
	End Property

	' AL.01 - Anv Leve Esportiva
	Public Property Get AuthorAL01
		If InStr( SessionPerm, "[91_AL01APR]" ) <> 0 Then
			 AuthorAL01 = True
		Else
			 AuthorAL01 = False
		End If
	End Property

	' Generico - Testa String
	Public Property Get Author( text )
		If InStr( SessionPerm, "[" & text & "]" ) <> 0 Then
			 Author = True
		Else
			 Author = False
		End If
	End Property



	'----------------------------------------------
	' Financeiro
	'
	' "0" - Usuario
	' "1" - Administrador
	' "2" - Master
	' "9" - Sem acesso
	'
	Public Function AuthorFinanceiro(user)

		' DB Connection
		Dim oDbFinanc	' Database Object To DB FDH (MSAccess)
		Set oDbFinanc = (new cDBAccess)("Financeiro")
		If oDbFinanc.ErrorNumber < 0 then
			oCtrlErr.Import( oDbFinanc.getObjErr() )
			AuthorFinanceiro = oDbFinanc.ErrorNumber
			Exit Function
		End If
	
		Dim querySQL
		querySQL =	"SELECT * FROM Pessoal WHERE nome = '" & LCase(user) & "'"

		Dim rsDiv
		Set rsDiv = oDbFinanc.getRecSetRd(querySQL)
		If rsDiv Is Nothing then
			oCtrlErr.Import( oDbFinanc.getObjErr() )
			AuthorFinanceiro = oDbFinanc.ErrorNumber
			Exit Function
		End If

		' Testa se achou para esse user
		If Not rsDiv.Eof then
			AuthorFinanceiro = rsDiv("Nv_Access")
		Else
			AuthorFinanceiro = -1
		End If

		If AuthorFinanceiro <> 0 And _
			AuthorFinanceiro <> 1 And _
			 AuthorFinanceiro <> 2 And _
			  AuthorFinanceiro <> 9 Then
			oCtrlErr.Error = "Unauthorized"
			AuthorFinanceiro = oCtrlErr.ErrorNumber
		End If

		oDbFinanc.Close()

	End Function

'    Dim cnxSql
'    Set cnxSql = OpenConn( sqlDBEventos )
'
'    Dim author
'    Set author = cnxSql.execute("select * from Acessos where usuarios = '" & user & "' order by usuarios asc")
'    If Not author.EOF Then
'        nome = tabela_eventos.fields("usuarios")
'        area = tabela_eventos.fields("area")
'        setor = tabela_eventos.fields("setor")
'		 divisao = tabela_eventos.fields("subsetor")
'
'        'permissao = tabela_eventos.fields("permissao")
'        ' Essa tabela de permissoes está só no código
'        ' 1- "Administrador"
'        ' 2- "Superintendente"
'        ' 3- "Gerente-Geral"
'        ' 4- "Gerente"
'        ' 5- "Chefe-Setor"
'        ' 6- "Usuário"
'        ' 7- "Usuário Avançado"
'        ' 8- "Sem Privilégios"
'
'    End If
'

End Class

%>
