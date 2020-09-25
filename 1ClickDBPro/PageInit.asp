<%
'1 Click DB Pro copyright 1997-2004 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com/

'**Start Encode**

'Begin PageInit

Option Explicit

Response.Buffer = True

'Declare constants for ocdCompatibility bitmap
Const ocdNoFrames = 1
Const ocdNoCookies = 16
Const ocdNoJavascript = 256

'Declare global configuration variables for Config.asp
Dim ocdFormNullToken, ocdShowImport, ocdQuotePrefix, ocdQuoteSuffix, ocdShowReports, ocdHideAutonumber,ocdConnectionShortcuts, ocdAllowBrand, ocdProcWizard, ocdAuditWizardPrefix, ocdProcCodeWizard, ocdAuditWizard, ocdBrandText, ocdBrandLogo, ocdDBTimeout, ocdShowTopMenu, ocdShowCheckedSearchFields, ocdForceExportDownload, ocdQueryWizardIsDefault, ocdMaxURLLength, ocdConnectWizardOnly, ocdShowTopFrame, ocdAllowExport, ocdReadOnly, ocdAllowProAdmin, ocdDisableTextDriver, ocdAdminPassword, ocdJETSQLReference, ocdUseAnsiQuotes, ocdMSSQLReference, ocdOraSQLReference, ocdAllowAdmin, ocdUseCustomEditPages, ocdMultipleFieldSort, ocdAccessTableEdits, ocdAllowBrowseRefresh, ocdDefaultTextCompare, ocdSessionTimeout, ocdShowDefaults, ocdShowHelp, ocdShowSQLSelector, ocdShowSQLConnector, ocdShowSQLCommander, ocdShowSQLExecutor, ocdLaunchPage, ocdRootDirCodeWiz, ocdSQLTableEdits, ocdShowGraph, ocdShowTableMenu, ocdFormEmptyStringIsNull, ocdShowSchema, ocdSecureCodeWiz, ocdHeaderHTML, ocdAllowCodeWiz, ocdFooterHTML, ocdADOConnection, ocdADOUsername, ocdADOPassword, ocdShowRelatedRecords, ocdSelectForeignKey, ocdShowQueryWizard, ocdUseFrameset, ocdForceCompatibility, ocdCompatibility, ocdMaxURL, ocdComputeTimeOut, ocdDemoExpires, ocdHomeAddress, ocdTestTableEditing, ocdStartTime, ocdEndTime, ocdGridHighlightSelected, ocdDebug, ocdCodePage, ocdAllowElf, ocdShowAdmin, ocdShowWizard, ocdConnectReport, ocdMotif, ocdCharset, ocdStyleSheet, ocdRunImportEventCode, ocdPageSizeDefault, ocdShowObjectOwner, ocdFormEStringToken, ocdShowPlanReference, ocdConnectURL, ocdDefaultMotif, ocdAuditConn, ocdMaxRecordsRetrieve, ocdMaxRecordsDisplay, ocdServerScriptTimeout, ocdShowDescription, ocdSchemaHideObjects, ocdCustomEditPages, ocdUseRegExKeywordSearch, ocdDBAuthenticate, ocdRequireSSL, ocdShowCompatibility, ocdConnectLogo, ocdShowKeywordSearch, ocdMaxRelatedValues, ocdBrowseAfterSave, ocdBrowseAfterCancel, ocdShowSQLText, ocdRenderAsHTML, ocdGridIcons, ocdWrapGrid, ocdExportLineBreaks

Call SetGlobalDefaults()

%>
<!--#INCLUDE FILE=Config.asp-->
<%

'Attempt Request audit, if ocdAuditConn string is present
Call LogIt()
	
'Declare global system variables
Dim ocdUseSQLRPC, ocdPageName, ocdSessionLCID, ocdIsDemo, ocdDatabaseType, ocdIsHome, ocdIsODBC, ocdDBMSVersion, ocdQSForNoCookie, ocdAppVersion, ocdTargetConn, ndnscSQLConnect, ndnscSQLUser, ndnscSQLPass, ndnscCompatibility

If Not ocdDebug Then
	On Error Resume Next
End If

Call init()

' End Include Logic

Sub init()

	If Not ocdDebug Then
		On Error Resume Next
	
	End If

	If ocdRequireSSL Then
		If UCase(Request.ServerVariables("HTTPS")) <> "ON" Then
			Response.Redirect("https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME") & "?" & Server.URLEncode(Request.QueryString))
		End If
	End If
	ocdAppVersion = "4.303467"
	If ocdHomeAddress = "" Then
		ocdHomeAddress = "204.2.108.116"
	End If
	Server.ScriptTimeout = ocdServerScriptTimeout
	ocdSecureCodeWiz = True
	ocdShowReports = False
	ocdSessionLCID = 0
	ocdTestTableEditing = False
	ocdIsDemo = False
	ocdUseSQLRPC = True
	ocdAllowProAdmin = ocdShowAdmin
	ocdAllowCodeWiz = ocdShowWizard
	ocdPageName = mid(Request.ServerVariables("PATH_INFO"), instrRev(Request.ServerVariables("PATH_INFO"), "/") + 1)
	If ocdUseCustomEditPages Then
		ocdCustomEditPages = "*"
	End If

	'BEGIN SESSION CHECKS
	If Not CBool(ocdCompatibility And ocdNoCookies) Then
		Select Case CStr(Request.ServerVariables("server_software"))
			Case "Microsoft-IIS/5.0", "Microsoft-IIS/5.1"
				Session.CodePage = ocdCodePage
		End Select
		Session.Timeout = ocdSessionTimeout
		If ocdSessionLCID = 0 Then
			ocdSessionLCID = session.lcid
		Else
			Session.lcid = ocdSessionLCID
		End If
		If ocdAdminPassword <> "" Then		  'Check for session password
			If Session("ocdAdminAuthorized") = "" And UCase(ocdPageName) <> "LOGON.ASP" And UCase(ocdPageName) <> "HELP.ASP" And UCase(ocdPageName) <> "LOGOUT.ASP" And UCase(ocdPageName) <> "LOGON.ASP"Then
				Session("ocdSQLConnect") = ""
				Session("ocdSQLUser") = ""
				Session("ocdSQLPass") = ""
				Session("ocdCompatibility") = ""
				If UCase(ocdPageName) = "CONNECT.ASP" Then
					response.redirect("Logon.asp")
				Else
					response.redirect("Logout.asp")
				End If
			End If
		End If
		If ocdStyleSheet = "" Then
		If UCase(ocdPageName) <> "COMMAND.ASP" And UCase(ocdPageName) <> "WIZARDIMPORT.ASP" Then
		If request.Form("ocdMotif") = "" Then
			If Session("ocdMotif") = "" Then
				ocdMotif = ocdDefaultMotif
			Else
				ocdMotif = Session("ocdMotif")
			End If
		Else
			ocdMotif = Request.Form("ocdMotif")
		End If
		Else
			
			If Session("ocdMotif") = "" Then
				ocdMotif = ocdDefaultMotif
			Else
				ocdMotif = Session("ocdMotif")
			End If		
		End If
		Session("ocdMotif") = ocdMotif
	End If
	End If
	If ocdDBAuthenticate And ocdADOUserName = "" Then
		ndnscSQLConnect = ocdADOConnection
		ndnscSQLUser = Session("ocdSQLUser")
		ndnscSQLPass = Session("ocdSQLPass")
		ndnscCompatibility = Session("ocdCompatibility")

	ElseIf ocdADOConnection <> "" Or CBool(ocdCompatibility And ocdNoCookies) Then
		ndnscSQLConnect = ocdADOConnection
		ndnscSQLUser = ocdADOUsername
		ndnscSQLPass = ocdADOPassword
		ndnscCompatibility = ocdCompatibility
	Else

		ndnscSQLConnect = Session("ocdSQLConnect")
		ndnscSQLUser = Session("ocdSQLUser")
		ndnscSQLPass = Session("ocdSQLPass")
		ndnscCompatibility = Session("ocdCompatibility")
	End If
	'END SESSION CHECKS
	If ocdMotif = "" and ocdStyleSheet = "" Then
		ocdMotif = ocdDefaultMotif
	End If
	If ocdAdminPassword <> "" Then
'		ocdLaunchPage = ""
	End If
	If Request.ServerVariables("LOCAL_ADDR") = ocdHomeAddress Then
		ocdIsHome = True
		ocdDisableTextDriver = True
		
	Else
		ocdIsHome = False
	End If
'	ocdConnectionShortCuts = "provider=Microsoft.Jet.OLEDB.4.0;data source=D:\webs\accesshelp.net\data\sample\timeandbilling.mdb|Time and Billing;;provider=Microsoft.Jet.OLEDB.4.0;data source=D:\webs\accesshelp.net\data\sample\membership.mdb|Membership;;provider=Microsoft.Jet.OLEDB.4.0;data source=D:\webs\accesshelp.net\data\sample\expenses.mdb|Business Expenses;;provider=Microsoft.Jet.OLEDB.4.0;data source=D:\webs\accesshelp.net\data\sample\donations.mdb|Donations"
	If (ocdBrandText = "") Then
		ocdBrandText = "1 Click DB"
	End If
	If (ocdBrandLogo = "") Then
		If ocdIsHome Then
			ocdBrandLogo = "<SPAN CLASS=Information>1 Click DB Pro Online</SPAN>"
		Else
			ocdBrandLogo = ocdBrandLogo & "<SPAN CLASS=Information>1 Click DB Pro Software</SPAN>"
		End If
	End If

	If ocdReadOnly Or Not ocdShowAdmin Then
		ocdShowSQLCommander = False
		ocdAccessTableEdits = False
		ocdSQLTableEdits = False
		ocdAllowCodeWiz = False
	End If
	If Not ocdAllowProAdmin Then
		ocdAccessTableEdits = False
		ocdSQLTableEdits = False
		ocdAllowAdmin = False
		ocdAllowCodeWiz = False
		ocdShowImport = False
	End If
	If Not ocdAllowCodeWiz Then
		ocdProcWizard = False
		ocdProcCodeWizard = False
		ocdAuditWizard = False
	Else
		ocdProcWizard = False
		ocdProcCodeWizard = False
	End If
	If ocdIsHome Then
		ocdLaunchPage = "http://www.standardreporting.net/1ClickDB/view.aspx?_@id=534173"
	End If
	If ocdFooterHTML = "" Then
		ocdFooterHTML =  ocdFooterHTML & " <p> <span class=""Information"">1&nbsp;Click&nbsp;DB&nbsp;Pro&nbsp;Software</span> " & "&nbsp;v" & ocdAppVersion & " - " & (CStr(Now())) & (" @ ") & (Request.ServerVariables("SERVER_NAME") & " </p>") 
	End If
	'Begin Set Response Headers
	Response.charset = ocdcharset
	'Response.CodePage = ocdCodePage
	If UCase(ocdPageName) <> "BROWSE.ASP" Or (UCase(ocdPageName) = "BROWSE.ASP" And Request.QueryString("ocdExportFormat_A") = "") Then
		Response.Expires = 0
		Response.ExpiresAbsolute = Now() - 1
		Call Response.addHeader("pragma", "no-cache")
		Call Response.addHeader("cache-control", "private")
		Response.CacheControl = "no-cache"
	Else 'set longer timeout using default cache control for exports
	End If
	'End Set Response Headers
	ocdQSForNoCookie = ""
	If Not (ocdStyleSheet <> "" And UCase(ocdMotif) <> "SYSTEM") Then
		Select Case UCase(ocdMotif)
			Case "CLASSIC", ""
				ocdStyleSheet = "ocdStyleSheet.css"
			Case "NIGHT"
				ocdStyleSheet = "ocdStyleSheetNight.css"
			Case "AUTUMN"
				ocdStyleSheet = "ocdStyleSheetAutumn.css"
			Case "SOFT BLUE"
				ocdStyleSheet = "ocdStyleSheetSoftBlue.css"
			Case "SYSTEM"
				ocdStyleSheet = ""
		End Select
	End If
	ocdQuotePrefix = """"
	ocdQuoteSuffix = """"
	If ndnscCompatibility = "" Then
		ndnscCompatibility = 0
	End If
	If (CBool(CInt(ocdCompatibility) And ocdNoJavaScript)) Then
		ndnscCompatibility = ocdNoJavaScript
	ElseIf (CBool(CInt(ocdCompatibility) And ocdNoFrames)) And Not (CBool(CInt(ndnscCompatibility) And ocdNoJavaScript)) Then
			ndnscCompatibility = ocdNoFrames
	End If
	If (CBool(CInt(ndnscCompatibility) And ocdNoJavaScript)) Then
		ocdShowDescription = False
	End If
	If ocdUseFrameset Then
		If (Not CBool(CInt(ndnscCompatibility) And ocdNoFrames)) And (Not CBool(CInt(ndnscCompatibility) And ocdNoJavaScript)) Then
		Else
			ocdUseFrameset = False
		End If
	Else
		If Not CBool(CInt(ndnscCompatibility) And ocdNoFrames) Then
			ndnscCompatibility = CInt(ndnscCompatibility) - ocdNoFrames
		End If
	End If
	Select Case UCase(ocdPageName)
		Case "BROWSE.ASP", "GENERATOR.ASP", "MTMCODE.ASP", "SCHEMA.ASP", "STRUCTURE.ASP", "EDIT.ASP", "SELECT.ASP", "COMMAND.ASP", "EXECUTE.ASP", "WIZARDSQLQUERY.ASP", "WIZARDSQLAUDIT.ASP", "MTMHEADING.ASP", "WIZARDASPCODE.ASP", "WIZARDIMPORT.ASP", "DESCRIBEFIELD.ASP"
			'need to know DB Properties
			
			If ndnscSQLConnect = "" or (ndnscSQLUser = "" and ocdDBAuthenticate) Then
				If ocdUseFrameset Or CBool(ocdCompatibility And ocdNoCookies) Then
					response.clear
					response.redirect("Logout.asp")
					response.end
				Else
					response.clear
					response.redirect("Connect.asp")
					response.end
				End If
			End If
			Set ocdTargetConn = server.CreateObject("ADODB.Connection")
	
			If ocdReadOnly Then
				ocdTargetConn.Mode = 1
			Else
				Select Case UCase(ocdPageName)
					Case "SCHEMA.ASP", "BROWSE.ASP"
						ocdTargetConn.Mode = 1 'readonly
					Case "EDIT.ASP", "WIZARDIMPORT.ASP"
						ocdTargetConn.Mode = 0 'readwrite
				End Select
			End If

			Call ocdTargetConn.Open(ndnscSQLConnect, ndnscSQLUser, ndnscSQLPass)
			
			If Err.Number <> 0 Then
				Call WriteHeader("NOCONNECT")
				If ocdUseFrameset Then
					Response.Write("<P>")
				End If
				Call WriteFooter("")
				Response.End()
			End If
			Select Case UCase(ocdTargetConn.Provider)
				Case "ADSDSOOBJECT"
					ocdDatabaseType = "ADSI"
					ocdAccessTableEdits = False
					ocdSQLTableEdits = False
					ocdAllowCodeWiz = False
				Case Else
					Select Case UCase(ocdTargetConn.Properties("DBMS Name"))
						Case "MS SQL SERVER", "MICROSOFT SQL SERVER"
							ocdDatabaseType = "SQLServer"
							If ocdUseAnsiQUotes Then
								ocdQuoteSuffix = """"
								ocdQuotePrefix = """"
							Else
								ocdQuoteSuffix = "]"
								ocdQuotePrefix = "["
							End If
							ocdAccessTableEdits = False
							ocdDBMSVersion = 7							' 
						Case "MYSQL"
							ocdDatabaseType = "MySQL"
							ocdAccessTableEdits = False
							ocdQuoteSuffix = "`"
							ocdQuotePrefix = "`"
							ocdSQLTableEdits = False
							ocdAllowCodeWiz = False
							ocdShowReports = False
						Case "MS JET", "ACCESS"
							ocdDatabaseType = "Access"
							ocdSQLTableEdits = False
							ocdShowReports = False
							ocdQuoteSuffix = "]"
							ocdQuotePrefix = "["
							If ocdTargetConn.Provider <> "MSDASQL.1" Then
								If Instr(UCase(ocdTargetConn.Properties("Extended Properties")),"EXCEL") <> 0 Then
									ocdDatabaseType = "Excel"
									ocdReadOnly = True
								End If
							End If
						Case "TEXT"
							ocdShowReports = False
							ocdAccessTableEdits = False
							ocdSQLTableEdits = False
							If ocdDisableTextDriver Then
								Call WriteHeader("")
								Call WriteFooter("Text Driver Disabled")
								response.end()
							Else
								ocdAllowCodeWiz = False
							End If
						Case Else
							ocdShowReports = False
							ocdAccessTableEdits = False
							ocdSQLTableEdits = False
							If instr(UCase(ocdTargetConn.Properties("DBMS Name")), "ORAC") = 0 Then
								ocdDatabaseType = "Unknown"
								ocdAllowCodeWiz = False
							Else
								ocdDatabaseType = "Oracle"
							End If
					End Select
			End Select

			If ocdTargetConn.Provider = "MSDASQL.1" Then
				ocdIsODBC = True
				ocdAccessTableEdits = False
				ocdShowReports = False
				ocdSQLTableEdits = False
			End If
	End Select
	'Response.Write ocdDatabaseType
End Sub

Sub WriteHeader(ByVal ocdAppStatus)
	Response.Write("<!DOCTYPE html PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">")

	Response.Write(vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF)
	Response.Write("<html>" & vbCRLF & "	<head>" & vbCRLF & "		<meta http-equiv=""Expires"" content=""Thu, 01 Dec 1994 16:00:00 GMT"">" & vbCRLF & "		<meta http-equiv=""Pragma"" content=""no-cache"">" & vbCRLF & "		<meta http-equiv=""Content-Type"" content=""text/html; charset=" & ocdCharSet & """>" & vbCRLF)
	Select Case UCase(ocdPageName) 
		Case "CONNECT.ASP", "LOGON.ASP"
			Response.Write("		<meta http-equiv=""Page-Enter"" content=""RevealTrans(Duration=.1,Transition=12)"">") & vbCRLF
	End Select
	Response.Write("		<title>" & ocdBrandText & " @ http://" & Request.ServerVariables("SERVER_NAME") & "</title>" & vbCRLF)
	Response.Write("<link rel=stylesheet type=""text/css"" href=""" & ocdStyleSheet & """>" & vbCRLF)
	Select Case UCase(ocdPageName)
		Case "CONNECT.ASP"
		Case Else

			If Not CBool(CInt(ndnscCompatibility) And ocdNoJavaScript) Then
				Response.Write("		<script type=""text/javascript"" language=""JavaScript"">" & vbCRLF)
				Response.Write("		<!--" & vbCRLF)
				Response.Write("			function stopError() {" & vbCRLF)
				Response.Write("				return true;" & vbCRLF)
				Response.Write(" 			}" & vbCRLF)
				Response.Write("			window.onError = stopError();" & vbCRLF)
				If ((UCase(ocdPageName) = "BROWSE.ASP" And UCase(Request.QueryString("ocdGridMode_A")) <> "FILTER" And UCase(Request.QueryString("ocdGridMode_A")) <> "SEARCH" And UCase(Request.QueryString("ocdGridMode_A")) <> "PROCESS") Or (UCase(ocdPageName) = "COMMAND.ASP") And Request("sqltext") <> "") Then
					Response.Write("			function FinishLoad() {" & vbCRLF)
					Response.Write("				if (document.all) {" & vbCRLF)
					Response.Write("				document.all.loading.style.visibility = 'hidden';" & vbCRLF)
					Response.Write("				}" & vbCRLF)
					Response.Write("			}" & vbCRLF)
					If Request.QueryString("ndreloadtime") <> "" And Request.QueryString("ndreloadtime") <> "0" And Request.QueryString("ndreloadaction") <> "Stop" Then
						If isnumeric(Request.QueryString("ndreloadtime")) Then
							Response.Write("			setTimeout(""reload()"",")
							Dim tir
							tir = CLng(Request.QueryString("ndreloadtime"))
							If tir < 30 Then
								tir = 30
							End If
							Response.Write(CLng(1000 * tir))
							Response.Write(");" & vbCRLF)
							Response.Write("			function reload() {" & vbCRLF)
							Response.Write("				location.href = location.href;" & vbCRLF)
							Response.Write("			}" & vbCRLF)
						End If
					End If
				End If
				Response.Write("		// -->" & vbCRLF)
				Response.Write("		</script>" & vbCRLF)
			End If
		End Select
		Response.write vbCRLF
		Response.Write("	</head>" & vbCRLF & "	<body")
		If Not CBool(CInt(ndnscCompatibility) And ocdNoJavaScript) Then
			If ((UCase(ocdPageName) = "BROWSE.ASP" And UCase(Request.QueryString("ocdGridMode_A")) <> "FILTER" And UCase(Request.QueryString("ocdGridMode_A")) <> "SEARCH" And UCase(Request.QueryString("ocdGridMode_A")) <> "PROCESS") Or (UCase(ocdPageName) = "COMMAND.ASP") And Request("sqltext") <> "") Then
				Response.Write(" onLoad=""javascript:FinishLoad();""")
				Response.Write(">")
				Response.write vbCRLF
				Response.Write("		<script type=""text/javascript"" LANGUAGE=""JavaScript"">" & vbCRLF)
				Response.Write("		<!--" & vbCRLF)
				Response.Write("			if (document.all) {" & vbCRLF)
				Response.write vbCRLF
				Response.Write("				document.write('<div id=""loading"" onclick=""javascript:FinishLoad();""><table class=""loading""><tr><td align=""center"" valign=""middle"" nowrap><span class=""Information"">")
				If (UCase(ocdPageName) = "COMMAND.ASP") Then
					Response.Write("Processing Commands...")
				Else
					Response.Write("Retrieving Data...")
				End If
				Response.Write("<\/span><\/td><\/tr><\/table><\/div>');" & vbCRLF)
				Response.Write("			}" & vbCRLF)
				Response.Write("		// -->" & vbCRLF)
				Response.Write("		</script>" & vbCRLF)
		Else
				Response.Write(">" & vbCRLF)
		End If
	Else	
		Response.Write(">" & vbCRLF)
	End If
	Response.write "		"
	If Not IsNull(ocdDemoExpires) Then
		Session.lcid = 1033
		If Date > ocdDemoExpires Then
			Session.lcid = ocdSessionlcid
			Call WriteFooter("This demo has expired.  The latest version can be found at <a href=http://1clickdb.com target=_top>1clickdb.com</a>.")
			Response.end()
		Else
			Session.lcid = ocdSessionLCID
		End If
	End If
	If ocdUseFrameset Then
		Select Case UCase(ocdPageName)
			Case "CODEWIZUPLOAD.ASP", "CODEWIZRESET.ASP", "CONNECT.ASP", "LAUNCH.ASP", "LOGON.ASP", "CODEWIZCONNECT.ASP", "DEFAULT.ASP"
				If Request.QueryString("sourcecontext") <> "import" Then
					If ocdShowTopMenu Then
						Call WriteTopMenu()
						Response.Write("<hr>")
					End If
				End If
			Case "SCHEMA.ASP"
				If ocdAppStatus = "NOCONNECT" Then
					Call WriteTopMenu()
					Response.Write("<hr>")
				End If
			Case "DBDESIGNSQLSERVER.ASP", "DBDESIGNMSACCESS.ASP"
				If Request.QueryString("DBCreate") <> "" Or Request.QueryString("DBCompact") <> "" Then
					Call WriteTopMenu()
					Response.Write("<hr>")
				ElseIf ocdShowTopMenu And ndnscSQLConnect = "" Then
					Call WriteTopMenu()
					Response.Write("<hr>")
				End If
		End Select
	Else
		Select Case UCase(ocdPageName)
			Case "HELP.ASP"
			Case "CODEWIZUPLOAD.ASP", "CODEWIZRESET.ASP", "CONNECT.ASP", "LOGON.ASP", "CODEWIZCONNECT.ASP", "DEFAULT.ASP"
				Call WriteTopMenu()
				Response.Write("<hr>")
			Case Else
				Call WriteTopMenu()
				
				Response.Write("<hr>")
				If Request.QueryString("DBCreate") = "" And Request.QueryString("DBCompact") = "" Then
				
					Response.Write("<TABLE BORDER=""0""><TR><TD NOWRAP VALIGN=TOP>")
					If ocdADOConnection = "" Then
						If UCase(ocdPageName) = "SCHEMA.ASP" And UCase(Request.QueryString("SHOW")) = "ADO" Then
							Response.Write("<a href=""Schema.asp?show=ADO""><strong>Database</strong></a> : ")
						ElseIf UCase(ocdPageName) <> "CONNECT.ASP" Then
							Response.Write("<a href=""Schema.asp?show=ADO"">Database</a> : ")
						End If
					Else
						If ocdAdminPassword <> "" Then
							Response.Write("<a href=logout.asp onclick=""javascript:return confirm('Are you sure?');"">Logout</a> : ")
						Else
							Response.Write("<strong>Show: </strong> ")
						End If
					End If
					If Not (ndnscSQLConnect = "" Or UCase(ocdPageName) = "CONNECT.ASP") Then
						Response.Write("<a href=""Schema.asp"">")
						If (UCase(ocdPageName) = "SCHEMA.ASP" And UCase(Request.QueryString("show")) = "TABLES") Or (UCase(ocdPageName) = "SCHEMA.ASP" And UCase(Request.QueryString("show")) = "") Then
							Response.Write("<strong>Tables</strong>")
						Else
							Response.Write("Tables")
						End If
						Response.Write("</a>")
						Response.Write(" - ")
						Response.Write("<a href=""Schema.asp?show=views"">")
						If UCase(ocdPageName) = "SCHEMA.ASP" And UCase(Request.QueryString("show")) = "VIEWS" Then
							Response.Write("<strong>Views</strong>")
						Else
							Response.Write("Views")
						End If
						Response.Write("</a>")
						Select Case ocdDatabaseType
							Case "Access", "SQLServer"
						Response.Write(" - ")
						Response.Write("<a href=""Schema.asp?show=procs"">")
						If UCase(ocdPageName) = "SCHEMA.ASP" And Request.QueryString("show") = "procs" Then
							Response.Write("<strong>Procedures</strong>")
						Else
							Response.Write("Procedures")
						End If
						Response.Write("</a>")
						End Select
						Response.Write(" : <a href=WizardImport.asp>Import</a>")
						If ocdAllowCodeWiz and not ocdReadOnly Then
						Response.Write(" - <a href=""WizardASPCode.asp")
						If CBool(ndnscCompatibility and ocdNoJavaScript) Then
							Response.Write("?objtoshow=Both")
						End If
						Response.write(""">Code</a>")
						End If
						If ocdDatabaseType = "SQLServer" and not ocdReadOnly Then
							Response.Write(" - <a href=WizardSQLAudit.asp>Audit</a>")
						End If
					End If
					Response.Write("</TD>")
					Response.Write("</TR></TABLE><P>")
				End If
		End Select
	End If
	Response.Write(vbCRLF)
End Sub

Sub WriteFooter(ByVal ocdAppStatus)
	Response.Write(" ")
	Select Case ocdAppStatus
		Case ""
			Select Case err.number
				Case 0
				Case Else
					Response.Write(DrawDialogBox("warning", "Warning Processing Request", "<p><img src=""appWarningSmall.gif"" alt=""warning""> <span class=""warning"">" & CStr(Err.Number) & "&nbsp;&nbsp;" & CStr(Err.Description) & "</span></p><p>Use your browser's back button to continue or try submitting your request again.</p>"))
			End Select
		Case Else
			Response.Write(DrawDialogBox("warning", "Warning Processing Request", "<P>" & ocdAppStatus & "</P><P>Use your browser's back button to continue or try submitting your request again."))
	End Select
	Response.Write(ocdFooterHTML)
	Response.Write(vbCRLF & "	</body>" & vbCRLF & "</html>" & vbCRLF)
	on error resume next
	ocdTargetConn.Close
	Set ocdTargetConn = Nothing
	Response.Flush()
	Response.Clear()
	Response.End()
End Sub
Sub WriteTopMenu()
	Dim strATarget, strBTarget
	If UCase(ocdPAgeName) = "MTMHEADING.ASP" Then
		strATarget="_parent"
		strBTarget="text"
	Else
		If ocdUseFrameset Then
			strATarget="_parent"
			strBTarget="_parent"
		Else
			strATarget="_self"
			strBTarget="_self"
		End If
	End If
	Response.Write ("<TABLE BORDER=""0""><TR><TD VALIGN=""MIDDLE"">&nbsp;</td>")
	If ocdLaunchPage <> "" Then
		Response.Write ("<TD VALIGN=""MIDDLE""><a href=""" & ocdLaunchPage & """ target=""" & strATarget & """><IMG SRC=""appStart.gif"" BORDER=""0"" WIDTH=""20"" HEIGHT=""20"" ALT=""Start""></a></TD><TD VALIGN=""MIDDLE""><A CLASS=""Menu"" HREF=""" & ocdLaunchPage & """ target=""" & strATarget & """>Start</a></TD><TD>&nbsp;</td>")
	End If
	If ocdAdminPassword <> "" and UCase(ocdPageName) <> "LOGON.ASP" Then
		Response.Write ("<TD VALIGN=""MIDDLE""><a href=""Logon.asp?action=logout"" target=""" & strATarget & """>")
		If ocdADOConnection = "" Then
			Response.Write("<IMG SRC=""appLogon.gif"" BORDER=""0"" WIDTH=""20"" HEIGHT=""20"" ALT=""Logon"">")
		Else
			Response.Write("<IMG SRC=""appConnect.gif"" BORDER=""0"" WIDTH=""20"" HEIGHT=""20"" ALT=""Logon"">")
		
		End If
		Response.Write("</a></TD><TD VALIGN=""MIDDLE""><A CLASS=""Menu"" HREF=""" & "Logon.asp?action=logout"" target=""" & strATarget & """>Logout</a></TD><TD>&nbsp;</td>")
	End If
	If ocdADOConnection = "" Or (ocdDBAuthenticate And ocdADOUserName = "") Then
		If Not CBOOL(ocdCompatibility AND ocdNoCookies) Then
			If Not (ocdAdminPassword <> "" and Session("ocdAdminAuthorized") = "") Then
				Response.Write ("<TD VALIGN=""MIDDLE""><a href=""Connect.asp"" target=""" & strATarget & """><IMG SRC=""AppConnect.gif"" BORDER=""0"" WIDTH=""20"" HEIGHT=""20"" ALT=""Connect to Database""></a></TD><TD VALIGN=""MIDDLE""><A CLASS=""Menu"" HREF=""Connect.asp"" target=""" & strATarget & """>Connect</a></TD><TD>&nbsp;</TD>")
			End If
		End If
	End If
			Select Case UCase(ocdPageName)
				Case "CONNECT.ASP","CODEWIZCONNECT.ASP","LOGON.ASP"
				Case Else
					If ocdShowSQLCommander and not ocdReadOnly  and not (ocdAdminPassword <> "" and Session("ocdAdminAuthorized") = "") Then
						Response.Write ("<TD VALIGN=""MIDDLE""><a href=""Command.asp"" target=""" & strBTarget & """><IMG SRC=""AppCommand.gif"" BORDER=""0"" WIDTH=""20"" HEIGHT=""20"" ALT=""Run SQL Command""></a></TD><TD VALIGN=""MIDDLE""><A CLASS=""Menu"" HREF=""Command.asp"" target=""" & strBTarget & """>Command</a></TD><TD>&nbsp;</TD>")
					End If
					If ocdShowSQLSelector and not (ocdAdminPassword <> "" and Session("ocdAdminAuthorized") = "") and UCase(ocdPageName) <> "LAUNCH.ASP" Then
						Response.Write ("<TD VALIGN=""MIDDLE""><a href=""Select.asp")
							Select Case ocdDatabaseType
								Case "Access","SQLServer","Oracle"
									If Not CBool(ndnscCompatibility and ocdNoJavaScript) Then
										Response.Write ("?ocdStartQueryWizard=yes")
									End If
							End Select
						Response.Write (""" target=""" & strBTarget & """><IMG SRC=""AppSelect.gif"" BORDER=""0"" WIDTH=""20"" HEIGHT=""20"" ALT=""Select Data""></a></TD><TD VALIGN=""MIDDLE""><A CLASS=""Menu"" HREF=""Select.asp")
							Select Case ocdDatabaseType
								Case "Access","SQLServer","Oracle"
									If Not CBool(ndnscCompatibility and ocdNoJavaScript) Then
										Response.Write ("?ocdStartQueryWizard=yes")
									End If
							End Select
						Response.Write (""" target=""" & strBTarget & """>Select</a></TD><TD>&nbsp;</TD>")
					End If
	End Select
	If ocdShowHelp Then
		Response.Write ("<TD VALIGN=""MIDDLE""><a href=""Help.asp""  style=""cursor:help"" target=""_blank""><IMG SRC=""AppHelp.gif"" BORDER=""0"" WIDTH=""20"" HEIGHT=""20"" ALT=""Help""></a></TD><TD VALIGN=""MIDDLE""><A CLASS=""Menu"" HREF=""Help.asp?"" style=""cursor:help"" target=""_blank"">Help</a></TD>")
	End If
	Response.Write ("<TD WIDTH=""100%"" NOWRAP> &nbsp; </TD><TD ALIGN=""RIGHT""  NOWRAP>" & ocdBrandLogo & "</TD></TR></TABLE>")
End Sub
Sub LogIt()
'	on error resume next
	If ocdAuditConn = "" Then
		exit sub
	End If
	Dim connAudit, strAuditSQL, blnLogForm, strFRM, eleF
	blnLogForm = True
	Select Case Mid(Request.ServerVariables("PATH_INFO"), InStrRev(Request.ServerVariables("PATH_INFO"), "/") + 1)
		Case "COMMAND.ASP"
			If Request.QueryString("loadit") <> "" Then
				blnLogForm = False
			End If
		Case "WIZARDIMPORT.ASP"
			If Request.QueryString("loadimportspec") <> "" Then
				blnLogForm = False
			End If
	End Select
	If blnLogForm Then
		strAuditSQL = "INSERT INTO [Request] ([ServerName],[ScriptName],[RemoteAddress],[SessionID],[QueryString],[FormVariables]) VALUES ('" & Replace(Request.ServerVariables("SERVER_NAME"),"'","''") & "','" & Replace(Request.ServerVariables("SCRIPT_NAME"),"'","''") & "','" & Replace(Request.ServerVariables("REMOTE_ADDR"),"'","''") & "','" & Session.SessionID & "','" & Replace(Request.QueryString,"'","''") & "','" & Replace(Request.Form,"'","''") & "')"
	Else
		strAuditSQL = "INSERT INTO [Request] ([ServerName],[ScriptName],[RemoteAddress],[SessionID],[QueryString]) VALUES ('" & Replace(Request.ServerVariables("SERVER_NAME"),"'","''") & "','" & Replace(Request.ServerVariables("SCRIPT_NAME"),"'","''") & "','" & Replace(Request.ServerVariables("REMOTE_ADDR"),"'","''") & "','" & Session.SessionID & "','" & Replace(Request.QueryString,"'","''") & "')"
	End If
	Set connAudit = Server.CreateObject("ADODB.Connection")
	connAudit.Open ocdAuditConn
	connAudit.Execute strAuditSQL
	connAudit.Close
	set connAudit = nothing
	Err.Clear()
End Sub
Function DrawDialogBox (strType, strCaption, strInfo)
	Dim strTemp
	strTemp = "<P ALIGN=CENTER><TABLE "
'	Select Case UCase(strType)
'		Case "WARNING"
			strTemp = strTemp & " "
'	End Select
	strTemp = strTemp & "CLASS=""DialogBox""><TR><TH Class=DialogBoxHeader NOWRAP ALIGN=LEFT>" & strCaption & " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </TH><TR CLASS=DialogBoxRow><TD VALIGN=TOP>"
	select Case UCase(strType)
		Case "WARNING"
			strTemp = strTemp & strInfo
			
		strTemp = strTemp &"</TD></TR></TABLE> "
		Case "DIALOG_START"
		Case "DIALOG_END"
			strTemp = "</TD></TR></TABLE> "
		Case Else
			strTemp = strTemp & strInfo
			strTemp = strTemp & "</TD></TR></TABLE> "
	End Select
	DrawDialogBox = strTemp
End Function
Sub SetGlobalDefaults()
	ocdServerScriptTimeout = 120
	ocdHomeAddress = ""
	ocdShowPlanReference = ""
	ocdShowObjectOwner = True
	ocdFormNullToken = ""
	ocdFormEStringToken  = """"""
	ocdShowImport = False
	ocdShowReports = True
	ocdConnectionShortcuts = ""
	ocdCodePage = 1252
	ocdAllowElf = False
	ocdShowAdmin = True
	ocdShowWizard = True
	ocdConnectReport = ""
	ocdMotif = ""
	ocdWrapGrid = False
	ocdRunImportEventCode = False
	ocdPageSizeDefault = 10
	ocdComputeTimeout = 1
	ocdShowDefaults = True
	ocdHideAutonumber = False
	ocdDefaultTextCompare = "="
	ocdAllowBrowseRefresh = True
	ocdAccessTableEdits = True 
	ocdMultipleFieldSort = True
	ocdFormEmptyStringIsNull = True
	ocdUseCustomEditPages = False
	ocdCustomEditPages = ""
	ocdSessionTimeout = 10
	ocdJETSQLReference = ""
	ocdMSSQLReference = ""
	ocdOraSQLReference = ""
	ocdAdminPassword = ""
	ocdDisableTextDriver = False
	ocdAllowAdmin = True
	ocdReadOnly = False
	ocdUseAnsiQuotes = True
	ocdAllowExport = True
	ocdShowTopFrame = True
	ocdConnectWizardOnly = True
	ocdMaxURLLength = 2000
	ocdQueryWizardIsDefault = True
	ocdForceExportDownload = True
	ocdShowCheckedSearchFields = True
	ocdSecureCodeWiz = True
	ocdDBTimeout = 30
	ocdAllowBrand = False
	ocdAuditWizardPrefix = "audit_"
	ocdBrandText = ""
	ocdBrandLogo = ""
	ocdProcWizard = True
	ocdProcCodeWizard = True
	ocdAuditWizard = true
	ocdLaunchPage = ""
	ocdRootDirCodeWiz = ""
	ocdSQLTableEdits = True
	ocdShowSQLExecutor = True
	ocdShowSQLCommander = True
	ocdShowSQLConnector = True
	ocdShowSQLSelector = True
	ocdShowGraph=False
	ocdSelectForeignKey = True
	ocdShowRelatedRecords = True
	ocdADOPassword = ""
	ocdADOUsername = ""
	ocdADOConnection = ""
	ocdFooterHTML = ""
	ocdHeaderHTML = ""
	ocdMaxRecordsRetrieve = 10000
	ocdMaxRecordsDisplay = 1000
	ocdShowTableMenu = True
	ocdShowHelp = True
	ocdShowDescription = True
	ocdCharSet = "iso-8859-1"
	ocdShowSchema = True
	ocdStyleSheet = "ocdStyleSheet.css"
	ocdMaxURL = 2000
	ocdShowKeywordSearch = True
	ocdUseRegExKeywordSearch = True
	ocdSchemaHideObjects = ""
	ocdCompatibility = 1 'ocdNoFrames
	ocdShowTopMenu = True
	ocdUseFrameset = False
	ocdShowQueryWizard = False
	ocdForceCompatibility = 0
	ocdDefaultMotif = "Classic"
	ocdDemoExpires = null
	ocdRequireSSL = False
	ocdDBAuthenticate = False
	ocdConnectLogo = ""
	ocdStyleSheet = ""
	ocdMaxRelatedValues = 1000
	ocdShowCompatibility = True
	ocdShowSQLSelector = True	'Enable ability to edit and run any SQL Select Statements; Default=true
	ocdShowTableMenu = True	'Show dropdown menu list of tables; Default=true; only works If tree menu is not active (compatibility level > 0)
	ocdShowHelp = True	'Show program help; Default=true
	ocdShowSchema = True	'Enable schema information and table list; Default=true
	ocdShowSQLExecutor = True
	ocdUseFrameset = True 'Use Compatibility switches for flexibility
	ocdReadOnly = False 'If true then allow only viewing but no modification of Tables and Views in the database
	ocdAllowProAdmin = True 'If true then activate the graphic interface for the Command screen, database design features, and stored procedure support
	ocdAllowCodeWiz =  True 'If true then activate all application and scripting support, only active for Wizard edition
	ocdConnectURL = "Connect.asp"
	ocdUseCustomEditPages = False
	ocdBrowseAfterSave = True
	ocdBrowseAfterCancel = True
	ocdShowSQLText = True
	ocdRenderAsHTML = False
	ocdGridIcons = True
	ocdDebug = False
	ocdExportLineBreaks = True
End Sub

%>
