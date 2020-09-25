<%@ LANGUAGE="VBScript.Encode"%>
<% ' Except for @ Directives, there should be no ASP or HTML codes above this line
' Setting LANGUAGE="VBScript.Encode" enables mixing of encoded And unencoded ASP scripts

'1 Click DB Pro copyright 1997-2004 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com/

    
'**Start Encode**

%>
<!--#INCLUDE FILE="PageInit.asp"-->
<%
Dim strAC, strAU, strAP, strCM
'--------------------
'Begin Page_Load
'--------------------

Call ProcessLogin()

'--------------------
'End Page_Load
'--------------------

'--------------------
'Begin Page_Render
'--------------------

Call WriteHeader("")

Call DisplayLogin()

Call WriteFooter("")

Response.End()

'--------------------
'End Page_Render 
'--------------------

'--------------------
'Begin Procedures
'--------------------

Sub SetConnectionInfo(ByVal strAction, ByVal strConnect, ByVal strUser, ByVal strPass, ByVal strCompatibility)
	'Helper function to set session variables for dynamic Database Connection

	Dim strSourceContext
	Select Case UCase(Request.QueryString("sourcecontext"))
		Case "IMPORT"
			strSourceContext = "import"
		Case Else
			strSourceContext = ""
	End Select
	Select Case UCase(strAction)
		Case "CONNECT"
			Select Case strSourceContext
				Case "import"
					Session("ocdSQLConnectImport") = strConnect
					Session("ocdSQLUserImport") = strUser
					Session("ocdSQLPassImport") = strPass
					Response.Clear()
					Response.Redirect("WizardImport.asp")
					Response.End()
				Case Else
					Session("ocdSQLConnect") = strConnect
					Session("ocdSQLUser") = strUser
					Session("ocdSQLPass") = strPass
					Session("ocdCompatibility") = strCompatibility
					Response.Clear()
					Response.Redirect("Schema.asp")
					Response.End()
			End Select
		Case "DISCONNECT"
			Select Case strSourceContext
				Case "import"
					Session("ocdSQLConnectImport") = ""
					Session("ocdSQLUserImport") = ""
					Session("ocdSQLPassImport") = ""
					Response.Clear()
					Response.Redirect("WizardImport.asp")
					Response.End()
				Case Else
					Session("ocdSQLConnect") = ""
					Session("ocdSQLUser") = ""
					Session("ocdSQLPass") = ""
					Session("ocdCompatibility") = ""
					Response.Clear()
					Response.Redirect("Connect.asp?nocache=" & Server.URLEncode(CStr(Now)))
					Response.End()
			End Select
	End Select
End Sub

Sub ProcessLogin()
	Dim strFilePath, strConnectReturn, strDisConnectReturn, strSourceContext, connDBList, strConnDBList, strConnectCaption, rsDBList, rsHasAccess, arrConnSC, arrConnSCName, eleConnSC

	strConnectReturn = "Schema.asp"
	strDisConnectReturn = "Connect.asp?nocache=" & Server.URLEncode(CStr(Now))

	Select Case UCase(Request.QueryString("sourcecontext"))
		Case "IMPORT"
			strSourceContext = "import"
		Case Else
			strSourceContext = ""
	End Select

	If ocdADOConnection <> "" And strSourceContext = "" And Not (ocdDBAuthenticate And ocdADOUserName = "") Then
		Response.Clear()
		Response.Redirect(strConnectReturn)
		Response.End()
	End If

	Select Case strSourceContext
		Case "import"
			strAC = Session("ocdSQLConnectImport")
			strAU = Session("ocdSQLUserImport")
			strAP = Session("ocdSQLPassImport")
			strCM = Session("ocdCompatibility")
		Case Else
			strAC = Session("ocdSQLConnect")
			strAU = Session("ocdSQLUser")
			strAP = Session("ocdSQLPass")
			strCM = Session("ocdCompatibility")
	End Select
	Select Case UCase(Request("Action"))
		Case "CONNECT" 
			If Not (ocdDBAuthenticate And ocdADOUserName = "") Then
				Select Case UCase(Request("Datasource"))
					Case "EXCEL", "ACCESS", "TEXT"
						If Request("filepath") <> "" Then
							If Len(Request("filepath"))> 1 Then
								If Left(Request("filepath"), 2) = "//" Then
									strFilePath = Request("filepath")
								ElseIf Left(Request("filepath"), 1) = "/" Then
									strFilePath = Server.MapPath(Request("filepath"))
								Else
									strFilePath = Request("filepath")
								End If
							Else
								If Left(Request("filepath"), 1) = "/" Then
									strFilePath = Server.MapPath(Request("filepath"))
								Else
									strFilePath = Request("filepath")
								End If
							End If
						Else
							strFilePath = Server.MapPath(".")
						End If
				End Select
				Select Case UCase(Request("Datasource"))
					Case "EXCEL"
						Select Case Request("AccessConnectType")
							Case "OLEDB4.0"
								ndnscSQLConnect = "provider=Microsoft.Jet.OLEDB.4.0;data source=" & strFilePath & ";Extended Properties=""" & Request("version") & ";HDR="
								If Request("HDR") <> "" Then
									ndnscSQLConnect = ndnscSQLConnect & "Yes"
								Else
									ndnscSQLConnect = ndnscSQLConnect & "No"
								End If
								ndnscSQLConnect = ndnscSQLConnect & ";"";"
							Case Else
								ndnscSQLConnect = "Provider=MSDASQL;Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & strFilePath & ";"
						End Select
						Select Case Request("AccessConnectType")
							Case "OLEDB4.0", "OLEDB3.51"
								If Request("systemdb") = "" Then
								ElseIf Left(Request("systemdb"), 1) = "/" Then
									ndnscSQLConnect = ndnscSQLConnect & "Jet OLEDB:System Database=" & Server.MapPath(Request("systemdb")) & ";"
								Else
									ndnscSQLConnect = ndnscSQLConnect & "Jet OLEDB:System Database=" & Request("systemdb") & ";"
								End If
								If Request("dbpass") <> "" Then
									ndnscSQLConnect = ndnscSQLConnect & "Jet OLEDB:Database Password=" & Request("dbpass") & ";"
								End If
							Case Else
								If Request("systemdb") = "" Then
								ElseIf Left(Request("systemdb"), 1) = "/" Then
									ndnscSQLConnect = ndnscSQLConnect & "Jet OLEDB:System Database=" & Server.MapPath(Request("systemdb")) & ";"
								Else
									ndnscSQLConnect = ndnscSQLConnect & "Jet OLEDB:System Database=" & Request("systemdb") & ";"
								End If
						End Select
					Case "TEXT"
						ndnscSQLConnect = "Provider=MSDASQL;Driver={Microsoft Text Driver (*.txt; *.csv)};Extensions=" & Request("extensions") & ";Dbq=" & strFilePath & ";"
					Case "SQLSERVER"
						If Request("sqldatabase") <> "" And Request("sqlserver") <> "" Then
							If Request("trusted") <> "" Then
								ndnscSQLuser = ""
								ndnscSQLPass = ""
							Else
								ndnscSQLuser = Request("user")
								ndnscSQLPass = Request("pass")
							End If
							Select Case Request("SQLConnectType")
								Case "SQLOLEDB"
									ndnscSQLConnect = "provider=SQLOLEDB;data source=" & Request("sqlserver") & ";initial catalog=" & Request("sqldatabase") & ";"
									If Request("netlibrary") <> "" Then
										ndnscSQLConnect = ndnscSQLConnect & "Network Library=" & Request("netlibrary") & ";"
									End If
									If Request("trusted") <> "" Then
										ndnscSQLConnect = ndnscSQLConnect & "Integrated Security=SSPI;"
									End If
									If Request("Encrypt") <> "" Then
										ndnscSQLConnect = ndnscSQLConnect & "Encryption for Data=True;"
									End If
								Case Else
									ndnscSQLConnect = "Provider=MSDASQL;driver={sql server};server=" & Request("sqlserver") & ";database=" & Request("sqldatabase") & ";"
									If Request("netlibrary") <> "" Then
										ndnscSQLConnect = ndnscSQLConnect & "Network=" & Request("netlibrary") & ";"
									End If
									If Request("trusted") <> "" Then
										ndnscSQLConnect = ndnscSQLConnect & "Trusted_Connection=yes;"
									End If
									If Request("Encrypt") <> "" Then
										ndnscSQLConnect = ndnscSQLConnect & "Encrypt=YES"
									End If
							End Select
						End If
					Case "ACCESS"
						Select Case Request("AccessConnectType")
							Case "OLEDB4.0"
								ndnscSQLConnect = "provider=Microsoft.Jet.OLEDB.4.0;data source=" & strFilePath & ";"
							Case "OLEDB3.51"
								ndnscSQLConnect = "provider=Microsoft.Jet.OLEDB.3.51;data source=" & strFilePath & ";"
							Case Else
								ndnscSQLConnect = "provider=MSDASQL;driver={Microsoft Access Driver (*.mdb)};dbq=" & strFilePath & ";"
						End Select
						Select Case Request("AccessConnectType")
							Case "OLEDB4.0", "OLEDB3.51"
								If Request("systemdb") = "" Then
								ElseIf Left(Request("systemdb"), 1) = "/" Then
									ndnscSQLConnect = ndnscSQLConnect & "Jet OLEDB:System Database=" & Server.MapPath(Request("systemdb")) & ";"

								Else
									ndnscSQLConnect = ndnscSQLConnect & "Jet OLEDB:System Database=" & Request("systemdb") & ";"
								End If
								If Request("dbpass") <> "" Then
									ndnscSQLConnect = ndnscSQLConnect & "Jet OLEDB:Database Password=" & Request("dbpass") & ";"
								End If
							Case Else
								If Request("systemdb") <> "" Then
									ndnscSQLConnect = ndnscSQLConnect & "SystemDB=" & Request("systemdb") & ";"
								End If
						End Select
					Case "ORACLE"
						If Request("servicename") <> "" Then
							If Request("trusted") <> "" Then
								ndnscSQLuser = ""
								ndnscSQLPass = ""
							Else
								ndnscSQLuser = Request("user")
								ndnscSQLPass = Request("pass")
							End If
							Select Case Request("ORAConnectType")
								Case "ORAOLEDB"
									ndnscSQLConnect = "provider=OraOLEDB.Oracle;data source=" & Request("servicename") & ";"
								Case "MSOLEDB"
									ndnscSQLConnect = "provider=msdaora;data source=" & Request("servicename") & ";"
								Case "MSODBC1"
									ndnscSQLConnect = "provider=MSDASQL;Driver={Microsoft ODBC Driver for Oracle};Connectstring=" & Request("servicename") & ";"
								Case "MSODBC2"
									ndnscSQLConnect = "provider=MSDASQL;Driver={Microsoft ODBC for Oracle};Server=" & Request("servicename") & ";"
								Case "ORAODBC"
									ndnscSQLConnect = "provider=MSDASQL;Driver={Oracle ODBC Driver};Server=" & Request("servicename") & ";"
							End Select
						End If
					Case "MYSQL"
						If Request("servername") <> "" And Request("dbname") <> "" Then
							ndnscSQLConnect = "Provider=MSDASQL;"
							If Request("351driver") <> "" Then
								ndnscSQLConnect = ndnscSQLConnect & "Driver=MySQL ODBC 3.51 Driver;"
							Else
								ndnscSQLConnect = ndnscSQLConnect & "Driver={mySQL};"
							End If
							ndnscSQLConnect = ndnscSQLConnect & "Server=" & Request("servername") & ";Option=" & Request("options") & ";Database=" & Request("dbname") & ";"
						End If
					Case Else
						If Request("ocdSCut") <> "" And ocdCOnnectionSHortCuts <> "" Then
							ndnscSQLConnect = Request("ocdSCut")
						ElseIf Request("connect") <> "" Then
							ndnscSQLConnect = Request("connect")
						End If
				End Select
			Else
				ndnscSQLConnect = ocdADOConnection
			End If
			If ndnscSQLConnect <> "" Then
				ndnscSQLUser = Request("user")
				ndnscSQLPass = Request("pass")
				ndnscCompatibility = 0
				If Request("nosession") <> "" Then
					ndnscCompatibility = ndnscCompatibility + ocdNoCookies
				End If
				If Request("noframes") <> "" Then
					ndnscCompatibility = ndnscCompatibility + ocdNoFrames
				End If
				If Request("nojavascript") <> "" Then
					ndnscCompatibility = ndnscCompatibility + ocdNoJavaScript
				End If
				If Not CBool((ndnscCompatibility And ocdNoCookies)) Then
					Call setConnectionInfo("Connect", ndnscSQLConnect, ndnscSQLUser, ndnscSQLPass, ndnscCompatibility)
				End If
			End If
		Case "DISCONNECT"
			Call setConnectionInfo("Disconnect", ndnscSQLConnect, ndnscSQLUser, ndnscSQLPass, ndnscCompatibility)
	End Select


	End Sub		

	'-------------------------------
	Public Sub DisplayLogin()

	Dim strFilePath, strSourceContext, connDBList, strConnDBList, strConnectCaption, rsDBList, rsHasAccess, arrConnSC, arrConnSCName, eleConnSC
	Select Case UCase(Request.QueryString("datasource"))
		Case "EXCEL"
			strConnectCaption = ("Microsoft Excel")
		Case "SQLSERVER"
			strConnectCaption = ("MS SQL Server")
		Case "ACCESS"
			strConnectCaption = ("Microsoft Access")
		Case "ORACLE"
			strConnectCaption = ("Oracle")
		Case "MYSQL"
			strConnectCaption = ("MySQL")
		Case "TEXT"
			strConnectCaption = ("ODBC Text Driver")
		Case Else
			strConnectCaption = ""
	End Select
	If strSourceContext = "import" Then
		strConnectCaption = strConnectCaption & (" Import")
	End If
	strConnectCaption = (" Connect to ") & strConnectCaption & " Data Source"
	Response.Write(DrawDialogBox("DIALOG_START", strConnectCaption, ""))
	Response.Write("<form method=""post"" action=""")
	Response.Write(Request.ServerVariables("SCRIPT_NAME"))
	Response.Write("?Datasource=")
	Response.Write(Server.URLEncode(Request.QueryString("datasource")))
	Response.Write("&amp;sourcecontext=")
	Response.Write(Server.URLEncode(Request.QueryString("sourcecontext")))
	Response.Write(""">")
	Response.Write("<table class=""DIALOGBOXTABLE"">")
	If ocdConnectLogo <> "" Then
		Response.Write("<tr class=""DialogBoxRow""><td valign=""top"" colspan=""2"">")
		Response.Write(ocdConnectLogo)
		Response.Write("</td></tr>")
	End If
	If Not ocdDBAuthenticate Then
		Select Case UCase(Request.QueryString("datasource"))
			Case "TEXT"
				Response.Write("<tr class=""DialogBoxRow""><td valign=""top""><span class=""FieldName"">Directory:</span></td><td valign=""bottom"" align=""left""><input type=""text"" class=""ConnectInput"" name=""filepath"" value=""")
				Response.Write(""" size=45 maxlength=""255""></td></tr>")
				Response.Write("<tr class=""DialogBoxRow""><td valign=""top""><span class=""FieldName"">Extensions:</span></td><td valign=""bottom"" align=""left""><input type=""text"" class=""ConnectInput"" name=""extensions"" value=""")
				Response.Write(Server.HTMLEncode("asc,csv,tab,txt"))
				Response.Write(""" size=45 maxlength=""255""></td></tr>")
			Case "EXCEL"
				Response.Write("<tr class=""DialogBoxRow""><td valign=""top""><span class=""FieldName"">Excel XLS:</span></td><td valign=""bottom"" align=""left""><input clas=""ConnectInput"" type=""text"" name=""filepath"" value=""")
				Response.Write(""" size=""45"" maxlength=""255"">")
				Response.Write("<input type=""hidden"" name=""AccessConnectType"" value=""OLEDB4.0"">")
				Response.Write("</td></tr>")
				
				
				Response.Write("<tr class=""DialogBoxRow""><td><span class=""FieldName"">Version:</span></td><td><select name=""Version"">")
				Response.Write("<option value=""Excel 3.0""")
				If Request("version") = "Excel 3.0" Then
					Response.Write(" selected")
				End If
				Response.Write(">Excel 3.0</option>")
				Response.Write("<option value=""Excel 4.0""")
				If Request("version") = "Excel 4.0" Then
					Response.Write(" selected")
				End If
				Response.Write(">Excel 4.0</option>")
				Response.Write("<option value=""Excel 5.0""")
				If Request("version") = "Excel 5.0" Then
					Response.Write(" selected")
				End If
				Response.Write(">Excel 5.0</option>")
				Response.Write("<option value=""Excel 8.0""")
				If Request("version") = "Excel 8.0" Or Request("version") = "" Then
					Response.Write(" selected")
				End If
				Response.Write(">Excel 8.0</option>")
				Response.Write("</select>&nbsp;&nbsp;&nbsp;&nbsp;<span class=""FieldName"">Header Row:</span> ")
				Response.Write("<input type=""checkbox"" name=""HDR""")
				If Request("HDR") <> "" Then
					Response.Write(" checked")
				End If
				Response.Write(">")
				Response.Write("</td></tr>")
			Case "SQLSERVER"
				Response.Write("<tr class=""DialogBoxRow""><td valign=""top"" nowrap align=""left""><span class=""FieldName"">SQL&nbsp;Server:</span></td><td valign=""bottom"" align=""left""><input type=""text"" class=""ConnectInput"" name=""sqlserver"" size=""35"" maxlength=""255"" value=""")
				Response.Write(Server.HTMLEncode(Request("SQLServer")))
				Response.Write("""	><br><small>Enter IP, DNS, or Windows Server Name</small>")
				Response.Write("</td></tr><tr class=""DialogBoxRow""><td valign=""top"" align=""left""><span class=""FieldName"">DB Name:</span></td><td valign=""bottom"" align=""left"">")
				If Request("Action") <> "" And Request("sqldatabase") = "" And Request("sqlserver") <> "" Then
					Set connDBList = Server.CreateObject("ADODB.Connection")
					Select Case Request("SQLConnectType")
						Case "SQLOLEDB"
							strConnDBList = "provider=SQLOLEDB;data source=" & Request("sqlserver") & ";"
							If Request("netlibrary") <> "" Then
								strConnDBList = strConnDBList & "Network Library=" & Request("netlibrary") & ";"
							End If
							If Request("trusted") <> "" Then
								strConnDBList = strConnDBList & "Integrated Security=SSPI;"
							End If
							If Request("Encrypt") <> "" Then
								strConnDBList = strConnDBList & "Encryption for Data=True;"
							End If
						Case Else
							strConnDBList = "Provider=MSDASQL;driver={sql server};server=" & Request("sqlserver") & ";"
							If Request("netlibrary") <> "" Then
								strConnDBList = strConnDBList & "Network=" & Request("netlibrary") & ";"
							End If
							If Request("trusted") <> "" Then
								strConnDBList = strConnDBList & "Trusted_Connection=yes;"
							End If
							If Request("Encrypt") <> "" Then
								strConnDBList = strConnDBList & "Encrypt=YES"
							End If
					End Select
					Call connDBList.Open(strConnDBList, Request("user"), Request("pass"))
					If Err.Number <> 0 Then
						Response.Write("<input class=""ConnectInput"" type=""text"" name=""sqldatabase"" size=""35"" maxlength=""255"" value=""")
						Response.Write(Server.HTMLEncode(Err.Description))
						Response.Write(""">")
						Err.Clear
					Else
						Set rsHasAccess = Server.CreateObject("ADODB.Recordset")
						Set rsDBList = connDBList.OpenSchema(1)	 'adSchemaCatalogs
						Response.Write("<select name=""sqldatabase"">")
						Do While Not rsDBList.eof
							'This method to check perms is too slow
							'					Set rsHasAccess = connDBList.execute ("SELECT HAS_DBACCESS('" & rsDBList.Fields("CATALOG_NAME").Value & "')") ', connDBList
							'					if rsHasAccess(0) = 1 THen
							Response.Write("<option value=""")
							Response.Write(Server.HTMLEncode(rsDBList.Fields("CATALOG_NAME").Value))
							Response.Write(""">")
							Response.Write(Server.HTMLEncode(rsDBList.Fields("CATALOG_NAME").Value))
							Response.Write("</option>")
							'					End If
							'					rsHasAccess.Close
							rsDBList.MoveNext()
						Loop
						Response.Write("</select>")
					End If
					connDBList.Close
					Set connDBList = Nothing
					If err <> 0 Then
						Err.clear()
					End If
				Else
					Response.Write("<input type=""text"" name=""sqldatabase"" class=""ConnectInput"" size=""35"" maxlength=""255"" value=""")
					Response.Write(Server.HTMLEncode(Request("sqldatabase")))
					Response.Write("""><br><small>Leave blank to list databases on SQL Server</small>")
				End If
				Response.Write("<input type=""hidden"" name=""SQLConnectType"" value=""SQLOLEDB"">")
				Response.Write("</td></tr>")
				Response.Write("<tr class=""DialogBoxRow""><td valign=""top"" align=""left""><span class=""FieldName"">Network:</span></td><td><select name=""netlibrary""><option value="""" ")
				If Request("netlibrary") = "" Then
					Response.Write("selected")
				End If
				Response.Write(">Default</option><option value=""dbnmpntw"" ")
				If Request("netlibrary") = "dbnmpntw" Then
					Response.Write("selected")
				End If
				Response.Write(">Named Pipes</option><option value=""dbmssocn"" ")
				If Request("netlibrary") = "dbmssocn" Then
					Response.Write("selected")
				End If
				Response.Write(">TCP/IP</option><option value=""dbmsspxn"" ")
				If Request("netlibrary") = "dbmsspxn" Then
					Response.Write("selected")
				End If
				Response.Write(">SPX/IPX</option><option value=""dbmsvinn"" ")
				If Request("netlibrary") = "dbmsvinn" Then
					Response.Write("selected")
				End If
				Response.Write(">Banyan Vines</option><option value=""dbmsrpcn"" ")
				If Request("netlibrary") = "dbmsrpcn" Then
					Response.Write("selected")
				End If
				Response.Write(">Multi-Protocol</option></select>&nbsp;&nbsp;&nbsp;<span class=""FieldName"">Trusted:</span>&nbsp;<input type=""checkbox"" name=""Trusted""")
				If Request("trusted") <> "" Then
					Response.Write(" selected")
				End If
				Response.Write(">&nbsp;&nbsp;&nbsp;<span class=""FieldName"">Encrypt:</span>&nbsp;<input type=""checkbox"" name=""Encrypt""")
				If Request("Encrypt") <> "" Then
					Response.Write(" selected")
				End If
				Response.Write("></td></tr>")
			Case "ACCESS"
				Response.Write("<tr class=""DialogBoxRow""><td valign=""top"" nowrap><span class=""FieldName"">Path&nbsp;to&nbsp;MDB:</span></td><td valign=""bottom"" align=""left""><input type=""text"" name=""filepath"" class=""ConnectInput"" value=""")
				Response.Write(""" size=""40"" maxlength=""255""><br><small>d:\inetpub\data\example.mdb syntax recommended<br>/webdir/example.mdb syntax will map from webroot</small></td></tr><tr class=""DialogBoxRow""><td valign=""top"" nowrap align=""left"" colspan=""2""><hr><small>The following information is required only for secured Access databases : </small></td></tr><tr class=""DialogBoxRow""><td valign=""top"" nowrap align=""left""><span class=""FieldName"">System&nbsp;DB:</span></td><td valign=""bottom"" align=""left""><input class=""ConnectInput"" type=""text"" name=""SystemDB"" size=""40"" maxlength=""255"" value=""""></td></tr>")
				Response.Write("<input type=""hidden"" name=""AccessConnectType"" value=""OLEDB4.0"">")
				Response.Write("<tr class=""DialogBoxRow""><td nowrap valign=""top"" align=""left""><span class=""FieldName"">DB&nbsp;Password:</span></td><td valign=""bottom"" align=""left""><input class=""ConnectInput"" type=Password name=""DBPass"" size=""35"" maxlength=""255"" value=""""></td></tr>")
			Case "ORACLE"
				Response.Write("<tr class=""DialogBoxRow""><td nowrap valign=""top"" align=""left""><span class=""FieldName"">Service&nbsp;Name:</span></td><td valign=""bottom"" align=""left""><input type=""text"" class=""ConnectInput"" name=""servicename"" size=""35"" maxlength=""255"" value=""")
				Response.Write(Server.HTMLEncode(Request("SQLServer")))
				Response.Write("""></td></tr>")
				Response.Write("<input type=""hidden"" name=""ORAConnectType"" value=""MSOLEDB"">")
			Case "MYSQL"
				Response.Write("<tr class=""DialogBoxRow""><td valign=""top"" align=""left""><span class=""FieldName"">Server Name:</span><td valign=""bottom"" align=""left""><input class=""ConnectInput"" type=""text"" name=""servername"" size=""35"" maxlength=""255"" value=""")
				Response.Write(Server.HTMLEncode(Request("servername")))
				Response.Write("""></td></tr>")
				Response.Write("<tr class=""DialogBoxRow""><td valign=""top"" align=""left""><span class=""FieldName"">DB Name:</span></td><td valign=""bottom"" align=""left""><input class=""ConnectInput"" type=""text"" name=""dbname"" size=""35"" maxlength=""255"" value=""")
				Response.Write(Server.HTMLEncode(Request("dbname")))
				Response.Write("""></td></tr>")
				Response.Write("<tr class=""DialogBoxRow""><td valign=""top"" align=""left""><span class=""FieldName"">Options:</span></td><td valign=""bottom"" align=""left"" nowrap><input class=""ConnectInput"" type=""text"" name=""options"" size=""5"" maxlength=""255"" value=""")
				Response.Write(Server.HTMLEncode("651"))
				Response.Write("""> &nbsp; <span class=""FieldName"">Use ODBC 3.51:</span> <input type=""checkbox"" name=""351Driver""></td></tr>")
			Case Else
				Response.Write("<tr class=""DialogBoxRow""><td valign=""top""><span class=""FieldName"">Connect Wizards:</span></td><td>")
				Response.Write(" (<a href=""Connect.asp?datasource=sqlserver&amp;sourcecontext=")
				Response.Write(Server.URLEncode(Request.QueryString("sourcecontext")))
				Response.Write(""">SQL Server</a>) (<a href=""Connect.asp?datasource=access&amp;sourcecontext=")
				Response.Write(Server.URLEncode(Request.QueryString("sourcecontext")))
				Response.Write(""">MS Access</a>) <br> (<a href=""Connect.asp?datasource=oracle&amp;sourcecontext=")
				Response.Write(Server.URLEncode(Request.QueryString("sourcecontext")))
				Response.Write(""">Oracle</a>) (<a href=""Connect.asp?datasource=mysql&amp;sourcecontext=")
				Response.Write(Server.URLEncode(Request.QueryString("sourcecontext")))
				Response.Write(""">MySQL</a>) (<a href=""Connect.asp?datasource=excel&amp;sourcecontext=")
				Response.Write(Server.URLEncode(Request.QueryString("sourcecontext")))
				Response.Write(""">Excel</a>)")
				If Not ocdDisableTextDriver Then
					Response.Write(" (<a href=""Connect.asp?datasource=text&amp;sourcecontext=")
					Response.Write(Server.URLEncode(Request.QueryString("sourcecontext")))
					Response.Write(""">Text</a>)")
				End If
				Response.Write("</td></tr>")
				If ocdConnectionShortCuts <> "" Then
					Response.Write("<tr class=""DialogBoxRow""><td valign=""top""><span class=""FieldName"">ShortCuts:</span></td><td>")
					arrConnSC = Split(ocdCOnnectionShortCuts, ";;")
					Response.Write("<select name=""ocdSCut"">")
					Response.Write("<option value=""""></option>")
					For Each eleConnSC In arrConnSC
						Response.Write("<option value=""")
						arrConnSCName = Split(eleConnSC, "|")
						Response.Write(Server.HTMLEncode(arrConnSCName(0)))
						Response.Write("""")
						If strAC = (Server.HTMLEncode(arrConnSCName(0))) Then
							Response.Write(" selected")
						End If
						Response.Write(">")
						Response.Write(Server.HTMLEncode(arrConnSCName(1)))
						Response.Write("</option>")
					Next
					Response.Write("</td></tr>")
				End If

				Response.Write("<tr class=""DialogBoxRow""><td valign=""top"" nowrap><span class=""FieldName"">Connect&nbsp;String:</span><br><small>(Can Be DSN)</small><P></td><td valign=""top""><textarea name=""connect"" rows=""2"" cols=""35"">")
				If strAC <> "" Then
					Response.Write(Server.HTMLEncode(strAC))
				End If
				Response.Write("</textarea></td></tr>")
		End Select
	End If
	Response.Write("<tr class=""DialogBoxRow""><td valign=""top"" align=""left""><span class=""FieldName"">User Name:</span></td><td valign=""bottom"" align=""left""><input type=""text"" class=""ConnectInput"" name=""user"" size=""35"" maxlength=""255"" value=""")
	If Request("user") <> "" Then
		Response.Write(Server.HTMLEncode(Request("user")))
	ElseIf Request("datasource") = "" Then
		Response.Write(Server.HTMLEncode(strAU))
	End If
	Response.Write("""></td></tr><tr class=""DialogBoxRow""><td align=""left"" valign=""top""><span class=""FieldName"">Password:</span></td><td align=""left"" valign=""bottom""><input class=""ConnectInput"" type=""password"" name=""pass"" size=""35"" maxlength=""255"" value=""")
	Response.Write("""></td></tr><tr><td nowrap>")
	If strSourceContext = "" And ocdShowCompatibility Then
		If Not CBool(ocdCompatibility And ocdNoFrames) Then
			Response.Write("<span class=""FieldName"">Compatibility: </span></td><td nowrap><input type=""checkbox"" ")
			If CBool(ndnscCompatibility And ocdNoFrames) Then
				Response.Write(" checked")
			End If
			Response.Write(" name=""NoFrames"">No&nbsp;Frames&nbsp;")
			Response.Write("<input type=""checkbox"" name=""NoJavascript"" ")
			If CBool(ndnscCompatibility And ocdNoJavaScript) Then
				Response.Write(" checked ")
			End If
			Response.Write(">No&nbsp;JavaScript&nbsp;")
			Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ")
		End If
	End If
	Response.Write ("</td></tr><tr class=""DialogBoxRow""><td colspan=""2"" valign=""top""><hr><input type=""hidden"" name=""ocdCSSFix""><input type=""submit"" name=""Action"" class=""Submit"" value=""Connect"">")
	If strAC <> "" And Request("datasource") = "" Then
		Response.Write("<input type=""submit"" class=""Submit"" name=""Action"" value=""Disconnect"">")
	End If
	Response.Write("&nbsp;&nbsp;&nbsp;")
	If ocdSQLTableEdits And UCase(Request.QueryString("datasource")) = "SQLSERVER" And strSourceContext = "" Then
		Response.Write("<a href=""DBDesignSQLServer.asp?DBCreate=Yes"" class=""menu""><img src=appNew.gif alt=NEW border=""0"">&nbsp;New&nbsp;Database</a> ")
	End If
	If ocdAccessTableEdits And UCase(Request.QueryString("datasource")) = "ACCESS" Then
		Response.Write("<a href=""DBDesignMSAccess.asp?DBCreate=Yes"" class=""menu""><img src=appnew.gif border=""0"" alt=new>&nbsp;New&nbsp;Database</a>")
	End If
	Response.Write("&nbsp;&nbsp;&nbsp;")

	If strSourceContext = "" And Not ocdRequireSSL Then
		If UCase(Request.ServerVariables("HTTPS")) = "ON" Then
			Response.Write("<a class=""menu"" href=""http://")
			Response.Write(Request.ServerVariables("SERVER_NAME"))
			Response.Write(Request.ServerVariables("SCRIPT_NAME"))
			Response.Write("?")
			Response.Write(Server.URLEncode(Request.QueryString))
			Response.Write("""><img src=""appNoSSL.gif"" border=""0"" alt=""Enable SSL"">&nbsp;Disable&nbsp;SSL&nbsp;Now</a>")
		Else
			Response.Write("<a class=""menu"" href=""https://")
			Response.Write(Request.ServerVariables("SERVER_NAME"))
			Response.Write(Request.ServerVariables("SCRIPT_NAME"))
			Response.Write("?")
			Response.Write(Server.URLEncode(Request.QueryString))
			Response.Write("""><img src=""appSSL.gif"" border=""0"" alt=""Enable SSL"">&nbsp;Enable&nbsp;SSL&nbsp;Now</a> ")
		End If
	End If
	If ocdIsHome Then
		Response.Write("<p><span class=""Warning"">Demo Connection - Not Secure</span></p> ")
	End If
	Response.Write("</td></tr></table>")
	Response.Write("</form> ")
	Response.Write(DrawDialogBox("DIALOG_END", "", ""))

End Sub

'-----------------
'End Procedures
'-----------------

%>