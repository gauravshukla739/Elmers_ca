<% ' Except for @ Directives, there should be no ASP Or HTML codes above this line
' Setting LANGUAGE = VBScript.Encode enables mixing of encoded And unencoded ASP scripts

'1 Click DB Pro copyright 1997-2004 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com/

        
'**Start Encode**

%>
<!--#INCLUDE FILE=PageInit.asp-->
<!--#INCLUDE FILE=ocdFunctions.asp-->
<%

'--------------------
'Begin Page_Load
'--------------------
'on error goto 0
'response.write err.number
'response.Write "T"
Call CheckTreeMenu()
'response.write "u"
'--------------------
'End Page_Load
'--------------------

'--------------------
'Begin Page_Render
'--------------------

Call WriteHeader("")

Call DisplaySchema()

Call WriteFooter("")

Response.End()

'--------------------
'End Page_Render 
'--------------------

'--------------------
'Begin Procedures
'--------------------

Sub CheckTreeMenu()
	If ((Not CBool(ndnscCompatibility And ocdNoFrames)) And (Not CBool(ndnscCompatibility And ocdNoJavaScript))) And Request.QueryString("show") = "" Then
		ocdTargetConn.Close
		Set ocdTargetConn = Nothing
		Err.Clear()
		Response.Clear()
		Response.Redirect("Frameset.asp")
	End If
End Sub

Sub DisplaySchema()
	Dim rsTemp, intCount, QS, ADOprp, srv, argSQLFrom, hideproc, argPCN, nicesqlprocname, arrtmpspnum, rsSchema, cFields, iSchema, brsTm,allrSQL, rsInfo, arrSchemaHideObjects, eleSchemaHideObjects, blnShowObject
	If (Not ocdShowSchema) Then
		Response.End()
	End If
	If (ocdSchemaHideObjects <> "") Then
		arrSchemaHideObjects = Split(ocdSchemaHideObjects,",")
	End If
	
	If (Err.Number <> 0) Then
		Call WriteFooter("")
	End If
	Set rsTemp = Server.CreateObject("ADODB.Recordset")
	If (ocdUseFrameset) Then
		Select Case (UCase(Request.QueryString("show")))
			Case "TABLES",""
				Response.Write("<span class=""Information"">Tables</span><p>")
			Case "VIEWS"
				Response.Write("<span class=""Information"">Views</span><p>")
			Case "PROCS"
				Response.Write("<span class=""Information"">Procedures</span><p>")
			Case "SYS"
				Response.Write("<span class=""information"">System Tables</span><p>")
		End Select
	End If
	Select Case (UCase(Request.QueryString("show")))
		Case "TABLES", "VIEWS","SYS", ""
			If (ocdDatabaseType = "Oracle" And UCase(ocdTargetConn.Properties("DBMS NAME")) <> "MYSQL") Then
				Set rsTemp = ocdTargetConn.Execute ("SELECT OBJECT_TYPE AS TABLE_TYPE, OBJECT_NAME AS TABLE_NAME, OWNER AS TABLE_SCHEMA FROM ALL_OBJECTS WHERE (OBJECT_TYPE = 'TABLE' Or OBJECT_TYPE = 'VIEW') And Not OWNER = 'SYS' And Not OWNER = 'WKSYS' And Not OWNER = 'MDSYS' And Not OWNER = 'OLAPSYS' And Not OWNER ='CTXSYS' And Not OWNER='SYSTEM'")
			Else
				Set rsTemp = ocdTargetConn.OpenSchema(20) 'adSchemaTables
			End If
			If (Err.Number <> 0) Then
				Call WriteFooter("")
			End If
			intCount = 0
			Response.Write("<table class=""Grid"">")
			Response.Write("<tr class=""GridHeader""><th>&nbsp;</th><th align=""left"">Name</th><th>Created</th></tr>")
			Do While (Not rsTemp.EOF)
				If (ocdDatabaseType = "SQLServer" Or ocdDatabaseType = "Oracle") Then
					argSQLFrom = ocdQuotePrefix & rsTemp.FIelds("TABLE_SCHEMA") & ocdQuoteSuffix & "." & ocdQuotePrefix & rsTemp.Fields("TABLE_NAME").Value & ocdQuoteSuffix
				Else
					argSQLFrom = ocdQuotePrefix & rsTemp.Fields("TABLE_NAME").Value & ocdQuoteSuffix
				End If
				blnShowObject = True
				If ((rsTemp.Fields("TABLE_TYPE").Value = "TABLE" And UCase(Left(rsTemp.Fields("TABLE_NAME").Value,4)) <> "MSYS" And ( UCase(Request.QueryString("show")) ="TABLES" Or UCase(Request.QueryString("show")) = "")) Or (rsTemp.Fields("TABLE_TYPE") = "VIEW" And UCase(Request.QueryString("show")) = "VIEWS") Or (UCase(Request.QueryString("show"))="SYS" And (rsTemp.Fields("TABLE_TYPE").Value = "SYSTEM TABLE" Or UCase(Left(rsTemp.Fields("TABLE_NAME").Value,4)) = "MSYS"))) Then
					If (ocdSchemaHideObjects <> "") Then
						For Each eleSchemaHideObjects In arrSchemaHideObjects
							If (eleSchemaHideObjects = argSQLFrom) Then
								blnShowObject = False
								Exit For
							End If
						Next
					End If
				Else
					blnShowObject = False
				End If
				If blnShowObject Then
					Response.Write("<tr")
					If intCount mod 2 = 0 Then
						Response.Write(" class=""GridOdd""")
					Else
						Response.Write(" class=""GridEven""")
					End If
					Response.Write(">")
					Response.Write("<td align=""left"" nowrap>")
					Response.Write("<a href=""Browse.asp?sqlfrom_A=")
					Response.Write(Server.URLEncode(argSQLFrom)) 
					Response.Write("""> ")
					Response.Write("<img border=""0"" src=""AppTable.gif"" alt=""Browse Data""></a>")
					Response.Write(" <a href=""Browse.asp?ocdGridMode_A=Search&amp;sqlfrom_A=" & Server.URLEncode(argSQLFrom))
					Response.Write("""><img src=""AppSearch.gif"" alt=""Search"" border=""0""></a> ")
					Response.Write(" <a href=""Structure.asp?sqlfrom=")
					Response.Write(Server.URLEncode(argSQLFrom))
					Response.Write("""><img src=""MENU_LINK_DEFAULT.GIF"" border=""0"" alt=""Properties""></a>")
					Response.Write("</td>")
					Response.Write("<td nowrap width=""100%"" >")
					Response.Write("<a href=""Browse.asp?sqlfrom_A=")
					Response.Write(Server.URLEncode(argSQLFrom)) 
					Response.Write(""">")
					Response.Write("<span class=""FieldName"">")
					If (Not ocdShowObjectOwner) Then
						Response.Write(Server.HTMLEncode(rsTemp.Fields("TABLE_NAME").Value))
					Else
						Response.Write(Server.HTMLEncode(argSQLFrom))
					End If
					Response.Write("</span></a>")
					Response.Write("</td>")
					Response.Write("<td valign=""top"" nowrap>")
					If (Not ocdDatabaseType = "Oracle") Then
						Response.Write(rsTemp.Fields("DATE_CREATED").Value)
						If (Err.Number <> 0) then ' This is Not always available
							Err.Clear()
							Response.Write("&nbsp;")
						End If
					End If
					Response.Write("</td>")
					Response.Write("</tr>")
					intCount = intCount + 1
				End If
				rsTemp.MoveNext
			Loop
			Response.Write("</table>")
			Response.Flush()
			If (UCase(Request.QueryString("show")) = "TABLES" Or UCase(Request.QueryString("show")) = "") Then
				If (ocdAccessTableEdits And (ocdTargetConn.Properties("DBMS Name") = "MS Jet" Or ocdTargetConn.Properties("DBMS Name") = "ACCESS")) Then
					If (Not CBool(CInt(ndnscCompatibility) And ocdNoCookies)) Then
						Response.Write(" <a href=""DBDesignMSAccess.asp?ocdaction=NEWTABLE"" class=""Menu""><img src=""appNew.gif"" border=""0"" alt=""New"">New Table</a>")
					End If
				End If
			End If	
			If (UCase(Request.QueryString("show")) = "TABLES" Or UCase(Request.QueryString("show")) = "") Then
				If (ocdSQLTableEdits And (ocdTargetConn.Provider = "SQLOLEDB.1")) Then
					If (Not CBool(CInt(ndnscCompatibility) And ocdNoCookies)) Then
						Response.Write("<a href=""DBDesignSQLServer.asp?action=newtable"" class=menu><img src=appNew.gif border=""0"" alt=New> New Table</a>")
					End If
				End If
			End If	
			If (UCase(Request.QueryString("show")) ="VIEWS") Then
				If (ocdAccessTableEdits And (ocdTargetConn.Properties("DBMS Name") = "MS Jet")) Then 
					If (Not CBool(CInt(ndnscCompatibility) And ocdNoCookies)) Then
						Response.Write("<a href=""DBDesignMSAccess.asp?ocdaction=CREATEVIEW"" class=menu><img src=appNew.Gif border=""0"" alt=New> New View</a>")
					End If
				End If
				If (ocdSQLTableEdits And (ocdTargetConn.Provider = "SQLOLEDB.1")) Then
					If (Not CBool(CInt(ndnscCompatibility) And ocdNoCookies)) Then
						Response.Write("<a href=""DBDesignSQLServer.asp?action=createview"" class=""menu""><img src=""appNew.gif"" border=""0""> New View</a>")
					End If
				End If
			End If	
			Response.Write("<p>")
		Case "PROCS"
			Set rsTemp = ocdTargetConn.OpenSchema(16) 'adSchemaProcedures
			intCount = 0
			Response.Write("<table class=""Grid"">")
			Response.Write("<tr class=""GridHeader""><th align=""left"">&nbsp;</th><th align=""left"">Name</th><th>Created</th></tr>")
			Do While (Not rsTemp.EOF)
				hideproc = False
				If (ocdDatabaseType = "Access") Then
					If (Len(rsTemp.Fields("PROCEDURE_NAME").Value) > 3) Then
						If (Left(rsTemp.Fields("PROCEDURE_NAME").Value,3) = "~sq")  Then
							hideproc = True
						End If
					End If
				End If
				If (ocdDatabaseType = "SQLServer") Then
					If (Len(rsTemp.Fields("PROCEDURE_NAME").Value) > 3) Then
						If (Left(rsTemp.Fields("PROCEDURE_NAME").Value,3) = "dt_")  Then
							hideproc = true
						End If
					End If
				End If
				blnShowObject = True
				If (ocdSchemaHideObjects <> "") Then
					For Each eleSchemaHideObjects In arrSchemaHideObjects
						If (eleSchemaHideObjects = FormatForSQL(rsTemp.Fields("PROCEDURE_NAME").Value,ocdDatabaseType,"ADDSQLIDENTIFIER")) Then
							blnShowObject = False
							Exit For
						End If
					Next
				Else
					blnShowObject = True
				End If
				If blnShowObject Then
					If Not hideproc then
						Response.Write("<tr")
						If intCount mod 2 = 0 Then
							Response.Write(" class=""GridOdd"" ")
						Else
							Response.Write(" class=""GridEven"" ")
						End If
						Response.Write(">")
						Response.Write("<td valign=""top"" nowrap>")
						If InStr(CStr(rsTemp.Fields("PROCEDURE_NAME").Value),";") > 0 Then
							arrtmpspnum = Split(Cstr(rsTemp.Fields("PROCEDURE_NAME").Value),";")
							nicesqlprocname = arrtmpspnum(0)
						Else
							nicesqlprocname = Cstr(rsTemp.Fields("PROCEDURE_NAME").Value)
						End If
						If ocdDatabaseType = "SQLServer" Then
							nicesqlprocname = ocdQuotePrefix & rsTemp.Fields("PROCEDURE_SCHEMA").Value & ocdQuoteSuffix & "." & ocdQuotePrefix & nicesqlprocname & ocdQuoteSuffix
						End If
					If Not ocdDatabaseType = "Oracle" Then
						Response.Write("<a href=""Execute.asp?sqlfrom=")
						Response.Write(Server.URLEncode(rsTemp.Fields("PROCEDURE_NAME").Value))
						Response.Write("""><img src=""Appproc.gif"" border=""0"" alt=""Execute Procedure""></a>&nbsp;")

						If ocddatabasetype= "Access" then
							Response.Write("<a href=""DBDesignMSAccess.asp?ocdaction=editproc&amp;sqlfrom=")
						Else
							Response.Write("<a href=""Command.asp?editsp=")
						End If
						Response.Write(Server.URLEncode( nicesqlprocname))  
						Response.Write("""><img src=MENU_LINK_DEFAULT.GIF alt=""Alter Procedure"" border=""0""></a>&nbsp;")
						If ocddatabasetype= "Access" then
	'						Response.Write("<a href=""DBDesignMSAccess.asp?action=dropproc&amp;sqlfrom=")
						Else
							Response.Write("<a href=""DBDesignSQLServer.asp?action=dropproc&amp;sqlfrom=")
						Response.Write(Server.URLEncode(nicesqlprocname))
						Response.Write("""><img src=GRIDLNKDELETE.GIF alt=""Drop Procedure"" border=""0""></a>")
						End If
					End If
					If ocdProcCodeWizard And ocdDatabaseType = "SQLServer" Then
						Response.Write("<a href=WizardProcASP.asp?sqlfrom=")
						Response.Write(Server.URLEncode( rsTemp.Fields("PROCEDURE_NAME").Value))
						For Each QS In Request.QueryString
							If Not UCase(QS) = "SQLFROM" Then
								Response.Write("&")
								Response.Write(QS)
								Response.Write("=")
								Response.Write(Server.URLEncode(Request.QueryString(QS)))
							End If
						Next
						Response.Write(">")
						Response.Write("Code&nbsp;Wizard</a>")
					End If
					Response.Write("</td>")
					Response.Write("<td align=""left"" valign=""top""  width=""100%"" nowrap>&nbsp;")
					If Not ocdDatabaseType = "Oracle" Then
					
						Response.Write("<a href=""Execute.asp?sqlfrom=")
						Response.Write(Server.URLEncode(rsTemp.Fields("PROCEDURE_NAME").Value))
						Response.Write(""">")
					End If
					Response.Write("<span class=""FieldName"">")
					Response.Write(nicesqlprocname)
					Response.Write("</span>")
					If Not ocdDatabaseType = "Oracle" Then
						Response.write "</a>"
					End if
					Response.Write("</td>")
					Response.Write("<td valign=""top"" nowrap>")
					Response.Write(rsTemp.Fields("DATE_CREATED").Value)
					If Err.Number <> 0 Then ' This is Not always available
						Err.clear
						Response.Write("&nbsp;")
					End If
					Response.Write("</td>")
					Response.Write("</tr>")
					End If
					intCount = intCount + 1
				End If
				rsTemp.MoveNext
			Loop
			Response.Write("</table>")
			Response.flush
			If UCase(Request.QueryString("show")) = "PROCS" Then
				If ocdAccessTableEdits And (ocdTargetConn.Properties("DBMS Name") = "MS Jet") Then 
					Response.Write("<p><a class=""Menu"" href=""DBDesignMSAccess.asp?ocdaction=CREATEPROC""><img src=AppNew.Gif alt=New border=""0""> New Procedure</a></p>")
				End If
				If ocdSQLTableEdits And (ocdTargetConn.Provider = "SQLOLEDB.1") Then
					Response.Write("<p><a href=""Command.asp?proposedsqltext=")
					Response.Write(Server.URLEncode("Create Procedure ""StoredProcedure1""" & vbCRLF & "/*" & vbCRLF & "	(" &  vbCRLF & "		@parameter1 datatype = default value," & vbCRLF & "		@parameter2 datatype OUTPUT" & vbCRLF & "	)" & vbCRLF & "*/" & vbCRLF & "As" & vbCRLF & "	/* set nocount on */" & vbCRLF & "	return " & vbCRLF ))
					Response.Write(""" class=menu><img src=appNew.gif border=""0"" alt=New> New Procedure</a></p>")
				End If
			End If	
		Case "SCHEMA"
			Set rsSchema = Server.CreateObject("ADODB.Recordset")
			If Request.QueryString("schemashow") = "" Then
				If ocdUseFrameset Then
					Response.Write("<span class=""information"">ADO Provider Schemas</span><p>")
					Response.Write("ADO data providers are Not required to support all schemas<p>")  
				End If
			Else
				Response.Write("<span class=""information""> ")
				Select Case CInt(Request.QueryString("schemashow"))
					Case -1
						Response.Write("Provider Specific")
					Case 0
						Response.Write("Asserts")
					Case 1
						Response.Write("Catalogs")
					Case 2
						Response.Write("Character Sets")
					Case 3
						Response.Write("Collations")
					Case 4
						Response.Write("Columns")
					Case 5
						Response.Write("Check Constraints")
					Case 6
						Response.Write("Constraint Column Usage")
					Case 7
						Response.Write("Constraint Table Usage")
					Case 8
						Response.Write("Key Column Usage")
					Case 9
						Response.Write("Referential Contraints")
					Case 10
						Response.Write("Table Constraints")
					Case 11
						Response.Write("Columns Domain Usage")
					Case 12
						Response.Write("Indexes")
					Case 13	
						Response.Write("Column Privileges")
					Case 14
						Response.Write("Table Privileges")
					Case 15
						Response.Write("Usage Privileges")
					Case 16
						Response.Write("Procedures")
					Case 17
						Response.Write("Schemata")
					Case 18
						Response.Write("SQL Languages")
					Case 19
						Response.Write("Statistics")
					Case 20
						Response.Write("Tables")
					Case 21
						Response.Write("Translations")
					Case 22
						Response.Write("Provider Types")
					Case 23
						Response.Write("Views")
					Case 24
						Response.Write("View Column Usage")
					Case 25
						Response.Write("View Table Usage")
					Case 26
						Response.Write("Procedure Parameters")
					Case 27
						Response.Write("Foreign Keys")
					Case 28
						Response.Write("Primary Keys")
					Case 29
						Response.Write("Procedure Columns")
				End Select
				Response.Write(" Schema</span><p>")
				Response.Flush()
				Call rsSchema.open (ocdTargetConn.OpenSchema(CInt(Request.QueryString("schemashow"))))
				ocdTargetConn.Close
				set ocdTargetConn = Nothing
				If Err <> 0 Then
					Response.Write("<TABLE><tr><td valign=""top""><img src=""appWarning.gif"" alt=""Warning"" border=""0""></td><TD><span class=""Warning"">Schema Not supported by this ADO Provider</span><p>ADO data providers are Not required to support all schemas</td></tr></table><p>")
				Else
					Response.Write("<table border=1 cellspacing=1 cellpadding=1><tr>")
  	  		'Header with name of fields
					For cFields=0 to rsSchema.Fields.count-1
   					Response.Write("<th bgcolor=silver> " & Trim(rsSchema.Fields(cFields).Name) & "</th>")
		   		Next
 		   		Response.Write("</tr>")
	    		If Not rsSchema.eof then        
     				Response.Write("<tr ><td nowrap>")
       			Response.Write(rsSchema.getstring(,,"</td><td nowrap>","</td></tr><tr><td>","&nbsp;"))
    	   		Response.Write("</td></tr>")
	    		End If
   		 		Response.Write("</table>")
					Response.Write("<p>")
				End If
				Response.Flush()
				Response.Clear()
				rsSchema.Close
				Err.Clear()
			End If
		Case "ADO"
			If ocdUseFrameset then
				Response.Write("<span class=""information""> Database</span><p>")
			End If
			Response.Write("<table border=""0"" CELLSPACING=5 CELLPADDING=5><tr><td valign=""top"" align=""left""><img src=AppDBServer.gif border=""0""></td><td valign=""top"" align=""left""><span class=""information"">" & ndnscSQLConnect & "</span><p>")
			If ocdAllowProAdmin Then
				If ocdDatabaseType="SQLServer" And 1=1 And ocdShowSQLCommander And ocdSQLTableEdits Then
					Response.Write(" <a href=""Command.asp?sqltext=")
					Response.Write(Server.URLEncode("sp_helpserver"))
					Response.Write(""">(Server)</a>&nbsp;&nbsp;&nbsp;")'
					Response.Write(" <a href=""Command.asp?sqltext=")
					Response.Write(Server.URLEncode("sp_helpdb """) & Server.URLEncode(ocdTargetConn.Properties("Current Catalog"))  & Server.URLEncode(""""))
					Response.Write(""">(DB)</a>&nbsp;&nbsp;&nbsp;")'
					Response.Write(" <a href=""Command.asp?sqltext=")
					Response.Write(Server.URLEncode("sp_monitor"))
					Response.Write(""">(Monitor)</a>&nbsp;&nbsp;&nbsp;")
					Response.Write(" <a href=""Command.asp?sqltext=")
					Response.Write(Server.URLEncode("sp_who2"))
					Response.Write(""">(Who)</a>&nbsp;&nbsp;&nbsp;")
					Response.Write(" <a href=""Command.asp?sqltext=")
					Response.Write(Server.URLEncode("select name, password from master..sysxlogins where (pwdcompare( name, master..sysxlogins.password, 0 ) = 1) union select master..sysxlogins.name, null from master..sysxlogins join master..syslogins on master..sysxlogins.sid=master..syslogins.sid where master..sysxlogins.password is null And master..syslogins.isntgroup=0 And master..syslogins.isntuser=0"))
					Response.Write(""">(Bad PWs)</a>&nbsp;&nbsp;&nbsp;")
					Response.Write(" <a href=""Command.asp?sqltext=")
					Response.Write(Server.URLEncode("sp_spaceused"))
					Response.Write(""">(Space)</a>&nbsp;&nbsp;&nbsp;")
					Response.Write(" <a href=""Command.asp?sqltext=")
					Response.Write(Server.URLEncode("DBCC CHECKDB"))
					Response.Write(""">(Check&nbsp;DB)</a>&nbsp;&nbsp;&nbsp;")
					Response.Write(" <a href=""Command.asp?sqltext=")
					Response.Write(Server.URLEncode("DBCC CHECKCATALOG"))
					Response.Write(""">(Check&nbsp;Catalog)</a>&nbsp;&nbsp;&nbsp;")
					Response.Write(" <a href=""Command.asp?sqltext=")
					Response.Write(Server.URLEncode("DBCC CHECKALLOC"))
					Response.Write(""">(Check&nbsp;Alloc)</a>&nbsp;&nbsp;&nbsp;")
					Response.Write(" <a href=""Command.asp?sqltext=")
					Response.Write(Server.URLEncode("sp_helprotect"))
					Response.Write(""">(Permissions)</a>&nbsp;&nbsp;&nbsp;")
					If ocdDBMSVersion > CDbl(6.9) Then
						Response.Write(" <a href=""Command.asp?sqltext=")
						Response.Write(Server.URLEncode("sp_createstats"))
						Response.Write(""">(Create&nbsp;Stats*)</a>&nbsp;&nbsp;&nbsp;")
						Response.Write(" <a href=""Command.asp?sqltext=")
						Response.Write(Server.URLEncode("sp_updatestats"))
						Response.Write(""">(Update&nbsp;Stats)</a>&nbsp;&nbsp;&nbsp;")
						Response.Write(" <a href=""Command.asp?sqltext=")
						Response.Write(Server.URLEncode("SELECT sysobjects.name, sysindexes.rows FROM sysindexes, sysobjects WHERE sysindexes.id = sysobjects.id And indid < 2 And sysobjects.xtype ='U'"))
						Response.Write(""">(Fast&nbsp;Count)</a>&nbsp;&nbsp;&nbsp;")
						Response.Write(" <a href=""Command.asp?sqltext=")
						Response.Write(Server.URLEncode("DBCC SHRINKDATABASE"))
						Response.Write(Server.URLEncode("('") & Server.URLEncode(ocdTargetConn.Properties("Current Catalog")) & Server.URLEncode("')"))
						Response.Write(""">(Shrink&nbsp;DB*)</a>&nbsp;&nbsp;&nbsp;")
						Response.Write(" <a href=""Command.asp?sqltext=")
						Response.Write(Server.URLEncode("master..xp_enum_oledb_providers"))
						Response.Write(""">(Providers*)</a>&nbsp;&nbsp;&nbsp;")
						Response.Write(" <a href=""Command.asp?sqltext=")
						Response.Write(Server.URLEncode("select name, alias, months, shortmonths, days, dateformat, datefirst from master..syslanguages"))
						Response.Write(""">(Languages*)</a>&nbsp;&nbsp;&nbsp;")
						Response.Write("<a href=""Command.asp?sqltext=")
						Response.Write(Server.URLEncode("DBCC PROCCACHE"))
						Response.Write(""">(Show Cache)</a>&nbsp;&nbsp;&nbsp;")
						Response.Write("<a href=""Command.asp?sqltext=")
						Response.Write(Server.URLEncode("DBCC FREEPROCCACHE"))
						Response.Write(""">(Free Cache *)</a>&nbsp;&nbsp;&nbsp;")

						Response.Write(" <a href=""Command.asp?sqltext=")
						Response.Write(Server.URLEncode("SELECT DISTINCT substring(msdb..sysdbmaintplan_history.plan_name,1,40) AS 'plan_name', substring(msdb..sysdbmaintplan_databases.database_name,1,50) as 'database_name',  substring(msdb..sysdbmaintplans.owner,1,15) as owner, msdb..sysdbmaintplans.date_created, substring(msdb..sysdbmaintplan_history.server_name,1,25) as 'server', substring(msdb..sysdbmaintplan_history.activity,1,35) as activity, 'succeeded'=case WHEN msdb..sysdbmaintplan_history.succeeded = 0 THEN 'No' WHEN msdb..sysdbmaintplan_history.succeeded = 1 THEN 'Yes' end, msdb..sysdbmaintplan_history.start_time,  msdb..sysdbmaintplan_history.end_time, msdb..sysdbmaintplan_history.message, msdb..sysdbmaintplan_history.error_number FROM msdb..sysdbmaintplan_history INNER JOIN msdb..sysdbmaintplan_databases ON msdb..sysdbmaintplan_history.plan_id = msdb..sysdbmaintplan_databases.plan_id INNER JOIN msdb..sysdbmaintplans ON msdb..sysdbmaintplan_history.plan_id = msdb..sysdbmaintplans.plan_id"))
						Response.Write(""">(Maintenance&nbsp;Plans*)</a>&nbsp;&nbsp;&nbsp;")
						Response.Write(" <a href=""Command.asp?sqltext=")
						Response.Write(Server.URLEncode("SELECT msdb..sysdtspackages.name, msdb..sysdtspackages.versionid, msdb..sysdtspackages.description, msdb..sysdtspackages.createdate, msdb..sysdtspackages.owner, msdb..sysdtspackagelog.starttime, msdb..sysdtspackagelog.endtime , msdb..sysdtspackagelog.elapsedtime, msdb..sysdtspackagelog.computer FROM msdb..sysdtspackages INNER JOIN msdb..sysdtspackagelog ON msdb..sysdtspackages.id = msdb..sysdtspackagelog.id"))
						Response.Write(""">(Logged&nbsp;DTS&nbsp;Packages*)</a>&nbsp;&nbsp;&nbsp;")
						Response.Write(" <a href=""Command.asp?sqltext=")
						Response.Write(Server.URLEncode("msdb..sp_enum_dtspackages"))
						Response.Write(""">(All&nbsp;DTS&nbsp;Packages*)</a>&nbsp;&nbsp;&nbsp;")
						Response.Write(" <a href=""Command.asp?sqltext=")
						Response.Write(Server.URLEncode("SELECT DISTINCT  SO.Name AS 'Master', S1.name AS 'Detail' FROM dbo.sysforeignkeys FK INNER JOIN dbo.sysobjects SO ON FK.rkeyID = SO.id INNER JOIN dbo.sysobjects S1 ON FK.fkeyID = S1.id"))
						Response.Write(""">(Foreign&nbsp;Keys*)</a>&nbsp;&nbsp;&nbsp;")
						Response.Write(" <a href=""Command.asp?sqltext=")
						Response.Write(Server.URLEncode("SELECT ""name"" AS ""Database"", SUBSTRING(CASE status & 1 WHEN 0 THEN '' ELSE ',autoclose' END +         CASE status & 4 WHEN 0 THEN '' ELSE ',select into/bulk copy' END + CASE status & 8 WHEN 0 THEN '' ELSE ',trunc. log on chkpt' END + CASE status & 16 WHEN 0 THEN '' ELSE ',torn page detection' END + CASE status & 32 WHEN 0 THEN '' ELSE ',loading' END + CASE status & 64 WHEN 0 THEN '' ELSE ',pre-recovery' END + CASE status & 128 WHEN 0 THEN '' ELSE ',recovering' END + CASE status & 256 WHEN 0 THEN '' ELSE ',not recovered' END + CASE status & 512 WHEN 0 THEN '' ELSE ',offline' END + CASE status & 1024 WHEN 0 THEN '' ELSE ',read only' END + CASE status & 2048 WHEN 0 THEN '' ELSE ',dbo USE only' END + CASE status & 4096 WHEN 0 THEN '' ELSE ',single user' END + CASE status & 32768 WHEN 0 THEN '' ELSE ',emergency mode' END + CASE status & 4194304 WHEN 0 THEN '' ELSE ',autoshrink' END + CASE status & 1073741824 WHEN 0 THEN '' ELSE ',cleanly shutdown' END + CASE status2 & 16384 WHEN 0 THEN '' ELSE ',ANSI NULL default' END + CASE status2 & 65536 WHEN 0 THEN '' ELSE ',concat NULL yields NULL' END + CASE status2 & 131072 WHEN 0 THEN '' ELSE ',recursive triggers' END + CASE status2 & 1048576 WHEN 0 THEN '' ELSE ',default TO local cursor' END + CASE status2 & 8388608 WHEN 0 THEN '' ELSE ',quoted identifier' END + CASE status2 & 33554432 WHEN 0 THEN '' ELSE ',cursor CLOSE on commit' END + CASE status2 & 67108864 WHEN 0 THEN '' ELSE ',ANSI NULLs' END + CASE status2 & 268435456 WHEN 0 THEN '' ELSE ',ANSI warnings' END + CASE status2 & 536870912 WHEN 0 THEN '' ELSE ',full text enabled' END, 2,8000) AS ""Options"" FROM master..sysdatabases "))
						Response.Write(""">(All&nbsp;DB&nbsp;Options*)</a>&nbsp;&nbsp;&nbsp;")
						Response.Write(" <a href=""Command.asp?sqltext=")
						Response.Write(Server.URLEncode("SELECT DISTINCT 	substring(msdb..sysjobs.name,1,100) AS [Job], 'Enabled'=case WHEN msdb..sysjobs.enabled = 0 THEN 'No' WHEN msdb..sysjobs.enabled = 1 THEN 'Yes' end, substring(msdb..sysjobschedules.name,1,30) AS [Schedule], 'Frequency'=case WHEN msdb..sysjobschedules.freq_type = 1 THEN 'Once' WHEN msdb..sysjobschedules.freq_type = 4 THEN 'Daily' WHEN msdb..sysjobschedules.freq_type = 8 THEN 'Weekly' WHEN msdb..sysjobschedules.freq_type = 16 THEN 'Monthly' WHEN msdb..sysjobschedules.freq_type = 32 THEN 'Monthly Relative' WHEN msdb..sysjobschedules.freq_type = 32 THEN 'On start SQL Server Agent' END,	'Subday Interval'=case WHEN msdb..sysjobschedules.freq_subday_type = 1 THEN 'Specified time'  WHEN msdb..sysjobschedules.freq_subday_type = 2 THEN 'Seconds' WHEN msdb..sysjobschedules.freq_subday_type = 4 THEN 'Minutes' WHEN msdb..sysjobschedules.freq_subday_type = 8 THEN 'Hours' END, cast(cast(msdb..sysjobschedules.active_start_date as varchar(15)) as datetime) as [Start Date],	 cast(cast(msdb..sysjobschedules.active_end_date as varchar(15)) as datetime) as [End Date],	 cast(cast(msdb..sysjobschedules.next_run_date as varchar(15)) as datetime) as [Next Run Date],	 msdb..sysjobschedules.next_run_time as [Next Run Time], msdb..sysjobschedules.date_created As [Created] FROM msdb..sysjobhistory INNER JOIN msdb..sysjobs ON  msdb..sysjobhistory.job_id = msdb..sysjobs.job_id INNER JOIN msdb..sysjobschedules ON msdb..sysjobs.job_id = msdb..sysjobschedules.job_id"))
						Response.Write(""">(Jobs*)</a>&nbsp;&nbsp;&nbsp;")
						If 1=1 Then 'ocdDBMSVersion > CDbl(7.5) Then
							Response.Write("<a href=""Command.asp?sqltext=")
							Response.Write(Server.URLEncode("USE MASTER" & vbCRLF & "SELECT DISTINCT CONVERT (smallint, l1.req_spid) AS spid, left(db_name(l1.rsc_dbid), 10) AS dbName, left(object_name(l1.rsc_objid), 20) AS ObjName, l1.rsc_indid AS IndId, substring (v.name, 1, 4) AS Type, substring (l1.rsc_text, 1, 16) AS Resource, substring (u.name, 1, 8) AS Mode, substring (x.name, 1, 5) AS Status FROM master.dbo.syslockinfo l1, master.dbo.syslockinfo l2, master.dbo.spt_values v, master.dbo.spt_values x, master.dbo.spt_values u WHERE l1.rsc_type = v.number And v.type = 'LR' And l1.req_status = x.number And x.type = 'LS' And l1.req_mode + 1 = u.number And u.type = 'L' And l1.rsc_type <>2 /* Not a DB lock */ And l1.rsc_dbid = l2.rsc_dbid And l1.rsc_bin = l2.rsc_bin And l1.rsc_objid = l2.rsc_objid And l1.rsc_indid = l2.rsc_indid And l1.req_spid <> l2.req_spid And l1.req_status <> l2.req_status ORDER BY substring (l1.rsc_text, 1, 16), substring (x.name, 1, 5)"))
							Response.Write(""">(Blocking&nbsp;Locks**)</a> &nbsp;&nbsp;&nbsp;")
							Response.Write(" <a href=""Command.asp?sqltext=")
							Response.Write(Server.URLEncode("select * from master..sysprocesses where open_tran > 0"))
							Response.Write(""">(Open&nbsp;Transactions**)</a>&nbsp;&nbsp;&nbsp;")
							Response.Write(" <a href=""Command.asp?sqltext=")
							Response.Write(Server.URLEncode("dbcc sqlperf (logspace)"))
							Response.Write(""">(Log&nbsp;Size**)</a>&nbsp;&nbsp;&nbsp;")
							Response.Write(" <a href=""Command.asp?sqltext=")
							Response.Write(Server.URLEncode("SELECT ServerProperty('Edition') As Edition, ServerProperty('LicenseType') As License, ServerProperty('NumLicenses') As ""#Licenses"", ServerProperty('IsClustered') As ""Clustered"""))
							Response.Write(""">(Server&nbsp;Props**)</a>&nbsp;&nbsp;&nbsp;")
						End If
					End If
				End If
				Response.Write("<a href=""" & ocdPageName & "?show=CGI" & """>(CGI Server Variables)</a>&nbsp;&nbsp;&nbsp;")
				Select Case ocdDatabaseType
					Case "Access", "SQLServer"
						Response.Write("<a href=""" & ocdPageName & "?show=sys"">(System&nbsp;Tables)</a>&nbsp;&nbsp;&nbsp; ")
				End Select
				Response.Write("<a href=""" & ocdPageName & "?show=adop"">(ADO&nbsp;Properties)</a>")
				If ocdDatabaseType = "SQLServer" Then
						Response.write("<p>* SQL Server 7+ required<br>** SQL Server 2000+ required</p>")
				End If
			End If
			Response.Write("</td></tr></table>")
			'**************
		Case "ADOP"
			Response.Write("<span class=""information"">ADO Connection Properties</span>")
			Response.Write("<p>")
			Response.Write("<span class=""FieldName"">Connect String</span> = ")
			Response.Write(Server.HTMLEncode(ndnscSQLConnect))
			Response.Write("<p>")
			'**************
			On Error Resume Next
			If Not ocdUseFrameset Or Request.QueryString("schemashow") = "" Then
				Response.Write("<form action=""" & ocdPageName & """ METHOD=GET>")
				Response.Write("<span class=""FieldName"">ADO Schema:</span> <SELECT NAME=schemashow>")
				Response.Write("<option value=""""></option>")
				For iSchema = 0 To 29
					Response.Write("<option value=""" & iSchema & """>")
					Select Case iSchema
						Case -1
							Response.Write("Provider Specific")
						Case 0
							Response.Write("Asserts")
						Case 1
							Response.Write("Catalogs")
						Case 2
							Response.Write("Character Sets")
						Case 3
							Response.Write("Collations")
						Case 4
							Response.Write("Columns")
						Case 5
							Response.Write("Check Constraints")
						Case 6
							Response.Write("Constraint Column Usage")
						Case 7
							Response.Write("Constraint Table Usage")
						Case 8
							Response.Write("Key Column Usage")
						Case 9
							Response.Write("Referential Contraints")
						Case 10
							Response.Write("Table Constraints")
						Case 11
							Response.Write("Columns Domain Usage")
						Case 12
							Response.Write("Indexes")
						Case 13	
							Response.Write("Column Privileges")
						Case 14
							Response.Write("Table Privileges")
						Case 15
							Response.Write("Usage Privileges")
						Case 16
							Response.Write("Procedures")
						Case 17
							Response.Write("Schemata")
						Case 18
							Response.Write("SQL Languages")
						Case 19
							Response.Write("Statistics")
						Case 20
							Response.Write("Tables")
						Case 21
							Response.Write("Translations")
						Case 22
							Response.Write("Provider Types")
						Case 23
							Response.Write("Views")
						Case 24
							Response.Write("View Column Usage")
						Case 25
							Response.Write("View Table Usage")
						Case 26
							Response.Write("Procedure Parameters")
						Case 27
							Response.Write("Foreign Keys")
						Case 28
							Response.Write("Primary Keys")
						Case 29
							Response.Write("Procedure Columns")
					End Select
					Response.Write("</option>")
				Next
				Response.Write("</select><input name=""ndAction"" type=""Submit"" class=""Submit"" value=""Show"">")
				Response.Write("<input type=""hidden"" name=""show"" value=""Schema"">")
				Response.Write("</form>")
			End If
			Response.Write("<span class=""FieldName"">ADO Version</span> = ")
			Response.Write(ocdTargetConn.version)
			Response.Write("<br>")
			Response.Write("<span class=""FieldName"">ADO Provider = </span>")
			Response.Write(ocdTargetConn.provider)
			Response.Write("<p>")
			For Each ADOprp In ocdTargetConn.Properties
				Response.Write("<span class=""FieldName"">")
				Response.Write(ADOprp.Name)
				Response.Write("</span> = ")
				Response.Write(ADOprp.Value)
				Response.Write("<br>")
			Next
		Case "CGI"
			If ocdUseFrameset Then
				Response.Write("<span class=""information"">CGI Server Variables</span><p>")
				Response.Write("<span class=""FieldName"">Server software</span> = ")
				Response.Write(request.servervariables("server_software"))
				Response.Write("<br>")
				Response.Write("<span class=""FieldName"">Script engine</span> = ")
				Response.Write(scriptengine())
				Response.Write(scriptenginemajorversion())
				Response.Write(".")
				Response.Write(scriptengineminorversion())
				Response.Flush()
				Response.Write("<br>")
				Response.Write("<span class=""FieldName"">Build</span> = ")
				Response.Write(scriptenginebuildversion())
				Response.Write("<p>")
			End If
			For Each srv In Request.ServerVariables
				Response.Write("<span class=""FieldName"">")
				Response.Write(srv)
				Response.Write("</span>")
				Response.Write(" = ")
				Response.Write(Server.HTMLEncode(Request.Servervariables(srv)))
				Response.Write("<br>")
			Next
	End Select
End Sub
%>
