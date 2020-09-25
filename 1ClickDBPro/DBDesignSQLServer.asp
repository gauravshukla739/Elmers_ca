<% ' Except for @ Directives, there should be no ASP Or HTML codes above this line
' Setting LANGUAGE = VBScript.Encode enables mixing of encoded and unencoded ASP scripts

'1 Click DB Pro copyright 1997-2003 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, Or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com
'Email inquiries to info@1ClickDB.com
        
'**Start Encode**

%>
<!--#INCLUDE FILE=PageInit.asp-->
<!--#INCLUDE FILE=ocdFunctions.asp-->
<!--#INCLUDE FILE=ocdManageSQLServer.asp-->
<%
'On Error Goto 0
Dim connNewDB, strStatus, strNewConnectionString, strSQLdbcreate, strWhatIsSQLFrom, strWhatIsRelatedTo, objSQL, strAction, strSQLFrom, connDesign, strAlterTableName, rsCreate, intTableID, intColID, strDCName, rsRelTbl, rsRelTblFld, rsRelFieldCount, rsRelIDX, OpenIdentifier, CloseIdentifier, fmAction, strNDURLConnect, strNDURLSQL, ndTargetConnType, rsdef, arrcrit, strcrit, deftype, rsFSchema, strUDTName, tnname, rsTableNames, strSelectOptions, fldF, rstxxxcn, strFKSQL, strIndSQL, COUNTITXX, fmNewTableKeyName, fmNewTableKeyType, fmNewTableKeySize, fmNewFieldTableName, fmNewFieldName, fmNewFieldType, fmNewFieldSize, qsFieldName, strSQL, rsXXTemp, ndAction, qsAction, strSQLStuff, dmosql, dmodb, strSv, strDB, strON, strOO, strSS, cat2tr, strViewName2tr, strSQLStufftr, rsuser2, tmpRedObj, strOName, rsuser, strUN, strNTName, irrra
	
		'response.end
If Request.QueryString("DBCreate") <> "" Then

	If ocdADOConnection <> "" Or ocdReadOnly Or Not ocdAllowProAdmin Then
		Response.Clear()
		Response.Redirect("Schema.asp")
		Response.end
	End If

	strStatus = ""
	If Request.Form("ndbtnCancel") <> "" Then
		Response.Clear
		Response.Redirect("Connect.asp")
		Resposne.End
	End If
	If Request.Form("Action") <> "" Then
	
		If strStatus = "" Then
			If Request("sqlserver") = "" Then
				strStatus = "Specify an SQL Server"		
			End If	
		End If
		If strStatus = "" Then
			If Request("sqldatabase") = "" Then
				strStatus = "Enter a Database to Create"		
			End If	
		End If
		If strStatus = "" Then
			If Request("sqldatabase") <> "" and Request("sqlserver") <> "" Then
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
							ndnscSQLConnect = ndnscSQLConnect &  "Integrated Security=SSPI;"
						End If
						strNewConnectionString = "provider=SQLOLEDB;data source=" & Request("sqlserver") & ";"
						If Request("netlibrary") <> "" Then
							strNewConnectionString = strNewConnectionString & "Network Library=" & Request("netlibrary") & ";"
						End If
						If Request("trusted") <> "" Then
							strNewConnectionString = strNewConnectionString &  "Integrated Security=SSPI;"
						End If
					Case Else
						ndnscSQLConnect = "driver={sql server};server=" & Request("sqlserver") & ";database=" & Request("sqldatabase") & ";"
						If Request("netlibrary") <> "" Then
							ndnscSQLConnect = ndnscSQLConnect &  "Network=" & Request("netlibrary") & ";"
						End If
						If Request("trusted") <> "" Then
							ndnscSQLConnect = ndnscSQLConnect &  "Trusted_Connection=yes;"
						End If
						strNewConnectionString = "driver={sql server};server=" & Request("sqlserver") & ";"
						If Request("netlibrary") <> "" Then
							strNewConnectionString = strNewConnectionString &  "Network=" & Request("netlibrary") & ";"
						End If
						If Request("trusted") <> "" Then
							strNewConnectionString = strNewConnectionString &  "Trusted_Connection=yes;"
						End If
				End Select
			End If
			Set connNewDB = Server.CreateObject("ADODB.Connection")
		
			Call connNewDB.open (strNewConnectionString, Request("user"), Request("pass"))
			connNewDb.commandtimeout = ocddbtimeout
			If Err.Number <> 0 Then
				strStatus = err.description
				Err.Clear
			End If
		End If
		If strStatus = "" Then
			strSQLDBCreate = "CREATE DATABASE """ & Request("sqldatabase") & """"
			If Request("initsize") <> "" Or Request("filename") <> "" Or Request("maxsize") <> "" Or Request("growsize") <> "" Then
				strSQLDBCreate = strSQLDBCREATE & " ON ("
				If Request("dbfilename") <> "" Then
					strSQLDBCreate = strSQLDBCREATE & "name='" & Request("dbfilename") & "',"
				End If
				If Request("filename") <> "" Then
					strSQLDBCreate = strSQLDBCREATE & "FILEname='" & Request("filename") & "',"
				End If
				If Request("initsize") <> "" Then
					strSQLDBCreate = strSQLDBCREATE & "SIZE=" & Request("initsize") & Request("initsizeunit") & ","
				End	If
				If Request("maxsize") <> "" Then
					strSQLDBCreate = strSQLDBCREATE & "MAXSIZE=" & Request("maxsize") & Request("maxsizeunit") & ","
				End If
				If Request("growsize") <> "" Then
					strSQLDBCreate = strSQLDBCREATE & "FILEGROWTH=" & Request("growsize") & Request("growsizeunit") & ","
				End If
				strSQLDBCreate = left(strSQLDBCreate,len(strSQLDBCreate)-1)
				strSQLDBCreate = strSQLDBCreate & ")"
			End If
			Call connNewDB.execute (strSQLdbcreate)
			If Err.Number <> 0 Then
				strStatus = err.description
				Err.Clear
			Else
				Session("ocdSQLUser") = Request("user")
				Session("ocdSQLPass") = Request("pass")
				Session("ocdSQLConnect") = ndnscSQLConnect
				ndnscSQLuser = Request("user")
				ndnscSQLPass = Request("pass")
				Response.Clear()
				Response.Redirect("Schema.asp")
				Response.end()
			End If
		End If
	End If
'	response.Write ("XXX")
	'Start Create DB Form
	Call WriteHeader("")
	Response.Write("<center><form method=post action=""")
	Response.Write(Request.ServerVariables("SCRIPT_NAME"))
	Response.Write("?DBCreate=Yes"">")
	Response.Write DrawDialogBox("DIALOG_START","Create New SQL Server Database","")
	If Request.Form <> "" Then
		If strStatus <> "" Then
			Response.Write("<table><tr><td><img src=""appWarning.gif"" alt=""Warning""></td><td><span class=PageStatus>")
			Response.Write(strStatus)
			Response.Write("</span></td></tr></table>")
		End If
		Response.Write("<p>")
	End If
	Response.Write("<table cellspacing=3 cellpadding=3 class=DialogInterior><tr><td valign=top align=left><strong> SQL Server:</strong></td><td valign=bottom align=left><input type=Text name=sqlserver size=35 maxlength=255 value=""")
	Response.Write(Server.HTMLEncode(Request("SQLServer")))
	Response.Write("""	></td></tr>")
	Response.Write("<tr><td valign=top align=left><strong>New DB Name:</strong></td><td valign=bottom align=left><input type=Text name=sqldatabase size=35 maxlength=255 value=""")
	Response.Write(Server.HTMLEncode(Request("sqldatabase")))
	Response.Write("""></td></tr>")
	Response.Write("<tr><td valign=top align=left><strong>DB Driver:</strong></td><td align=left valign=bottom><input type=Radio name=SQLConnectType value=""SQLOLEDB"" ")
	If Request("SQLConnectType") = "SQLOLEDB"  Or Request("SQLConnectType") = "" Then
		Response.Write("checked")
	End If
	Response.Write(">SQLOLEDB&nbsp;&nbsp;&nbsp;SQL&nbsp;Server&nbsp;7.0+&nbsp;only</td></tr><tr><td valign=top align=left><strong>Network:</strong></td><td><select name=netlibrary><option value="""" ")
	If Request("netlibrary") = "" Then
		Response.Write "selected"
	End If
	Response.Write ">Default</option><option value=""dbnmpntw"" "
	If Request("netlibrary") = "dbnmpntw" Then
		Response.Write "selected"
	End If
	Response.Write ">Named Pipes</option><option value=""dbmssocn"" "
	If Request("netlibrary") = "dbmssocn" Then
		Response.Write "selected"
	End If
	Response.Write ">TCP/IP</option><option value=""dbmsspxn"" "
	If Request("netlibrary") = "dbmsspxn" Then
		Response.Write "selected"
	End If
	Response.Write ">SPX/IPX</option><option value=""dbmsvinn"" "
	If Request("netlibrary") = "dbmsvinn" Then
		Response.Write "selected"
	End If
	Response.Write ">Banyan Vines</option><option value=""dbmsrpcn"" "
	If Request("netlibrary") = "dbmsrpcn" Then
		Response.Write "selected"
	End If
	Response.Write ">Multi-Protocol</option></select>&nbsp;&nbsp;&nbsp;<strong>Trusted:</strong>&nbsp;<input type=CHECKBOX name=Trusted "
	If Request.Form("trusted") <> "" Then
		Response.Write "selected"
	End If
	Response.Write "> </td></tr>"
	Response.Write("<tr><td valign=top  align=left><strong>User Name:</strong></td><td valign=bottom align=left><input type=text name=user size=35 maxlength=255 value=""")
	Response.Write(Server.HTMLEncode(Request("User")))
	Response.Write("""></td></tr><tr><td align=left valign=top><strong>Password:</strong></td><td align=left valign=bottom><input type=Password name=pass size=35 maxlength=255 value=""")
	Response.Write(Server.HTMLEncode(Request("Pass")))
	Response.Write("""></td></tr>")
	Response.Write("<tr><td colspan=2 valign=top><p><input type=submit name=Action class=submit value=""Create Database""> Optional Parameters Below</td></tr>")
	Response.Write("<tr><td valign=top align=left><strong>File Name:</strong></td><td valign=bottom align=left><input type=Text name=dbfilename size=42 maxlength=255 value=""")
	Response.Write(Server.HTMLEncode(Request("dbfilename")))
	Response.Write("""></td></tr>")
	Response.Write("<tr><td valign=top align=left><strong>OS File Name:</strong></td><td valign=bottom align=left><input type=Text name=filename size=42 maxlength=255 value=""")
	Response.Write(Server.HTMLEncode(Request("filename")))
	Response.Write("""></td></tr>")
	Response.Write("<tr><td valign=top align=left><strong>Initial Size:</strong></td><td valign=bottom align=left><input type=Text name=initsize size=10 maxlength=255 value=""")
	Response.Write(Server.HTMLEncode(Request("initsize")))
	Response.Write(""">")
	Response.Write("<select name=initsizeunit>")
	Response.Write("<option value=""KB""")
	If Request("initsizeunit") = "KB" Then
		Response.Write(" selected")
	End If
	Response.Write(">KB</option>")
	Response.Write("<option value=""MB""")
	If Request("initsizeunit") = "" Or Request("initsizeunit") = "MB" Then
		Response.Write(" selected")
	End If
	Response.Write(">MB</option>")
	Response.Write("<option value=""GB""")
	If Request("initsizeunit") = "GB" Then
		Response.Write(" selected")
	End If
	Response.Write(">GB</option>")
	Response.Write("<option value=""TB""")
	If Request("initsizeunit") = "TB" Then
		Response.Write(" selected")
	End If
	Response.Write(">TB</option>")
	Response.Write("</select>")
	Response.Write("</td></tr>")
	Response.Write("<tr><td valign=top align=left><strong>Maximum Size:</strong></td><td valign=bottom align=left><input type=Text name=maxsize size=10 maxlength=255 value=""")
	Response.Write(Server.HTMLEncode(Request("maxsize")))
	Response.Write(""">")
	Response.Write("<select name=maxsizeunit>")
	Response.Write("<option value=""KB""")
	If Request("maxsizeunit") = "KB" Then
		Response.Write(" selected")
	End If
	Response.Write(">KB</option>")
	Response.Write("<option value=""MB""")
	If Request("maxsizeunit") = "" Or Request("maxsizeunit") = "MB" Then
		Response.Write(" selected")
	End If
	Response.Write(">MB</option>")
	Response.Write("<option value=""GB""")
	If Request("maxsizeunit") = "GB" Then
		Response.Write(" selected")
	End If
	Response.Write(">GB</option>")
	Response.Write("<option value=""TB""")
	If Request("maxsizeunit") = "TB" Then
		Response.Write(" selected")
	End If
	Response.Write(">TB</option>")
	Response.Write("</select>")
	Response.Write("</td></tr>")
	Response.Write("<tr><td valign=top align=left><strong>File Growth:</strong></td><td valign=bottom align=left><input type=Text name=growsize size=10 maxlength=255 value=""")
	Response.Write(Server.HTMLEncode(Request("growsize")))
	Response.Write(""">")
	Response.Write("<select name=growsizeunit>")
	Response.Write("<option value=""KB""")
	If Request("growsizeunit") = "KB" Then
		Response.Write(" selected")
	End If
	Response.Write(">KB</option>")
	Response.Write("<option value=""MB""")
	If Request("growsizeunit") = "" Or Request("growsizeunit") = "MB" Then
		Response.Write(" selected")
	End If
	Response.Write(">MB</option>")
	Response.Write("<option value=""GB""")
	If Request("growsizeunit") = "GB" Then
		Response.Write(" selected")
	End If
	Response.Write(">GB</option>")
	Response.Write("<option value=""TB""")
	If Request("growsizeunit") = "TB" Then
		Response.Write(" selected")
	End If
	Response.Write(">TB</option>")
	Response.Write("<option value=""%""")
	If Request("growsizeunit") = "%" Then
		Response.Write(" selected")
	End If
	Response.Write(">%</option>")
	Response.Write("</select>")
	Response.Write("</td></tr>")
	Response.Write("</table>")
	Response.Write DrawDialogBox("DIALOG_END","","")
	Response.Write("</form></center>")
	Call WriteFooter("")
End If
'end create new database form


'start design functions

If (Not ocdSQLTableEdits) Or ocdReadOnly Then
	Call WriteHeader("")
	Call WriteFooter(ocdBrandText & " is Read Only")
	Response.End()
ElseIf Not ocdSQLTableEdits Then
	Call WriteHeader("")
	Call WriteFooter("Access Denied")
	Response.End()
End If



If Request.Form("cancel") <> "" Then
	Select Case Request.QueryString("action")
		Case "dropproc" 
			Response.Redirect("Schema.asp?show=procs")
		Case Else
			Response.Redirect("Structure.asp?sqlfrom=" &  Request.QueryString("sqlfrom"))
	End Select
	Response.End()
End If

Set objSQL = New ocdManageSQLServer
objSQL.SQLConnect = ndnscSQLConnect
objSQL.SQLUser = ndnscSQLUser
objSQL.SQLPass = ndnscSQLPass
objSQL.SQLObject = Request.QueryString("sqlfrom")
Call objSQL.Open()
strSQLFrom = Request.QueryString("sqlfrom")
strAction = Request.QueryString("action")

Select Case UCase(strAction)
	'These require confirm
	Case "DELETETABLE","DELETERELATION", "DROPTABLE","DELETEFIELD","DELETEINDEX","DROPTRIG","DROPVIEW","DROPPROC"
		If Request.Form("confirm") = "" Then
			Call WriteHeader ("")
			Response.Write "<form action=" & Request.servervariables("SCRIPT_NAME") & "?" & Request.QueryString & "  method=post>"
			Response.Write("<center><table WIDTH=""60%"" class=""DialogBox""><tr><TH STYLE=""text-align:left;background-color:navy;color:white;"" align=left><DIV STYLE=""color:white;"">Confirm Action</DIV></TH><tr><td BGCOLOR=Silver valign=top>")
			Response.Write("<table><tr><td valign=top><img src=appWarning.gIf border=0 alt=""Warning""></td><td>&nbsp;</td><td valign=top>")
			Select Case UCase(strAction)
				Case "DELETETABLE"
					Response.Write("<p><span class=PageStatus>Are you sure you want to permanently remove all records from the table ")
					Response.Write(objSQL.SQLObject)
					Response.Write("? </span> <p>This action will Not affect the structure of the table but cannot be undone.")
				Case "DROPTABLE"
	
					Response.Write("<p><span class=PageStatus>Are you sure you want to permanently  drop the table ")
					Response.Write(Server.HTMLEncode(objSQL.SQLObject))
					Response.Write(" from this database?  </span><p>This action will remove the structure and all data records in the table and cannot be undone.<p>")
				Case "DELETEFIELD"
					Response.Write("<p><span class=PageStatus>Are you sure you want to permanently drop the field ")
					Response.Write(Server.HTMLEncode(Request.QueryString("fieldname")))
					Response.Write(" in table ")
					Response.Write(Server.HTMLEncode(objSQL.SQLObject))
					Response.Write(" ?</span><p>  This action will affect all records in the table and cannot be undone.")
				Case "DELETEINDEX"
					Response.Write("<p><span class=PageStatus>Are you sure you want to permanently drop the index ")
					Response.Write(Server.HTMLEncode(Request.QueryString("indexname")))
					Response.Write(" for table ")
					Response.Write(Server.HTMLEncode(objSQL.SQLObject))
					Response.Write(" in this database?  </span><p>This action cannot be undone.<p>")
				Case "DROPPROC"
					Response.Write("Are you sure you want to permanently remove the procedure ")
					Response.Write(Server.HTMLEncode(objSQL.SQLObject))
					Response.Write(" from this database?  </span><p>This action  cannot be undone.")

				Case "DROPTRIG"
					Response.Write("Are you sure you want to permanently remove the trigger ")
					Response.Write(Request.QueryString("trigname"))
					Response.Write(" on table ")
					Response.Write(Server.HTMLEncode(objSQL.SQLObject))
					Response.Write(" in this database?  </span><p>This action  cannot be undone.<p>")
				Case "DELETERELATION"
					Response.Write("<span class=PageStatus>Are you sure you want to permanently drop the relation ")
					Response.Write(Request("relationname"))
					Response.Write(" for table ")
					Response.Write(objSQL.SQLObject)
					Response.Write(" ? </span><p> This action cannot be undone.")
				Case "DROPVIEW"
					Response.Write("<span class=PageStatus>Are you sure you want to permanently drop the view ")
					Response.Write(Server.HTMLEncode(objSQL.SQLObject))
					Response.Write("  from this database?  </span><p>This action will remove the structure of the query and cannot be undone.  No records will be deleted.")
				Case Else
					Response.Write("NOT DEFINED")
			End Select 
			Response.Write("<p><input type=Submit name=Confirm  class=submit value=""OK"">&nbsp;<input type=Submit  class=submit  name=Cancel value=Cancel></td></tr></table></td></tr></table></center></form>")
			Call WriteFooter("")
			Response.End()
		Else
			Select Case Request.QueryString("action")
 '''''''''DELETE FIELD	
				Case "deletefield"
					 objSQL.DropField(Request.QueryString("fieldname"))
					If Err.Number = 0 Then
						Response.Redirect("Structure.asp?sqlFrom=" & objSQL.SQLObject)
					End If
				
'''''''''DELETE TABLE			
				Case "deletetable"
				
					objSQL.ADOConnection.execute "DELETE FROM " & objSQL.SQLObject
					If Err.Number = 0 Then
						Response.Redirect("Structure.asp?sqlFrom=" & objSQL.SQLObject)
					Else
						Call WriteHeader("")
						Call WriteFooter("")
					End If
'''''''''''DELETE RELATION
				Case "deleterelation"
					strSQL = "ALTER TABLE " & objSQL.SQLObject & " DROP CONSTRAINT " & ocdQUoteSuffix & Request("relationname") & ocdQUotePrefix & ""
					objSQL.ADOConnection.execute (strSQL)
					If Err.Number = 0 Then
						Response.Redirect("Structure.asp?sqlFrom=" & objSQL.SQLObject)
					Else
						Call WriteHeader("")
						Response.Write Server.HTMLEncode(strSQL)
						Call WriteFooter("")
					End If
				Case "deleteindex"
''''''''''DELETE INDEX
					objSQL.DropIndex(Request.QueryString("indexname"))
					If Err.Number = 0 Then
						Response.Redirect("Structure.asp?sqlfrom=" & objSQL.SQLObject)
					End If
''''''''''''DROP TABLE                  			
				Case "droptable"
					strSQL = "DROP TABLE " & objSQL.SQLObject
					objSQL.ADOConnection.execute strSQL
					If Err.Number = 0 Then
						Response.Redirect("Schema.asp?show=tables")
					Else
						Call WriteHeader("")
						Response.Write Server.HTMLEncode(strSQL)
						Call WriteFooter("")
					End If
''''''''''''''''DROP PROC
				Case "dropproc"
					strSQL = "DROP PROCEDURE " & objSQL.SQLObject
					objSQL.ADOConnection.execute strSQL
					If Err.Number = 0 Then
		
						Response.Redirect("Schema.asp?show=procs")
					Else
						Call WriteHeader("")
						Response.Write Server.HTMLEncode(strSQL)
						Call WriteFooter("")
					End If


'''''''''' DROP TRIGGER			
				Case "droptrig"
					strSQL = "DROP TRIGGER " & Request("trigname")
					objSQL.ADOConnection.execute (strSQL)
					If Err.Number = 0 Then
						Response.Redirect("Structure.asp?sqlFrom=" & objSQL.SQLObject)
Else
						Call WriteHeader("")
						Response.Write Server.HTMLEncode(strSQL)
						Call WriteFooter("")
					End If
''''''''''''''''DROP VIEW
				Case "dropview"
					 objSQL.ADOConnection.execute "DROP VIEW " & objSQL.SQLObject
					If Err.Number = 0 Then
						Response.Redirect("Schema.asp?show=views")
					End If
			End Select		
			If Err.Number <> 0 Then
				Call WriteHeader("")
				Call WriteFooter("")
				Response.End()
			End If
		End If
End Select





set connDesign = Server.CreateObject("ADODB.Connection")
connDesign.Open ndnscSQLConnect, ndnscSQLUser, ndnscSQLPass
connDesign.COmmandTimeout = ocddbtimeout

If Request("table")<>"" Then
	strAlterTableName = Request("table")
Else
	strAlterTableName = Request("altertablename")
End If
If ocdSQLTableEdits AND (connDesign.Provider = "SQLOLEDB.1") Then
Else
	'Response.end
End If

If Request("ocdAction") <> "" Then 
	Call WriteHeader("")
End If

Select Case UCase(Request("ocdAction"))
	CASE "ADDRELATION"

		Select Case UCase(Request("whatrelation"))
			Case "MANY-TO-ONE"
				strWhatIsSQLFrom = Request("SQLFrom")
				strWhatIsRelatedTo = Request("RelatedTo")
			Case Else
				strWhatIsRelatedTo  = Request("SQLFrom")
				strWhatIsSQLFrom = Request("RelatedTo")
		End Select
%>
<span class=INFORMATION>Add New Relation to Database</span>
<form action=<%=ocdPageName%>?sqlfrom=<%=server.urlencode(strWhatIsSQLFrom)%> method=post>
<input name="relatedto" type=Hidden value="<%=Server.HTMLEncode(strWhatIsRelatedTo)%>">
<strong>Name: </strong><input type=Text name=sqlrelationname value="" SIZE=40><br>
<table BORDER=1>
<tr><TH>Primary Key in <%=strWhatIsRelatedTo%></th>
<TH>Related Column in <%=strWhatIsSQLFrom%></th>
</tr>
<%
		set rsRelTbl = Server.CreateObject("ADODB.Recordset")
		set rsRelIDX = Server.CreateObject("ADODB.Recordset")
		set rsRelIDX = connDesign.openSchema(12) 'indexes
		rsRelFieldCount = 0
		rsRelTbl.open "SELECT * from " & strWhatIsSQLFrom & " WHERE 1=2", connDesign
		do while Not rsRelIDX.eof

			If (rsRelIDX("table_name")) = GetSQLIDFPart(strWhatIsRelatedTo,"SQLOBJECTNAME", ocdQuotePrefix,ocdQuoteSuffix) and (rsRelIDX("table_schema")) = GetSQLIDFPart(strWhatIsRelatedTo,"SQLOBJECTOWNER", ocdQuotePrefix,ocdQuoteSuffix) and  rsRelIDX("primary_key") = True Then
				rsRelFieldCount = rsRelFieldCount + 1
				Response.Write "<tr>"
				Response.Write "<td>"
				Response.Write rsRelIDX("COLUMN_NAME")
				Response.Write "<input type=Hidden name=""pkfld"
				Response.Write rsRelFieldCount
				Response.Write """ value="""
				Response.Write rsRelIDX("COLUMN_NAME")
				Response.Write """>"
				Response.Write "</td>"
				Response.Write "<td>"
				Response.Write "<select name=""fkfld"
				Response.Write rsRelFieldCount
				Response.Write """>"
				for each rsRelTblFld in rsRelTbl.fields
					Response.Write "<option value="""
					Response.Write Server.HTMLEncode(rsRelTblFld.Name)
					Response.Write """>"
					Response.Write Server.HTMLEncode(rsRelTblFld.Name)
					Response.Write "</option>"
				Next
				Response.Write "</select>"
				Response.Write "</td>"
				Response.Write "</tr>"
			End If
			rsRelIDX.movenext
		loop
%>
</table>
	<input name=relcheckdata type=checkbox> Check Existing Data
	<br>
<input name=relrepl type=checkbox> Enable constraint for replication
<br>

<%
%>

<input name=reldeleterule type=checkbox> Cascade Delete (SQL Server 2000+ required)
<br>
<input name=relupdaterule type=checkbox> Cascade Update (SQL Server 2000+ required)
<p>
<input type=Submit class=submit name=Action value="Create Relation">
<p>
</form>
<%
	Case "NEWINDEX"
%>

<span class=INFORMATION><A HREF=Schema.asp?show=tables>Tables</A> : <A HREF=Structure.asp?sqlfrom=<%=server.urlencode(Request.QueryString("sqlfrom"))%>><%=Request("sqlfrom")%></A> :  Add New Index</span>
<form action="<%=Request.servervariables("SCRIPT_NAME")%>" method=post>
<input type=HIDDEN name=sqlfrom value="<%=Server.HTMLEncode(Request("sqlfrom"))%>">
<strong>New Index Name:</strong> <input type=Text name=NewIndexName>
<select name=IndexUnique><option value=""></option>
<option value="Unique">Unique</option>
<option value="Primary Key">Primary Key</option>
</select>
<select name=IndexClustered>
<option value="NONCLUSTERED" selected>Non Clustered</option>
<option value="CLUSTERED">Clustered</option>
</select><br>
<input type=checkbox name=indexignoredup> Ignore Duplicate Key
<br>

<input type=checkbox name=indexdrop> Drop Existing Index
<br>
<input type=checkbox name=indexnostats> Don't Recompute Statistics
<br>
<input type=checkbox name=indexsorttempdb> Sort In TempDB
<br>
<input type=Checkbox name=INDEXPAD> Pad Index
<br>
Fill Factor <input name=IndexFill SIZE=3 MAXLENGTH=3 type=TEXT>% 
<p>
<input type=Submit class=submit name=Action value="Create Index">
<p><BLOCKQUOTE>
<%
		set rsTableNames = Server.CreateObject("ADODB.Recordset")
		rsTableNames.open "Select * from " & Request("sqlfrom") & " WHERE 1=2", connDesign ', 3 'adOpenStatic
		strSelectOptions = "<option value=""""></option>"
		for each fldF in rsTableNames.fields
			strSelectOptions = strSelectOptions & "<option value=""" & Server.HTMLEncode(fldF.Name) & """>" & Server.HTMLEncode(fldF.Name) & "</option>"
		next
%>
<table>
<% for rstxxxcn = 1 to 10 %>
<tr><td><strong>Field:</strong></td><td><select name=idxf<%=Trim(cstr(rstxxxcn))%>><%=strSelectOptions%></select></td><td><select name=idxf<%=Trim(cstr(rstxxxcn))%>sort><option value="ASC">Ascending</option><option value="DESC">Descending</option></select></td></tr>
<%next%>
</table>
</blockquote>
</form>

<%
	Case "NEWFIELD", "EDITFIELD"

	objSQL.open
	'Response.Write connDesign.connectionstring
	
	If UCase(Request("ocdAction")) = "EDITFIELD" Then
		tnname = Request("Table")
	Else
		tnName = Request("sqlfrom")
	End If
	
	set rsFSchema = Server.CreateObject("ADODB.Recordset")
	
	strUDTName = ""

	If UCase(Request("ocdAction")) = "EDITFIELD" Then
	objSQL.SQLObject = tnName
'	on error goto 0
	set rsFSchema = connDesign.OpenSchema(4, array(null,objSQL.getsqlobjectowner(),objSQL.getsqlobjectname(), Cstr(Request("fieldName"))))
If Not rsFSchema.eof Then
If Not isnull(rsFSchema("DOMAIN_NAME")) Then
strUDTName = rsFSchema("DOMAIN_NAME")
End If
End If	


		If strUDTName <> "" Then
		
		deftype = strUDTName

		Else
		Select Case CLng(Request("deftype"))
			case 0 'adEmpty
				deftype = "UNKNOWN"
			case 16 'adTinyInt 
				deftype = "TINYINT"
			case 2 'adSmallInt
				deftype = "SMALLINT"
			case 3 'adInteger 
				deftype = "INT"
			case 20 'adBigInt 
				deftype =  "BIGINT"
			case 17 'adUnsignedTinyInt 
				deftype =  "TINYINT"
			case 18 'adUnsignedSmallInt
				deftype =  "SMALLINT"
			case 19 'adUnsignedInt
				deftype =  "INT"
			case 21 'adUnsignedBigInt
				deftype = "BIGINT"
			case 4 'adSingle
				deftype = "REAL"
			case 5 'adDouble
				deftype = "FLOAT"
			case 6 'adCurrency
				deftype = "MONEY"
			case 14 'adDecimal
				deftype ="DECIMAL"
			case 131 'adNumeric
				deftype = "NUMERIC"
			case 11 'adBoolean
				deftype = "BIT"
			case 10 'adError
				deftype = "UNKNOWN"
			case 132 'adUserDefined
				deftype = "UNKNOWN"
			case 12 'adVariant
				deftype = "UNKNOWN"
			case 9 'adIDispatch
				deftype ="UNKNOWN"
			case 13 'adIUnknown
				deftype ="UNKNOWN"
			case 72 'adGUID
				deftype = "UNIQUEIDENTIFIER"
			case 7 'adDate
				deftype = "DATETIME"
			case 133 'adDBDate
				deftype = "DATETIME"
			case 134 'adDBTime
				deftype = "DATETIME"
			case 135 'adDBTimeStamp
				deftype = "DATETIME"
			case 8 'adBSTR
				deftype = "UNKNOWN"
			case 129 'adChar
				deftype = "CHAR"
			case 200 'adVarChar
				deftype = "VARCHAR"
			case 201 'adLongVarChar
				deftype = "TEXT"
			case 130 'adWChar
				deftype = "NCHAR"
			case 202 'adVarWChar
				deftype = "NVARCHAR"
			case 203 'adLongVarWChar
				deftype = "NTEXT"
			case 128 'adBinary 
				deftype = "BINARY"
			case 204 'adVarBinary
				deftype = "VARBINARY"
			case 205 'adLongVarBinary
				deftype = "IMAGE"
			case else
				deftype = "UNKNOWN"
'				deftype = CSTR(fldF.Type)
		End Select
End If
		%>
<span class=Information><A HREF=Schema.asp?show=tables>Tables</A> : 
<A HREF="Structure.asp?sqlfrom=<%=server.urlencode(Request.QueryString("table"))%>"><%=Server.HTMLEncode(Request.QueryString("table"))%></A> : <%=Server.HTMLEncode(Request("fieldname"))%> 
</span>
		
		<%
		Else
%>
<span class=Information><A HREF=Schema.asp?show=tables>Tables</A> : 
<A HREF="Structure.asp?sqlfrom=<%=server.urlencode(Request.QueryString("sqlfrom"))%>"><%=Server.HTMLEncode(Request.QueryString("sqlfrom"))%></A> : Add New Field
</span>
<%End If%>
<script type="text/javascript" language="JavaScript">
  <!--
  function CheckIDs(){
  mydoc =document.forms[0];
	  myFieldType = mydoc.NewFieldType.options[document.forms[0].NewFieldType.selectedIndex].value;
if (myFieldType == "uniqueidentifier") {
document.forms[0].allownulls.selectedIndex = 1;
document.forms[0].is_rowguid.disabled = false;
if ( document.forms[0].identity.checked == true )
				  {
				   document.forms[0].identity.checked == false;
					document.forms[0].identity_seed.value = "";
										document.forms[0].identity_increment.value = "";
									document.forms[0].identity_seed.disabled = true;
					document.forms[0].identity_increment.disabled = true;					
document.forms[0].identity.disabled = true;					
				  }
}else{
document.forms[0].is_rowguid.disabled = true;

}
  
if  ( (myFieldType == "decimal") || (myFieldType == "int") || (myFieldType == "numeric") || (myFieldType == "tinyint") || (myFieldType == "smallint"))
		{
		if ( document.forms[0].identity.checked == true )
				  {
document.forms[0].NewFieldScale.value = "0";
					document.forms[0].identity_seed.value = "1";					document.forms[0].identity_increment.value = "1";
					document.forms[0].identity_seed.value = "1";
					document.forms[0].identity_seed.disabled = false;
					document.forms[0].identity_increment.disabled = false;					
				  }
				  else
				  {
				document.forms[0].identity.disabled = false;
	document.forms[0].identity_increment.value = "";
					document.forms[0].identity_seed.value = "";
					document.forms[0].identity_seed.disabled = true;
					document.forms[0].identity_increment.disabled = true;					
				  }

		
		document.forms[0].NewFieldScale.disabled = false;	
				document.forms[0].NewFieldPrecision.disabled = false;	
		}
		else
		{
		document.forms[0].NewFieldScale.disabled = true;	
				document.forms[0].NewFieldPrecision.disabled = true;	
document.forms[0].identity_increment.value = "";
					document.forms[0].identity_seed.value = "";
		document.forms[0].identity.checked = false;	document.forms[0].identity_increment.value = "";
					document.forms[0].identity_seed.disabled = true;
				document.forms[0].identity.disabled = true;
	document.forms[0].identity_increment.disabled = true;					


		}
  if ((myFieldType == "varchar") || (myFieldType == "char") || (myFieldType == "nvarchar") || (myFieldType == "nchar")) {
		
			document.forms[0].NewFieldSize.disabled = false;
		} else{
				document.forms[0].NewFieldSize.disabled = true;
		}
				
  
  }
function UpdateControls()
	{
		mydoc =document.forms[0];
	  myFieldType = mydoc.NewFieldType.options[document.forms[0].NewFieldType.selectedIndex].value;
CheckIDs();

if (myFieldType == "bit") {
document.forms[0].allownulls.selectedIndex = 1;
}
		if (myFieldType == "binary"){
		
		document.forms[0].NewFieldSize.value = "10";
		document.forms[0].NewFieldPrecision.value = "0";
		document.forms[0].NewFieldScale.value = "0";
		}
		
		if (myFieldType == "bit"){
		
		document.forms[0].NewFieldSize.value = "1";
		document.forms[0].NewFieldPrecision.value = "0";
		document.forms[0].NewFieldScale.value = "0";
		document.forms[0].allownulls.selectedindex = 1;
		}
			
			if (myFieldType == "char"){
		
		document.forms[0].NewFieldSize.value = "10";
		document.forms[0].NewFieldPrecision.value = "0";
		document.forms[0].NewFieldScale.value = "0";
		}	
				if (myFieldType == "datetime"){
		
		document.forms[0].NewFieldSize.value = "8";
		document.forms[0].NewFieldPrecision.value = "0";
		document.forms[0].NewFieldScale.value = "0";
		}			
				
						if (myFieldType == "decimal"){
		
		document.forms[0].NewFieldSize.value = "9";
		document.forms[0].NewFieldPrecision.value = "18";
		document.forms[0].NewFieldScale.value = "0";
		}			
				
								if (myFieldType == "float"){
		
		document.forms[0].NewFieldSize.value = "8";
		document.forms[0].NewFieldPrecision.value = "53";
		document.forms[0].NewFieldScale.value = "0";
		}	
				
						if (myFieldType == "image"){
		
		document.forms[0].NewFieldSize.value = "16";
		document.forms[0].NewFieldPrecision.value = "0";
		document.forms[0].NewFieldScale.value = "0";
		}			
				
				if (myFieldType == "int"){
		
		document.forms[0].NewFieldSize.value = "4";
		document.forms[0].NewFieldPrecision.value = "10";
		document.forms[0].NewFieldScale.value = "0";
		}		
		if (myFieldType == "money"){
		
		document.forms[0].NewFieldSize.value = "8";
		document.forms[0].NewFieldPrecision.value = "19";
		document.forms[0].NewFieldScale.value = "4";
		}		
		
		
		if (myFieldType == "nchar"){
		
		document.forms[0].NewFieldSize.value = "10";
		document.forms[0].NewFieldPrecision.value = "0";
		document.forms[0].NewFieldScale.value = "0";
		}		
		
		if (myFieldType == "ntext"){
		
		document.forms[0].NewFieldSize.value = "16";
		document.forms[0].NewFieldPrecision.value = "0";
		document.forms[0].NewFieldScale.value = "0";
		}				
		
		if (myFieldType == "numeric"){
		
		document.forms[0].NewFieldSize.value = "9";
		document.forms[0].NewFieldPrecision.value = "18";
		document.forms[0].NewFieldScale.value = "0";
		}	
		
		if (myFieldType == "nvarchar"){
		
		document.forms[0].NewFieldSize.value = "50";
		document.forms[0].NewFieldPrecision.value = "0";
		document.forms[0].NewFieldScale.value = "0";
		}	
		if (myFieldType == "real"){
		
		document.forms[0].NewFieldSize.value = "4";
		document.forms[0].NewFieldPrecision.value = "24";
		document.forms[0].NewFieldScale.value = "0";
		}	
		
		if (myFieldType == "smalldatetime"){
		
		document.forms[0].NewFieldSize.value = "4";
		document.forms[0].NewFieldPrecision.value = "0";
		document.forms[0].NewFieldScale.value = "0";
		}	
		
		if (myFieldType == "smallint"){
		
		document.forms[0].NewFieldSize.value = "2";
		document.forms[0].NewFieldPrecision.value = "5";
		document.forms[0].NewFieldScale.value = "0";
		}	
		
		if (myFieldType == "smallmoney"){
		
		document.forms[0].NewFieldSize.value = "4";
		document.forms[0].NewFieldPrecision.value = "10";
		document.forms[0].NewFieldScale.value = "4";
		}	
		
				if (myFieldType == "text"){
		
		document.forms[0].NewFieldSize.value = "16";
		document.forms[0].NewFieldPrecision.value = "0";
		document.forms[0].NewFieldScale.value = "0";
		}	
		
				if (myFieldType == "timestamp"){
		
		document.forms[0].NewFieldSize.value = "8";
		document.forms[0].NewFieldPrecision.value = "0";
		document.forms[0].NewFieldScale.value = "0";
		}	
		
				if (myFieldType == "tinyint"){
		
		document.forms[0].NewFieldSize.value = "1";
		document.forms[0].NewFieldPrecision.value = "3";
		document.forms[0].NewFieldScale.value = "0";
		}	
				if (myFieldType == "uniqueidentifier"){
		
		document.forms[0].NewFieldSize.value = "16";
		document.forms[0].NewFieldPrecision.value = "0";
		document.forms[0].NewFieldScale.value = "0";
		document.forms[0].fdefault.value = "NEWID()";
		}	
				if (myFieldType == "varbinary"){
		
		document.forms[0].NewFieldSize.value = "50";
		document.forms[0].NewFieldPrecision.value = "0";
		document.forms[0].NewFieldScale.value = "0";
		}	
				if (myFieldType == "varchar"){
		
		document.forms[0].NewFieldSize.value = "50";
		document.forms[0].NewFieldPrecision.value = "0";
		document.forms[0].NewFieldScale.value = "0";
		}	
		

		  
				  

		
		}
  // -->

</script>
<form action="<%=Request.Servervariables("SCRIPT_NAME")%>" method=post>
<input type=Hidden name=altertablename value="<%=Server.HTMLEncode(tnname)%>">
<input type=Hidden name=fieldname value="<%=Server.HTMLEncode(Request("fieldname"))%>">
            <table border="1" class=Grid>
              <tr class=gridheader>
<TH>Field Name</TH>
<TH>Type</TH>
<TH>Length</TH>
<TH>Precision</TH>
<TH>Scale</TH>
              </tr>
			              
              <tr class=gridrowodd>

                <td><input type="text" name="NewFieldName" size="20" maxlength="64" value="<%=Server.HTMLEncode(Request("fieldname"))%>"></td>
                <td><select onchange="UpdateControls();" size="1" name="NewFieldType">
                    <option value="binary" <%If UCase(deftype) = "BINARY" Then Response.Write " selected "%>>binary</option>
                    <option value="bit"<%If UCase(deftype) = "BIT" Then Response.Write " selected "%>>bit</option>
                    <option value="char"<%If UCase(deftype) = "CHAR" Or deftype="" Then Response.Write " selected "%>>char</option>
                    <option value="datetime"<%If UCase(deftype) = "DATETIME" Then Response.Write " selected "%>>datetime</option>
                    <option value="decimal"<%If UCase(deftype) = "DECIMAL" Then Response.Write " selected "%>>decimal</option>
                    <option value="float"<%If UCase(deftype) = "FLOAT" Then Response.Write " selected "%>>float</option>
                    <option value="image"<%If UCase(deftype) = "IMAGE" Then Response.Write " selected "%>>image</option>
                    <option value="int"<%If UCase(deftype) = "INT" Then Response.Write " selected "%>>int</option>
                    <option value="money"<%If UCase(deftype) = "MONEY" Then Response.Write " selected "%>>money</option>
                    <option value="nchar"<%If UCase(deftype) = "NCHAR" Then Response.Write " selected "%>>nchar</option>
                    <option value="ntext"<%If UCase(deftype) = "NTEXT" Then Response.Write " selected "%>>ntext</option>
                    <option value="numeric" <%If UCase(deftype) = "NUMERIC" Then Response.Write " selected "%>>numeric</option>
                    <option value="nvarchar" <%If UCase(deftype) = "NVARCHAR" Then Response.Write " selected "%>>nvarchar</option>
                    <option value="real"<%If UCase(deftype) = "REAL" Then Response.Write " selected "%>>real</option>
                    <option value="smalldatetime" <%If UCase(deftype) = "SMALLDATETIME" Then Response.Write " selected "%>>smalldatetime</option>
                    <option value="smallint" <%If UCase(deftype) = "SMALLINT" Then Response.Write " selected "%>>smallint</option>
                    <option value="smallmoney" <%If UCase(deftype) = "SMALLMONEY" Then Response.Write " selected "%>>smallmoney</option>
                    <option value="text"<%If UCase(deftype) = "TEXT" Then Response.Write " selected "%>>text</option>
                    <option value="timestamp" <%If UCase(deftype) = "TIMESTAMP" Then Response.Write " selected "%>>timestamp</option>
                    <option value="tinyint" <%If UCase(deftype) = "TINYINT" Then Response.Write " selected "%>>tinyint</option>
                    <option value="uniqueidentifier"<%If UCase(deftype) = "UNIQUEIDENTIFIER" Then Response.Write " selected "%>>uniqueidentifier</option>
                    <option value="varbinary"<%If UCase(deftype) = "VARBINARY" Then Response.Write " selected "%>>varbinary</option>
                    <option value="varchar" <%If UCase(deftype) = "VARCHAR" Then Response.Write " selected "%>>varchar</option>
<% If strUDTName <> "" Then%>
                    <option value="<%=Server.HTMLEncode(strUDTName)%>" selected><%=Server.HTMLEncode(strUDTName)%></option>
<%End If%>
                  </select></td>
                <td><input type="text" name="NewFieldSize" size="3" maxlength="4"  value="<%=Server.HTMLEncode(Request("defsize"))%>"></td>
                <td><input type="text" name="NewFieldPrecision" size="3" maxlength="3" value="<%=Server.HTMLEncode(Request("defprecision"))%>"></td>
                <td><input type="text" name="NewFieldScale" size="3" maxlength="3" value="<%=Server.HTMLEncode(Request("defscale"))%>" ></td></tr></table>
<%If 1=1 Then 'UCase(Request("strAction")) = "NEWFIELD" Or (UCase(Request("strAction")) = "EDITFIELD"  and Request("defidentity") = "") Then %>
  <table border="1" class=Grid>
              <tr class=gridheader>
<TH>ID</TH>
<TH>ID Seed</TH>
<TH>ID Incr</TH>
<TH>GUID</TH>
<TH>Required</TH>
<TH>Default</TH>
              </tr>
			              
              <tr>
                <td> <input type="checkbox" name="identity" <%
If Request("defidentity") <> "" Then

	Response.Write " checked "

End If


%> onclick="javascript:UpdateControls()"></td>
                <td><input type="text" name="identity_seed" size="3" maxlength="50" value=1></td>
                <td><input type="text" name="identity_increment" size="3" maxlength="50" value="1"></td>
                <td><input type="checkbox" name="is_rowguid"></td>
<%End If%>
			
                <td><select name=allownulls><option value="NULL" <%
If UCase(Request("nulldef")) = "NULL" Or Request("nulldef") = "" Then
	Response.Write " selected "
End If
%>>False</option><option value="NOTNULL" <%
If UCase(Request("nulldef")) = "NOTNULL" Then
	Response.Write " selected "
End If
%>>True</option></select></td>
                <td><input type="text" name="fdefault" size="24" maxlength="255" value="<%

strcrit = ",," & Request.QueryString("table") & "," & Request("fieldname")
arrcrit = split(strcrit,",")

set rsdef = Server.CreateObject("ADODB.Recordset")
 set rsdef = connDesign.OpenSchema(4) ',array(Empty,Empty,(Request.QueryString("table")),(Request("fieldname"))))
If Err.Number <> 0 Then
	Call WriteFooter("")
End If

do while Not rsdef.eof 
 	If GetSQLIDFPart(Request.QueryString("Table"),"SQLOBJECTNAME", ocdQuotePrefix,ocdQuoteSuffix) = rsdef("TABLE_NAME") Then
	If Request("fieldname") = rsdef("COLUMN_NAME") Then
	If Not isnull(rsdef("COLUMN_DEFAULT")) Then
	Response.Write Server.HTMLEncode(rsdef("COLUMN_DEFAULT"))
	End If
	End If
	End If
	rsdef.movenext
loop
%>"></td>

              </tr>
              
            </table><br>
<%
		Response.Write "Use 'single quotes' syntax for literal text values in defaults"
%>
<p>
<%	

If strUDTName <> "" Then%>
<span class=warning><img src=appwarningsmall.gif> Cannot Edit User Defined Type (UDT) Field </span>
<%else
If UCase(Request("ocdAction")) = "EDITFIELD" Then %>
<input type=SUBMIT  class=submit name='Action' value='Alter Field'>
<script type="text/javascript" language="JavaScript">
 
 CheckIDs();
  
  // -->
 </script>
<p>
</P>
<%else%>
<input type=SUBMIT  class=submit  name='Action' value='Create New Field'>

<script type="text/javascript" language="JavaScript">
  <!--
  UpdateControls();
  -->
 </script>
 <p>
Use the <A HREF=Command.asp class=menu>Command</A> screen to create a new field based on a User Defined Type (UDT).</P>
 <%End If
 End If%>
 </form><p>
<%
	
	Call WriteFooter("")
	Response.end
	
Case ""

	
		OpenIdentifier = """"
		CloseIdentifier = """"
	


If instr(1,(ndnscSQLConnect),("Microsoft.Jet.OLEDB")) > 0 Then
	ndTargetConnType = "JET.OLEDB"
End If
	

	
strNDURLConnect = "?connect=" & Server.URLEncode(Request.QueryString("connect")) & "&user=" & Server.URLEncode(Request.QueryString("user")) & "&pass=" & Server.URLEncode(Request.QueryString("pass") )

strNDURLSQL = "&sqlselect=" & Server.URLEncode(Request.QueryString("sqlselect")) & "&sqlfrom=" & Server.URLEncode(Request.QueryString("sqlfrom")) & "&sqlwhere=" & Server.URLEncode(Request.QueryString("sqlwhere"))& "&sqlorderby=" & Server.URLEncode(Request.QueryString("sqlorderby"))

fmNewTableKeyName = Request.Form("addnewtablekeyname")
fmNewTableKeyType = Request.Form("addnewtablekeytype")
fmNewTableKeySize = Request.Form("addnewtablekeysize")
qsFieldName = Request.QueryString("fieldname")
fmNewFieldName = Request.Form("newfieldname")
fmNewFieldType = Request.Form("newfieldtype")
fmNewFieldSize = Request.Form("newfieldsize")

ndAction = Request.Form("action")
If Request.QueryString("action") <> "" Then
	ndAction = Request.QueryString("action")
End If

set rsXXTemp=Server.CreateObject("ADODB.Recordset")
Select Case ndAction
'''''''''''''''''''CREATE INDEX
	CASE "Create Index"

If Request("IndexUnique") = "Primary Key" Then
strIndSQL = "ALTER TABLE "
		strIndSQL = strIndSQL & Request("sqlfrom")
		OBJsql.sqlObject = Request("SQLfROM")
		strIndSQL = strIndSQL & " ADD CONSTRAINT " & Request("NewIndexName") & " PRIMARY KEY "
strIndSQL = strIndSQL & Request("indexclustered")
strIndSQL = strIndSQL & " ("
If Request("idxf1") = "" Then

Else
		for countitxx = 1 to 10
		If Request("idxf" & Trim(cstr(countitxx))) <> "" Then
			strIndSQL = strINDSQL & "[" & Request("idxf" & Trim(cstr(countitxx))) & "],"
	
		End If
	
		Next
		strIndSQL = left(strIndSQL, len(strIndSQL)-1)
End If		
strIndSQL = strIndSQL & ")"
		
Else 
strIndSQL = "CREATE "
If Request("IndexUnique") <> "" Then
	strIndSQL = strIndSQL & " UNIQUE "
End If
strIndSQL = strIndSQL & Request("indexclustered")
	strIndSQL = strIndSQL & " INDEX "
		strIndSQL = strIndSQL & Request("NewIndexName")

		strIndSQL = strIndSQL & " "
			strIndSQL = strIndSQL & " ON "
		strIndSQL = strIndSQL & Request("sqlfrom")

		strIndSQL = strIndSQL & " ("
		
			for countitxx = 1 to 10
		If Request("idxf" & Trim(cstr(countitxx))) <> "" Then
			strIndSQL = strINDSQL & "[" & Request("idxf" & Trim(cstr(countitxx))) & "]"
			strIndSQL = strINDSQL & " " & Request("idxf" & Trim(cstr(countitxx)) & "sort") & ","
		End If
next
		strIndSQL = left(strIndSQL, len(strIndSQL)-1)
		strIndSQL = strIndSQL & ")"
		If Request("IndexOptions") <> "" Or Request("indexpad") <> "" Or Request("indexignoredup") <> "" Or Request("indexdrop") <> "" Or Request("indexnostat") <> "" Or Request("indexsortempdb") Or (Request("indexfill") <> "" and Request("indexfill") <> "0") Then
				strIndSQL = strIndSQL & " WITH "
				If Request("indexpad") <> "" Then
					strIndSQL = strIndSQL & " PAD_INDEX,"
				End If
				If Request("indexfill") <> "" and Request("indexfill") <> "0" Then
				strIndSQL = strIndSQL & " FILLFACTOR=" & Request("indexfill") & ","
				End If
				If Request("indexignoredup") <> "" Then
				strIndSQL = strIndSQL & " IGNORE_DUP_KEY,"
				End If
								If Request("indexdrop") <> "" Then
				strIndSQL = strIndSQL & " DROP_EXISTING,"
				End If
				If Request("indexnostats") <> "" Then
				strIndSQL = strIndSQL & " STATISTICS_NORECOMPUTE,"
				End If
								If Request("indexsorttempdb") <> "" Then
				strIndSQL = strIndSQL & " SORT_IN_TEMPDB,"
				End If
				strINDSQL = left(strIndSQL,len(strIndSQL)-1)
		End If
End If
		connDesign.Execute (strIndSQL)
		If Err.Number <> 0 Then
			Call WriteHeader("")
			Response.Write strIndSQL

			Call WriteFooter("")
		End If

	Response.Redirect("Structure.asp?sqlfrom=" & Request("sqlfrom"))
		Case "Create Relation"

		strFKSQL = "ALTER TABLE "
		strFKSQL = strFKSQL & Request("sqlfrom")

		strFKSQL = strFKSQL & " "
		If Request("relcheckdata") <> "" Then
		strFKSQL = strFKSQL & " WITH CHECK "
			
		Else
		strFKSQL = strFKSQL & " WITH NOCHECK "
		
		End If		
		strFKSQL = strFKSQL & " ADD CONSTRAINT [" & Request("sqlrelationname") & "] FOREIGN KEY "
		strFKSQL = strFKSQL & " ("
		for irrra = 1 to 10
		If Request("fkfld" & Trim(cstr(irrra))) <> "" Then
			strFKSQL = strFKSQL & "[" & Request("fkfld" & Trim(cstr(irrra))) & "],"
	
		End If
	
		next
		strFKSQL = left(strFKSQL, len(strFKSQL)-1)
		strFKSQL = strFKSQL & ")"

		
		strFKSQL = strFKSQL & " REFERENCES "
		strFKSQL = strFKSQL & Request("relatedto")

		strFKSQL = strFKSQL & "  "
		strFKSQL = strFKSQL & " ("
		for irrra = 1 to 10
		If Request("pkfld" & Trim(cstr(irrra))) <> "" Then
			strFKSQL = strFKSQL & "[" & Request("pkfld" & Trim(cstr(irrra))) & "],"
	
		End If
					Next
		strFKSQL = left(strFKSQL, len(strFKSQL)-1)
		strFKSQL = strFKSQL & ")"
		
			If Request("relupdaterule") <> "" Then
    strFKSQL = strFKSQL & " ON UPDATE CASCADE"
End If
If Request("reldeleterule") <> "" Then
     strFKSQL = strFKSQL & " ON DELETE CASCADE"
End If
If Request("relrepl") <> "" Then
	
Else
     strFKSQL = strFKSQL & " Not FOR REPLICATION"
End If
		connDesign.Execute (strFKSQL)
		If Err.Number <> 0 Then
			Call WriteHeader("")
			Response.Write strFKSQL
			Call WriteFooter("")
		End If
	Response.Redirect("Structure.asp?sqlfrom=" & (Request("sqlfrom")))
	
		Response.End()
	Case "Create New Field"	, "Alter Field"
			set rsCreate = Server.CreateObject("ADODB.Recordset")

		If ndAction = "Alter Field" Then
			If Request.Form("Fieldname") <> fmNewFieldName Then
				strSQL = "exec sp_rename '[" &  GetSQLIDFPart(strAlterTableName, "SQLOBJECTOwner",ocdQuoteSuffix,ocdQUotePrefix) & "].[" &   GetSQLIDFPart(strAlterTableName, "SQLOBJECTName",ocdQuoteSuffix,ocdQUotePrefix) & "].[" & Request.Form("FieldName") & "]','" & fmNewFieldName & "','COLUMN'"
				connDesign.execute strSQL
				If Err.Number <> 0 Then
					Call WriteHeader("")
					Call WriteFooter("")
				End If
				strSQL = ""
			End If
			
			If (fmNewFieldType <> "" and ndAction = "Alter Field") Or ndAction <> "Alter Field"  Then
				set rsCreate = connDesign.execute ("Select ""ID"" from sysobjects where ""name""='" & replace(GetSQLIDFPart(strAlterTableName,"SQLOBJECTNAME", ocdQuotePrefix,ocdQuoteSuffix),"'","''") & "'")
				If Err.Number <> 0 Then
					Call WriteHeader("")
					Call WriteFooter("")
				End If
				If Not rsCreate.eof Then
					intTableID = rsCreate("ID")
					rsCreate.close
					set rsCreate = connDesign.Execute ("SELECT ""CDEFAULT"" from syscolumns where ""name""='" & replace(fmNewFieldName,"'","''") & "'" & " and ""id"" = " & intTableID )
				If Err.Number <> 0 Then
					Call WriteHeader("")
					Call WriteFooter("")
				End If
					If Not rsCreate.eof Then
						intColID = rsCreate("CDEFAULT")
						rsCreate.close
						set rsCreate = connDesign.Execute ("SELECT ""NAME"" from sysobjects where  ""id"" = " & intColID )
						If Not rsCreate.eof Then
							strDCName = rsCreate("NAME")
							rsCreate.close
							connDesign.Execute "ALTER TABLE " & strAlterTableName & " DROP Constraint " & OpenIdentifier & strDCName & CloseIdentifier & ""
				If Err.Number <> 0 Then
					Call WriteHeader("")
					Call WriteFooter("")
				End If
						End If
					End If
				End If		
			End If
		End If
		If (fmNewFieldType <> "" and ndAction = "Alter Field") Or ndAction <> "Alter Field" Then
			If ndAction = "Alter Field" Then
				strSQL = "ALTER TABLE " & strAlterTableName &  " ALTER COLUMN "
			Else
				strSQL = "ALTER TABLE " & strAlterTableName & " ADD  "
			End If
			strSQL=  strSQL & ocdQuotePrefix & fmNewFieldName & ocdQUoteSuffix 
			Select Case UCase(fmNewFieldType)
				Case "VARCHAR","CHAR","NVARCHAR","NCHAR"
					strSQL=  strSQL & " " & UCase(fmNewFieldType) & "(" & Request("newfieldsize") & ")"
		
				Case  "DECIMAL","NUMERIC"
				
					If Request("newfieldscale") <> "" Or Request("newfieldprecision") <> "" Then
					If Request("newfieldscale") <> "" and Request("newfieldprecision") <> "" Then
					strSQL=  strSQL & " " & UCase(fmNewFieldType) & "(" & Request("newfieldprecision") & "," & Request("newfieldscale") & ")"
					
					Else
					strSQL=  strSQL & " " & UCase(fmNewFieldType) & "(" & Request("newfieldprecision") & ")"
					End If
					Else
					strSQL=  strSQL & " " & UCase(fmNewFieldType)
					End If
				Case Else
					strSQL=  strSQL & " " & UCase(fmNewFieldType)
				
		 End Select
		 If Request("identity") <> "" and ndAction <> "Alter Field" Then
		 	strSQL = strSQL & " IDENTITY (" & Request("identity_seed") & "," & Request("identity_increment") & ") "

		 	
		 End If
			If UCase(Request("allownulls")) = "NULL" Then
				strSQL = strSQL & " "
			Else
				strSQL = strSQL & " Not NULL "
				If ndAction = "Alter Field" Then
				Else
				End If
			End If
			If ndAction = "Alter Field" Then
				connDesign.execute strSQL
								If Err.Number <> 0 Then
					Call WriteHeader("")
					Response.Write strSQL

					Call WriteFooter("")
				End If
				'now make default
				If Request("fdefault") <> "" Then
					strSQL = "ALTER TABLE " & strAlterTableName & "  ADD CONSTRAINT ""DF"
strSQL = strSQL & replace(left(GetSQLIDFPart(strAlterTableName,"SQLOBJECTNAME", ocdQuotePrefix,ocdQuoteSuffix),25)," ","_") & "_" & replace(left(fmNewFieldName,25)," ","_") & "_" & Replace(Replace(right(cstr(now),9)," ","_"),":","-")
strSQL= strSQL & """ DEFAULT " & replace(Request("fdefault"),"'","'") & " FOR " & OpenIdentifier & fmNewFieldName & CloseIdentifier & ""

					connDesign.execute strSQL
									If Err.Number <> 0 Then
					Call WriteHeader("")
					Response.Write strSQL
					Call WriteFooter("")
				End If
				End If
			Else
				If Request("identity") = ""  Then
				strSQL = strSQL & " DEFAULT " 
				If Request("fdefault") <> "" Then
					strSQL = strSQL & Replace(Request("fdefault"),"'","'")
				Else
					strSQL = strSQL & " NULL "
				End If
				End If
				connDesign.execute strSQL
								If Err.Number <> 0 Then
					Call WriteHeader("")
				Response.Write Server.HTMLEncode(strSQL)

					Call WriteFooter("")
				End If
			End If
			strSQL = ""
		End If
	'End If
		Response.Redirect "Structure.asp?sqlfrom=" & strAlterTableName
		Response.end
'''''''''RENAME TABLE
		Case "renametable","renameview"
			If Request.Form("DoAction") = "" Then
				Call WriteHeader ("")
				%>
				<span class=Information><%
				If Request.QueryString("action") = "renametable" Then
				%><A HREF=Schema.asp?Show=tables>Tables</A><%
				else
				%><A HREF=Schema.asp?Show=views>Views</A><%
				End If%> : <A HREF=Structure.asp?sqlfrom=<%=Server.URLEncode(objSQL.SQLObject)%>> <%=objSQL.SQLObject%></A> : Rename </span>
				<form action=<%=Request.servervariables("SCRIPT_NAME") & "?action=" & Request.QueryString("action") & "&sqlfrom=" & Server.URLEncode(objSQL.SQLObject)%> method=post>
<strong>New Object Name </strong> <input SIZE=40 type=Table name=rnname value="<%=Server.HTMLEncode(objSQL.GetSQLObjectName)%>"> <p>
<input type=Submit name=DoAction  class=submit value="OK">&nbsp;<input type=Submit  class=submit name=Cancel value=Cancel></form>
				<%
				Call WriteFooter("")
			Else
'				on error resume next
				objSQL.ADOConnection.Execute ("sp_rename '" & replace(Request.QueryString("sqlfrom"),"'","''") & "' , '" & Replace(Request.Form("rnName"),"'","''") & "', 'OBJECT'")
				If Err.Number = 0 Then
					If Request.QueryString("action") = "renametable" Then
					Response.Redirect("Schema.asp?show=tables")
					Else
										Response.Redirect("Schema.asp?show=views")
					End If
				Else
					Call WriteHeader("")
					Call WriteFooter("")
				End If
			End If
			
			
			
'''''''''COPY TABLE
		Case "copytable","copyview"
			If Request.Form("DoAction") = "" Then
				Call WriteHeader ("")
				%>
				<span class=Information><A HREF=Schema.asp?Show=tables>Tables</A> : <A HREF=Structure.asp?sqlfrom=<%=Server.URLEncode(objSQL.SQLObject)%>> <%=objSQL.SQLObject%></A> : Copy </span>
				<form action=<%=Request.servervariables("SCRIPT_NAME") & "?action=" & Request.QueryString("action") & "&sqlfrom=" & Server.URLEncode(objSQL.SQLObject)%> method=post>
<strong>Copy </strong><select name=cptype><option value="Stucture">Structure Only</option><option value="Data" selected>Structure and Data</option></select><strong> to new table </strong><input type=Table name=cpname> <p>
<input type=Submit name=DoAction  class=submit value="OK">&nbsp;<input type=Submit  class=submit name=Cancel value=Cancel></form>
				<%
				Call WriteFooter("")
			Else
				on error resume next
				Call objSQL.CopyObject (Request("cptype"), Request("cpname"))

				If Err.Number = 0 Then
					Response.Redirect("Schema.asp?show=tables")
				Else
					Call WriteHeader("")
					Call WriteFooter("")
				End If
			End If
		CASE "createview"

			If Request.Form("doaction") = "" Then
				Call WriteHeader("")
%>
<span class=INFORMATION><A HREF=Schema.asp?show=views>Views</A> : Add New</span>

<form action="<%=Request.Servervariables("SCRIPT_NAME")%>?<%=Request.QueryString%>" method=post>
<strong>Name : </strong><input type=Text name=newviewname value="User_View1" SIZE=40><br>
<strong>SQL Text :</strong> <br>
<textarea name=SQLCommandText COLS=50 ROWS=7><%
If Request("SQLCommandText") <> "" Then
	Response.Write(Server.HTMLEncode(Request("SQLCommandText")))
Else
	Response.Write(Server.HTMLEncode(Request.QueryString("proposedsqlviewtext")))
End If
%></textarea>
<br>
<input type=Submit class=submit name=DoAction value="Create Query">
</form>
<%

				Call WriteFooter("")
	Response.end	
			Else
				objSQL.ADOConnection.Execute "CREATE VIEW " & Request.Form("newviewname") & " AS " &  Request.Form("SQLCommandText")
				If Err.Number <> 0 Then
					Call WriteHeader("")
					Call WriteFooter("")
				Else	
					Response.Redirect "Schema.asp?show=views"
					Response.end
				End If
			End If
''''''''''NEW TABLE
		Case "newtable"
			If Request.Form("doaction") = "" Then
				Call WriteHeader("")
%>
<span class=Information><A HREF=Schema.asp?show=tables>Tables</A> : Add New</span><form action="<%=Request.Servervariables("SCRIPT_NAME")%>?<%=Request.QueryString%>" method=post><strong>Name:</strong> <input type=TEXT name=AlterTableName SIZE=34 maxlength=64 value="NewTable">
<br><input  class=submit  type=SUBMIT name='DoAction' value='Create New Table'></form><p>			
			<%
				Call WriteFooter("")
			
			Else

			strNTName = Request.Form("AlterTableName")
			If instr(strNTName,".") = 0 Then
				strNTName = ocdQuotePrefix & FormatForSQL(strNTName, "SQLServer", "RemoveSQLIdentifier")  & ocdQuoteSuffix
			End If
			strSQL = "CREATE TABLE " & strNTName & " (" & ocdQuotePrefix & "RecordID" & ocdQuoteSuffix & " INTEGER IDENTITY Not NULL CONSTRAINT " & ocdQuotePrefix & "pk" & FormatForSQL(strNTName, "SQLServer", "RemoveSQLIdentifier") & "RecordID" & ocdQuoteSuffix & " PRIMARY KEY)"
				
				objSQL.ADOConnection.execute strSQL
				If Err.Number <> 0 Then
				
					Call WriteHeader("")
					Response.Write(strSQL)
					Call WriteFooter("")
				Else
				'If Instr(strNTName,".") = 0 Then
set rsuser = objSQL.ADOConnection.execute("SELECT USER_NAME()")
				
'				Response.Write err.description
				strUN = rsUser(0)
				rsUser.Close
				Set rsUser = Nothing
				Err.Clear

				strOName = GetSQLIDFPart(Request.Form("AlterTableName"), "SQLOBJECTOWNER", ocdquoteprefix, ocdquotesuffix)

				If strOName = "" Then
				set rsuser2 = objSQL.ADOConnection.execute("SELECT USER_NAME()")
				If Err.Number <> 0 Then
				Response.Write err.description
				End If
				strOName = rsUser2(0)
				rsUser2.Close
				Set rsUser = Nothing
				Err.Clear


				End If
				If strOName <> "" Then
								tmpRedObj = ocdQuotePrefix & strOName & ocdQuoteSuffix & "." 
End If
				tmpRedObj = tmpRedObj & ocdQuotePrefix & GetSQLIDFPart(Request.Form("AlterTableName"), "SQLOBJECTName", ocdquoteprefix, ocdquotesuffix) & ocdQuoteSuffix
				
					Response.Redirect "Structure.asp?sqlfrom=" & Server.URLEncode(tmpRedObj)
					Response.end
			End If
			End If
'''''''''''''EDIT TRIG
		Case "edittrig"
			If Request.Form("DoAction") = "" Then
				Call WriteHeader ("")
				%>
<span class=INFORMATION>Tables : <A HREF=Structure.asp?sqlfrom=<%=server.urlencode(Request.QueryString("SQLFrom"))%>><%=Server.HTMLEncode(Request("sqlfrom"))%></A> : <%=Request("trigname")%></span>
<form action="<%=request.servervariables("SCRIPT_NAME")%>?<%=Request.QueryString%>" method=post>
<input type=HIDDEN name=sqlfrom value="<%=Server.HTMLEncode(Request("sqlfrom"))%>">
<input type=HIDDEN name=trigname value="<%=Server.HTMLEncode(Request("trigname"))%>">
<strong>Trigger SQL Text:</strong> <br>
<textarea name=SQLCommandText COLS=60 ROWS=10><%
Response.Write Server.HTMLEncode(objSQL.GetHelpText( Request.QueryString("trigname")))
%></textarea>
<br>
<input type=Submit class=submit name=DoAction value="Update Trigger">
<input type=Submit class=submit name=Cancel value="Cancel">

<p>On update, the existing trigger will be dropped and the above statement will
be executed to create a new trigger.  This may affect permissions.
</form>
<%
					Call WriteFooter("")
				Else

					strSQLStufftr = "DROP TRIGGER " & Request("trigname")
					strViewName2tr = (Request("sqlfrom"))
					objSQL.ADOCOnnection.execute strSQLStufftr
					If Err.Number <> 0 Then
						Call WriteHeader("")
						Call WriteFooter("")
					Else
						objSQL.ADOConnection.execute Request.Form("SQLCommandText")
						If Err.Number <> 0 Then
							Call WriteHeader("")
							Call WriteFooter("")
						End If
					End If
					Response.Redirect "Structure.asp?sqlfrom=" &  Server.UrlEncode(Request("sqlfrom"))

				End If
				
		CASE "scriptobject"
			on error resume next
			If Request.Form("DoAction") = "" Then
				Call WriteHeader("")
%>
<span class=INFORMATION>T-SQL for Object <%=Server.HTMLEncode(Request("sqlfrom"))%></span>
<form action="<%=request.servervariables("SCRIPT_NAME")%>?<%=Request.QueryString%>" method=post>
<input type=HIDDEN name=sqlfrom value="<%=Server.HTMLEncode(Request("sqlfrom"))%>">
<br>
<textarea name=SQLCommandText COLS=50 ROWS=7><%
				Set dmosql = Server.CreateObject("SQLDMO.SQLServer")
				Set dmodb = Server.CreateObject("SQLDMO.Database")
				If Err.Number <> 0 Then
					Response.Write "This feature requires SQLDMO to be installed on the webserver"
					Err.Clear()
				Else
					strSV = Cstr(connDesign.Properties("Data Source"))
					strDB = Cstr(connDesign.Properties("Initial Catalog"))
					strOo = objSQL.GetSQLObjectOwner
					strON = objSQL.GetSQLObjectName
					If UCase(connDesign.Properties("Integrated Security")) ="SSPI" Then
						dmosql.loginsecure = true
						dmosql.connect  strSV
					Else
						dmosql.connect  strSV, ndnscSQLUser , ndnscSQLPass
					End If
					Set dmodb = dmosql.Databases(strDB, stroo)
					Err.Clear()
					strSS = dmodb.Views(strON).Script
					If Err.Number <> 0 Then
						Err.Clear()
						strSS = ""
						strSS = dmodb.Tables(strON).Script
					End If	
					If Err.Number <> 0 Then
					Else
						Response.Write err.description
						Response.Write Server.HTMLEncode(strSS)
					End If
					Err.Clear()
					set dmodb = nothing
					dmosql.DisConnect

					Set dmosql = Nothing
					Err.Clear()
				End If
				Response.Write("</textarea><br></form>")
				Call WriteFooter("")	
			Else
				If Err.Number <> 0 Then
					Call WriteHeader("")
					Call WriteFooter("")
				Else
					If Err.Number <> 0 Then
						Call WriteHeader("")
						Call WriteFooter("")
					Else
					End If
				End If
				Response.Write  "<p><strong>redirect 6</strong></P>"
				Response.end
			End If
					

''''''''''''''EDIT VIEW

		Case "editview"
			If Request.Form("DoAction") = "" Then
				Call WriteHeader("")
				Response.Write("<span class=INFORMATION><A HREF=Schema.asp?show=views>Views</A> : ")
				Response.write(Request("sqlfrom"))
				Response.Write("</span><form action=""")
				Response.write(request.servervariables("SCRIPT_NAME"))
				Response.write("?")
				Response.Write(Request.QueryString)
				Response.Write(""" method=post><input type=HIDDEN name=sqlfrom value=""")
				Response.Write(Server.HTMLEncode(Request("sqlfrom")))
				Response.Write("""><strong>SQL Text:</strong> <br><textarea name=SQLCommandText cols=50 rows=7>")
				Response.Write Server.HTMLEncode(replace(objSQL.GetHelpText(Request.QueryString("sqlfrom")),"CREATE VIEW","ALTER VIEW",1,1,1))
				Response.Write("</textarea><br><input type=Submit class=submit name=DoAction value=""Update Query""><input type=Submit class=submit name=Cancel value=""Cancel""><p></form>")
				Call WriteFooter("")	
			Else
				objSQL.ADOConnection.execute Request("SQLCommandText")
				If Err.Number <> 0 Then
					Call WriteHeader("")
					Response.Write Server.HTMLEncode(Request("SQLCommandText"))
					Call WriteFooter("")
				Else
					Response.Redirect("Schema.asp?show=views")
					Response.End()
				End If
				Response.Write "<p><strong>redirect 6</strong></P>"
				Response.End()
			End If
		End Select
		Response.Write strSQL
		Response.End()
		If strSQL <> "" Then
			set rsXXTemp = connDesign.Execute (strSQL)
		End If
		If Err.Number <> 0 Then
			Call WriteHeader("")
			Response.Write Server.HTMLEncode(strSQL)
			Call WriteFooter("")
		End If
		Response.Redirect("Structure.asp?sqlfrom=" & objSQL.SQLObject)
	Case Else
End Select
If Request("ocdAction") <> "" Then
	Call WriteFooter("")
End If
%>
