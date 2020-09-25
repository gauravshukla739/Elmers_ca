<% ' Except for @ Directives, there should be no ASP or HTML codes above this line
' Setting LANGUAGE = VBScript.Encode enables mixing of encoded and unencoded ASP scripts

'1 Click DB Pro copyright 1997-2003 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com
'Email inquiries to info@1ClickDB.com
        
'**Start Encode**

%>
<!--#INCLUDE FILE=PageInit.asp-->
<!--#INCLUDE FILE=ocdFunctions.asp-->
<%

Dim strAlterTableName, strAction, strDSN, connDB, blnDemo, rsTemp, strStatus, strNewDBName, strRelatedFrom, strRelatedTo, kyForeign, cmdTemp, strViewName, catADOX, fldTemp, OpenIdentifier, CloseIdentifier, strFieldName, strSQL, strNewTableKeyName, strNewTableKeyType, strNewTableKeySize, strNewFieldName, strNewFieldType, strNewFieldSize, intTemp, strTemp, strDefaultType, rsRelatedTable, intFieldCount, rsSchema, intIDXFCount
If Not ocdAccessTableEdits Or ocdReadOnly then
	Call WriteHeader("")
	Call WriteFooter("Permission Denied")
	Response.End
End If
		OpenIdentifier = "["
		CloseIdentifier = "]"
If Request.QueryString("DBCompact") <> "" Or Request.Form("ndbtnCompact") <> "" Then
	Call WriteHeader("")
	Call WriteFooter("Online Compacting not supported")
	Response.End
End If
If Request.QueryString("DBCreate") <> "" Then
	If ocdIsHome then
		Response.Clear
		Call WriteHeader("")
		Call WriteFooter("Feature Disabled in this Demonstration")
		Response.End
	End If
	If ocdADOConnection <> "" Then
		Response.Clear
		Response.Redirect "Schema.asp"
	End If
	strNewDBName = Request.Form("DBtocreate")
	strStatus = ""
	If Request.Form("ndbtnCancel") <> "" Then
		Response.Clear

		Response.Redirect ("Connect.asp")
	End If
	If Request.Form("ndbtnCreate") <> "" Then
		If ocdIsHome  Then
				Call WriteHeader("")
				Call WriteFooter("Disabled In Demo")
				Response.End
				strStatus = "Creating new Access Databases has been disabled in this demo."		
		End If
		If strStatus = "" Then
			If strNewDBName = "" Then
				strStatus = "Enter a Database to Create"
			End If	
		End If
		If strStatus = "" Then
			If Left(strNewDBName,1) = "/" Then
				strNewDBName = Server.Mappath(strNewDBName)
				If Err.Number <> 0 Then
					strStatus = Err.Description
				End If
			End If
		End If
		If strStatus = "" Then
			Set catADOX = Server.CreateObject("ADOX.Catalog" )
			If Err.Number <> 0 Then
				strStatus = ("<B>An ADOX catalog could not be created on the server.  This may be a permissions problem or the object may not be installed.</b>")
			End If
		End If
		If strStatus = "" Then
			Call catADOX.Create ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strNewDBName)
			If Err.Number <> 0 Then
				strStatus = Err.Description
				Set catADOX = nothing
			End If
		End If
		If strStatus = "" Then
			Set catADOX = nothing
			ndnscSQLuser = ""
			ndnscSQLPass = ""
			ndnscSQLConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strNewDBName & ";"
			Session("ocdSQLUser") = ndnscSQLuser
			Session("ocdSQLPass") = ndnscSQLPass
			Session("ocdSQLConnect") = ndnscSQLConnect
			Response.Redirect ("Schema.asp")
		End If
		Err.clear
	End If
	Call WriteHeader("")
	Response.Write ("<CENTER><FORM ACTION=""")
	Response.Write (request.servervariables("SCRIPT_NAME"))
	Response.Write ("?DBCreate=True"" METHOD=post>")
	If Request.Form <> "" Then
		If strStatus <> "" Then	
			Response.Write ("<table><tr><td><img src=appWarning.gif alt=Warning></td><td><SPAN CLASS=PageStatus>")
			Response.Write (strStatus)
			Response.Write ("</SPAN></td></tr></table>")
		End If
		Response.Write ("<p>")
	End If
	Response.Write DrawDialogBox("DIALOG_START","Create New Access Database","")
	Response.Write ("<table><tr><td NOWRAP valign=top><SPAN CLASS=FieldName>Path to MDB:</SPAN></TD><td><input Name=DBtocreate SIZE=50 VALUE=""")
	Response.Write (Server.HTMLEncode(Request.Form("DBtocreate")))
	Response.Write ("""></TD></TR><tr><td></td><td><input TYPE=Submit CLASS=submit Name=ndbtnCreate Value=""Create""><input TYPE=Submit Name=ndbtnCancel CLASS=submit Value=""Cancel""></td></tr></table></TD></TR><tr><td COLSPAN=2>")
	Response.Write DrawDialogBox("DIALOG_END","","")
	Response.Write ("</FORM></CENTER>")
	Call WriteFooter("")
End If
Set connDB = Server.CreateObject("ADODB.Connection")
connDB.Open ndnscSQLConnect, ndnscSQLUser, ndnscSQLPass
If ocdAccessTableEdits And (connDB.Properties("DBMS Name") = "MS Jet" Or connDB.Properties("DBMS Name") = "ACCESS") Then
Else
	Call WriteHeader("")
	Call WriteFooter("UNKNOWN PROVIDER")
	Response.End
End If
If Err.number <> 0 Then
	Call WriteHeader("")
	Call WriteFooter("")
	Response.End
End If
strDSN = Cstr(connDB.Properties("Data Source Name"))
blnDemo = IsNumeric(mid(strDSN,instrrev(strDSN,"\")+1, len(strDSN)- (3+(instrrev(strDSN,"\")+1))))
err.clear
If Request("table") <>"" then
	strAlterTableName = Request("table")
Else
	strAlterTableName = Request("altertablename")
End If
If Request("ocdAction") <> "" Then
	Call WriteHeader("")
End If
Select Case UCASE(Request("ocdAction"))
	Case "ADDRELATION"
		Select Case UCASE(Request("whatrelation"))
			Case "MANY-TO-ONE"
				strRelatedFrom = Request("SQLFrom")
				strRelatedTo = Request("RelatedTo")
			Case Else
				strRelatedTo  = Request("SQLFrom")
				strRelatedFrom = Request("RelatedTo")
		End Select
		Response.Write("<SPAN CLASS=INFORMATION>Add New Relation to Database</span>")
		Response.Write("<FORM action=" & ocdPageName & "?sqlfrom=" & server.urlencode(strRelatedFrom)& " method=post>")
		Response.Write("<input NAME=relatedto TYPE=Hidden VALUE=""")
		Response.Write(Server.HTMLEncode(strRelatedTo))
		Response.Write("""><b>Name: </b><input TYPE=Text NAME=sqlrelationname VALUE="""" SIZE=40><BR><table BORDER=1><tr><TH>Primary Key in ")
		Response.Write(Server.HTMLEncode(strRelatedTo))
		Response.Write("</th><TH>Related Column in ")
		Response.Write(Server.HTMLEncode(strRelatedFrom))
		Response.Write("</th></tr>")
		Set rsRelatedTable = Server.CreateObject("ADODB.Recordset")
		Set rsSchema = Server.CreateObject("ADODB.Recordset")
		Set rsSchema = connDB.openSchema(12) 'indexes
		intFieldCount = 0
		rsRelatedTable.open "SELECT * from [" & strRelatedFrom & "] WHERE 1=2", connDB
		Do While Not rsSchema.EOF
			If UCASE(rsSchema("table_name")) = UCase(strRelatedTo) and rsSchema("primary_key") = True Then
				intFieldCount = intFieldCount + 1
				Response.Write "<tr>"
				Response.Write "<td>"
				Response.Write rsSchema("COLUMN_NAME")
				Response.Write "<input TYPE=Hidden NAME=""pkfld"
				Response.Write intFieldCount
				Response.Write """ VALUE="""
				Response.Write rsSchema("COLUMN_NAME")
				Response.Write """>"
				Response.Write "</td>"
				Response.Write "<td>"
				Response.Write "<SELECT NAME=""fkfld"
				Response.Write intFieldCount
				Response.Write """>"
				For Each fldTemp in rsRelatedTable.fields
					Response.Write "<OPTION VALUE="""
					Response.Write Server.HTMLEncode(fldTemp.Name)
					Response.Write """>"
					Response.Write Server.HTMLEncode(fldTemp.Name)
					Response.Write "</OPTION>"
				Next
				Response.Write "</select>"
				Response.Write "</td>"
				Response.Write "</tr>"
			End If
			rsSchema.MoveNext
		Loop
		Response.Write("</table><input NAME=reldeleterule TYPE=checkbox> Cascade Delete<BR><input NAME=relupdaterule TYPE=checkbox> Cascade Update<p><input TYPE=Submit CLASS=submit Name=Action Value=""Create Relation""><p></FORM>")
	Case "CREATEVIEW"
		Response.Write("<SPAN CLASS=INFORMATION>Add New View</span><FORM action=" & ocdPageName & " method=post><b>Name : </b><input TYPE=Text NAME=table VALUE=""" & Server.HTMLEncode(Request("newviewname")) & """ SIZE=40  MAXLENGTH=255><p><b>SQL Text : </b> <BR><TEXTAREA Name=SQLCommandText COLS=50 ROWS=7>")
		Response.Write(Server.HTMLEncode(Request("SQLCommandText")))
		Response.Write("</TEXTAREA><p><input TYPE=Submit CLASS=submit Name=Action Value=""Save View""><p>Saving SQL Text containing syntax or spelling errors will cause MS Access to automatically reclassify a new or existing View as a Procedure of the same name. Access Views and Procedures created from ADOX may not be visible from the Database Window Queries List.</FORM>")
	Case "CREATEPROC"
		Response.Write("<SPAN CLASS=INFORMATION>Add New Procedure to Database</span><FORM action=" & ocdPageName & " method=post><b>Name: </b><input TYPE=Text NAME=sqlfrom VALUE="""" SIZE=40><BR><b>Procedure SQL Text:</b> <BR><TEXTAREA Name=SQLCommandText COLS=50 ROWS=7></TEXTAREA><BR><input TYPE=Submit CLASS=submit Name=Action Value=""Create Procedure""><p><p>Saving SQL Text containing with no parameters may cause MS Access to automatically reclassify a new or existing Procedure as a View of the same name. Access Views and Procedures created from ADOX may not be visible from the Database Window Queries List.</FORM>")
	Case "EDITVIEW"
		Response.Write("<SPAN CLASS=INFORMATION>Views : " & Request("sqlfrom") & " : Edit SQL Text</span><FORM action=" & ocdPageName & " method=post>")
		Set catADOX = Server.CreateObject("adox.catalog")
		Set catADOX.ActiveConnection = connDB
		If Err.Number <> 0 then
			Call WriteFooter("This function requires ADOX to be installed on your web server.")
		Else
			Response.Write("<input TYPE=HIDDEN NAME=sqlfrom VALUE=""" & Server.HTMLEncode(Request("sqlfrom")) & """><b>View SQL Text:</b> <BR><TEXTAREA Name=SQLCommandText COLS=50 ROWS=7>")
strViewName = Replace(Replace((Request("sqlfrom")),"]",""),"[","")
			Response.Write Server.HTMLEncode(catADOX.views(strViewName).Command.CommandText) 
			Response.Write("</TEXTAREA><BR><input TYPE=Submit CLASS=submit Name=Action Value=""Update Query"">")
		End If
		Response.Write("</FORM>")
		Call WriteFooter("")
	Case "EDITPROC"
		Response.Write "<SPAN CLASS=INFORMATION>Procedures : " & Request("sqlfrom") & "</span> <a href=""DBDesignMSAccess.asp?Action=dropproc&amp;sqlfrom=" &  Server.URLEncode(Request("sqlfrom")) & """>Drop Procedure</a>"
		Response.Write "<FORM action=" & ocdPageName & " method=post>"
		strViewName = (Request("sqlfrom"))
		Set catADOX = Server.CreateObject("adox.catalog")
		Set catADOX.ActiveConnection = connDB
		If Err.Number <> 0 then
			Response.Write "This function requires ADOX to be installed on your web server."
			Err.Clear
		Else
			Response.Write "<input TYPE=HIDDEN NAME=sqlfrom VALUE=""" &Server.HTMLEncode(Request("sqlfrom")) & """>"
			Response.Write "<b>Procedure SQL Text:</b> <BR>"
			Response.Write "<TEXTAREA Name=SQLCommandText COLS=50 ROWS=7>"
			Response.Write Server.HTMLEncode(catADOX.procedures(strViewName).Command.CommandText) 
			Response.Write "</TEXTAREA>"
			Response.Write "<BR>"
			Response.Write "<input TYPE=Submit CLASS=submit Name=Action Value=""Update Procedure"">"
			Response.Write("<p>Saving SQL Text containing syntax or spelling errors will cause MS Access to automatically reclassify a new or existing View as a Procedure of the same name. Access Views and Procedures created from ADOX may not be visible from the Database Window Queries List.  ")
		End If
		Response.Write "</FORM>"
	Case "NEWINDEX"
		Response.Write("<span class=""information"">Create Index on Table  ")
		Response.Write(Server.HTMLEncode(Request("sqlfrom")))
		Response.Write("</span><form action=""" & ocdPageName & """ method=post><input type=""hidden"" name=""sqlfrom"" VALUE=""")
		Response.Write(Server.HTMLEncode(Request("sqlfrom")))
		Response.Write("""><span class=""information"">Name:</span> <input type=""text"" name=""NewIndexName""><select name=""IndexUnique""><option value=""""></option><option value=""Unique"">Unique</option></select><select name=""IndexOptions""><option value=""""></option><option value=""Primary"">Primary Key</option><option value=""Disallow Null"">Disallow Null</option><option value=""Ignore Null"">Ignore Null</option></select><input type=""submit"" class=""submit"" name=""Action"" value=""Create Index""><p>")
		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		rsTemp.open "Select * from [" & Replace(Replace(Request("sqlfrom"),"]",""),"[","") & "] WHERE 1=2", connDB ', 3 'adOpenStatic
		strTemp = "<option value=""""></option>"
		For Each fldTemp in rsTemp.fields
			strTemp = strTemp & "<option value=""" & Server.HTMLEncode(fldTemp.Name) & """>" & Server.HTMLEncode(fldTemp.Name) & "</option>"
		next
		Response.Write("<table>")
		For intIDXFCount = 1 to 10
			Response.Write("<tr><td><span class=""fieldName"">Field " & Trim(Cstr(intIDXFCount)) & " :</span></td><td><select name=""idxf" & Trim(Cstr(intIDXFCount)) & """>")
			Response.Write(strTemp)
			Response.Write("</select></td><td><select name=""idxf" & Trim(Cstr(intIDXFCount)) & "sort""><option value=""ASC"">Ascending</option><option value=""DESC"">Descending</option></select></td></tr>")
		Next
		Response.Write("</table></form>")
	Case "NEWTABLE"
		Response.Write("<SPAN CLASS=Information>Add New Table:</span><FORM action=""" & ocdPageName & """ method=post><b>Name:</b> <input TYPE=TEXT NAME=AlterTableName SIZE=34 maxlength=64 Value=""NewTable""><BR><b>Primary Key:</b> <input TYPE=TEXT NAME=AddNewTableKeyName Size=24 maxlength=64 Value=""RecordID""><SELECT Name=AddNewTableKeyType><OPTION VALUE='Autonumber' selected>Autonumber<OPTION VALUE='Currency'>Currency<OPTION VALUE='Date'>Date/Time<OPTION VALUE='Double'>Double<OPTION VALUE='Integer'>Integer<OPTION VALUE='Long'>Long Integer<OPTION VALUE='Memo'>Memo<OPTION VALUE='Text'>Text</SELECT><b>Size:</b> <input TYPE=TEXT NAME=AddNewTableKeySize Size=3 maxlength=3><BR><input  CLASS=submit  TYPE=SUBMIT NAME='Action' VALUE='Create New Table'></FORM><p>")
	Case "NEWFIELD"
		Response.Write("<SPAN CLASS=Information>Add Field To Table ")
		Response.Write(Server.HTMLEncode(Request.QueryString("sqlfrom")))
		Response.Write(":</SPAN><FORM action=""")
		Response.Write(ocdPageName)
		Response.Write(""" method=post><b>Name:</b> <input TYPE=TEXT NAME=NewFieldName Size=24 maxlength=64 Value=""NewField""><SELECT Name=NewFieldType><OPTION VALUE='Autonumber' >Autonumber<OPTION VALUE='BIT'>Boolean<OPTION VALUE='Currency'>Currency<OPTION VALUE='Date'>Date/Time<OPTION VALUE='Double'>Double<OPTION VALUE='Integer'>Integer<OPTION VALUE='Long'>Long Integer<OPTION VALUE='Memo'>Memo<OPTION VALUE='Text' selected>Text</SELECT><b>Size:</b> <input TYPE=TEXT NAME=NewFieldSize Size=3 maxlength=3 value=""50""><input TYPE=HIDDEN NAME=altertablename  value=""")
		Response.Write(Server.HTMLEncode(Request.QueryString("sqlfrom")))
		Response.Write(""">")
		If 1=2 Then 'NOT NULL JET SQL No longer works
		Response.Write("<SELECT NAME=allownulls><OPTION VALUE=""NULL""></OPTION><OPTION VALUE=""NOTNULL"" SELECTED>Required</OPTION></select>")
		End If
		If connDB.provider ="Microsoft.Jet.OLEDB.4.0" Then
			Response.Write "<BR><b>Default:</b> <input NAME=fdefault TYPE=Text> Use double quotes to enclose text values"
		End If
		Response.Write("<BR><input TYPE=SUBMIT  CLASS=submit  NAME='Action' VALUE='Create New Field'></FORM><p><SPAN CLASS=Notice>To rename or change the required property on an Access field, first create a new field then copy the data and remove the old field.</SPAN><p>")
	Case "EDITFIELD"
		Select Case CLng(Request("defType"))
			Case 0 'adEmpty
				strDefaultType = "Empty"
			Case 16 'adTinyInt 
				strDefaultType = "INTEGER"
			Case 2 'adSmallInt
				strDefaultType = "INTEGER"
			Case 3 'adInteger 
				strDefaultType = "LONG"
			Case 20 'adBigInt 
				strDefaultType =  "LONG"
			Case 17 'adUnsignedTinyInt 
				strDefaultType =  "UnsignedTinyInt"
			Case 18 'adUnsignedSmallInt
				strDefaultType =  "UnsignedSmallInt"
			Case 19 'adUnsignedInt
				strDefaultType =  "UnsignedInt"
			Case 21 'adUnsignedBigInt
				strDefaultType = "UnsignedBigInt"
			Case 4 'adSingle
				strDefaultType = "Single"
			Case 5 'adDouble
				strDefaultType = "Double"
			Case 6 'adCurrency
				strDefaultType = "CURRENCY"
			Case 14 'adDecimal
				strDefaultType ="DOUBLE"
			Case 131 'adNumeric
				strDefaultType = "DOUBLE"
			Case 11 'adBoolean
				strDefaultType = "BOOLEAN"
			Case 10 'adError
				strDefaultType = "Error"
			Case 132 'adUserDefined
				strDefaultType = "UserDefined"
			Case 12 'adVariant
				strDefaultType = "Variant"
			Case 9 'adIDispatch
				strDefaultType ="IDispatch"
			Case 13 'adIUnknown
				strDefaultType ="IUnknown"
			Case 72 'adGUID
				strDefaultType = "GUID"
			Case 7 'adDate
				strDefaultType = "DATE"
			Case 133 'adDBDate
				strDefaultType = "DATE"
			Case 134 'adDBTime
				strDefaultType = "DATE"
			Case 135 'adDBTimeStamp
				strDefaultType = "DATE"
			Case 8 'adBSTR
				strDefaultType = "TEXT"
			Case 129 'adChar
				strDefaultType = "TEXT"
			Case 200 'adVarChar
				strDefaultType = "TEXT"
			Case 201 'adLongVarChar
				strDefaultType = "MEMO"
			Case 130 'adWChar
				strDefaultType = "TEXT"
			Case 202 'adVarWChar
				strDefaultType = "TEXT"
			Case 203 'adLongVarWChar
				strDefaultType = "MEMO"
			Case 128 'adBinary 
				strDefaultType = "Binary"
			Case 204 'adVarBinary
				strDefaultType = "VarBinary"
			Case 205 'adLongVarBinary
				strDefaultType = "LongVarBinary"
			Case else
'				strDefaultType = CSTR(fldTemp.Type)
		end select
		Response.Write "<SPAN CLASS=Information>"
		Response.Write "Tables : "
		Response.Write Server.HTMLEncode(Request.QueryString("table")) 
		Response.Write " : "
		Response.Write Server.HTMLEncode(Request("fieldname"))
		Response.Write "</SPAN>"
		Response.Write "<FORM action=" & ocdPageName & " method=post>"
		Response.Write "<input TYPE=HIDDEN NAME=NewFieldName Value="""
		Response.Write Server.HTMLEncode(Request("fieldname"))
		Response.Write """> "
		Response.Write "<table BORDER=1><tr><td>"
		Response.Write "<B>Type: </b>"
		Response.Write "</TD><td>"
		Response.Write " <SELECT Name=NewFieldType><OPTION VALUE='Autonumber' "
		Select Case UCASE(strDefaultType)
			Case "AUTONUMBER"
				Response.Write " SELECTED "
		End select
		Response.Write ">Autonumber"
		Response.Write "<OPTION VALUE='BIT' "
		Select Case strDefaultType
			Case "BOOLEAN"
				Response.Write " SELECTED "
		End select
		Response.Write ">Boolean"
		Response.Write "<OPTION VALUE='Currency' "
		Select Case strDefaultType
			Case "CURRENCY"
				Response.Write " SELECTED "
		End select
		Response.Write ">Currency"
		Response.Write "<OPTION VALUE='Date' "
		Select Case strDefaultType
			Case "DATE"
				Response.Write " SELECTED "
		End select
		Response.Write ">Date/Time<OPTION VALUE='Double' "
		Select Case strDefaultType
			Case "DOUBLE"
				Response.Write " SELECTED "
		End select
		Response.Write ">Double<OPTION VALUE='Integer' "
		Select Case strDefaultType
			Case "INTEGER"
				Response.Write " SELECTED "
		End select
		Response.Write ">Integer<OPTION VALUE='Long' "
		Select Case strDefaultType
			Case "LONG"
				Response.Write " SELECTED "
		End select
		Response.Write ">Long Integer<OPTION VALUE='Memo' "
		Select Case strDefaultType
			Case "MEMO"
				Response.Write " SELECTED "
		End select
		Response.Write ">Memo<OPTION VALUE='Text' "
		Select Case strDefaultType
			Case "TEXT"
				Response.Write " SELECTED "
		End select
		Response.Write ">Text</SELECT>"
		Response.Write "</TD></TR><tr><td>"
		Response.Write "<b>Size:</b></TD><td><input TYPE=TEXT NAME=NewFieldSize Size=3 maxlength=3 value="""
		Response.Write Request("defsize")
		Response.Write """></TD></TR><tr><td><B>Required:</B><input TYPE=HIDDEN NAME=altertablename  value=""" 
		Response.Write Server.HTMLEncode(Request.QueryString("table")) 
		Response.Write """></TD><td>"
		If CDbl(connDB.Version) > CDbl("2.5") Then
			If Request("nulldef") = "NOTNULL" Then
				Response.Write " " & True & " "
			Else
				Response.Write " " & False & " "
			End If
		Else
			If Request("nulldef") = "NOTNULL" Then
				Response.Write " " & True & " "
			Else
				Response.Write " <SELECT NAME=allownulls><OPTION VALUE=""NULL"" "
				If Request("nulldef") = "NULL" Then
					Response.Write " SELECTED "
				End If
				Response.Write ">" & False & "</OPTION><OPTION VALUE=""NOTNULL"" "
				If Request("nulldef") = "NOTNULL" Then
					Response.Write " SELECTED "
				End If
				Response.Write ">" & True & "</OPTION></select>"
			End If
		End If
		If connDB.provider ="Microsoft.Jet.OLEDB.4.0" Then
			Response.Write "</TD></TR><tr><td valign=top><b>Default:</b></TD><td valign=top><input TYPE=TEXT NAME=fdefault VALUE="""
			Set rsTemp = Server.CreateObject("ADODB.Recordset")
			Set rsTemp = connDB.OpenSchema(4) ',array(Empty,Empty,(Request.QueryString("table")),(Request("fieldname"))))
			do while not rsTemp.eof 
			 	If FormatForSQL(Request.QueryString("table"), "Access", "RemoveSQLIdentifier") = rsTemp("TABLE_NAME") Then
					If FormatForSQL(Request("fieldname"), "Access", "RemoveSQLIdentifier") = rsTemp("COLUMN_NAME") Then
						If not isnull(rsTemp("COLUMN_DEFAULT")) Then
							Response.Write Server.HTMLEncode(rsTemp("COLUMN_DEFAULT"))
						End If
					End If
				End If
				rsTemp.movenext
			loop
			Response.Write """> <BR><SPAN CLASS=Notice>Use double quotes to enclose literal text values</SPAN></TD></TR>"
		End If
		Response.Write "</TABLE>"
		Response.Write "<BR><input TYPE=SUBMIT  CLASS=submit NAME='Action' VALUE='Alter Field'></FORM><p>"
		Response.Write "<SPAN CLASS=Notice>To rename or change the required property on an Access field, first create a new field then copy the data and remove the old field.</SPAN><p>"
	Case ""

		strNewTableKeyName = Request.Form("addnewtablekeyname")
		strNewTableKeyType = Request.Form("addnewtablekeytype")
		strNewTableKeySize = Request.Form("addnewtablekeysize")
		strFieldName = Request.QueryString("fieldname")
		strNewFieldName = Request.Form("newfieldname")
		strNewFieldType = Request.Form("newfieldtype")
		strNewFieldSize = Request.Form("newfieldsize")
		If Request.QueryString("action") <> "" Then
			strAction = Request.QueryString("action")
		Else
			strAction = Request.Form("action")
		End If
		Set rsTemp=Server.CreateObject("ADODB.Recordset")
		Select Case strAction
			Case "Create Index"
				strSQL = "CREATE "
				If Request("IndexUnique") <> "" Then
					strSQL = strSQL & " UNIQUE "
				End If
				strSQL = strSQL & " INDEX ["
				strSQL = strSQL & Request("NewIndexName")
				strSQL = strSQL & "] "
				strSQL = strSQL & " ON ["
				strSQL = strSQL & Replace(Replace(Request("sqlfrom"),"[",""),"]","")
				strSQL = strSQL & "] ("
				If Request("idxf1") <> "" Then
					strSQL = strSQL & "[" & Request("idxf1") & "]"
					strSQL = strSQL & " " & Request("idxf1sort") & ","
				End If
				If Request("idxf2") <> "" Then
					strSQL = strSQL & "[" & Request("idxf2") & "]"
					strSQL = strSQL & " " & Request("idxf2sort") & ","
				End If
				If Request("idxf3") <> "" Then
					strSQL = strSQL & "[" & Request("idxf3") & "]"
					strSQL = strSQL & " " & Request("idxf3sort") & ","
				End If
				If Request("idxf4") <> "" Then
					strSQL = strSQL & "[" & Request("idxf4") & "]"
					strSQL = strSQL & " " & Request("idxf4sort") & ","
				End If
				If Request("idxf5") <> "" Then
					strSQL = strSQL & "[" & Request("idxf5") & "]"
					strSQL = strSQL & " " & Request("idxf5sort") & ","
				End If
				If Request("idxf6") <> "" Then
					strSQL = strSQL & "[" & Request("idxf6") & "]"
					strSQL = strSQL & " " & Request("idxf6sort") & ","
				End If
				If Request("idxf7") <> "" Then
					strSQL = strSQL & "[" & Request("idxf7") & "]"
					strSQL = strSQL & " " & Request("idxf7sort") & ","
				End If
				If Request("idxf8") <> "" Then
					strSQL = strSQL & "[" & Request("idxf8") & "]"
					strSQL = strSQL & " " & Request("idxf8sort") & ","
				End If
				If Request("idxf9") <> "" Then
					strSQL = strSQL & "[" & Request("idxf9") & "]"
					strSQL = strSQL & " " & Request("idxf9sort") & ","
				End If
				If Request("idxf10") <> "" Then
					strSQL = strSQL & "[" & Request("idxf10") & "]"
					strSQL = strSQL & " " & Request("idxf10sort") & ","
				End If
				strSQL = Left(strSQL, len(strSQL)-1)
				strSQL = strSQL & ")"
				If Request("IndexOptions") <> "" Then
					strSQL = strSQL & " WITH "
					strSQL = strSQL & Request("IndexOptions")
				End If
				Call notforsample()
				connDB.Execute (strSQL)
				If err<>0 then
					Call WriteHeader("")
					Response.Write strSQL
					Call WriteFooter("")
				End If
				Response.Redirect ("Structure.asp?sqlfrom=" & server.URLEncode(Request("sqlfrom")))
			Case "Create New Table"
				Select Case UCASE(strNewTableKeyType)
					Case "AUTONUMBER"
						strSQL= "CREATE TABLE " & OpenIDentifier & strAlterTableName & CloseIdentifier & "(" & OpenIdentifier & strNewTableKeyName & CloseIdentifier & " COUNTER NOT NULL CONSTRAINT " & OpenIdentifier & "pk" & strAlterTableName & CloseIdentifier & " PRIMARY KEY)"
					Case "CURRENCY"
						strSQL= "CREATE TABLE " & OpenIdentifier & strAlterTableName & CloseIdentifier & " (" & OpenIdentifier & strNewTableKeyName & CloseIdentifier & " CURRENCY NOT NULL CONSTRAINT " & OpenIdentifier & "pk" & strAlterTableName & CloseIdentifer & " PRIMARY KEY)"
					Case "DATE"
						strSQL= "CREATE TABLE "& OpenIDentifier & strAlterTableName & CloseIdentifier & " (" & OpenIdentifier & strNewTableKeyName & CloseIdentifier & " DATETIME NOT NULL CONSTRAINT " & OpenIdentifier & "pk" & strAlterTableName & CloseIdentifier & " PRIMARY KEY)"
					Case "DOUBLE"
						strSQL= "CREATE TABLE " & OpenIdentifier & strAlterTableName & CloseIdentifier & " (" & OpenIdentifier & strNewTableKeyName & CloseIdentifier & " DOUBLE NOT NULL CONSTRAINT " & OpenIdentifier & "pk" & strAlterTableName & CloseIdentifier & " PRIMARY KEY)"
					Case "INTEGER"
						strSQL= "CREATE TABLE " & OpenIdentifier & strAlterTableName & CloseIdentifier & " (" & OpenIdentifier & strNewTableKeyName & CloseIdentifier & " SHORT NOT NULL CONSTRAINT " & OpenIdentifier & "pk" & strAlterTableName & CloseIdentifier & " PRIMARY KEY)"
					Case "LONG"
						strSQL= "CREATE TABLE " & OpenIdentifier & strAlterTableName & CloseIdentifier & " (" & OpenIdentifier & strNewTableKeyName & CloseIdentifier & " LONG NOT NULL CONSTRAINT " & OpenIdentifier & "pk" & strAlterTableName & CloseIdentifier & " PRIMARY KEY)"
					Case "MEMO"
						strSQL= "CREATE TABLE " & OpenIdentifier & strAlterTableName & CloseIdentifier & " (" & OpenIdentifier & strNewTableKeyName & CloseIdentifier & " LONGTEXT NOT NULL CONSTRAINT " & OpenIdentifier & "pk" & strAlterTableName & CloseIdentifier & " PRIMARY KEY)"
					Case "TEXT"
						strSQL= "CREATE TABLE " & OpenIdentifier & strAlterTableName & CloseIdentifier & " (" & OpenIdentifier & strNewTableKeyName & CloseIdentifier & " TEXT(" & strNewTableKeySize & ") NOT NULL CONSTRAINT " & OpenIDentifier & "pk" & strAlterTableName & CloseIdentifier & " PRIMARY KEY)"
				End select
			Case "Create Relation"
				Call notforsample()
				Set catADOX = Server.CreateObject("ADOX.Catalog")
				If err <> 0 Then
					err.clear
					Call WriteHeader("")
					Response.Write "<table><tr><td valign=top><img src=""appWarning.gif"" alt=""Warning""></td><td valign=top>"
					Response.Write "Could not create an ADOX catalog on the server.  Either this object is not installed or there is a problem with permissions."
					Response.Write "</td></tr></table>"
					Call WriteFooter("")
				End If
				Set kyForeign = Server.CreateObject("ADOX.Key")
		  	catADOX.ActiveConnection = connDB
		  	kyForeign.Name = Cstr(Request("sqlrelationname"))
		  	kyForeign.Type = 2 'adKeyForeign
		  	kyForeign.RelatedTable = Request("relatedto")		
				If Request("relupdaterule") <> "" Then
					kyForeign.UpdateRule = 1 'adRICascade
				End If
				If Request("reldeleterule") <> "" Then
					kyForeign.DeleteRule = 1 'adRICascade
				End If
				For intTemp = 1 To 12
					If Request("fkfld" & Cstr(intTemp)) <> "" Then
						kyForeign.Columns.Append Cstr(Request("fkfld" & Cstr(intTemp)))
						kyForeign.Columns(Cstr(Request("fkfld" & Cstr(intTemp)))).RelatedColumn = Cstr(Request("pkfld" & Cstr(intTemp)))
					End If
				Next
				catADOX.Tables(CSTR(Request("sqlfrom"))).Keys.Append kyForeign
				If err<>0 then
					Call WriteHeader("")
					Call WriteFooter("")
				End If
				Response.Redirect ("Structure.asp?sqlfrom=" & server.URLEncode(Request("sqlfrom")))
				Response.Write err.description
			Case "Create Query","Save View"
				Call notforsample()
				strSQL = Request.Form("SQLCommandText")
				strViewName = (Request("table"))
				Set catADOX = Server.CreateObject("adox.catalog")
				If err <> 0 Then
					err.clear
					Call WriteHeader("")
					Response.Write "<table><tr><td valign=top><img src=""appWarning.gif"" alt=""Warning""></td><td valign=top>"
					Response.Write "Could not create an ADOX catalog on the server.  Either this object is not installed or there is a problem with permissions."
					Response.Write "</td></tr></table>"
					Call WriteFooter("")
				End If
				Set catADOX.ActiveConnection = connDB
				Set cmdTemp = Server.CreateObject("adodb.command")
				Set cmdTemp.ActiveConnection = connDB
				cmdTemp.CommandType = &H0001 'adCmdText
				cmdTemp.CommandText = strSQL
				catADOX.views.append strViewName, cmdTemp
				If err <> 0 Then
					Call WriteHeader("")
					Call WriteFooter("")
				End If
				Response.Redirect "Structure.asp?sqlfrom=" & server.urlencode(Request("table"))
			Case "Create Procedure"
				Call notforsample()
				Set catADOX = Server.CreateObject("adox.catalog")
				If err <> 0 Then
					err.clear
					Call WriteHeader("")
					Response.Write "<table><tr><td valign=top><img src=""appWarning.gif"" alt=""Warning""></td><td valign=top>"
					Response.Write "Could not create an ADOX catalog on the server.  Either this object is not installed or there is a problem with permissions."
					Response.Write "</td></tr></table>"
					Call WriteFooter("")
				End If
				Set catADOX.ActiveConnection = connDB
				Set cmdTemp = Server.CreateObject("adodb.command")
				Set cmdTemp.ActiveConnection = connDB
				cmdTemp.CommandText = Cstr(Request.Form("SQLCommandText"))
			 catADOX.procedures.append Cstr(Request("sqlfrom")), cmdTemp
				If err <> 0 Then
					Call WriteHeader("")
					Call WriteFooter("")
				End If
				Response.Redirect "Schema.asp?show=procs"
			Case "Update Query"
				Call notforsample()
				strSQL = Request.Form("SQLCommandText")
				strViewName = Replace(Replace((Request("sqlfrom")),"]",""),"[","")
				Set catADOX = Server.CreateObject("adox.catalog")
				Set catADOX.ActiveConnection = connDB
				Set cmdTemp = Server.CreateObject("adodb.command")
				Set cmdTemp.ActiveConnection = connDB
				cmdTemp.CommandType = &H0001 'adCmdText
				cmdTemp.CommandText = strSQL
				catADOX.views(strViewName).Command = cmdTemp
				If err <> 0 Then
					Call WriteHeader("")
					Call WriteFooter("")
				End If
				Response.Redirect "Structure.asp?sqlfrom=" & Server.URLEncode(Request("sqlfrom"))
			Case "Update Procedure"
				Call notforsample()
				Set catADOX = Server.CreateObject("adox.catalog")
				Set catADOX.ActiveConnection = connDB
				Set cmdTemp = Server.CreateObject("adodb.command")
				Set cmdTemp.ActiveConnection = connDB
				cmdTemp.CommandText =  Cstr(Request.Form("SQLCommandText"))
				catADOX.procedures(Cstr(Request("sqlfrom"))).Command = cmdTemp 
				If err <> 0 Then
					Call WriteHeader("")
					Call WriteFooter("")
				End If
				Response.Redirect "Schema.asp?show=procs"
			Case "Create New Field"	, "Alter Field"
				Call notforsample()
				If strAction = "Alter Field" Then
					strSQL = "ALTER TABLE [" &  Replace(Replace(strAlterTableName,"]",""),"[","") & "] ALTER COLUMN "
				Else
					strSQL = "ALTER TABLE [" & Replace(Replace(strAlterTableName,"]",""),"[","") & "] ADD COLUMN "
				End If
				strSQL = strSQL & OpenIdentifier & strNewFieldName & CloseIdentifier 
				Select Case UCASE(strNewFieldType)
					Case "AUTONUMBER"
					 strSQL = strSQL & " COUNTER"
					Case "CURRENCY"
						strSQL = strSQL & " CURRENCY"
					Case "BIT"
						strSQL = strSQL & " BIT"
					Case "DATE"
						strSQL = strSQL & " DATETIME"
					Case "DOUBLE"
						strSQL = strSQL & " DOUBLE"
					Case "INTEGER"
						strSQL = strSQL & " SHORT"
					Case "LONG"
						strSQL = strSQL & " LONG"
					Case "MEMO"
						strSQL = strSQL & " LONGTEXT"
					Case "TEXT"
						strSQL = strSQL & " TEXT(" & strNewFieldSize & ")"
				end select
				'NOT NULL no longer works for JET SQL
				If 1=2 Then 'request("Action")  <> "Alter Field" Then
					If UCASE(Request("allownulls")) = "NULL" Then
						strSQL = strSQL & " NULL"
					Else
						strSQL = strSQL & " CONSTRAINT [rq" & strNewFieldName & "] NOT NULL"
					End If
				End If
				If connDB.provider ="Microsoft.Jet.OLEDB.4.0" Then
					If Request("fdefault") <> "" Then
						strSQL = strSQL & " DEFAULT " 
						strSQL = strSQL & Request("fdefault")
					Else
						'strSQL = strSQL & " NULL "
					End If
				End If
				Call notforsample()
				'response.write strSQL
				'response.end
				connDB.Execute strSQL
				If err <> 0 Then
					Response.Clear
					Call WriteHeader("")
					Response.Write strSQL
					Call WriteFooter("")
					Response.End
				End If
				Response.clear
				Response.Redirect ("Structure.asp?sqlfrom=" & server.URLEncode(strAlterTableName ))
				Response.End
			Case "deletefield"
				If Request.Form("cancel") <> "" Then
					Response.Redirect ("Structure.asp?sqlfrom=" & server.URLEncode(strAlterTableName ))
				ElseIf Request.Form("confirm") = "" Then
					Call WriteHeader ("")
					Response.Write("<FORM action=""")
					Response.Write(ocdPageName)
					Response.Write("?" & Request.QueryString)
					Response.Write(""" method=post>")
					Response.Write(drawdialogbox("DIALOG_START","Please Confirm",""))
					Response.Write("<table ><tr><td valign=top><img src=appWarning.gif alt=Warning></td><td>Are you sure you want to permanently drop the field ")
					Response.Write(Server.HTMLEncode(strFieldName))
					Response.Write(" in table ")
					Response.Write(Server.HTMLEncode(strAlterTableName))
					Response.Write(" ?  This action will affect all records in the table and cannot be undone.</td></tr></table><p><input TYPE=Submit NAME=Confirm  CLASS=submit Value=""OK"">&nbsp;<input TYPE=Submit  CLASS=submit  Name=Cancel Value=Cancel>")
					Response.Write(drawdialogbox("DIALOG_END","Please Confirm",""))
					Response.Write("</FORM>")
					Call WriteFooter ("")
				Else
					err.clear
					strSQL = "ALTER TABLE " & strAlterTableName & " DROP " & OpenIdentifier & strFieldName & CloseIdentifier & ""
				End If
			Case "deleterelation"
				strFieldName = Request("relationname")
				strAlterTableName = Request("sqlfrom")
				If Request.Form("cancel") <> "" Then
					Response.Redirect ("Structure.asp?sqlfrom=" & server.URLEncode(strAlterTableName ))
				ElseIf Request.Form("confirm") = "" Then
					Call WriteHeader ("")
					Response.Write("<FORM action=""")
					Response.Write(ocdPageName & "?" & Request.QueryString)
					Response.Write(""" method=post><table><tr><td valign=top><img src=appWarning.gif alt=Warning></td><td>Are you sure you want to permanently drop the relation ")
					Response.Write(Server.HTMLEncode(strFieldName))
					Response.Write(" in table ")
					Response.Write(Server.HTMLEncode(strAlterTableName))
					Response.Write(" ?  This action cannot be undone.</td></tr></table><p><input TYPE=Submit NAME=Confirm  CLASS=submit Value=""OK"">&nbsp;<input TYPE=Submit  CLASS=submit  Name=Cancel Value=Cancel></FORM>")
					Call WriteFooter ("")
				Else
					strSQL = "ALTER TABLE " & OpenIdentifier & Replace(Replace(strAlterTableName,"]",""),"[","") & CloseIdentifier & " DROP CONSTRAINT " & OpenIdentifier & Request("relationname") & CloseIdentifier & ""
				End If
			Case "deleteindex"
				strSQL =		"DROP INDEX " & OpenIdentifier &  Request("indexname") & CloseIdentifier & " ON " & OpenIdentifier & Replace(Replace(strAlterTableName,"]",""),"[","") & CloseIdentifier & ""
			Case "indexfield"
				strSQL = "CREATE INDEX " & OpenIdentifier & "ind" & strFieldName & CloseIdentifier & " ON " & OpenIdentifier & Replace(Replace(strAlterTableName,"]",""),"[","") & CloseIdentifier & " (" & OpenIdentifier & strFieldName & CloseIdentifier & ")"
			Case "droptable"
				If Request.Form("cancel") <> "" Then
					Response.Redirect ("Structure.asp?sqlfrom=" & server.URLEncode(strAlterTableName ))
				ElseIf Request.Form("confirm") = "" Then
					Call WriteHeader ("")
					Response.Write("<FORM action=""")
					Response.Write(ocdPageName & "?" & Request.QueryString)
					Response.Write(""" method=post><table><tr><td valign=top><img src=appWarning.gif alt=Warning></td><td valign=top><b>Are you sure you want to permanently drop the table ")
					Response.Write(Server.HTMLEncode(strAlterTableName))
					Response.Write("from this database?  <p></b>This action will remove the structure and all records in the table cannot be undone.</td></tr></table><p><input TYPE=Submit NAME=Confirm  CLASS=submit Value=""OK"">&nbsp;<input TYPE=Submit  CLASS=submit Name=Cancel Value=Cancel></FORM>")
					Call WriteFooter ("")
				Else
					Call notforsample()
					strAlterTableName = Replace(Replace(strAlterTableName,"]",""),"[","")
					strSQL = "DROP TABLE " & OpenIdentifier & strAlterTableName & CloseIdentifier & ""
					connDB.execute strSQL
					If Err.Number <> 0 Then
						Call WriteHeader("")
						Response.Write strSQL
						Call WriteFooter("")
					Else
						Response.Redirect ("Schema.asp?show=tables")
					End If
				End If
			Case "dropproc"
				If Request.Form("cancel") <> "" Then
					Response.Redirect ("Schema.asp?show=procs")
				ElseIf Request.Form("confirm") = "" Then
					Call WriteHeader ("")
					Response.Write("<FORM action=""")
					Response.Write(ocdPageName & "?" & Request.QueryString)
					Response.Write(""" method=post><table><tr><td valign=top><img src=appWarning.gif alt=Warning></td><td valign=top><b>Are you sure you want to permanently drop the Procedure ")
					Response.Write(Server.HTMLEncode(Request("SQLFrom")))
					Response.Write(" from this database?  <p></b>This action cannot be undone.</td></tr></table><p><input TYPE=Submit NAME=Confirm  CLASS=submit Value=""OK"">&nbsp;<input TYPE=Submit  CLASS=submit Name=Cancel Value=Cancel></FORM>")
					Call WriteFooter ("")
				Else
					strViewName = (Request("sqlfrom"))
					Set catADOX = Server.CreateObject("adox.catalog")
					Set catADOX.ActiveConnection = connDB
					If err <> 0 then
						Response.Write "This function requires ADOX to be installed on your web server."
						err.clear
					End If
					catADOX.procedures.Delete strViewName '.Delete
					Response.Redirect ("Schema.asp?show=procs")
				End If
			Case "dropview"
				If Request.Form("cancel") <> "" Then
				Response.clear
				'	Response.Redirect ("Schema.asp?show=views")
					Response.Redirect ("Structure.asp?sqlfrom=" & Server.URLEncode(FormatForSQL(strAlterTableName, "Access", "RemoveSQLIdentifier")))
					Response.End
				ElseIf Request.Form("confirm") = "" Then
					Call WriteHeader ("")
					Response.Write("<form action=""")
					Response.Write(ocdPageName & "?" & Request.QueryString)
					Response.Write(""" method=post><table><tr><td valign=top><img src=appWarning.gif alt=Warning></td><td valign=top><b>Are you sure you want to permanently drop the view ")
					Response.Write(Server.HTMLEncode(strAlterTableName))
					Response.Write(" from this database?  <p></b>This action will remove the structure and all records in the table cannot be undone.</td></tr></table><p><input TYPE=Submit NAME=Confirm  CLASS=submit Value=""OK"">&nbsp;<input TYPE=Submit  CLASS=submit Name=Cancel Value=Cancel></FORM>")
					Call WriteFooter ("")
				Else
'					strSQL = FormatForSQL(Request.QueryString("table"), "Access", "RemoveSQLIdentifier")
					strSQL ="DROP VIEW " & OpenIdentifier & FormatForSQL(strAlterTableName, "Access", "RemoveSQLIdentifier") &  CloseIdentifier & ""
'					response.end
					Call notforsample()
					connDB.execute strSQL
					if err.number = 0 Then
					
					Response.Redirect ("Schema.asp?show=views")
					Else
						Call WriteHeader("")
						response.write strSQL
						Call writefooter("")
						End if
				End If
			Case "copytable"
				If Request("cancel") <> "" Then
					Response.Redirect ("Structure.asp?sqlfrom=" & server.URLEncode(strAlterTableName ))
				ElseIf Request.Form("confirm") = "" or Request.Form("cpname") = "" Then
					Call WriteHeader ("")
					Response.Write("<SPAN CLASS=Information>Copy Table ")
					Response.Write(Server.HTMLEncode(strAlterTableName))
					Response.Write(" </span><FORM action=""")
					Response.Write(ocdPageName & "?action=copytable&amp;table=" & Server.URLEncode(Request.QueryString("table")))
					Response.Write(""" method=post><B>Copy </b><SELECT Name=cptype><OPTION VALUE=""Stucture"">Structure Only</OPTION><OPTION VALUE=""Data"" SELECTED>Structure and Data</OPTION></select><b> to new table </b><input TYPE=Table NAME=cpname> <BR><b>in database</b> <input Name=whatdb VALUE=""")
					Response.Write(connDB.Properties("Data Source Name"))
					Response.Write(""" SIZE=40 MAXLENGTH=255><p><input TYPE=Submit NAME=Confirm  CLASS=submit Value=""OK"">&nbsp;<input TYPE=Submit  CLASS=submit Name=Cancel Value=Cancel></FORM>")
					Call WriteFooter ("")
				Else
					If Request("cptype") = "Data" Then
					strAlterTableName = FormatForSQL(strAlterTableName, "Access", "REMOVESQLIDENTIFIER")
						If Request("whatdb") = "" Then
							Call notforsample()
							connDB.execute "SELECT * INTO " & OpenIdentifier & Request("cpname") & CloseIdentifier & " FROM " & OpenIdentifier & strAlterTableName & CloseIdentifier & ""
						Else
							Call notforsample()
							response.write "SELECT * INTO " & OpenIdentifier & Request("cpname") & CloseIdentifier & " IN '" & Request("whatdb") & "' FROM " & OpenIdentifier & strAlterTableName & CloseIdentifier & ""
							connDB.execute "SELECT * INTO " & OpenIdentifier & Request("cpname") & CloseIdentifier & " IN '" & Request("whatdb") & "' FROM " & OpenIdentifier & strAlterTableName & CloseIdentifier & ""
							'Response.Write("XX")
							'response.write err.number
							'response.write err.description
							'Response.End()
							End If
						Else
							If Request("whatdb") = "" Then
								Call notforsample()
								connDB.execute "SELECT * INTO " & OpenIdentifier & Request("cpname") & CloseIdentifier & " FROM " & OpenIdentifier & strAlterTableName & CloseIdentifier & " WHERE 1=2"		
							Else
								Response.Write "SELECT * INTO " & OpenIdentifier & Request("cpname") & CloseIdentifier & " IN '" & Request("whatdb") & "' FROM " & OpenIdentifier & strAlterTableName & CloseIdentifier & " WHERE 1=2"		
								connDB.execute "SELECT * INTO " & OpenIdentifier & Request("cpname") & CloseIdentifier & " IN '" & Request("whatdb") & "' FROM " & strAlterTableName &  " WHERE 1=2"		
							End If
						End If
						strAlterTableName = ""
						Response.Redirect ("Browse.asp?sqlfrom_a=" & Server.URLEncode(Request("cpname")))
					End If
			Case "deletetable"
				If Request.Form("cancel") <> "" Then
					Response.Redirect ("Structure.asp?sqlfrom=" & server.URLEncode(strAlterTableName ))
				ElseIf Request.Form("confirm") = "" Then
					Call WriteHeader ("")
					Response.Write("<FORM action=""")
					Response.Write(ocdPageName & "?" & Request.QueryString)
					Response.Write(""" method=post><table><tr><td valign=top><img src=appWarning.gif alt=Warning></td><td valign=top><b>Are you sure you want to permanently remove all records from the table ")
					Response.Write(Server.HTMLEncode(strAlterTableName))
					Response.Write("?  <p></b>This action will not affect the structure of the table but cannot be undone.</td></tr></table><p><input TYPE=Submit  CLASS=submit NAME=Confirm Value=""OK"">&nbsp;<input TYPE=Submit  CLASS=submit Name=Cancel Value=Cancel></FORM>")
					Call WriteFooter ("")
				Else
					Call notforsample
					connDB.execute "DELETE FROM " & OpenIdentifier & strAlterTableName & CloseIdentifier & ""
					Response.Redirect ("Structure.asp?sqlfrom=" & server.URLEncode(strAlterTableName))
				End If
		End select
		Call notforsample
		If strSQL="" THen
			Call WriteFooter("")
		End If
		Set rsTemp = connDB.Execute (strSQL)
		If Err <> 0 then
			Call WriteHeader("")
			Response.Write strSQL
			Call WriteFooter("")
		End If
		Response.Redirect ("Structure.asp?sqlfrom=" & server.URLEncode(strAlterTableName ))
End Select
If Request("ocdAction") <> "" Then
	Call WriteFooter("")
End If
Response.End
Sub notforsample()
	If ocdIsHome and not blnDemo then
		Response.Clear
		Call WriteHeader("")
		Call WriteFooter("Disabled for Example Database.  Upload your own Access .MDB file to demonstrate this feature online.")
		Response.End
		End If
End Sub
%>
