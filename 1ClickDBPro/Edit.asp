<%@ LANGUAGE = VBScript.Encode %>
<% ' Except for @ Directives, there should be no ASP or HTML codes above this line
' Setting LANGUAGE = VBScript.Encode enables mixing of encoded And unencoded ASP scripts

'1 Click DB Pro copyright 1997-2004 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com/

    
'**Start Encode**
%>
<!--#INCLUDE FILE="PageInit.asp"-->
<!--#INCLUDE FILE="ocdForm.asp"-->
<!--#INCLUDE FILE="ocdGrid.asp"-->
<!--#INCLUDE FILE="ocdFunctions.asp"-->
<%
'--------------------
'Begin Page_Load
'--------------------

Dim objForm

'--------------------
'End Page_Load
'--------------------

'--------------------
'Begin Page_Render
'--------------------

Call WriteHeader("")

Call DisplayEdit()

Call WriteFooter("")

'--------------------
'End Page_Render 
'--------------------

'--------------------
'Begin Procedures
'--------------------
Sub DisplayEdit()
	Dim rsDef, evDef, evDefresult, hasdef, fkrelatedfield, fkrelatedtable, fldF, strName, intSize, tQS, bintSize, rsFK, HasFK, intFKColumnCount, strFKColumnName, strPKTables, elePKName, eleFKName, strFKTables, arrPKTables, arrFKTables, prevPKTable, prevFKTable, cat, tblCat, astrTemp, keytblcat, colkeytblcat, rsdefeval, arrrsdef, intrsdef, rtmpqs, intI, arrHomeField, arrfkRelatedField, prevcolumn, objGrid, blnHANF, homefield, showrelated, intcountgrids, strSQLTName, varFormNum, arrCEP, eleCEP, blnCEP, blnBrowseAfterSave, blnBrowseAfterCancel
	Set objForm = New ocdForm
	varFormNum = 0
	objForm.DatabaseType = ocdDatabaseType
	objForm.MaxRelatedValues = ocdMaxRelatedValues
	objForm.FormNullToken = ocdFormNullToken
	'objForm.ADOConnection = ocdTargetConn
	objForm.SQLConnect = ndnscSQLConnect 'ADO Connect String, including uid And pw If necessary
	objForm.Debug = ocdDebug
	objForm.SQLUser = ndnscSQLUser
	objForm.SQLPass = ndnscSQLPass
	objForm.AllowMultiDelete = True
	objForm.SQLSelect = "*" 'Database Field List 
	objForm.SQLFrom = Request.QueryString("sqlFrom")'Database Table Name
	If ocdReadOnly Then
		objForm.AllowEdit = False
		objForm.AllowAdd = False
		objForm.AllowDelete = False
	Else
		objForm.AllowEdit = True
		objForm.AllowAdd = True
		objForm.AllowDelete = True
	End If
	objForm.HTMLCheckField = "<span class=Warning> Check  </span>"
	objForm.HTMLAttribSaveBtn = "TYPE=""Submit"" Value=""Save"" class=""Submit"""
	objForm.HTMLAttribCancelBtn = "TYPE=""Submit"" Value=""Cancel"" class=""Submit"""
	objForm.HTMLAttribNewBtn = "TYPE=""Submit"" Value=""New"" class=""Submit"""
	objForm.HTMLAttribDeleteBtn = "TYPE=""Submit"" Value=""Delete"" class=""Submit"""
	objForm.CallOnCancel = True
	blnBrowseAfterSave = ocdBrowseAfterSave
	blnBrowseAfterCancel = ocdBrowseAfterCancel

	Call objForm.Open()


	If (ocdDataBaseType =  "SQLServer") And Request.QueryString("ocdShowRelated") <> "Yes" Then
		ocdSelectForeignKey = false
		ocdShowRelatedRecords = false
	End If
	hasFK = False
	Select Case ocdDatabaseType
		Case "Access"
			strSQLTName = CStr(FormatForSQL(objForm.SQLFrom, ocddatabasetype, "RemoveSQLIdentifier"))
		Case "SQLServer"
			strSQLTName = GetSQLIDFPart(objForm.SQLFrom,"SQLOBJECTNAME", ocdQuotePrefix,ocdQuoteSuffix)
	End Select
	If (ocdSelectForeignKey Or ocdShowRelatedRecords ) And (objForm.ADOConnection.provider ="Microsoft.Jet.OLEDB.4.0" Or objForm.ADOConnection.provider ="SQLOLEDB.1") Then
		Set rsFK = objForm.ADOConnection.OpenSchema(27)
		If Err.Number = 0 Then
			If Not rsFK.eof Then
				HasFK = True
				strPKTables = ""
				strFKTables = ""
				prevPKTable = ""
				prevFKTable = ""
				Do While Not rsFK.eof
					If (rsFK.Fields("PK_TABLE_NAME").Value) = strSQLTName And (rsFK.Fields("FK_TABLE_NAME").Value) <> (prevFKTABLE) Then
						prevFKTable = (rsFK.Fields("FK_TABLE_NAME").Value)
						strFKTables = strFKTables & (rsFK.Fields("FK_TABLE_NAME").Value) & ","
					End If
					If (rsFK.Fields("FK_TABLE_NAME").Value) = strSQLTName And (rsFK.Fields("FK_NAME").Value) <> (prevPKTABLE) Then
						prevPKTable = (rsFK.Fields("FK_NAME").Value)
						strPKTables = strPKTables & (rsFK.Fields("FK_NAME").Value) & ","
					End If
					rsFK.movenext
				Loop
				If Len(strPKTables) > 0 Then
					strPKTables = Left(strPKTables, Len(strPKTables)-1)
				End If
				If Len(strFKTables) > 0 Then
					strFKTables = Left(strFKTables, Len(strFKTables)-1)
				End If
			Else
				rsFK.close
				Set rsFK = Nothing
			End If
		Else
			Set rsFK = Nothing
			Err.Clear
		End If
	End If
	If Request.QueryString("SQLFrom") = "" Then
		Response.Clear
		Response.Redirect ("Schema.asp"  )
	End If
	
	Response.Flush
	Response.Write("<span class=""Information""> ")
	If Request.QueryString("sqlid") = "" And Request.QueryString("SQLWHERE") = "" Then
		Response.Write("Add Record To ")
	Else
		Response.Write("Edit Record In ")
	End If
	Response.Write(" <a href=""Browse.asp?sqlfrom_a=" & Server.URLEncode(Request.QueryString("sqlfrom")) & "&amp;")
	For Each tQS In Request.QueryString
		If UCase(tQS) <> "SQLID" And UCase(tQS) <> "SQLFROM" And UCase(tQS) <> "NDBTNDELETE" And UCase(tQS) <> "SQLFROM_A" And UCase(tQS) <> "ACTION" And UCase(tQS) <> "SQLWHERE" Then
			Response.Write(tQS  & "=" &  Server.URLEncode(Request.QueryString(tQS)) & "&amp;")
		End If
	Next
	Response.Write(""">")
	Response.Write("" & Server.HTMLEncode(objForm.SQLFrom) )
	Response.Write("</a>")
	Response.Write("</span>")
	Select Case UCase(ocdDatabaseType)
		Case "SQLSERVER"
			Response.Write(" <a class=""menu"" href=""" & ocdPageName & "?")
			If Request.QueryString("OCDSHOWRELATED") = "Yes" Then
				Response.Write("ocdShowRelated=&amp;")
			Else
				Response.Write("ocdShowRelated=Yes&amp;")
			End If
			For Each tQS in Request.QueryString
				If UCase(tQS) <> "OCDSHOWRELATED" THen
					Response.Write(tQS  & "=" &  Server.URLEncode(Request.QueryString(tQS)) & "&amp;")
				End If
			next
			If Request.QueryString("OCDSHOWRELATED") = "Yes" Then
				Response.Write(""">(Hide ")
			Else
				Response.Write(""">(Show ")
			End If
			Response.Write("Related)</a>")
	End Select
	Response.Write("<p>")
	Response.Flush
	'start writing main body
	objForm.Display("STATUS")
	objForm.Display("START")
	Response.Write("<table>")
	If ocdShowDefaults And Request.QueryString("Sqlid") = "" And Request.QueryString("sqlwhere") = "" Then
		If (objForm.ADOConnection.provider ="Microsoft.Jet.OLEDB.4.0") Then
			Set rsdef = objForm.ADOConnection.OpenSchema(4,Array(Empty,Empty,CSTR(FormatForSQL(objForm.SQLFrom, ocddatabasetype, "RemoveSQLIdentifier")))) 'columns
			If rsdef.eof Then
				ocdShowDefaults = False
			Else
				arrrsdef = rsdef.getrows (,,Array("TABLE_NAME","COLUMN_NAME","COLUMN_DEFAULT"))
				rsdef.close
				Set rsdef = nothing
			End If
			Set rsdefeval = server.createobject("ADODB.Recordset")
		ElseIf (objForm.ADOConnection.provider = "SQLOLEDB.1")Then
			Set rsdef = objForm.ADOConnection.OpenSchema(4,Array(Empty,Empty,CSTR(GetSQLIDFPart(objForm.SQLFrom,"SQLOBJECTNAME", ocdQuotePrefix,ocdQuoteSuffix)))) 'columns
			If rsdef.eof Then
				ocdShowDefaults = False
			Else
				arrrsdef = rsdef.getrows (,,Array("TABLE_NAME","COLUMN_NAME","COLUMN_DEFAULT"))
				rsdef.close
				Set rsdef = Nothing
			End If
			Set rsdefeval = Server.CreateObject("ADODB.Recordset")
		Else
			ocdShowDefaults = False
		End If					
	Else
		ocdShowDefaults = False
	End If				
	' format each field according to its type
	blnHANF = False
	For Each fldF In objForm.ADORecordset.Fields
		strName = fldF.Name
		intSize = fldF.DefinedSize
		If intSize = -1 Then
			intSize=50
		End If
		intFKColumnCount = 0
		strFKColumnName = ""
		fkrelatedtable = ""
		fkrelatedfield = ""
		blnHANF = False
		If UCase(objForm.ADOConnection.provider)  = "MICROSOFT.JET.OLEDB.3.51" or objForm.DatabaseType = "IXS" Then
			'look out, field properties bomb with this provider
		Else
			If UCase(strName) = UCase(objForm.SQLSelectID) And ocdHideAutonumber Then
				blnHANF = True
			End If
		End If
		If Not blnHANF Then
			Select Case fldF.Type
				Case 205, 128, 204 'adLongVarBinary, adBinary, adVarBinary
					Response.Write("<tr><td nowrap valign='top' align=right>")
					Response.Write("<span class=""FieldName"">" & strName & ":</span>")
					Response.Write(" &nbsp;&nbsp;")
					Response.Write("</td>")
					Response.Write("<td align=""left"" valign=baseline>")
					Response.Write("<span class=Information>Binary&nbsp;Data</span> ")
					Response.Write("</td></tr>")
				Case Else
					hasdef=false
					If ocdShowDefaults And Request.QueryString("sqlid") = "" And Request.QueryString("sqlwhere") = "" And Not ocdDatabaseType = "Oracle" Then
						intrsdef = 0
						Do While intrsdef < ubound(arrrsdef,2)
							If ocdDataBaseType = "Access" Then				
				 				astrTemp =  FormatForSQL((Request.QueryString("sqlfrom")),ocddatabasetype,"REMOVESQLIDENTIFIER")
							ElseIf ocdDataBaseType = "SQLServer" THen
								astrTemp =  GetSQLIDFPart(Request.QueryString("SQLFROM"),"SQLOBJECTNAME", ocdQuotePrefix,ocdQuoteSuffix)				
							End If
							If astrTemp = (arrrsdef(0,intrsdef)) Then
								If UCase(strName) = UCase(arrrsdef(1,intrsdef)) Then
									If Not IsNull(arrrsdef(2,intrsdef)) Then
										evdef = arrrsdef(2,intrsdef)
										hasdef = True
										Exit Do
									End If
								End If
							End If
							intrsdef = intrsdef + 1
						Loop
						If Not hasdef Then
							evdefresult = "" 
						Else
							Call rsdefeval.Open ("Select " & evdef & " as expr1", objForm.ADOConnection)
							evdefresult = rsDefeval.Fields(0).Value
							rsdefeval.close
						End If
					Else
						evdefresult = ""
					End If
					If IsNull(evdefresult) then
						evdefresult = ""
					End If
					If ocdSelectForeignKey And HasFK And Not ocdReadOnly Then
						rsFK.MoveFirst
						Do While Not rsFK.EOF
							If (rsFK.Fields("FK_TABLE_NAME").Value) = strSQLTName And rsFK.Fields("FK_COLUMN_NAME").Value = strNAME Then
								intFKColumnCount = intFKColumnCount + 1
								strFKColumnName = strName
								fkrelatedtable = rsFK.Fields("PK_TABLE_NAME").Value
								fkrelatedfield = rsFK.Fields("PK_COLUMN_NAME").Value
							End If
							rsFK.movenext
						Loop
					End If
					Response.Write("<tr><td nowrap valign='top' align=right>")
					If (Not CBool(CInt(ndnscCompatibility) And ocdNoJavaScript)) And ocdDatabaseType = "Access" And ocdShowDescription Then
	Response.Write(vbCRLF & "<script TYPE=""text/javascript"" Language=""JavaScript"">" & vbCRLF)		
									Response.Write("document.write ('<img alt=\""Describe Field\"" src=\""appHelpSmall.gif\"" Border=0 onClick=\""javascript:window.open(\'DescribeField.asp?sqlfrom=" & Server.URLEncode(objForm.SQLFrom) & "&amp;SQLField=" & Server.URLEncode(strName) & "\', \'describe\',\'height=200,width=300,scrollbars=yes\')\"">');" & vbCRLF)
									Response.Write("</script>" & vbCRLF)
					End If
					Response.Write("<span class=""FIELDNAME"">" & strName & ":</span>")
					If CBool(fldF.Attributes And &H00000020) Then 'adFldIsNullable
						Response.Write(" &nbsp;&nbsp;")
					Else
						Response.Write(" <span class=""Warning"">*</span>")
					End If
					Response.Write("</td>")
					If intfkcolumncount = 1 Then 'multicolumns Not supported as dropdowns
						Response.Write("<td align=""left"" valign=""top"">")
						If objForm.ADOConnection.provider ="Microsoft.Jet.OLEDB.4.0" Then
							Call objForm.DisplayFieldAsRelatedValues(Replace(fldF.Name,"""","""""") ,"Select [" & fkRelatedField & "] From [" & fkRelatedTable & "] Order By [" & fkRelatedField & "]",evdefresult,"class=DataEntry")
						Else
							Call objForm.DisplayFieldAsRelatedValues(Replace(fldF.Name,"""","""""") ,"Select """ & fkRelatedField & """ From """ & fkRelatedTable & """ Order By """ & fkRelatedField & """",evdefresult,"class=DataEntry")
						End If
						Response.Write("</td></tr>")
					Else
						Select Case fldF.Type
							Case 201, 203 'adLongVarChar, adLongVarWChar
								Response.Write("<td align=""left"" valign=""top"">")
									Call objForm.DisplayFieldAsMemo(strName,evdefresult,"ROWS =""5"" COLS=""35"" class=""DataEntry"" ")
								If Not CBool(CInt(ndnscCompatibility) And ocdNoJavaScript) And Not ocdReadOnly Then
									Response.Write("&nbsp;<A HREF="""" onclick=""javascript:window.open('ocdZoomText.asp?CallingForm=" & varformnum & "&amp;TextField=" & Server.URLEncode("ocdTF" & strName) & "', 'zoomtext','height=400,width=600,scrollbars=yes');return false""><IMG ALT=""Zoom Text"" SRC=""GRIDLNKEDIT.GIF"" Border=0></A>")
									Response.Write(vbCRLF & "<script TYPE=""text/javascript"" Language=""JavaScript"">" & vbCRLF)		
									Response.Write("if (parseInt(navigator.appVersion) >= 4) {" & vbCRLF)
									Response.Write("	if (navigator.appName == ""Microsoft Internet Explorer"") {" & vbCRLF)
									Response.Write("document.write ('<IMG ALT=\""HTML Edit\"" SRC=\""AppHTMLEdit.gif\"" Border=0 onClick=\""javascript:window.open(\'ocdHTMLEdit.asp?CallingForm=" & varformnum & "&amp;TextField=" & server.urlencode("ocdTF" & strName) & "\', \'zoomtext\',\'height=400,width=600,scrollbars=yes\')\"">');" & vbCRLF)
									Response.Write("	}" & vbCRLF)
									Response.Write("}" & vbCRLF)
									Response.Write("</SCRIPT>" & vbCRLF)
								End If
								Response.Write("</td></tr>")
							Case 11 'adBoolean
								Response.Write("<td align=""left"" valign=""top"">")
								If Not CBool(fldF.Attributes And &H00000020) Then
									Call objForm.DisplayFieldAsCheckBox(strName,True,False,True,"")
								Else
									Call objForm.DisplayFieldAsTextBox(strName,"","SIZE=""5"" MAXLENGTH=""12"" class=""DataEntry""")
								End If
								Response.Write("</td></tr>")
							Case  133, 135, 134, 7 'adDBDate, adDBTimeStamp, adDBTime, adDate
								Response.Write("<td align=""left"" valign=""top"">")
									Call objForm.DisplayFieldAsTextBox(strName,evdefresult, "SIZE=""20"" MAXLENGTH=""50"" class=""DataEntry"" ")
								If Not cbool(cint(ndnscCompatibility) And ocdNoJavaScript) And Not ocdReadOnly Then

									Response.Write("<img onclick=""javascript:window.open('ocdPickDate.asp?CallingForm=" & varformnum & "&amp;DateField=" & server.urlencode("ocdTF" & strName) & "&amp;InitialDate=' + document.forms[" & varformnum & "].elements['" & ("ocdTF" & strName) & "'].value, 'calendar','height=260,width=250,scrollbars=yes');"" width=""17"" height=""17"" alt=""Click for Calendar"" SRC=AppCalendar.gif border=""0"">")
								End If
								Response.Write("</td></tr>")
							Case 6 'adCurrency
								Response.Write("<td align=""left"" valign=""top"">")
									Call objForm.DisplayFieldAsTextBox(strName,evdefresult, "SIZE=""12"" MAXLENGTH=""50"" class=""DataEntry"" ")
								Response.Write("</td></tr>")
							Case 20, 14, 5, 131, 4, 2, 16, 21, 19, 18, 17, 3 'adBigInt, adDecimal, adDouble, adNumeric, adSingle, _
							' adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, _
							' adUnsignedSmallInt, adUnsignedTinyInt,adInteger
								Response.Write("<td align=""left"" valign=""top"">")
									Call objForm.DisplayFieldAsTextBox(strName,evdefResult, "size=""24"" maxlength=""50"" class=""DataEntry"" ")
								Response.Write("</td></tr>")
							Case Else					
								Response.Write("<td align=""left"" valign=""top"">")
								If intSize > 35 then
									bintSize = 35
								Else
									bintSize = intSize
								End If
								Call objForm.DisplayFieldAsTextBox(strName,evdefresult, "size=""" & bintSize & """ maxlength=""" & intSize & """ class=""DataEntry"" ")
								Response.Write("</td></tr>")
						End Select
					End If
			End Select
		End If
		Response.Flush
		Response.Clear
	Next
	If HasFK Then
		rsFK.close
		Set rsFK = Nothing
	End If
	If ocdShowDefaults Then
		Set rsdefeval = Nothing
	End If
	Response.Write("</table><p>")
	If Not ocdReadOnly Then
		objForm.Display("BUTTONS")
		If objForm.EditMode <> "Add" And objForm.FormStatus= "" Then
		If (Not CBool(CInt(ndnscCompatibility) And ocdNoJavaScript)) And objForm.AllowAdd Then
		Response.Write("<input type=""submit"" value=""Duplicate"" onclick=""javascript:document.forms[0].action='")

					Response.Write(request.servervariables("SCRIPT_NAME") & "?sqlid=")
					dim qs
					For Each QS In Request.QueryString
						If UCASE(QS) <> "SQLID" And UCASE(QS) <> "SQLWHERE" Then
							Response.Write("&amp;" & QS & "=" & Server.URLEncode(Request.QueryString(QS)))
						End If
					Next
					Response.Write("';"" class=""submit"">")
					End If
	End If
	End if
	objForm.DISPLAY("END")	' And finally return the table
	Response.Write("<P><span class=""Warning"">*</span>indicates required field")
	Response.Write("<P>")
	Response.Flush
	If Err.Number <> 0 Then
		Call writefooter("")
	End If
	If ocdShowRelatedRecords And Request.Form("ocdEditDelete") = "" Then
		Select Case ocdDatabaseType
			Case "Access","SQLServer"
				
				If ((objForm.ADOConnection.provider ="Microsoft.Jet.OLEDB.4.0" OR objForm.ADOConnection.provider ="SQLOLEDB.1") And (Request.QueryString("SQLID") <> "" or Request.QueryString("SQLWHERE") <> "") And Not objForm.ADORecordset.eof) Then ' OR objForm.ADOConnection.provider ="SQLOLEDB.1"
				Set cat = Server.CreateObject("ADOX.Catalog")
				If Err.Number <> 0 Then
					Response.Write("Detailed view of related records is Not available, an ADOX catalog could Not be created.")
					Call writefooter("")
				Else
					intcountgrids = 1
					cat.ActiveConnection = objForm.ADOConnection
					If strPKTables <> "" Then '*
						homefield = ""
						fkrelatedfield = ""
						arrPKTables = split(strPKTables,",") '*
						Set tblcat = cat.Tables(strSQLTName)
						For Each elePKName In arrPKTables
							set keytblcat = tblcat.Keys(elePKName)
							If keytblcat.type = 2 Then
								If fkrelatedfield = "" then					
									For Each colkeytblcat in keytblcat.Columns
										If keytblcat.RelatedTable <> "" THen 
											prevcolumn = ""
											If homefield = "" then
												homefield = colkeytblcat.Name
												fkrelatedfield =colkeytblcat.RelatedColumn
											ElseIf prevcolumn <> homefield then
												homefield = homefield & "," &  colkeytblcat.Name
												fkrelatedfield = fkrelatedfield & "," &colkeytblcat.RelatedColumn
											End If
											prevcolumn = colkeytblcat.Name
											fkrelatedtable = FormatForSQL(keytblcat.RelatedTable,ocdDatabaseType,"ADDSQLIDENTIFIER")
										End If
									Next
								End If
								If fkRelatedTable <> "" Then
									Set objGrid = New ocdGrid
									objGrid.HTMLGridButtons = "First|First;;prev|Prev;;next|Next;;last|Last;;New|New"
									objGrid.HTMLSortASCLink= ""	'HTML to display inside sort ascending link
									objGrid.HTMLSortDESCLink= ""	'HTML to display inside sort descending link
									objGrid.HTMLFilterLink= ""

									Response.Write("<span class=Information>Related Record in <A HREF=""Browse.asp?sqlfrom_a=" & server.urlencode(ocdQuotePrefix & 	FormatForSQL(fkRelatedTable, ocdDatabaseType,"REMOVESQLIDENTIFIER") & ocdQuoteSuffix) & "&amp;")
									For Each tQS In Request.QueryString
										If UCase(tQS) <> "SQLID" And UCase(tQS) <> "SQLFROM" And UCase(tQS) <> "NDBTNDELETE" And UCase(tQS) <> "SQLFROM_A" And UCase(tQS) <> "SQLORDERBY_A" And UCase(tQS) <> "SQLWHERE_A" And UCase(tQS) <> "SQLGROUPBY_A" And UCase(tQS) <> "SQLHAVING_A" And UCase(tQS) <> "ACTION" And UCase(tQS) <> "SQLWHERE" Then
											Response.Write(tQS  & "=" &  Server.URLEncode(Request.QueryString(tQS)) & "&amp;")
										End If
									Next
									Response.Write("""> " & fkRelatedTable & "</a></span><P>")
									objGrid.SQLConnect = ndnscSQLConnect
									objGrid.SQLUser = ndnscSQLUser
									objGrid.SQLPass = ndnscSQLPass
									objGrid.GridID = "Default" & intcountgrids
									objGrid.SQLSelect = "*"
									objGrid.SQLFrom = fkRelatedTable
									objGrid.GridHideAutonumber = ocdHideAutonumber
									If InStr(homefield,",") = 0 Then
										Select Case objForm.ADORecordset.Fields(homefield).Type
											Case 20, 14, 5, 131, 4, 2, 16, 21, 19, 18, 17, 3 
									'adBigInt, adDecimal, adDouble, adNumeric, adSingle, _
									' adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, 
									' adUnsignedSmallInt, adUnsignedTinyInt,adInteger
												If IsNull(objForm.ADORecordset.Fields(homefield).Value) Then
													objGrid.SQLWhereExtra = "[" & fkrelatedfield & "] Is Null"
												Else
													objGrid.SQLWhereExtra = "[" & fkrelatedfield & "] =" & objForm.ADORecordset.Fields(homefield).Value
												End If
											Case Else
												If IsNull(objForm.ADORecordset.Fields(homefield).Value) Then
													objGrid.SQLWhereExtra = "[" & fkrelatedfield & "] Is Null"
												Else
													objGrid.SQLWhereExtra = "[" & fkrelatedfield & "] ='"  & Replace(objForm.ADORecordset.Fields(homefield).Value,"'","''") & "'"
												End If
										End Select
									Else
										objGrid.SQLWhereExtra = ""
										arrhomefield = Split(homefield,",")
										arrfkrelatedfield = Split(fkrelatedfield,",")
										For intI = 0 To UBound (arrhomefield)
											Select Case objForm.ADORecordset.Fields(arrhomefield(intI)).Type
												Case 20, 14, 5, 131, 4, 2, 16, 21, 19, 18, 17, 3 
										'adBigInt, adDecimal, adDouble, adNumeric, adSingle, _
										' adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, 
										' adUnsignedSmallInt, adUnsignedTinyInt,adInteger
													If IsNull(objForm.ADORecordset.Fields(arrhomefield(intI)).Value) Then
														objGrid.SQLWhereExtra = objGrid.SQLWhereExtra & "[" & arrfkrelatedfield(intI) & "] Is Null And "
													Else
														objGrid.SQLWhereExtra = objGrid.SQLWhereExtra & "[" & arrfkrelatedfield(intI) & "] =" & objForm.ADORecordset.Fields(arrhomefield(intI)).Value & " And "
													End If
												Case Else
													If IsNull(objForm.ADORecordset.Fields(arrhomefield(intI)).Value) Then
														objGrid.SQLWhereExtra = objGrid.SQLWhereExtra & "[" & arrfkrelatedfield(intI) & "] Is Null And "
													Else
														objGrid.SQLWhereExtra = objGrid.SQLWhereExtra & "[" & arrfkrelatedfield(intI) & "] ='"  & Replace(objForm.ADORecordset.Fields(arrhomefield(intI)).Value,"'","''") & "'" & " And "
													End If
											End Select
										Next
										objGrid.SQLWhereExtra = left(objGrid.SQLWhereExtra,len(objGrid.SQLWhereExtra)-5)
									End If
									objGrid.AllowEdit = True
									objGrid.AllowDelete = False
									objGrid.AllowAdd = False
									objGrid.AllowExport = False
									blnCEP = False
									Select Case ocdCustomEditPages
										Case "*"
											blnCEP = True
										Case ""
											blnCEP = False
										Case Else
											arrCEP = Split(ocdCustomEditPages,",")
											For Each eleCEP In arrCEP
	'										Response.Write objGrid.SQLFrom
												If eleCEP = objGrid.SQLFrom Then
													blnCEP = True
													Exit For
												End If
											Next
									End Select
									If blnCEP Then
										objGrid.FormEdit = Replace(Replace(Replace(Replace(Replace(Replace(objGrid.SQLFrom," ","_"),"/","_"),"""",""),".","_"),"[",""),"]","") & "_Edit.asp"
									Else
										objGrid.FormEdit = "Edit.asp"
									End If
									objGrid.SQLSelectIDName = ""
									objGrid.SQLSelectPK = ""
									objGrid.Open
									objGrid.Display("GRID") 
									Response.Write("<P>")
									Response.flush
									intcountgrids = intcountgrids + 1
									homefield = ""

									fkrelatedfield = ""
									set objGrid = nothing
								End If
							Else 'not fkey
								homefield = ""
								fkrelatedfield = ""
							End If
							homefield = ""
							fkrelatedfield = ""
						Next
					End If
					If strFKTables <> "" Then '*
						homefield = ""
						fkrelatedfield = ""
						arrFKTables = split(strFKTables,",") '*
						For Each eleFKName In arrFKTables '*
							Set tblcat = cat.Tables(eleFKNAME)
							For Each keytblcat in tblcat.Keys
								If keytblcat.type = 2 Then
									For Each colkeytblcat In keytblcat.Columns
										If (keytblcat.RelatedTable) = strSQLTName Then
											showrelated = true
											fkrelatedtable = keytblcat.RelatedTable
											If homefield = "" Then
												homefield = colkeytblcat.Name
												fkrelatedfield =colkeytblcat.relatedcolumn
											Else
												homefield = homefield & "," & colkeytblcat.Name
												fkrelatedfield =fkrelatedfield & "," & colkeytblcat.relatedcolumn
											End If
										Else 
											showrelated = False
										End If
									Next
									If showrelated Then
										Set objGrid = New ocdGrid
										objGrid.HTMLGridButtons = "First|First;;prev|Prev;;next|Next;;last|Last;;New|New"
										objGrid.HTMLSortASCLink= ""	'HTML to display inside sort ascending link
										objGrid.HTMLSortDESCLink= ""	'HTML to display inside sort descending link
										objGrid.HTMLFilterLink= ""
										objGrid.SQLConnect = ndnscSQLConnect
										objGrid.SQLUser = ndnscSQLUser
										objGrid.GridHideAutonumber = ocdHideAutonumber
										objGrid.SQLPass = ndnscSQLPass
										objGrid.GridID = "Default" & intcountgrids
										objGrid.SQLSelect = "*"
										objGrid.SQLFrom = FormatForSQL(CStr(tblCat.Name),ocdDatabaseType,"ADDSQLIDENTIFIER")
										If instr(homefield,",") = 0 Then
											If IsNull(objForm.ADORecordset.Fields(fkrelatedfield).Value) THen	
												If  (objForm.ADOConnection.provider ="Microsoft.Jet.OLEDB.4.0") Then
													objGrid.SQLWhereExtra = "[" & homefield & "] Is Null"
												Else
													objGrid.SQLWhereExtra = """" & homefield & """ Is Null" 
												End If
											Else
												Select Case objForm.ADORecordset.Fields(fkrelatedfield).Type
													Case 20, 14, 5, 131, 4, 2, 16, 21, 19, 18, 17, 3 
											'adBigInt, adDecimal, adDouble, adNumeric, adSingle, _
											' adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, 
											' adUnsignedSmallInt, adUnsignedTinyInt,adInteger
														If  (objForm.ADOConnection.provider ="Microsoft.Jet.OLEDB.4.0") Then
												objGrid.SQLWhereExtra = "[" & homefield & "] = " & objForm.ADORecordset.Fields(fkrelatedfield).Value
														Else
															objGrid.SQLWhereExtra = """" & homefield & """ = " & objForm.ADORecordset.Fields(fkrelatedfield).Value
														End If
													Case Else
														If  (objForm.ADOConnection.provider ="Microsoft.Jet.OLEDB.4.0") Then
												objGrid.SQLWhereExtra = "[" & homefield & "] ='"  & Replace(objForm.ADORecordset.Fields(fkrelatedfield).Value,"'","''") & "'"
														Else
															objGrid.SQLWhereExtra = """" & homefield & """ ='"  & Replace(objForm.ADORecordset.Fields(fkrelatedfield).Value,"'","''") & "'"
														End If
												End Select
											End If
										Else
											arrhomefield = split (homefield,",")
											arrfkrelatedfield = Split(fkrelatedfield,",")
											objGrid.SQLWhereExtra = ""
											For intI = 0 To UBound(arrhomefield)
												Select Case objForm.ADORecordset.Fields(arrfkrelatedfield(intI)).Type
													Case 20, 14, 5, 131, 4, 2, 16, 21, 19, 18, 17, 3 
											'adBigInt, adDecimal, adDouble, adNumeric, adSingle, _
											' adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, 
											' adUnsignedSmallInt, adUnsignedTinyInt,adInteger
														If  (objForm.ADOConnection.provider ="Microsoft.Jet.OLEDB.4.0") Then
																objGrid.SQLWhereExtra = objGrid.SQLWhereExtra & "[" & arrhomefield(intI) & "]"
																If IsNull(objForm.ADORecordset.Fields(arrfkrelatedfield(intI)).Value) Then
																	objGrid.SQLWhereExtra = objGrid.SQLWhereExtra & " Is Null And "
																Else
																	objGrid.SQLWhereExtra = objGrid.SQLWhereExtra & "=" & objForm.ADORecordset.Fields(arrfkrelatedfield(intI)).Value & " And "
																End If 
															Else
																If IsNull(objForm.ADORecordset.Fields(arrfkrelatedfield(intI)).Value) Then
													objGrid.SQLWhereExtra = objGrid.SQLWhereExtra & """" & arrhomefield(intI) & """ Is Null And "
																Else
																	objGrid.SQLWhereExtra = objGrid.SQLWhereExtra & """" & arrhomefield(intI) & """ =" & objForm.ADORecordset.Fields(arrfkrelatedfield(intI)).Value & " And "
																End If
															End If
														Case Else
															If  (objForm.ADOConnection.provider ="Microsoft.Jet.OLEDB.4.0") Then
																If IsNull(objForm.ADORecordset.Fields(arrfkrelatedfield(intI)).Value) Then
																	objGrid.SQLWhereExtra = objGrid.SQLWhereExtra & "[" & arrhomefield(intI) & "] Is Null And "
																Else
																	objGrid.SQLWhereExtra = objGrid.SQLWhereExtra & "[" & arrhomefield(intI) & "] ='"  & Replace(objForm.ADORecordset.Fields(arrfkrelatedfield(intI)).Value,"'","''") & "'" & " And "
																End If
															Else
																If IsNull(objForm.ADORecordset.Fields(arrfkrelatedfield(intI)).Value) Then
																	objGrid.SQLWhereExtra = objGrid.SQLWhereExtra & """" & arrhomefield(intI) & """ = Is Null And "
																Else
																	objGrid.SQLWhereExtra = objGrid.SQLWhereExtra & """" & arrhomefield(intI) & """ ='"  & Replace(objForm.ADORecordset.Fields(arrfkrelatedfield(intI)).Value,"'","''") & "'" & " And "
																End If
															End If
													End Select
												Next
												objGrid.SQLWhereExtra = Left(objGrid.SQLWhereExtra,Len(objGrid.SQLWhereExtra)-5)
											End If
											objGrid.SQLPageSize = ocdPageSizeDefault
											objGrid.SQLPage = ""
											objGrid.AllowEdit = True
											If Not ocdReadOnly Then
												objGrid.AllowDelete = True
											Else
												objGrid.AllowDelete = True
											End If
											objGrid.AllowAdd = True
											objGrid.SQLSelectIDName = ""
											objGrid.SQLSelectPK = ""
											Select Case ocdCustomEditPages
												Case "*"
													blnCEP = True
												Case ""
													blnCEP = False
												Case Else
													arrCEP = Split(ocdCustomEditPages,",")
													For Each eleCEP In arrCEP
														If eleCEP = objGrid.SQLFrom Then
															blnCEP = True
															Exit For
														End If
													Next
											End Select
											If blnCEP Then
												objGrid.FormEdit = Replace(Replace(Replace(Replace(Replace(Replace(objGrid.SQLFrom," ","_"),"/","_"),"""",""),".","_"),"[",""),"]","") & "_Edit.asp"
											Else
												objGrid.FormEdit = "Edit.asp"
											End If
											objGrid.Open
											Response.Write("<span class=""information"">")
											Response.Write(objGrid.SQLRecordCount)
											Response.Write(" Related Records in <a href=""Browse.asp?sqlfrom_a=" & server.urlencode(ocdQuotePrefix & tblcat.name & ocdQuoteSuffix) & "&amp;" )
											For Each tQS In Request.QueryString
												If UCase(tQS) <> "SQLID" And UCase(tQS) <> "SQLFROM" And UCase(tQS) <> "NDBTNDELETE" And UCase(tQS) <> "SQLFROM_A"  And UCase(tQS) <> "SQLORDERBY_A" And UCase(tQS) <> "SQLWHERE_A" And UCase(tQS) <> "SQLGROUPBY_A" And UCase(tQS) <> "SQLHAVING_A"  And UCase(tQS) <> "ACTION" And UCase(tQS) <> "SQLWHERE" Then
													Response.Write(tQS  & "=" &  Server.URLEncode(Request.QueryString(tQS)) & "&amp;")
												End If
											Next
											Response.Write( """>")
											Response.Write(tblcat.name & "</a></span><br>")
											objGrid.Display("BUTTONS") 
											objGrid.Display("GRID") 
											Response.Write("<P>")
											Response.flush
											intcountgrids = intcountgrids + 1
											homefield = ""
											fkrelatedfield = ""
											set objGrid = Nothing
										End If
									Else 'not fkey
										homefield = ""
										fkrelatedfield = ""
									End If
								Next
							Next
						End If
	'					Set objGrid = Nothing
						Set tblcat = Nothing
						Set keytblcat = Nothing
						Set cat = Nothing
					End If'adox catalog Not created
				End If
		End Select
	End If 'check If related records should be displayed
	objForm.Close()
End Sub
'React to pseudo-events

Sub ocdBeforeUpdate ()
	'not used	
End Sub

Sub ocdAfterUpdate()
	If blnBrowseAfterSave Then
			Call RedirectToBrowse()
	End If
end sub

Sub ocdBeforeInsert ()
	'not used	
End sub

Sub ocdAfterInsert()
	If blnBrowseAfterSave Then
			Call RedirectToBrowse()
	End If
End Sub

Sub ocdAfterDelete()
	Dim strADURL, tmpadqs
	If Request.QueryString("SQLFROM_A") <> "" Then
		strADURL = "Browse.asp?"
		For Each tmpadqs In Request.QueryString
			Select Case UCase(tmpadqs) 
				Case "SQLFROM","SQLSELECT","SQLWHERE","SQLID"
				Case Else
					strADURL = strADURL & tmpadqs & "=" & Server.URLEncode(Request.QueryString(tmpadqs)) & "&"
			End Select
		Next
		response.clear
		response.redirect strADURL
	End If
End Sub

Sub ocdOnCancel()
	If blnBrowseAfterCancel Then
			Call RedirectToBrowse()
	End If
End Sub


Sub ocdBeforeDelete()
	
	Dim tmpeqs
	Response.Write("<FORM ACTION=""")
	Response.Write(Request.ServerVariables("SCRIPT_NAME") & "?")
	For Each tmpeqs In Request.QueryString
		If UCase(tmpeqs) <> "OCDEDITDELETE" Then
			Response.Write(tmpeqs & "=" & Server.URLEncode(Request.QueryString(tmpeqs)) & "&")
		End If
	Next
	Response.Write(""" method=post>")
	Response.Write("<table align=""center"" WIDTH=""50%"" class=""DialogBox"" cellpadding=""5""><tr><th style=""text-align:left;background-color:navy;color:white;"" align=""left""><div style=""color:white;"">Confirm Delete</div></th><tr><td bgcolor=""Silver"" valign=""top"">")
	Response.Write("<P><img src=""appWarningSmall.gif"" border=""0"" alt=""Warning"">")
	If objForm.SQLID <> "" Or objForm.SQLWhere <> "" Then
		Response.Write("<strong>Are you sure you want to delete ")
		If instr(objForm.SQLID,",") > 0 Then
			Response.Write("these records ")
		Else
			Response.Write("this record ")
		End If
		Response.Write("?</Strong><p>This action cannot be undone.<P><INPUT type=""submit"" SPAN class=Submit Name=ocdEditConfirm Value=""OK"">&nbsp;<input type=""submit"" name=""ocdEditCancel"" class=Submit Value=""Cancel""><INPUT TYPE=hidden Name=ocdEditCancelPage class=Submit Value=""Browse.asp""><input type=""hidden"" name=""ocdEditDelete"" class=""Submit"" value=""Delete""></td></tr></table>")		
	Else
		Response.Write("<b>No Records Were Selected</b><P>You may use your browser's back button to continue.</td></tr></table></td></tr></TABLE></CENTER>")		
	End If
	Response.Write("</form>")
	Call WriteFooter("")
	Response.End
End Sub

Sub RedirectToBrowse()
	Dim strADURL, tmpadqs
	If Request.QueryString("SQLFROM_A") <> "" Then
		strADURL = "Browse.asp?"
		For Each tmpadqs in Request.QueryString
			Select Case UCASE(tmpadqs) 
				Case "SQLFROM","SQLSELECT","SQLWHERE", "SQLID"
				Case Else
					strADURL = strADURL & tmpadqs & "=" & Server.URLEncode(Request.QueryString(tmpadqs)) & "&"
			End Select
		Next
		Response.Clear
		Response.Redirect(strADURL)
	End If
End Sub

%>