<% ' Except for @ Directives, there should be no ASP or HTML codes above this line
' Setting LANGUAGE = VBScript.Encode enables mixing of encoded And unencoded ASP scripts

'1 Click DB Pro copyright 1997-2004 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com/

        
'**Start Encode**

%>
<!--#INCLUDE FILE=PageInit.asp-->
<!--#INCLUDE FILE=ocdDoubleBox.asp-->
<!--#INCLUDE FILE=ocdFunctions.asp-->
<%
Dim GRIDID, qsTemp, strRedirect, strTempField, fldTemp, rsTemp, objDoubleBox, blnHasTotals, arrSchemaHideObjects, eleSchemaHideObjects, blnShowObject
GRIDID = "A"
strRedirect = ""
If ocdSchemaHideObjects <> "" Then
	arrSchemaHideObjects = Split(ocdSchemaHideObjects,",")
End If
If Request("ShowAll1") <> "" Then
	strRedirect = "Browse.asp?sqlfrom_" & GRIDID & "=" & Server.URLEncode(Request.QueryString("sqlfrom_" & GRIDID))
ElseIf Request("totalit") <> "" or Request("setit") <> "" Then
	strRedirect = strRedirect & "Browse.asp?SQLSelect_" & GRIDID & "=" & Server.URLEncode(Request.QueryString("SQLSelect_" & GRIDID)) & "&SQLFrom_" & GRIDID & "=" & Server.URLEncode(Request.QueryString("sqlfrom_" & GRIDID))
	strRedirect = strRedirect &  "&SQLSelectSum_" & GRIDID & "=" & Server.URLEncode(Request.QueryString("sqlgrandtotals_" & GRIDID))
	strRedirect = strRedirect & "&SQLSelectMax_" & GRIDID & "=" & Server.URLEncode(Request.QueryString("sqlmaximums_" & GRIDID))
	strRedirect = strRedirect & "&SQLSelectMin_" & GRIDID & "=" & Server.URLEncode(Request.QueryString("sqlminimums_" & GRIDID))
	strRedirect = strRedirect & "&SQLSelectAvg_" & GRIDID & "=" & Server.URLEncode(Request.QueryString("sqlaverages_" & GRIDID))
	If Request("setit") <> "" Then
		strRedirect = strRedirect & "&ocdGridMode_A=Search"
	End If
ElseIf Request("selectfields") <> "" And Request.QueryString("sqlselect_" & GRIDID) = "" Then
	strRedirect = (Request.ServerVariables("SCRIPT_NAME") & "?sqlfrom_" & GRIDID & "=" & Server.URLEncode(Request.QueryString("sqlfrom_" & GRIDID)))
ElseIf Request("ShowALL") <> "" And Request("sqlselect_" & GRIDID) <> "" And Request("sqlselect_" & GRIDID) <> "" Then
	strRedirect = "Browse.asp?"
	strRedirect = strRedirect & "sqlfrom_" & GRIDID & "="
	strRedirect = strRedirect & Server.URLEncode(Request.QueryString("sqlfrom_" & GRIDID))
	strRedirect = strRedirect & "&sqlselect_" & GRIDID & "="
	strRedirect = strRedirect & Server.URLEncode(Request("sqlselect_" & GRIDID))
ElseIf Request("ShowDups") <> "" And Request("sqlselect_" & GRIDID) <> ""  Then
	strRedirect = "Browse.asp?"
	strRedirect = strRedirect & "sqlfrom_" & GRIDID & "="
	strRedirect = strRedirect & Server.URLEncode(Request.QueryString("sqlfrom_" & GRIDID))
	strRedirect = strRedirect & "&sqlselect_" & GRIDID & "="
	strRedirect = strRedirect & Server.URLEncode(Request("sqlselect_" & GRIDID) & ", Count(*) as Dups")
	strRedirect = strRedirect & "&sqlgroupby_" & GRIDID & "="
	strRedirect = strRedirect & Server.URLEncode(Request("sqlselect_" & GRIDID))
	strRedirect = strRedirect & "&sqlhaving_" & GRIDID & "="
	strRedirect = strRedirect & Server.URLEncode("Count(*)>1")
ElseIf Request("SelectFields") <> "" And Request("Sqlselect_" & GRIDID) <> "" Then
	strRedirect = "Browse.asp?ocdGridMode_" & GRIDID & "=Search&"
	strRedirect = strRedirect & "sqlfrom_" & GRIDID & "="
	strRedirect = strRedirect & Server.URLEncode(Request.QueryString("sqlfrom_" & GRIDID))
	strRedirect = strRedirect & "&sqlselect_" & GRIDID & "="
	strRedirect = strRedirect & Server.URLEncode(Request("Sqlselect_" & GRIDID))
End If

If strRedirect <> "" Then
	Response.Clear()
	Response.Redirect(strRedirect)
	Response.End
End If

Call WriteHeader("")

'Set ocdTargetConn = Server.CreateObject("ADODB.Connection")

'Call ocdTargetConn.open (ndnscSQLConnect , ndnscSQLUser , ndnscSQLPass)

If Err.Number <> 0 Then
	Call WriteFooter("")
End If

Set rsTemp = Server.CreateObject("ADODB.Recordset")

If Request.QueryString("sqlfrom_" & GRIDID) <> "" And Request("sqlselect_" & GRIDID) <> "" And Request.QueryString("showobj") = "" And Request("ComputeFields") <> "" Then
	'***************Select Totals
	Response.Write("<span class=Information>Set Totals for Selected Numeric Fields in ")
	Response.Write Server.HTMLEncode(Request.QueryString("sqlfrom_" & GRIDID))
	Response.Write("</span>")
	Response.Write("<FORM METHOD=get ACTION=""" & Request.ServerVariables("SCRIPT_NAME") &  """>")
	Call rsTemp.Open("SELECT " & Request.QueryString("sqlselect_" & GRIDID) & " from " & clns(Request.QueryString("sqlfrom_" & GRIDID)) &  " WHERE 1=2", ocdTargetConn)
	If strTempField <> "" Then
		strTempField = Left(strTempField,Len(strTempField)-1)
	End If
	strTempField = ""
	If strTempField <> "" Then
		strTempField = Left(strTempField,Len(strTempField)-1)
	End If
	For Each qsTemp In Request.QueryString
		Select Case UCASE(qsTemp)
			Case "sqlfrom_" & GRIDID
			Case Else
				Response.Write("<input type=Hidden name=""" & qsTemp & """ value=""" & Server.HTMLEncode(Request.QueryString(qsTemp)) & """>")
		End Select
	Next
	strTempField = ""
	blnHasTotals = False
	Response.Write("<p><table border=""1"">")
	Response.Write("<tr><th>Field</th><th>SUM</th><th>AVG</th><th>MIN</th><th>MAX</th></tr>")
	For Each fldTemp in rsTemp.Fields
		Select Case fldTemp.Type
									Case 2, 3, 4, 5, 14, 16, 17, 18, 19, 20, 128, 6	 'adSmallInt, adInteger, adSingle, adDouble, adDecimal, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adBigInt, adUnsignedBigInt, adNumeric,  adCurrency, 
Response.Write("<tr>")
Response.Write("<td>")
Response.Write(Server.HTMLEncode(Replace(clns(fldTemp.Name),"""","""")))
Response.Write("</td>")
Response.Write("<td>")
Response.Write("<input type=""Checkbox"" name=""sqlgrandtotals_" & GRIDID & """ value=""" & Server.HTMLEncode(Replace(clns(fldTemp.Name),"""",""""))& """>")
Response.Write("</td>")
Response.Write("<td>")
Response.Write("<input type=""Checkbox"" name=""sqlaverages_" & GRIDID & """ value=""" & Server.HTMLEncode(Replace(clns(fldTemp.Name),"""",""""))& """>")
Response.Write("</td>")
Response.Write("<td>")
Response.Write("<input type=""Checkbox"" name=""sqlminimums_" & GRIDID & """ value=""" & Server.HTMLEncode(Replace(clns(fldTemp.Name),"""",""""))& """>")
Response.Write("</td>")
Response.Write("<td>")
Response.Write("<input type=""Checkbox"" name=""sqlmaximums_" & GRIDID & """ value=""" & Server.HTMLEncode(Replace(clns(fldTemp.Name),"""",""""))& """>")
Response.Write("</td>")
Response.Write("</tr>")
		blnHasTotals = True
	Case Else
	End Select
	Next
	Response.Write "</table>"
	If Not blnHasTotals Then
		Response.write "<p>None of the Selected Fields are Numeric</p>"
	End If'Response.Write blnHasTotals
	Response.Write("<p><input type=""submit"" name=""SetIt"" value=""Set Criteria &gt;"" class=""submit""> ")
	Response.Write("<input type=""submit"" name=""TotalIt"" value=""Show All &gt;&gt;"" class=""submit""><p>")
	Response.Write "</form>"
	Call WriteFooter("")
	'*************** end totals 
ElseIf Request.QueryString("sqlfrom_" & GRIDID) <> "" And Request.QueryString("showobj") = "" Then '("SelectTables") <> "" Then
	'********************** Start Field Select
	Response.Write("<span class=Information>Select Fields From ")
	Response.Write(Server.HTMLEncode(Request.QueryString("sqlfrom_" & GRIDID)))
	Response.Write("</span>")
	Response.Write("<form method=""get"" action=""" & Request.ServerVariables("SCRIPT_NAME") &  """>")
	For Each qsTemp In Request.QueryString
		Select Case UCASE(qsTemp)
			Case "SQLSELECT_" & GRIDID
			Case Else
				Response.Write("<input type=Hidden name=""" & qsTemp & """ value=""" & Server.HTMLEncode(Request.QueryString(qsTemp)) & """>")
		End Select
	Next
	Call rsTemp.Open("SELECT * FROM " & clns(Request.QueryString("sqlfrom_" & GRIDID)) &  " WHERE 1=2", ocdTargetConn)
	strTempField = ""
	For Each fldTemp in rsTemp.Fields
		strTempField = strTempField & Replace(clns(fldTemp.Name),"""","""") & ","
	Next
	If strTempField <> "" Then
		strTempField = Left(strTempField,Len(strTempField)-1)
	End If
	Set objDoubleBox = New DoubleBox
	objDoubleBox.doublebox_Name = "sqlselect_" & GRIDID
	objDoubleBox.doublebox_fields = strTempField
	objDoubleBox.doublebox_size = 12
	objDoubleBox.doublebox_lHeader ="<span class=FieldName>&nbsp;</span>"
	objDoubleBox.doublebox_rHeader ="<span class=FieldName>Selected Fields</span>"
	objDoubleBox.DrawDoubleBox
	Response.Write("<p><input type=Submit name=""ComputeFields"" value=""Set Totals &gt;"" class=submit> ")
	Response.Write(" <input type=Submit name=""SelectFields"" value=""Set Criteria &gt;"" class=submit><p>")
	Response.Write("<p><input type=Submit name=""ShowDups"" value=""Show Duplicates &gt;&gt;"" class=submit>")
	Response.Write(" <input type=Submit name=""ShowAll"" value=""Show All &gt;&gt;"" class=submit><p>")
	'*************** end field select
Else
	'**************************start table/view select
	Response.Write("<span class=Information>Select From ")
	Select Case UCase(Request.QueryString("objtoshow"))
		Case "QUERIES"
			Response.Write "View"
		Case "BOTH"
			Response.Write "Table or View"
		Case Else
			Response.Write "Table"
	End Select
	Response.Write("</span> ")
	Response.Write "<form action=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString & """>"
	For Each qsTemp in Request.QueryString
		Select Case UCASE(qsTemp)
			Case "sqlfrom_" & GRIDID
			Case Else
				Response.Write("<input type=""hidden"" name=""" & qsTemp & """ value=""" & Server.HTMLEncode(Request.QueryString(qsTemp)) & """>")
		End Select
	Next
	Response.Write("<select name=""sqlfrom_" & GRIDID & """ size=""12"">")
	If ocdDatabaseType = "Oracle" Then
		Set rsTemp = ocdTargetConn.Execute ("SELECT OBJECT_TYPE AS TABLE_TYPE, OBJECT_NAME AS TABLE_NAME, OWNER AS TABLE_SCHEMA FROM ALL_OBJECTS WHERE (OBJECT_TYPE = 'TABLE' OR OBJECT_TYPE = 'VIEW') And NOT OWNER = 'SYS' And NOT OWNER = 'WKSYS' And NOT OWNER = 'MDSYS' And NOT OWNER = 'OLAPSYS' And NOT OWNER ='CTXSYS' And NOT OWNER='SYSTEM'")
	Else
		Set rsTemp = ocdTargetConn.OpenSchema(20) 'adSchemaTables
	End If
	Dim intShowObjCount
	intShowObjCount = 0
	Do While Not rsTemp.EOF
	blnShowObject = True
	response.write rsTemp("TABLE_TYPE")
		If ((rsTemp.Fields("TABLE_TYPE").Value = "TABLE" OR rsTemp.Fields("TABLE_TYPE").Value = "VIEW") And UCASE(Left(rsTemp.Fields("TABLE_NAME").Value,4)) <> "MSYS")  Then
			If ((Request.QueryString("objtoshow") = "Both" or Request.QueryString("objtoshow") = "" or Request.QueryString("objtoshow") = "Tables") And (rsTemp("TABLE_TYPE").VALUE = "TABLE")) or ((Request.QueryString("objtoshow") = "Queries" or Request.QueryString("objtoshow") = "Both") And rsTemp("TABLE_TYPE").VALUE = "VIEW") Then

			
			If ocdSchemaHideObjects <> "" Then

 					For Each eleSchemaHideObjects In arrSchemaHideObjects
						If eleSchemaHideObjects = FormatForSQL(rsTemp.Fields("TABLE_NAME").VALUE,ocdDatabaseType,"ADDSQLIDENTIFIER") Then
							blnShowObject = False
							Exit For
						End If
					Next
				End If
			Else
				blnShowObject = False
			End If
			If blnShowObject Then
				intShowObjCount = intShowObjCount + 1
				Response.Write("<option value=""" & Server.HTMLEncode(rsTemp.Fields("TABLE_NAME").VALUE) & """")
				If intShowObjCount = 1 Then
					Response.Write(" selected")
				End If
				Response.Write(">")
				Response.Write(Server.HTMLEncode(rsTemp.Fields("TABLE_NAME").Value))
				Response.Write("</option>")			
			End If
		End If
		rsTemp.movenext
	Loop
	Response.Write("</select><p>")
	Response.Write("<input name=""SelectTables"" type=""submit"" value=""Select Fields &gt;"" class=""submit"">")
	Response.Write(" <input type=""submit"" name=""ShowAll1"" value=""Show All &gt;&gt;"" class=""submit""><p>")
	Response.Write("</form>")
	Response.Write "<form action=""" & Request.ServerVariables("SCRIPT_NAME") & """>"
	For Each qsTemp in Request.QueryString
		Select Case UCASE(qsTemp)
			Case "SQLFROM_" & UCASE(GRIDID), "OBJTOSHOW"
			Case Else
				Response.Write("<input type=""hidden"" name=""" & qsTemp & """ value=""" & Server.HTMLEncode(Request.QueryString(qsTemp)) & """>")
		End Select
	next
	Response.Write(" Show: ")
	Response.Write("<input name=""objtoshow"" value=""Tables"" type=""radio""  onclick=""javascript:document.forms[1].submit()""")
	If Request.QueryString("objtoshow") = "" or Request.QueryString("objtoshow") = "Tables" Then 
		Response.Write(" checked")
	End If
	Response.Write(">Tables")
	Response.Write("<input name=""objtoshow"" value=""Queries"" type=""radio"" onclick=""javascript:document.forms[1].submit()""")
	If Request.QueryString("objtoshow") = "Queries" Then 
		Response.Write(" checked")
	End If
	Response.Write(">Views")
	Response.Write("<input name=""objtoshow"" value=""Both"" type=""radio"" onclick=""javascript:document.forms[1].submit()"" ")
	If Request.QueryString("objtoshow") = "Both" Then 
		Response.Write("checked")
	End If
	Response.Write(" >Both")
End If
Response.Write("</form>")

Call WriteFooter("")

Response.End

Function clns (strSQLTableString) 
	If instr(strSQLTableString,"""") = 0 And ((instr(strSQLTableString,".") = 0 And ocdDatabaseType = "Oracle") or ocdDatabaseType <> "Oracle") And instr(strSQLTableString,"[")=0 Then
	Select Case ocdDatabaseType 
		Case "Access" 
		strSQLTableString = Replace(strSQLTableString,"]","")
		strSQLTableString = Replace(strSQLTableString,"[","")
		clns = "[" & strSQLTableString & "]"
		Case "MySQL"
		strSQLTableString = Replace(strSQLTableString,"`","")
		clns = "`" & strSQLTableString & "`"
		Case "IXS","ADSI"
		clns = strSQLTableString 
		Case Else '"SQLServer","Oracle"
		strSQLTableString = Replace(strSQLTableString,"]","")
		strSQLTableString = Replace(strSQLTableString,"[","")
		strSQLTableString = Replace(strSQLTableString,"""","")
		clns = """" & strSQLTableString & """"
	End Select
	Else
		clns = strSQLTableString
	End If
end function

%>
