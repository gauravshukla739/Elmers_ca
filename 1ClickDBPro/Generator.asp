<% ' Except for @ Directives, there should be no ASP or HTML codes above this line
' Setting LANGUAGE = VBScript.Encode enables mixing of encoded and unencoded ASP scripts

'1 Click DB Pro copyright 1997-2004 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com/

        
'**Start Encode**

%>
<!--#INCLUDE FILE="PageInit.asp"-->
<!--#INCLUDE FILE="ocdFunctions.asp"-->
<%
Server.ScriptTimeOut = 30
dim strWriteDetail, strWriteBrowse, strwriteEdit, strTemp, generr, strOutput, BrowseTemplateText, objBrowseTemplateFile, genwarn, objFSO, I, strSQL, rsTemp, ahnIDField, fldF, strASPFolder, strASPTemplateFolder, qsTable, strMeta, connNDTargetCW, rsNDTempCW, arrNDSchemocd, EditTemplateStartText, strEditText, objEditTemplateStartFile
dim arrNDSchemocdFields(1)
genwarn = ""
generr = ""
If Not ocdAllowCodeWiz Then
	response.clear
	response.redirect ("Schema.asp")
End if
Call WriteHeader("")
If request.querystring("sqlfrom_a") = "" Then
	Call WriteFooter("Choose a Table or View")
End if
%>
 <FORM><TABLE BORDER=0><TR><TD VALIGN=TOP><strong>Select Code then Copy and Paste</strong> the code below into new .asp files using the names suggested below.  These files must be placed in folder containing a copy of the open source <A HREF="http://www.standardreporting.net/download/1ClickDBASPLib.zip">1&nbsp;Click&nbsp;DB&nbsp;ASP&nbsp;Library</A> to run on your webserver. <TD VALIGN=TOP></TD></TR></TABLE>
</P>
<%
response.flush
if err <> 0 then
	call codewizerr
end if
Response.Write("<TABLE><TR><TD VALIGN=TOP>")
set objFSO=Server.createobject("Scripting.FileSystemObject")
if err <> 0 then
	generr = "Code Wizard could not read asp template<P>" & err.description
	call codewizerr
end if
strASPTemplateFolder = server.mappath(left(Request.ServerVariables("PATH_INFO"),instrRev(Request.ServerVariables("PATH_INFO"),"/")) )
if genwarn <> "" Then
	genwarn = genwarn & "<P>There were initialization errors <P>"
End if
Select Case UCASE(request("cwBrowseTemplate"))
	Case "CUSTOM"
		set objBrowseTemplateFile = objFSO.OpenTextFile(Cstr(strASPTemplateFolder & "\_BROWSETEMPLATECUSTOM.ASP"))
	Case "BUTTONS"
		set objBrowseTemplateFile = objFSO.OpenTextFile(Cstr(strASPTemplateFolder & "\_BROWSETEMPLATEBUTTONS.ASP"))
	Case "QUICK"
		set objBrowseTemplateFile = objFSO.OpenTextFile(Cstr(strASPTemplateFolder & "\_BROWSETEMPLATEQUICK.ASP"))
	Case "DEFAULT"
		set objBrowseTemplateFile = objFSO.OpenTextFile(Cstr(strASPTemplateFolder & "\_BrowseTemplate.asp"))
		set objEditTemplateStartFile = objFSO.OpenTextFile(Cstr(strASPTemplateFolder & "\_EditTemplate.asp"))
		EditTemplateStartText = objEditTemplateStartFile.ReadAll
		objEditTemplateSTartFile.close
	Case "EDIT"			
		set objEditTemplateStartFile = objFSO.OpenTextFile(Cstr(strASPTemplateFolder & "\_EditTemplate.asp"))
		EditTemplateStartText =  objEditTemplateStartFile.ReadAll
		objEditTemplateSTartFile.close
End Select
Select Case UCASE(request("cwBrowseTemplate"))
	Case "CUSTOM","BUTTONS","QUICK","DEFAULT"
		BrowseTemplateText = objBrowseTemplateFile.ReadAll 'All
		objBrowseTemplateFile.close		
End Select
if err <> 0 then 
	generr = "Code Wizard could not open template file, check web server permissions.<P>" & err.description
	call codewizarderror
	err.clear
end if
ahnidfield = ""
set rsNDTempCW = server.CreateObject("ADODB.Recordset")
set connNDTargetCW = server.CreateObject("ADODB.Connection")
connNDTargetCW.Open ndnscSQLConnect, ndnscSQLUser, ndnscSQLPass
if err <>0 then 
	generr = "Code Wizard could not open connection to database<P>" & err.description
	call codewizarderror
	err.clear
end if
ocdDatabaseType = getDatabaseType(connNDTargetCW)
set rsNDTempCW = connNDTargetCW.OpenSchema(20)
if err <>0 then 
	call codewizerr
end if
arrNDSchemocdFields(0) = "TABLE_TYPE"
arrNDSchemocdFields(1) = "TABLE_NAME"
arrNDSchemocd = rsNDTempCW.GetRows(,,arrNDSchemocdFields)
if err <>0 then 
	generr = "Code Wizard could not get database schema info <P>" & err.description
	call codewizerr	
end if
rsNDTempCW.Close
set rsNDTempCW = nothing
set rsTemp = server.CreateObject("ADODB.Recordset")
If err <> 0 then
	call codewizerr
End if
qsTable = FormatForSQL(Request.Querystring("SQLFrom_A"), ocdDatabaseType, "ADDSQLIDENTIFIER") 
'Determine which field is the ID
If ocdDatabaseType = "SQLServer" Then
	strSQL = "Select * from " & qsTable & " Where 1=2 "
Else
	strSQL = "Select * from " & qsTable & " Where 1=2 "
End If
response.write strSQL
rsTemp.Open strSQL,connNDTargetCW
If err <> 0 Then
	Genwarn = genwarn & "Code Wizard could not create page(s) for <I>" & qsTable &  "</i>: " & err.description & "<P>"
		err.clear
Else
	ahnIDField = ""
	If UCASE(connNDTargetCW.provider)  <> "MICROSOFT.JET.OLEDB.3.51" Then
		For Each fldF In rsTemp.Fields
			If fldF.Properties("ISAUTOINCREMENT") ="True" Then 
				ahnIDField = """" & CStr(fldF.Name)  & """"
				Exit For
			End If
		Next
	End If
	rsTemp.close
	if ocdDatabaseType = "SQLServer" Then
		strSQL = "Select * from " & qsTable & " Where 1=2"
	Else
		strSQL = "Select * from " & qsTable & " Where 1=2"
	End if
	rsTemp.Open strSQL,connNDTargetCW
	strEditText = Replace(EditTemplateStartText,"{{ocdSQLFrom}}", Replace(qsTable,"""","""""") )
	strEditText = Replace(strEditText,"{{ocdSQLSelect}}", request.querystring("sqlselect_A"))
	strEditText = Replace(strEditText,"{{ocdSQLWhereExtra}}", request.querystring("sqlwhere_A"))
	if request("ocdStartBrowse") <> "" and request("createaddeditpage") = "yes" and request("ocdAllowAdd") = "" THen
		strEditText = Replace(strEditText,"{{ocdAllowAdd}}", "False")
	Else
		strEditText = Replace(strEditText,"{{ocdAllowAdd}}", "True")
	End if
	if request("ocdStartBrowse") <> "" and request("createaddeditpage") = "yes" and request("ocdAllowEdit") = "" THen
		strEditText = Replace(strEditText,"{{ocdAllowEdit}}", "False")
	Else
		strEditText = Replace(strEditText,"{{ocdAllowEdit}}", "True")
	End if
	if request("ocdStartBrowse") <> "" and request("createaddeditpage") = "yes" and request("ocdAllowDelete") = "" THen
	 strEditText = Replace(strEditText,"{{ocdAllowDelete}}", "False")
	Else
		strEditText = Replace(strEditText,"{{ocdAllowDelete}}", "True")
	End if
	strMeta = "<TABLE>" & vbCRLF
	For Each fldF in rsTemp.Fields
		strMeta = strMeta & vbCRLF
		strMeta = strMeta & "   "    &"<TR>" & vbCRLF & "   "    & "   "    & "<TD valign=top align=right><SPAN CLASS=FieldName>" & fldF.Name & ":</SPAN></TD>" & vbCRLF
		strmeta = strMeta & "   "    & "   "    & "<TD align=left valign=top>"
		err.clear
		Select Case fldF.Type
			Case 205, 128, 204
				strMeta = strMeta & "<SPAN CLASS=Information>Binary Field</SPAN>"
			Case 201, 203
				strMeta = strMeta & "<"
				strMeta = strMeta & "%"
				strMeta = strMeta & "Call objForm.DisplayFieldAsMemo(""" & Replace(fldF.Name,"""","""""") & ""","""",""ROWS =""""5"""" COLS=""""40"""" "")" 
				strMeta = strMeta & "%"
				strMeta = strMeta & ">"
			Case 133, 135, 134, 7
				strMeta = strMeta & "<"
				strMeta = strMeta & "%"
				strMeta = strMeta & "Call objForm.DisplayFieldAsTextBox(""" & Replace(fldF.Name,"""","""""") & ""","""", ""SIZE=""""20"""" MAXLENGTH=""""50"""""")"
				strMeta = strMeta & "%"
				strMeta = strMeta & ">" 
			Case 11
				strMeta = strMeta & "<"
				strMeta = strMeta & "%"
				strMeta = strMeta & "Call objForm.DisplayFieldAsCheckBox(""" & Replace(fldF.Name,"""","""""") & """,True,False,True,"""")"	
				strMeta = strMeta & "%"
				strMeta = strMeta & ">" 
			Case 6
				strMeta = strMeta & "<"
				strMeta = strMeta & "%"
				strMeta = strMeta & "Call objForm.DisplayFieldAsTextBox(""" & Replace(fldF.Name,"""","""""") & ""","""", ""SIZE=""""12"""" MAXLENGTH=""""36"""""")"
				strMeta = strMeta & "%"
				strMeta = strMeta & ">" 
			Case 20, 14, 5, 131, 4, _
						 2, 16, 21, 19, _
					 18, 17,3
				strMeta = strMeta & "<"
				strMeta = strMeta & "%"
				strMeta = strMeta & "Call objForm.DisplayFieldAsTextBox(""" & Replace(fldF.Name,"""","""""") & ""","""", ""SIZE=""""12"""" MAXLENGTH=""""36"""""")"
				strMeta = strMeta & "%"
				strMeta = strMeta & ">"
			Case Else
				strMeta = strMeta & "<"
				strMeta = strMeta & "%"
				strMeta = strMeta & "Call objForm.DisplayFieldAsTextBox(""" & Replace(fldF.Name,"""","""""") & ""","""", ""MAXLENGTH=""""" & fldF.DefinedSize & """"" SIZE="
				strMeta = strMeta &	""""""
				if fldF.DefinedSize > 50 then 
					strMeta = strMeta & "50"
				Else
					strMeta = strMeta & fldF.DefinedSize
				End if
				strMeta = strMeta & """"""")"
				strMeta = strMeta & "%"
				strMeta = strMeta & ">" 
		End Select
		strMeta = strMeta 
		strMeta = strMeta & "</TD>" & vbCRLF & "   "    & "</TR>"& vbCRLF
	Next
	strMeta = strMeta & vbCRLF & "</TABLE>" & vbCRLF & vbCRLF 
	strEditText =  Replace(strEditText,"{{ocdFormFields}}", strMeta) 
	response.clear
	if request("ocdAllowEdit") <> "" Then
If Not CBool(ndnscCompatibility and ocdNoJavaScript) Then
%>
<input class=submit type=button value="Select Code" onClick="javascript:this.form.txtEdit.focus();this.form.txtEdit.select();">
<%
End If
%>
<SPAN CLASS=FIELDNAME>
<%
		Response.write SafeFileName(qsTable) &"_Edit.asp<P>" 
%></SPAN><P>
<TEXTAREA name=txtEdit ROWS=4 COLS=50><%
		response.write server.htmlencode(strEditText)
		Response.Write("</TEXTAREA></P><HR>")
		response.flush
	End if
	'*************************END EDIT
	'------------------------CONSTRUCT DETAIL
	strWriteDetail = ""
	if ahnIDField <> ""  then
		rsTemp.close
		strMeta = ""
		strSQL = "Select * from " & qsTable & " Where 1=2"
		rsTemp.Open strSQL,connNDTargetCW
		if right(strASPFOlder,1) = "\" then
			strASPFolder = left(StrASPFOlder,len(StrASPFolder)-1)
		End if	
		strWriteDetail = strWritedetail &  "<"
		strWriteDetail = strWritedetail &  "%" & vbCRLF
		strWriteDetail = strWritedetail &  "Option Explicit" & vbCRLF
		strWriteDetail = strWritedetail &  "on error resume next	" & vbCRLF
		strWriteDetail = strWritedetail &  "Response.Buffer=True	" & vbCRLF
		strWriteDetail = strWritedetail &  "'ocdFormat.asp Contains WriteFooter("""") and WriteHeader("""") Formatting Functions %"
		strWriteDetail = strWritedetail &  ">" & vbCRLF
		strWriteDetail = strWritedetail &  "<!--#INCLUDE FILE=ocdFormat.asp-->" & vbCRLF
		strWriteDetail = strWritedetail & "<!--#INCLUDE FILE=ocdConnectInfo.asp-->" & vbCRLF
		strWriteDetail = strWritedetail & vbCRLF & "<%" & vbCRLF
		strWriteDetail = strWritedetail & "dim connDetail, rsDetail, strSQLDetail, varValue" & vbCRLF
		strWriteDetail = strWritedetail & "set connDetail = Server.CreateObject(""ADODB.Connection"")" & vbCRLF
		strWriteDetail = strWritedetail & "connDetail.Open ocdSQLConnect, ocdSQLUser, ocdSQLPass 'ADO Connect String, including uid and pw if necessary" & vbCRLF
		strWriteDetail = strWritedetail &  "Call WriteHeader("""")" & vbCRLF
		If ocdDatabaseType <> "SQLServer" Then
			if ahnIDField <> "" Then
				strWriteDetail = strWritedetail &  "if  isnumeric(Request.Querystring(""SQLID"")) Then " & vbCRLF
				strWriteDetail = strWritedetail &  "strSQLDetail = ""Select * From " & qsTable & " WHERE [" & Replace(ahnIDField,"""","") & "]="" &  Replace(Request.Querystring(""SQLID""),"";"","""")" & vbCRLF 'SQL String" & vbCRLF
				strWriteDetail = strWritedetail & vbTab & "set rsDetail = Server.CreateObject(""ADODB.Recordset"")" & vbCRLF
				strWriteDetail = strWritedetail & vbTab &  "rsDetail.Open strSQLDetail, connDetail 'read only forward cursor recordset" & vbCRLF
				strWriteDetail = strWritedetail & vbTab &  "if err.number <> 0 Then " & vbCRLF
				strWriteDetail = strWritedetail &  vbTab &  vbTab &  "call writefooter("""")" & vbCRLF
				strWriteDetail = strWritedetail &  vbTab &  vbTab &  "response.end" & vbCRLF
				strWriteDetail = strWritedetail &  vbTab & "end if" & vbCRLF
				strWriteDetail = strWritedetail & vbTab &  "if  rsDetail.eof  Then " & vbCRLF
				strWriteDetail = strWritedetail &  vbTab &  vbTab &  "call writefooter(""Not Found"")" & vbCRLF
				strWriteDetail = strWritedetail &  vbTab &  vbTab &  "response.end" & vbCRLF
				strWriteDetail = strWritedetail &  vbTab & "end if" & vbCRLF
				strWriteDetail = strWritedetail &   "else" & vbCRLF
				strWriteDetail = strWritedetail &  vbTab &  "call writefooter(""Not Found"")" & vbCRLF
				strWriteDetail = strWritedetail &  vbTab &  "response.end" & vbCRLF
				strWriteDetail = strWritedetail &   "end if" & vbCRLF
			End if
		Else
			strWriteDetail = strWritedetail &  "if  isnumeric(Request.Querystring(""SQLID"")) Then " & vbCRLF
			strWriteDetail = strWritedetail &  vbTab & "strSQLDetail = ""Select * From " & Replace(qsTable,"""","""""") & " WHERE """"" & Replace(ahnIDField,"""","") & """""="" &  Request.Querystring(""SQLID"")  'SQL String" & vbCRLF
			strWriteDetail = strWritedetail & vbTab & "set rsDetail = Server.CreateObject(""ADODB.Recordset"")" & vbCRLF
			strWriteDetail = strWritedetail & vbTab &  "rsDetail.Open strSQLDetail, connDetail 'read only forward cursor recordset" & vbCRLF
			strWriteDetail = strWritedetail & vbTab &  "if err.number <> 0 Then " & vbCRLF
			strWriteDetail = strWritedetail &  vbTab &  vbTab &  "call writefooter("""")" & vbCRLF
			strWriteDetail = strWritedetail &  vbTab &  vbTab &  "response.end" & vbCRLF
			strWriteDetail = strWritedetail &  vbTab & "end if" & vbCRLF
			strWriteDetail = strWritedetail & vbTab &  "if  rsDetail.eof  Then " & vbCRLF
			strWriteDetail = strWritedetail &  vbTab &  vbTab &  "call writefooter(""Not Found"")" & vbCRLF
			strWriteDetail = strWritedetail &  vbTab &  vbTab &  "response.end" & vbCRLF
			strWriteDetail = strWritedetail &  vbTab & "end if" & vbCRLF
			strWriteDetail = strWritedetail &   "else" & vbCRLF
			strWriteDetail = strWritedetail &  vbTab &  "call writefooter(""Not Found"")" & vbCRLF
			strWriteDetail = strWritedetail &  vbTab &  "response.end" & vbCRLF
			strWriteDetail = strWritedetail &   "end if" & vbCRLF
		End if
		strWriteDetail = strWritedetail & "%"
		strWriteDetail = strWritedetail &  ">" & vbCRLF
		strWriteDetail = strWritedetail &  "<!-- Start Write Error Status -->" & vbCRLF
		strWriteDetail = strWritedetail &  vbCRLF	
		strWriteDetail = strWritedetail &  vbCRLF & vbCRLF & "<P>" & vbCRLF & vbCRLF
		strWriteDetail = strWritedetail &  "<"
		strWriteDetail = strWritedetail &  "%" & vbCRLF
		strWriteDetail = strWritedetail &  "if err.number <> 0  Then " & vbCRLF
		strWriteDetail = strWritedetail &  "call writefooter("""")" & vbCRLF
		strWriteDetail = strWritedetail &  "response.end" & vbCRLF
		strWriteDetail = strWritedetail &  "end if" & vbCRLF
		strWriteDetail = strWritedetail &  "%"
		strWriteDetail = strWritedetail &  ">" & vbCRLF
		strMeta = strMeta & "<TABLE>" & vbCRLF
		For Each fldF in rsTemp.Fields
			strMeta = strMeta & vbCRLF
			strMeta = strMeta & "   "    &"<TR>" & vbCRLF & "   "    & "   "    & "<TD valign=top align=right><SPAN CLASS=FieldName>" & fldF.Name & ":</SPAN></TD>" & vbCRLF
			strmeta = strMeta & "   "    & "   "    & "<TD align=left valign=top>"
			Select Case fldF.Type
				Case 205, 128, 204
					strMeta = strMeta & "<SPAN CLASS=Information>Binary Field</SPAN>"
				Case 201, 203
					strMeta = strMeta & "<"
					strMeta = strMeta & "%" & vbCRLF
					strMeta = strMeta & "varValue = rsDetail(""" & fldF.Name & """)" & vbCRLF
					strMeta = strMeta & "if not isnull(varValue) Then" & vbCRLF
					strMeta = strMeta & "Response.write Server.HTMLEncode(varValue)" & vbCRLF
					strMeta = strMeta & "End if" & vbCRLF
					strMeta = strMeta & "%"
					strMeta = strMeta & ">"
				Case 133, 135, 134, 7
					strMeta = strMeta & "<"
					strMeta = strMeta & "%" & vbCRLF
					strMeta = strMeta & "varValue = rsDetail(""" & fldF.Name & """)" & vbCRLF
					strMeta = strMeta & "if not isnull(varValue) Then" & vbCRLF
					strMeta = strMeta & "Response.write Server.HTMLEncode(varValue)" & vbCRLF
					strMeta = strMeta & "End if" & vbCRLF
					strMeta = strMeta & "%"
					strMeta = strMeta & ">"
				Case 11
					strMeta = strMeta & "<"
					strMeta = strMeta & "%" & vbCRLF
					strMeta = strMeta & "varValue = rsDetail(""" & fldF.Name & """)" & vbCRLF
					strMeta = strMeta & "if not isnull(varValue) Then" & vbCRLF
					strMeta = strMeta & "Response.write Server.HTMLEncode(varValue)" & vbCRLF
					strMeta = strMeta & "End if" & vbCRLF
					strMeta = strMeta & "%"
					strMeta = strMeta & ">"
				Case 6
					strMeta = strMeta & "<"
					strMeta = strMeta & "%" & vbCRLF
					strMeta = strMeta & "varValue = rsDetail(""" & fldF.Name & """)" & vbCRLF
					strMeta = strMeta & "if not isnull(varValue) Then" & vbCRLF
					strMeta = strMeta & "Response.write Server.HTMLEncode(varValue)" & vbCRLF
					strMeta = strMeta & "End if" & vbCRLF
					strMeta = strMeta & "%"
					strMeta = strMeta & ">"
				Case 20, 14, 5, 131, 4, _
				 2, 16, 21, 19, _
				 18, 17,3
					strMeta = strMeta & "<"
					strMeta = strMeta & "%" & vbCRLF
					strMeta = strMeta & "varValue = rsDetail(""" & fldF.Name & """)" & vbCRLF
					strMeta = strMeta & "if not isnull(varValue) Then" & vbCRLF
					strMeta = strMeta & "Response.write Server.HTMLEncode(varValue)" & vbCRLF
					strMeta = strMeta & "End if" & vbCRLF
					strMeta = strMeta & "%"
					strMeta = strMeta & ">"
				Case Else
					strMeta = strMeta & "<"
					strMeta = strMeta & "%" & vbCRLF
					strMeta = strMeta & "varValue = rsDetail(""" & fldF.Name & """)" & vbCRLF
					strMeta = strMeta & "if not isnull(varValue) Then" & vbCRLF
					strMeta = strMeta & "Response.write Server.HTMLEncode(varValue)" & vbCRLF
					strMeta = strMeta & "End if" & vbCRLF
					strMeta = strMeta & "%"
					strMeta = strMeta & ">"
			End Select
			strMeta = strMeta 
			strMeta = strMeta & "</TD>" & vbCRLF & "   "    & "</TR>"& vbCRLF
			strWriteDetail = strWritedetail &  strMeta
			strMeta = ""
		Next
		strWriteDetail = strWritedetail &  "</TABLE>" & vbCRLF
		strWriteDetail = strWritedetail &  "<p>" & vbCRLF
		strWriteDetail = strWritedetail &  vbCRLF
		strWriteDetail = strWritedetail &  "<A HREF=""<"
		strWriteDetail = strWritedetail &  "%=Replace(Request.ServerVariables(""SCRIPT_NAME""),""_Detail.asp"",""_Edit.asp"") & ""?"" & Replace(Request.Querystring,""&"",""&amp;"")%"
		strWriteDetail = strWritedetail & 	">"">Edit this Record</A>" & vbCRLF
		strWriteDetail = strWritedetail &  vbCRLF
		strWriteDetail = strWritedetail &  "<p>" & vbCRLF
		strWriteDetail = strWritedetail &  vbCRLF
		strWriteDetail = strWritedetail &  vbCRLF & "<"
		strWriteDetail = strWritedetail &  "%" & vbCRLF
		strWriteDetail = strWritedetail &  "Call writefooter("""")" & vbCRLF
		strWriteDetail = strWritedetail &  "%"
		strWriteDetail = strWritedetail &  ">" & vbCRLF
		If Not CBool(ndnscCompatibility and ocdNoJavaScript) Then
%>
<input class=submit type=button value="Select Code" onClick="javascript:this.form.txtDetail.focus();this.form.txtDetail.select();">
<%
End If
%>
<SPAN CLASS=FIELDNAME>
<%
		Response.write SafeFileName(qsTable) &"_Detail.asp<P>" 
%></SPAN><P>
<TEXTAREA name=txtDetail ROWS=4 COLS=50><%
		response.write server.htmlencode(strWriteDetail)
		Response.Write("</TEXTAREA></P>")
		response.write "<hr>"
		response.flush
	end if
	response.flush
'---------------------------Construct Browse
	if request.querystring("createbrowsepage") = "yes" and (request("ocdStartBrowseShowDetail") = "" and request("ocdStartBrowseShowEdit") = "" )  Then
		strOutput = browsetemplatetext
		strOutput = Replace(strOutput,"{{ocdSQLFrom}}", (Replace(qsTable,"""","""""")))
		strOutput = Replace(strOutput,"{{ocdSQLSelect}}", request.querystring("sqlselect_A"))
		strOutput = Replace(strOutput,"{{ocdSQLWhereExtra}}", request.querystring("sqlwhere_A"))
'		strOutput = Replace(strOutput,"{{ocdSQLSelectName}}", request.querystring("ocdSQLNames"))
		strOutput = Replace(strOutput,"{{ocdSQLOrderByDefault}}", request.querystring("sqlorderby_A"))
		if request.querystring("createdetailpage") = "yes" Then
			If not (ahnIDField = "" or ahnIDField = """""" ) Then
				'objOutFile.Write "objGrid.AllowDetail = True" & vbCRLF
			Else
				strTemp =  "objGrid.AllowDetail = False" & vbCRLF
			end if
		Else
			strTemp =  "objGrid.AllowDetail = False" & vbCRLF
		End if
		strOutput = Replace(strOutput,"{{ocdAllow}}", strTemp)
		strTemp = "objGrid.FormEdit = """ & SafeFileName(qsTable) &"_Edit.asp""" & vbCRLF
		strTemp = strTemp &  "objGrid.FormDetail = """ & SafeFileName(qsTable) &"_Detail.asp""" & vbCRLF
		strOutput = Replace(strOutput,"{{ocdForms}}", strTemp)
		if request("ocdStartBrowse") <> "" and request("createaddeditpage") = "yes" and request("ocdAllowAdd") = "" THen
			strOutput = Replace(strOutput,"{{ocdAllowAdd}}", "False")
		Else
			strOutput = Replace(strOutput,"{{ocdAllowAdd}}", "True")
		End if
		if request("ocdStartBrowse") <> "" and request("createaddeditpage") = "yes" and request("ocdAllowEdit") = "" THen
			strOutput = Replace(strOutput,"{{ocdAllowEdit}}", "False")
		Else
			strOutput = Replace(strOutput,"{{ocdAllowEdit}}", "True")
		End if
		if request("ocdStartBrowse") <> "" and request("createaddeditpage") = "yes" and request("ocdAllowDelete") = "" THen
			strOutput = Replace(strOutput,"{{ocdAllowDelete}}", "False")
		Else
			strOutput = Replace(strOutput,"{{ocdAllowDelete}}", "True")
		End if
		if request("ocdStartBrowse") <> "" and request("createdetailpage") = "yes" and request("ocdAllowDetail") = "" THen
			strOutput = Replace(strOutput,"{{ocdAllowDetail}}", "False")
		Else
			strOutput = Replace(strOutput,"{{ocdAllowDetail}}", "True")
		End if
		response.clear		
		If Not CBool(ndnscCompatibility and ocdNoJavaScript) Then
%>
<input class=submit type=button value="Select Code" onClick="javascript:this.form.txtBrowse.focus();this.form.txtBrowse.select();">
<%
End If
%>
<SPAN CLASS=FIELDNAME>
<%
		Response.write SafeFileName(qsTable) &"_Browse.asp<P>" 
%></SPAN><P>
<TEXTAREA name=txtBrowse ROWS=4 COLS=50><%
		response.write server.htmlencode(strOutput)
		Response.Write("</TEXTAREA></P>")
		response.flush
	end if 'construct browse
'-------------------------------End Construct Browse
	rsTemp.close
End if 'could not open recordset
response.clear
%>
</TD><TD VALIGN=TOP>
<P>Use <B>Ctrl-C</B>, <B>Ctrl-V</B> for copy and paste.<P>ASP source code can be edited with Visual Studio, FrontPage, Macromedia MX, or Notepad.</P><small>Browse/Search/Export and Add/Edit/Delete ASP files must be generated separately.  The name of the Edit page must match the name specified in the Browse page by the command objGrid.FormEdit = "ObjectName_Edit.asp". </small>
</TD></TR></TABLE></FORM>
<%
if genwarn <> "" Then
	call writefooter(genwarn)
	response.write "<hr><p><span class=""warning"">Warnings :</span></p>"
	response.write genwarn
	response.write "<hr>"
end if
response.flush

Call writefooter("")

Sub CodeWizErr ()
	call writeheader("")
	response.write "<SPAN CLASS=Information>1 Click DB Code Wizard Was Not Successful</SPAN>"
	Response.write "<P><TABLE><TR><TD VALIGN=TOP ALIGN=LEFT><IMG SRC=appWarning.gif ALT=Warning></td><TD VALIGN=TOP ALIGN=LEFT>"
	if generr <> "" Then
		Response.write  Generr 
	Else
		response.write err.description
		err.clear
	End if
	Response.write "</td></tr></table>	"
	call writefooter("")
	response.end
end sub
function safefilename(byVal strToFormat)
	safefilename =  Replace(Replace(Replace(Replace(FormatForSQL(strToFormat,ocdDatabaseType,"REMOVESQLIDENTIFIER")," ","_"),"/","_"),"\","_"),".","_") 
end function
%>
