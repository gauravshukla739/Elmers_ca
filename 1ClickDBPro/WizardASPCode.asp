
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

<%
Dim GRIDID, qsTemp, strRedirect, connTarget, strTempField, fldTemp, rsTemp, objDoubleBox, blnHasTotals, intCount
GRIDID = "A"
strRedirect = ""
Call writeheader("")
Set connTarget = Server.Createobject("ADODB.Connection")
Call connTarget.open (ndnscSQLConnect , ndnscSQLUser , ndnscSQLPass)
If err.number <> 0 Then
	Call WriteFooter("")
End If
Set rsTemp = Server.createobject("ADODB.Recordset")
If Request.QueryString("sqlfrom_a") = "" Then
	Response.Write("<span class=Information>Select ")
	Select Case UCase(Request.QueryString("objtoshow"))
		Case "QUERIES"
			Response.Write "View"
		Case "BOTH"
			Response.Write "Table Or View"
		Case Else
			Response.Write "Table"
	End Select
	Response.Write(" to Generate ASP Code</span> ")
	Response.Write "<form action=""Generator.asp"" method=""get"">"
	Response.Write("<input type=""hidden"" name=""createbrowsepage"" value=""yes"">")
	Response.Write("<input type=""hidden"" name=""cwBrowseTemplate"" value=""Default"">")
	Response.Write("<input name=""createdetailpage"" value=""yes"" type=""hidden"">")
	Response.Write("<select name=""sqlfrom_" & GRIDID & """ size=""12"">")
	If ocdDatabaseType = "Oracle" Then
		Set rsTemp = connTarget.Execute ("SELECT OBJECT_TYPE AS TABLE_TYPE, OBJECT_NAME AS TABLE_NAME, OWNER AS TABLE_SCHEMA FROM ALL_OBJECTS WHERE (OBJECT_TYPE = 'TABLE' Or OBJECT_TYPE = 'VIEW') And NOT OWNER = 'SYS' And NOT OWNER = 'WKSYS' And NOT OWNER = 'MDSYS' And NOT OWNER = 'OLAPSYS' And NOT OWNER ='CTXSYS' And NOT OWNER='SYSTEM'")
	Else
		Set rsTemp = connTarget.OpenSchema(20) 'adSchemaTables
	End If
	intCount = 0
	Do While Not rsTemp.EOF
		If ((rsTemp.Fields("TABLE_TYPE").Value = "TABLE" Or rsTemp.Fields("TABLE_TYPE").Value = "VIEW") And UCase(LEFT(rsTemp.Fields("TABLE_NAME").Value,4)) <> "MSYS")  Then
			If ((Request.QueryString("objtoshow") = "Both" Or Request.QueryString("objtoshow") = "" Or Request.QueryString("objtoshow")="Tables") And (rsTemp("TABLE_TYPE").Value = "TABLE")) Or ((Request.QueryString("objtoshow")="Queries" Or Request.QueryString("objtoshow")="Both") And rsTemp("TABLE_TYPE").Value = "VIEW") Then
				Response.Write("<option value=""" & Server.HTMLEncode(rsTemp.Fields("TABLE_NAME").Value) & """")
				If intCount = 0 Then
					Response.Write(" selected")
				End If
				Response.Write(">")
Response.Write(Server.HTMLEncode(rsTemp.Fields("TABLE_NAME").Value))
				Response.Write("</option>")			
				intCount = intCount + 1
			End If	
		End If
		rsTemp.MoveNext
	Loop
	Response.Write("</select><p>")
	Response.Write("<input type=Submit name=ocdStartBrowse class=""submit"" value=""")
	Response.Write(Server.HTMLEncode("Generate Code >>"))
	Response.Write(""">")

Response.Write(" &nbsp; <input type=""hidden"" name=""createaddeditpage"" value=""yes"">")
		Response.Write("Add: <input name=ocdAllowAdd type=""checkbox"" CHECKED> &nbsp; ")
		Response.Write("Edit: <input name=ocdAllowEdit type=""checkbox"" CHECKED> &nbsp; ")
		Response.Write("Delete: <input name=ocdAllowDelete type=""checkbox"" CHECKED> ")
	Response.Write("</form>")
	If Not CBool(CInt(ndnscCompatibility) And ocdNoJavaScript) Then
	Response.Write "<FORM ACTION=""" & Request.ServerVariables("SCRIPT_NAME") & """>"
	for each qsTemp in Request.QueryString
		Select Case UCase(qsTemp)
			Case "SQLFROM_" & UCase(GRIDID), "OBJTOSHOW"
			Case Else
				Response.Write("<input type=Hidden name=""" & qsTemp & """ value=""" & Server.HTMLEncode(Request.QueryString(qsTemp)) & """>")
		End Select
	next
	Response.Write(" Show: ")
	Response.Write("<input name=""objtoshow"" value=""Tables"" type=""radio"" onclick=""javascript:document.forms[1].submit()""")
	If Request.QueryString("objtoshow") = "" Or Request.QueryString("objtoshow") = "Tables" Then 
		Response.Write(" checked")
	End If
	Response.Write(">Tables")
	Response.Write("<input name=""objtoshow"" value=""Queries"" type=""radio"" onclick=""javascript:document.forms[1].submit()""")
	If Request.QueryString("objtoshow") = "Queries" Then 
		Response.Write(" checked")
	End If
	Response.Write(">Views")
	Response.Write("<input name=""objtoshow"" value=""Both"" type=""radio"" onclick=""javascript:document.forms[1].submit()""")
	If Request.QueryString("objtoshow") = "Both" Then 
		Response.Write(" checked")
	End If
	Response.Write(">Both")
	End If
	Response.Write("</form>")
End If
Call WriteFooter("")
%>