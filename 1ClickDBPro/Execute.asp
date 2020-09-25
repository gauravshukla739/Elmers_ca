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
<!--#INCLUDE FILE=ocdCommand.asp-->
<%
Dim blnHasParams, strProcName, arrProcName,  cmdStruct, prmTemp, objCommand, catADOX, cmdExec
blnHasParams = False
If Not ocdShowSQLExecutor Then
	Call WriteHeader("")
	Call WriteFooter("Permission Denied")
	Response.End
End If
Call WriteHeader("")
Response.Write("<span class=""Information"">Procedures : " & Server.HTMLEncode(Request.Querystring("SQLFROM")) & "</span>")
Response.Write("<form action=""" & request.servervariables("SCRIPT_NAME") & "?" & Request.QueryString & """ method=""post""><input type=""hidden"" name=""x1"" value=on>")
If instr(Cstr(request.querystring("SQLFrom")),";") = 0 Then
	strProcName = request.querystring("SQLFrom")
Else
	arrProcName = split(Cstr(request.querystring("SQLFrom")),";")
	strProcName = arrProcName(0)
End If
'Set ocdTargetConn = Server.CreateObject("ADODB.Connection")
'Call ocdTargetConn.Open (ndnscSQLConnect, ndnscSQLUser, ndnscSQLPass)
If (ocdDatabaseType = "SQLServer" ) Then
   Set cmdStruct = Server.CreateObject("ADODB.Command")
   Set cmdStruct.ActiveConnection = ocdTargetConn
   ' Specify the name of the stored procedure you wish to call
    cmdStruct.CommandText =  GetSQLIDFPart(Request.Querystring("SQLFrom"),"SQLOBJECTNAME", ocdQuotePrefix,ocdQuoteSuffix)
    cmdStruct.CommandType = &H0004 'adCmdStoredProc
    cmdStruct.Parameters.Refresh
Else
	Set catADOX = server.createobject("adox.catalog")
	If Err.Number <> 0 then
		Call WriteFooter("Access Procedures can not be run because an ADOX catalog object could not be created on the server")
	End If		
	catADOX.ActiveConnection = ocdTargetConn
End If
Response.Write("<table border=""1"">") 
If (ocdDatabaseType <> "SQLServer") Then
	If catADOX.Procedures(Cstr(request.querystring("sqlfrom"))).command.parameters.count > 0 Then
		blnHasParams = true
		For Each prmTemp In catADOX.Procedures(Cstr(request.querystring("sqlfrom"))).command.parameters
			Response.Write("<tr><td>")
			Response.Write(Server.HTMLEncode(prmTemp.name))
			Response.Write("</td><td>")
			Select Case prmTemp.type
				Case 0 'adEmpty
					Response.Write("Empty")
				Case 16 'adTinyInt 
					Response.Write("TinyInt")
				Case 2 'adSmallInt
					Response.Write("SmallInt")
				Case 3 'adInteger 
					Response.Write("Integer")
				Case 20 'adBigInt 
					Response.Write("BigInt")
				Case 17 'adUnsignedTinyInt 
					Response.Write("UnsignedTinyInt")
				Case 18 'adUnsignedSmallInt
					Response.Write("UnsignedSmallInt")
				Case 19 'adUnsignedInt
					Response.Write("UnsignedInt")
				Case 21 'adUnsignedBigInt
					Response.Write("UnsignedBigInt")
				Case 4 'adSingle
					Response.Write("Single")
				Case 5 'adDouble
					Response.Write("Double")
				Case 6 'adCurrency
					Response.Write("Currency")
				Case 14 'adDecimal
					Response.Write("Decimal")
				Case 131 'adNumeric
					Response.Write("Numeric")
				Case 11 'adBoolean
					Response.Write("Boolean")
				Case 10 'adError
					Response.Write("Error")
				Case 132 'adUserDefined
					Response.Write("UserDefined")
				Case 12 'adVariant
					Response.Write("Variant")
				Case 9 'adIDispatch
					Response.Write("IDispatch")
				Case 13 'adIUnknown
					Response.Write("IUnknown")
				Case 72 'adGUID
					Response.Write("GUID")
				Case 7 'adDate
					Response.Write("Date")
				Case 133 'adDBDate
					Response.Write("DBDate")
				Case 134 'adDBTime
					Response.Write("DBTime")
				Case 135 'adDBTimeStamp
					Response.Write("DBTimeStamp")
				Case 8 'adBSTR
					Response.Write("BSTR")
				Case 129 'adChar
					Response.Write("Char")
				Case 200 'adVarChar
					Response.Write("VarChar")
				Case 201 'adLongVarChar
					Response.Write("LongVarChar")
				Case 130 'adWChar
					Response.Write("WChar")
				Case 202 'adVarWChar
					Response.Write("VarWChar")
				Case 203 'adLongVarWChar
					Response.Write("LongVarWChar")
				Case 128 'adBinary 
					Response.Write("Binary")
				Case 204 'adVarBinary
					Response.Write("VarBinary")
				Case 205 'adLongVarBinary
					Response.Write("LongVarBinary")
				Case else
					Response.Write("Type " & prmTemp.type)
			End Select
			Response.Write("</TD>")
			Response.Write("<TD>")
			Select Case prmTemp.type
				Case 129, 200, 201, 130, 202
					Response.Write(prmTemp.size)
				Case Else
					Response.Write "&nbsp;"
			End Select
			Response.Write("</TD><TD>")
			Response.Write("<INPUT TYPE=HIDDEN NAME=""ocdInPN" & server.htmlencode(prmTemp.name) & """ VALUE=""" & server.htmlencode(prmTemp.name) & """>")
			Response.Write("<INPUT TYPE=HIDDEN NAME=""ocdInPD" & server.htmlencode(prmTemp.name) & """ VALUE=""" & server.htmlencode(prmTemp.direction) & """>")
			Response.Write("<INPUT TYPE=HIDDEN NAME=""ocdInPS" & server.htmlencode(prmTemp.name) & """ VALUE=""" & server.htmlencode(prmTemp.size) & """>")
			Response.Write("<INPUT TYPE=HIDDEN NAME=""ocdInPT"  & server.htmlencode(prmTemp.name) & """ VALUE=""" & server.htmlencode(prmTemp.type) & """>")
      Select Case  prmTemp.direction 
				Case &H0001, &H0003
					Response.Write("<INPUT TYPE=TEXT NAME=""ocdInPV"  & server.htmlencode(prmTemp.name) & """ VALUE=""")
					Response.Write(server.htmlencode(request.Form("ocdINPV" & server.htmlencode(prmTemp.name))))
					Response.Write(""">")
			End Select
			Response.Write("</td></TR>")
		Next
	End If
Else
	For Each prmTemp In cmdStruct.Parameters
    Select Case  prmTemp.direction 
			Case &H0000, &H0001, &H0003
				Response.Write("<TR><TD>")
				Response.Write(prmTemp.name)
				Response.Write("</TD><TD>")
				Select Case prmTemp.type
					Case 0 'adEmpty
						Response.Write("Empty")
					Case 16 'adTinyInt 
						Response.Write("TinyInt")
					Case 2 'adSmallInt
						Response.Write("SmallInt")
					Case 3 'adInteger 
						Response.Write("Integer")
					Case 20 'adBigInt 
						Response.Write("BigInt")
					Case 17 'adUnsignedTinyInt 
						Response.Write("UnsignedTinyInt")
					Case 18 'adUnsignedSmallInt
						Response.Write("UnsignedSmallInt")
					Case 19 'adUnsignedInt
						Response.Write("UnsignedInt")
					Case 21 'adUnsignedBigInt
						Response.Write("UnsignedBigInt")
					Case 4 'adSingle
						Response.Write("Single")
					Case 5 'adDouble
						Response.Write("Double")
					Case 6 'adCurrency
						Response.Write("Currency")
					Case 14 'adDecimal
						Response.Write("Decimal")
					Case 131 'adNumeric
						Response.Write("Numeric")
					Case 11 'adBoolean
						Response.Write("Boolean")
					Case 10 'adError
						Response.Write("Error")
					Case 132 'adUserDefined
						Response.Write("UserDefined")
					Case 12 'adVariant
						Response.Write("Variant")
					Case 9 'adIDispatch
						Response.Write("IDispatch")
					Case 13 'adIUnknown
						Response.Write("IUnknown")
					Case 72 'adGUID
						Response.Write("GUID")
					Case 7 'adDate
						Response.Write("Date")
					Case 133 'adDBDate
						Response.Write("DBDate")
					Case 134 'adDBTime
						Response.Write("DBTime")
					Case 135 'adDBTimeStamp
						Response.Write("DBTimeStamp")
					Case 8 'adBSTR
						Response.Write("BSTR")
					Case 129 'adChar
						Response.Write("Char")
					Case 200 'adVarChar
						Response.Write("VarChar")
					Case 201 'adLongVarChar
						Response.Write("LongVarChar")
					Case 130 'adWChar
						Response.Write("WChar")
					Case 202 'adVarWChar
						Response.Write("VarWChar")
					Case 203 'adLongVarWChar
						Response.Write("LongVarWChar")
					Case 128 'adBinary 
						Response.Write("Binary")
					Case 204 'adVarBinary
						Response.Write("VarBinary")
					Case 205 'adLongVarBinary
						Response.Write("LongVarBinary")
					Case Else
						Response.Write("Type " & prmTemp.type)
				End Select
				Response.Write("</TD><TD>")
		    Select Case  prmTemp.direction 
					Case  &H0003
						Response.Write("<SELECT NAME=""ocdInPD" & server.htmlencode(prmTemp.name) & """>")
		 				Response.Write("<OPTION VALUE=""" & server.htmlencode(&H0001) & """ SELECTED> Input")
		 				Response.Write("<OPTION VALUE=""" & server.htmlencode(&H0002) & """ > Output")
					Case Else
			    	Response.Write("<INPUT TYPE=HIDDEN NAME=""ocdInPD" & server.htmlencode(prmTemp.name) & """ VALUE=""" & server.htmlencode(prmTemp.direction) & """>")
				End Select
				Response.Write("</TD><TD>")
				Response.Write(prmTemp.size)
				Response.Write("</TD><TD>")
				Response.Write("<INPUT TYPE=HIDDEN NAME=""ocdInPN" & server.htmlencode(prmTemp.name) & """ VALUE=""" & server.htmlencode(prmTemp.name) & """>")
				Response.Write("<INPUT TYPE=HIDDEN NAME=""ocdInPS" & server.htmlencode(prmTemp.name) & """ VALUE=""" & server.htmlencode(prmTemp.size) & """>")
				Response.Write("<INPUT TYPE=HIDDEN NAME=""ocdInPT"  & server.htmlencode(prmTemp.name) & """ VALUE=""" & server.htmlencode(prmTemp.type) & """>")
  	    Select Case  prmTemp.direction 
					Case &H0001, &H0003
						Response.Write("<INPUT TYPE=TEXT NAME=""ocdInPV"  & server.htmlencode(prmTemp.name) & """ VALUE=""")
						Response.Write(server.htmlencode(request.Form("ocdINPV" & server.htmlencode(prmTemp.name))))
						Response.Write(""">")
				End Select
				Response.Write("</td></TR>")
			Case Else
		    Select Case  prmTemp.direction 
					Case  &H0003
						Response.Write("<SELECT NAME=""ocdInPD" & server.htmlencode(prmTemp.name) & """>")
		 				Response.Write("<OPTION VALUE=""" & server.htmlencode(&H0001) & """ SELECTED> Input")
		 				Response.Write("<OPTION VALUE=""" & server.htmlencode(&H0002) & """ > Output")
					Case Else
			    	Response.Write("<INPUT TYPE=HIDDEN NAME=""ocdInPD" & server.htmlencode(prmTemp.name) & """ VALUE=""" & server.htmlencode(prmTemp.direction) & """>")
				End Select
				Response.Write("<INPUT TYPE=HIDDEN NAME=""ocdInPN" & server.htmlencode(prmTemp.name) & """ VALUE=""" & server.htmlencode(prmTemp.name) & """>")
				Response.Write("<INPUT TYPE=HIDDEN NAME=""ocdInPS" & server.htmlencode(prmTemp.name) & """ VALUE=""" & server.htmlencode(prmTemp.size) & """>")
				Response.Write("<INPUT TYPE=HIDDEN NAME=""ocdInPT"  & server.htmlencode(prmTemp.name) & """ VALUE=""" & server.htmlencode(prmTemp.type) & """>")
		End Select
	Next
End If
Response.Write("</TABLE>")
Response.Write("<P><INPUT TYPE=Submit VALUE=""Run Procedure"" CLASS=Submit></FORM>")
If Not Request.Form = "" Then
	Set cmdExec = Server.CreateObject("ADODB.Command")
	Set cmdExec.ActiveConnection = ocdTargetConn
   ' Specify the name of the stored procedure you wish to call
	If (ocdTargetConn.Provider <> "SQLOLEDB.1") Then
   	cmdExec.CommandText =  "[" & strProcName  & "]"
	Else
   	cmdExec.CommandText = """" & strProcName & """"
	End If
  cmdExec.CommandType = 4 '&H0004 'adCmdStoredProc
	If (ocdDatabaseType <> "SQLServer") Then
		For Each prmTemp In catADOX.Procedures(Cstr(request.querystring("sqlfrom"))).command.parameters
			Select Case  prmTemp.direction 
				Case &H0000
				Case &H0001
					Call cmdExec.Parameters.Append (cmdExec.CreateParameter(prmTemp.name, prmTemp.type, prmTemp.direction,prmTemp.size, (Request.Form("ocdInPV" & prmTemp.name))))
					Response.Write(Request.Form("ocdInPSV" & prmTemp.name))
				Case &H0002
				Case &H0003
					Call cmdExec.Parameters.Append (cmdExec.CreateParameter(prmTemp.name, prmTemp.type, prmTemp.direction,prmTemp.size, (Request.Form("ocdInPV" & prmTemp.name))))
				Case &H0004
				  Call cmdExec.Parameters.Append (cmdExec.CreateParameter(prmTemp.name, prmTemp.type, prmTemp.direction, prmTemp.size))
			End Select
		Next
	Else
		For Each prmTemp In cmdStruct.Parameters
	    Select Case  prmTemp.direction 
				Case &H0000
				Case &H0001, &H0003
					If Request.Form("ocdInPV" & prmTemp.name) <> "" Then
						Call cmdExec.Parameters.Append (cmdExec.CreateParameter(prmTemp.name, prmTemp.type, prmTemp.direction,prmTemp.size, (Request.Form("ocdInPV" & prmTemp.name))))
						Response.Write(Request.Form("ocdInPSV" & prmTemp.name))
					End If
				Case &H0002
				Case  &H0003
					Call cmdExec.Parameters.Append (cmdExec.CreateParameter(prmTemp.name, prmTemp.type, CLng(Request.Form("ocdInPV" & prmTemp.name)), prmTemp.size))
				Case &H0004
			    Call cmdExec.Parameters.Append (cmdExec.CreateParameter(prmTemp.name, prmTemp.type, prmTemp.direction, prmTemp.size))
			End SElect
		Next
	End If
	Set objCommand = New ocdCommand
	Set objCommand.ADOCommand = cmdExec
	objCommand.Display("")
End If
ocdTargetConn.Close
Set ocdTargetConn = nothing
If Err.number <> 0 Then
	Response.Write "<span class=warning>" & err.description & "</span>"
	err.clear
End If
Response.Write("</BIG></BIG></BIG></FONT></FONT><P>")
call WriteFooter("")
%>      