<% ' Except for @ Directives, there should be no ASP or HTML codes above this line
' Setting LANGUAGE = VBScript.Encode enables mixing of encoded and unencoded ASP scripts

'1 Click DB Pro copyright 1997-2004 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com/

        
'**Start Encode**

%>
<!--#INCLUDE FILE=PageInit.asp-->
<!--#INCLUDE FILE=ocdFunctions.asp-->
<!--#INCLUDE FILE=ocdLoadBinary.asp-->
<%
Response.Buffer = True
Server.ScriptTimeout = 9900
Dim strSQLConnectImport, strSQLUserImport, strSQLPassImport, rsTemp, GRIDID, strSelect, rsField, ocdConnImport, strImpCStat, blnConnectOk, intBatchSize, blnFail, varEval, varEvalExpr, intICount, intSICount, varValue, fvTemp2, fvTemp, rsTemp2, rsField2, rsImportSource, rsImportTarget, arrFile, strLine, objloader, varFileData, arrLine, strImportFrom, strTargetTable, strImportType, strIgnoreFieldErrors, arrFieldMap, blnOkayToMap

If Request.QueryString("loadimportspec") <> "" Then
	Call WriteHeader("")
	Set objloader = new Loadit
	If Err.Number <> 0 Then
		Call WriteFooter("Load Error")
	End If
	varFileData = objloader.GetFileInput()
	If Err.Number <> 0 Then
		Call WriteFooter("Load Error")
	End If
	If varFileData = "" Then
		Call WriteFooter("Load Error")
	End If
	arrFile = split(varFileData, vbCRLF)
	If UBound(arrFile) = 0 Then
		Call WriteFooter("Load Error")
	End if
	Set objloader = Nothing
	If Err.Number <> 0 Then
		Call WriteFooter("LoadError")
	End If
	For Each strLine In arrFile
		If Trim(strLine) = "" Then
		ElseIf Trim(UCASE(strLine)) = "**BEGIN 1 CLICK DB IMPORT" Then
			exit for
		Else
			Call WriteFooter("Incorrect File Format")
		End If
		If Err.Number <> 0 Then
			Call WriteFooter("Incorrect File Format")
		End If
	Next
	If Err.Number <> 0 Then
		Call WriteFooter("Incorrect File Format")
	End If
	For Each strLine In arrFile
		If Trim(strLine)="" Then
		Else
			arrLine = split(strLine,"=")
			select case Replace(UCase(Trim(Cstr(arrLine(0))))," ","")
				Case "IMPORTSOURCE"
					session("ocdSQLCOnnectImport") = Mid(strLine,instr(strLIne,"=")+1)
				Case "IMPORTUSER"
					session("ocdSQLUserImport") = arrLine(1)
				Case "IMPORTPASSWORD"
					session("ocdSQLPassImport") = arrLine(1)
				Case "IMPORTFROM"
					strImportFrom = arrLine(1)
				Case "TARGETTABLE"
					strTargetTable = arrLine(1)
				Case "IMPORTTYPE"
					strImportType = arrLine(1)
				Case "IGNOREFIELDERRORS"
					strIgnoreFieldErrors = arrLine(1)
				Case "FIELDMAPPING"
					exit for			
				Case Else
			End Select
		End If
	Next
	Response.Write "<FORM ACTION=""" & ocdPageName & "?" 
	Response.Write "sqlfromimport_A=" & server.urlencode(strImportFrom)
	Response.Write "&amp;sqlfrom_A=" & server.urlencode(strTargetTable)
	Response.Write """ METHOD=POST>"
	Response.Write "<SPAN CLASS=Information>Import Specification Loaded</SPAN> "
	Response.Write " <INPUT TYPE=SUbmit CLASS=Submit Name=loadImportSpec Value=""Proceed &gt; &gt;"">"
	Response.Write "<INPUT TYPE=Hidden NAME=ImportTrans VALUE=""" & server.HTMLEncode(strImportType) & """>"
	Response.Write "<INPUT TYPE=Hidden NAME=TranContinueOnError VALUE=""" & server.htmlencode(strIgnoreFieldErrors) & """>"
	blnOkayToMap = False
	For Each strLine in arrFile
		If not blnOkayToMap then
			If Replace(UCASE(Trim(strLine))," ","") = "FIELDMAPPING=" Then
				blnOkayToMap = True
			End If
		Else
			If Left(strLine,5) = "**END" Then
				Exit for
			Else
				arrFieldMap = split(strLine,"::")
				Response.Write "<INPUT TYPE=HIDDEN NAME=""ocdImFld"
				Response.Write server.htmlencode(Trim(Cstr(arrFieldMap(0))))
				Response.Write """ VALUE="""
				Response.Write server.htmlencode(Trim(Cstr(arrFieldMap(1))))
				Response.Write """>"
				Response.Write "<INPUT TYPE=HIDDEN NAME=""ocdImExp"
				Response.Write server.htmlencode(Trim(Cstr(arrFieldMap(0))))
				Response.Write """ VALUE="""
				Response.Write server.htmlencode(Trim(Cstr(arrFieldMap(2))))
				Response.Write """>"
			End If
		End If
	Next
	Response.Write "<PRE>"
	Response.Write Server.htmlencode(varFileData)
	Response.Write "</PRE>"
	Response.Write "</FORM>"
	Call WriteFooter("")
	Response.end
End If

If Request.Form("Action") =  "Connect" Then
	Session("ocdSQLCOnnectImport") = Request.Form("ADOConnectImportSource")
	Session("ocdSQLUserImport") = Request.Form("ADOUserImportSource")
	Session("ocdSQLPassImport") = Request.Form("ADOPassImportSource")
End If

strSQLCOnnectImport = session("ocdSQLCOnnectImport")
strSQLUserImport = session("ocdSQLUserImport")
strSQLPassImport = session("ocdSQLPassImport")
GRIDID = "A"
intBatchSize = 50
blnConnectOK = False
strImpCStat = ""
strSelect = ""

Set rsTemp = server.createobject("ADODB.Recordset")
Set ocdConnImport = server.createobject("ADODB.Connection")

Call WriteHeader("")

'********* Begin Show Import DataSource Connect
If strSQLConnectImport <> ""  Then
	Call ocdConnImport.open (strSQLCOnnectImport, strSQLUserImport , strSQLPassImport)
	If Err.Number <> 0 then 
		strImpCStat = err.description
		Err.clear
		Response.Write("<p>" & session("ocdSQLCOnnectImport") & "</p>")
		Session("ocdSQLCOnnectImport") = ""
		Session("ocdSQLUserImport") = ""
		Session("ocdSQLPassImport") = ""
	Else
		blnCOnnectOK = True
	End If
End If
If (Not blnConnectOk) Or Request.QueryString("reconnect") <> "" Then 
		If Not blnConnectOK then
			Response.Write "<p><span class=""warning"">Import Data Source Connection Required"
			If strImpCstat <> "" THen
				Response.Write " : " & strImpCStat
			End If
			Response.Write "</span></p>"
		End If		
		Response.Write("<UL><LI><A HREF=Connect.asp?sourcecontext=import class=menu><IMG SRC=appConnect.gif BORDER=0>Connect Wizard</A> &lt;-- Click to specify ADO connection information<P></LI><LI><form method=""post"" enctype=""multipart/form-data"" action=""" & ocdPageName & "?loadimportspec=true""> or Browse..., Select and Load a previously saved 1 Click DB Import Spec <BR>	<input type=""file"" name=""file"" size=""25""> 		<input type=""submit"" value=""Load Import Spec""></form></LI></UL>")
	Call WriteFooter("")
	Response.End
End If 
'********* End Show Import DataSource Connect
'Call ocdTargetConn.open (ndnscSQLConnect , ndnscSQLUser , ndnscSQLPass)
if Request.QueryString("sqlfromimport_" & GRIDID) = "" Then
	Response.Write("<SPAN CLASS=Information>Import Source : </SPAN>")
	Response.Write(Session("ocdSQLConnectImport"))
	Response.Write("<A HREF=WizardImport.asp?reconnect=yes>")
	Response.Write( "(Click to Change)</A> ")
	Response.Write("<P><SPAN CLASS=Information>Select Data Import From : </SPAN></P>")
	Response.Write("<BLOCKQUOTE>")
	Response.Write("<FORM METHOD=GET ACTION=""" & 	Request.servervariables("SCRIPT_NAME") & """>")
	Response.Write("<SELECT NAME=""sqlfromimport_" & GRIDID & """ SIZE=12>")
	Set rsTemp = ocdConnImport.OpenSchema(20) 'adSchemaTables
	If err <> 0 then
		Call WriteFooter("")
		Response.end
	End If 
	ocdDataBaseType = getDatabaseType(ocdConnImport)
	Do While Not rsTemp.EOF
		Select Case ocdDatabaseType
			Case "Access"
				If ((rsTemp.Fields("TABLE_TYPE").Value = "TABLE" AND UCASE(LEFT(rsTemp.Fields("TABLE_NAME").Value,4)) <> "MSYS"))  Then
					Response.Write("<OPTION VALUE=""" & server.htmlencode("[" & rsTemp.Fields("TABLE_NAME").VALUE & "]") & """>")
					Response.Write(Server.HTMLEncode(rsTemp.Fields("TABLE_NAME").Value))
					Response.Write("</OPTION>")			
				End If
			Case "SQLServer"
				If ((rsTemp.Fields("TABLE_TYPE").Value = "TABLE" AND UCASE(LEFT(rsTemp.Fields("TABLE_NAME").Value,4)) <> "MSYS"))  Then
					Response.Write("<OPTION VALUE=""" & server.htmlencode("""" & rsTemp.Fields("TABLE_SCHEMA").VALUE & """" & "." & """" & rsTemp.Fields("TABLE_NAME").VALUE & """") & """>")
					Response.Write(Server.HTMLEncode("""" & rsTemp.Fields("TABLE_SCHEMA").Value & """" & "." & """" & rsTemp.Fields("TABLE_NAME").Value & """"))
					Response.Write("</OPTION>")			
				End If
			Case "Oracle"
				If ((rsTemp.Fields("TABLE_TYPE").Value = "TABLE" AND UCASE(LEFT(rsTemp.Fields("TABLE_NAME").Value,4)) <> "MSYS"))  Then
					Response.Write("<OPTION VALUE=""" & server.htmlencode("""" & rsTemp.Fields("TABLE_SCHEMA").VALUE & """" & "." & """" & rsTemp.Fields("TABLE_NAME").VALUE & """") & """>")
					Response.Write(Server.HTMLEncode("""" & rsTemp.Fields("TABLE_SCHEMA").Value & """" & "." & """" & rsTemp.Fields("TABLE_NAME").Value & """"))
					Response.Write("</OPTION>")			
				End If
			Case ELse
				If ((rsTemp.Fields("TABLE_TYPE").Value = "TABLE" AND UCASE(LEFT(rsTemp.Fields("TABLE_NAME").Value,4)) <> "MSYS"))  Then
					Response.Write("<OPTION VALUE=""" & server.htmlencode(rsTemp.Fields("TABLE_NAME").VALUE) & """>")
					Response.Write(Server.HTMLEncode(rsTemp.Fields("TABLE_NAME").Value))
					Response.Write("</OPTION>")			
				End If
		End Select
		rsTemp.movenext
	Loop
	Response.Write("</SELECT><P>")
	Response.Write("<INPUT NAME=""SelectTables"" TYPE=Submit Value=""Next &gt;"" Class=submit>")
	Response.Write("<P>")
	Response.Write("</FORM>")
Elseif Request.QueryString("sqlfrom_" & GRIDID) = "" and Request.QueryString("sqlfromimport_" & GRIDID) <> "" THen 
	Response.Write("<SPAN CLASS=Information>Import Source : </SPAN>")
	Response.Write Session("ocdSQLConnectImport")
	Response.Write "<A HREF=WizardImport.asp?reconnect=yes>"
	Response.Write(  "(Click to Change)"& "</A> <BR> ")
	Response.Write "<SPAN CLASS=Information>Import From : </SPAN>  "
	Response.Write server.htmlencode(Request.QueryString("sqlfromimport_" & GRIDID)) 
	Response.Write "<A HREF=""WizardImport.asp"">"
	Response.Write( "(Click to Change)" & "</A> ")
	Response.Write "<P><SPAN CLASS=Information>Select Target Table for Insert Into : </SPAN></P>"
	Response.Write("<BLOCKQUOTE>")
	Response.Write "<FORM METHOD=GET ACTION=""" & Request.servervariables("SCRIPT_NAME") & """>"
	For Each fvTemp in Request.QueryString
		Select case UCASE(fvTemp)
			Case "SQLFROM_" & UCASE(GRIDID), "OBJTOSHOW"
			Case Else
				Response.Write("<INPUT TYPE=Hidden NAME=""" & fvTemp & """ VALUE=""" & server.htmlencode(Request.QueryString(fvTemp)) & """>")
		End Select
	Next
	Response.Write("<SELECT NAME=""sqlfrom_" & GRIDID & """ SIZE=12>")
	set rsTemp = ocdTargetConn.OpenSchema(20) 'adSchemaTables
	If err <> 0 then
		Call WriteFooter("")
		Response.end
	End If 
'	ocdDataBaseType = getDataBaseType(ocdTargetConn)
	Do While Not rsTemp.EOF
		Select Case ocdDatabaseType
			Case "Access"
				If ((rsTemp.Fields("TABLE_TYPE").Value = "TABLE" AND UCASE(LEFT(rsTemp.Fields("TABLE_NAME").Value,4)) <> "MSYS"))  Then
					Response.Write("<OPTION VALUE=""" & server.htmlencode("[" & rsTemp.Fields("TABLE_NAME").VALUE & "]") & """>")
					Response.Write(Server.HTMLEncode(rsTemp.Fields("TABLE_NAME").Value))
					Response.Write("</OPTION>")			
				End If
			Case "SQLServer"
				If ((rsTemp.Fields("TABLE_TYPE").Value = "TABLE" AND UCASE(LEFT(rsTemp.Fields("TABLE_NAME").Value,4)) <> "MSYS"))  Then
					Response.Write("<OPTION VALUE=""" & server.htmlencode("""" & rsTemp.Fields("TABLE_SCHEMA").VALUE & """" & "." & """" & rsTemp.Fields("TABLE_NAME").VALUE & """") & """>")
					Response.Write(Server.HTMLEncode("""" & rsTemp.Fields("TABLE_SCHEMA").Value & """" & "." & """" & rsTemp.Fields("TABLE_NAME").Value & """"))
					Response.Write("</OPTION>")			
				End If
			Case "Oracle"
				If ((rsTemp.Fields("TABLE_TYPE").Value = "TABLE" AND UCASE(LEFT(rsTemp.Fields("TABLE_NAME").Value,4)) <> "MSYS"))  Then
					Response.Write("<OPTION VALUE=""" & server.htmlencode("""" & rsTemp.Fields("TABLE_SCHEMA").VALUE & """" & "." & """" & rsTemp.Fields("TABLE_NAME").VALUE & """") & """>")
					Response.Write(Server.HTMLEncode("""" & rsTemp.Fields("TABLE_SCHEMA").Value & """" & "." & """" & rsTemp.Fields("TABLE_NAME").Value & """"))
					Response.Write("</OPTION>")			
				End If
			Case ELse
				If ((rsTemp.Fields("TABLE_TYPE").Value = "TABLE" AND UCASE(LEFT(rsTemp.Fields("TABLE_NAME").Value,4)) <> "MSYS"))  Then
					Response.Write("<OPTION VALUE=""" & server.htmlencode(rsTemp.Fields("TABLE_NAME").VALUE) & """>")
					Response.Write(Server.HTMLEncode(rsTemp.Fields("TABLE_NAME").Value))
					Response.Write("</OPTION>")			
				End If
		End Select
		rsTemp.movenext
	Loop
	Response.Write("</SELECT><P>")
	Response.Write("<INPUT NAME=""SelectTables"" TYPE=Submit Value=""Next &gt;"" Class=submit>")
Else
	set rsTemp = server.createobject("ADODB.Recordset")
	set rsTemp2 = server.createobject("ADODB.Recordset")
	set rsImportSource = server.createobject("ADODB.Recordset")
	set rsImportTarget = server.createobject("ADODB.Recordset")
	'********************** Begin Show Field Match
	Response.Write("<SPAN CLASS=Information>Import Source : </SPAN>")
	Response.Write Server.HTMLEncode(Session("ocdSQLConnectImport"))
	Response.Write "<A HREF=WizardImport.asp?reconnect=yes>"
	Response.Write(  "(Click to Change)" & "</A> <BR> ")
	Response.Write "<SPAN CLASS=Information>Import From : </SPAN>  "
	Response.Write server.htmlencode(Request.QueryString("sqlfromimport_" & GRIDID)) 
	Response.Write "<A HREF=""WizardImport.asp"">"
	Response.Write( "(Click to Change)" & "</A> <BR>")
	Response.Write "<SPAN CLASS=Information>Target Table : </SPAN>  "
	Response.Write server.htmlencode(Request.QueryString("sqlfrom_" & GRIDID)) 
	Response.Write "<A HREF=""WizardImport.asp?sqlfromimport_" & GRIDID & "=" & server.urlencode(Request.QueryString("sqlfromimport_" & GRIDID)) & """>"
	Response.Write( "(Click to Change)" & "</A><P>")
	Response.Write("<FORM method=post action=""")
	Response.Write(Request.ServerVariables("SCRIPT_NAME"))
	Response.Write("?" & Request.QueryString & """>")
	Response.Write "<INPUT TYPE=Submit Name=Import Value=""Import Now !"" CLASS=Submit>"
	Response.Write "&nbsp;&nbsp;&nbsp;"
	Response.Write "<SELECT NAME=ImportTrans>"
	Response.Write "<OPTION VALUE="""" "
	If Request.Form("ImportTrans") = "" THen
		Response.Write " SELECTED "
	End If
	Response.Write ">Continue Import On Error</OPTION>"
	Response.Write "<OPTION VALUE=""FailOnError"" "
	If Request.Form("ImportTrans") = "FailOnError" THen
		Response.Write " SELECTED "
	End If
	Response.Write ">Stop Import On First Error</OPTION>"
	Response.Write "<OPTION VALUE=""TranContinueOnError"" "
	If Request.Form("ImportTrans") = "TranContinueOnError" THen
		Response.Write " SELECTED "
	End If
	Response.Write ">Continue Transaction On Error</OPTION>"
	Response.Write "<OPTION VALUE=""TranFailOnError"" "
	If Request.Form("ImportTrans") = "TranFailOnError" THen
		Response.Write " SELECTED "
	End If
	Response.Write ">Fail Transaction On Error</OPTION>"
	Response.Write "</SELECT>"
	Response.Write "&nbsp;&nbsp;&nbsp;"
	Response.Write "Ignore Field Errors <INPUT NAME=ImportIgnoreFIeldErrors TYPE=CHECKBOX "
	If Request.Form("ImportIgnoreFIeldErrors") <> "" THen
		Response.Write " CHECKED "
	End If
	Response.Write " TYPE=Text>"
	Response.Write "<P>"
	Response.flush
	If Request.Form("import") <> "" Then 
		Response.Write "<P><HR><P><SPAN CLASS=Information>Import Results</SPAN><P>"
		Response.Write "<P><B>Import Start : </B>"
		Response.Write now
		Response.Write "<P>"
		Select Case Request.Form("ImportTrans") 
			Case "TranFailOnError","TranContinueOnError"
				ocdTargetConn.BeginTrans
		End Select
		If ocdRunImportEventCode Then
			server.execute "WizardImport_StartTrans.asp"
		End If
		If UCase(ocdTargetConn.Provider) = "MSDORA" Then
			rsImportTarget.CursorLocation = 3
		End If
		call rsImportTarget.open ("SELECT * from " & (Request.QueryString("sqlfrom_" & GRIDID)) &  " WHERE 1=2", ocdTargetConn,3,3)
		If not Response.isclientconnected Then
			call closeobjects()
		End If
		call rsImportSource.open ("SELECT * from " & (Request.QueryString("sqlfromimport_" & GRIDID)) &  " ", ocdConnImport)
		If not Response.isclientconnected Then
			call closeobjects()
		End If
		If ocdRunImportEventCode Then
			server.execute "WizardImport_AfterLoad.asp"
		End If
		intICount = 0
		intSICount = 0
		blnFail = False
		do while not rsImportSource.eof and not blnFail
			If not Response.isclientconnected Then
				call closeobjects()
			End If
	 		intIcount = intICount + 1
			If intICOunt < 11 Then
				Response.Write intIcount & "-"
			ElseIf intICOunt > 10 and intIcount < 101 Then
				If intICount Mod 10 = 0 Then
					Response.Write intIcount & "-"
				End If
			Elseif intICount > 100 and intICount < 501 Then
				If intICount Mod 50 =0 Then
					Response.Write intIcount & "-"
				End If
			Elseif intICount > 500 and intICount < 1501 Then
				If intICount Mod 100 =0 Then
					Response.Write intIcount & "-"
				End If
			Else
				If intICount Mod 250 =0 Then
					Response.Write intIcount & "-"
				End If
			End If
			If ocdRunImportEventCode Then
				server.execute "WizardImport_BeforeInsert.asp"
			End If
			rsImportTarget.addnew
			For Each fvTemp2 in Request.Form
				If UCase(LEFT(fvTemp2,8)) = "OCDIMFLD"  and not blnFail Then
					If Cstr(Request.Form(fvTemp2)) = "" AND Cstr(Request.Form("ocdimexp" & mid(fvTemp2,9))) <> "" Then
						varEvalExpr = Cstr(Request.Form("ocdimexp" & mid(fvTemp2,9)))
						varEval = Eval(varEvalExpr)
						rsImportTarget(mid(fvTemp2,9)) = varEval 
					Elseif Cstr(Request.Form(fvTemp2)) <> "" Then
						varValue = rsImportSource(Cstr(Request.Form(fvTemp2)))
						rsImportTarget(mid(fvTemp2,9)) = varValue 
					End If
					If err.number = 0  THen
					Else
						If Request.Form("importtrans") <> "" and Request.Form("importignorefielderrors") = "" Then
							blnFail = True
						End If
						Response.Write "<P><SPAN CLASS=Warning><IMG SRC=appWarningSmall.gif ALT=Warning> Record " & intICOunt & " "
						Response.Write mid(fvTemp2,9)
						Response.Write "&lt;--"
						If Cstr(Request.Form(fvTemp2)) = "--Use Expression--" Then
							Response.Write varEvalExpr
							Response.Write " = "
							Response.Write varValue
							Response.Write "</SPAN> "
						Else
							Response.Write Request.Form(fvTemp2)
							Response.Write " = "
							Response.Write varValue
						End If
						Response.Write "</SPAN> "
						Response.Write " : "
						Response.Write err.number
						Response.Write " : "
						Response.Write err.description
						err.clear
						Response.Write "</P>"
					End If
				End If
			next	
			If not blnFail Then
				rsImportTarget.Update
				If err.number = 0  THen
					intSICOunt = intSICOunt + 1
				Else
					Response.Write "<P><span class=warning><IMG SRC=appWarningSmall.gif ALT=Warning> Record " & intICOunt & " Insert</SPAN> : "
					Response.Write err.number
					Response.Write " : " 
					Response.Write err.description
					Response.Write "</P>"
					Select Case Request.Form("ImportTrans") 
						Case "TranFailOnError","FailOnError"
							blnFail =True
					End Select
					rsImportTarget.CancelUpdate
					err.clear			
				End If
				Response.flush
			Else
				rsImportTarget.cancelUpdate
				err.clear
			End If
			If ocdRunImportEventCode Then
				server.execute "WizardImport_AfterInsert.asp"
			End If
			If not Response.isclientconnected Then
				call closeobjects()
			End If
			If not blnFail Then
				If intICount mod intBatchSize = 0 Then
					set rsImportTarget = Nothing
					set rsImportTarget = server.createobject("ADODB.Recordset")		
					call rsImportTarget.open ("SELECT * from " & (Request.QueryString("sqlfrom_" & GRIDID)) &  " WHERE 1=2", ocdTargetConn,3,3)
					If err.number <> 0 Then
						Response.Write err.description
						Response.end
					End If
				End If
				rsImportSource.movenext
			End If
		loop
		Response.Write "<P>" & intICOunt & " Record(s) processed</P>"
		Response.Write "<P>" & intSICOunt & " Successful Insert(s)</P>"
		If not Response.isclientconnected Then
			call closeobjects()
		End If
		Select Case Request.Form("ImportTrans") 
			Case "TranFailOnError"
				If blnFail Then
					ocdTargetConn.RollbackTrans
					Response.Write "<P>Transaction Rollback</P>"
				ELse
					ocdTargetConn.Committrans
					Response.Write "<P>Transaction Committed</P>"
				End If
			Case "TranContinueOnError"
				ocdTargetConn.Committrans
				Response.Write "<P>Transaction Committed</P>"
			Case "FailOnError"
				If blnFail Then
					Response.Write "<P>Import Stopped On Error</P>"
				End If
		End Select
		If err.number <> 0 Then
			Response.Write "<P><SPAN CLASS=Warning><IMG SRC=appWarningSmall.gif ALT=Warning> Transaction Error</SPAN> "
			Response.Write err.number
			Response.Write " : " 
			Response.Write err.description
			Response.Write "</P>"
			err.clear
		End If
		rsImportSource.Close
		set rsImportSource = Nothing
		rsImportTarget.close
		set rsImportTarget = Nothing
		err.clear
		Response.Write "<P><B>Import End : </B>"
		Response.Write now
		Response.Write "<P><HR><P>"
		If ocdRunImportEventCode Then
			server.execute "WizardImport_CompleteTrans.asp"
		End If
		If not Response.isclientconnected Then
			call closeobjects()
		End If
	End If
	If Request.Form("SaveSpec") <> "" Then
		Response.Write "<P><SPAN CLASS=Information>Save this Import Spec into a Text File: </SPAN>"
		Response.Write "<P><TABLE><TR><TD VALIGN=TOP><TEXTAREA  COLS=35 ROWS=5 Name=txt>"
		Response.Write Server.HTMLEncode ("**BEGIN 1 CLICK DB IMPORT" & vbCRLF)
		Response.Write Server.HTMLEncode ("Import Source=" &  Session("ocdSQLConnectImport") & vbCRLF)
		Response.Write Server.HTMLEncode ("Import User=" & session("ocdSQLUserImport") & vbCRLF)
		Response.Write Server.HTMLEncode ("Import Password=" &  session("ocdSQLPassImport") & vbCRLF)
		Response.Write Server.HTMLEncode ("Import From=" & Request.QueryString("sqlfromimport_" & GRIDID) & vbCRLF)
		Response.Write Server.HTMLEncode ("Target Table=" & Request.QueryString("sqlfrom_" & GRIDID) & vbCRLF)
		Response.Write Server.HTMLEncode ("Import Type=" & Request.Form("ImportTrans") & vbCRLF)
		Response.Write Server.HTMLEncode ("Ignore Field Errors=" & Request.Form("ImportIgnoreFIeldErrors") & vbCRLF)
		Response.Write Server.HTMLEncode ("Field Mapping=" & vbCRLF)
		For Each fvTemp in Request.Form
			If UCase(LEFT(fvTemp,8)) = "OCDIMFLD"  Then
				Response.Write Server.htmlencode(mid(fvTemp,9)) & " :: "
				Response.Write server.htmlencode(Request.Form(fvTemp))
				Response.Write " :: "
				Response.Write Server.HTMLEncode(Cstr(Request.Form("ocdimexp" & mid(fvTemp,9))) )& vbCRLF
			End If
		next
		Response.Write Server.HTMLEncode ("**END 1 CLICK DB IMPORT")
		Response.Write "</TEXTAREA><TD VALIGN=TOP><input class=submit type=button value=""Select All Code"" onClick=""javascript:this.form.txt.focus();this.form.txt.select();""><P>Use <B>Ctrl-C</B>, <B>Ctrl-V</B> for copy and paste. </TD></TR></TABLE></P>"
	End If
	Response.Write "<SPAN CLASS=Information>Import Field Mapping</SPAN> "
	Response.Write "<INPUT TYPE=Submit NAME=SaveSpec VALUE=""Generate Import Spec"" CLASS=Submit>"
	call rsTemp.open ("SELECT * from " & (Request.QueryString("sqlfrom_" & GRIDID)) &  " WHERE 1=2", ocdTargetConn)
'Response.Write "SELECT * from " & (Request.QueryString("sqlfromimport_" & GRIDID)) &  " WHERE 1=2"
	call rsTemp2.open ("SELECT * from " & (Request.QueryString("sqlfromimport_" & GRIDID)) &  " WHERE 1=2", ocdConnImport)
	If err <> 0 Then
		writefooter("")
	End If
	Response.Write "<TABLE><TR CLASS=GridHeader><TH ALIGN=LEFT>Destination</TH><TH></TH><TH ALIGN=LEFT>Import Source (use blank expresion to skip field)</TH></TR>"
	For Each rsField in rsTemp.Fields
		Response.Write "<TR CLASS=GridOdd>"
		Response.Write "<TD><SPAN CLASS=FieldName>"
		Response.Write rsField.Name
		Response.Write "</SPAN></TD>"
		Response.Write "<TD NOWRAP>&lt;---&nbsp;&nbsp;</TD>"
		Response.Write "<TD>"
		Response.Write "<SELECT NAME=""ocdImFld" & rsField.Name & """>"
		Response.Write "<OPTION VALUE="""" "
		If Request.Form = "" Then
			If Request.Form("ocdImFld" & rsField.Name) = "" Then
				Response.Write(" Selected ")
			End If
		Else
			If Request.Form("ocdImFld" & rsField.Name) = "" Then
				Response.Write(" Selected ")
			End If
		End If
		Response.Write ">--Use Expression--</OPTION>"
		For Each rsField2 in rsTemp2.Fields
			Response.Write "<OPTION VALUE=""" & rsField2.Name & """ "
			If Request.Form("import") = "" and Request.Form("SaveSpec") = "" AND Request.Form("loadImportSpec") = "" Then
				If rsField.Name = rsField2.Name Then
					Response.Write " SELECTED "
				End If

			Else
				If Request.Form("ocdImFld" & rsField.Name) = rsField2.Name Then
					Response.Write " SELECTED "
				End If
			End If
			Response.Write ">"
			Response.Write rsField2.Name
			Response.Write "</OPTION>"
		next
		Response.Write "</SELECT>"
		Response.Write "<INPUT NAME=""ocdImExp" & rsField.Name & """ VALUE="""
		Response.Write server.htmlencode(Request.Form("ocdImExp" & rsField.Name))
		Response.Write """>"
		Response.Write "</TD>"
		Response.Write "</TR>"
	next
	Response.Write "</TABLE>"
	rsTemp.close
	set rsTemp = Nothing
	rsTemp2.close
	set rsTemp2 = Nothing
End If
Response.Write("</FORM>")
Response.Write("</BLOCKQUOTE>")
'Response.Write Request.Form
Call WriteFooter("")
Response.end
Sub CloseObjects()
	on error resume next
	rsTemp.close
	rsTemp2.close
	rsImportSource.close
	rsImportTarget.close
	set rsTemp = Nothing
	set rsTemp2 = Nothing
	set rsImportSource = Nothing
	set rsImportTarget = Nothing
	set ocdTargetConn = Nothing
	set ocdConnImport = Nothing
	err.clear
	Response.end
end sub

%>
