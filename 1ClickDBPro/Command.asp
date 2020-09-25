<% ' Except for @ Directives, there should be no ASP or HTML codes above this line
' Setting LANGUAGE = VBScript.Encode enables mixing of encoded And unencoded ASP scripts

'1 Click DB Pro copyright 1997-2004 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com/

'Email inquiries to info@1ClickDB.com
        
'**Start Encode**

%>
<!--#INCLUDE FILE=PageInit.asp-->
<!--#INCLUDE FILE=ocdFunctions.asp-->
<!--#INCLUDE FILE=ocdCommand.asp-->
<!--#INCLUDE FILE=ocdShowPlan.asp-->
<!--#INCLUDE FILE=ocdLoadBinary.asp-->
<!--#INCLUDE FILE=ocdManageSQLServer.asp-->
<%

'Command.asp is used to execute arbitrary SQL Batches from a FORM post.
'SQL may be typed in manually executed from uploaded text files
'Graphic Display of Estimated Execution SHOWPLAN is supported
Dim strSQLText, qsTemp, strShowPlan, strTimeout, objLoader, strFileData, objSQL, strSPHelp, strMDBName, lngTimeout, rsTemp, strCommand, objCommand, intBatchCount, strCheckedCommand, cmdTemp, intBatchCurrent, arrCommands, objSP
Call WriteHeader("")
If Not ocdShowSQLCommander Or Not ocdAllowProAdmin then
	Call WriteFooter("Admin Not Enabled")
	Response.End()
End If
If Request.QueryString("loadit") <> "" Then
	Set objLoader = new Loadit
	strFileData = objLoader.GetFileInput()
	If objLoader.TotalBytes = 0 Then
		Set objLoader = Nothing
		Call WriteFooter("A Zero Length File Was Uploaded")
	ElseIf strFileData = "" Then
		Set objLoader = Nothing
		Call WriteFooter("Uploaded File Is Too Big")
	End If
End If

If Request.QueryString("loadit") = "" Then
	If request("EstExPlan") <> "" Then
		strShowPlan = "text"
	End If
	strSQLText =  Request("sqltext")
	strTimeout = request("ocdTimeOut")
Else
	strTimeout = Request.QueryString("ocdTimeOut")
End If
If Request.QueryString("EstExPlan") <> "" Then
	strShowPlan = "text"
End If
Response.Flush()
Response.Write("<span class=""information"">Execute SQL Commands:</span>")
If ocdDatabaseType = "Oracle" And ocdOraSQLReference <> "" Then
	Response.Write(" <a href=""" & ocdOraSQLReference & """ target=""_blank"" class=""menu""><img src=""appHelpSmall.gif"" border=""0"" alt=""help"">PL/SQL </a><br>")
ElseIf ocdDatabaseType = "SQLServer" And ocdMSSQLReference <> "" Then
	Response.Write(" <a href=""" & ocdMSSQLReference & """ target=""_blank"" class=""menu""><img src=""appHelpSmall.gif"" border=""0"" alt=""help"">Transact-SQL</a><br>")
ElseIf ocdJETSQLReference <> "" And ocdDatabaseType = "Access" Then
	Response.Write(" <a href=""" & ocdJETSQLReference & """ target=""_blank"" class=""menu""><img src=""appHelpSmall.gif"" border=""0"" alt=""help"">JET SQL</a><br>")
End If
Response.Write("<form action=""" & ocdPageName & "?nocache=" & Server.URLEncode(Now()))
For Each qsTemp in Request.QueryString
	Select Case UCase(qsTemp) 
		Case "SQLTEXT", "SHOWPLAN", "PROPOSEDSQLTEXT" ,"LOADIT" 
			If Request.QueryString(qsTemp) <> "" Then
				Response.Write("&amp;" & qsTemp & "=" & Server.URLEncode(Request.QueryString(qsTemp)))
			End If
	End Select
Next
Response.Write(""" method=""post""><textarea rows=""6"" cols=""50"" name=""sqltext"">" )
If strFileData <> "" Then
	Response.Write(Replace(Server.HTMLEncode((strFileData)),"&#65279;",""))
Elseif strSQLText <> "" Then
	Response.Write(Server.HTMLEncode(strSQLText))
ElseIf Request.QueryString("editsp") <> "" and ocdDatabaseType = "SQLServer" Then
	Set objSQL = New ocdManageSQLServer
	objSQL.SQLConnect = ndnscSQLConnect
	objSQL.SQLUser = ndnscSQLUser
	objSQL.SQLPass = ndnscSQLPass
	objSQL.Open
	strSPHelp = Trim(objSQL.GetHelpText( Request.QueryString("editsp")))
	Do Until UCase(left(strSPHelp, 1)) = "C"
		strSPHelp = mid(strSPHelp,2)
	Loop
	strSPHelp = "ALTER " & Mid(strSPHelp, 8)
	Response.Write Server.HTMLEncode(strSPHelp)
	set objSQL = Nothing
Else
	Response.Write(Server.HTMLEncode(Request.QueryString("proposedsqltext")))
End If
Response.Write("</textarea>")
If Not CBool(ndnscCompatibility And ocdNoJavaScript) Then
	Response.Write("&nbsp;<a href="""" onClick=""javascript:window.open('ocdZoomText.asp?CallingForm=" & "0" & "&amp;TextField=sqltext', 'zoomtext','height=400,width=600,scrollbars=yes');return false""><img src=""GRIDLNKEDIT.GIF"" border=""0"" alt=""Zoom""></a>")
End If
Response.Write("<br>")
Select Case ocdDatabaseType
	Case "Access","SQLServer"
		Response.Write " Timeout: <select name=""ocdTimeout"">"
		Response.Write "<option value=""0"""
		If strTimeout = "0" Then
			Response.Write " selected"
		End If
		Response.Write ">Never</option>"
		Response.Write "<option value=1 "
		If strTimeout = "1" Then
			Response.Write " selected"
		End If
		Response.Write ">1 Second</option>"
		If ocdServerScriptTimeout > 2 Then
			Response.Write "<option value=""3"" "
			If strTimeout = "3" Then
				Response.Write " selected"
			End If
			Response.Write ">3 Seconds</option>"
			If ocdServerScriptTimeout > 4 Then
				Response.Write "<option value=5 "
				If strTimeout = "5" or strTimeout = "" Then
					Response.Write " selected"
				End If
				Response.Write ">5 Seconds</option>"
				Response.Write "<option value=10 "
				If ocdServerScriptTimeout > 9 Then
					If strTimeout = "10" Then
						Response.Write " selected"
					End If
					Response.Write ">10 Seconds</option>"
					If ocdServerScriptTimeout > 19 Then
						Response.Write "<option value=20 "
						If strTimeout = "20" Then
							Response.Write " selected"
						End If
						Response.Write ">20 Seconds</option>"
						If ocdServerScriptTimeout > 29 Then
							Response.Write "<option value=30 "
							If strTimeout = "30" Then
								Response.Write " selected"
							End If
							Response.Write ">30 Seconds</option>"
							If ocdServerScriptTimeOut > 179 Then
								Response.Write "<option value=60 "
								If strTimeout = "60" Then
									Response.Write " selected"
								End If
								Response.Write ">1 Minute</option>"
								If ocdServerScriptTimeOut > 179 Then
									Response.Write "<option value=180 "
									If strTimeout = "180" Then
										Response.Write " selected"
									End If
									Response.Write ">3 Minutes</option>"
									If ocdServerScriptTimeOut > 299 Then
										Response.Write "<option value=300 "
										If strTimeout = "300" Then
											Response.Write " selected"
										End If
										Response.Write ">5 Minutes</option>"
										If ocdServerScriptTimeOut > 599 Then
											Response.Write "<option value=600 "
											If strTimeout = "600" Then
												Response.Write " selected"
											End If
											Response.Write ">10 Minutes</option>"
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		Response.Write "</select>&nbsp;&nbsp; "
End Select
If ocdDatabaseType = "SQLServer" Then
	If not ocdDBMSVersion < CDbl(7) Then
		Response.Write "Show Plan : "
		Response.Write "<input type=""checkbox"" name=""EstExPlan"""
			If strShowPlan = "text" Then
				Response.Write " checked"
			End If
		Response.Write ">"
	End If
End If
Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;<input type=""submit"" value=""Go!"" class=""submit"">")
Response.Write("</form>")
If UCase(ocdDatabaseType) = "SQLSERVER" Then 
	Response.Write("<form method=""POST"" enctype=""multipart/form-data"" action=""" & ocdPageName & "?loadit=true""><input type=""file"" name=""file"" size=""25""><input type=""submit"" value=""Load .SQL File""></form>")
End If
Response.Flush 
If strTimeout = "" Then
	lngTimeout = 5
Else
	lngTimeout = CLng(strTimeout)
End If
If Not strSQLText = "" Then
	Response.Write("<span class=""information"">")
	If strShowPlan = "text" Then
		Response.Write "Estimated Execution Plan:"
	Else
		Response.Write("Results:")
	End If
	Response.Write("</span><p>")
	If strShowPlan = "text" Then
		Response.Write "Hover mouse over plan icons for detailed statistics <a href=""" & ocdShowPlanReference & """ target=""_blank"" class=""menu""><img border=""0"" src=""appHelpSmall.gif"" alt=""Help""> Show Plan</a><p>"
	End If
	Select Case ocdDatabaseType
		Case "SQLServer","Access"
			ocdTargetConn.CommandTimeout = lngtimeout
	End Select
	If ocdIsHome And ocdDatabaseType = "Access" Then
	strMDBName = Cstr(ocdTargetConn.Properties("Data Source Name"))
		If IsNumeric(mid(strMDBName,instrrev(strMDBName,"\")+1, len(strMDBName)- (3+(instrrev(strMDBName,"\")+1)))) Then
		Else
			Call WriteFooter("Not Enabled for this Example")
			Response.end
		End If
		Err.clear
	End If 
	Set rsTemp = Server.Createobject("ADODB.Recordset") 
	Select Case strShowPlan
		Case "all", "text"
			ocdTargetConn.Execute "SET SHOWPLAN_ALL ON"
		Case "statsprof"
			ocdTargetConn.Execute "SET STATISTICS PROFILE ON"
		Case "time","timeonly"
			ocdTargetConn.Execute "SET STATISTICS TIME ON"
	End Select
	arrCommands = Split(strSQLText,vbCRLF & "GO",-1,1)
	intBatchCurrent = 1
	intBatchCount = UBound(arrCommands)
	For Each strCommand In arrCommands
		strCheckedCommand = replace(strCommand, vbCRLF, " ")
		strCheckedCommand = Trim(strCheckedCommand)
		If InStr(1,strCheckedCommand,"set quoted_identifier",1) > 0 And strShowplan = "text" Then
			strCheckedCommand = ""
			Response.Write "<p><hr>"
			Response.Write "<span class=""information"">Query Batch " & intBatchCurrent & "</span> "   & Server.HTMLEncode(strCommand) & "<p>"
			intBatchCurrent = intBatchCurrent + 1
		End If
		If strCheckedCommand  <> "" Then
			Set cmdTemp = server.createobject("ADODB.Command")
			cmdTemp.CommandTimeout = strTimeout
			cmdTemp.CommandText = strCommand
				Response.Write "<p><hr>"
				Response.Write "<span class=""information"">Query Batch " & intBatchCurrent & "</span> "   & Server.HTMLEncode(Now()) & " --&gt; " & Server.HTMLEncode(strCommand) & "<p>" 
			intBatchCurrent = intBatchCurrent + 1
			Set cmdTemp.ActiveConnection = ocdTargetConn
			If UCase(strShowPlan) = "TEXT" THen
				Set objSP = New ocdShowPlan
				Set objSP.ADOCommand = cmdTemp
				objSP.Display(strShowPlan)
				If Err.Number <> 0 Then
					Response.Write("<p><span class=""warning""><img src=""appWarningSmall.gif""> ")
					Response.Write(Err.Description)
					Response.Write("</span><br>Further results from this query batch could not be retrieved by ADO</p>")
					Err.Clear()
				End If
				If ubound(arrCommands) > 1 Then
					Response.Write "<hr>"
					Response.Flush
				End If
				Set cmdTemp = Nothing
				Set objSP = Nothing
				Err.Clear()
			Else
				Set objCommand = New ocdCommand
				objCommand.ShowSortLinks = False
				Set objCommand.ADOCommand = cmdTemp
				objCommand.Display(strShowPlan)
				If Err.Number <> 0 Then
					Response.Write("<p><span class=""warning""><img src=appWarningSmall.gif> ")
					Response.Write(Err.Description)
					Response.Write "</span><br>Further results from this query batch could not be retrieved by ADO</p>"
					Err.Clear()
				End If
				If UBound(arrCommands) > 1 Then
					Response.Write "<hr>"
					Response.flush()
				End If
				Set cmdTemp = Nothing
				Set objCommand = Nothing
				Err.Clear()
			End If
		End If
	Next
	Select Case strShowPlan
	 	Case "text"
			ocdTargetConn.Execute "SET SHOWPLAN_TEXT OFF"
		Case "all"
			ocdTargetConn.Execute "SET SHOWPLAN_ALL OFF"
		Case "statsprof"
			ocdTargetConn.Execute "SET STATISTICS PROFILE OFF"
		Case "time","timeonly"
			ocdTargetConn.Execute "SET STATISTICS TIME OFF"
	End Select		
	Response.Write("<p>Complete</p>")
	ocdTargetConn.Close
	Set ocdTargetConn = Nothing
End If
Response.Write("<p>")
Call WriteFooter("")
%>  