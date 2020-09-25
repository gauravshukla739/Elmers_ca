<%

'If you are reading this text in your web browser, please check to see if your 
'webserver is properly configured to execute ASP scripts from this directory.

On Error Resume Next
Response.Buffer = True
Response.Clear()
%>
<!--#INCLUDE FILE=Config.asp-->
<%
If CInt(scriptenginemajorversion()) < 5 Then
	Response.Write("1 Click DB requires Microsoft Scripting Engine version 5.0 or better to be installed on your web server.  This upgrade is available free from Microsoft at <a href=""http://msdn.microsoft.com/scripting"">http://msdn.microsoft.com/scripting/</a>")
	Response.Write("<p>")
	Response.Write("Current scripting engine is: ")
	Response.Write(scriptengine())
	Response.Write(scriptenginemajorversion())
	Response.Write(".")
	Response.Write(scriptengineminorversion())
	Response.Write(" Build:")
	Response.Write(scriptenginebuildversion())
	Response.End()
End If
Set connTest = Server.CreateObject("ADODB.Connection")
If Err.Number <> 0 Then
	Response.Write("1 Click DB Requires Microsoft Data Access Components version 2.1 or better.  These drivers are available free from Microsoft at <a href=""http://microsoft.com/data/"">http://microsoft.com/data/</a>")
	Response.End()
Else
	Set connTest = Nothing
End If

If Err.Number = 0 Then
	If Not CBool(CInt(ocdCompatibility) And 16) Then '16=ocdNoCookies ; Session State is required
		Dim varRetVal
		varRetval = Session("ocdCompatibility") 
		If Err.Number <> 0 Then
			Response.Write("This copy of 1 Click DB is currently configured to use Session variables on this web server, but an attempt to use these variables has resulted an program error. The most likely cause for this is that Session variables have been disabled on this webserver by the webmaster.<p> To run 1 Click DB with no session variables add 16 to the current value of the ocdCompatibility variable in the Config.asp configuration file.  ocdCompatibility is a bitmasked variable and the number 16 stands for the internal constant ocdNoCookies.  No-session configuration also requires the value of the ocdADOConnection to be hardcoded in the same Config.asp configuration file.<p>Please note : the 1 Click DB Classic ASP libraries for Browse/Search/Export/Add/Edit/Delete DO NOT REQUIRE session state enabled on a webserver and by default Code Wizard apps use an global variable in a #INCLUDE file name ocdConnectInfo.asp.")
			Err.Clear()
			Response.End()
		End If  
	Else
		Response.Clear()
		Response.Redirect("Schema.asp")
		Response.End()
	End If
End If

If Not Request.ServerVariables("LOCAL_ADDR") = "207.21.247.253" Then 'Default redirection behavior
	If ocdLaunchPage = "" Then
		If ocdADOConnection = "" Then
			Response.Clear()
			Response.Redirect("Connect.asp")
			Response.End()
		Else
			Response.Clear()
			Response.Redirect("Schema.asp")
			Response.End()
		End if
	Else
		Response.Clear()
		Response.Redirect(ocdLaunchPage)
		Response.End()
	End if
Else 'this behavior for 1 Click DB home server is hardcoded
	Response.Clear()
	Response.Redirect("/view.aspx?_@id=534173")
	Response.End()
End if
%>