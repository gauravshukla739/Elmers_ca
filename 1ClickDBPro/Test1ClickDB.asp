<%@ LANGUAGE = VBScript.Encode %>
<% ' Except for @ commands, there should be no ASP or HTML codes above this line
' Setting LANGUAGE = VBScript.Encode enables mixing of encoded and unencoded ASP scripts

'1 Click DB Pro copyright 1997-2004 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com/

        
'**Start Encode**

'If you are reading this text in your web browser through http:// 
'then this webserver directory is not configured to execute ASP scripts.

on error resume next
response.buffer = true
response.clear
Response.write "<h3>1 Click DB Compatibility Test</h3><P><HR><P>"
Response.write "Server Software: "
Response.write 	request.servervariables("server_software")
Response.write "<P><HR><P>"
Response.write "Testing Script Engine...<P>"
Response.flush
if CINT(scriptenginemajorversion()) < 5 Then
	response.write "<FONT COLOR=RED>1 Click DB requires Microsoft Scripting Engine version 5 or better.  This upgrade is available free from Microsoft at <A HREF=http://msdn.microsoft.com/scripting>http://msdn.microsoft.com/scripting</a>.  You may also be able to use an older version of 1 Click DB software.  Contact <A HREF=http://AccessHelp.net>AccessHelp.net</a> for details.</font>"
Else
	response.write "...Scripting Engine okay.<P>1 Click DB ASP files can be executed on this web server."
End if
Response.write "<P>"
Response.write "Current scripting engine is: "
Response.write (scriptengine())
Response.write " "
Response.write (scriptenginemajorversion())
Response.write (".")
Response.write (scriptengineminorversion())
response.write (" Build:")
Response.write (scriptenginebuildversion())
Response.flush
Response.write "<P><HR><P>Testing ADO Provider...<P>"
Response.flush
dim connTest
set connTest = server.createobject("ADODB.Connection")
If err.number <> 0 Then
	Response.write "1 Click DB Requires Microsoft Data Access Components version 2.1 or better.  These drivers are available free from Microsoft at <A HREF=http://microsoft.com/data>http://microsoft.com/data</a>"
Else
	Response.write "...ADO Provider found.<P>1 Click DB should be able to connect to databases from this web server."
	Response.write "<P>ADO Provider Version: "
	Response.write connTest.Version 
End if
Response.flush
Response.write "<P><HR><P>Testing Session State...<P>"
if err = 0 Then
	dim varRetVal
	varRetval = Session("DummyForTesting") 
	if err <> 0 Then
		Response.write "<FONT COLOR=RED>1 Click DB requires Session state to be enabled on your web server to run.  Note that ASP scripts <B>created</B> by 1 Click DB Code Wizard do not require Sessions to run.  1 Click DB Pro can also be configured to run with minimal privileges without using cookies, although these still must be enabled on the server.</FONT>"
		err.clear
	Else
		Response.write "...Session State okay."
		Response.write "<P>LCID (2048 is default): "
		response.write session.LCID
		Response.write "<P>Code Page: "
		response.write session.codepage
		err.clear
	End if  
end if


response.write "<P><HR><P>Testing File System Object...<P>"
Response.flush
dim fsoTest
set fso = server.createobject("Scripting.FileSystemObject")
if err <> 0 then
	Response.write "<FONT COLOR=RED>Could not create Scripting.FileSystemObject on this server.  It has either been uninstalled by your administrator or you do not have adequate permission to use it.  The 1 Click DB Code Wizard will not be able to create ASP files on this server.  However, if no other errors are reported you will be able to run files created by Code Wizard on this server.  Consider using another PWS or IIS web server for developing files for this web server with Code Wizard.</font>"
Else
	Response.write "...File System Object okay.<P>Assuming you have the correct permissions directories, you will be able to create ASP files with Code Wizard templates on this server."
End if
Response.write "<P><HR><P>Testing Language Output interface..."
Response.write "<P>Web Browser HTTP_ACCEPT_LANGUAGE: "
response.write request.servervariables("HTTP_ACCEPT_LANGUAGE")
Response.write "<P>Web Server Charset: "  
response.write response.charset
err.clear
Response.write "<P><HR><P>1 Click Db Web Server Compatibility Tests Complete. <P> Please email any error reports to <A HREF=mailto:support@1clickdb.com>support@1clickdb.com</a> or contact <A HREF=http://AccessHelp.net>AccessHelp.net</a> for support.<P>"
%>