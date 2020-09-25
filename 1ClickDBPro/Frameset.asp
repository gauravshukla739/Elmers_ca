<%@ LANGUAGE = VBScript.Encode %>
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
<%
Response.Write(vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF & vbCRLF)

'Response.Write ("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Frameset//EN"">") & vbCRLF
Response.Write ("<html><head><title>")
Response.Write(ocdBrandText)
Response.Write("</title></head>")
If ocdShowTopFrame Then
	Response.Write("<frameset bordercolor=""white"" rows=""55,*"">")
	Response.Write("<frame src=""MTMHEADING.ASP"" framespacing=""0"" name=""heading"" frameborder=""1"">")
End If
Response.Write("<frameset bordercolor=""white"" cols=""160,*""><frameset  bordercolor=""white"" rows=""*,0""><frame src=""MTMMENU_EMPTY.ASP"" name=""menu"" scrolling=""Auto"" frameborder=""0""><frame src=""MTMCODE.ASP"" name=""code"" scrolling=""no"" frameborder=""0""></frameset>")
If Request("contenturl") = "" Then
	Response.Write("<frame src=""Schema.asp?show=tables"" name=""text"" scrolling=""Auto"" frameborder=""1"">")
Else
	Response.Write("<frame src=""")
	Response.Write(server.htmlencode(request("contenturl")))
	Response.Write(""" name=""text"" scrolling=""Auto"" frameborder=""1"">")
End If
Response.Write("</frameset><noframes><body>Your browser does not support frames <a href=""Connect.asp"">Click here and choose No Frames compatibility to continue</a>.</body></noframes>")
If ocdShowTopFrame Then
	Response.Write ("</frameset>")
End If
Response.Write("</html>")
On Error Resume Next
ocdTargetConn.Close
Set ocdTargetConn = Nothing
%>
