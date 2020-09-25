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
Dim arrPasswords, elePassword
If ocdAdminPassword = ""  Then
	Response.Clear
	Response.Redirect(ocdConnectURL)
	Response.End
End If
arrPasswords = Split(ocdAdminPassword,",")

	Select Case UCase(Request("Action"))
		Case "LOGON", ""
			If Request.Form <> "" Then
				For Each elePassword In arrPasswords
					If Request.Form("AdminPassword") = elePassword then
						Session("ocdAdminAuthorized") = "True"
						Response.Clear
						Response.Redirect(ocdConnectURL)
						Response.End
					End If
				Next
			End If
		Case "LOGOUT"
			Session("ocdAdminAuthorized") = ""
			Session("ocdSQLConnect") = ""
			Session("ocdSQLUser") = ""
			Session("ocdSQLPass") = ""
			Session("ocdCompatibility") = ""
			Response.Clear
			If ocdLaunchPage = "" Then
				Response.Redirect(ocdPageName)
			Else
				Response.Redirect(ocdLaunchPage)
			End If
			Response.End
	End Select
Call WriteHeader("")
Session("ocdAppInit") = True
Response.Write("<CENTER><FORM method=post action=""")
Response.Write(Request.ServerVariables("SCRIPT_NAME"))
Response.Write("?")
Response.Write(Server.URLEncode(Request.QueryString("datasource")))
Response.Write("""")
Response.Write(DrawDialogBox("DIALOG_START","Please Logon",""))
Response.Write("<TABLE CELLSPACING=3 CELLPADDING=3>")
If Request.Form <> "" Then
	Response.Write("<TR><TD COLSPAN=2><P><SPAN CLASS=Warning><IMG SRC=appWarningSmall.gif ALT=Warning> Incorrect Password</SPAN></P></TD></TR>")
End If
Response.Write("<TR><TD>")
Response.Write("<SPAN CLASS=FieldName>Authorization:</TD><TD>")
If ocdReadOnly Then
	Response.Write("Read")
ElseIf ocdAllowProAdmin Then
	Response.Write("Admin")
Else
	Response.Write("Edit")
End If
Response.Write("</TD></TR>")
Response.Write("<TR><TD ALIGN=LEFT VALIGN=TOP><b>Password:</b></TD><TD ALIGN=LEFT VALIGN=BOTTOM><input type=Password name=adminpassword size=35 maxlength=255 VALUE=""")
If Request("pass") <> "" then
	Response.Write(Server.HTMLEncode(Request("pass")))
End If
Response.Write("""></TD></TR>")
Response.Write("<TR><TD colspan=2 VALIGN=TOP><P><input type=submit name=Action CLASS=Submit Value=""Logon"">")
if Session("ocdAdminAuthorized") = "True" Then
	Response.Write("<input type=submit  CLASS=Submit name=Action Value=""Logout"">")
End If
Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
If UCase(Request.ServerVariables("HTTPS")) = "ON" Then
	Response.Write("<A HREF=""http://" )
	Response.Write(Request.ServerVariables("SERVER_NAME"))
	Response.Write(Request.ServerVariables("SCRIPT_NAME"))
	Response.Write("?") 
	Response.Write(Server.URLEncode(Request.QueryString))
	Response.Write(""">Exit https secure</a>")
Else
	Response.Write("<A HREF=""https://")
	Response.Write(Request.ServerVariables("SERVER_NAME"))
	Response.Write(Request.ServerVariables("SCRIPT_NAME"))
	Response.Write("?")
	Response.Write(Server.URLEncode(Request.QueryString))
	Response.Write(""">Make https secure</a>")
End If
Response.Write("</td></tr></table></TD></TR></TABLE></form></CENTER>")
Call WriteFooter("")
%>
