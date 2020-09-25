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
if not ocdShowSQLSelector then
	response.end
End if
if (ocdDatabaseType = "Access" or ocdDatabaseType = "SQLServer") and Not CBool(CInt(ndnscCompatibility) and ocdNoJavaScript) and Not CBool(CInt(ndnscCompatibility) and ocdNoCookies) Then
	'response.redirect ("WizardSQLQuery.asp")
end if
if request.querystring("ocdStartQueryWizard") <> "" Then
	response.redirect ("WizardSQLQuery.asp")
End if
dim ndsQS
Call Writeheader("")

Response.Write ("<SPAN Class=Information>Edit SQL Select</SPAN>")

Response.write ("<P><FORM action=""Browse.asp"" method=get><TABLE><TR><TD VALIGN=TOP ALIGN=RIGHT><B>SELECT</B> </TD><TD VALIGN=TOP ALIGN=LEFT><TEXTAREA cols=40 rows=3 name=sqlselect_A>")
if Request.QueryString("sqlselect_A") = "" then
	Response.Write ("*")
Else
	Response.Write (Request.QueryString("sqlselect_A"))
End if
Response.Write ("</TEXTAREA></TD></TR><TR><TD VALIGN=TOP ALIGN=RIGHT><B>FROM</B> </TD><TD VALIGN=TOP ALIGN=LEFT><TEXTAREA name=sqlfrom_A cols=40 rows=2>")
Response.write (Server.HTMLEncode( Request.QueryString("sqlfrom_A") ) )
Response.Write ("</TEXTAREA></TD></TR><TR><TD VALIGN=TOP ALIGN=RIGHT><B>WHERE</B> </TD><TD VALIGN=TOP ALIGN=LEFT><TEXTAREA cols=40 rows=3 id=textarea1 name=sqlwhere_A>")
Response.Write (Server.HTMLEncode(Request.QueryString("sqlwhere_A")))
Response.Write ("</TEXTAREA></TD></TR><TR><TD VALIGN=TOP ALIGN=RIGHT><B>GROUP BY</B> </TD><TD VALIGN=TOP ALIGN=LEFT><INPUT name=sqlgroupby_A size=40  VALUE=""")
Response.Write (server.HTMLEncode ( Request.QueryString("sqlgroupby_A") ))
Response.Write ("""></TD></TR><TR><TD VALIGN=TOP ALIGN=RIGHT><B>HAVING</B> </TD><TD VALIGN=TOP ALIGN=LEFT><INPUT name=sqlhaving_A size=40 VALUE=""")
Response.Write (server.HTMLEncode ( Request.QueryString("sqlhaving_A") ))
Response.Write ("""></TD></TR><TR><TD VALIGN=TOP ALIGN=RIGHT><B>ORDER BY</B> </TD><TD VALIGN=TOP ALIGN=LEFT><INPUT name=sqlorderby_A size=40 VALUE=""")
Response.Write (server.HTMLEncode ( Request.QueryString("sqlorderby_A") ))
Response.Write ("""></TD></TR><TR><TD VALIGN=TOP ALIGN=LEFT><B>PAGE SIZE</b></td><TD><INPUT type=text size=4 Name=SQLPageSize_A VALUE=""")
if request.querystring("SQLPageSize_A") <> "" Then
	Response.write (Server.HTMLEncode(Request.Querystring("SQLPageSize_A")))
Else
	Response.write ("10")
End if
Response.write ("""> ")
Response.write (" <INPUT type=submit CLASS=Submit VALUE=""Browse""> ")
Response.write ("</td></tr></TABLE>")

for each ndsQS in request.querystring
	Select Case UCASE(ndsQS)
		Case "SQLSELECT_A", "SQLFROM_A", "SQLWHERE_A", "SQLORDERBY_A", "SQLGROUPBY_A", "SQLHAVING_A", "SQLPAGE_A", "SQLID_A", "DATABASETYPE","GRIDID", "FILTERFIELDTYPE","FILTERFIELDSIZE", "SQLPAGESIZE_A","NDSEARCHRECORDS","FILTERFIELD","SCRIPT"
		Case Else
			Response.write ("<INPUT TYPE=HIDDEN NAME=""" )
			Response.write (ndsQS )
			Response.write (""" VALUE=""" )
			Response.write (server.htmlencode(request.querystring(ndsQS)))
			Response.write (""">")
	End Select
next
Response.write ("</FORM>")
call writefooter("")
%>
