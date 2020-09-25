<%@ LANGUAGE = VBScript.Encode %>
<% ' Except for @ Directives, there should be no ASP or HTML codes above this line

'1 Click DB technology is fully protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'IMPORTANT : THIS CODE USES PASS-THROUGH SECURITY  !
'To enforce application security, set logins and permissions
'for all web server and database users as appropriate.

'For more information see : http://1ClickDB.com

Option Explicit
On Error Resume Next
Response.Buffer=True   

%>

<!--#INCLUDE FILE=ocdFormat.asp-->
<!--#INCLUDE FILE=ocdForm.asp-->
<!--#INCLUDE FILE=ocdConnectInfo.asp-->
<!--#INCLUDE FILE=ocdFunctions.asp-->

<%

'Notes on #INCLUDE files:
'ocdFormat.asp contains writeheader("") and writefooter("") functions
'ocdConnectInfo.asp contains global database connection string variable ocdSQLConnect
'ocdForm.asp contains 1 Click DB data form object

'Initialize Form Object

Dim objForm
Set objForm = New ocdForm

objForm.SQLConnect = ocdSQLConnect 'ADO Connect String, including uid and pw if necessary
objForm.SQLUser = ocdSQLUser
objForm.SQLPass =ocdSQLPass

objForm.SQLSelect = "{{ocdSQLSelect}}" 'Database Field List 
objForm.SQLFrom = "{{ocdSQLFrom}}" 'Database Table Name
objForm.SQLWhereExtra = "{{ocdSQLWhereExtra}}"

'Fire All Events
objForm.CallBeforeDelete = True
objForm.CallAfterDelete = True
objForm.CallPreDelete = True
objForm.CallBeforeUpdate = True
objForm.CallAfterUpdate = True
objForm.CallBeforeInsert =True
objForm.CallAfterInsert = True
objForm.CallOnCancel = True

'Set Default Interface Behavior
objForm.AllowAdd = {{ocdAllowAdd}}
objForm.AllowEdit = {{ocdAllowEdit}}
objForm.AllowDelete = {{ocdAllowDelete}}
objForm.Open

Sub ocdBeforeUpdate ()

End Sub

Sub ocdAfterUpdate()

End Sub

Sub ocdBeforeInsert ()

End Sub

Sub ocdAfterInsert()

End Sub

Sub ocdBeforeDelete()
'Event is initiated when request("ocdEditDelete") <> "" AND request("ocdEditConfirm") = ""
   WriteHeader("")
   dim tmpeqs
   Response.write "<FORM ACTION=""" &    request.servervariables("SCRIPT_NAME") & "?"
   for each tmpeqs in request.querystring
      if UCase(tmpeqs) <> "OCDEDITDELETE" Then
         Response.write tmpeqs & "=" & Server.URLEncode(request.querystring(tmpeqs)) & "&amp;"
      end if
   next
   Response.write """ method=post>"
   Response.write "<TABLE><TR><TD VALIGN=TOP><IMG SRC=appWarning.gif ALT=Warning></td><TD VALIGN=TOP><b>Are you sure you want to delete this record?</b> <P>This action cannot be undone.<P><INPUT TYPE=Submit Name=ocdEditConfirm Value=OK Class=Submit>&nbsp;<INPUT TYPE=submit Name=ocdEditCancel Class=Submit Value=Cancel><INPUT TYPE=hidden Name=ocdEditDelete Value=Delete></td></tr></table></form>"
   Call WriteFooter("")
   Response.end
End Sub

Sub ocdPreDelete()

End Sub

Sub ocdAfterDelete()

End Sub

Sub ocdOnCancel()

End Sub

Call WriteHeader("")
%>
<!--INSERT CUSTOM HTML-->
<%objForm.Display("START")%>
<%objForm.Display("STATUS")%>
{{ocdFormFields}}
<%objForm.Display("BUTTONS")%>
<%objForm.Display("END")%>
<!--INSERT CUSTOM HTML-->
<%
objForm.Close()
Set objForm = nothing
%>
<!--INSERT CUSTOM HTML-->
<%
Call Writefooter("")

'There should be no ASP or HTML code below this line %>