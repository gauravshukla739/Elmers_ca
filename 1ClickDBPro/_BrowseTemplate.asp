<%@ LANGUAGE = VBScript.Encode %>
<% ' Except for @ Directives, there should be no ASP or HTML codes above this line
' Setting LANGUAGE = VBScript.Encode enables mixing of encoded ASP scripts with plain text commands


'1 Click DB technology is fully protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'IMPORTANT : THIS CODE USES PASS-THROUGH SECURITY  !
'To enforce application security, set logins and permissions
'for all web server and database users as appropriate.

'Page Settings
Option Explicit
On Error Resume Next
Response.Buffer = True	

%>

<!--#INCLUDE FILE=ocdFormat.asp-->
<!--#INCLUDE FILE=ocdGrid.asp-->
<!--#INCLUDE FILE=ocdFunctions.asp-->
<!--#INCLUDE FILE=ocdConnectInfo.asp-->

<%

'Notes on #INCLUDE files:
'ocdFormat.asp contains writeheader("") and writefooter("") functions 
'ocdConnectInfo.asp contains global database connection string variable ocdSQLConnect
'ocdGrid.asp contains 1 Click DB data grid object

'Initialize Grid Object
Dim objGrid 
set objGrid = New ocdGrid

'Set Database Information
objGrid.SQLConnect = ocdSQLConnect 'string variable from ocdConnectInfo.asp
objGrid.SQLUser = ocdSQLUser
objGrid.SQLPass =ocdSQLPass

objGrid.SQLSelect = "{{ocdSQLSelect}}" 'comma delimited list of field names, if blank all fields will be returned
objGrid.SQLFrom = "{{ocdSQLFrom}}" 'table name or sql join
objGrid.SQLOrderByDefault = "{{ocdSQLOrderByDefault}}" 'default column sort
objGrid.SQLWhereExtra = "{{ocdSQLWhereExtra}}" 'extra where clause restriction

'Set Default Interface Behavior
objGrid.AllowAdd = {{ocdAllowAdd}}
objGrid.AllowEdit = {{ocdAllowEdit}}
objGrid.AllowDelete = {{ocdAllowDelete}}
objGrid.AllowDetail = {{ocdAllowDetail}}

'.EditForm is also used for Add and Delete targets
'Form links only active when corresponding .Allow is true
'The specified files must already have been created with from an Edit template exist or a 404 error will result

{{ocdForms}}

'If using for searches or exports, write no html before opening grid.  objGrid.GridMode is now either ="Search" ="Browse" or ="Filter"

'Writeheader Displays standard formatting for top of the page

'Process grid information and retrieve data, this will fire a page redirect for print and exports and after processing filter and search form criteria.
objGrid.Open
'Place custom HTML or ASP Code as indicated:
%>
	<!-- Insert HTML for All Modes -->
<%
'Display individual grid elements depending on grid mode
'Valid Modes are SEARCH, BROWSE, FILTER, and EXPORT
Select Case UCase(objGrid.GridMode)
	Case "EXPORT"
		objGrid.Display("GRID")
		response.end
	Case "FILTER" 'Single Field Search
		Call WriteHeader("")
%>
	<!-- Insert HTML for FILTER Mode -->
	
	<SPAN CLASS=Information>Set Filter</SPAN>

	<%objGrid.Display("Filter")%>
	
	<!-- End Insert HTML for FILTER Mode -->
<%	
	Case "SEARCH" 'Query by Form search page for all or specified fields
		Call WriteHeader("")
%>
	<!-- Insert HTML for SEARCH Mode -->

	<SPAN CLASS=Information>Set Criteria</SPAN>

	<%objGrid.Display("Search")%>

	<!-- End Insert HTML for SEARCH Mode -->
<%
	Case "BROWSE" 'HTML table view of active fields and records
		Call WriteHeader("")
		objGrid.HTMLAttribGridEven = "class=""GridEven"" onmouseover=""this.className='GridSelect';"" onmouseout=""this.className='GridEven';"" "
		objGrid.HTMLAttribGridOdd = "class=""GridOdd"" onmouseover=""this.className='GridSelect';"" onmouseout=""this.className='GridOdd';"" "
		objGrid.HTMLGridButtons = "first|<IMG SRC=""GRIDSMBTNFIRST.GIF"" ALT=""First"" BORDER=""0"" WIDTH=""18"" HEIGHT=""18"">;;prev|<IMG SRC=""GRIDSMBTNPREV.GIF"" ALT=""Previous"" BORDER=""0"" WIDTH=""14"" HEIGHT=""18"">;;paging|smbutton;;next|<IMG SRC=""GRIDSMBTNNEXT.GIF"" BORDER=""0"" ALT=""Next"" WIDTH=""14"" HEIGHT=""18"">;;last|<IMG SRC=""GRIDSMBTNLAST.GIF"" BORDER=""0"" ALT=""Last"" WIDTH=""18"" HEIGHT=""18"">;;new|<IMG SRC=""GRIDSMBTNNEW.GIF"" ALT=""New"" BORDER=""0"" WIDTH=""18"" HEIGHT=""18"">;;search|<IMG SRC=""GRIDSMBTNSEARCH.GIF"" ALT=""Search Records"" BORDER=""0"" WIDTH=""18"" HEIGHT=""18"">;;drilldown|<IMG SRC=""GRIDSMBTNDRILLDOWN.GIF"" ALT=""Drill Down"" BORDER=""0"" WIDTH=""18"" HEIGHT=""18"">;;reset|<IMG SRC=""GRIDSMBTNRESET.GIF"" ALT=""Reset"" BORDER=""0"" WIDTH=""18"" HEIGHT=""18"">;;print|<IMG SRC=""GRIDSMBTNPRINT.GIF"" WIDTH=18 HEIGHT=18 BORDER=""0"" ALT=""Print"">;;excel|<IMG SRC=""GRIDSMBTNEXCEL.GIF"" BORDER=""0"" ALT=""Export to Excel"" WIDTH=""18"" HEIGHT=""18"">;;xml|<IMG SRC=""GRIDSMBTNXML.GIF"" BORDER=""0"" ALT=""Export to XML"" WIDTH=""18"" HEIGHT=""18"">" 

		objGrid.HTMLEditLink = "<IMG SRC=""GRIDLNKEDIT.GIF"" border=""0"" HEIGHT=""12"" WIDTH=""12"" ALT=""Edit"">"
		objGrid.HTMLDeleteLink = "<IMG SRC=""GRIDLNKDELETE.GIF"" border=0 HEIGHT=12 WIDTH=12 ALT=""Delete"">"
		objGrid.HTMLDetailLink = "<IMG SRC=""GRIDLNKDETAIL.GIF"" border=0 HEIGHT=12 WIDTH=12 ALT=""Detail"">"
		objGrid.HTMLSortASCLink = "<IMG SRC=""GRIDLNKASC.GIF"" BORDER=""0"" ALT=""Sort Ascending"" WIDTH=""11"" HEIGHT=""11"">"
		objGrid.HTMLSortDESCLink = "<IMG SRC=""GRIDLNKDESC.GIF"" BORDER=""0"" ALT=""Sort Descending"" WIDTH=""11"" HEIGHT=""11"">"
		objGrid.HTMLFilterLink = "<IMG SRC=""GRIDLNKFILTER.GIF"" BORDER=""0"" ALT=""Filter on This Field"" WIDTH=""11"" HEIGHT=""11"">"

		objGrid.HTMLTrueValue = "<img src=""GRIDVALTRUE.GIF"" border=""0"" alt=""True"">"
		objGrid.HTMLFalseValue = "<img src=""GRIDVALFALSE.GIF"" border=""0"" alt=""False"">"

%>
	<!-- Insert HTML for BROWSE Mode -->	

	<%objGrid.Display("Buttons")%>

	<!-- End Insert HTML for BROWSE Mode  -->	

	<%
	'Display Main Data Table
	call objGrid.Display("Grid")
	%>

	<!-- Insert HTML for BROWSE Mode  -->
	
	<%
	'display keyword search box below Grid
	call objGrid.Display("Keyword")
	%>

	<!-- Insert HTML for BROWSE Mode  -->
<%
	'show links to remove grid sort and/or criteria
	'this code block can be safely deleted
	dim QS
	If request.querystring("sqlwhere_" & objGrid.GRIDID) <> "" Then
		Response.write ("<B>Criteria:</b> <SPAN CLASS=SQLText>" & Server.HTMLEncode(Request.Querystring("sqlwhere_" &objGrid.GridID)) & "</SPAN>&nbsp;&nbsp;<A HREF=""" & Request.Servervariables("SCRIPT_NAME")   & "?sqlwhere_" & objGrid.GridID & "=")
		for each QS in Request.Querystring
			If  UCASE(QS) <> "NDACTION_" & objGrid. GridID AND UCASE(QS) <> "SQLWHERE_" & objGrid.GridID Then
				Response.write  ("&amp;" & QS  & "=" & Server.URLEncode(Request.Querystring(QS)))
			End if
		next
		Response.write ("""><IMG SRC=""GRIDLNKNOFILTER.GIF"" BORDER=0 WIDTH=12 HEIGHT=12 ALT=""Remove Criteria""></A><P>")
	end if
	If Request.Querystring("sqlorderby_" & objGrid.GridID) <> "" Then
		Response.write ("<B>Order&nbsp;By:</b> <SPAN CLASS=SQLText>" & Server.HTMLEncode(Request.Querystring("sqlorderby" & "_" & objGrid.GridID)) & "</SPAN>&nbsp;&nbsp;<A HREF=""" & Request.Servervariables("SCRIPT_NAME")  & "?sqlorderby" & "_" & objGrid.GridID & "=")
		for each QS in Request.Querystring
			If  UCASE(QS) <> "NDACTION" & "_" & objGrid.GridID AND UCASE(QS) <> "SQLORDERBY" & "_" & objGrid.GridID Then
				Response.write ("&amp;" & QS  & "=" & Server.URLEncode(Request.Querystring(QS)))
			End if
		next
		Response.write ("""><IMG SRC=""GRIDLNKNOSORT.GIF"" WIDTH=12 HEIGHT=12 border=0 ALT=""Remove Order By""></A><P>")
	end if
	'end links to remove grid sort and/or criteria
%>
	<!-- Insert HTML for BROWSE Mode  -->
<%
End Select
%>
	<!-- Insert HTML for All Modes  -->
<%
Call WriteFooter("")
'end Full Control custom asp or HTML

'clean up 1 Click DB VBScript Grid object
objGrid.Close
set objGrid = Nothing

'There should be no ASP or HTML code below this line%>
