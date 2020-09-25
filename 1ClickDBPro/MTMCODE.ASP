<% ' Except for @ commands, there should be no ASP or HTML codes above this line
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
<%

Dim strMTMFooter, connSchema, intMCount, strMName, strSQLFrom, rsSchema, intI, arrSchema, strSQLFromName, strMTMenuText, blnShowObject, eleSchemaHideObjects, arrSchemaHideObjects
Dim arrSchemaFields (2) 

Call ConfigureTreeMenu()

%>

<html>
<head>
<title><%=ocdBrandText%></title>

<script type="text/javascript" src="mtmcode.js">
</script>

<script type="text/javascript">
<!--
// Morten's JavaScript Tree Menu
// version 2.3.0, dated 2001-04-30
// http://www.treemenu.com/

// Copyright (c) 2001, Morten Wang & contributors
// All rights reserved.

// This software is released under the BSD License which should accompany
// it in the file "MTMCOPYING.TXT".  If you do not have this file you can access
// the license through the WWW at http://www.treemenu.com/license.txt

/******************************************************************************
* User-configurable options.                                                  *
******************************************************************************/

// Menu table width, either a pixel-value (number) or a percentage value.
var MTMTableWidth = "100%";

// Name of the frame where the menu is to appear.
var MTMenuFrame = "menu";

// Variable for determining how a sub-menu gets a plus-sign.
// "Never" means it never gets a plus sign, "Always" means always,
// "Submenu" means when it contains another submenu.
var MTMSubsGetPlus = "Always";

// variable that defines whether the menu emulates the behaviour of
// Windows Explorer
var MTMEmulateWE = true;

// Directory of menu images/icons
var MTMenuImageDirectory = "";

// Variables for controlling colors in the menu document.
// Regular BODY atttributes as in HTML documents.
var MTMBGColor = "#FFFFFF";
var MTMBackground = "";
var MTMTextColor = "#000000";

// color for all menu items
var MTMLinkColor = "#300066";

// Hover color, when the mouse is over a menu link
var MTMAhoverColor = "#990000";

// Foreground color for the tracking & clicked submenu item
var MTMTrackColor ="#000000";
var MTMSubExpandColor = "#666699";
var MTMSubClosedColor = "#666699";

// All options regarding the root text and it's icon
var MTMRootIcon = "AppDB.gif";
var MTMenuText = "<a href=Schema.asp?show=ado target=text class=Menu>DB Properties</a>";
var MTMRootColor = "#000000";
var MTMRootFont = "Tahoma, Arial, sans-serif";
var MTMRootCSSize = "8pt";
var MTMRootFontSize = "-1";

// Font for menu items.
var MTMenuFont = "Tahoma, Arial, sans-serif";
var MTMenuCSSize = "8pt";
var MTMenuFontSize = "-1";

// Variables for style sheet usage
// 'true' means use a linked style sheet.
//var MTMLinkedSS = false;
//var MTMSSHREF = "style/menu.css";

var MTMLinkedSS = true;
var MTMSSHREF = "<%=ocdStyleSheet%>";

// Additional style sheet properties if you're not using a linked style sheet. 
// See the documentation for details on IDs, classes & elements used in the menu.
// Empty string if not used.
var MTMExtraCSS = "";

// Header & footer, these are plain HTML.
// Leave them to be "" if you're not using them

var MTMHeader = "";
var MTMFooter = "<%=strMTMFooter%>";

// Whether you want an open sub-menu to close automagically
// when another sub-menu is opened.  'true' means auto-close
var MTMSubsAutoClose = false;

// This variable controls how long it will take for the menu
// to appear if the tracking code in the content frame has
// failed to display the menu. Number if in tenths of a second
// (1/10) so 10 means "wait 1 second".
var MTMTimeOut = 15;

// Cookie usage.  First is use cookie (yes/no, true/false).
// Second is cookie name to use.
// Third is how many days we want the cookie to be stored.

var MTMUseCookies = false;
var MTMCookieName = "MTMCookie";
var MTMCookieDays = 3;

// Tool tips.  A true/false-value defining whether the support
// for tool tips should exist or not.
var MTMUseToolTips = true;

/******************************************************************************
* User-configurable list of icons.                                            *
******************************************************************************/

var MTMIconList = null;
MTMIconList = new IconList();

/******************************************************************************
* User-configurable menu.                                                     *
******************************************************************************/

var menu = null;
menu = new MTMenu();

MTMenuText = "<%=strMTMenuText%>";

// Main menu.

<%

Call BuildTreeMenu()
%>
-->
</script>
</head>
<body onload="MTMStartMenu()" bgcolor="#ffffff" text="#000000" link="#300066" vlink="#666699" alink="#990000">
</body>
</html>
<%

Sub ConfigureTreeMenu()

	If err.number <> 0 Then
		Response.Write err.description
		Response.Write err.number
		Response.End
	End If
	If ocdSchemaHideObjects <> "" Then
		arrSchemaHideObjects = split(ocdSchemaHideObjects,",")
	End If
	If Not (ocdStyleSheet <> "" And UCase(ocdMotif) = "") Then
		strMTMFooter = "<form target=\""_parent\"" action=\""Frameset.asp\"" method=\""post\""><p align=\""center\"">Style : <select onchange=\""document.forms[0].submit();\"" name=\""ocdMotif\""><option value=\""Classic\"" "
		Select Case UCase(ocdMotif)
			Case "","CLASSIC"
				strMTMFooter = strMTMFooter & " selected"
		End Select
		strMTMFooter = strMTMFooter & ">Classic</option><option value=\""Autumn\"" "
		if UCASE(ocdMotif) = "AUTUMN" Then
			strMTMFooter = strMTMFooter & " selected"
		End If
		strMTMFooter = strMTMFooter & ">Autumn</option><option value=\""Night\"" "
		if UCASE(ocdMotif) = "NIGHT" Then
			strMTMFooter = strMTMFooter & " selected"
		End If
		strMTMFooter = strMTMFooter & ">Night</option><option value=\""System\"" "
		if UCASE(ocdMotif) = "SYSTEM" Then
			strMTMFooter = strMTMFooter & " selected"
		End If
		strMTMFooter = strMTMFooter & ">None</option></select></p>"
	End If
	strMTMFooter = strMTMFooter & "<p align=\""center\""><a href=\""MTMCODE.ASP\"" target=\""code\"" class=\""Menu\"">Refresh TreeMenu</a></p></form>"
	if ocdIsHome Then
		'	strMTMFooter = strMTMFooter & "<a href=\""bugreport.htm\""  class=\""menu\"" target=\""_blank\""><img src=\""AppBug.gif\"" border=\""0\"" alt=\""Bugs\"">Report&nbsp;Bugs</a>"
	End If
	if ndnscSQLConnect = "" Then
		Response.Write "No Connection"
		Response.End
	End If
	set connSchema = server.CreateObject("ADODB.Connection")
	call connSchema.Open (ndnscSQLConnect , ndnscSQLUser , ndnscSQLPass)
	if err <>0 then
		Response.Write err.description '"Could not connect to database.  <A HREF=connect.asp target=_parent>Click here to continue.</a>"
		Response.End
	End If
	strMTMenuText = "<a href=Schema.asp?show=ado target=text class=menu>"
	Select Case ocdDatabaseType
		Case "ADSI"
			strMTMenuText = strMTMenuText & "ADSI"
		Case Else
			strMTMenuText = strMTMenuText & Server.HTMLEncode(getDatabaseDisplayName(ocdDatabaseType,ocdIsODBC,connSchema))
			response.write err.description
	End Select
	strMTMenuText = strMTMenuText  & "</a>"
	if ocdDatabaseType = "Oracle" Then
		set rsSchema = connSchema.execute ("SELECT OBJECT_TYPE AS TABLE_TYPE, OBJECT_NAME AS TABLE_NAME, OWNER AS TABLE_SCHEMA FROM ALL_OBJECTS WHERE (OBJECT_TYPE = 'TABLE' OR OBJECT_TYPE = 'VIEW') AND NOT OWNER = 'SYS' AND NOT OWNER = 'WKSYS' AND NOT OWNER = 'MDSYS' AND NOT OWNER = 'OLAPSYS' AND NOT OWNER ='CTXSYS' AND NOT OWNER='SYSTEM'")
		if err <>0 then 
			Response.clear
			Response.write err.description
			Response.Write "Could not get database info.  <A HREF=Connect.asp target=_parent>Click here to continue.</a>"	
			Response.end
		End If
		arrSchemaFields (0) = "TABLE_TYPE"
		arrSchemaFields (1) = "TABLE_NAME"
		arrSchemaFields (2) = "TABLE_SCHEMA"
		arrSchema = rsSchema.GetRows(,, arrSchemaFields )
		rsSchema.Close
		Set rsSchema = nothing
	ElseIf ocdDatabaseType <> "IXS" Then
		Set rsSchema = connSchema.OpenSchema(20) 	
		If rsSchema.eof then
			redim arrSchema(2,1)
		Else
			arrSchemaFields (0) = "TABLE_TYPE"
			arrSchemaFields (1) = "TABLE_NAME"
			arrSchemaFields (2) = "TABLE_SCHEMA"
			arrSchema = rsSchema.GetRows(,, arrSchemaFields )
		End If
		rsSchema.Close
		Set rsSchema = nothing
	End If
	connSchema.close
	Set connSchema = nothing
	ocdTargetConn.Close
	Set ocdTargetConn = Nothing
	if err <> 0 then 
	'	Response.clear
		response.write err.description
		Response.Write "Could not get database info.  <A HREF=Connect.asp target=_parent>Click here to continue.</a>"	
		Response.end
	End If	
End Sub

Sub BuildTreeMenu()
	if ocdDatabaseType = "ADSI" Then
		'no tree menu
	Elseif ocdDatabaseType = "IXS" Then
		Response.Write "menu.MTMAddItem(new MTMenuItem(""Scope"",""Browse.asp?sqlfrom_A=SCOPE()&sqlselect_A=doctitle,rank,write,characterization,dockeywords,filename,size,path,vpath"",""text"","""",""AppTable.gif""));" & vbCRLF
		Response.Write "menu.MTMAddItem(new MTMenuItem(""CGI Variables"", ""Schema.asp?show=cgi"", ""text""));" & vbCRLF
	Else
		Response.Write "menu.MTMAddItem(new MTMenuItem(""Tables"",""Schema.asp?show=tables"",""text"","""",""AppTable.gif""));" & vbCRLF
		Response.Write "menu.MTMAddItem(new MTMenuItem(""Views"",""Schema.asp?show=views"",""text"","""",""AppTable.gif""));" & vbCRLF
		if ocdAllowProAdmin and not ocdReadOnly Then
			If ocdDatabaseType <> "Oracle" and ocdDatabaseType <> "MySQL" Then
				Response.Write "menu.MTMAddItem(new MTMenuItem(""Procedures"", ""Schema.asp?show=procs"", ""text"","""",""AppProc.gif""));" & vbCRLF
			End If
		End If
		if ocdShowImport Then
			Response.Write  vbCRLF & "menu.MTMAddItem(new MTMenuItem(""Import Wizard"", ""WizardImport.asp"", ""text"","""",""AppImport.gif""));" & vbCRLF
		End If
		if ocdAllowCodeWiz and not ocdReadOnly Then
			Response.Write  vbCRLF & "menu.MTMAddItem(new MTMenuItem(""Code Wizard"", ""WizardASPCode.asp"", ""text"","""",""AppWizard.gif""));" & vbCRLF
		End If
		if ocdAllowCodeWiz and not ocdReadOnly and ocdDatabaseType = "SQLServer" Then
			Response.Write  vbCRLF & "menu.MTMAddItem(new MTMenuItem(""Audit Wizard"", ""WizardSQLAudit.asp"", ""text"","""",""AppAudit.gif""));" & vbCRLF
	End If

	Response.Write vbCRLF & "var tablelist = null;" & vbCRLF
	Response.Write vbCRLF & "tablelist = new MTMenu();" & vbCRLF
	If ocdDataBaseType = "SQLServer" or ocdDatabaseType = "Oracle" Then
		for intI = 0 to UBound( arrSchema ,2)
			if arrSchema (0, intI ) = "TABLE" and UCASE(left(arrSchema(1,intI),4)) <> "MSYS" then
				strSQLFrom = ocdQuotePrefix & arrSchema(2,intI) & ocdQUoteSuffix & "." & ocdQUotePrefix & arrSchema(1,intI) & ocdQuoteSuffix
blnShowObject = True
If ocdSchemaHideObjects <> "" Then

					For Each eleSchemaHideObjects In arrSchemaHideObjects

						If eleSchemaHideObjects = strSQLFrom Then
							blnShowObject = False

							
						End If
					Next
			Else
				blnShowObject = True
			End If

				strSQLFromName = arrSchema(1,intI)
if blnShowObject Then
				Response.Write (vbCRLF & "tablelist.MTMAddItem(new MTMenuItem(""" & server.htmlencode(strSQLFromName) & """,""Browse.asp?sqlfrom_A=" & server.urlencode(strSQLFrom) & """,""text""));" & vbCRLF)
End If
 			End If
		next
	Else
		for intI = 0 to UBound( arrSchema ,2)
			if arrSchema (0, intI ) = "TABLE" and UCASE(left(arrSchema(1,intI),4)) <> "MSYS" then
				Response.Write (vbCRLF & "tablelist.MTMAddItem(new MTMenuItem(""" & server.htmlencode( arrSchema(1,intI) ) & """,""Browse.asp?sqlfrom_A=" & server.urlencode(ocdQuotePrefix & arrSchema(1,intI) & ocdQUoteSuffix) & """,""text""));" & vbCRLF) 
			End If
		next
	End If
	If ocdDatabaseType <> "Oracle" And Not ocdReadOnly Then
		If ocdDatabaseType = "Access" And ocdAccessTableEdits Then
			Response.Write (vbCRLF & "tablelist.MTMAddItem(new MTMenuItem(""New Table"",""DBDesignMSAccess.asp?ocdAction=NEWTABLE"",""text"","""",""AppNew.gif""));" & vbCRLF)
		ElseIf ocdDatabaseType = "SQLServer" And ocdSQLTableEdits And Not ocdIsODBC Then
			Response.Write (vbCRLF & "tablelist.MTMAddItem(new MTMenuItem(""New Table"",""DBDesignSQLServer.asp?action=newtable"",""text"","""",""AppNew.gif""));" & vbCRLF)
		End If
	End If
	Response.Write vbCRLF & "menu.items[0].MTMakeSubmenu(tablelist,"""",""AppTable.gif"",""AppTable.gif"");" & vbCRLF
	strMName = ""
	intMCount = 0
	if ocdDatabaseType = "SQLServer" Then
		for intI = 0 to UBound( arrSchema ,2)
			if arrSchema (0, intI ) = "TABLE" and UCASE(left(arrSchema(1,intI),4)) <> "MSYS" then
strMName = safename(arrSchema(1,intI))
				strSQLFrom = ocdQuotePrefix & arrSchema(2,intI) & ocdQUoteSuffix & "." & ocdQUotePrefix & arrSchema(1,intI) & ocdQuoteSuffix
			blnShowObject = True

			If ocdSchemaHideObjects <> "" Then

					For Each eleSchemaHideObjects In arrSchemaHideObjects

						If eleSchemaHideObjects = strSQLFrom Then
							blnShowObject = False

							
						End If
					Next
			Else
				blnShowObject = True
			End If

			If blnShowObject Then

				Response.Write vbCRLF

				Response.Write "var tablelist_" & strMName & " = null;" & vbCRLF
				Response.Write "tablelist_" & strMName & " = new MTMenu();" & vbCRLF
				Response.Write (vbCRLF & "tablelist_" & strMName & ".MTMAddItem(new MTMenuItem(""Search"",""Browse.asp?ocdGridMode_A=Search&sqlfrom_A=" & server.urlencode(strSQLFrom) & """,""text"","""",""AppSearch.gif""));" & vbCRLF)
				Response.Write (vbCRLF & "tablelist_" & strMName & ".MTMAddItem(new MTMenuItem(""Properties"",""Structure.asp?sqlfrom=" & server.urlencode(strSQLFrom) & """,""text""));" & vbCRLF)
				If Not ocdReadOnly Then
					If ocdUseCustomEditPages Then
						Response.Write (vbCRLF & "tablelist_" & strMName & ".MTMAddItem(new MTMenuItem(""New Record"",""" & Replace(Replace(arrSchema(1,intI)," ","_"),"/","_") & "_edit.asp"",""text"","""",""AppNew.gif""));" & vbCRLF)
					Else
						Response.Write (vbCRLF & "tablelist_" & strMName & ".MTMAddItem(new MTMenuItem(""New Record"",""Edit.asp?sqlfrom=" & server.urlencode(strSQLFrom) & """,""text"","""",""AppNew.gif""));" & vbCRLF)
					End If
				End If
				Response.Write "tablelist.items[" & intMCount & "].MTMakeSubmenu(tablelist_" & strMName & ","""",""AppTable.gif"",""AppTable.gif"");" & vbCRLF
				intMCount = intMCount + 1
			End If
			End If
		Next
	Else
		For intI = 0 To UBound( arrSchema ,2)
			If arrSchema (0, intI ) = "TABLE" and UCASE(left(arrSchema(1,intI),4)) <> "MSYS" Then
strMName = safename(arrSchema(1,intI))
			blnShowObject = True

			

			If blnShowObject Then

				Response.Write vbCRLF
				Response.Write "var tablelist_" & strMName & " = null;" & vbCRLF
				Response.Write "tablelist_" & strMName & " = new MTMenu();" & vbCRLF
				Response.Write (vbCRLF & "tablelist_" & strMName & ".MTMAddItem(new MTMenuItem(""Search"",""Browse.asp?ocdGridMode_A=Search&sqlfrom_A=" & server.urlencode(ocdQuotePrefix & arrSchema(1,intI) & ocdQuoteSuffix) & """,""text"","""",""AppSearch.gif""));" & vbCRLF)
				Response.Write (vbCRLF & "tablelist_" & strMName & ".MTMAddItem(new MTMenuItem(""Properties"",""Structure.asp?sqlfrom=" & server.urlencode(ocdQuotePrefix & arrSchema(1,intI) & ocdQuoteSuffix) & """,""text""));" & vbCRLF)
				if not ocdReadOnly Then
					if ocdUseCustomEditPages Then
						Response.Write (vbCRLF & "tablelist_" & strMName & ".MTMAddItem(new MTMenuItem(""New Record"",""" & Replace(Replace(arrSchema(1,intI)," ","_"),"/","_") & "_edit.asp"",""text"","""",""AppNew.gif""));" & vbCRLF)
					Else
						Response.Write (vbCRLF & "tablelist_" & strMName & ".MTMAddItem(new MTMenuItem(""New Record"",""Edit.asp?sqlfrom=" & server.urlencode(ocdQuotePrefix & arrSchema(1,intI) & ocdQuoteSuffix) & """,""text"","""",""AppNew.gif""));" & vbCRLF)
					End If
				End If
				Response.Write "tablelist.items[" & intMCount & "].MTMakeSubmenu(tablelist_" & strMName & ","""",""AppTable.gif"",""AppTable.gif"");" & vbCRLF
				intMCount = intMCount + 1
			End If
			End If
		next
	End If
	Response.Write vbCRLF & "var viewlist = null;" & vbCRLF
	Response.Write "viewlist = new MTMenu();" & vbCRLF
	If ocdDataBaseType = "SQLServer" Then
		for intI = 0 to UBound( arrSchema ,2)
			if arrSchema (0, intI ) = "VIEW" and UCASE(left(arrSchema(1,intI),4)) <> "MSYS" then
				strSQLFrom = ocdQuotePrefix & arrSchema(2,intI) & ocdQUoteSuffix & "." & ocdQUotePrefix & arrSchema(1,intI) & ocdQuoteSuffix
				if ocdShowObjectOwner Then
					strSQLFromName = strSQLFrom
				Else
					strSQLFromName = arrSchema(1,intI)
				End If
				Response.Write (vbCRLF & "viewlist.MTMAddItem(new MTMenuItem(""" & arrSchema(1,intI) & """,""Browse.asp?sqlfrom_A=" & server.urlencode(strSQLFrom) & """,""text""));" & vbCRLF) 
			End If
		next
	Else
		for intI = 0 to UBound( arrSchema ,2)
			if arrSchema (0, intI ) = "VIEW" and UCASE(left(arrSchema(1,intI),4)) <> "MSYS" then
				Response.Write (vbCRLF & "viewlist.MTMAddItem(new MTMenuItem(""" & server.htmlencode(arrSchema(1,intI)) & """,""Browse.asp?sqlfrom_A=" & server.urlencode(arrSchema(1,intI)) & """,""text""));" & vbCRLF) 
			End If
		next
	End If
	If not ocdDatabaseType = "Oracle" and not ocdReadOnly Then
		If ocdDatabaseType = "Access" and ocdAccessTableEdits Then
			if ocdShowQueryWizard and ocdQueryWizardIsDefault Then
				Response.Write (vbCRLF & "viewlist.MTMAddItem(new MTMenuItem(""New View"",""WizardSQLQuery.asp"",""text"","""",""AppNew.gif""));" & vbCRLF)
			Else
				Response.Write (vbCRLF & "viewlist.MTMAddItem(new MTMenuItem(""New View"",""DBDesignMSAccess.asp?ocdAction=CREATEVIEW"",""text"","""",""AppNew.gif""));" & vbCRLF)
			End If
		Elseif ocdDatabaseType = "SQLServer" Then 'and ocdSQLTableEdits and not ocdIsODBC Then
			If ocdShowQueryWizard and ocdQueryWizardIsDefault Then
				Response.Write (vbCRLF & "viewlist.MTMAddItem(new MTMenuItem(""New View"",""WizardSQLQuery.asp"",""text"","""",""AppNew.gif""));" & vbCRLF)
			Else
				Response.Write (vbCRLF & "viewlist.MTMAddItem(new MTMenuItem(""New View"",""DBDesignSQLServer.asp?Action=createview"",""text"","""",""AppNew.gif""));" & vbCRLF)
			End If
		End If
	End If
	Response.Write vbCRLF & "menu.items[1].MTMakeSubmenu(viewlist,"""",""AppView.gif"",""AppView.gif"");" & vbCRLF
	strMName = ""
	intMCount = 0
	If ocdDatabaseType = "SQLServer" Then
		For intI = 0 to UBound( arrSchema ,2)
			If arrSchema (0, intI ) = "VIEW" then
				strSQLFrom = ocdQuotePrefix & arrSchema(2,intI) & ocdQUoteSuffix & "." & ocdQUotePrefix & arrSchema(1,intI) & ocdQuoteSuffix
				strMName = safename(arrSchema(1,intI))
				Response.Write vbCRLF
				Response.Write "var viewlist_" & strMName & " = null;" & vbCRLF
				Response.Write "viewlist_" & strMName & " = new MTMenu();" & vbCRLF
				Response.Write (vbCRLF & "viewlist_" & strMName & ".MTMAddItem(new MTMenuItem(""Search"",""Browse.asp?ocdGridMode_A=Search&sqlfrom_A=" & server.urlencode(strSQLFrom) & """,""text"","""",""AppSearch.gif""));" & vbCRLF)
				If ocdAllowProAdmin Then
					Response.Write (vbCRLF & "viewlist_" & strMName & ".MTMAddItem(new MTMenuItem(""Properties"",""Structure.asp?sqlfrom=" & server.urlencode(strSQLFrom) & """,""text""));" & vbCRLF)
				End If
				Response.Write "viewlist.items[" & intMCount & "].MTMakeSubmenu(viewlist_" & strMName & ","""",""AppView.gif"",""AppView.gif"");" & vbCRLF
				intMCount = intMCount + 1
			End If
		Next
	Else
		For intI = 0 to UBound( arrSchema ,2)
			if arrSchema (0, intI ) = "VIEW" and UCASE(left(arrSchema(1,intI),4)) <> "MSYS" then
				strMName = safename(arrSchema(1,intI))
				Response.Write vbCRLF
				Response.Write "var viewlist_" & strMName & " = null;" & vbCRLF
				Response.Write "viewlist_" & strMName & " = new MTMenu();" & vbCRLF
				Response.Write (vbCRLF & "viewlist_" & strMName & ".MTMAddItem(new MTMenuItem(""Search"",""Browse.asp?ocdGridMode_A=Search&sqlfrom_A=" & server.urlencode(arrSchema(1,intI)) & """,""text"","""",""AppSearch.gif""));" & vbCRLF)
				if ocdAllowProAdmin Then
					Response.Write (vbCRLF & "viewlist_" & strMName & ".MTMAddItem(new MTMenuItem(""Properties"",""Structure.asp?sqlfrom=" & server.urlencode(arrSchema(1,intI)) & """,""text""));" & vbCRLF)
				End If
				Response.Write "viewlist.items[" & intMCount & "].MTMakeSubmenu(viewlist_" & strMName & ","""",""AppView.gif"",""AppView.gif"");" & vbCRLF
				intMCount = intMCount + 1
			End If
		next
	End If
End If	


End Sub
%>
