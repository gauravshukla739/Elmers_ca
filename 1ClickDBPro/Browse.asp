<%@ LANGUAGE = VBScript.Encode %>
<%' Except for @ Directives, there should be no ASP or HTML codes above this line
' Setting LANGUAGE = VBScript.Encode enables mixing of encoded and unencoded ASP scripts

'1 Click DB Pro copyright 1997-2004 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com/

'**Start Encode**

'--------------------
'Begin INCLUDE
'--------------------
%>
<!--#INCLUDE FILE=PageInit.asp-->
<!--#INCLUDE FILE=ocdFunctions.asp-->
<!--#INCLUDE FILE=ocdGrid.asp-->
<%
'--------------------
'End INCLUDE
'--------------------

'--------------------
'Begin Page_Load
'--------------------

Dim objGrid

'--------------------
'End Page_Load
'--------------------

'--------------------
'Begin Page_Render
'--------------------

Call WriteHeader("")

Call DisplayBrowse()

Call WriteFooter("")

Response.End()

'--------------------
'End Page_Render 
'--------------------

'--------------------
'Begin Procedures
'--------------------

Sub DisplayBrowse()
	Dim strSQLV, qsTemp, objGrid, arrCEP, eleCEP, blnCEP

	Set objGrid = New ocdGrid

	objGrid.GridID = "A"
	objGrid.ADOMaxRecords = ocdMaxRecordsRetrieve
	objGrid.GridMaxRecordsDisplay = ocdMaxRecordsDisplay
	objGrid.Debug = ocdDebug
	Set objGrid.ADOConnection = ocdTargetConn 
	objGrid.AllowMultiDelete = True
	objGrid.GridHideAutonumber = ocdHideAutonumber
	objGrid.ADORecordsetTimeout = ocdDBTimeout
	objGrid.ADOComputeTimeout  = ocdComputeTimeout
	objGrid.SQLSelect = Request.QueryString("sqlselect_" & objGrid.GRIDID)
	objGrid.SQLSelectName = ""
	objGrid.ExportLineBreaks = ocdExportLineBreaks

	objGrid.SQLFrom =  Request.QueryString("sqlfrom_" & objGrid.GridID)
	objGrid.SQLGroupBy = Request.QueryString("sqlgroupby_" & objGrid.GridID)
	objGrid.SQLHaving = Request.QueryString("sqlhaving_" & objGrid.GridID)
	objGrid.UseRegExKeywordSearch = ocdUseRegExKeywordSearch
	If Not ocdReadOnly Then
		objGrid.AllowDelete = True
		objGrid.AllowAdd = True
		objGrid.AllowEdit = True
	Else
		objGrid.AllowDelete = False
		objGrid.AllowAdd = False
		objGrid.AllowEdit = False
	End If
	objGrid.AllowMultiSelect = False
	objGrid.FormSelect = ""
	objGrid.SQLSelectSum = Request.QueryString("SQLSelectSum_" & objGrid.GridID)
	objGrid.SQLSelectMin = Request.QueryString("SQLSelectMin_" & objGrid.GridID)
	objGrid.SQLSelectMax = Request.QueryString("SQLSelectMax_" & objGrid.GridID)
	objGrid.SQLSelectAvg = Request.QueryString("SQLSelectAvg_" & objGrid.GridID)
	objGrid.SQLGroupBy = Request.QueryString("sqlgroupby_" & objGrid.GridID)
	objGrid.SQLHaving = Request.QueryString("sqlhaving_" & objGrid.GridID)
	Select Case UCase(ocdMotif)
		Case "SYSTEM"
			objGrid.HTMLGridButtons = "paging|;;search|<span class=""ocdGridBtn"">Search</span>;;drilldown|<span class=""ocdGridBtn"">DrillDown</span>;;reset|<span class=""ocdGridBtn"">Reset</span>;;first|<span class=""ocdGridBtn"">First</span>;;prev|<span class=""ocdGridBtn"">Prev</span>;;next|<span class=""ocdGridBtn"">Next</span>;;last|<span class=""ocdGridBtn"">Last</span>;;new|<span class=""ocdGridBtn"">New</span>;;print|<span class=""ocdGridBtn"">Print</span>;;excel|<span class=""ocdGridBtn"">Excel</span>;;xml|<span class=""ocdGridBtn"">XML</span>"
		Case Else
			If UCase(objGrid.GridMode) = "BROWSE" Then
				objGrid.RenderAsHTML = ocdRenderAsHTML
				If ocdGridIcons Then
					objGrid.ShowDescription = ocdShowDescription
					objGrid.HTMLGridButtons = "first|<img src=""GRIDSMBTNFIRST.GIF"" alt=""First Page"" border=""0"" width=""18"" height=""18"">;;prev|<img src=""GRIDSMBTNPREV.GIF"" alt=""Previous Page"" border=""0"" width=""14"" height=""18"">;;paging|smbutton;;next|<img src=""GRIDSMBTNNEXT.GIF"" border=""0"" alt=""Next Page"" width=""14"" height=""18"">;;last|<img src=""GRIDSMBTNLAST.GIF"" border=""0"" alt=""Last Page"" width=""18"" height=""18"">;;new|<img src=""GRIDSMBTNNEW.GIF"" alt=""New Record"" border=""0"" width=""18"" height=""18"">;;search|<img src=""GRIDSMBTNSEARCH.GIF"" alt=""Advanced Search"" border=""0"" width=""18"" height=""18"">;;drilldown|<img src=""GRIDSMBTNDRILLDOWN.GIF"" alt=""Drill Down"" border=""0"" width=""18"" height=""18"">;;reset|<img src=""GRIDSMBTNRESET.GIF"" alt=""Show All"" border=""0"" width=""18"" height=""18"">;;print|<img src=""GRIDSMBTNPRINT.GIF"" width=18 height=18 border=""0"" alt=""Print All"">;;excel|<img src=""GRIDSMBTNEXCEL.GIF"" border=""0"" alt=""Export All to Excel"" width=""18"" height=""18"">;;xml|<img src=""GRIDSMBTNXML.GIF"" border=""0"" alt=""Export All to XML"" width=""18"" height=""18"">" 'custom|<input TYPE=""Image"" src=""GridBtnSave.gif"" onclick=""javascript:window.external.AddFavorite(location.href,document.title)"" alt=""Save"">;;
					objGrid.HTMLEditLink = "<img src=""GRIDLNKEDIT.GIF"" border=""0"" height=""14"" width=""14"" alt=""Edit"">"
					objGrid.HTMLDeleteLink = "<img src=""GRIDLNKDELETE.GIF"" border=""0"" height=""14"" width=""14"" alt=""Delete"">"
					objGrid.HTMLDetailLink = "<img src=""GRIDLNKDETAIL.GIF"" border=""0"" height=""12"" width=""12"" alt=""Detail"">"
					
					objGrid.HTMLTrueValue = "<img src=""GridValTrue.gif"" border=""0"" alt=""True"">"
					objGrid.HTMLFalseValue = "<img src=""GridValFalse.gif"" border=""0"" alt=""False"">"
				End If
				objGrid.HTMLSortASCLink = "<img src=""GRIDLNKASC.GIF"" border=""0"" alt=""Sort Ascending"" width=""11"" height=""11"">"
				objGrid.HTMLSortDESCLink = "<img src=""GRIDLNKDESC.GIF"" border=""0"" alt=""Sort Descending"" width=""11"" height=""11"">"
				objGrid.HTMLFilterLink = "<img src=""GRIDLNKFILTER.GIF"" border=""0"" alt=""Filter This Field"" width=""11"" height=""11"">"
				If ocdGridHighlightSelected And (Not CBool(CInt(ndnscCompatibility) And ocdNoJavaScript)) Then
					objGrid.HTMLAttribGridEven = ("class=""GridEven"" onmouseover=""javascript:this.className='GridSelect';"" onmouseout=""javascript:this.className='GridEven';""")
					objGrid.HTMLAttribGridOdd = ("class=""GridOdd"" onmouseover=""javascript:this.className='GridSelect';"" onmouseout=""javascript:this.className='GridOdd';""")
				Else
					objGrid.HTMLAttribGridEven = ("class=""GridEven""")
					objGrid.HTMLAttribGridOdd = ("class=""GridOdd""")
				End If
			End If
	End Select
	objGrid.SearchMultiSort = ocdMultipleFieldSort
	objGrid.SearchDefaultTextCompare = ocdDefaultTextCompare
	objGrid.SearchCheckAll = ocdShowCheckedSearchFields
	If ocdWrapGrid Then
		objGrid.HTMLAttribGridCell = "valign=""top"""
	Else
		objGrid.HTMLAttribGridCell = "valign=""top"" nowrap "
	End If
	objGrid.ExportForceDownload = ocdForceExportDownload
	objGrid.SQLPageSizeDefault = ocdPageSizeDefault
	blnCEP = False
	Select Case ocdCustomEditPages
		Case "*"
			blnCEP = True
		Case ""
			blnCEP = False
		Case Else
			arrCEP = Split(ocdCustomEditPages,",")
			For Each eleCEP In arrCEP
				If eleCEP = objGrid.SQLFrom Then
					blnCEP = True
					Exit For
				End If
			Next
	End Select
	If blnCEP Then
		objGrid.FormEdit = Replace(Replace(Replace(Replace(Replace(Replace(objGrid.SQLFrom," ","_"),"/","_"),"""",""),".","_"),"[",""),"]","") & "_Edit.asp"
	Else
		objGrid.FormEdit = "Edit.asp"
	End If

	objGrid.Open

	If CInt(objGrid.SQLPageSize) = 1 Then
		objGrid.HTMLGridVertical = True
	End If
	If Err.Number <> 0 then 
		If Not (UCase(Request.QueryString("ocdGRIDMODE_A")) <> "PROCESS" and UCase(Request.QueryString("ocdGRIDMODE_A")) <> "EXPORT") Then
			Call WriteHeader("")
		End If
		Call WriteFooter("")
	End If
	Select Case UCase(objGrid.GridMode)
		Case "EXPORT"
			If UCase(Request.QueryString("ocdExportFormat_" & objGRID.GridID)) = "PRINT" Then
				objGrid.HTMLExportStart = "<html><head><title>1 Click DB Export</title><link rel=""stylesheet"" type=""text/css"" href=""ocdStyleSheetExport.css""></head><body onload=""javascript:window.print();"">"
			Else
				objGrid.HTMLExportStart = "<html><head><title>1 Click DB Export</title><body>"
			End If
			objGrid.HTMLExportEnd = "</body></html>"
			objGrid.Display("GRID")
			Response.End()
		Case "SEARCH" '
			Response.Write("<span class=""Information"">Search in ")
			Response.Write(Server.HTMLEncode(objGrid.SQLFrom))
			Response.Write("</span>")
			objGrid.Display("SEARCH")
		Case "BROWSE" 'Table View of Selected Records
			If ocdShowSQLText Then			
				If (ocdDatabaseType = "SQLServer" And ocdSQLTableEdits) Or (ocdDatabaseType = "Access" And ocdAccessTableEdits) Then
					If ocdAccessTableEdits Then
						Response.Write("<form method=""post"" action=""DBDesignMSAccess.asp?ocdaction=createview"">")
					Else
						Response.Write("<form method=""post"" action=""DBDesignSQLServer.asp?action=createview"">")
					End If
	'				Response.Write("<input type=""hidden"" name=""doaction"" value=""yes"">")
					Response.Write("<input type=""hidden"" name=""SQLCommandText"" value=""")
					If ocdDatabaseType = "SQLServer" Then
						strSQLV = "SELECT TOP " & objGrid.ADOMaxRecords & " " & objGrid.SQLSelect
					Else
						strSQLV = "SELECT " & objGrid.SQLSelect
					End If
					strSQLV = strSQLV & " FROM "  
					strSQLV = strSQLV & objGrid.SQLFrom
					If (Not objGrid.SQLWhere = "") Then
						strSQLV = strSQLV & " WHERE " & objGrid.SQLWhere & ""
					End If
					If Not objGrid.SQLGroupBy = "" then 
						strSQLV = strSQLV & " GROUP BY " & objGrid.SQLGroupBy 
					End If
					If Not objGrid.SQLHaving = "" then 
						strSQLV = strSQLV & " HAVING " & objGrid.SQLHaving 
					End If
					If Not objGrid.SQLOrderBy = "" then 
						strSQLV = strSQLV & " ORDER BY " & objGrid.SQLOrderBy 
					End If
					Response.Write(Server.HTMLEncode(strSQLV))
					Response.Write(""">")
				End If
				Response.Write("<a href=""Select.asp?sqlfrom_A=")
				Response.Write(Server.URLEncode(objGrid.SQLFrom))
				For Each qsTemp In Request.QueryString
					If UCase(qsTemp) <> "SQLFROM_A" Then
						Response.Write("&amp;" & qsTemp & "=" & server.urlencode(Request.QueryString(qsTemp)))
					End If
				next
				Response.Write(""">")
				Response.Write("<img src=""GRIDLNKEDIT.GIF"" border=""0"" alt=""Edit SQL Select""></a> ")
				If (ocdDatabaseType = "SQLServer" And ocdSQLTableEdits) Or (ocdDatabaseType = "Access" And ocdAccessTableEdits) Then
					If Not CBool(CInt(ndnscCompatibility) And ocdNoJavaScript) Then
						Response.Write(" <input type=""image"" src=""appView.gif"" name=""Action"" value=""save..."" alt=""Save As View"" label=""Save As View"">")		
					Else
						Response.Write(" <input type=""submit"" class=""submit"" name=""Action"" value=""save..."">")		
					End If
				End If
				Response.Write(" <span class=""Information"">" ) 
				Response.Write(Trim(Server.HTMLEncode(objGrid.SQLText)) )
				Response.Write("</span>")
				Response.Write("</form>")
				Response.Write("")
			End If
			Response.Write("<table border=""0""><tr><td align=""top"" valign=""Top"" nowrap>")
			objGrid.Display("Buttons") 
			Response.Write("</td><td nowrap valign=""top"">")
			objGrid.Display("Keyword")
			Response.Write("</td><td valign=""top"">")
			If Not CBool(CInt(ndnscCompatibility) and ocdNoJavaScript) Then
				If ocdAllowBrowseRefresh Then
					Response.Write("<form method=""get"" action=""")
					Response.Write(Request.ServerVariables("SCRIPT_NAME"))
					Response.Write(""">")
					For Each qsTemp In Request.QueryString
						If UCase(qsTemp) <> "NDRELOADTIME" and UCASE(qsTemp) <> "NDRELOADSET" and UCASE(qsTemp) <> "NDRELOADACTION" Then
							Response.Write("<input name=""" & qsTemp & """")
							Response.Write(" value=""")
							Response.Write(Server.HTMLEncode(Request.QueryString(qsTemp)))
							Response.Write(""" type=""hidden"">")
						End If
					Next
					Response.write("</td><td width=""100%"">&nbsp;</td><td align=""left"" valign=""middle"" nowrap>")
					Response.Write("<span class=""FieldName""><small>Refresh<br>Every&nbsp;: </small>")
					Response.Write("</span> </td><td nowrap valign=""middle"">")
					If Not ((UCase(Request("ndreloadaction")) = "SET" or Request("ndreloadtime") <> "") and UCase(Request("ndreloadaction")) <> "STOP") Then
						Response.Write("<input type=""text"" class=""PagingControl"" name=""ndreloadtime"" value=""")
					End If
					If Request.QueryString("NDRELOADTIME") <> "" Then
						If UCase(Request("ndreloadaction")) <> "STOP" Then
							If CLng(Request.QueryString("NDRELOADTIME")) > 30 Then
								Response.Write(Server.HTMLEncode(Request.QueryString("NDRELOADTIME")))
							Else
								Response.Write("30")
							End If
						End If
					End If
					If Not((UCase(Request("ndreloadaction")) = "SET" or Request("ndreloadtime") <> "") and UCase(Request("ndreloadaction")) <> "STOP") Then
						Response.Write(""" size=""2"" maxlength=""5"">")
					End If
					Response.Write("<span class=""fieldname"">s</span>&nbsp;<input name=""ndReloadAction"" type=""submit"" class=""submit"" value=""")
					If (UCase(Request("ndreloadaction")) = "SET" or Request("ndreloadtime") <> "") and UCase(Request("ndreloadaction")) <> "STOP" Then
						Response.Write("stop")
					Else
						Response.Write("set")
					End If
					Response.Write("""> </td><td valign=""top""></form>")
			
				End If
			End If
			Response.Write("</td></tr></table>")
			Response.Flush()
			objGrid.Display("Grid") 
			objGrid.Display("TOTALS")
			If objGrid.AllowEdit = False And Not ocdReadOnly Then
				Response.Write("<p>An ")
				If ocdDatabaseType = "Access" Then
					Response.Write(" autonumber field ")
					If not ocdIsODBC Then
						Response.Write(" or eligible primary key ")
					End If
				ElseIf ocdDatabaseType = "SQLServer" Then
					Response.Write(" identity field ")
					If not ocdIsODBC Then
						Response.Write(" or eligible primary key ")
					End If
				ElseIf ocdDatabaseType = "MySQL" Then
					Response.Write(" autoincrement field ")
				Else
					Response.Write(" eligible primary key ")
				End If
				Response.Write(" is required to update data.</p>")
			End if
		Case "FILTER" 'Single Field Search
			Response.Write("<span class=Information>Set Filter</span>")
			objGrid.Display("Filter")
	End Select
	Call objGrid.Close()
	Set objGrid = Nothing
End Sub

'--------------------
'End Procedures
'--------------------

%>