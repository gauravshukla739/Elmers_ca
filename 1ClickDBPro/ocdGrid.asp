<%
'1 Click DB ASP Library - SQL to ASP Grid Object
'copyright 1997-2004 David J. Kawliche

'1 Click DB ASP Library source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'Use of this software and/or source code is strictly at your own risk.
'All warranties are specifically disclaimed except as required by law.

'IMPORTANT : THIS CODE USES PASS-THROUGH DATABASE SECURITY !
'
'To enforce application security, set appropriate logins 
'and permissions for all web server and database users.

'For more information see : http://1ClickDB.com

'**Start Encode**

Class ocdGrid
	'
	'This class renders an HTML interface to Browse - Search - Export SQL Queries using Active Server Pages (ASP) and Active Data Object (ADO) for MS Access, SQL Server, Oracle and other OLEDB providers including ODBC
	'
	Public GridID	'defaults To "A", alpha-numeric string To uniquely identify multiple grids on a page
	Public GridMode	'READ ONLY: 1 Click DB can be "BROWSE", "FILTER","SEARCH", "PROCESS" or "EXPORT", To Set .GridMode, use HTML syntax For links and form variables
	Public ADOConnection 'active database connection For Grid
	Public ADORecordSet 'active database cursor For Grid
	Public SQLConnect	'If ADOConnection is not Set prior To calling .Open , Then use this info To connect
	Public SQLUser	'database user name If using .SQLConnect
	Public SQLPass	'database password If using .SQLConnect
	Public SQLSelect	'comma delimited list of fields To retrieve, all fields mentioned In other SQLSelect... properties must be here
	Public SQLFrom 'REQUIRED: table To be displayed, SQL Join clauses are supported but are generally not updateable
	Public GridHideAutonumber	'If true, hide the autonumber field For display but include the name In .SQLSelect To enable editing

	Public GridFieldNames 'If true Then display grid header row
	Public GridFormatNumbers 'defaults To "" (database default) can also be "COMMAS"
	Public GridMaxMemo 'maximum number of characters To be displayed For a memo field
	Public SearchMultiSort 'If true, Order By's are additive
	Public SearchPersistKeyword 'remember criteria submitted on browse keyword search input, To avoid confusion, should only be used If ShowSearch and filter capabilities are off
	Public SearchCheckAll 'If true, all search fields are checked by default
	Public SearchSortSize 'number of order by fields To display In search form
	Public SearchDefaultTextCompare 'Default operator selected For Text field searches
	Public ADORecordsetTimeout  'how long before giving up on a database request, In seconds
	Public ADOComputeTimeout 'Seconds before giving up on retrieving aggregate information, including SQLRecordCount when optimized 
	Public ADOMaxRecords 'Absolute maximum number of records to attempt retrieval.
	Public GridMaxRecordsDisplay 'Absolute maximum number of records to display
	Public GridMaxURLLength 'try To intercept overly complex querystrings
	Public DatabaseType	'READ ONLY: identification of current database platform
	Public ExportForceDownload 'If true, sets MIME header on export pages as an attachment
	Public SQLSelectHide 'comma delimited list of fields To retrieve but hide from display
	Public SQLSelectName 'comma delimited list of "Nice" field names To display In Header, must have exactly the same number of entries as defined In SQLSelect
	Public SQLSelectAliases 'comma delimited list of selected field names that are aliases, suppresses search and grid field controls that can cause syntax errors
	Public SQLSelectSearchDropDown 'comma delimited list of fields that should be created as dropdowns of existing values In a search page
	Public SQLSelectFilterHide 'comma delimited list of no filter fields 
	Public SQLSelectSortHide 'comma delimited list of no sort fields
	Public SQLSelectSearchHide 'comma delimited list of no search fields
	Public SQLSelectSearchUncheck 'comma delimited list of fields To display unchecked by default In search screen
	Public SQLSelectFormat 	'1 Click DB field formatting variables
	Public SQLSelectSum 'comma delimited list of fields For computed Grand Total
	Public SQLSelectAvg 'comma delimited list of fields For computed Group Average
	Public SQLSelectMin 'comma delimited list of fields For computed Smallest Value
	Public SQLSelectMax 'comma delimited list of fields For computed Largest Value
	Public QuoteSuffix 'usually ansi quoted " or [
	Public QuotePrefix 'usually ansi quoted " or ]
	Public SQLWhere 'SQL WHERE clause without the word WHERE restricting which records To display
	Public SQLWhereExtra 'SQL WHERE clause without the keyword WHERE Used To restrict what rows from a table are displayed In addition To any criteria specified by SQLWhere.  Good For securing subsets of records from certain users
	Public SQLHaving 'SQL HAVING clause without the word HAVING used For display of aggregate queries
	Public SQLGroupBy 'SQL GROUP BY clause without the words GROUP BY used For display of aggregate querieds
	Public SQLOrderBy 'SQL ORDER BY without the words ORDER BY clause used For determining display order of records
	Public SQLOrderByDefault 'If no other order by is specified, use this
	Public SQLPage 'what page of records To display, If number is invalid, the first page is shown
	Public SQLPageSize 'number of records To display per page, If hard coded, will use default or querystring values
	Public SQLPageSizeDefault 'default number of records To display per page
	Public SQLRecordCount 'READ ONLY: how many records In current recordset
	Public SQLPageCount	'READ ONLY: how many pages In current recordset
	Public SQLText 'READ ONLY: actual SQL used To create main recordset
	Public AllowEdit 'If true, show edit links, If recordSet keys are not available, setting this To True will have no effect
	Public AllowDelete 'If true, show delete links, If recordSet keys are not available, setting this To True will have no effect
	Public AllowSearch 'If true, allow search/filter form processing, can be enabled For one grid per page
	Public AllowExport 'If true,  allow print and other exports, can be enabled For one grid per page
	Public AllowDetail 'If true,  show detail links
	Public AllowAdd	'If true show Add record button
	Public AllowMultiDelete 'If true, show check boxes to select records to delete
	Public SQLSelectID 'Deprecated, use SQLSelectIDName instead
	Public Criteria
	Public FormSearch 'ASP file for search requests, defaults To current page
	Public FormEdit	'ASP file for edit requests
	Public FormDelete 'ASP file for delete requests
	Public FormDetail 'ASP file For detail requests
	Public HTMLAttribGrid 'can be used To specify CELLPADDING and CELLSPACING attributes, CLASS is harcoded To "Grid"
	Public ShowDescription
	Public HTMLAttribButtonPanel
	Public HTMLAttribGridEven
	Public HTMLAttribGridCellAlign
	Public HTMLAttribGridOdd
	Public HTMLAttribBtnMultiDelete 'attributes for multidelete button
	Public HTMLAttribBtnMultiSelect
	Public HTMLGridButtons 'list of buttons To be displayed
	Public HTMLTextCompare 'list of valid text operators For search
	Public HTMLTrueValue 'If not set, grid will display database default representation
	Public HTMLFalseValue 'If not set, grid will display database default representation
	Public HTMLPagingStart
	Public HTMLPagingEnd
	Public HTMLTemplateColumn
	Public HTMLTemplateHeader

	Public HTMLDetailLink 'customize detail links
	Public HTMLNullValue 'replace null values with this
	Public HTMLBinaryValue 'replace binary values with this
	Public HTMLSortASCLink 'customize sort ascending links
	Public HTMLSortDESCLink	'customize sort descending links
	Public HTMLFilterLink 'customize filter links
	Public HTMLEditLink 'customize edit links
	Public HTMLDeleteLink 'customize delete links
	Public HTMLAttribEditLink 'customize edit links
	Public HTMLAttribDeleteLink 'customize delete links
	Public HTMLMemoContinues 'what To show at end of long memo
	Public HTMLAttribFindBtn 'defines HTML tag attributes For Find Records button
	Public HTMLAttribCancelBtn 'defines HTML tag attributes For Cancel button
	Public MaxKeywordFields
	Public AllowMultiSelect
	Public HTMLAttribGridCell
	Public HTMLExportStart 'in quick mode this starts export files
	Public HTMLGridVertical
	Public HTMLAfterFieldName
	Public HTMLExportEnd 'in quick mode this ends export forms
	Public HTMLAscText 'text displayed To select ASCending sort on search page
	Public HTMLDescText 'text displayed To select DESCending sort on search page
	Public HTMLSearchBetween 'text shown by default as between on date searches
	Public HTMLSearchBetweenAnd 'text shown between start and end inputs on date searches
	Public HTMLSearchAll 'text shown next To radio button To select "all" logic on searches
	Public HTMLSearchAny 'text shown next To radio button To select "any" logic on searches
	Public HTMLSearchAnd 'text shown next To radio button To add new restriction To current criteria
	Public HTMLSearchOr 'text shown next To radio button To add new match condition To current criteria
	Public HTMLSearchReSet 'text shown next To radio button To replace old criteria with new criteria
	Public HTMLSearchKeyword 'html For "Keyword" search field display 
	Public SearchKeywordTextFields
	Public SearchKeywordNumberFields
	Public ShowKeywordSearch
	Public 	RenderAsHTML 'Warning ! Setting this option to TRUE for untrusted data introduces Cross Site Scripting vulnerabilities.
	Public FormSelect
	Public ForceOptimization
	Public ForceNoOptimization

	Public UseRegExKeywordSearch
	Public ExportLineBreaks
	Public SQLSelectPK
	Public SQLSelectIDName 'single numeric key field for unique id links
	Public OracleDateFormat
	
	Private GridUseGetRows 'If true .ADORecordSet will be discarded immediately after moving data To local arrays
	Private arrSQLSelect
	Private arrSQLSelectType
	Private arrSQLSelectFormat
	Private arrSQLSelectName
	Private arrSQLSelectFilterHide
	Private arrSQLSelectSortHide
	Private arrstrSQLSelectPKType
	Private arrstrSQLSelectPKPos
	Private arrSQLSelectPK
	Private arrComputeData
	Private arrGridData
	Private intSQLSelectIDPos
	Private intSQLSelectIDPosType	
	Private intCurrentField
	Private strSQLSelectPKType
	Private strSQLSelectPKPos
	Private strSCRIPT_NAME
	Private strCData
	Public Debug 

	Private Sub Class_Initialize 'Set default values
	
			GridMaxRecordsDisplay = 1000
		HTMLAttribButtonPanel = " border=0 cellpadding=2 cellspacing=2 "
		HTMLAfterFieldName = "<br>"
		HTMLAttribGridCell = "valign=""top"" nowrap "
		Debug = False
		ForceOptimization = False
		AllowMultiSelect = False
		SearchKeywordTextFields = ""
		SearchKeywordNumberFields = ""
		HTMLGridVertical = False
		FormSelect = ""
		HTMLAttribGridCellAlign = ""
		ADOMaxRecords = 10000
		SQLSelectID = -1
		ADOComputeTimeOut = 10
		HTMLTemplateColumn = ""
		HTMLTemplateHeader = ""
		GridFieldNames = True
		UseRegExKeywordSearch = False
		SQLSelectSum = ""
		SQLSelectMin = ""
		FormEdit = ""
		ShowKeywordSearch = True
		FormDelete = ""
		ForceNoOptimization = False
		HTMLPagingStart = ""
		HTMLPagingEnd = ""
		AllowAdd = True
		RenderAsHTML = False 'Warning ! Setting this option to TRUE for untrusted data introduces Cross Site Scripting vulnerabilities.
		MaxKeywordFields = 32
		AllowEdit = True
		AllowDelete = True
		ShowDescription = False
		SQLSelectMax = ""
		SQLSelectAvg = ""
		ExportForceDownload = True
		SQLSelectSearchDropDown = ""
		SQLSelectSearchHide = ""
		ADORecordsetTimeout = 30
		SQLSelectPK = ""
		SQLSelectSearchUncheck = ""
		FormSearch = ""
		GridFormatNumbers = ""
		SQLSelectAliases = ""
		GridMaxURLLength = 2000
		GRIDUseGetRows = True
		ExportLineBreaks = True
		GridMaxMemo = 255
		SearchPersistKeyword = False
		SearchCheckAll = True
		AllowDetail = False
		SQLOrderByDefault = ""
		FormDetail = ""
		strSCRIPT_NAME = Request.ServerVariables("SCRIPT_NAME")
		SearchSortSize = 3
		SearchMultiSort = False
		SQLPageSizeDefault = 10
		AllowMultiDelete = False
		SQLSelectHide = ""
		DatabaseType = ""
		GridID = "A"
		SQLWhereExtra = ""
		SQLPageSize = ""
		SQLSelectName = ""
		SQLSelectIDName = ""
		intSQLSelectIDPos = -1
		GridMode = "Browse"
		QuoteSuffix = """"
		QuotePrefix = """"
		HTMLAttribEditLink = ""
		HTMLAttribDeleteLink = ""
		Select Case UCase(Request.QueryString("ocdGridMode_" & GRIDID))
			Case "FILTER"
				GridMode = "Filter"
			Case "SEARCH"
				GridMode = "Search"
			Case "EXPORT"
				GridMode = "Export"
			Case "PROCESS"
				GridMode = "Process"
		End Select
		OracleDateFormat = "DD-MON-YYYY"
	
%><!--#INCLUDE FILE=ocdGrid_Lang.asp--><%
	End Sub

	Private Sub Class_Terminate()	 'clean up ADO objects
		On Error Resume Next
		Call ADORecordset.close()
		Set ADORecordSet = Nothing
		If SQLConnect <> "" Then		  'disconnect if connection not Set to exisiting connection
			ADOConnection.Close()
			Set ADOConnection = Nothing
		End If
		Call Err.clear()
	End Sub

	Public Sub Open() 'Connect to DB and retrieve data
		Dim blnComputable, blnTemp, blnTemp2, rsIDX, rsCompute, intFieldCount, intTemp, strCSQL, strSQL, strTemp, strTemp2, eleTemp, eleTemp2, intRA, arrTemp, fldTemp
		'---BEGIN Initialize and check GRID variables
		If SQLSelectFormat <> "" Then
			arrSQLSelectFormat = Split(SQLSelectFormat, ",")
		End If
		If SQLSelectAliases <> "" Then
			If SQLSelectSearchHide <> "" Then
				SQLSelectSearchHide = SQLSelectSearchHide & "," & SQLSelectAliases
			Else
				SQLSelectSearchHide = SQLSelectAliases
			End If
			If SQLSelectSortHide <> "" Then
				SQLSelectSortHide = SQLSelectSortHide & "," & SQLSelectAliases
			Else
				SQLSelectSortHide = SQLSelectAliases
			End If
			If SQLSelectFilterHide <> "" Then
				SQLSelectFilterHide = SQLSelectFilterHide & "," & SQLSelectAliases
			Else
				SQLSelectFilterHide = SQLSelectAliases
			End If
		End If
		If SQLSelectFilterHide <> "" Then
			arrSQLSelectFilterHide = Split(SQLSelectFilterHide, ",")
		End If
		If SQLSelectSortHide <> "" Then
			arrSQLSelectSortHide = Split(SQLSelectSortHide, ",")
		End If
		If SQLSelect = "" Then
			SQLSelect = "*"
		End If
		SQLSelect = FormatForSQL(SQLSelect,DataBaseType,"CleanUserSQL")
		SQLFrom = FormatForSQL(SQLFrom,DataBaseType,"CleanUserSQL")
		If SQLPage = "" Then
			SQLPage = Request.QueryString("SQLPAGE_" & GRIDID)
		End If
		If IsNumeric(SQLPage) Then
			SQLPage = CInt(SQLPage)
			If SQLPage = 0 Then
				SQLPage = 1
			End If
		Else
			SQLPage = 1
		End If
		If SQLPageSize = "" Then
			If Request.QueryString("sqlpagesize" & "_" & GridID) <> "" And IsNumeric(Request.QueryString("sqlPagesize" & "_" & GridID)) Then
				SQLPageSize = CLng(Request.QueryString("sqlpagesize" & "_" & GridID))
			Else
				SQLPageSize = SQLPageSizeDefault
			End If
		Else
			If Not IsNumeric(SQLPageSize) Then
				SQLPageSize = SQLPageSizeDefault
			Else
				SQLPageSize = CLng(SQLPageSize)
			End If
		End If
		If SQLPageSize > GridMaxRecordsDisplay Then 
			SQLPageSize = GridMaxRecordsDisplay
		End If
		SQLGroupBy = FormatForSQL(SQLGroupBy,DataBaseType,"CleanUserSQL")
		SQLHaving = FormatForSQL(SQLHaving,DataBaseType,"CleanUserSQL")
		SQLSelectSum =  FormatForSQL(SQLSelectSum,DataBaseType,"CleanUserSQL")
		SQLSelectMin = FormatForSQL(SQLSelectMin,DataBaseType,"CleanUserSQL")
		SQLSelectMax = FormatForSQL(SQLSelectMax,DataBaseType,"CleanUserSQL")
		SQLSelectAvg = FormatForSQL(SQLSelectAvg,DataBaseType,"CleanUserSQL")
		If SQLOrderBy = "" Then
			If Request.QueryString("sqlorderby" & "_" & GridID) <> "" Then
				SQLOrderBy = Request.QueryString("sqlorderby" & "_" & GridID)
			Else
				SQLOrderBy = SQLOrderByDefault
			End If
		End If
		SQLOrderBy = FormatForSQL(SQLOrderBy,DataBaseType,"CleanUserSQL")
		'Initiaize and check database connection
		If Not IsObject(ADOConnection) Then
			Set ADOConnection = Server.CreateObject("ADODB.Connection")
			ADOConnection.Mode = 1			 'adModeRead
			Call ADOConnection.Open(SQLConnect, SQLUser, SQLPass)
		End If
		If DatabaseType = "" Then
			DatabaseType = getDatabaseType(ADOConnection)
		End If
		Select Case UCase(DatabaseType)
			Case "IXS", "ADSI"
				AllowEdit = False
				AllowDelete = False
				AllowAdd = False
			Case "ACCESS", "SQLSERVER"
		End Select
		If Not AllowEdit Then
			HTMLEditLink = ""
		End If
		If Not AllowDelete Then
			HTMLDeleteLink = ""
		End If
		If SQLWhere = "" Then
			SQLWhere = Request.QueryString("SQLWHERE_" & GRIDID)
		End If
			'check For batched commands
		SQLWhere = FormatForSQL(SQLWhere,DataBaseType,"CleanUserSQL")
'		Response.write sqlwhere
		'---END Initiaize and check Grid Variables
		Select Case UCase(GridMode)
			Case "FILTER"
				Exit Sub
		End Select
		'Check eligibility for query optimization and aggregates
		blnComputable = False 
		If UCase(GridMode) = "BROWSE" Then
	
			If ForceOptimization Then
				blnComputable = True
			ElseIf ForceNoOptimization Then
							blnComputable = False
			Else
				If InStr(SQLSelect, ")") = 0 And InStr(SQLSelect, "(") = 0 And SQLGroupBy = "" And SQLHaving = "" And DatabaseType <> "IXS" And DatabaseType <> "ADSI" And ((Len(sqlselect) > 2) Or (sqlselect = "*")) Then

					If UCase(Left(sqlselect, 3)) <> "TOP" and UCase(Left(sqlselect, 8)) <> "DISTINCT" Then

						blnComputable = True
					End If
				End If
			End If
		End if
		If UCase(GridMode) = "BROWSE" Then
			If InStr(SQLSelect, ")") = 0 And InStr(SQLSelect, "(") = 0 And SQLGroupBy = "" And SQLHaving = "" And DatabaseType <> "IXS" And DatabaseType <> "ADSI" And ((Len(sqlselect) > 2) Or (sqlselect = "*")) Then
				If UCase(Left(sqlselect, 3)) <> "TOP" and UCase(Left(sqlselect, 8)) <> "DISTINCT" Then
					blnComputable = True
				End If
			End If
		End If
		If Debug Then 
			response.write "comput" & blnComputable
		End If
		If UCase(GridMode) <> "BROWSE" Then
			
			blnComputable = False
		End If
'		If SQLWhereExtra <> ""
'			blnComputable = False
'		End If
		If DatabaseTYpe = "MySQL" Then
			blncomputable = False
		End if
		If Debug Then 
			response.write "comput" & blnComputable
		End If
		'----BEGIN construct Main SQL Statement
'		response.write blncomputable
		If blnComputable Then
			If DatabaseType = "Access" Then
				strSQL = "SELECT TOP " & (SQLPageSize * SQLPage) & " " & SQLSelect
			Else
				strSQL = "SELECT " & SQLSelect
			End If
		Else
			strSQL = "SELECT " & SQLSelect
			GRIDUseGetRows = False
		End If
		strSQL = strSQL & " FROM "
		If InStr(1, SQLFrom, ",") = 0 And InStr(1, SQLFrom, "=") = 0 Then
			SQLFrom = FormatForSQL(SQLFrom, DatabaseType, "AddSQLIdentifier")
		End If
		strSQL = strSQL & SQLFrom
		If UCase(GridMode) = "PROCESS" Or UCase(GridMode) = "SEARCH" Then
			strSQL = strSQL & " WHERE 1=2"
		ElseIf SQLWhereExtra <> "" Then
			If SQLWhere = "" Then
				strSQL = strSQL & " WHERE " & SQLWhereExtra
			Else
				strSQL = strSQL & " WHERE (" & SQLWhere & ") AND " & SQLWhereExtra & ""
			End If
		ElseIf (Not SQLWhere = "") Then
			strSQL = strSQL & " WHERE " & SQLWhere & ""
		End If
		If Not SQLGroupBy = "" Then
			strSQL = strSQL & " GROUP BY " & SQLGroupBy
		End If
		If Not SQLHaving = "" Then
			strSQL = strSQL & " HAVING " & SQLHaving
		End If
		If Not SQLOrderBy = "" Then
			strSQL = strSQL & " ORDER BY " & SQLOrderBy
		End If
		If SQLPage = "" Then
			SQLPage = Request.QueryString("sqlpage" & "_" & GridID)
		End If
		If Not IsNumeric(SQLPage) Then
			SQLPage = 1
		End If
		If Debug Then
			response.write strSQL
		End If
		'----END construct Main SQL Statement
		'----BEGIN construct Computed SQL Statment
		If blnComputable Then
			strCSQL = "SELECT COUNT(*) AS " & FormatForSQL("Record Count", DatabaseType, "AddSQLIdentifier") & ","
			If SQLSelectSum <> "" Then
				arrTemp = Split(SQLSelectSum, ",")
				For Each eleTemp In arrTemp
					strCSQL = strCSQL & "SUM(" & FormatForSQL(eleTemp, DatabaseType, "AddSQLIdentifier") & ") As " & FormatForSQL("SUM " & FormatForSQL(Replace(eleTemp, " ", ""), DatabaseType, "RemoveSQLIdentifier"), DatabaseType, "AddSQLIdentifier") & ","
				Next
			End If
			If SQLSelectAvg <> "" Then
				arrTemp = Split(SQLSelectAvg, ",")
				For Each eleTemp In arrTemp
					strCSQL = strCSQL & "AVG(" & FormatForSQL(eleTemp, DatabaseType, "RemoveSQLIdentifier") & ") As " & FormatForSQL("AVG " & FormatForSQL(Replace(eleTemp, " ", "") , DatabaseType, "RemoveSQLIdentifier"), DatabaseType, "AddSQLIdentifier") & ","
				Next
			End If
			If SQLSelectMin <> "" Then
				arrTemp = Split(SQLSelectMin, ",")
				For Each eleTemp In arrTemp
					strCSQL = strCSQL & "MIN(" & FormatForSQL(eleTemp, DatabaseType, "RemoveSQLIdentifier") & ") As " & FormatForSQL("MIN " & FormatForSQL(Replace(eleTemp, " ", ""), DatabaseType, "RemoveSQLIdentifier"), DatabaseType, "AddSQLIdentifier") & ","
				Next
			End If
			If SQLSelectMax <> "" Then
				arrTemp = Split(SQLSelectMax, ",")
				For Each eleTemp In arrTemp
					strCSQL = strCSQL & "MAX(" & FormatForSQL(eleTemp, DatabaseType, "RemoveSQLIdentifier") & ") As " & FormatForSQL("MAX " & FormatForSQL(Replace(eleTemp, " ", "") , DatabaseType, "RemoveSQLIdentifier"), DatabaseType, "AddSQLIdentifier") & ","
				Next
			End If
			strCSQL = Left(strCSQL, Len(StrCSQL) - 1)
			strCSQL = strCSQL & " FROM " & SQLFrom
			If SQLWhereExtra <> "" Then
				If SQLWhere = "" Then
					strCSQL = strCSQL & " WHERE " & SQLWhereExtra
				Else
					strCSQL = strCSQL & " WHERE (" & SQLWhere & ") AND " & SQLWHEREExtra & ""
				End If
			ElseIf (Not SQLWhere = "") Then
				strCSQL = strCSQL & " WHERE " & SQLWhere & ""
			End If
			'----END construct Computation SQL Statement
			'Retrieve SQL Computations
			Set rsCompute = Server.CreateObject("ADODB.Recordset")
			rsCompute.CursorLocation = 2			 'adUseServer
			Select Case DatabaseType
				Case "Access", "SQLServer"
					ADOConnection.CommandTimeout = ADOComputeTimeout
			End Select
			If Debug then
				Response.Write strCSQL
			Else
				On Error Resume Next
			End If
			Call rsCompute.Open(strCSQL, ADOConnection)
			If Err.Number <> 0 Then
				Select Case Err.Number
					Case -2147217871
						GridUseGetRows = False
						If SQLSelectMin <> "" Or SQLSelectMax <> "" Or SQLSelectAvg <> "" Or SQLSelectSum <> "" Then
							strCData = "Timeout on Summary Totals"
						End If
						Err.Clear
					Case -2147217911
						'Response.Write("Access Denied")
						'err.clear
						Exit Sub
					Case Else
						If Not Debug Then
						On Error GoTo 0
						Err.Raise(13)
						End If
						Exit Sub
				End Select
			Else
				'Process computations, including optimized recordcount
				If rsCompute.Fields.count > 1 Then
					strCData = "<table>"
				End If
				For Each fldTemp In rsCompute.Fields
					If fldTemp.Name = "Record Count" Then
						SQLRecordCount = CLng(rsCompute(0))
					Else
						'Create HTML for Grouping summary
						strCData = strCData & "<tr>"
						strCData = strCData & "<td><span class=""FieldName"">" & fldTemp.Name & "</span></td>"
						strCData = strCData & "<td align=""right"">" 
						If not isnull(fldTemp.Value) and UCase(GridFormatNumbers) = "COMMAS" Then
							strCData = strCData &	FormatNumber(fldTemp.Value, 0, 0, 0, -1)
						Else
							strCData = strCData & fldTemp.Value
						End if
						strCData = strCData & "</td>"
						strCData = strCData & "</tr>"
					End If
				Next
				If rsCompute.Fields.Count > 1 Then
					strCData = strCData & "</table>"
				End If
				rsCompute.Close()
			End If
			Set rsCompute = Nothing
		End If
		If GridUseGetRows Then
			Set ADORecordSet = Server.CreateObject("ADODB.Recordset")
			ADORecordset.Maxrecords = CLng(SQLPageSize * SQLPage)
			ADORecordset.CursorLocation = 2			  'adUseServer 3 
			If Debug Then
				response.write strSQL
			Else
				On Error Resume Next
			End If
			
				Select Case DatabaseType
					Case "SQLServer"
						ADOConnection.CommandTimeout = ADORecordsetTimeout
						dim cmdTarget
						set cmdTarget = Server.CreateObject("ADODB.Command")
						cmdTarget.CommandType = &H0001 'Text
						cmdTarget.CommandTimeout = ADORecordsetTimeout
						cmdTarget.CommandText = strSQL
						cmdTarget.ActiveConnection = ADOConnection
						Set ADORecordset = cmdTarget.Execute()
					Case Else
					
				Select Case DatabaseType
					Case "Access", "SQLServer"
						If Debug then
							response.write "SetTimeout=" & ADORecordsetTimeout
						End If
						ADOConnection.CommandTimeout = ADORecordsetTimeout

				End Select
'adUseClient
			Call ADORecordset.Open(strSQL, ADOConnection, 0, 1, &H1)				  'adOpenForwardOnly,  adLockReadOnly, adCmdText
			End Select
			Select Case Err.Number
				Case 0

				Case -2147217871
					'Timeout, pass to user
					exit sub

				Case Else
		
					If Not Debug Then
						Err.Clear
						on error goto 0
						Err.raise(13)
					End If
					exit sub
				
			End Select
			on error goto 0
			If SQLWhereExtra <> "" And Not ForceOptimization Then
'				ADORecordset.Filter = SQLWhereExtra
			End If
		Else
			'retrieve full dataset
			Select Case DatabaseType
				Case "Access", "SQLServer"
					ADOConnection.CommandTimeout = ADORecordsetTimeout
			End Select
			Set ADORecordSet = Server.CreateObject("ADODB.Recordset")
			ADORecordset.CursorLocation = 2			  'adUseServer 3 'adUseClient
			If UCase(GridMode) <> "EXPORT" Then
				ADORecordset.maxrecords = ADOMaxRecords + 1
			Else
					ADORecordset.maxrecords = GridMaxRecordsDisplay + 1
			End If
			ADORecordset.CursorLocation = 3
			On Error Resume Next
			Call ADORecordset.Open(strSQL, ADOConnection, 3, 3, &H1			)				 'adOpenStatic, adLockReadOnly, adCmdText
			If Debug Then
				Response.Write(Err.Number & Err.Description)
			End If
			Select Case Err.Number
				Case 0
				Case -2147217871
					exit sub
				Case Else
					Err.Clear
					on error goto 0
					Err.raise(13)
					exit sub
			End Select
			on error goto 0
			If SQLWhereExtra <> "" And Not ForceOptimization Then
'				ADORecordset.Filter = SQLWhereExtra
			End If
			SQLRecordCount = ADORecordset.RecordCount
		End If
		'Compute paging and final counts
		If IsNumeric(SQLRecordCount) Then
			If SQLRecordCount > ADOMaxRecords Then
				SQLPageCount = CLng(Fix(CLng(ADOMaxRecords) / CLng(SQLPageSize)))
				If (CLng(ADOMaxRecords) / CLng(SQLPageSize)) <> Fix(CLng(ADOMaxRecords) / CLng(SQLPageSize)) Then
					SQLPageCount = SQLPageCount + 1
				End If
			Else
				SQLPageCount = CLng(Fix(CLng(SQLRecordCount) / CLng(SQLPageSize)))
				If (CLng(SQLRecordCount) / CLng(SQLPageSize)) <> Fix(CLng(SQLRecordCount) / CLng(SQLPageSize)) Then
					SQLPageCount = SQLPageCount + 1
				End If
			End If
		Else
			SQLPageCount = CLng(Fix(CLng(ADOMaxRecords) / CLng(SQLPageSize)))
			If (CLng(ADOMaxRecords) / CLng(SQLPageSize)) <> Fix(CLng(ADOMaxRecords) / CLng(SQLPageSize)) Then
				SQLPageCount = SQLPageCount + 1
			End If
		End If
		If 1 > SQLPage Then
			SQLPage = 1
		ElseIf SQLPage > SQLPageCount Then
			SQLPage = SQLPageCount
		End If
		If UCase(Left(Cstr(strSQL), 12 + Len(Cstr(SQLPageSize)))) = "SELECT TOP " & Trim(Cstr(SQLPageSize * SQLPage)) & " " And UCase(left(SQLSelect, 4)) <> "TOP " Then
			SQLText =  "SELECT " & (Mid(Cstr(strSQL), 13 + Len(Cstr(SQLPageSize * SQLPage))))
		Else
			SQLText = strSQL
		End if
		strTemp = ""
		strTemp2 = ""
		intFieldCount = 0
		'Check for primary key and retireve field types to Private array
		For Each fldTemp In ADORecordset.Fields
			strTemp = strTemp & fldTemp.Name & ";;"
			strTemp2 = strTemp2 & CStr(fldTemp.Type) & ","
			If UCase(GridMode) = "BROWSE"  Then 'And intSQLSelectIDPos = -1
				If SQLSelectIDName <> "" Then
					If fldTemp.Name = SQLSelectIDName Then
						intSQLSelectIDPos = intFieldCount
					End If
				ElseIf SQLSelectID <> -1 Then
					If intFieldCount = Cint(SQLSelectID) Then
						SQLSelectIDName = fldTemp.Name
						intSQLSelectIDPos = SQLSelectIDName
					End If
				Else
					Select Case DatabaseType
						Case "Access", "SQLServer","MySQL"
							If InStr(SQLSelect, ")") = 0 And InStr(SQLSelect, "(") = 0 And SQLGroupBy = "" And SQLHaving = "" Then
								If UCase(ADOConnection.provider) = "MICROSOFT.JET.OLEDB.3.51" Or DatabaseType = "IXS" Then
									'look out, field properties bomb with this provider
									intSQLSelectIDPos = -1
								Else
									'Server Cursor required
									If CBool(fldTemp.Properties("ISAUTOINCREMENT")) = True Then
										intSQLSelectIDPos = intFieldCount
										SQLSelectIDName = fldTemp.Name
									End If
								End If
							End If
					End Select
				End If
			End If
			intFieldCount = intFieldCount + 1
		Next
		If strTemp = ";;" Then
			strTemp = "-;;"
		End If
		If Len(strTemp) > 1 Then
			arrSQLSelect = Split(Left(strTemp, Len(strTemp) - 2), ";;")
			arrSQLSelectType = Split(Left(strTemp2, Len(strTemp2) - 1), ",")
		End If
		
		If UCase(GridMode) = "EXPORT" And UCase(Request.QueryString("ocdExportFormat_" & GRIDID)) = "XML" Then
			Response.Clear
			Response.ContentType = "text/xml"
			If SQLRecordCount > GridMaxRecordsDisplay Then
				Response.Write("<xml><warning>Display Limited Exceeded</warning></xml>")
			Else
				Call ADORecordset.Save(response, 1)
			End If
			Call Close()
			Response.End()
		End If
		If GridUseGetRows Then
			If Not ADORecordset.eof Then
				Select Case UCase(GridMode)
					Case "PROCESS", "SEARCH"
						'no data To retrieve
					Case "BROWSE"
						'move To the right page	and retrieve data
							If Not (SQLPage = 1 Or SQLPage = 0) Then
								On Error Resume Next
								ADORecordset.move(CLng((SQLPage - 1) * SQLPageSize))
								Err.Clear()
								On Error GoTo 0
							End If
							arrGridData = ADORecordset.GetRows(SQLPageSize)
					Case "EXPORT"
						Select Case UCase(Request.QueryString("ocdExportFormat_" & GRIDID))
							Case Else
								arrGridData = ADORecordset.GetRows()
						End Select
						SQLRecordcount = UBound(arrGridData, 2) + 1
				End Select
				Call ADORecordset.Close()
				Set ADORecordSet = Nothing
			Else
				SQLRecordCount = 0
			End If
		Else
			SQLRecordCount = ADORecordset.RecordCount
		End If
		If SQLSelectName <> "" Then
			arrSQLSelectName = split(SQLSelectName,",")
			SetCriteria()
		End If
		Select Case UCase(GridMode)
			Case "PROCESS"
				Call ProcessSearch()
			Case "SEARCH"
				Exit Sub
			Case "EXPORT"			 'Set Mime Headers as appropriate
				
				Response.clear()
				
				AllowMultiDelete = False
				AllowMultiSelect = False
				Select Case UCase(Request.QueryString("ocdExportFormat_" & GRIDID))
					Case "EXCEL", "TEXT"
						Response.Clear

						If Not Debug Then
							If ExportForceDownload Then
								Select Case UCase(Request.QueryString("ocdExportFormat_" & GRIDID))
									Case "EXCEL"
											
										Call Response.AddHeader("content-disposition", "attachment; filename=export.xls")
									Case "TEXT"
										Call Response.AddHeader("content-disposition", "attachment; filename=export.txt")
								End Select
							End If
							Response.ContentType = "application/vnd.ms-excel"
						End If
						'Response.Write(HTMLExportStart)
					Case "TSV"
						Response.ContentType = ("application/tab-separated-values")
						If ExportForceDownload Then
							Call Response.AddHeader("content-disposition", "attachment; filename=export.txt")
						Else
							Call Response.AddHeader("content-disposition", "inline;filename=export.txt")
						End If
					Case "PRINT"
						Response.Write(HTMLExportStart)
				End Select
				Exit Sub
			Case "BROWSE"
				If CInt(SQLPageSize) = 1 Then
					AllowMultiDelete = False
				End If
				'If no key information required but still not defined, Then attempt To autodetect
				strSQLSelectPKType = ""
				If (intSQLSelectIDPos = -1 And SQLSelectPK = "" And (AllowEdit Or AllowDetail Or AllowDelete) ) Then				'try To determine pk fields
					SQLSelectPK = getPKFields(ADOConnection, DatabaseType, SQLFrom, QuoteSuffix, QUotePrefix)
				End If
					If SQLSelectPK <> "" Then
						'from field list determine other props on rs
						arrSQLSelectPK = Split(SQLSelectPK, ",")
						blnTemp = False
						If CInt(SQLPageSize) = 1 Then
							AllowMultiDelete = False
						End If
						For Each eleTemp2 In arrSQLSelectPK
							intTemp = 0
							blnTemp2 = False
							For Each eleTemp In arrSQLSelect
								If UCase(eleTemp) = UCase(eleTemp2) Then
									strSQLSelectPKPos = strSQLSelectPKPos & intTemp & ","
									strTemp2 = arrSQLSelectType(intTemp)
									strSQLSelectPKType = strSQLSelectPKType & CStr(strTemp2) & ","
									If strTemp2 = "135" Then
										blntemp2 = False
										strSQLSelectPKType = ""
										strSQLSelectPKPos = ""
										SQLSelectPK = ""
										Exit For
									End If
									blntemp2 = True
									Exit For
								End If
								intTemp = intTemp + 1
							Next
							If Not blnTemp2 Then
								blnTemp = True
								SQLSelectPK = ""
							End If
						Next
					End If
					If SQLSelectPK <> "" Then
						strSQLSelectPKType = Left(strSQLSelectPKType, Len(strSQLSelectPKType) - 1)
						arrstrSQLSelectPKType = Split(strSQLSelectPKType, ",")
						strSQLSelectPKPos = Left(strSQLSelectPKPos, Len(strSQLSelectPKPos) - 1)
						arrstrSQLSelectPKPos = Split(strSQLSelectPKPos, ",")
					End If
'				End If
				If intSQLSelectIDPos = -1 Or Not AllowDelete Or Not GridUseGetRows Or UCase(GridMode) <> "BROWSE" Then
					AllowMultiDelete = False
				End If
				If intSQLSelectIDPos = -1 And SQLSelectPK = "" Then
					'no key To identify individual records is available
					AllowAdd = False
					AllowEdit = False
					AllowDelete = False
					AllowDetail = False
				End If
		End Select
		If CInt(SQLPageSize) = 1 Then
			AllowMultiDelete = False
		End If
	End Sub

	Public Sub Display(ByVal strTemplate)
		'Generic method to write HTML interface blocks
		Dim qstTemp, intRowCount, intSPFC, intLoopSize, strTemp, strCDp, strFname, strInfqstTemp, strPKqstTemp, arrSQLSelectPK, arrSQLSelectHide, tmpSQLPKString, varPKFieldType, fldTemp, blnTemp, eleTemp, intPKPos, arrHyperLink, strCD, strName, varValue, intType, strRTempStart, strRTemp, strToEval, varEval, strRTempEnd
		If Err.number <> 0 Then
			Exit Sub
		End If
		Select Case Trim(UCase((strTemplate)))
			Case "QUICK", ""
				Call Open()
				If UCase(GridMode) = "FILTER" Then
					Call DisplayFilter(GRIDID)
					Exit Sub
				End If
			Case "FILTER"
				If UCase(GridMode) = "FILTER" Then
					Call DisplayFilter(GRIDID)
					Exit Sub
				End If
		End Select
		Select Case UCase((strTemplate))
			Case "QUICK", "", "SEARCH"
				If UCase(GridMode) = "SEARCH" Then
					Call DisplaySearch(GRIDID)
					Exit Sub
				End If
		End Select
		GridID = UCase(GridID)
		If SQLSelectHide = "" Then
			SQLSelectHide = Request.QueryString("SQLSELECTHIDE_" & GRIDID)
		End If
		If Not UCase(GridMode) = "EXPORT" Then
			Select Case UCase(strTemplate)
				Case "QUICK", "", "BUTTONS"
					If UCase(GridMode) <> "EXPORT" Then
						Call DisplayGridButtons()
					End If
			End Select
		End If
		If SQLSelectHide <> "" Then
			arrSQLSelectHide = Split(SQLSelectHide, ",")
		End If
		intCurrentField = 0
		Select Case UCase(strTemplate)
			Case "QUICK", "", "GRID"
				If SQLSelectName <> "" Then
					ArrSQLSelectName = Split(SQLSelectName, ",")
				End If
				Select Case UCase(GridMode)
					Case "EXPORT"
						AllowAdd = False
						AllowEdit = False
						AllowDetail = False
						AllowDelete = False
						AllowMultiDelete = False
						AllowMultiSelect = False
						Select Case UCase(Request.QueryString("ocdExportFormat_" & GRIDID))
							Case "PRINT"
								Response.Write(HTMLExportStart)
						End Select
				End Select
				'Start Building results table
				If FormDelete = "" Then
					FormDelete = FormEdit
				End If
				If AllowMultiDelete Then
					Response.Write("<form method=""get"" action=""" & FormDelete & """><input type=hidden name=sqlfrom value=""" & Server.HTMLEncode(SQLFrom) & """>")
				ElseIf AllowMultiSelect Then
					Response.Write("<form method=""get"" action=""" & FormSelect & """>")
				End If
				If IsNumeric(SQLRecordCount) Then
					If SQLRecordCount = 0 Then
						intLoopSize = 0
					Else
						If GridUseGetRows Then
							intLoopSize = UBound(arrGridData, 2) + 1
						Else
							If UCase(GridMode) = "EXPORT" Then
								If IsNumeric(SQLRecordCount) Then
									intLoopSize = SQLRecordCount
								Else
									intLoopSize = ADOMaxRecords
								End If
							if intLoopSize > GridMaxRecordsDisplay Then
								intLoopSize = GridMaxRecordsDisplay
						End If
							Else
								intLoopSize = SQLPageSize
							End If
						End If
					End If
				Else
					If GridUseGetRows Then
						intLoopSize = UBound(arrGridData, 2) + 1
						If UCase(GridMode) = "EXPORT" Then
							if intLoopSize > GridMaxRecordsDisplay Then
								intLoopSize = GridMaxRecordsDisplay
							End If
						End If
					Else
						If UCase(GridMode) = "EXPORT" Then
							If IsNumeric(SQLRecordCount) Then
								intLoopSize = SQLRecordCount
							Else
								intLoopSize = ADOMaxRecords
							End If
							if intLoopSize > GridMaxRecordsDisplay Then
								intLoopSize = GridMaxRecordsDisplay
							End If
						Else

							intLoopSize = SQLPageSize
						End If
					End If
				End If
				If intLoopSize = GridMaxRecordsDisplay AND UCase(GridMode) <> "BROWSE" Then
					Response.Write("<p>Display Limit Exceeded - First " & GridMaxRecordsDisplay & " Records Displayed</p>")
				End if
				If Not HTMLGridVertical Then
					Response.Write("<table class=""Grid"" " & HTMLAttribGrid & ">")
				End If
				If GridFieldNames And Not HTMLGridVertical Then
					Response.Write("<tr " & "class=""GridHeader""" & ">")
					If UCase(GridMode) <> "EXPORT" Then
					Response.Write("<th")
					Response.Write(">")
					If AllowMultiSelect Then
						Response.Write("<input name=""OCDMULTISELECT"" " & HTMLAttribBtnMultiSelect & ">")
					Else
						Response.Write("&nbsp;")
					End If
					Response.Write("</th>")
					End If
					If HTMLTemplateColumn <> "" Then 
						Response.Write("<th>")
						Response.Write(HTMLTemplateHeader)
						Response.Write("</th>")
					End If
					'Create Column Header
					For Each fldTemp In arrSQLSelect
						blnTemp = True
						If (CInt(intCurrentField) = CInt(intSQLSelectIDPos)) And GridHideAutonumber Then
						Else
							If SQLSelectHide <> "" Then
								For eleTemp = 0 To UBound(arrSQLSelectHide)
									If UCase(arrSQLSelect(intCurrentField)) = UCase(arrSQLSelectHide(eleTemp)) Then
										blnTemp = False
										Exit For
									End If
								Next
							End If
							If blnTemp Then
								If SQLSelectName = "" Then
									Call DisplayFieldName(ArrSQLSelect(intCurrentField), ArrSQLSelectType(intCurrentField), 50, "")
								Else
									Call DisplayFieldName(ArrSQLSelect(intCurrentField), ArrSQLSelectType(intCurrentField), 50, ArrSQLSelectName(intCurrentField))
								End If
							End If
						End If
						intCurrentField = intCurrentField + 1
					Next
					If UCase(GridMode) <> "EXPORT" Then
					Response.Write("<th>")
					If AllowMultiDelete And IsNumeric(SQLRecordCount) Then
						If SQLRecordCount <> 0 Then
							Response.Write("<input name=""OCDEDITDELETE"" " & HTMLAttribBtnMultiDelete & ">")
							For Each qstTemp In Request.QueryString
								Select Case UCase(qstTemp)
									Case "OCDEDITDELETE", "SQLID"
									Case "SQLFROM"
										Response.Write("<input type=""hidden"" name=""SQLFrom"" value=""" & Server.URLEncode(FormatForSQL(SQLFrom, DatabaseType, "REMOVESQLIDENTIFIER")) & """>")
									Case Else
										Response.Write("<input type=""hidden"" name=""" & qstTemp & """ value=""" & Server.htmlencode(Request.QueryString(qstTemp)) & """>")
								End Select
							Next
							Response.Write("</th>")
							End If
							Response.Write("</tr>")
							End If
						End If
					End If
				If SQLRecordCount = 0 Then
					Response.Write("</table>")
					Call Close()
					Exit Sub
				End If
				intRowCount = 1
				Select Case UCase(FormEdit)
					Case "EDIT.ASP", "FREEEDIT.ASP"
						strInfqstTemp = "&amp;"
						strInfqstTemp = strInfqstTemp & "sqlfrom=" & Server.URLEncode(SQLFrom)
				End Select
				For Each qstTemp In Request.QueryString
					Select Case UCase(qstTemp)
						Case "SQLID", "SQLWHERE", "SQLFROM", "OCDACTION" & "_" & GridID
						Case Else
							strInfqstTemp = strInfqstTemp & "&amp;" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp))
					End Select
				Next
				If SQLSelectPK <> "" Then
					arrSQLSelectPK = Split(SQLSelectPK, ",")
				End If
				If Not GridUseGetrows And UCase(GridMode) = "BROWSE" Then
					If Not ADORecordset.Eof Then
						If SQLPage > 1 Then
							ADORecordset.Move(CLng(SQLPageSize * (SQLPage - 1)))
						End If
					End If
				End If
				For intRowCount = 1 To CInt(intLoopSize)
					If HTMLGridVertical Then
						Response.Write("<table><tr>")
					Else
						Response.Write("<tr ")
						If intRowCount Mod 2 = 0 Then
							Response.Write(HTMLAttribGridEven)
						Else
							Response.Write(HTMLAttribGridOdd)
						End If
						Response.Write(">")
					End If
					If HTMLGridVertical Then
						Response.Write("<td colspan=""2"">")
					Else
						If Not UCase(GridMode) = "EXPORT" Then
						Response.Write("<td>")
						End If
					End If
					If (intSQLSelectIDPos <> -1 Or SQLSelectPK <> "") And (AllowEdit Or AllowDetail Or AllowDelete Or AllowMultiSelect Or AllowMultiDelete) Then
						strPKqstTemp = ""
						If SQLSelectIDName <> "" Then
								If Not GridUseGetRows Then
									strPKqstTemp = strPKqstTemp & "sqlid=" & Server.URLEncode(ADORecordset.Fields(intSQLSelectIDPos).Value)
								Else
									If Not isnull(arrGridData(intSQLSelectIDPos, intRowCount - 1)) Then
										strPKqstTemp = strPKqstTemp & "sqlid=" & Server.URLEncode(arrGridData(intSQLSelectIDPos, intRowCount - 1))
									Else
										AllowEdit = False
										AllowDelete = False
									End If
								End If
						ElseIf SQLSelectPK <> "" Then
							strPKqstTemp = strPKqstTemp & "sqlwhere="
							strCDp = ""
							tmpSQLPKString = ""
							For intSPFC = 0 To UBound(arrSQLSelectPK)
								tmpSQLPKString = tmpSQLPKString & FormatForSQL(arrSQLSelectPK(intSPFC), DatabaseType, "AddSQLIdentifier") & "="
								intPKPos = arrstrSQLSelectPKType(intSPFC)
								intPKPos = CInt(intPKPos)
								Select Case intPKPos
									Case 2, 3, 4, 5, 14, 16, 17, 18, 19, 20, 21, 128, 131, 204, 6, 11									 'adSmallInt, adInteger, adSingle, adDouble, adDecimal, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adBigInt, adUnsignedBigInt, adBinary, adNumeric, adVarBinary, adLongVarBinary, adCurrency, adBoolean
										strCDp = ""
									Case 135, 7, 133, 134									 'adDBTimeStamp, adDate, adDBDate,  adDBTime
										If DatabaseType = "Access" Then
											strCDp = "#"
										Else
											strCDp = "'"
										End If
									Case 8, 129, 130, 200, 201, 202, 203									 'adBSTR, adChar, adWChar, adVarChar, adLongVarChar, adVarWChar, adLongVarWChar						
										strCDp = "'"
								End Select
								If GridUseGetRows Then
									Select Case intPKPos
										Case 8, 129, 130, 200, 201, 202, 203										'adBSTR, adChar, adWChar, adVarChar, adLongVarChar, adVarWChar, adLongVarWChar		
											tmpSQLPKString = tmpSQLPKString & strCDp & Replace(arrGridData(arrstrSQLSelectPKPos(intSPFC), intRowCount - 1), "'", "''") & strCDp & " AND "
										Case 135, 7, 133, 134										'adDBTimeStamp, adDate, adDBDate,  adDBTime
											tmpSQLPKString = tmpSQLPKString & strCDp & FormatForSQL(CStr(arrGridData(arrstrSQLSelectPKPos(intSPFC), intRowCount - 1)), DatabaseType, "SafeDate") & strCDp & " AND "
										Case Else
											tmpSQLPKString = tmpSQLPKString & strCDp & arrGridData(arrstrSQLSelectPKPos(intSPFC), intRowCount - 1) & strCDp & " AND "
									End Select
								Else
									Select Case intPKPos
										Case 8, 129, 130, 200, 201, 202, 203										'adBSTR, adChar, adWChar, adVarChar, adLongVarChar, adVarWChar, adLongVarWChar		
											tmpSQLPKString = tmpSQLPKString & strCDp & Replace(ADORecordset.Fields(arrSQLSelectPK(intSPFC)).Value, "'", "''") & strCDp & " AND "
										Case 135, 7, 133, 134										'adDBTimeStamp, adDate, adDBDate,  adDBTime
											tmpSQLPKString = tmpSQLPKString & strCDp & FormatForSQL(CStr(ADORecordset.Fields(arrSQLSelectPK(intSPFC)).Value), DatabaseType, "SafeDate") & strCDp & " AND "
										Case Else
											tmpSQLPKString = tmpSQLPKString & strCDp & ADORecordset.Fields(arrSQLSelectPK(intSPFC)).Value & strCDp & " AND "
									End Select
								End If
							Next
							If tmpSQLPKString <> "" Then
								tmpSQLPKString = Left(tmpSQLPKString, Len(tmpSQLPKString) - 5)
								strPKqstTemp = strPKqstTemp & Server.URLEncode(tmpSQLPKString)
							End If
						End If
						strPKqstTemp = strPKqstTemp & strInfqstTemp
						If (AllowEdit Or AllowDelete) And (intSQLSelectIDPos <> -1 Or ((SQLSelectPK <> ""))) Then
							Response.Write("<a href=""" & FormEdit & "?" & strPKqstTemp & "&"" " & HTMLAttribEditLink & ">" & HTMLEditLink & "</a>")
						End If
						If AllowDetail And AllowEdit Then
							Response.Write("&nbsp;")
						End If
						If AllowDetail And (intSQLSelectIDPos <> -1 Or ((SQLSelectPK <> ""))) Then
							Response.Write("<a href=""" & FormDetail & "?" & strPKqstTemp & """>" & HTMLDetailLink & "</a>")
						End If
						If AllowMultiSelect Then
							Response.Write("<input type=""checkbox"" name=""SQLID"" value=""" & Server.URLEncode(arrGridData(intSQLSelectIDPos, intRowCount - 1)) & """>")
						End If
					End If
					If HTMLGridVertical Then
						Response.Write("</td></tr>")
					Else
						If Not UCase(GridMode) = "EXPORT" Then
							Response.Write("</td>")
						End If
					End If
					'Start Build Template column
					If HTMLTemplateColumn <> "" Then
						If HTMLGridVertical Then
							Response.Write("<td colspan=""2"">")
						Else
							Response.Write("<td>")
						End If
						strRTemp = HTMLTemplateColumn
						Do While instr(strRTemp,"{%") > 0
							strRTempStart = left(strRTemp,instr( strRTemp,"{%")-1)
							strRTempEnd = mid(strRTemp,instr(instr(strRTemp,"{%")+2, strRTemp,"%}")+2)
							strToEval = mid(strRTemp,instr(strRTemp,"{%")+2, instr(instr(strRTemp,"{%")+2, strRTemp,"%}") - instr(strRTemp,"{%")-2)
							varEval = Null
							intcurrentfield = 0
							For Each fldTemp In arrSQLSelect
								If UCase(fldTemp) = UCase(strToEval) Then
									If GridUseGetRows Then
										varEval = arrGridData(intCurrentField, intRowCount - 1)
									Else
										varEval = ADORecordset.Fields(intCurrentField).Value
									End If
									Exit For
								End If
								intcurrentfield = intcurrentfield + 1
							Next
							If err.number <> 0 Then
								strRTemp = strRTempStart & "#ERR" &  strRtempEnd
								err.clear
							Else
								If Not Isnull(varEval) Then
									strRTemp = strRTempStart & varEval & strRtempEnd
								Else
									strRTemp = strRTempStart & strRtempEnd
								End If
							End if
							strRTempEnd = ""
							strRTempStart = ""
						Loop
						Response.write(strRTemp)
						If HTMLGridVertical Then
							Response.Write("</td></tr>")
						Else
							Response.Write("</td>")
						End If
					End If
					'End Build Template Column
					intCurrentField = 0
					'response.write sqlselect
					
					If SQLSelectName = "" Then
						strFName = arrSQLSelect(intCurrentField)
					Else
						strFName = arrSQLSelectName(intCurrentField)
					End If
					For Each fldTemp In arrSQLSelect
						blnTemp = True
						If SQLSelectHide <> "" Then
							For eleTemp = 0 To UBound(arrSQLSelectHide)
								If UCase(ArrSQLSelect(intCurrentField)) = UCase(arrSQLSelectHide(eleTemp)) Then
									blnTemp = False
									Exit For
								End If
							Next
						End If
						If blnTemp Then
							If GridUseGetRows Then
								strName = ArrSQLSelect(intCurrentField)
								varValue = arrGridData(intCurrentField, intRowCount - 1)
								intType = arrSQLSelectType(intCurrentField)
							Else
								strName = ArrSQLSelect(intCurrentField)
								varValue = ADORecordset.Fields(intCurrentField).Value
								intType = arrSQLSelectType(intCurrentField)
							End If
						End If
						'display field value
						If strName <> "" Then
							If intSQLSelectIDPos = -1 Then
								intSQLSelectIDPos = -1
							End If
							If (CInt(intCurrentField) = CInt(intSQLSelectIDPos)) And GridHideAutonumber Then
							Else
								If HTMLGridVertical Then
									Response.Write("<tr>")
									If SQLSelectName = "" Then
										Call DisplayFieldName(arrSQLSelect(intCurrentField), arrSQLSelectType(intCurrentField), 50, arrSQLSelect(intCurrentField))
									Else
										Call DisplayFieldName(arrSQLSelect(intCurrentField), arrSQLSelectType(intCurrentField), 50, arrSQLSelectName(intCurrentField))
									End If
									Response.Write("<td valign=""top"">")
								Else
									Response.Write("<td " )
									If HTMLAttribGridCellAlign = "" Then
										Select Case intType
											Case 11										'adBoolean
												Response.Write("align=""center""")
											Case 2, 3, 4, 5, 14, 16, 17, 18, 19, 20, 21, 128, 131, 204, 6										'adSmallInt, adInteger, adSingle, adDouble, adDecimal, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adBigInt, adUnsignedBigInt, adBinary, adNumeric, adVarBinary, adLongVarBinary, adCurrency, 
												Response.Write("align=""right""")
											Case Else
												Response.Write("align=""left""")
										End Select
									End If
									Response.Write(" " & HTMLAttribGridCell & ">")
								End If
								Select Case intType								  'Check For fields that can't be displayed
									Case 128, 204, 205									 'adBinary, adVarBinary, adLongVarBinary  Then Don't touch binary fields
										Response.Write(HTMLBinaryValue)
									Case Else									 'retrieve value and check For null
										If IsNull(varValue) Then
											Response.Write(HTMLNullValue)
										ElseIf varValue = "" Then
											Response.Write(HTMLNullValue)
										ElseIf VarType(varValue) = vbNull Then
											Response.Write(HTMLNullValue)
										Else
											Select Case intType											  'Check For fields that don't support drilldown
												Case 11												 'adBoolean
													If HTMLTrueValue <> "" and UCase(GridMode) <> "EXPORT" Then
														If varValue Then
															Response.Write(HTMLTrueValue)
														Else
															Response.Write(HTMLFalseValue)
														End If
													Else
														Response.Write(varValue)
													End If
												Case 203, 201, 135, 7, 133, 134, 72												 'adLongVarChar 

													If UCase(GridMode) = "EXPORT" And UCase(Request.QueryString("ocdExportFormat_" & GRIDID)) <> "PRINT" And IsNumeric(varValue) Then
														Response.Write("=""" & (Replace(Server.HTMLEncode(varValue),"""","""""")) & """")
ElseIf SQLSelectFormat = "" Then
														If UCase(GridMode) = "EXPORT" And ExportLineBreaks Then
																Response.Write((Replace(Server.HTMLEncode(varValue),vbCRLF,"<br>")))
																'Response.Write((Replace(Server.HTMLEncode(varValue),vbCRLF,"&#10;")))
																
														Else
															If Len(varValue) > GridMaxMemo Then
																If RenderAsHTML Then
Response.Write(Left(varValue, GridMaxMemo) & HTMLMemoContinues)

																Else
																	 Response.Write(Server.HTMLEncode(Left(varValue, GridMaxMemo)) & HTMLMemoContinues)
																End If
															Else
																If RenderAsHTML Then
Response.Write(varValue)

																Else
																	Response.Write(Server.HTMLEncode(varValue))
																End If
															End If
														End if
													Else
														Select Case UCase(arrSQLSelectFormat(intCurrentField))
															Case "LINEBREAKS"
																Response.Write(Replace(Server.HTMLEncode(varValue), vbCRLF, "<br>"))
															Case "HYPERLINK"
																arrHyperlink = Split(varValue, "#")
																If arrHyperLink(0) = "" Then
																	Response.Write("<a href=""" & arrHyperLink(1) & """>" & arrHyperLink(1) & "</a>")
																Else
																	Response.Write("<a href=""" & arrHyperLink(1) & """>" & arrHyperLink(0) & "</a>")
																End If
															Case "HTML"
																Response.Write(varValue)
															Case Else
															If Len(varValue) > GridMaxMemo Then
																If RenderAsHTML Then
Response.Write(Left(varValue, GridMaxMemo) & HTMLMemoContinues)

																Else
																	 Response.Write(Server.HTMLEncode(Left(varValue, GridMaxMemo)) & HTMLMemoContinues)
																End If
															Else
																If RenderAsHTML Then
Response.Write(varValue)

																Else
																	Response.Write(Server.HTMLEncode(varValue))
																End If
															End If
														End Select
													End If
												Case Else												 
													'Fields eligible For drilldown
													If UCase(GridMode) = "BROWSE" And SQLRecordCount > 1 And (UCase(Request.QueryString("drilldown" & "_" & GridID)) = "YES" And UCase(strName) <> UCase(intSQLSelectIDPos)) Then
														' These Set the drill down values 
														Response.Write("<a href=""" & strSCRIPT_NAME)
														Select Case intType
															Case 2, 3, 4, 5, 14, 16, 17, 18, 19, 20, 21, 128, 131, 204, 6, 11															 'adSmallInt, adInteger, adSingle, adDouble, adDecimal, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adBigInt, adUnsignedBigInt, adBinary, adNumeric, adVarBinary, adLongVarBinary, adCurrency, adBoolean
																strCD = ""
															Case 135, 7, 133, 134															 'adDBTimeStamp, adDate, adDBDate,  adDBTime
																If DatabaseType = "Access" Then
																	strCD = "#"
																Else
																	strCD = "'"
																End If
															Case 8, 129, 130, 200, 201, 202, 203															 'adBSTR, adChar, adWChar, adVarChar, adLongVarChar, adVarWChar, adLongVarWChar						
																strCD = "'"
														End Select
														Select Case intType
															Case 135, 7, 133, 134															 'adDBTimeStamp, adDate, adDBDate,  adDBTime
																If SQLWhere = "" Then
																	Response.Write("?sqlwhere" & "_" & GridID & "=" & Server.URLEncode(FormatForSQL(strName, DatabaseType, "AddSQLIdentifier") & " = " & strCD & FormatForSQL(CStr(varValue), DatabaseType, "SafeDate") & strCD))
																Else
																	Response.Write("?sqlwhere" & "_" & GridID & "=" & Server.URLEncode("(" & SQLWhere & ") AND (" & FormatForSQL(strName, DatabaseType, "AddSQLIdentifier") & " = " & strCD & FormatForSQL(CStr(varValue), DatabaseType, "SafeDate") & strCD & ")"))
																End If
															Case Else
																If SQLWhere = "" Then
																	Response.Write("?sqlwhere" & "_" & GridID & "=" & Server.URLEncode(FormatForSQL(strName, DatabaseType, "AddSQLIdentifier") & " = " & strCD & Replace(varValue, "'", "''") & strCD))
																Else
																	Response.Write("?sqlwhere" & "_" & GridID & "=" & Server.URLEncode("(" & SQLWhere & ") AND (" & FormatForSQL(strName, DatabaseType, "AddSQLIdentifier") & " = " & strCD & Replace(varValue, "'", "''") & strCD & ")"))
																End If
														End Select
														For Each qstTemp In Request.QueryString
															If UCase(qstTemp) <> "SQLWHERE" & "_" & GridID And UCase(qstTemp) <> "OCDACTION" & "_" & GridID Then
																Response.Write("&amp;" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp)))
															End If
														Next
														Response.Write(""">")
													End If
													Select Case intType
														Case 6														  'adCurrency 
															Response.Write(Server.HTMLEncode(FormatCurrency(varValue)))
														Case 8, 129, 130, 200, 201, 202, 203														  'adBSTR, adChar, adWChar, adVarChar, adLongVarChar, adVarWChar, adLongVarWChar	

													If UCase(GridMode) = "EXPORT" And IsNumeric(varValue) And UCase(Request.QueryString("ocdExportFormat_" & GRIDID)) <> "PRINT" Then
														Response.Write("=""" & (Replace(Server.HTMLEncode(varValue),"""","""""")) & """")

															ElseIf SQLSelectFormat <> "" Then
																Select Case UCase(arrSQLSelectFormat(intCurrentField))
																	Case "HTML"
																		Response.Write(varValue)
																	Case Else
																If RenderAsHTML Then
																	Response.Write(varValue)
																Else
																		Response.Write(Server.HTMLEncode(varValue))
																End If
																End Select
															Else
																If RenderAsHTML Then
																	Response.Write(varValue)

																Else
																	Response.Write(Server.HTMLEncode(varValue))
																End If
															End If
														Case 2, 3, 4, 5, 14, 16, 17, 18, 19, 20, 21, 128														 'adSmallInt, adInteger, adSingle, adDouble, adDecimal, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adBigInt, adUnsignedBigInt, adBinary, adNumeric, 
															If SQLSelectFormat <> "" Then
																Select Case UCase(arrSQLSelectFormat(intCurrentField))
																	Case "PERCENT"
																		Response.Write(FormatPercent(varValue, 0, 0, 0, 0))
																	Case "COMMAS"
																		Response.Write(FormatNumber(varValue, 0, 0, 0, -1))
																	Case Else
																		Response.Write(Server.HTMLEncode(varValue))
																End Select
															ElseIf UCase(GridFormatNumbers) = "COMMAS" Then
																Response.Write(FormatNumber(varValue, 0, 0, 0, -1))
															Else
																Response.Write(Server.HTMLEncode(varValue))
															End If
														Case Else
															Response.Write(Server.HTMLEncode(varValue))
													End Select
													If (Request.QueryString("drilldown" & "_" & GridID) = "yes" And UCase(strName) <> UCase(intSQLSelectIDPos)) Then
														Response.Write("</a>")
													End If
											End Select
										End If
								End Select
								Response.Write("</td>")
								If HTMLGridVertical Then
									Response.Write("</tr>")
								End If
							End If
							'end display field value
							strName = ""
						End If
						intCurrentField = intCurrentField + 1
					Next
					If UCase(GRIDMODE) <> "Export" Then
					Response.Write("<td>")
					If AllowDelete And (intSQLSelectIDPos <> -1 Or ((SQLSelectPK <> ""))) Then
						tmpSQLPKString = ""
						Response.Write("<a href=""" & FormDelete & "?" & strPKqstTemp & "&amp;ocdEditDelete=delete"" " & HTMLAttribDeleteLink & " >" & HTMLDeleteLink & "</a>")
						If AllowMultiDelete Then
							Response.Write("<input type=""checkbox"" name=""sqlid"" value=""" & Server.URLEncode(arrGridData(intSQLSelectIDPos, intRowCount - 1)) & """>")
						End If
					Else
						Response.Write("&nbsp;")
					End If
					Response.Write("</td>")
					End If
					Response.Write("</tr>")
					If Not Response.IsClientConnected Then
						Call Close()
						Response.Clear()
						Response.End()
					End If
					If Not GridUseGetRows Then
						ADORecordset.MoveNext()
						If ADORecordset.Eof Then
							Exit For
						End If
					End If
				Next
				Response.Write("</table>")
				If AllowMultiDelete Or AllowMultiSelect Then
					Response.Write("</form>")
				End If
		End Select
		Select Case UCase(strTemplate)
			Case "QUICK", "", "TOTALS", "EXPORT"
				Response.Write(strCData)
				Select Case UCase(Request.QueryString("ocdExportFormat_" & GRIDID))
					Case "EXCEL"
					Response.end()
					Case "PRINT"
						Response.Write("</body></html>")
				End Select
		End Select
		Select Case UCase(strTemplate)
			Case "QUICK", ""
				If UCase(GridMode) = "EXPORT" Then
					'	Response.end
				End If
		End Select
		Select Case UCase(strTemplate)
			Case "KEYWORD"
				If ShowKeywordSearch Then
					Call DisplayKeywordSearch()
				End If
		End Select
	End Sub

	Public Sub Close()
		On Error Resume Next
		ADORecordset.close()
		Set ADORecordSet = Nothing
		ADOConnection.close()
		Set ADOConnection = Nothing
		Err.clear()
	End Sub

	Private Sub DisplayFieldName(ByVal strFieldName, ByVal intFieldType, ByVal intFieldSize, ByVal strDisplayName)
		'Create grid table header cell
		Dim qstTemp, blnTemp, eleTemp, strTemp
		If strFieldName = intSQLSelectIDPos And GridHideAutonumber = True Then
			Exit Sub
		End If
		If HTMLGridVertical Then
			Response.Write("<td valign=""top"" nowrap align=""left""><span class=""FieldName"">")
		Else
			Response.Write("<th valign=""top"" nowrap align=""left"" class=""GridHeader"">")
		End If
		Select Case intFieldType
			Case 128, 204, 205			'adBinary, adVarBinary, adLongVarBinary 
				If strDisplayName = "" Then
					Response.Write(Replace(strFieldName, " ", "&nbsp;"))
				Else
					Response.Write(strDisplayName)
				End If
				Response.Write(HTMLAfterFieldName)
			Case Else
				' Make sorting links
				Response.Write("")
				If SQLSelectName = "" Then
					Response.Write(Replace(strFieldName, " ", "&nbsp;"))
				Else
					Response.Write(ArrSQLSelectName(intCurrentField))
				End If
				Response.Write(HTMLAfterFieldName)
				If HTMLSortASCLink <> "" And UCase(GridMode) <> "EXPORT" Then
					blnTemp = True
					If SQLSelectSortHide <> "" Then
						For Each eleTemp In arrSQLSelectSortHide
							If UCase(strFieldName) = UCase(eleTemp) Then
								blnTemp = False
								Exit For
							End If
						Next
					End If
					If blnTemp Then
						strTemp = ""
						For Each qstTemp In Request.QueryString
							If UCase(qstTemp) <> "SQLORDERBY" & "_" & GridID And UCase(qstTemp) <> "OCDACTION" & "_" & GridID Then
								strTemp = strTemp & "&amp;" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp))
							End If
						Next
						Select Case intFieldType
							Case 128, 205, 204, 201, 203							'adLongVarChar
								'don't build sort links For a memo or binary
							Case Else
								'Ascending
								Response.Write("<a href=""" & strSCRIPT_NAME & "?sqlorderby" & "_" & GridID & "=")
								If SQLOrderBy = "" Or Not SearchMultiSort Then
									Response.Write(Server.URLEncode(FormatForSQL(strFieldName, DatabaseType, "AddSQLIdentifier") & " ASC"))
								Else
									Response.Write(Server.URLEncode(SQLOrderBy) & ",+" & Server.URLEncode(FormatForSQL(strFieldName, DatabaseType, "AddSQLIdentifier") & " ASC"))
								End If
								Response.Write(strTemp & """>" & HTMLSortASCLink & "</a>")
								'Create descending sort link
								Response.Write("&nbsp;<a href=""" & strSCRIPT_NAME & "?sqlorderby" & "_" & GridID & "=")
								If SQLOrderBy = "" Or Not SearchMultiSort Then
									Response.Write(Server.URLEncode(FormatForSQL(strFieldName, DatabaseType, "AddSQLIdentifier") & " DESC"))
								Else
									Response.Write(Server.URLEncode(SQLOrderBy) & ",+" & Server.URLEncode(FormatForSQL(strFieldName, DatabaseType, "AddSQLIdentifier") & " DESC"))
								End If
								Response.Write(strTemp & """>" & HTMLSortDESCLink & "</a>")
						End Select
					End If
				End If
				If HTMLFilterLink <> "" And UCase(GridMode) <> "EXPORT" Then
					blnTemp = True
					If SQLSelectFilterHide <> "" Then
						For Each eleTemp In arrSQLSelectFilterHide
							If UCase(strFieldName) = UCase(eleTemp) Then
								blnTemp = False
								Exit For
							End If
						Next
					End If
					If blnTemp Then
						Response.Write("&nbsp;<a href=""" & strSCRIPT_NAME & "?ocdGridMode_" & GRIDID & "=Filter&amp;ocdFilterFieldName_" & GRIDID & "=" & Server.URLEncode(strFieldName))
						If SQLSelectName <> "" Then
							Response.Write("&amp;ocdFilterFieldDisplayName_" & GRIDID & "=" & Server.URLEncode(ArrSQLSelectName(intCurrentField)))
						End If
						Response.Write("&amp;ocdFilterFieldtype_" & GRIDID & "=" & Server.URLEncode(intFieldType) & "&amp;ocdFilterFieldSize_" & GRIDID & "=" & Server.URLEncode(intFieldSize))
						For Each qstTemp In Request.QueryString
							Select Case UCase(qstTemp)
								Case "GRIDID", UCase("OCDGRIDMODE_" & GridID), UCase("OCDFILTERFIELDNAME_" & GridID), UCase("OCDFILTERFIELDDISPLAYNAME_" & GridID), "OCDACTION_" & GridID, "OCDFILTERFIELDTYPE_" & GridID, "OCDFILTERFIELDSIZE_" & GridID, "DATABASETYPE_" & GridID, "SCRIPT_" & GridID
								Case Else
									Response.Write("&amp;" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp)))
							End Select
						Next
						Response.Write(""">" & HTMLFilterLink & "</a>")
					End If
				End If
				Response.Write(" ")
		End Select
		If ShowDescription And DatabaseType = "Access" And UCase(GridMode) <> "EXPORT" Then
			Response.Write("<A HREF="""" onclick=""javascript:window.open('DescribeField.asp?sqlfrom=" & Server.URLEncode(SQLFrom) & "&amp;sqlfield=" & Server.URLEncode(strFieldName) & "', 'describe','height=200,width=300,scrollbars=yes');return false""><IMG ALT=""Describe Field"" SRC=""appHelpSmall.gif"" Border=0></A>&nbsp;")
		End If
		If HTMLGridVertical Then
			Response.Write("</span></td>")
		Else
			Response.Write("</th>")
		End If
	End Sub

	Public Sub DisplayKeywordSearch()
		'Build HTML FORM for submitting keywords
		Dim qstTemp,blnTemp, arrSQLSelectSearchHide, eleSQLSelectSearchHide,strName, arrSQLSelectHideFields, eleSQLSelectHideFields
		

		Response.Write("<form method=""Post"" action=""" & strSCRIPT_NAME & "?ocdGridMode_" & GRIDID & "=Process&amp;")
		For Each qstTemp In Request.QueryString
			If UCase(qstTemp) <> UCase(("OCDGRIDMODE_" & GRIDID)) And UCase(qstTemp) <> UCase("OCDFILTERFIELD_" & GRIDID) And UCase(qstTemp) <> UCase("OCDACTION_" & GRIDID) And UCase(qstTemp) <> UCase("OCDKEYWORD_" & GRIDID) Then
				Response.Write(qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp)) & "&amp;")
			End If
		Next
		Response.Write(""">")
		Response.Write("<table border=""0""><tr>" & "<td nowrap>")
		Response.write("</td><td nowrap>")
		Response.Write("<input name=""ocdSearchPageSize_" & SQLPageSize & """ value=""NEITHER"" type=""hidden"">")
		Response.Write("<input name=""ocdSearchExistslogic_" & GRIDID & """ value=""NEITHER"" type=""hidden"">")
		Response.Write("<input name=""ocdSearchlogic_" & GRIDID & """ value=""AND"" type=""hidden"">")
		If Request.QueryString("sqlselecthide_" & GRIDID) <> "" Then
			arrSQLSelectHideFields = Split(CStr(Request.QueryString("sqlselecthide_" & GRIDID)), ",")
		End If
		If SQLSelectSearchUncheck <> "" Then
			arrSQLSelectSearchUncheck = Split(SQLSelectSearchUncheck, ",")
		End If
		If SQLSelectSearchHide <> "" Then
			arrSQLSelectSearchHide = Split(SQLSelectSearchHide, ",")
		End If
		If SQLSelectName <> "" Then
			ArrSQLSelectName = Split(SQLSelectName, ",")
		End If
			for each strName in arrSQLSelect
			blnTemp = False
			Response.Write("<input type=""hidden"" name=""ocdSearchshowchk_" & GRIDID & Server.htmlencode(strName) & """ value=""")
			If Request.QueryString("sqlselecthide_" & GRIDID) = "" Then
					If SearchCheckAll Then
						If SQLSelectSearchUncheck = "" Then
							Response.Write("on")
						Else
							blnTemp = True
							For Each eleshowcheck In arrSQLSelectSearchUncheck
								If UCase(eleshowcheck) = UCase(strName) Then
									blnTemp = False
								End If
							Next
							If blnTemp Then
								Response.Write("on")
							End If
						End If
					End If
				Else
					blnTemp = False
					For Each eleSQLSelectHideFields In arrSQLSelectHideFields
						If UCase(eleSQLSelectHideFields) = UCase(strName) Then
							blnTemp = True
						End If
					Next
					If Not blnTemp Then
						Response.Write("on")
					End If
				End If
				Response.Write(""" > <input type=""hidden"" name=""ocdSearchexischk_" & GRIDID & Server.htmlencode(strName) & """ value=""Exists"" >")
		next
		Response.Write("<input name=""ocdKeyword_" & GRIDID & """ value=""" & Server.htmlencode(Request.QueryString("ocdKeyword_" & GRIDID)) & """ size=""12"" maxlength=""255""> " & "<input name=""ocdFind_" & GRIDID & """ type=""submit"" value=""search"" class=""submit"">" & "<input name=""ocdFindBrowse"" type=""hidden"" value=""yes"">")
		Response.Write("</td><td>")
		Response.Write("</td></tr></table>")
		Response.Write("</form>")
	End Sub

	Private Sub DisplayGridButtons()
		Dim qstTemp, arrShowGridButtons, eleShowGridButtons, tmpExportqstTemp, tmpstrPageLinkqstTemp, arrGRIDBE, intNewStart, intNewSize
		If HTMLGridButtons = "" Then
			Exit Sub
		End If
		If FormSearch = "" Then
			FormSearch = strSCRIPT_NAME
		End If
		tmpstrPageLinkqstTemp = ""
		If Request.QueryString("SQLSelect_" & GRIDID) <> "" Then
			tmpExportqstTemp = "&amp;sqlselect_" & GridID & "=" & Server.URLEncode(Request.QueryString("SQLSelect_" & GRIDID)) 
		End If
		If Request.QueryString("SQLFrom_" & GRIDID) <> "" Then
			tmpExportqstTemp = tmpExportqstTemp & "&amp;sqlfrom_" & GridID & "=" & Server.URLEncode(Request.QueryString("SQLFrom_" & GRIDID)) 
		End If
		If Request.QueryString("SQLWhere_" & GRIDID) <> "" Then
			tmpExportqstTemp = tmpExportqstTemp & "&amp;sqlwhere_" & GridID & "=" & Server.URLEncode(Request.QueryString("SQLWhere_" & GRIDID)) 
		End If
		If Request.QueryString("SQLOrderBy_" & GRIDID) <> "" Then
			tmpExportqstTemp = tmpExportqstTemp & "&amp;sqlorderby_" & GridID & "=" & Server.URLEncode(Request.QueryString("SQLOrderBy_" & GRIDID)) 
		End If
		If Request.QueryString("SQLSelectHide_" & GRIDID) <> "" Then
		tmpExportqstTemp = tmpExportqstTemp & "&amp;sqlselecthide_" & GridID & "=" & Server.URLEncode(Request.QueryString("SQLSelectHide_" & GRIDID))
		End If
		For Each qstTemp In Request.QueryString
			Select Case UCase(qstTemp)
				Case UCase("SQLWHERE_" & GridID), UCase("SQLPAGE_" & GridID), UCase("SQLORDERBY_" & GridID), UCase("SQLSELECT_" & GridID), UCase("OCDACTION_" & GridID), UCase("SQLFROM_" & GridID), UCase("SQLSELECTHIDE_" & GridID)
				Case Else
					tmpExportqstTemp = tmpExportqstTemp & "&amp;" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp))
			End Select
		Next
		For Each qstTemp In Request.QueryString
			Select Case UCase(qstTemp)
				Case UCase("SQLPAGE_" & GridID)
				Case Else
					tmpstrPageLinkqstTemp = tmpstrPageLinkqstTemp & "&amp;" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp))
			End Select
		Next
		arrShowGRIDButtons = Split(HTMLGridButtons, ";;")
		Response.Write("<table " & HTMLAttribButtonPanel & "><tr>")
		For Each eleShowGridButtons In arrShowGridButtons
			Response.Write("<td align=""center"" valign=""top"" nowrap>")
			arrGRIDBE = Split(eleShowGridButtons, "|")
			Response.Write(" ")
			Select Case UCase(arrGRIDBE(0))
				Case "SEARCH"
					Response.Write(" <a href=""" & FormSearch & "?ocdGridMode_" & GRIDID & "=Search")
					For Each qstTemp In Request.QueryString
						If UCase(qstTemp) <> ("OCDGRIDMODE_" & UCase(GRIDID)) And UCase(qstTemp) <> "OCDACTION_" & GridID Then
							Response.Write("&amp;" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp)))
						End If
					Next
					Response.Write(""">")
					If UBound(arrGridBE) > 0 Then
						Response.Write(arrGridBE(1))
					Else
						Response.Write(arrGridBE(0))
					End If
					Response.Write("</a> ")
				Case "DRILLDOWN"
					If UCase(Request.QueryString("DrillDown" & "_" & GridID)) = "YES" Then
						Response.Write(" <a href=""" & strSCRIPT_NAME & "?DrillDown" & "_" & GridID & "=")
						For Each qstTemp In Request.QueryString
							If UCase(qstTemp) <> "OCDACTION_" & GridID And UCase(qstTemp) <> "DRILLDOWN_" & GridID Then
								Response.Write("&amp;" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp)))
							End If
						Next
						Response.Write(""">" & Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1) & "</a>")
					Else
						Response.Write("<a href=""" & strSCRIPT_NAME & "?DrillDown_" & GridID & "=yes")
						For Each qstTemp In Request.QueryString
							If UCase(qstTemp) <> "OCDACTION_" & GridID And UCase(qstTemp) <> "DRILLDOWN_" & GridID Then
								Response.Write("&amp;" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp)))
							End If
						Next
						Response.Write("""> " & Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1) & "</a>")
					End If
				Case "RESET"
					If SQLOrderBy <> "" Or SQLWhere <> "" Then
						Response.Write(" <a href=""" & FormSearch & "?sqlorderby_" & GRIDID & "=&amp;sqlwhere_" & GRIDID & "=")
						For Each qstTemp In Request.QueryString
							If UCase(qstTemp) <> "SQLORDERBY_" & GRIDID And UCase(qstTemp) <> "SQLWHERE_" & GridID And UCase(qstTemp) <> "OCDACTION_" & GridID Then
								Response.Write("&amp;" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp)))
							End If
						Next
						Response.Write("""> ")
					End If
					Response.Write(Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1))
					If SQLOrderBy <> "" Or Request.QueryString("SQLWHERE_" & GRIDID) <> "" Then
						Response.Write("</a>")
					End If
				Case "FIRST"
					If SQLPage > 1 Then
						Response.Write(" <a href=""" & strSCRIPT_NAME & "?sqlpage" & "_" & GridID & "=" & "1" & tmpstrPageLinkqstTemp & """>")
					End If
					Response.Write(Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1))
					If SQLPage > 1 Then
						Response.Write("</a>")
					End If
					Response.Write(" ")
				Case "PREV"
					Response.Write(" ")
					If SQLPage > 1 Then
						Response.Write("<a href=""" & strSCRIPT_NAME & "?sqlpage_" & GridID & "=" & CStr(SQLPage - 1) & tmpstrPageLinkqstTemp & """>")
					End If
					Response.Write(Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1))
					If SQLPage > 1 Then
						Response.Write("</a>")
					End If
				Case "CUSTOM"
					Response.Write(Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1))
				Case "NEXT"
					If SQLPage < SQLPageCount Then
						Response.Write(" <a href=""" & strSCRIPT_NAME)
						Response.Write("?sqlpage_" & GridID & "=" & CStr(SQLPage + 1) & tmpstrPageLinkqstTemp & """>")
					End If
					Response.Write(Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1))
					If SQLPage < SQLPageCount Then
						Response.Write("</a> ")
					End If
				Case "LAST"
					If SQLPage < SQLPageCount Then
						Response.Write(" <a href=""" & strSCRIPT_NAME & "?sqlpage_" & GridID & "=" & CStr(SQLPageCount) & tmpstrPageLinkqstTemp & """>")
					End If
					Response.Write(Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1))
					If SQLPage < SQLPageCount Then
						Response.Write("</a>")
					End If
				Case "PRINT"
					Response.Write(" <a href=""" & strSCRIPT_NAME & "?ocdGridMode_" & GRIDID & "=Export&amp;ocdExportFormat_" & GRIDID & "=PRINT" & tmpExportqstTemp & """>" & Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1) & "</a> ")
				Case "PAGING"
					If UCase((Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1))) = "SMBUTTON" Then
						Response.Write(" <table width=""26"" border=""1"" cellspacing=""0"" cellpadding=""2""><tr><td height=""12"" nowrap align=""center"" valign=""middle""><small>")
					ElseIf UCase((Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1))) = "BUTTON" Then
						Response.Write(" <table width=""26"" border=""1"" cellspacing=""0"" cellpadding=""3""><tr><td height=""12"" nowrap align=""center"" valign=""middle""><small>")
					End If
					Response.Write(HTMLPagingStart)
					If IsNumeric(SQLRecordCount) Then
						If SQLRecordCount > 0 Then
							Response.Write(" ")
							Response.Write(SQLPageSize * (SQLPage - 1) + 1)
							If SQLPageSize > 1 Then
								If GridMaxRecordsDisplay > 4 Then
		  						Response.Write(" <a href=""" & strSCRIPT_NAME & "?sqlpagesize" & "_" & GridID & "=")
									If SQLPageSize < 11 And GridMaxRecordsDisplay > 4Then
										intNewSize =5
									ElseIf SQLPageSize < 21 And GridMaxRecordsDisplay > 9 Then
										intNewSize = 10
									ElseIf SQLPageSize < 51 And GridMaxRecordsDisplay > 19 Then
										intNewSize = 20
									ElseIf SQLPageSize < 101 And GridMaxRecordsDisplay > 49 Then
										intNewSize = 50
									ElseIf SQLPageSize < 251 And GridMaxRecordsDisplay > 99 Then
										intNewSize = 100
									ElseIf SQLPageSize < 501 And GridMaxRecordsDisplay > 249 Then
										intNewSize = 250
									ElseIf SQLPageSize < 1001 And GridMaxRecordsDisplay > 499 Then
										intNewSize = 500
									ElseIf SQLPageSize > 1000 Then
										intNewSize = 1000
									End If
									Response.Write(intNewSize)
									If SQLPage = 1 Then
										intNewStart = 1
									Else
										If intNewSize = 0 Then
											intNewSize = 1
										End If
										intNewStart = Clng(SQLPageSize * (SQLPage - 1) + 1) \ CLng(intNewSize) + 1
									End If
									Response.Write("&amp;sqlpage_" & GRIDID & "=" & Cstr((intNewStart)))
									For Each qstTemp In Request.QueryString
										If UCase(qstTemp) <> "OCDACTION_" & GridID And UCase(qstTemp) <> "SQLPAGESIZE_" & GridID And UCase(qstTemp) <> "SQLPAGE_" & GridID Then
											Response.Write("&amp;" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp)))
										End If
									Next
									Response.Write(""" style=""text-decoration:none;"">")
									Response.Write("<small>-</small>")
									Response.Write("</a> ")
								Else
									Response.Write(" - ")
								End If
								If SQLRecordCount < (SQLPageSize * (SQLPage)) Then
									Response.Write(SQLRecordCount)
								Else
									Response.Write(SQLPageSize * (SQLPage))
								End If
								If GridMaxRecordsDisplay > 4  And GridMaxRecordsDisplay > SQLPageSize Then
  								Response.Write(" <a href=""" & strSCRIPT_NAME & "?sqlpagesize" & "_" & GridID & "=")
									If SQLPageSize < 9 And GridMaxRecordsDisplay > 9 Then
										intNewSize = 10
									ElseIf SQLPageSize < 19 And GridMaxRecordsDisplay > 19 Then
										intNewSize = 20
									ElseIf SQLPageSize < 49 And GridMaxRecordsDisplay > 49 Then
										intNewSize = 50
									ElseIf SQLPageSize < 99 And GridMaxRecordsDisplay > 99 Then
										intNewSize = 100
									ElseIf SQLPageSize < 249 And GridMaxRecordsDisplay > 249 Then
										intNewSize = 250
									ElseIf SQLPageSize < 499 And GridMaxRecordsDisplay > 499 Then
										intNewSize = 500
									ElseIf SQLPageSize < 999 And GridMaxRecordsDisplay > 999 Then
										intNewSize = 1000
									Else
										intNewSize  = GridMaxRecordsDisplay 
									End If
									Response.Write(intNewSize)
									If SQLPage = 1 Then
										intNewStart = 1
									Else
										If intNewSize = 0 Then
											intNewSize = 1
										End If
										intNewStart = Clng(SQLPageSize * (SQLPage - 1) + 1) \ CLng(intNewSize) + 1
									End If
									Response.Write("&amp;sqlpage_" & GRIDID & "=" & Cstr((intNewStart)))
									For Each qstTemp In Request.QueryString
										If UCase(qstTemp) <> "OCDACTION_" & GridID And UCase(qstTemp) <> "SQLPAGESIZE_" & GridID And UCase(qstTemp) <> "SQLPAGE_" & GridID Then
											Response.Write("&amp;" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp)))
										End If
									Next
									Response.Write(""" style=""text-decoration:none;"">")
									Response.Write("<small>+</small>")
									Response.Write("</a> ")
								End If
							End If
							Response.Write(" : ")
							If Not GRIDUseGetRows And SQLRecordCount > ADOMaxRecords Then
							Response.Write(ADOMaxRecords)
							Response.Write("+")
							Else
								Response.Write(SQLRecordCount)
							End If
						Else
							Response.Write("0 : 0")
						End If
					Else
						Response.Write(" ")
						Response.Write((SQLPageSize * (SQLPage - 1)) + 1)
						Response.Write(" - ")
						Response.Write(SQLPageSize * (SQLPage))
						Response.Write(" : ")
						Response.Write(Server.htmlencode("???"))
					End If
					Response.Write(HTMLPagingEnd)
					If UCase((Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1))) = "BUTTON" Or UCase((Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1))) = "SMBUTTON" Then
						Response.Write("</small></td></tr></table> ")
					End If
				Case "TEXT"
					'Text export deprecated because of common behavior for IE browsers to 
					'not respect mime type handling for anything it thinks could be text/html
					Response.Write(" <a href=""" & strSCRIPT_NAME & "?ocdGridMode_" & GRIDID & "=Export&amp;ocdExportFormat_" & GRIDID & "=text" & tmpExportqstTemp & """ ")
					Response.Write(">" & Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1) & "</a> ")
				Case "XML"
					Select Case CStr(Request.servervariables("server_software"))
						Case "Microsoft-IIS/5.0", "Microsoft-IIS/5.1", "Microsoft-IIS/6.0"
							Response.Write(" <a href=""" & strSCRIPT_NAME & "?ocdGridMode_" & GRIDID & "=Export&amp;ocdExportFormat_" & GRIDID & "=xml" & tmpExportqstTemp & """ ")
							Response.Write(">" & Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1) & "</a> ")
					End Select
				Case "EXCEL"
					Response.Write(" <a href=""" & strSCRIPT_NAME & "?ocdGridMode_" & GridID & "=Export&amp;ocdExportFormat_" & GRIDID & "=excel" & tmpExportqstTemp & """ ")
					Response.Write(">" & Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1) & "</a> ")
				Case "NEW"
					If (intSQLSelectIDPos <> -1 Or SQLSelectPK <> "") And AllowAdd Then
						If UCase(FormEdit) = "EDIT.ASP" Or UCase(FormEdit) = "FREEEDIT.ASP" Then
							Response.Write(" <a href=""" & FormEdit & "?sqlfrom=" & Server.URLEncode(sqlfrom))
						Else
							Response.Write(" <a href=""" & FormEdit & "?sqlid=")
						End If
						For Each qstTemp In Request.QueryString
							If UCase(qstTemp) <> "SQLWHERE" And UCase(qstTemp) <> "SQLID" And UCase(qstTemp) <> "OCDACTION_" & GridID And UCase(qstTemp) <> "SQLFROM" Then
								Response.Write("&amp;" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp)))
							End If
						Next
						Response.Write(""" " & HTMLAttribEditLink & "> " & Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1) & "</a>")
					End If
				Case "GRAPH"
					Response.Write(" <a href=""" & "GraphIt.asp" & "?Graphtype=3&amp;GridID=" & GRIDID & "&amp;" & "sqlselect_" & GridID & "=" & Server.URLEncode(SQLSelect) & "&amp;sqlfrom_" & GridID & "=" & Server.URLEncode(SQLFrom) & "&amp;sqlwhere_" & GridID & "=" & Server.URLEncode(SQLWhere) & "&amp;sqlorderby_" & GridID & "=" & Server.URLEncode(SQLOrderBy) & "&amp;sqlselecthide_" & GridID & "=" & Server.URLEncode(SQLSelectHide))
					For Each qstTemp In Request.QueryString
						Select Case UCase(qstTemp)
							Case "GRIDID", "SQLWHERE_" & GridID, "SQLORDERBY_" & GridID, "SQLSELECT_" & GridID, "OCDACTION_" & GridID, "SQLFROM_" & GridID, "SQLSELECTHIDE" & "_" & GridID
							Case Else
								Response.Write("&amp;" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp)))
						End Select
					Next
					Response.Write("""> " & Mid(eleShowGridButtons, InStr(eleShowGridButtons, "|") + 1) & "</a>")
			End Select
			Response.Write(" </td>")
		Next
		Response.Write("</tr></table>")
	End Sub

	Private Sub ProcessSearch()
		'Accepts special Form Variables, finishes with redirect to browse page
		Dim strRedirect, strKeywordNumberFields, strKeywordTextFields, strSQLSearchWhere, strKeyword, strSQLSearchFrom, qstTemp, intTemp, intFieldType, intTemp2, arrTemp, fmTemp, eleTemp, strKeywordCurrencyFields, arrCF, eleCF, blnISC, strSFieldName, intLenIdent
		strKeywordCurrencyFields = ""
		If Request.Form("ocdCancel_" & GRIDID) <> "" Then
			strRedirect = strSCRIPT_NAME & "?sqlfrom_" & GRIDID & "=" & Request.QueryString("sqlfrom_" & GRIDID)
			'strip out any unwanted querystrings
			For Each qstTemp In Request.QueryString
				If UCase(qstTemp) <> UCase("SQLFROM_" & GRIDID) And UCase(qstTemp) <> UCase("OCDFILTERFIELDNAME_" & GRIDID) And UCase(qstTemp) <> UCase("OCDGRIDMODE_" & GRIDID) And UCase(qstTemp) <> UCase("OCDFILTERFIELDTYPE_" & GRIDID) And UCase(qstTemp) <> "OCDFILTERSIZE" And UCase(qstTemp) <> "DATABASETYPE" And UCase(qstTemp) <> "GRIDID" And UCase(qstTemp) <> "SCRIPT" And UCase(qstTemp) <> "OCDACTION_" & GRIDID And UCase(qstTemp) <> "OCDFILTERDISPLAYNAME" Then
					strRedirect = strRedirect & "&" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp))
				End If
			Next
			Call Close()
			Response.redirect(strRedirect)
		End If
		strSQLSearchWhere = ""
		If Request.Form("ocdKeyword_" & GRIDID) <> "" and ShowKeywordSearch Then
			If SearchKeywordTextFields <> "" Or SearchKeywordNumberFields <> "" Then
				strKeywordTextFields = SearchKeywordTextFields & ","
				strKeywordNumberFields = SearchKeywordNumberFields & ","
			Else
				strKeywordTextFields = ""
				strKeywordNumberFields = ""
				strKeywordCurrencyFields = ""
				'create list of keyword fields
				intTemp = 0
				For Each eleTemp In arrSQLSelect
					Select Case CInt(arrSQLSelectType(intTemp))
						Case 129, 200, 201, 130, 202, 203						 'adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar 
							strKeywordTextFields = strKeywordTextFields & eleTemp & ","
						Case 2, 3, 4, 5, 14, 16, 17, 18, 19, 20, 21, 131, 6, 11, 135, 7, 133, 134						  'adSmallInt, adInteger, adSingle, adDouble, adDecimal, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adBigInt, adUnsignedBigInt, adBinary, adNumeric, adVarBinary, adLongVarBinary, adCurrency, adBoolean'adDBTimeStamp, adDate, adDBDate,  adDBTime
							strKeywordNumberFields = strKeywordNumberFields & eleTemp & ","
							If CInt(arrSQLSelectType(intTemp)) = 6 Then
								If DatabaseType = "SQLServer" Then
									strKeywordCurrencyFields = strKeywordCurrencyFields & eleTemp & ","
								End If
							End If
					End Select
					intTemp = intTemp + 1
					If intTemp >= MaxKeywordFields then
						exit for
					end if
				Next
			End If
			If strKeywordTextFields <> "" Or strKeywordNumberFields <> "" Then
			strKeyword = Request.Form("ocdKeyword_" & GRIDID)
				If UseRegExKeywordSearch And (DatabaseType = "Access" Or DatabaseType = "SQLServer") And (Instr(UCase(strKeyword)," AND ") > 0 Or Instr(Ucase(strKeyword)," OR ") > 0 Or Instr(UCase(strKeyword),"NOT ") > 0 ) Then
				'---------------
					Dim objExpr, strNDKeywordSearch, strNotStuff, strANDStuff, strORStuff, SQLSearchWhere, intExpr, strCurrWord, intKWNF, arrKWNF, intKWF, arrKWF
					If strKeywordTextFields <> "" Then
						strKeywordTextFields = Left(strKeywordTextFields,len(strKeywordTextFields) - 1)
						arrKWF = split(strKeywordTextFields,",")				
					End If
					If strKeywordNumberFields <> "" Then
						strKeywordNumberFields  = left(strKeywordNumberFields, len(strKeywordNumberFields) -1)
						arrKWNF = split(strKeywordNumberFields,",")				
					End If
					strNDKeywordSearch = strKeyword
					Set objExpr = New RegExp
					objExpr.IgnoreCase = True
					objExpr.Global = True
					objExpr.Pattern = "(((NOT\s+)\s*)\b[-\w']+\b)|(((NOT)\s*)+?[""].*[""])"
					Set strNOTStuff = objExpr.Execute(strNDKeywordSearch)
					strNDKeywordSearch = objExpr.Replace(strNDKeywordSearch, "")
					objExpr.Pattern = "((AND\s)\s*\b[-\w']+\b)|((AND)\s*[""].*[""])" 
					Set strANDStuff = objExpr.Execute(strNDKeywordSearch)
					strNDKeywordSearch = objExpr.Replace(strNDKeywordSearch, "")
					objExpr.Pattern = "(((OR\s+)\s*)?\b[-\w']+\b)|(((OR)\s*)?[""].*[""])"
					Set strORStuff = objExpr.Execute(strNDKeywordSearch)
					strNDKeywordSearch = objExpr.Replace(strNDKeywordSearch, "")
					If strANDStuff.Count = 0 Then
						'nothing to do
					Else
						SQLSearchWhere = SQLSearchWhere & "("
						For intExpr = 0 To strANDStuff.Count - 1
							strCurrword = strANDStuff.Item(intExpr).Value
							objExpr.Pattern = "^(AND)\s*"
							strCurrword = objExpr.Replace(strCurrword, "")
							strCurrword = Replace(strCurrword, """","")
							strCurrword = Replace(strCurrword, "'", "''")
							If intExpr <> 0 Then
								SQLSearchWhere = SQLSearchWhere & " AND "
		 					End If
							If strKeywordTextFields <> "" Then
								For intKWF = 0 To UBound(arrKWF)
									SQLSearchWhere = SQLSearchWhere &  FormatForSQL((arrKWF(intKWF)), DatabaseType, "ADDSQLIDENTIFIER") & " LIKE '%" & strCurrword & "%' OR " 
								Next
							End if
							Select Case DatabaseType
								Case "Access","SQLServer"
									If strKeywordNumberFields <> "" Then
										If isnumeric(strCurrword) Then							
											For intKWNF = 0 to UBound(arrKWNF)
												If DatabaseType = "Access" Then
													SQLSearchWhere = SQLSearchWhere & "IIf(isnull(" & FormatForSQL((arrKWNF(intKWNF)),DatabaseType,"ADDSQLIDENTIFIER") & "),'',CStr(" & FormatForSQL((arrKWNF(intKWNF)), DatabaseType,"ADDSQLIDENTIFIER") & ")) LIKE '%" & strCurrword & "%' OR " 
												Else
													SQLSearchWhere = SQLSearchWhere & "Cast(isnull(" & FormatForSQL((arrKWNF(intKWNF)),DatabaseType,"ADDSQLIDENTIFIER") & ",0) as VarChar) LIKE '%" & strCurrword & "%' OR " 
												End if
											Next
										End If
									End If
							End Select
							If Right(SQLSearchWhere,1) <> "(" Then
								SQLSearchWhere = left(SQLSearchWhere,Len(SQLSearchWhere)-4)
							End if
						Next
						If Right(SQLSearchWhere,1) = "(" Then
							SQLSearchWhere = Left(SQLSearchWhere,Len(SQLSearchWhere)-1)
						Else
							SQLSearchWhere = SQLSearchWhere & ")"
						End if
					End If
					If strORStuff.Count = 0 Then
					'nothing to do
					Else
						If SQLSearchWhere = "" Then
							SQLSearchWhere = "("
						Else
							SQLSearchWhere = SQLSearchWhere & " AND ("
						End If
						For intExpr = 0 To strORStuff.Count - 1
							strCurrword = strORStuff.Item(intExpr).Value
							objExpr.Pattern = "^(OR)\s*"
							strCurrword = objExpr.Replace(strCurrword, "")			
							strCurrword = Replace(strCurrword, """","")
							strCurrword = Replace(strCurrword, "'", "''")
							If intExpr <> 0 Then
								SQLSearchWhere = SQLSearchWhere & " OR "
							End If
							If strKeywordTextFIelds <> "" Then
								For intKWF = 0 to UBound(arrKWF)
									SQLSearchWhere = SQLSearchWhere &  FormatForSQL((arrKWF(intKWF)), DatabaseType, "ADDSQLIDENTIFIER") & " LIKE '%" & strCurrword & "%' OR "
								Next
							End if
							Select Case DatabaseType
								Case "Access","SQLServer"
									If strKeywordNumberFields <> "" Then
										If isnumeric(strCurrword) Then	
											For intKWNF = 0 To UBound(arrKWNF)
												If DatabaseType = "Access" Then
													SQLSearchWhere = SQLSearchWhere &  "IIf(isnull(" & FormatForSQL((arrKWNF(intKWNF)),DatabaseType,"ADDSQLIDENTIFIER") & "),'',CStr(" & FormatForSQL((arrKWNF(intKWNF)), DatabaseType,"ADDSQLIDENTIFIER") & ")) LIKE '%" & strCurrword & "%' OR "
												Else
													SQLSearchWhere = SQLSearchWhere & "Cast(isnull(" & FormatForSQL((arrKWNF(intKWNF)), DatabaseType, "ADDSQLIDENTIFIER") & ",0) as VarChar) LIKE '%" & strCurrword & "%' OR "
												End if
											Next
										End If
									End If
							End Select
							If right(SQLSearchWhere,1) <> "(" Then
								SQLSearchWhere = left(SQLSearchWhere,len(SQLSearchWhere)-4)
							End if
						Next
						If right(SQLSearchWhere,1) = "(" Then
							SQLSearchWhere = left(SQLSearchWhere,len(SQLSearchWhere)-1)
						Else
							SQLSearchWhere = SQLSearchWhere & ")"
						End if
					End If	
					If strNOTStuff.Count = 0 Then
						'nothing to do
					Else
						If SQLSearchWhere = "" Then
							SQLSearchWhere = "NOT ("
						Else
							SQLSearchWhere = SQLSearchWhere & " AND NOT ("
						End If
						For intExpr = 0 To strNOTStuff.Count - 1
							strCurrword = strNOTStuff.Item(intExpr).Value
							objExpr.Pattern = "^(NOT)\s*"
							strCurrword = objExpr.Replace(strCurrword, "")
							strCurrword = Replace(strCurrword, """","")
							strCurrword = Replace(strCurrword, "'", "''")
							If intExpr <> 0 Then
								SQLSearchWhere = SQLSearchWhere & " OR "
							End If
							If strKeywordTextFields <> "" Then
								For intKWF = 0 to UBound(arrKWF)
									SQLSearchWhere = SQLSearchWhere &  FormatForSQL((arrKWF(intKWF)), DatabaseType,"ADDSQLIDENTIFIER") & " LIKE '%" & strCurrword & "%' OR " 
								Next
							End if
							Select Case DatabaseType
								Case "Access","SQLServer"
									If strKeywordNumberFields <> "" Then
										If isnumeric(strCurrword) Then	
											For intKWNF = 0 to UBound(arrKWNF)
												If DatabaseType = "Access" Then
													SQLSearchWhere = SQLSearchWhere &  "IIf(isnull(" & FormatForSQL((arrKWNF(intKWNF)), DatabaseType, "ADDSQLIDENTIFIER") & "),'',CStr(" & FormatForSQL((arrKWNF(intKWNF)),DatabaseType,"ADDSQLIDENTIFIER") & ")) LIKE '%" & strCurrword & "%' OR " 
												Else
													SQLSearchWhere = SQLSearchWhere & "Cast(isnull(" & FormatForSQL((arrKWNF(intKWNF)),DatabaseType, "ADDSQLIDENTIFIER") & ",0) as VarChar) LIKE '%" & strCurrword & "%' OR " 
												End if
											Next
										End if
									End if
							End Select
							If Right(sqlsearchwhere,1) <> "(" Then
								SQLSearchWhere = left(SQLSearchWhere,len(SQLSearchWhere)-4)
							End if
						Next
						If right(SQLSearchWhere,1) = "(" Then
							SQLSearchWhere = left(SQLSearchWhere,len(SQLSearchWhere)-1)
						Else
							SQLSearchWhere = SQLSearchWhere & ")"
						End if
					End If
					strSQLSearchWhere = SQLSearchWhere
				'End RegExp Keyword search
				Else
				'Start Process regular keyword search
					strKeyword = Request.Form("ocdKeyword_" & GRIDID)
					If strKeywordTextFields <> "" Then
						strKeywordTextFields = Left(strKeywordTextFields, Len(strKeywordTextFields) - 1)
						arrTemp = Split(strKeywordTextFields, ",")
						For intTemp2 = 0 To UBound(arrTemp)
							strSQLSearchWhere = strSQLSearchWhere & FormatForSQL(arrTemp(intTemp2), DatabaseType, "AddSQLIdentifier") & " LIKE '%" & strKeyword & "%' OR "
						Next
					End If
					arrCF = Split(strKeywordCurrencyFields, ",")
					'Begin build criteria in strSQLSearchWhere
					Select Case DatabaseType
						Case "Access", "SQLServer"
							If strKeywordNumberFields <> "" Then
								strKeywordNumberFields = Left(strKeywordNumberFields, Len(strKeywordNumberFields) - 1)
								arrTemp = Split(strKeywordNumberFields, ",")
								If IsNumeric(strKeyword) Then
									For intTemp2 = 0 To UBound(arrTemp)
										If DatabaseType = "Access" Then
											strSQLSearchWhere = strSQLSearchWhere & "IIf(isnull(" & FormatForSQL(arrTemp(intTemp2), DatabaseType, "AddSQLIdentifier") & "),'',CStr(" & FormatForSQL(arrTemp(intTemp2), DatabaseType, "AddSQLIdentifier") & ")) LIKE '%" & strKeyword & "%' OR "
										Else
											blnIsC = False
											For Each eleCF In arrCF
												If UCase(eleCF) = UCase(arrTemp(intTemp2)) Then
													blnIsC = True
													Exit For
												End If
											Next
											If blnIsC Then
												strSQLSearchWhere = strSQLSearchWhere & "CONVERT(varchar(255)," & FormatForSQL(arrTemp(intTemp2), DatabaseType, "AddSQLIdentifier") & ") LIKE '%" & strKeyword & "%' OR "
											Else
												strSQLSearchWhere = strSQLSearchWhere & "" & FormatForSQL(arrTemp(intTemp2), DatabaseType, "AddSQLIdentifier") & " LIKE '%" & strKeyword & "%' OR "
											End if
									End If
								Next
								End If
							End If
					End Select
					If right(strSQLSearchWhere, 1) <> "(" And strSQLSearchWhere <> "" Then
						strSQLSearchWhere = Left(strSQLSearchWhere, Len(strSQLSearchWhere) - 4)
					End If
				End If 'keyword fields found
			End If 'regular keyword search
		End If 'keyword submitted
		For Each fmTemp In Request.Form
			'Check for database search fields in Form Variables
			If Len(fmTemp) > (14 + Len(GRIDID)) Then
				strSFieldName = Mid(fmTemp, (14 + Len(GRIDID)))
				If UCase(Left(fmTemp, (13 + Len(GRIDID)))) = UCase("OCDSEARCHTXT_" & GRIDID) Then
					If (Not Request.Form(fmTemp) = "") Or (UCase(Request.Form("ocdSearchSec_" & GRIDID & strSFieldName)) = "IS NULL" Or UCase(Request.Form("sec" & Mid(fmTemp, 4))) = "IS NOT NULL") Then
						strSQLSearchWhere = strSQLSearchWhere & " " & Request.Form("ocdSearchlogic_" & GRIDID) & " "
						strSQLSearchWhere = strSQLSearchWhere & " " & FormatForSQL(strSFieldName, DatabaseType, "AddSQLIdentifier") & " "
						intTemp = 0
						For Each eleTemp In arrSQLSelect
							If UCase(arrSQLSelect(intTemp)) = UCase(strSFieldName) Then
								Exit For
							End If
							intTemp = intTemp + 1
						Next
						Select Case CInt(arrSQLSelectType(intTemp))
							Case 16, 2, 3, 20, 17, 18, 19, 21, 4, 5, 6, 14, 131, 11							  'adTinyInt, adSmallInt, adInteger,adBigInt,adUnsignedTinyInt, adUnsignedSmallInt,adCurrency,adSingle,adDouble,adDecimal,adNumeric
								strSQLSearchWhere = strSQLSearchWhere & " " & Request.Form("ocdSearchsec_" & GRIDID & strSFieldName) & " " & Request.Form(fmTemp) & " "
							Case 7, 133, 134, 135 'adDBTimeStamp, adDate, adDBDate,  adDBTime
								If UCase(Request.Form("ocdSearchtxt_" & GRIDID & strSFieldName)) = "NULL" Then
									strSQLSearchWhere = strSQLSearchWhere & " Is " & Request.Form("ocdSearchslg_" & GRIDID & strSFieldName) & " Null"
								ElseIf UCase(Request.Form("ocdSearchtxt_" & GRIDID & strSFieldName)) = "NOT NULL" Then
									strSQLSearchWhere = strSQLSearchWhere & " Is Not Null"
								Else
									If DatabaseType = "Access" Then
										If Request.Form("ocdSearchtx2_" & GRIDID & strSFieldName) = "" Then
											strSQLSearchWhere = strSQLSearchWhere & " " & Request.Form("ocdSearchslg_" & GRIDID & strSFieldName) & " = #" & Request.Form(fmTemp) & "#"
										Else
											strSQLSearchWhere = strSQLSearchWhere & " " & Request.Form("ocdSearchslg_" & GRIDID & strSFieldName) & " BETWEEN #" & Request.Form(fmTemp) & "# AND #" & Request.Form("ocdSearchtx2_" & GRIDID & strSFieldName) & "#"
										End If
										ElseIf DatabaseType = "Oracle" Then
										'JG
										If Request.Form("ocdSearchtx2_" & GRIDID & strSFieldName) = "" Then
											strSQLSearchWhere = strSQLSearchWhere & " " & Request.Form("ocdSearchslg_" & GRIDID & strSFieldName) & " = TO_DATE('" & Request.Form(fmTemp) & ", '" & OracleDateFormat & "') "
										Else
										strSQLSearchWhere = strSQLSearchWhere & " " & Request.Form("ocdSearchslg_" & GRIDID & strSFieldName) & " 	BETWEEN  TO_DATE('" & Request.Form(fmTemp) & "', '" & OracleDateFormat & "') AND  TO_DATE('" & Request.Form("ocdSearchtx2_" & GRIDID & strSFieldName) & "', '" & OracleDateFormat & "')"
										End If

									Else
										If Request.Form("ocdSearchtx2_" & GRIDID & strSFieldName) = "" Then
											strSQLSearchWhere = strSQLSearchWhere & " " & Request.Form("ocdSearchslg_" & GRIDID & strSFieldName) & " = '" & Request.Form(fmTemp) & "'"
										Else
											strSQLSearchWhere = strSQLSearchWhere & " " & Request.Form("ocdSearchslg_" & GRIDID & strSFieldName) & " BETWEEN '" & Request.Form(fmTemp) & "' AND '" & Request.Form("ocdSearchtx2_" & GRIDID & strSFieldName) & "'"
										End If
									End If
								End If
							Case Else
								Select Case UCase(Request.Form("ocdSearchsec_" & GRIDID & strSFieldName))
									Case "IS NULL"
										strSQLSearchWhere = strSQLSearchWhere & " Is " & Request.Form("ocdSearchslg_" & GRIDID & strSFieldName) & " Null"
									Case "IN"
										strSQLSearchWhere = strSQLSearchWhere & " " & Request.Form("ocdSearchslg_" & GRIDID & strSFieldName) & " In (" & Request.Form(fmTemp) & ")"
									Case "IS"
										strSQLSearchWhere = strSQLSearchWhere & " Is " & Request.Form("ocdSearchslg_" & GRIDID & strSFieldName) & " " & Replace(Request.Form(fmTemp), "'", "''") & ""
									Case "CONTAINS"
										strSQLSearchWhere = strSQLSearchWhere & " " & Request.Form("ocdSearchslg_" & GRIDID & strSFieldName) & " Like '%" & Replace(Request.Form(fmTemp), "'", "''") & "%'"
									Case "STARTS WITH"
										strSQLSearchWhere = strSQLSearchWhere & " " & Request.Form("ocdSearchslg_" & GRIDID & strSFieldName) & " Like '" & Replace(Request.Form(fmTemp), "'", "''") & "%'"
									Case "LIKE"
										strSQLSearchWhere = strSQLSearchWhere & " " & Request.Form("ocdSearchslg_" & GRIDID & strSFieldName) & " Like '" & Replace(Request.Form(fmTemp), "'", "''") & "'"
									Case Else
										If Request.Form("ocdSearchslg_" & GRIDID & strSFieldName) <> "" Then
										strSQLSearchWhere = strSQLSearchWhere & " <> '" & Replace(Request.Form(fmTemp), "'", "''") & "'"
									Else
										strSQLSearchWhere = strSQLSearchWhere & " = '" & Replace(Request.Form(fmTemp), "'", "''") & "'"
									End If
								End Select 'Text Comparison
						End Select 'Field Type
					End If 'Process null comparisons and comparison with value
				End If 'Process Form Variables whose Name indicates a database search field
			End If 'Process Form Variables whose Name could have have search field
		Next 'Form Variable
		Select Case Left(strSQLSearchWhere, 4)
			Case " AND", " OR "
				If Request.Form("ocdSearchlogic_" & GRIDID) = "AND" Then
					strSQLSearchWhere = Mid(strSQLSearchWhere, 7)
				Else
					strSQLSearchWhere = Mid(strSQLSearchWhere, 6)
				End If
		End Select
		strSQLSearchWhere = trim(FormatForSQL(strSQLSearchWhere,DatabaseType,"CLEANUSERSQL"))
		strSQLSearchFrom = ""
		intTemp = 0
		For Each eleTemp In arrSQLSelect
			If Request.Form("ocdSearchexischk_" & GRIDID & UCase(eleTemp)) <> "" And Request.Form("ocdSearchshowchk_" & GRIDID & UCase(eleTemp)) = "" Then
				strSQLSearchFrom = strSQLSearchFrom & eleTemp & ","
			End If
			intTemp = intTemp + 1
		Next
		If strSQLSearchFrom <> "" Then
			strSQLSearchFrom = Left(strSQLSearchFrom, Len(strSQLSearchFrom) - 1)
		End If
		strRedirect = strSCRIPT_NAME & "?"
		strRedirect = strRedirect & "sqlorderby_" & GRIDID & "="
		If Request.Form("ocdSearchOrderBy_" & GRIDID) <> "" Then
			If DatabaseType = "Access" Then
				strRedirect = strRedirect & Server.URLEncode("[")
			Else
				strRedirect = strRedirect & Server.URLEncode("""")
			End If
			strRedirect = strRedirect & Server.URLEncode(Request.Form("ocdSearchOrderBy_" & GRIDID))
			If DatabaseType = "Access" Then
				strRedirect = strRedirect & Server.URLEncode("]")
			Else
				strRedirect = strRedirect & Server.URLEncode("""")
			End If
			strRedirect = strRedirect & Server.URLEncode(" " & Request.Form("ocdSearchOrderByOrder_" & GRIDID))
			If SearchSortSize > 1 Then
				For intTemp = 2 To SearchSortSize
					If Request.Form("ocdSearchOrderBy_" & GRIDID & CStr(intTemp)) <> "" Then
						If DatabaseType = "Access" Then
							strRedirect = strRedirect & Server.URLEncode(", [")
						Else
							strRedirect = strRedirect & Server.URLEncode(", """)
						End If
						strRedirect = strRedirect & Server.URLEncode(Request.Form("ocdSearchOrderBy_" & GRIDID & CStr(intTemp)))
						If DatabaseType = "Access" Then
							strRedirect = strRedirect & Server.URLEncode("]")
						Else
							strRedirect = strRedirect & Server.URLEncode("""")
						End If
						strRedirect = strRedirect & Server.URLEncode(" " & Request.Form("ocdSearchOrderByOrder_" & GRIDID & CStr(intTemp)))
					End If
				Next
			End If
		Else
			strRedirect = strRedirect & Server.URLEncode(SQLOrderBy)
		End If
		strRedirect = strRedirect & "&sqlwhere_" & GRIDID & "="
		If Request.QueryString("sqlwhere_" & GRIDID) <> "" Then
			If Request.Form("ocdSearchexistslogic_" & GRIDID) = "NEITHER" Then
				If Request.Form("ocdFindBrowse_" & GRIDID) = "" Then
					If strSQLSearchWhere <> "" Then
						strRedirect = strRedirect & "(" & Server.URLEncode(Trim(strSQLSearchWhere)) & ")"
					End If
				Else
					If strSQLSearchWhere <> "" Then
						strRedirect = strRedirect & " (" & Server.URLEncode(Trim(strSQLSearchWhere)) & ")"
						If SearchPersistKeyword Then
							strRedirect = strRedirect & "&ocdKeyword_" & GRIDID & "=" & Server.URLEncode(Request.Form("ocdKeyword_" & GRIDID))
						End If
					Else
						strRedirect = strRedirect & ""
					End If
				End If
			Else
				If Request.Form("ocdFindBrowse_" & GRIDID) = "" Then
					If strSQLSearchWhere <> "" Then
						strRedirect = strRedirect & "((" & Server.URLEncode(Request.QueryString("sqlwhere_" & GRIDID)) & ") " & Request.Form("ocdSearchExistslogic_" & GRIDID) & " (" & Server.URLEncode(Trim(strSQLSearchWhere)) & "))"
					Else
						strRedirect = strRedirect & Server.URLEncode(Request.QueryString("sqlwhere_" & GRIDID))
					End If
					If SearchPersistKeyword Then
						strRedirect = strRedirect & "&ocdKeyword_" & GRIDID & "=" & Server.URLEncode(Request.Form("ocdKeyword_" & GRIDID))
					End If
				Else
					If strSQLSearchWhere <> "" Then
						strRedirect = strRedirect & " (" & Server.URLEncode(Trim(strSQLSearchWhere)) & ")"
						If SearchPersistKeyword Then
							strRedirect = strRedirect & "&ocdKeyword_" & GRIDID & "=" & Server.URLEncode(Request.Form("ocdKeyword_" & GRIDID))
						End If
					End If
				End If
			End If
		Else
			If strSQLSearchWhere <> "" Then
				strRedirect = strRedirect & "(" & Server.URLEncode(Trim(strSQLSearchWhere)) & ")"
				If SearchPersistKeyword Then
					strRedirect = strRedirect & "&ocdKeyword_" & GRIDID & "=" & Server.URLEncode(Request.Form("ocdKeyword_" & GRIDID))
				End If
			End If
		End If
		strRedirect = strRedirect & "&sqlselecthide_" & UCase(GRIDID) & "=" & Server.URLEncode(strSQLSearchFrom)
		strRedirect = strRedirect & "&sqlpagesize_" & GRIDID & "=" & Server.URLEncode(Request.Form("ocdSearchPageSize_" & GRIDID))
		For Each qstTemp In Request.QueryString
			Select Case UCase(qstTemp)
				Case UCase("OCDGRIDMODE_" & GRIDID), UCase("SQLPAGESIZE_" & GRIDID), UCase("OCDFILTERFIELDNAME_" & GRIDID), "SQLWHERE_" & UCase(GRIDID), "SQLSELECTHIDE_" & UCase(GRIDID), "SQLORDERBY_" & UCase(GRIDID), "OCDFILTERFIELDTYPE_" & UCase(GRIDID), "OCDFILTERFIELDSIZE_" & UCase(GRIDID), "DATABASETYPE_" & UCase(GRIDID), "OCDACTION_" & UCase(GRIDID), "OCDKEYWORD_" & GRIDID, "SCRIPT_" & UCase(GRIDID)
				Case Else
					strRedirect = strRedirect & "&" & qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp))
			End Select
		Next
		Call Close()
		If Len(strRedirect) > GridMaxURLLength Then
			Err.raise(1030) 'Identifier Too Long, browser will not process querystring
			Exit Sub
		End If
		Response.clear()
		Response.redirect(strRedirect)
	End Sub
	
	Private Sub SetCriteria()
		Dim eleTemp, intArrPos, strFName, strFMap
		If SQLSelectName = "" Then
			Criteria = SQLWhere
		Else
			Criteria = SQLWhere
			intArrPos = 0
			For Each eleTemp in arrSQLSelect
				strFName = Trim(eleTemp)
				strFName = FormatForSQL(strFName, DatabaseType, "AddSQLIdentifier")
				strFMap = Trim(arrSQLSelectName(intArrPos))
				strFMap = FormatForSQL(strFMap, DatabaseType, "AddSQLIdentifier")
				Criteria = Replace(Criteria,strFName,strFMap)
				intArrPos = intArrPos + 1
			Next
		End If
	End Sub
	
	Private Sub DisplayFilter(ByVal strGridID)
		Dim intSize, intType, strName, qstTemp, arrFilterDropDownFields, eleFilterDropDownFields, showdropdown
		If SQLSelectName <> "" Then
			ArrSQLSelectName = Split(SQLSelectName, ",")
			ArrSQLSelect = Split(SQLSelect, ",")
		End If
		Response.Write("<form method=""Post"" action=""" & strSCRIPT_NAME & "?ocdGridMode_" & GRIDID & "=Process&amp;")
		For Each qstTemp In Request.QueryString
			Select Case UCase(qstTemp)
				Case UCase("OCDFILTERFIELDTYPE_" & GRIDID), UCase("OCDFILTERFIELDNAME_" & GRIDID), UCase("OCDFILTERFIELDDISPLAYNAME_" & GRIDID), UCase("OCDGRIDMODE_" & GRIDID), UCase("OCDFILTERFIELDSIZE_" & GRIDID)
				Case Else
					Response.Write(qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp)) & "&amp;")
			End Select
		Next
		Response.Write(""">")
		Response.Write("<input type=""hidden"" name=""ocdSearchPageSize_" & GRIDID & """ value=""" & server.htmlencode(request.querystring("sqlpagesize_" & GRIDID)) & """>")
		' format each field according To its type
		strName = Request.QueryString("ocdFilterFieldName_" & GRIDID)
		intSize = CLng(Request.QueryString("ocdFiltersize_" & GRIDID))
		intType = CInt(Request.QueryString("ocdFilterFieldtype_" & GRIDID))
		Select Case intType
			Case 128, 205, 204			'adBinary adVarBinary adLongVarBinary
				'do nothing
			Case Else
				If Request.QueryString("ocdFilterFielddisplayname_" & GRIDID) <> "" Then
					'no html encode
					Response.Write("<span class=""FieldName"">" & (Request.QueryString("ocdFilterFielddisplayname_" & GRIDID)) & "</span>")
				Else
					Response.Write("<span class=""FieldName"">")
					If SQLSelectName <> "" Then
						Response.Write(strName)
					Else
						Response.Write(Server.htmlencode(strName))
					End If
					Response.Write("</span>")
				End If
				Response.Write("&nbsp;")
				ShowDropDown = False
				Call DisplaySearchItem(strName, intType, intSize, showdropdown)
		End Select
		Response.Write("<p>")
		If SQLWhere <> "" Then
			SetCriteria()
			Response.Write("</p><input name=""ocdSearchexistslogic_" & GRIDID & """ value=""AND"" type=""radio"" checked>" & HTMLSearchAnd & "<input name=""ocdSearchexistslogic_" & GRIDID & """ value=""OR"" type=""radio"">" & HTMLSearchOr)
			Response.Write("<input name=""ocdSearchexistslogic_" & GRIDID & """ value=""NEITHER"" type=""radio"" checked>" & HTMLSearchReSet & "&nbsp;&nbsp;&nbsp;&nbsp;")
			Response.Write(Server.HTMLEncode(Criteria) & "<p>")
		End If
		Response.Write("<input name=""ocdFind_" & GRIDID & """" & HTMLAttribFindBtn & ">&nbsp;")
		Response.Write("<input name=""ocdCancel_" & GRIDID & """" & HTMLAttribCancelBtn & ">")
		Response.Write("</p></form>")
	End Sub

	Private Sub DisplaySearch(ByVal strGridID)
		Dim intSize, fldType, strName, qstTemp, arrSQLSelectSearchHide, intArrPos, arrSQLSelectSearchDropDown, eleSQLSelectSearchDropDown, eleSearchFieldHide, eleShowCheck, arrSQLSelectHideFields, eleSQLSelectHideFields, intOBLoop, intTemp, eleTemp, blnTemp, arrSQLSelectSearchUncheck
		If SQLSelectSearchUncheck <> "" Then
			arrSQLSelectSearchUncheck = Split(SQLSelectSearchUncheck, ",")
		End If
		If SQLSelectName <> "" Then
			arrSQLSelectName = Split(SQLSelectName, ",")
		End If
		Response.Write("<form method=""post"" action=""" & strSCRIPT_NAME & "?ocdGridMode_" & GridID & "=Process&amp;")
		For Each qstTemp In Request.QueryString
			If UCase(qstTemp) <> UCase("OCDGRIDMODE_" & GRIDID) Then
				Response.Write(qstTemp & "=" & Server.URLEncode(Request.QueryString(qstTemp)) & "&amp;")
			End If
		Next
		Response.Write(""">")
		If SQLSelectSearchUncheck <> "" Then
			arrSQLSelectSearchUncheck = Split(SQLSelectSearchUncheck, ",")
		End If
		If SQLSelectSearchHide <> "" Then
			arrSQLSelectSearchHide = Split(SQLSelectSearchHide, ",")
		End If
		Response.Write("<table>")
		If ShowKeywordSearch And UBound(arrSQLSelect) > 0 Then
			Response.Write("<tr><td align=""right"" nowrap>" & HTMLSearchKeyword & "</td><td colspan=""3""><input name=""ocdKeyword_" & GRIDID & """ size=""30""></td></tr>")
		End If
		If SQLSelectName <> "" Then
			ArrSQLSelectName = Split(SQLSelectName, ",")
		End If
		If Request.QueryString("sqlselecthide_" & GRIDID) <> "" Then
			arrSQLSelectHideFields = Split(CStr(Request.QueryString("sqlselecthide_" & GRIDID)), ",")
		End If
		intArrPos = 0
		If SQLSelectSearchDropDown <> "" Then
			arrSQLSelectSearchDropDown = Split(SQLSelectSearchDropDown, ",")
		End If
		intTemp = 0
		For Each eleTemp In arrSQLSelect
			blnTemp = True
			strName = eleTemp
			intSize = 50			 'FldF.DefinedSize
			fldtype = CInt(arrSQLSelectType(intTemp))
			If SQLSelectSearchHide <> "" Then
				For Each elesearchfieldhide In arrSQLSelectSearchHide
					If UCase(strName) = UCase(elesearchfieldhide) Then
						blnTemp = False
						Exit For
					End If
				Next
			End If
			If fldtype = 128 Or fldtype = 205 Or fldtype = 204 Or Not blnTemp Then			 'adBinary adVarBinary adLongVarBinary
				'do nothing
			ElseIf GridHideAutonumber And UCase(strName) = UCase(intSQLSelectIDPos) Then
				'do nothing
			Else
				Response.Write("<tr><td nowrap valign=""middle"" align=""right"">")
				If ShowDescription And DatabaseType = "Access" Then
					Response.Write("<A HREF="""" onclick=""javascript:window.open('DescribeField.asp?sqlfrom=" & Server.URLEncode(SQLFrom) & "&amp;sqlfield=" & Server.URLEncode(strName) & "', 'describe','height=200,width=300,scrollbars=yes');return false""><IMG ALT=""Describe Field"" SRC=""appHelpSmall.gif"" Border=0></A>&nbsp;")
				End If
				Response.Write("<span class=""FieldName"">")
				If SQLSelectName <> "" Then
					Response.Write(ArrSQLSelectName(intArrPos))
				Else
					Response.Write(Server.htmlencode(strName))
				End If
				'Parse Show/Hide fields
				Response.Write(":</span></td><td nowrap><input type=""checkbox"" name=""ocdSearchshowchk_" & GRIDID & Server.htmlencode(strName) & """ ")

				If Request.QueryString("sqlselecthide_" & GRIDID) = "" Then
					If SearchCheckAll Then
						If SQLSelectSearchUncheck = "" Then
							Response.Write(" checked ")
						Else
							blnTemp = True
							For Each eleshowcheck In arrSQLSelectSearchUncheck
								If UCase(eleshowcheck) = UCase(strName) Then
									blnTemp = False
								End If
							Next
							If blnTemp Then
								Response.Write(" checked ")
							End If
						End If
					End If
				Else
					blnTemp = False
					For Each eleSQLSelectHideFields In arrSQLSelectHideFields
						If UCase(eleSQLSelectHideFields) = UCase(strName) Then
							blnTemp = True
						End If
					Next
					If Not blnTemp Then
						Response.Write(" checked ")
					End If
				End If
				Response.Write(" > <input type=""hidden"" name=""ocdSearchexischk_" & GRIDID & Server.htmlencode(strName) & """ value=""Exists"" > </td><td align=left valign=MIDDLE nowrap>")
				blnTemp = False
				If SQLSelectSearchDropDown <> "" Then
					For Each eleSQLSelectSearchDropDown In arrSQLSelectSearchDropDown
						If Trim(UCase(eleSQLSelectSearchDropDown)) = Trim(UCase(strName)) Then
							blnTemp = True
						End If
					Next
				End If
				Call DisplaySearchItem(strName, fldType, intSize, blnTemp)
				Response.Write("</td></tr>")
			End If
			intArrPos = intArrPos + 1
			intTemp = intTemp + 1
		Next
		Response.Write("</table>")
		Response.Write("<p><input name=""ocdFind_" & GRIDID & """" & HTMLAttribFindBtn & ">&nbsp;")
		If UBound(arrSQLSelect) > 0 Then
			Response.Write("<input name=""ocdSearchlogic_" & GRIDID & """ value=""AND"" type=""radio"" checked >" & HTMLSearchAll & "<input name=""ocdSearchlogic_" & GRIDID & """ value=""OR"" type=""radio"" >" & HTMLSearchAny & "")
		End If
		If SQLWhere <> "" Then
			SetCriteria()
			Response.Write("<br><input name=""ocdSearchexistslogic_" & GRIDID & """ value=""AND"" type=""radio"" checked >" & HTMLSearchAnd & "<input name=""ocdSearchexistslogic_" & GRIDID & """ value=""OR"" type=""radio"" >" & HTMLSearchOr)
			Response.Write(Server.HTMLEncode(Criteria))
			Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;<input name=""ocdSearchexistslogic_" & GRIDID & """ value=""NEITHER"" type=""radio"" checked >Reset")
		End If
		If UBound(arrSQLSelect) < (SearchSortSize - 1) Then
			SearchSortSize = UBound(arrSQLSelect) + 1
		End If
		Response.Write("<p><table><tr><td align=""right"" valign=""top"" nowrap><span class=""FieldName"">Order By:</span></td><td valign=""top"" nowrap colspan=3>")
		For intOBLoop = 1 To SearchSortSize
			If intOBLoop > 1 Then
				Response.Write("<br>")
			End If
			Response.Write("<select name=""ocdSearchOrderBy_" & GRIDID)
			If intOBLoop > 1 Then
				Response.Write(intOBLoop)
			End If
			Response.Write("""><option value=""""></option>")
			intTemp = 0
			For Each eleTemp In arrSQLSelect
				blnTemp = True
				If SQLSelectSearchHide <> "" Then
					For Each elesearchfieldhide In arrSQLSelectSearchHide
						If UCase(eletemp) = UCase(elesearchfieldhide) Then
							blnTemp = False
							Exit For
						End If
					Next
				End If
				If blnTemp Then
					Select Case CInt(arrSQLSelectType(intTemp))
						Case 128, 205, 204, 201, 203
						Case Else
							Response.Write("<option value=""" & Server.htmlencode(eleTemp) & """>") 
							If SQLSelectName = "" Then
								Response.Write(eletemp)
							Else
								Response.Write(ArrSQLSelectName(intTemp))
							End If
							Response.Write("</option>")
					End Select
				End If
				intTemp = intTemp + 1
			Next
			Response.Write("</select> <select name=""ocdSearchOrderByOrder_" & GRIDID)
			If intOBLoop > 1 Then
				Response.Write(intOBLoop)
			End If
			Response.Write("""><option value=""ASC"" selected>" & HTMLAscText & "</option><option value=""DESC"">" & HTMLDescText & "</option></select>")
		Next
		Response.Write("</td></tr><tr><td align=""right"" nowrap><span class=""FieldName"">Page Size:</span></td><td nowrap colspan=""3"">")
		Response.Write("<input name=""ocdSearchPageSize_" & GRIDID & """ size=""5"" value=""" & Server.htmlencode(SQLPageSize) & """> ")
		Response.Write("</tr></table></form>")
	End Sub
	
	Private Sub DisplaySearchItem(ByVal strName, ByVal fldType, ByVal fldSize, ByVal blnUseDropDown)
		Dim strSearchSQL, intI, rsConstraint, arrGetRows, eleTcomps, arrTComps
		If blnUseDropDown Then
			set rsConstraint = Server.CreateObject("ADODB.Recordset")
			strSearchSQL = "SELECT DISTINCT" & FormatForSQL(strName, DatabaseType, "AddSQLIdentifier")
			strSearchSQL = strSearchSQL & " FROM "
			If InStr(1, SQLFrom, ",") = 0 And InStr(1, SQLFrom, "=") = 0 Then
				SQLFrom = FormatForSQL(SQLFrom, DatabaseType, "AddSQLIdentifier")
			End If
			strSearchSQL = strSearchSQL & SQLFrom
			If SQLWhereExtra <> "" Then
				strSearchSQL = strSearchSQL & " WHERE " & SQLWHEREEXTRA & " AND (" & FormatForSQL(strName, DatabaseType, "AddSQLIdentifier") & " Is Not Null)"
			Else
				strSearchSQL = strSearchSQL & " WHERE " & FormatForSQL(strName, DatabaseType, "AddSQLIdentifier") & " Is Not Null"
			End If
			strSearchSQL = strSearchSQL & " ORDER BY " & FormatForSQL(strName, DatabaseType, "AddSQLIdentifier")
			On Error Resume Next
			Call rsConstraint.Open(strSearchSQL, ADOConnection)
			If Err.Number <> 0 Then
				Err.Clear
				On Error GoTo 0
				Err.raise(13)
			End If
			On Error GoTo 0
			If rsConstraint.eof Then
				rsConstraint.close()
				Set rsConstraint = Nothing
				Response.Write(" -- ")
			Else
				arrGetRows = rsConstraint.GetRows
				rsConstraint.Close()
				Set rsConstraint = Nothing
				Response.Write(" = <select name=""ocdSearchtxt_" & GRIDID & Server.HTMLEncode(strName) & """>")
				Response.Write("<option value="""" selected></option>")
				For intI = 0 To UBound(arrGetRows, 2)
					Response.Write("<option value=""")
					Response.Write(Server.HTMLEncode(arrGetRows(0, intI)))
					Response.Write(""" >")
					Response.Write(Server.HTMLEncode(arrGetRows(0, intI)))
					Response.Write("</option>")
				Next
				Response.Write("</select>")
			End If
		Else
			Select Case fldType
				Case 7, 133, 134, 135				'adDBTimeStamp, adDate, adDBDate,  adDBTime
					Response.Write("<select name=""ocdSearchslg_" & GRIDID & Server.HTMLEncode(strName) & """><option value="""" selected></option><option value=""NOT"">NOT</option></select> " & HTMLSearchBetween & " <input type=""text"" size=12 Maxlength=50 name=""ocdSearchtxt_" & GRIDID & Server.HTMLEncode(strName) & """ > " & HTMLSearchBetweenAnd & " <input type=""text"" size=12 Maxlength=50 name=""ocdSearchtx2_" & GRIDID & Server.HTMLEncode(strName) & """ >")
				Case 11				'adBoolean
					Response.Write("<select name=""ocdSearchsec_" & GRIDID & Server.HTMLEncode(strName) & """><option value=""="" selected>=</option></select>")
					If DatabaseType = "SQLServer" Then
						Response.Write("<select name=""ocdSearchtxt_" & GRIDID & Server.HTMLEncode(strName) & """ ><option value=""""></option><option value=""1"">True</option><option value=""0"">False</option></select>")
					Else
						Response.Write("<select name=""ocdSearchtxt_" & GRIDID & Server.HTMLEncode(strName) & """ ><option value=""""></option><option value=""True"">True</option><option value=""False"">False</option></select>")
					End If
				Case 16, 2, 3, 20, 17, 18, 19, 21, 4, 5, 6, 14, 131				  'adTinyInt, adSmallInt, adInteger,adBigInt,adUnsignedTinyInt, adUnsignedSmallInt,adCurrency,adSingle,adDouble,adDecimal,adNumeric
					Response.Write("<select name=""ocdSearchsec_" & GRIDID & Server.HTMLEncode(strName) & """ ><option value=""="" selected>=<option value=""&gt;"">&gt;<option value=""&lt;"">&lt;<option value=""&lt;="">&lt;=<option value=""&gt;="">&gt;=<option value=""&lt;&gt;"">&lt;&gt;<option value=""Is Null"">Is Null<option value=""Is Not Null"">Is Not Null</select><input type=""text"" name=""ocdSearchtxt_" & GRIDID & Server.HTMLEncode(strName) & """ >")
				Case Else
					arrTcomps = Split(HTMLTextCompare, ";;")
					Response.Write("<select name=""ocdSearchslg_" & GRIDID & Server.HTMLEncode(strName) & """ ><option value="""" selected></option><option value=""NOT"">NOT</option></select> <select name=""ocdSearchsec_" & GRIDID & Server.HTMLEncode(strName) & """>")
					For Each eleTcomps In arrTComps
						If Not ((Left(eleTcomps, InStr(eleTcomps, "|") - 1) = "=" Or UCase(Left(eleTcomps, InStr(eleTcomps, "|") - 1)) = "IN") And (DatabaseType = "SQLServer" And (fldType = 201 Or fldType = 203))) Then
							Response.Write("<option value=""" & Left(eleTcomps, InStr(eleTcomps, "|") - 1) & """")
							If UCase(SearchDefaultTextCompare) = UCase(Left(eleTcomps, InStr(eleTcomps, "|") - 1)) Then
								Response.Write(" selected")
							End If
							Response.Write(">" & Mid(eleTcomps, InStr(eleTcomps, "|") + 1) & "</option>")
						End If
					Next
					Response.Write("</select>")
					Response.Write("<input type=""text"" name=""ocdSearchtxt_" & GRIDID & Server.HTMLEncode(strName) & """>")
			End Select
		End If
	End Sub
End Class
%>
