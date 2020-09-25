<%'There should be no ASP or HTML code above this line

'1 Click DB Variables for Application Configuration 

'1 Click DB technology is fully protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'Use of this software and source code is strictly at your own risk.
'All warranties are specifically disclaimed except as required by law.

'IMPORTANT : THIS APPLICATION USES PASS THROUGH DATABASE SECURITY  !
'
'To enforce application security, set logins and permissions
'for all web server and database users as appropriate.
'
'Use object and configuration properties _only_ to customize 
'application appearance and interactions with other components. 
'For more information see : http://1ClickDB.com
'Configuration variables can either be hard-coded or set dynamically 

ocdReadOnly = False 'If true, modifications to your database are discouraged.  For security, also set user and database permissions as appropriate
ocdShowAdmin = True 'If true, show advanced design interfaces
ocdShowImport = True 'If true, show Import Wizard
ocdShowWizard = False 'If true, show Wizard controls
ocdAdminPassword = "" 'comma delimited list of passwords required to create dynamic connections.  Session cookies required for this option.
ocdADOConnection = ""' e.g : "provider=Microsoft.Jet.OLEDB.4.0;data source=c:\inetpub\data\northwind2000.mdb;" 'hard coded connection string, if not set session cookies will be used to gather dynamic connection information
ocdADOUsername = "" 'hard coded db user name, only active if ocdADOConnection is set
ocdADOPassword = "" 'hard coded db password, only active if ocdADOConnection is set
ocdPageSizeDefault = 10 'default page size for browse grids
ocdBrandLogo = "" '"Your HTML Link/Logo Here" 'replace "1 Click DB Pro" HTML in upper right corner of interface
ocdBrandText = "1 Click DB Pro" 'replace "1 Click DB Pro" TEXT in the interface
ocdFooterHTML = "" 'replace bottom disclaimer on all pages
ocdSessionTimeout = 50 'Session timeout in minutes
ocdDBTimeout = 5 'Seconds before giving up on returning results from an SQL query
ocdComputeTimeout = 10 'Seconds before giving up on returning results from an computed statistics for a query
ocdMultipleFieldSort = False 'Set to False for non-additive Order By sorts
ocdAllowBrowseRefresh = True 'False to disable Browse Grid Refresh Functions
ocdGridHighlightSelected = True 'if true, enable highlight selected records in Browse grids
ocdDefaultTextCompare = "Starts With" 'Default comparision operator for text field, valid values are "=","Starts With", "Contains","Like","In","Is Null"
ocdShowDefaults = True 'Enable display of default values when adding new records
ocdCodePage = 1252
ocdFormNullToken = ""
ocdShowRelatedRecords = True	'Show browse of related records from edit screens; Default=true
ocdSelectForeignKey = True 'Create dropdown select box for picking single field foreign keys, if false show regular text field; Default=true
ocdJETSQLReference = "http://msdn.microsoft.com/library/default.asp?url=/library/en-us/office97/html/output/F1/D2/S5A318.asp?frame=true" 'For Command screen line
ocdMSSQLReference = "http://msdn.microsoft.com/library/default.asp?url=/library/en-us/tsqlref/ts_tsqlcon_6lyk.asp?frame=true" 'For Command screen line
ocdOraSQLReference = "http://technet.oracle.com/doc/lite/sqlref/html/sqltoc.htm"
ocdShowPlanReference = "http://www.microsoft.com/technet/treeview/default.asp?url=/technet/prodtechnol/sql/maintain/featusability/showplan.asp"
ocdDisableTextDriver = True 'Use True to discourage Text driver connections. Remove driver from server when not required or if security is an issue.
ocdAllowExport = True
ocdHideAutoNumber = False ' If true, autonumber/identity fields are hidden for all tables/views
ocdShowCheckedSearchFields = True
ocdMaxRecordsRetrieve = 10000
ocdConnectionShortCuts = ""
ocdMaxRecordsDisplay = 5000
ocdMaxURLLength = 2000 'IE 5.0 limit is 2083, other browsers may be more or less if application sees a violation an "Identifier too long" exception is thrown.  URLs that exceed your browser's capabilities usually result in non-specific error messages.
ocdAuditWizardPrefix = "audit_" 'prefix used before all audit wizard object names
ocdSchemaHideObjects = "" 'Comma delimited list of objects to hide from schema view, must be appropriately delmited as [TableName] for MS Access or "Owner"."TableName" for SQL Server and Oracle
ocdCustomEditPages = "" 'if "" all edit page requests go to dynamic Edit.asp page, if "*" all edit page requests are forwarded to custom file named ObjectName_Edit.asp, otherwise use as a comma delmited list of objects names to be forwarded to custom file named ObjectName_Edit.asp.  Object names must be appropriately delmited
ocdCharSet = "iso-8859-1" 
ocdStyleSheet = "" 'Hard code this to override defaults and motif styles
ocdRunImportEventCode = False 'supported only as custom option
'Session.LCID = 1033 'Force 1 Click DB ASP code to use English_United_States locale settings
ocdLaunchPage = ""
ocdWrapGrid = False 'if true, let web browser control cell wrapping
ocdServerScriptTimeOut = 120
ocdShowSQLText = True
ocdRenderAsHTML = False 'Warning ! Setting this option to TRUE for untrusted data introduces Cross Site Scripting vulnerabilities.
ocdGridIcons = True 'if true, use standard grid icons instead of plain text links
ocdDebug = False
'There should be no ASP or HTML code below this line
%>
