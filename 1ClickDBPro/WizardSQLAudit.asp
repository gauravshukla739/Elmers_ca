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

if ocdIsODBC Then
call writeheader("")
call writefooter("An OLEDB Connection is required to run the Audit Wizard")
end if

if request.querystring("SQLFrom") = "" Then
call writeheader("")
Response.write ("<SPAN CLASS=Information>Select Table to Enable Audit Triggers")
Response.write ("</span><BLOCKQUOTE>")
dim fmTemp
dim GRIDID
GRIDID = Request.QueryString("GRIDID")
	Response.write "<FORM ACTION=""" & Request.servervariables("SCRIPT_NAME") & """>"
	for each fmTemp in Request.Querystring
		select case UCASE(fmtemp)
			Case "SQLFROM"
			Case Else
				Response.write ("<INPUT TYPE=Hidden NAME=""" & fmTemp & """ VALUE=""" & server.htmlencode(request.querystring(fmtemp)) & """>")
		end select
	next

	
'		Response.write ("<INPUT NAME=""SelectAnotherTable"" TYPE=Submit Value=""Select Another ")
'	select case request.querystring("objtoshow")
'		Case "Both"
'			Response.write "Table/Query"
'		Case "Queries"
'			Response.write "Query"
'		Case Else
'			Response.write "Table"
'	End Select
'	Response.write (" &gt;&gt;""><P>")
'response.end
Response.write ("<SELECT NAME=""sqlfrom"" SIZE=12>")
dim rsTemp
dim xConn
set xConn = server.CreateObject("ADODB.Connection")

xConn.Open ndnscSQLConnect, ndnscSQLUser, ndnscSQLPass
set rsTemp = xConn.OpenSchema(20) 'adSchemaTables

Do While Not rsTemp.EOF
	If ((rsTemp.Fields("TABLE_TYPE").Value = "TABLE" ) AND UCASE(LEFT(rsTemp.Fields("TABLE_NAME").Value,4)) <> "MSYS")  Then
		Response.write ("<OPTION VALUE=""" & server.htmlencode(rsTemp.Fields("TABLE_NAME").VALUE) & """>")
		Response.write (Server.HTMLEncode(rsTemp.Fields("TABLE_NAME").Value))
		Response.write ("</OPTION>")			
	End if
	rsTemp.movenext
Loop
Response.write ("</SELECT><P>")

Response.write ("<INPUT NAME=""SelectTables"" TYPE=Submit Value=""Audit &gt;"" Class=submit><P>")
Response.write ("</FORM>")


response.write "</BLOCKQUOTE>"
call writefooter("")
response.end
End if
dim histrecscreated
histrecscreated = false
dim okaytogo
okaytogo = false
dim histexists
histexists = false
dim histhasrecords
histhasrecords = false
dim tmpcolname
dim varFormNum
dim strTrig
strTrig = ""
if not cbool(cint(ndnscCompatibility) and ocdNoJavaScript) Then

if ocdUseFrameset then
	varFormNum = "1"
Else
	varFormNum = "2"
End if
End if
if ocdReadOnly then
	call writeheader("")
'	Response.write ("1 Click DB is Read Only")
	call writefooter("READONLY")
	response.end
Elseif not ocdShowSQLExecutor then
	call writeheader("")
	Response.write "Access denied"
	call writefooter("")
	response.end
End if
dim ndOpenConn
	set ndOpenConn = server.CreateObject("ADODB.Connection")
	if ocdReadOnly then
		ndOpenConn.Mode = 1 'adOpenRead
	End if
	call ndOpenConn.Open (ndnscSQLConnect, ndnscSQLUser, ndnscSQLPass)
dim rsHistExists
set rsHistExists = ndOpenConn.OpenSchema(20,array(empty,empty,CSTR(ocdAuditWizardPrefix  & NOSQLIDentifier(request.querystring("sqlfrom")))))  ' adOpenTables
if not rsHistExists.eof then
HistExists = True

	dim rsHistHasRecords
	set rsHistHasRecords = ndOpenConn.execute ("SELECT TOP 1 * from """ & ocdAuditWizardPrefix &  NOSQLIDentifier(request.querystring("sqlfrom")) & """")
	if not rsHistHasRecords.eof then
		histhasrecords = true
	End if
End if
set rsHistExists = nothing
dim strInsertIntoSQL
dim strSelectIntoSQL
dim rsOTabDef
dim fldOTabDef
set rsOTabDef = server.createobject("ADODB.Recordset")
if request.form("InsertCurrent") <> "" and ((histhasrecords and request.form("overwriteHistRecords") <> "") or (not histhasrecords) )Then

strInsertIntoSQL = "Insert Into """ & ocdAuditWizardPrefix &  NoSQLIDentifier(request.QUerystring("SQLFrom")) & """ ("
rsOTabDef.open "Select * From """ & NoSQLIdentifier(request.querystring("SQLFrom")) & """ where 1=2", ndOpenConn
for each fldOtabDef in rsOtabDef.Fields
select case fldOtabDef.Type
Case 201,203,128,204,205
Case Else
'if fldOtabDef.Properties("ISAUTOINCREMENT") = "True" Then
'strInsertIntoSQL = strInsertIntoSQL & "Convert(INT,""" & fldOtabDef.Name & """) AS """ & fldOtabDef.Name & ""","
'Else
strInsertIntoSQL = strInsertIntoSQL & """" & fldOtabDef.Name & ""","
'End if
End Select
next
strInsertIntoSQL = strInsertIntoSQL & """" & ocdAuditWizardPrefix & "startdatetime"", """ & ocdAuditWizardPrefix & "enddatetime"", """ & ocdAuditWizardPrefix & "startusername"", """ & ocdAuditWizardPrefix & "startappname"", """ & ocdAuditWizardPrefix & "starthostname"""


'strInsertIntoSQL = left(strInsertIntoSQL,len(strInsertIntoSQL)-1)
strInsertIntoSQL = strInsertIntoSQL & ") SELECT "
for each fldOtabDef in rsOtabDef.Fields
select case fldOtabDef.Type
Case 201,203,128,204,205
Case Else
'if fldOtabDef.Properties("ISAUTOINCREMENT") = "True" Then
'strInsertIntoSQL = strInsertIntoSQL & "Convert(INT,""" & fldOtabDef.Name & """) AS """ & fldOtabDef.Name & ""","
'Else
strInsertIntoSQL = strInsertIntoSQL & """" & fldOtabDef.Name & ""","
'End if
End Select
next
strInsertIntoSQL = strInsertIntoSQL & " getdate() AS """ & ocdAuditWizardPrefix & "startdatetime"", '9/9/9999' AS """ & ocdAuditWizardPrefix & "enddatetime"", suser_sname() as """ & ocdAuditWizardPrefix & "startusername"", app_name() as """ & ocdAuditWizardPrefix & "startappname"", host_name() as """ & ocdAuditWizardPrefix & "starthostname"""
'strInsertIntoSQL = left(strInsertIntoSQL,len(strInsertIntoSQL)-1)
 strInsertIntoSQL = strInsertIntoSQL & " FROM """ & NoSQLIDentifier(request.QUerystring("SQLFrom")) & """"
'response.write strInsertIntoSQL
ndOpenConn.execute "TRUNCATE TABLE """ & ocdAuditWizardPrefix &  NoSQLIDentifier(request.QUerystring("SQLFrom")) & """"
if err.number = 0 then

ndOpenConn.execute strInsertIntoSQL
if err.number = 0 Then
	histhasrecords = true
	histrecscreated = true
End if
end if
End if
if request.form("MakeHistoryTrigger") <> "" Then
strTrig = strTrig & "CREATE TRIGGER """ & ocdAuditWizardPrefix & "trigger_" & NoSQLIdentifier(request.querystring("sqlfrom")) & """ ON """ & NoSQLIdentifier(request.querystring("sqlfrom")) & """" & vbCRLF
strTrig = strTrig & "FOR INSERT, UPDATE, DELETE" & vbCRLF
strTrig = strTrig & "NOT FOR REPLICATION" & vbCRLF
strTrig = strTrig & "AS" & vbCRLF

strTrig = strTrig & "DECLARE " & vbCRLF 
'strTrig = strTrig & "@CustomerID NCHAR(5)," & vbCRLF
strTrig = strTrig & "@TrigTime DateTime" & vbCRLF

'strTrig = strTrig & "set @CustomerID = (select CustomerID From DELETED)" & vbCRLF
strTrig = strTrig & "set @TrigTime = getDate()" & vbCRLF
'strTrig = strTrig & "BEGIN TRANSACTION" & vbCRLF

strTrig = strTrig & "UPDATE """ & ocdAuditWizardPrefix & NoSQLIdentifier(request.querystring("sqlfrom")) & """ SET " & ocdAuditWizardPrefix & "enddatetime = (@TrigTime), " & ocdAuditWizardPrefix & "endappname = (APP_Name()), " & ocdAuditWizardPrefix & "endusername = (SUSER_SName()), " & ocdAuditWizardPrefix & "endhostname = (HOST_NAME()) FROM deleted,""" & ocdAuditWizardPrefix &  NoSQLIdentifier(request.querystring("sqlfrom")) & """ WHERE "
'	dim tmpcolname

	dim rsIDX
	set rsIDX = server.createobject("ADODB.Recordset")
	set rsIDX = ndOpenConn.openSchema(12,Array(empty,empty,empty,empty,Cstr(NoSQLIdentifier(Request.Querystring("SQLFrom"))))) 'indexes
'response.write "SDF"
'	response.end
	do while not rsIDX.eof
	If UCASE(rsIDX("table_name")) = UCase(NoSQLIdentifier(Request.Querystring("SQLFrom"))) and rsIDX("primary_key") = True Then
okaytogo = true
tmpcolname = rsIDX.Fields("COLUMN_NAME").Value
						

strTrig = strTrig & """" & ocdAuditWizardPrefix &  NoSQLIdentifier(request.querystring("sqlfrom")) & """.""" & tmpcolName & """ = deleted.""" & tmpcolName & """ AND "
	End if
	rsIDX.movenext
	Loop
			if not okaytogo then
					call writeheader("")
					call writefooter("Primary Key Not Found")
				end if
strTrig = strTrig & ocdAuditWizardPrefix & "enddatetime = '9/9/9999'" & vbCRLF



strTrig = strTrig & "INSERT  INTO """ & ocdAuditWizardPrefix & NoSQLIdentifier(request.querystring("sqlfrom")) & """ ("
rsOTabDef.open "Select * From """ & NoSQLIdentifier(request.querystring("SQLFrom")) & """ where 1=2", ndOpenConn
for each fldOtabDef in rsOtabDef.Fields
select case fldOtabDef.Type
Case 201,203,128,204,205
Case Else
strTrig = strTrig & """" & fldOtabDef.name & ""","
end select
next
strTrig = strTrig & """" & ocdAuditWizardPrefix & "startdatetime"", """ & ocdAuditWizardPrefix & "enddatetime"", """ & ocdAuditWizardPrefix & "startusername"", """ & ocdAuditWizardPrefix & "startappname"", """ & ocdAuditWizardPrefix & "starthostname"""
strTrig = strTrig & ")"
strTrig = strTrig & " SELECT "
for each fldOtabDef in rsOtabDef.Fields
select case fldOtabDef.Type
Case 201,203,128,204,205
Case Else
strTrig = strTrig & """" & fldOtabDef.name & ""","
end select
next
strTrig = strTrig & "@TrigTime,'9/9/9999',suser_sname(),app_name(),host_name()"

strTrig = strTrig & " FROM inserted" & vbCRLF

'strTrig = strTrig & "COMMIT TRANSACTION" & vbCRLF

end if
if request.form("MakeHistoryTable") <> "" Then
if request.form("OverwriteHist") <> "" Then 
ndOpenConn.execute "DROP TABLE """ & ocdAuditWizardPrefix &  NoSQLIDentifier(request.QUerystring("SQLFrom")) & """"
HistHasRecords = False
End if
strSelectIntoSQL = "Select "
rsOTabDef.open "Select * From """ & NoSQLIdentifier(request.querystring("SQLFrom")) & """ where 1=2", ndOpenConn
for each fldOtabDef in rsOtabDef.Fields
select case fldOtabDef.Type
Case 201,203,128,204,205
Case Else
if fldOtabDef.Properties("ISAUTOINCREMENT") = "True" Then
strSelectIntoSQL = strSelectIntoSQL & "Convert(INT,""" & fldOtabDef.Name & """) AS """ & fldOtabDef.Name & ""","
Else
strSelectIntoSQL = strSelectIntoSQL & """" & fldOtabDef.Name & ""","
End if
End Select
next
strSelectIntoSQL = left(strSelectIntoSQL,len(strSelectIntoSQL)-1)
strSelectIntoSQL = strSelectIntoSQL & " into """ & ocdAuditWizardPrefix  & NoSQLIDentifier(request.QUerystring("SQLFrom")) & """ From """ & NoSQLIdentifier(request.querystring("SQLFrom")) & """ where 1=2"
'response.write strSelectIntoSQL
ndOpenConn.execute strSelectIntoSQL
if err.number = 0 Then'ndOpenConn.execute "Select * into ""ocdb_History_" & NoSQLIDentifier(request.QUerystring("SQLFrom")) & """ From """ & NoSQLIdentifier(request.querystring("SQLFrom")) & """ where 1=2"
'ndOpenConn.Execute "ALTER TABLE ""ocdb_History_" & NoSQLIDentifier(request.QUerystring("SQLFrom")) & """ ALTER COLUMN EmployeeID INTEGER"
ndOpenConn.Execute "ALTER TABLE """ & ocdAuditWizardPrefix & NoSQLIDentifier(request.QUerystring("SQLFrom")) & """ ADD " & ocdAuditWizardPrefix & "recordid INTEGER IDENTITY"
ndOpenConn.Execute "ALTER TABLE """ & ocdAuditWizardPrefix & NoSQLIDentifier(request.QUerystring("SQLFrom")) & """ ADD " & ocdAuditWizardPrefix & "startdatetime DATETIME"
ndOpenConn.Execute "ALTER TABLE """ & ocdAuditWizardPrefix & NoSQLIDentifier(request.QUerystring("SQLFrom")) & """ ADD " & ocdAuditWizardPrefix & "enddatetime DATETIME"
ndOpenConn.Execute "ALTER TABLE """ & ocdAuditWizardPrefix & NoSQLIDentifier(request.QUerystring("SQLFrom")) & """ ADD " & ocdAuditWizardPrefix & "startusername VARCHAR(255)"
ndOpenConn.Execute "ALTER TABLE """ & ocdAuditWizardPrefix & NoSQLIDentifier(request.QUerystring("SQLFrom")) & """ ADD " & ocdAuditWizardPrefix & "startappname VARCHAR(255)"
ndOpenConn.Execute "ALTER TABLE """ & ocdAuditWizardPrefix & NoSQLIDentifier(request.QUerystring("SQLFrom")) & """ ADD " & ocdAuditWizardPrefix & "starthostname VARCHAR(255)"
ndOpenConn.Execute "ALTER TABLE """ & ocdAuditWizardPrefix & NoSQLIDentifier(request.QUerystring("SQLFrom")) & """ ADD " & ocdAuditWizardPrefix & "endusername VARCHAR(255)"
ndOpenConn.Execute "ALTER TABLE """ & ocdAuditWizardPrefix & NoSQLIDentifier(request.QUerystring("SQLFrom")) & """ ADD " & ocdAuditWizardPrefix & "endappname VARCHAR(255)"
ndOpenConn.Execute "ALTER TABLE """ & ocdAuditWizardPrefix & NoSQLIDentifier(request.QUerystring("SQLFrom")) & """ ADD " & ocdAuditWizardPrefix & "endhostname VARCHAR(255)"
HistExists = True
HistHasRecords = False
End if
end if
Call WriteHeader("")
Response.Write ("<SPAN CLASS=Information>Enable Auditing for " & Request.Querystring("SQLFROM") & "</SPAN>")
Response.write "<FORM ACTION=""" & request.servervariables("SCRIPT_NAME") & "?" & request.querystring & """ METHOD=POST>"
response.write "<INPUT TYPE=Submit Name=MakeHistoryTable Value=""Create History Table"" Class=submitbtn>"
if HistExists Then
Response.write " "
Response.write "<INPUT TYPE=CHeckbox NAME=OverwriteHist> Drop Existing"
Response.write "<P>"
Response.write "<INPUT TYPE=Submit NAME=InsertCurrent VALUE=""Insert Current Data"" CLASS=Submitbtn>"
if HistHasRecords Then

Response.write " "
Response.write "<INPUT TYPE=Checkbox NAME=OverwriteHistRecords> Delete Existing"

End if
End if
response.write "<P>"
if histexists then
response.write "<INPUT TYPE=Submit Name=MakeHistoryTrigger Value=""Generate Audit Script"" CLASS=SubmitBtn>"
end if
Response.write "</FORM>"
if strTrig <> "" Then
Response.write "<FORM ACTION=""Command.asp"" METHOD=POST>"
response.write "<TEXTAREA NAME=sqltext ROWS=10 COLS=50>"
response.write server.htmlencode(strTrig)
response.write "</TEXTAREA>"
Response.write "<BR>"
Response.write "<INPUT TYPE=Submit NAME=sbmbtn VALUE=""Create this Trigger"" class=submitbtn>"
Response.write "</FORM>"
Response.write "<P>"
Response.write "SQL Server 2000 allows multiple triggers to be created for each data modification event. In SQL Server 7.0 the default behavior for CREATE TRIGGER is to add additional triggers to existing triggers, if the trigger names differ. If trigger names are the same, SQL Server returns an error message. However, if the compatibility level is equal to or less than 65, any new triggers created with the CREATE TRIGGER statement replace any existing triggers of the same type, even if the trigger names are different. "
End if
Response.Write ("</BIG></BIG></BIG></FONT></FONT><P>")
if (histhasrecords and request.form("overwriteHistRecords") = "" and not histrecscreated) Then
	call writefooter("History Table already has data")
'	response.write "dont"
end if

call writefooter("")
function NoSQLIdentifier (strSQLTableString) 
	dim tmpstrSQLTableString
	tmpstrSQLTableString = strSQLTableString
	Select Case ocdDatabaseType
	Case "Oracle"
'	if isOracle THen
		tmpstrSQLTableString = mid(strSQLTableString,instr(strSQLTableString,".") + 2)
		tmpstrSQLTableString = left(tmpstrSQLTableString, len(tmpstrSQLTableString)-1)
		NoSQLIdentifier = tmpstrSQLTableString
'	Else
'	if DatabaseType = "Access" Then
	Case "Access"
		tmpstrSQLTableString = Replace(tmpstrSQLTableString,"]","")
	tmpstrSQLTableString = Replace(tmpstrSQLTableString,"[","")
		NoSQLIdentifier = tmpstrSQLTableString
	'ElseIf DatabaseType = "SQLServer" Then
	Case "MySQL"
		tmpstrSQLTableString = Replace(tmpstrSQLTableString,"`","")
		NoSQLIdentifier = tmpstrSQLTableString 
	
	Case Else
		tmpstrSQLTableString = Replace(tmpstrSQLTableString,"""","")
		NoSQLIdentifier = tmpstrSQLTableString 
'	Else
'		NoSQLIdentifier = tmpstrSQLTableString
'	End if
'	End if
	End Select
end function
function NiceSQLIdentifier (strSQLTableString) 
	if instr(strSQLTableString,"""") = 0 and ((instr(strSQLTableString,".") = 0 and ocdDatabaseType = "Oracle") or ocdDatabaseType <> "Oracle") and instr(strSQLTableString,"[")=0 Then
	Select Case ocdDatabaseType 
		Case "Access" 
		strSQLTableString = Replace(strSQLTableString,"]","")
	strSQLTableString = Replace(strSQLTableString,"[","")
		NiceSQLIdentifier = "[" & strSQLTableString & "]"
		Case "MySQL"
		strSQLTableString = Replace(strSQLTableString,"`","")
'	strSQLTableString = Replace(strSQLTableString,"[","")
		NiceSQLIdentifier = "`" & strSQLTableString & "`"
		Case "IXS","ADSI"
		NiceSQLIdentifier = strSQLTableString 
		Case Else '"SQLServer","Oracle"
		strSQLTableString = Replace(strSQLTableString,"""","")
		NiceSQLIdentifier = """" & strSQLTableString & """"
'		Case Else
'		NiceSQLIdentifier = strSQLTableString
	End Select
	Else
		NiceSQLIdentifier = strSQLTableString
	
	End if
end function
%>      
CREATE PROCEDURE sp_blahdd
@startdate datetime = Null,
@enddate datetime = Null,
@orderid int = null
AS
if @startdate is Null 
BEGIN
set @startdate = '8/8/9999'
END
if @orderid is NOT Null 
BEGIN
if @enddate is Null 
BEGIN
SELECT * From "ocdb_History_Order Details" WHERE ocdb_History_StartDateTime <= @startdate and ocdb_History_EndDateTime > @startdate and orderid = @orderid
END
ELSE
BEGIN
SELECT * From "ocdb_History_Order Details" WHERE ocdb_History_StartDateTime >= @startdate and ocdb_History_StartDateTime < @enddate and orderid = @orderid
END
END
ELSE
BEGIN
if @enddate is Null 
BEGIN
SELECT * From "ocdb_History_Order Details" WHERE ocdb_History_StartDateTime <= @startdate and ocdb_History_EndDateTime > @startdate 
END
ELSE
BEGIN
SELECT * From "ocdb_History_Order Details" WHERE ocdb_History_StartDateTime >= @startdate and ocdb_History_StartDateTime < @enddate 
END
END
GRANT INSERT ON [dbo].[AUDIT_LOG] TO [public]
GO

DENY REFERENCES , SELECT , DELETE , UPDATE ON [dbo].[AUDIT_LOG] TO [public] CASCADE 
GO

subst suser_sname for user_name