<%
'1 Click DB ASP Library - Common SQL and ASP Functions
'copyright 1997-2003 David J. Kawliche

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

'These functions are intended for internal 1 Click DB operations only:

Function FormatForSQL(ByVal strToFormat, ByVal DatabaseType, ByVal FormatType)
	Select Case UCase(FormatType)
		Case "CLEANUSERSQL"
			If strToFormat = "" Then
				Exit Function
			End If
			Dim strReserved, arrReserved, eleReserved, arrReserved2, regEx
			strToFormat = Replace(strToFormat, vbCRLF," ")
			strToFormat = Replace(strToFormat, vbCR," ")
			strToFormat = Replace(strToFormat, vbLF," ")
			strToFormat = Replace(strToFormat, vbTab," ")
			Do While Instr(strToFormat,"  ") <> 0 
				strToFormat = Replace(strToFormat, "  "," ")
			Loop
			arrReserved = Array(";","--","DBCC ","sp_","xp_","fn_")
			arrReserved2 = Array("SELECT","FROM","ORDER\sBY","GROUP\sBY","WHERE","DELETE","DROP","UNION","OPEN","OPENQUERY","INSERT","TRUNCATE","UPDATE","BULK","EXEC","EXECUTE","INTO","KILL","OPENDATASET","OPENROWSET","OPENXML","UPDATETEXT")
			Set regEx = New RegExp
			regEx.Global = true
			regEx.IgnoreCase = true
			regEx.Pattern = "  "
'			strToFormat = regEx.Replace(strToFormat, " ") 
			for each eleReserved in arrReserved
			regEx.Pattern = eleReserved
			strToFormat = regEx.Replace(strToFormat, "' + '" & eleReserved & "' + '") 
			next
			for each eleReserved in arrReserved2
			regEx.Pattern = eleReserved & "\s"
			strToFormat = regEx.Replace(strToFormat, "' + '" & eleReserved & " ' + '") 
			next
			for each eleReserved in arrReserved2
			regEx.Pattern = eleReserved & "\("
			strToFormat = regEx.Replace(strToFormat, "' + '" & eleReserved & "(' + '") 
			next
			'regEx.Pattern =  "(NOT)?(\s*\(*)\s*(\w+)\s*(=|&lt;&gt;|&lt;|&gt;|LIKE|IN)\s*(\(([^\)]*)\)|'([^']*)'|(-?\d*\.?\d+))(\s*\)*\s*)(AND|OR)?"

			FormatForSQL = Trim(strToFormat)
		Case "ADDSQLIDENTIFIER"
			'response.write instr(strToFormat, """") = 0 And ((instr(strToFormat, ".") = 0 And DatabaseType = "Oracle") Or DatabaseType <> "Oracle") And instr(strToFormat, "[") = 0
			If instr(strToFormat, """") = 0 And ((instr(strToFormat, ".") = 0 And DatabaseType = "Oracle") Or DatabaseType <> "Oracle") And instr(strToFormat, "[") = 0 Then
				Select Case DatabaseType
					Case "Access"
						strToFormat = Replace(strToFormat, "]", "")
						strToFormat = Replace(strToFormat, "[", "")
						FormatForSQL = "[" & strToFormat & "]"
					Case "MySQL"
						strToFormat = Replace(strToFormat, "`", "")
						FormatForSQL = "`" & strToFormat & "`"
					Case "IXS", "ADSI"
						FormatForSQL = strToFormat
					Case Else				'"SQLServer","Oracle"
						strToFormat = Replace(strToFormat, """", "")
						FormatForSQL = """" & strToFormat & """"
				End Select
			Else		  'no change
				FormatForSQL = strToFormat
			End If
		Case "REMOVESQLIDENTIFIER"
			Select Case DataBaseType
				Case "Oracle"
					strToFormat = mid(strToFormat, instr(strToFormat, ".") + 2)
					strToFormat = left(strToFormat, len(strToFormat) - 1)
					FormatForSQL = strToFormat
				Case "Access"
					strToFormat = Replace(strToFormat, "]", "")
					strToFormat = Replace(strToFormat, "[", "")
					FormatForSQL = strToFormat
				Case "MySQL"
					strToFormat = Replace(strToFormat, "`", "")
					FormatForSQL = strToFormat
				Case Else
					strToFormat = Replace(strToFormat, """", "")
					FormatForSQL = strToFormat
			End Select
		Case "SAFEDATE"
			Dim varRetVal
			varRetVal = ""
			If isDate(strToFormat) Then
				varRetVal = Day(strToFormat) & "-" & UCase(Monthname(Month(strToFormat), True)) & "-" & Year(strToFormat)
			End If
			FormatForSQL = varRetVal
		Case Else
			FormatForSQL = "Unknown Format"
	End Select
End Function

Function getDatabaseType(ByRef ADOConnection)
	Select Case UCASE(ADOConnection.Provider)
		Case "ADSDSOOBJECT"
			getDatabaseType = "ADSI"
		Case Else
			Select Case UCASE(ADOConnection.Properties("DBMS Name"))
				Case "MICROSOFT INDEX SERVER"
					getDatabaseType = "IXS"
				Case "MS SQL SERVER", "MICROSOFT SQL SERVER"
					getDatabaseType = "SQLServer"
				Case "MYSQL"
					getDatabaseType = "MySQL"
				Case "MS JET", "ACCESS"
					getDatabaseType = "Access"
				Case Else
					If instr(UCASE(ADOConnection.Properties("DBMS Name")), "ORAC") = 0 Then
						getDatabaseType = "Unknown"
					Else
						getDatabaseType = "Oracle"
					End If
			End Select
	End Select
End Function

Function getPKFields(ByRef ADOConnection, ByVal DatabaseType, ByVal ObjectName, ByVal QuotePrefix, ByVal QuoteSuffix)
	On Error Resume Next
	Dim strObjectOwner, strObjectNAME, rsIDX, strSQLIDX
	strObjectOwner = GetSQLIDFPart(ObjectName, "SQLOBJECTOWNER", quoteprefix, quotesuffix)
	strObjectName = GetSQLIDFPart(ObjectName, "SQLOBJECTNAME", quoteprefix, quotesuffix)
	getPKFields = ""
	Set rsIDX = server.createobject("ADODB.Recordset")
	Select Case UCASE(DatabaseType)
		Case "ORACLE"
			Set rsIDX = server.createobject("ADODB.Recordset")
			Call rsIDX.open("SELECT ALL_CONS_COLUMNS.COLUMN_NAME AS COLNAME FROM ALL_CONSTRAINTS, ALL_CONS_COLUMNS WHERE ALL_CONSTRAINTS.OWNER = ALL_CONS_COLUMNS.OWNER AND ALL_CONSTRAINTS.CONSTRAINT_NAME = ALL_CONS_COLUMNS.CONSTRAINT_NAME AND ALL_CONSTRAINTS.TABLE_NAME = ALL_CONS_COLUMNS.TABLE_NAME AND CONSTRAINT_TYPE = 'P' AND ALL_CONS_COLUMNS.TABLE_NAME = '" & FormatForSQL(ObjectName, DatabaseType, "RemoveSQLIdentifier") & "' AND ALL_CONSTRAINTS.TABLE_NAME = '" & FormatForSQL(ObjectName, DatabaseType, "RemoveSQLIdentifier") & "'", ADOConnection)
			Do While Not rsIDX.eof
				getPKFields = getPKFields & rsIDX("COLNAME") & ","
				Call rsIDX.movenext()
			Loop
			Call rsIDX.close()
			Set rsIDX = Nothing
		Case "SQLSERVER"
			On Error Resume Next
			If strObjectOwner <> "" Then
				Set rsIDX = ADOConnection.Execute("Select COLUMN_NAME from INFORMATION_SCHEMA.TABLE_CONSTRAINTS INNER JOIN INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE On INFORMATION_SCHEMA.TABLE_CONSTRAINTS.CONSTRAINT_NAME = INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE.CONSTRAINT_NAME Where INFORMATION_SCHEMA.TABLE_CONSTRAINTS.TABLE_NAME = '" & replace(strObjectName, "'", "''") & "' AND INFORMATION_SCHEMA.TABLE_CONSTRAINTS.TABLE_SCHEMA = '" & replace(strObjectOwner, "'", "''") & "'  AND INFORMATION_SCHEMA.TABLE_CONSTRAINTS.CONSTRAINT_TYPE = 'PRIMARY KEY'")			  'indexes
			Else
				Set rsIDX = ADOConnection.Execute("select COLUMN_NAME from INFORMATION_SCHEMA.TABLE_CONSTRAINTS INNER JOIN INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE On INFORMATION_SCHEMA.TABLE_CONSTRAINTS.CONSTRAINT_NAME = INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE.CONSTRAINT_NAME where INFORMATION_SCHEMA.TABLE_CONSTRAINTS.TABLE_NAME = '" & replace(strObjectName, "'", "''") & "' AND INFORMATION_SCHEMA.TABLE_CONSTRAINTS.CONSTRAINT_TYPE = 'PRIMARY KEY'")
			End If
			If err <> 0 Then
				getPKFields = ""
				Exit Function
			Else
				Do While Not rsIDX.eof
					getPKFields = getPKFields & rsIDX.Fields("COLUMN_NAME").Value & ","
					Call rsIDX.movenext()
				Loop
				Call rsIDX.close()
				Set rsIDX = Nothing
			End If
			Err.Clear
			
		Case "ACCESS"
			If (ADOConnection.provider <> "MSDASQL.1") Then
				Set rsIDX = ADOConnection.openSchema(12, array(empty, empty, empty, empty, CStr(FormatForSQL(ObjectName, DatabaseType, "RemoveSQLIdentifier"))))		  'indexes
				Do While Not rsIDX.eof
					If UCASE(rsIDX.Fields("TABLE_NAME").Value) = UCase(FormatForSQL(ObjectName, DatabaseType, "RemoveSQLIdentifier")) And rsIDX.Fields("PRIMARY_KEY").Value = True Then
						getPKFields = getPKFields & rsIDX.Fields("COLUMN_NAME").Value & ","
					End If
					Call rsIDX.movenext()
				Loop
				Call rsIDX.close()
				Set rsIDX = Nothing
			Else
				Set rsIDX = ADOConnection.openSchema(12)				 'adSchemaindexes, usually not populated for ODBC connects
				Do While Not rsIDX.eof
					If UCASE(rsIDX.Fields("table_name").Value) = UCase(FormatForSQL(ObjectName, DatabaseType, "RemoveSQLIdentifier")) And rsIDX.Fields("primary_key").Value = True Then
						getPKFields = getPKFields & rsIDX.Fields("COLUMN_NAME").Value & ","
					End If
					Call rsIDX.movenext()
				Loop
				Call rsIDX.close()
				Set rsIDX = Nothing
			End If
		Case "MYSQL"
		Case Else
			Set rsIDX = ADOConnection.openSchema(12)			  'adSchemaindexes, usually not populated for ODBC connects
			Do While Not rsIDX.eof
				If UCASE(rsIDX.Fields("table_name").Value) = UCase(FormatForSQL(ObjectName, DatabaseType, "RemoveSQLIdentifier")) And rsIDX.Fields("primary_key").Value = True Then
					getPKFields = getPKFields & rsIDX.Fields("COLUMN_NAME").Value & ","
				End If
				Call rsIDX.movenext()
			Loop
			Call rsIDX.close()
			Set rsIDX = Nothing
	End Select
	If getPKFields <> "" Then
		getPKFields = left(getPKFields, len(getPKFields) - 1)
	End If
	If err.number <> 0 Then
		err.clear()
		getPKFields = ""
	End If
End Function

Function SafeName(ByVal strToFormat)
	Dim tmpval
	tmpval = CStr(strToFormat)
	SafeName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(tmpval, " ", "_"), "-", "_"), "#", "_"), "$", "_"), "'", "_"), """", "_"), ".", "_"), "/", "_"), "&", "_"), "*", "_"), "%", "_"), "=", "_"), "<", "_"), ">", "_"), "!", "_"), "?", "_"), "|", "_"), ":", "_"), "\", "_"),")","_"),"(","_")
End Function

Function getDataBaseDisplayName(ByVal ocdgnDatabaseType, ByVal ocdgnIsODBC, ByVal ocdgnDatabaseConn)
	Dim tmpPVar
	If ocdgnDatabaseType = "Access" Or ocdgnDatabaseType = "Excel" Then
		If ocdgnIsODBC Then
			getDataBaseDisplayName = mid(ocdgnDatabaseConn.Properties("Current Catalog"), instrrev(ocdgnDatabaseConn.Properties("Current Catalog"), "\") + 1)
		Else
			tmpPVar = mid(ocdgnDatabaseConn.Properties("Data Source"), instrrev(ocdgnDatabaseConn.Properties("Data Source"), "\") + 1)
			If instr(tmpPVar, ".") > 0 Then
				tmpPVar = left(tmpPVar, instr(tmpPvar, ".") - 1)
			End If
			getDataBaseDisplayName = tmpPVar
		End If
	ElseIf ocdgnDatabaseType = "SQLServer" Then
		getDataBaseDisplayName = ocdgnDatabaseConn.Properties("Current Catalog")
	Else
		If (ocdgnDatabaseConn.Properties("Data Source")) = "" Then
			getDataBaseDisplayName = "DB Properties"
		Else
			getDataBaseDisplayName = ocdgnDatabaseConn.Properties("Data Source")
		End If
	End If
End Function

Function GetSQLIDFPart(ByVal strSQLIdf, ByVal strIdfPart, ByVal strPrefix, ByVal strSuffix)
	Dim arrPieces, strTemp
	arrPieces = split(strSQLIdf, ".")
	strTemp = ""
	Select Case UCASE(strIDfPart)
		Case "SQLOBJECTNAME"
			If Ubound(arrPieces) > -1 Then
				strTemp = arrPieces(UBound(arrPieces))
			End If
		Case "SQLOBJECTOWNER"
			If Ubound(arrPieces) > 0 Then
				strTemp = arrPieces(UBound(arrPieces) - 1)
			End If
		Case Else
	End Select
	strTemp = replace(strTemp, strPrefix, "")
	strTemp = replace(strTemp, strSuffix, "")
	GetSQLIDFPart = strTemp
End Function

%>
