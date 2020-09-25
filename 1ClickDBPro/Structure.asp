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
<!--#INCLUDE FILE=ocdFunctions.asp-->
<%
if ocdIsHome Then
	ocdIsHome = False
End If
'on error goto 0
Dim qs, rsTable, fldTemp, connStructure, strSQL, intRowCount, rsIDX, strIDXName, rsViewType, blnIsTable, blnIsView, blnIsProc, arrNDSchema, strFullSQLIDF, strSQLObjectName, strSQLObjectOwner, strProcName, arrTemp, rsTemp, intCountFields, varValue, rsCheckFields, rsCCInf, rsTrInf, rsFK, intK, rsNDTemp1
'on error goto 0
Dim arrNDSchemaFields (1) 
Dim arrNDSchemaFields2 (2) 
blnIsTable = False
blnIsView = False
blnIsProc = False
strFullSQLIDF = request.querystring("sqlfrom")
strSQLObjectName = GetSQLIDFPart(strFullSQLIDF, "SQLOBJECTName",ocdQuoteSuffix,ocdQUotePrefix)
strSQLObjectOwner = GetSQLIDFPart(strFullSQLIDF, "SQLOBJECTOWNER",ocdQuoteSuffix,ocdQUotePrefix)
Set connStructure = server.CreateObject("ADODB.Connection")
Call connStructure.Open(ndnscSQLConnect, ndnscSQLUser, ndnscSQLPass)
Select Case ocdDatabaseType 
	Case "SQLServer", "Access"
		connStructure.CommandTimeout = ocdDBTimeout
End Select
Set rsViewType = server.createobject("ADODB.Recordset")
If ocdDatabaseType = "Oracle" Then
	rsViewType.open "SELECT OBJECT_TYPE FROM ALL_OBJECTS WHERE OBJECT_NAME ='" & NoSQLIdentifier(Request.QuerySTring("sqlfrom")) & "'", connStructure
	If rsViewType.Fields("OBJECT_TYPE").VALUE = "TABLE" Then
		blnIsTable = True
	Elseif rsViewTYpe.Fields("OBJECT_TYPE").VALUE = "VIEW" Then
		blnIsView = True
	End If
Else
	If ocdDatabaseType = "SQLServer" Then
		Set rsViewType = connStructure.Execute ("Select xtype from sysobjects inner join sysusers on sysobjects.uid = sysusers.uid WHERE sysusers.name = '" & strSQLObjectOwner & "' AND sysobjects.name = '" & strSQLObjectName & "'")
		If Not rsViewType.eof then
			Select Case Trim(rsViewType(0))
				Case "V"
					blnIsView = True
				Case "U"
					blnIsTable = True
			End Select
		End If
	Elseif ocdDatabaseType = "Access" Then
		Set rsViewTYpe = connStructure.OpenSchema(20,array(empty,empty,CSTR(NOSQLIDentifier(request.querystring("sqlfrom"))))) 'adSchemaTables
	Else
		Set rsViewTYpe = connStructure.OpenSchema(20) 'adSchemaTables
	End If
End If
If Not ocdDatabaseType = "Oracle" And Not ocdDatabaseType = "SQLServer"  Then
	Do While Not rsViewType.eof
		If UCase(rsViewType.Fields("TABLE_NAME").Value) = UCASE(NOSQLIDentifier(request.querystring("sqlfrom"))) Then
			If rsViewTYPE.Fields("TABLE_TYPE").Value = "TABLE" Then
				blnIsTable = True
			ElseIf rsViewTYPE.Fields("TABLE_TYPE").Value = "VIEW" Then
				blnIsView = True
			End If
		End If
		rsViewTYpe.movenext
	Loop
End If
If Not (blnIsTable or blnIsView) And Not ocdDatabaseType = "Excel" And Not ocdDatabaseType = "Oracle" Then
	Set rsViewTYpe = connStructure.OpenSchema(16) 'adSchemaProcedures
	Do while not rsViewType.eof
		If Replace(UCASE(rsViewType.Fields("PROCEDURE_NAME").Value),"","") = UCASE(request.querystring("sqlfrom")) Then
			blnIsProc = True
		End If
		rsViewTYpe.movenext
	Loop
End If
Set rsTable = Server.CreateObject("ADODB.Recordset")
Call writeheader("")
response.Flush
Response.Write("<SPAN CLASS=Information>")
If blnIsView Then
	Response.Write(" Views : ")
Elseif blnIsProc Then
	Response.Write(" Procedures : ")
Else
	Response.Write(" <A HREF=Schema.asp?show=tables>Tables</A> : ")
End If
If blnIsTable or blnIsView Then
	Response.Write("<A HREF=""browse.asp?sqlfrom_A=")
	Response.Write(Server.URLEncode(request.querystring("sqlfrom"))) 
	For Each QS In request.querystring
		If Not UCASE(QS) = "SQLFROM_A" And Not UCASE(QS) = "SQLFROM" Then
			Response.Write("&amp;")
			Response.Write(QS)
			Response.Write("=")
			Response.Write(server.urlencode(Request.querystring(QS)))
		End If
	Next
	Response.Write(""">")
	Response.Write(Server.HTMLEncode(Request.QueryString("sqlfrom")))
	Response.Write("</A>")
Else
	Response.Write(Server.HTMLEncode(Request.QueryString("sqlfrom")))
End If
Response.Write("</SPAN><BR>")
If blnIsView Then
	If ocdAccessTableEdits And (connStructure.Properties("DBMS Name") = "MS Jet") Then 
		Response.Write("<A HREF=DBDesignMSAccess.asp?ocdaction=editview&sqlfrom=")
		Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
		Response.Write(">(Edit SQL Text)</A> ")
	End If
	If ocdSQLTableEdits  Then
		Response.Write("<A HREF=DBDesignSQLServer.asp?action=editview&sqlfrom=")
		Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
		Response.Write(">(Edit SQL Text)</A> ")
	End If
End If
Response.flush
If blnIsProc and not (connStructure.Properties("DBMS Name") = "MS Jet" OR connStructure.Properties("DBMS Name") = "ACCESS") Then 'not blnIsProc Then
	If instr(request.querystring("Sqlfrom"),";") > 0 Then
		arrTemp = split(Cstr(request.querystring("SQLFROM")),";")
		strProcName = arrTemp(0)
	Else
		strProcName = Cstr(request.querystring("SQLFROM"))
	End If
	If ocdSQLTableEdits Then
		Response.Write("<A HREF=DBDesignSQLServer.asp?action=dropproc&table=")
		Response.Write(Server.URLEncode(strProcName))
		Response.Write(" CLASS=Menu>(Drop)</A>")
	End If
	Set rsTemp = server.createobject("ADODB.Recordset")
	call rsTemp.Open ("sp_sproc_columns @procedure_name=""" & request.querystring("sqlfrom") & """", connStructure)
	Response.Write("<table border=0 cellspacing=1 cellpadding=1><tr>")
  	'Header with name of fields
	For intCountFields=0 to rsTemp.Fields.count-1
		Response.Write("<th> " & Trim(rsTemp.Fields(intCountFields).Name) & "</th>")
   	Next
   	Response.Write("</tr>")
		Response.Write("<TR>")
		do while not rsTemp.eof
			For intCountFields=0 to rsTemp.Fields.count-1
				varValue = ""
				varValue = rsTemp.Fields(intCountFields).Value
     		Response.Write("<TD>") 
				If not isnull(varValue) Then
					Response.Write(server.htmlencode(varValue))
				End If
				Response.Write  ("</TD>")
   		Next
	   	Response.Write("</TR>")
			rsTemp.movenext
		loop
	 	Response.Write("</table>")
		If ocdSQLTableEdits and ocdShowSQLCommander Then
			Response.Write("<P>")
			Response.Write("<A HREF=DBDesignSQLServer.asp?ocdaction=EDITPROC&sqlfrom=" & server.urlencode(strProcName))
			Response.Write(">")
		End If
		Response.Write("(Edit SQL Text)")
		Response.Write("</a>")
		Response.Write("<P>Only column information about input and output parameters for the stored procedure are displayed. If the  columns are indeterminate, no information will be returned.<P>")
	Else
		strSQL = "Select " 
		If request.querystring("SQLSelect") = "" Then
			strSQL = strSQL & "*"
		Else
			strSQL = strSQL & request.querystring("SQLSelect")
		End If
		strSQL = strSQL & " from " & FormatForSQL(Request.QueryString("sqlfrom"),ocdDatabaseType,"ADDSQLIDENTIFIER") & " WHERE 1=2"
		Call rsTable.Open (strSQL, connStructure, 3,3) '0, 1, &H0001)
 ',adOpenForwardOnly, adLockReadOnly,adCmdText
		If err <> 0 Then
			Call writefooter("")
		End If
		If err <> 0 Then
			err.clear
		Elseif rsTable.fields.count > 0 Then
			Response.Write(" ")
			If rsTable.fields.count = 1 Then
				Response.Write("1 Field - ")
			Else
				Response.Write(rsTable.fields.count)
				Response.Write(" Fields - ")
			End If
			If err = 0 Then
				connStructure.CommandTimeOut = 1
				set rsCheckFields = connStructure.execute ("Select Count(*) from " & FormatForSQL(Request.QueryString("sqlfrom"),ocdDatabaseType,"ADDSQLIDENTIFIER") )
				If err = 0 Then
					Response.Write(rsCheckFields.Fields(0).Value)
					Response.Write(" Record")
					If isnull(rsCheckFields.Fields(0).Value) Then
						Response.Write("s")
					Else
						If CLNG(rsCheckFields.Fields(0).Value) <> CLNG(1) Then
							Response.Write("s")
					End If
				End If
				rsCheckFields.close
				set rsCheckFields = nothing
			Else
'				on error goto 0
				err.clear
				Response.Write("Many Records")
				set rsCheckFields = nothing
			End If
			connStructure.CommandTimeOut = ocdDBTimeout
		End If
		Response.Write(" ")
	End If
	If blnIsTable  or blnIsView or blnIsProc Then
		If ocdAccessTableEdits AND (connStructure.Properties("DBMS Name") = "MS Jet" OR connStructure.Properties("DBMS Name") = "ACCESS") Then
			If not blnIsProc Then
				Response.Write("<A HREF=DBDesignMSAccess.asp?action=copytable&table=")
				Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
				Response.Write(">(Copy)</A>")
				Response.Write(" <A HREF=DBDesignMSAccess.asp?action=deletetable&table=")
				Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
				Response.Write(">(Empty)</A>")
			End If
			If blnIsTable Then
				Response.Write(" <A HREF=DBDesignMSAccess.asp?action=droptable&table=")
				Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
				Response.Write(">(Drop)</A>")
			Elseif blnIsProc Then
				Response.Write(" <A HREF=DBDesignMSAccess.asp?action=droptable&table=")
				Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
				Response.Write(">(Drop)</A>")
			Else
				Response.Write(" <A HREF=DBDesignMSAccess.asp?action=dropview&table=")
				Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
				Response.Write(">(Drop)</A>")
			End If
			Response.Write(" ")
		End If
		If ocdSQLTableEdits  Then
			Response.Write("<A HREF=DBDesignSQLServer.asp?action=scriptobject&sqlfrom=")
			Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
			Response.Write(">(Generate SQL)</A> ")
			If blnIsTable Then
				Response.Write(" <A HREF=""DBDesignSQLServer.asp?action=renametable&amp;sqlfrom=")
				Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
				Response.Write(""">(Rename)</A>")
			Elseif blnIsView Then
				Response.Write(" <A HREF=""DBDesignSQLServer.asp?action=renameview&amp;sqlfrom=")
				Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
				Response.Write(""">(Rename)</A>")
			End If
			Response.Write(" <A HREF=""DBDesignSQLServer.asp?action=copytable&amp;sqlfrom=")
			Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
			Response.Write(""">(Copy)</A>")
			Response.Write(" <A HREF=""DBDesignSQLServer.asp?action=deletetable&amp;sqlfrom=")
			Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
			Response.Write(""">(Empty)</A>")
			If blnIsTable Then
				Response.Write(" <A HREF=DBDesignSQLServer.asp?action=droptable&amp;sqlfrom=")
				Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
				Response.Write(">(Drop)</A><BR>")
			Else
				Response.Write(" <A HREF=DBDesignSQLServer.asp?action=dropview&sqlfrom=")
				Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
				Response.Write(">(Drop)</A><BR>")
			End If
		End If
	End If
	If (ocdSQLTableEdits AND (ocdDatabaseType = "SQLServer") and ocdShowSQLCommander) AND (blnIsTable ) Then
		If blnIsTable Then
					Response.Write("<A HREF=Command.asp?sqltext=")
			Response.Write(server.urlencode("sp_help ('") & Server.URLEncode(Request.Querystring("sqlfrom")) & Server.URLEncode("')"))
			Response.Write("><FONT SIZE=1>(Describe)</FONT></a>&nbsp;&nbsp;&nbsp;")
			Response.Write("<A HREF=Command.asp?sqltext=")
			Response.Write(server.urlencode("DBCC CHECKTABLE ('") & Server.URLEncode(Request.Querystring("sqlfrom")) & Server.URLEncode("')"))
			Response.Write("><FONT SIZE=1>(Check)</FONT></a>&nbsp;&nbsp;&nbsp;")
		End If
		Response.Write("<A HREF=Command.asp?sqltext=")
		Response.Write(server.urlencode("DBCC UPDATEUSAGE ('") & Server.URLEncode(connStructure.Properties("Current Catalog")) & server.urlencode("','") & Server.URLEncode(Request.Querystring("sqlfrom")) & Server.URLEncode("')"))
		Response.Write("><FONT SIZE=1>(Update&nbsp;Usage)</FONT></a>&nbsp;&nbsp;&nbsp;")
		Response.Write("<A HREF=Command.asp?sqltext=")
		Response.Write(server.urlencode("UPDATE STATISTICS ") & Server.URLEncode(Request.Querystring("sqlfrom")) & Server.URLEncode(""))
		Response.Write("><FONT SIZE=1>(Update&nbsp;Stats)</FONT></a>&nbsp;&nbsp;&nbsp;")
		If blnIsView then
			Response.Write "<FONT SIZE=1>Not supported by all editions of SQL Server</FONT>"
		End If
	End If
	If (ocdSQLTableEdits AND (ocdDatabaseType = "SQLServer") and ocdShowSQLCommander) AND (blnIsTable) Then
		If ocdDBMSVersion > CDbl(6.9) Then
			If ocdDBMSVersion > CDbl(7.9) Then
				Response.Write("<A HREF=Command.asp?sqltext=")
				Response.Write(server.urlencode("DBCC CHECKCONSTRAINTS ('") & Server.URLEncode(Request.Querystring("sqlfrom")) & Server.URLEncode("')"))
				Response.Write("><FONT SIZE=1>(Check&nbsp;Constraints)</FONT></a>&nbsp;&nbsp;&nbsp;")
				Response.Write("<A HREF=Command.asp?sqltext=")
				Response.Write(server.urlencode("DBCC SHOWCONTIG ('") & Server.URLEncode(Request.Querystring("sqlfrom")) & Server.URLEncode("')"))
				Response.Write("><FONT SIZE=1>(Show&nbsp;Contig)</FONT></a>&nbsp;&nbsp;&nbsp;")
				Response.Write("<A HREF=Command.asp?sqltext=")
				Response.Write(server.urlencode("DBCC CLEANTABLE ('") & Server.URLEncode(connStructure.Properties("Current Catalog")) & server.urlencode("','") & Server.URLEncode(Request.Querystring("sqlfrom")) & Server.URLEncode("')"))
				Response.Write("><FONT SIZE=1>(Clean)</FONT></a>&nbsp;&nbsp;&nbsp;")
		End If
	End If
end	If
if (ocdSQLTableEdits AND (ocdDatabaseType = "SQLServer")) AND (blnIsTable or blnIsView) Then
	Response.Write("<BR>")
End If
Response.Write("<P><TABLE  CLASS=Grid><TR CLASS=GridHeader>")
If blnIsTable Then
	If ocdAccessTableEdits AND (connStructure.Properties("DBMS Name") = "MS Jet" OR connStructure.Properties("DBMS Name") = "ACCESS") Then
		Response.Write("<TH ALIGN=LEFT>&nbsp;</TH>")
	End If
	If ocdSQLTableEdits  Then
		Response.Write("<TH ALIGN=LEFT>&nbsp;</TH>")
	End If
End If
Response.Write("<TH ALIGN=LEFT>Field&nbsp;Name</TH><TH ALIGN=LEFT>Type</TH><TH ALIGN=LEFT>Size</TH><TH ALIGN=LEFT>Required</TH></TR>")
intRowCount=0
for each fldTemp in rsTable.Fields
	Response.Write("<TR")
	If intRowCOunt mod 2 <> 0 Then
		Response.Write(" class=GridEven ")
	Else
		Response.Write(" class=GridOdd ")
	End If
	Response.Write(">")
	If blnIsTable Then
		If ocdAccessTableEdits AND ocdDatabaseType = "Access" Then
			Response.Write("<TD VALIGN=MIDDLE ALIGN=CENTER>")
				If connStructure.provider ="Microsoft.Jet.OLEDB.4.0" Then
					Response.Write("<A HREF=DBDesignMSAccess.asp?ocdaction=EditField&fieldname=")
					Response.Write(Server.URLEncode(fldTemp.Name))
					Response.Write("&defsize=" & fldTemp.DefinedSize & "&deftype=")
					Response.Write(fldTemp.Type)
					Response.Write("&nulldef=")
					If fldTemp.Attributes and &H00000020 Then 'adFldIsNullable
						Response.Write("NULL")
					Else
						Response.Write("NOTNULL")
					End If
					Response.Write("&table=")
					Response.Write(Server.URLEncode(request.querystring("sqlfrom")))
					Response.Write("><IMG SRC=GRIDLNKEDIT.GIF BORDER=0 ALT=Modify></a>&nbsp;")
				End If
				Response.Write("<A HREF=DBDesignMSAccess.asp?action=deletefield&fieldname=")
				Response.Write(Server.URLEncode(fldTemp.Name))
				Response.Write("&table=")
				Response.Write(Server.URLEncode(Request.Querystring("sqlfrom")))
				Response.Write("><IMG SRC=GRIDLNKDELETE.GIF Border=0 Alt=Drop></a>")
				Response.Write("</TD>")
			End If
			If ocdSQLTableEdits  Then
				Response.Write("<TD ALIGN=CENTER>")
					If ocdDBMSVersion > CDbl(6.9) Then
						Response.Write("<A HREF=DBDesignSQLServer.asp?ocdaction=EditField&fieldname=")
						Response.Write(Server.URLEncode(fldTemp.Name))
						Response.Write("&defsize=" & fldTemp.DefinedSize & "&deftype=")
						Response.Write(fldTemp.Type)
						Response.Write("&defscale=" & fldTemp.NumericScale)
						Response.Write("&defprecision=" & fldTemp.Precision)
						Response.Write("&nulldef=")
						If fldTemp.Attributes and &H00000020 Then 'adFldIsNullable
							Response.Write("NULL")
						Else
							Response.Write("NOTNULL")
						End If
						Response.Write("&table=")
						Response.Write(Server.URLEncode(request.querystring("sqlfrom")))
						If fldTemp.Properties("ISAUTOINCREMENT") = "True" Then			
							Response.Write "&defidentity=on"
						End If
						Response.Write("><IMG BORDER=0 ALT=Modify SRC=GRIDLNKEDIT.GIF></a>")
					End If
					Response.Write(" <A HREF=DBDesignSQLServer.asp?action=deletefield&fieldname=")
					Response.Write(Server.URLEncode(fldTemp.Name))
					Response.Write("&sqlfrom=")
					Response.Write(Server.URLEncode(Request.Querystring("sqlfrom")))
					Response.Write("><IMG BORDER=0 ALT=DROP SRC=GRIDLNKDELETE.GIF></a> ")
					If fldTemp.Properties("ISAUTOINCREMENT") = "True" Then
						Response.Write("<BR><A HREF=Command.asp?sqltext=")
						Response.Write(Server.URLEncode("DBCC CHECKIDENT ('") & Server.URLEncode(request.querystring("sqlfrom")) & Server.URLEncode("', RESEED)"))
						Response.Write("><FONT SIZE=1>(Reseed)</FONT></a>&nbsp;")
					End If
					Response.Write("</TD>")
				End If
			End If
			Response.Write("<TD NOWRAP>")
			Response.Write(fldTemp.Name)
			Response.Write("</TD><TD>")
			select case fldTemp.Type
				case 0 'adEmpty
					Response.Write("Empty")
				case 16 'adTinyInt 
					Response.Write("TinyInt")
				case 2 'adSmallInt
					Response.Write("SmallInt")
				case 3 'adInteger 
					Response.Write("Integer")
				case 20 'adBigInt 
					Response.Write("BigInt")
				case 17 'adUnsignedTinyInt 
					Response.Write("UnsignedTinyInt")
				case 18 'adUnsignedSmallInt
					Response.Write("UnsignedSmallInt")
				case 19 'adUnsignedInt
					Response.Write("UnsignedInt")
				case 21 'adUnsignedBigInt
					Response.Write("UnsignedBigInt")
				case 4 'adSingle
					Response.Write("Single")
				case 5 'adDouble
					Response.Write("Double")
				case 6 'adCurrency
					Response.Write("Currency")
				case 14 'adDecimal
					Response.Write("Decimal")
				case 131 'adNumeric
					Response.Write("Numeric")
				case 11 'adBoolean
					Response.Write("Boolean")
				case 10 'adError
					Response.Write("Error")
				case 132 'adUserDefined
					Response.Write("UserDefined")
				case 12 'adVariant
					Response.Write("Variant")
				case 9 'adIDispatch
					Response.Write("IDispatch")
				case 13 'adIUnknown
					Response.Write("IUnknown")
				case 72 'adGUID
					Response.Write("GUID")
				case 7 'adDate
					Response.Write("Date")
				case 133 'adDBDate
					Response.Write("DBDate")
				case 134 'adDBTime
					Response.Write("DBTime")
				case 135 'adDBTimeStamp
					Response.Write("DBTimeStamp")
				case 8 'adBSTR
					Response.Write("BSTR")
				case 129 'adChar
					Response.Write("Char")
				case 200 'adVarChar
					Response.Write("VarChar")
				case 201 'adLongVarChar
					Response.Write("LongVarChar")
				case 130 'adWChar
					Response.Write("WChar")
				case 202 'adVarWChar
					Response.Write("VarWChar")
				case 203 'adLongVarWChar
					Response.Write("LongVarWChar")
				case 128 'adBinary 
					Response.Write("Binary")
				case 204 'adVarBinary
					Response.Write("VarBinary")
				case 205 'adLongVarBinary
					Response.Write("LongVarBinary")
				case else
					Response.Write "Type" & fldTemp.Type
			end select
			Response.Write("</TD><TD>")
			If fldTemp.DefinedSize < 12000 Then
				Response.Write(fldTemp.DefinedSize)
			Else
				Response.Write(fldTemp.DefinedSize\(1024 * 1024))
				Response.Write "Mb"
			End If
			Response.Write("</TD><TD>")
			If CBool(fldTemp.Attributes and &H00000020) Then 'adFldIsNullable
				Response.Write False
			Else
				Response.Write True
			End If
			Response.Write("</TD></TR>")
			intRowCount = intRowCount + 1
		next
		Response.Write("</table>")
	End If
	response.flush
	If blnIsTable Then
		If ocdAccessTableEdits AND (connStructure.Properties("DBMS Name") = "MS Jet" OR connStructure.Properties("DBMS Name") = "ACCESS") Then
			Response.Write("<A HREF=DBDesignMSAccess.asp?ocdaction=NEWFIELD&sqlfrom=")
			Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
			Response.Write(" Class=Menu><IMG SRC=appNew.gif Border=0 ALT=New Field> New Field</A><p>")
		End If
		If ocdSQLTableEdits Then
			Response.Write("<A HREF=DBDesignSQLServer.asp?ocdaction=NEWFIELD&sqlfrom=")
			Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
			Response.Write(" CLASS=Menu><IMG BORDER=0 ALT=New SRC=AppNew.gif> New Field</A><p>")
		End If
	End If
	If blnIsView Then
		If ocdAccessTableEdits AND (connStructure.Properties("DBMS Name") = "MS Jet") Then ' OR connStructure.Properties("DBMS Name") = "ACCESS") Then
			Response.Write("<A HREF=DBDesignMSAccess.asp?ocdaction=EditView&sqlfrom=")
			Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
			Response.Write(">(Edit SQL Text)</A>")
		End If
	If ocdSQLTableEdits Then
		Response.Write("<A HREF=DBDesignSQLServer.asp?action=editview&sqlfrom=")
		Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
		Response.Write(">(Edit SQL Text)</A><p>")
	End If
End If
If blnIsProc Then
	If ocdAccessTableEdits AND (connStructure.Properties("DBMS Name") = "MS Jet") Then' OR connStructure.Properties("DBMS Name") = "ACCESS") Then
		Response.Write("<A HREF=DBDesignMSAccess.asp?ocdaction=EditProc&sqlfrom=")
		Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
		Response.Write(">(Edit Procedure)</A><p>")
	End If
End If
	If not response.isclientconnected then 
		call CleanUpObjects
	End If


'Table Indexes
If not ocdReadOnly and blnIsTable and not ocdDatabaseType = "Oracle" and not ocdDatabaseType = "MySQL" Then
	set rsIDX = server.createobject("ADODB.Recordset")
	If ocdDatabaseType = "SQLServer" Then
	set rsIDX = connStructure.Execute("sp_helpindex '" & ocdQuotePrefix & strSQLObjectOwner & ocdQuoteSuffix & "." & ocdQuotePrefix & strSQLObjectName & ocdQUoteSuffix & "'") 'adSchemaIndexes
	Elseif ocdDatabaseType = "Access"  Then
		set rsIDX = connStructure.openSchema(12,array(empty,empty,empty,empty,cstr(replace(replace(request.querystring("sqlfrom"),"[",""),"]","")))) 'adSchemaIndexes
	Else
	set rsIDX = connStructure.openSchema(12) 'adSchemaIndexes
	End If
	strIDXName = ""
	Response.Write("<P><SPAN CLASS=Information>Indexes")
	Response.Write(" </SPAN><BR>")
	If ocdDataBaseType = "SQLServer" Then
		Response.Write("<table Class=Grid><tr Class=GridHeader><TH ALIGN=LEFT>&nbsp;</TH><TH ALIGN=LEFT>Name</TH><TH ALIGN=LEFT>Description</TH><TH ALIGN=LEFT>Index Columns</TH></tr>")
	intRowCount = 0 
	If rsIDX.State <> 0 THen
	do while not rsIDX.eof
		intRowCount = intRowCOunt + 1
		Response.Write "<TR "
		If intRowCount Mod 2 <> 0 then
				Response.Write("class=GridOdd")
			Else
				Response.Write("class=GridEven")
			End If
			
			Response.Write ">"
		Response.Write "<TD VALIGN=TOP ALIGN=CENTER>"
	If ocdSQLTableEdits Then
				Response.Write("<A HREF=DBDesignSQLServer.asp?action=deleteindex") 				
				Response.Write("&amp;indexname=" & Server.URLEncode(rsIDX.Fields("index_name").Value) & "&amp;sqlfrom=" & Server.URLEncode(Request.Querystring("sqlfrom")) & "><FONT SIZE=1>(Drop)</FONT></a>")
				If ocdShowSQLCommander Then
				Response.Write("&nbsp;<A HREF=Command.asp?sqltext=")
				Response.Write(server.urlencode("DBCC SHOW_STATISTICS ('") & Server.URLEncode(Request.Querystring("sqlfrom")) & "','" & Server.URLEncode(rsIDX.Fields("index_name").Value) & "')")
				Response.Write("><FONT SIZE=1>(Stats)</FONT></a><BR>")
			If ocdDBMSVersion > CDbl(7.9) Then
				Response.Write("<A HREF=Command.asp?sqltext=")
				Response.Write(server.urlencode("DBCC INDEXDEFRAG (0,'") & Server.URLEncode(Request.Querystring("sqlfrom")) & "','" & Server.URLEncode(rsIDX.Fields("index_name").Value) & "')")
				Response.Write("><FONT SIZE=1>(Defrag)</FONT></a>&nbsp;")
			End If
					Response.Write("<A HREF=Command.asp?sqltext=")
				Response.Write(server.urlencode("DBCC DBREINDEX ('") & Server.URLEncode(Request.Querystring("sqlfrom")) & "','" & Server.URLEncode(rsIDX.Fields("index_name").Value) & "')")
				Response.Write("><FONT SIZE=1>(Reindex)</FONT></a> ")
				End If
			else
				Response.Write("&nbsp;")
			End If
			Response.Write "</TD>"
			Response.Write "<TD VALIGN=TOP>"
			Response.Write rsIdx(0)
			Response.Write "</TD>"
			Response.Write "<TD VALIGN=TOP>"
			Response.Write rsIdx(1)
			Response.Write "</TD>"
			Response.Write "<TD VALIGN=TOP>"
			Response.Write rsIdx(2)
			Response.Write "</TD>"
			Response.Write "</TR>"
		rsIDX.moveNext
		Loop
		End If
		Response.Write "</TABLE><P>"
	Else
	Response.Write("<table Class=Grid><tr Class=GridHeader><TH ALIGN=LEFT>&nbsp;</TH><TH ALIGN=LEFT>Index Name</TH><TH ALIGN=LEFT>Unique</TH><TH ALIGN=LEFT>Primary Key</TH><TH ALIGN=LEFT>Index Columns</TH></tr>")
	intRowCOunt = 0
	do while not rsIDX.eof
		If UCASE(rsIDX.Fields("table_name").Value) = UCase(Replace(replace(request.querystring("sqlfrom"),"[",""),"]","")) and UCase(strIDXName) <> UCase(rsIDX.Fields("index_name").Value) then
			strIDXName = rsIDX.Fields("index_name").Value
			Response.Write("<tr ")
			If intRowCount Mod 2 = 0 then
				Response.Write("class=GridOdd")
			Else
				Response.Write("class=GridEven")
			End If
			Response.Write("><td NOWRAP ALIGN=CENTER VALIGN=TOP>")
			If ocdAccessTableEdits AND (connStructure.Properties("DBMS Name") = "MS Jet" OR connStructure.Properties("DBMS Name") = "ACCESS") Then
				Response.Write("<A HREF=DBDesignMSAccess.asp?action=deleteindex&indexname=" & Server.URLEncode(rsIDX("index_name")) & "&table=" & Server.URLEncode(Request.Querystring("sqlfrom")) & "><IMG SRC=GRIDLNKDELETE.GIF ALT=Drop BORDER=0></a>")
			Elseif ocdSQLTableEdits Then
				Response.Write("<A HREF=DBDesignSQLServer.asp?action=deleteindex&primarykey=")
				Response.Write(rsIDX.Fields("primary_key").Value)
				Response.Write("&indexname=" & Server.URLEncode(rsIDX.Fields("index_name").Value) & "&table=" & Server.URLEncode(Request.Querystring("sqlfrom")) & "><FONT SIZE=1>(Drop)</FONT></a>")
				If ocdShowSQLCommander Then
				Response.Write("&nbsp;<A HREF=Command.asp?sqltext=")
				Response.Write(server.urlencode("DBCC SHOW_STATISTICS (""") & Server.URLEncode(Request.Querystring("sqlfrom")) & """,""" & Server.URLEncode(rsIDX.Fields("index_name").Value) & """)")
				Response.Write("><FONT SIZE=1>(Stats)</FONT></a><BR>")
			If ocdDBMSVersion > CDbl(7.9) Then
				Response.Write("<A HREF=Command.asp?sqltext=")
				Response.Write(server.urlencode("DBCC INDEXDEFRAG (0,""") & Server.URLEncode(Request.Querystring("sqlfrom")) & """,""" & Server.URLEncode(rsIDX.Fields("index_name").Value) & """)")
				Response.Write("><FONT SIZE=1>(Defrag)</FONT></a>&nbsp;")
			End If
					Response.Write("<A HREF=Command.asp?sqltext=")
				Response.Write(server.urlencode("DBCC DBREINDEX (""") & Server.URLEncode(Request.Querystring("sqlfrom")) & """,""" & Server.URLEncode(rsIDX.Fields("index_name").Value) & """)")
				Response.Write("><FONT SIZE=1>(Reindex)</FONT></a> ")
				End If
			else
				Response.Write("&nbsp;")
			End If
			Response.Write("</td><td ALIGN=LEFT VALIGN=TOP>")
			Response.Write(rsIDX.Fields("index_name").value)
			Response.Write("</td><td ALIGN=LEFT VALIGN=TOP>")
			Response.Write(rsIDX.Fields("unique").Value)
			Response.Write("</td><td ALIGN=LEFT VALIGN=TOP>")
			Response.Write(rsIDX.Fields("primary_key").Value)
			Response.Write("</td><td ALIGN=LEFT VALIGN=TOP>")
			Response.Write(rsIDX.Fields("column_name").Value)
		Elseif UCase(rsIDX.Fields("table_name").Value) = UCase(replace(replace(request.querystring("sqlfrom"),"]",""),"[","")) then
			Response.Write("<BR>")
			Response.Write(rsIDX.Fields("column_name").Value)
		End If
		rsIDX.movenext
		If not rsIDX.eof then
			If Ucase(strIDXName) = UCase(rsIDX.Fields("index_name").Value) and UCase(rsIDX.Fields("table_name").Value) = UCase(request.querystring("sqlfrom")) Then
			Elseif UCase(rsIDX.Fields("table_name").Value) = UCase(request.querystring("sqlfrom")) Then
				Response.Write("</TD></TR>")
				intRowCount = intRowCount + 1
			End If
		Else
			Response.Write("</TD></TR>")
			intRowCount = intRowCount + 1
		End If
	loop
	Response.Write("</table>")
	End If
	If rsIDX.state <> 0 Then
	rsIDX.close
	End If
	set rsIDX = nothing
	If connStructure.Provider = "MSDASQL.1" AND ((connStructure.Properties("DBMS Name") = "MS Jet" OR connStructure.Properties("DBMS Name") = "ACCESS")) Then
		Response.Write("Indexes can be created but not viewed using ODBC.  Also related tables can not be detected and existing fields may not be modified with ODBC. For better functionality use JET OLEDB 4.0 drivers instead<BR>")
	End If
	If ocdAccessTableEdits AND (connStructure.Properties("DBMS Name") = "MS Jet" OR connStructure.Properties("DBMS Name") = "ACCESS") Then
		Response.Write("<A HREF=DBDesignMSAccess.asp?ocdaction=NEWINDEX&sqlfrom=")
		Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
		Response.Write(" CLASS=Menu><IMG border=0 SRC=appNew.gif ALT=New> New Index</A><p>")
	End If
	If ocdSQLTableEdits Then
		Response.Write("<A HREF=DBDesignSQLServer.asp?ocdaction=NEWINDEX&sqlfrom=")
		Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
		Response.Write(" Class=Menu><IMG BORDER=0 ALT=New SRC=appNew.gif> New Index</A><p>")
	End If
	Response.flush
	If not response.isclientconnected then 
		call CleanUpObjects
	End If

	If blnIsTable and ocdDatabaseType = "SQLServer" Then
		Response.Write("<P><SPAN CLASS=Information>Check Constraints")
		Response.Write("</SPAN><BR><table Class=Grid><tr Class=GridHeader><TH NOWRAP ALIGN=LEFT>&nbsp;</TH><TH NOWRAP ALIGN=Left>Constraint Name</th><TH NOWRAP ALIGN=LEFT></TH><TH NOWRAP ALIGN=LEFT></TH><TH NOWRAP ALIGN=Left></th>")
		Response.Write("</tr>")
		set rsCCInf = server.createobject("ADODB.Recordset")
		set rsCCInf = connStructure.Execute ("Select sysobjects.name as ""Constraint_Name"" from sysobjects inner join sysusers on sysobjects.uid = sysusers.uid WHERE sysusers.name = '" & strSQLObjectOwner & "' AND sysobjects.name = '" & strSQLObjectName & "' AND (sysobjects.xtype) Like 'C%'")
 'adschematableconstraints
		do while not rsCCInf.eof
			Response.Write("<TR>")
			Response.Write("<TD>")
			Response.Write("&nbsp;")
			Response.Write("</TD>")
			Response.Write("<TD>")
			Response.Write(Server.HTMLEncode(rsCCInf.Fields("constraint_name").Value))
			Response.Write("</td>")
			Response.Write("<TD>")
			Response.Write("</td>")
			Response.Write("<TD>")
			Response.Write("</td>")
			Response.Write("<TD>")
			Response.Write("</td>")
			Response.Write("</TR>")
			rsCCInf.movenext
		loop
		Response.Write("</TABLE>")
				If ocdSQLTableEdits and ocdShowSQLCommander Then
		Response.Write("<A HREF=Command.asp?proposedsqltext=")
		Response.Write(server.urlencode("ALTER TABLE " & (request.querystring("Sqlfrom")) & " ADD CONSTRAINT [CONSTRAINT NAME] CHECK ( expression ) "))
		Response.Write(" Class=Menu><IMG BORDER=0 ALT=New SRC=AppNew.gif> New Constraint</A>")
		End If
'		rsTRInf.close
'		set rsTRInf = nothing
'	End If
		rsCCInf.close
		set rsCCInf = nothing
		Response.flush
	End If
	If not response.isclientconnected then 
		call CleanUpObjects
	End If

	If blnIsTable and ocdDatabaseType = "SQLServer" Then
If ocdDBMSVersion > CDbl(6.9) Then
		Response.Write("<P><SPAN CLASS=Information>Triggers")
		Response.Write("</SPAN><BR><table Class=Grid><tr Class=GridHeader><TH NOWRAP ALIGN=LEFT>&nbsp;</TH><TH NOWRAP ALIGN=Left>Trigger Name</th><TH NOWRAP ALIGN=LEFT>Delete</TH><TH NOWRAP ALIGN=LEFT>Insert</TH><TH NOWRAP ALIGN=Left>Update</th>")
		Response.Write("</tr>")
		set rsTrInf = server.createobject("ADODB.Recordset")
		call rsTRInf.open ("sp_helptrigger '" & request.querystring("sqlfrom") & "'", connStructure)
		do while not rsTrInf.eof
			Response.Write("<TR>")
			Response.Write("<TD>")
			If ocdSQLTableEdits  Then
'				Response.Write("<FONT SIZE=1>")
				Response.Write("<A HREF=DBDesignSQLServer.asp?action=edittrig&amp;sqlfrom=")
				Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
				Response.Write("&amp;trigName=" & server.URLEncode(ocdQuotePrefix &  rsTrInf.FIelds("trigger_owner") & ocdqUOTEsuffix & "." & ocdQuotePrefix & rsTrInf.Fields("trigger_name").Value & ocdqUOTEsuffix ) )
				Response.Write("><IMG BORDER=0 ALT=Modify SRC=GRIDLNKEDIT.GIF></A>")
				Response.Write(" ")
				Response.Write("<A HREF=DBDesignSQLServer.asp?action=droptrig&sqlfrom=")
				Response.Write(Server.URLEncode(Request.querystring("sqlfrom")))
				Response.Write("&trigName=" & server.URLEncode(rsTrInf.Fields("trigger_name").Value))
				Response.Write("><IMG BORDER=0 ALT=DROP SRC=GRIDLNKDELETE.GIF></A>")

			else
				Response.Write("&nbsp;")
			End If
			Response.Write("</TD>")
			Response.Write("<TD>")
			Response.Write(Server.HTMLEncode(rsTrInf.Fields("trigger_name").Value))
			Response.Write("</td>")
			Response.Write("<TD>")
			Response.Write(Server.HTMLEncode(rsTrInf.Fields("isdelete").Value))
			Response.Write("</td>")
			Response.Write("<TD>")
			Response.Write(Server.HTMLEncode(rsTrInf.Fields("isinsert").Value))
			Response.Write("</td>")
			Response.Write("<TD>")
			Response.Write(Server.HTMLEncode(rsTrInf.Fields("isupdate").Value))
			Response.Write("</td>")
			Response.Write("</TR>")
			rsTrInf.movenext
		loop
		Response.Write("</TABLE>")
		'Response.Write("<BR>")
		If ocdSQLTableEdits and ocdShowSQLCommander Then
		Response.Write("<A HREF=Command.asp?proposedsqltext=")
		Response.Write(server.urlencode("CREATE TRIGGER [TRIGGER NAME] ON " & (request.querystring("Sqlfrom")) & " FOR INSERT, UPDATE, DELETE AS "))
		Response.Write(" CLASS=Menu><IMG BORDER=0 ALT=New SRC=AppNew.gif> New Trigger</A>")
		End If
		rsTRInf.close
		set rsTRInf = nothing
	End If
	End If
	response.flush
	If not response.isclientconnected then 
		call CleanUpObjects
	End If

	If (connStructure.provider ="Microsoft.Jet.OLEDB.4.0" OR connStructure.provider ="SQLOLEDB.1") Then 
		Response.Write("<P><SPAN CLASS=Information>Related Tables ")
		Response.Write("</SPAN><BR><table Class=Grid><tr Class=GridHeader><TH NOWRAP ALIGN=LEFT>&nbsp;</TH><TH NOWRAP ALIGN=Left>Key Name</th><TH NOWRAP ALIGN=LEFT>Related Table</TH><TH NOWRAP ALIGN=LEFT>Relationship</TH><TH NOWRAP ALIGN=Left>Update Rule</th><TH NOWRAP ALIGN=LEFT>Delete Rule</th><TH NOWRAP ALIGN=LEFT>Key Columns</TH>")
		Response.Write("</tr>")
		intRowCOunt = 1
		set rsFK = server.createobject("ADODB.Recordset")
		set rsFK = connStructure.OpenSchema(27) 'foreign keys
		If not rsFK.EOF THEN
			do while not rsFK.eof
				If ocdDatabaseType = "SQLServer" Then
								If (rsFK.FIelds("FK_TABLE_SCHEMA").Value) = strSQLObjectOwner and (rsFK.FIelds("FK_TABLE_NAME").Value) = strSQLObjectName and UCase(strIDXName) <> Ucase(rsFK.Fields("PK_TABLE_NAME").Value) then
					strIDXName = rsFK.Fields("PK_TABLE_NAME").Value
					Response.Write("<tr ")
					If intRowCount Mod 2 = 0 then
						Response.Write("class=GridOdd")
					Else
						Response.Write("class=GridEven")
					End If
					Response.Write("><td VALIGN=TOP>")
					If (connStructure.Properties("DBMS Name") = "MS Jet" OR connStructure.Properties("DBMS Name") = "ACCESS") and ocdAccessTableEdits Then
						Response.Write("<FONT SIZE=1><A HREF=DBDesignMSAccess.asp?action=deleterelation&relationname=" & Server.URLEncode(rsFK.Fields("FK_NAME").Value) & "&table=" & Server.URLEncode(Request.Querystring("sqlfrom")) & "><IMG SRC=GRIDLNKDELETE.GIF border=0 ALT=Drop></a></FONT>")
					Elseif (connStructure.Provider = "SQLOLEDB.1" and ocdSQLTableEdits) Then
						Response.Write("<FONT SIZE=1><A HREF=DBDesignSQLServer.asp?action=deleterelation&relationname=")
						Response.Write(server.urlencode(rsFK.Fields("FK_NAME").Value))
						Response.Write("&sqlfrom=" & Server.URLEncode(Request.Querystring("sqlfrom")) & "><IMG SRC=GRIDLNKDELETE.GIF border=0 ALT=Drop></a></FONT>")
					else
						Response.Write("&nbsp;")
					End If
					Response.Write("</td><td ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("FK_NAME").Value)
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("PK_TABLE_NAME").Value)
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write("Many-to-One")
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("UPDATE_RULE").Value)
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("DELETE_RULE").Value)
					Response.Write("</td><td ALIGN=LEFT NOWRAP VALIGN=TOP>")
					Response.Write(rsFK.Fields("FK_COLUMN_NAME").Value & " &lt;&mdash; " & rsFK.Fields("PK_COLUMN_NAME").Value)
				Elseif (rsFK.FIelds("FK_TABLE_SCHEMA").Value) = strSQLObjectOwner and (rsFK.FIelds("FK_TABLE_NAME").Value) = strSQLObjectName and UCase(strIDXName) = Ucase(rsFK.Fields("PK_TABLE_NAME").Value) then
					Response.Write("<BR>")
					Response.Write(rsFK.Fields("FK_COLUMN_NAME").Value & " &lt;&mdash; " & rsFK.Fields("PK_COLUMN_NAME").Value)
				Else
				End If
				rsFK.movenext
				If not rsFK.eof then
					If UCase(strIDXName) = UCase(rsFK.Fields("PK_TABLE_NAME").Value) and (rsFK.FIelds("FK_TABLE_SCHEMA").Value) = strSQLObjectOwner and (rsFK.FIelds("FK_TABLE_NAME").Value) = strSQLObjectName Then
					Elseif (rsFK.Fields("FK_TABLE_SCHEMA").Value) = (strSQLObjectOwner) and (rsFK.Fields("FK_TABLE_NAME").Value) = (strSQLObjectName) and UCase(strIDXName) <> UCase(rsFK.Fields("PK_TABLE_NAME").Value) Then
					Response.Write("</TD></TR>")
					intRowCount = intRowCount + 1
					Else
					End If
				Else
					Response.Write("</TD></TR>")
					intRowCount = intRowCount + 1
				End If
				'****************
				
				Else
				If Ucase(rsFK.Fields("FK_TABLE_NAME").Value) = UCASE(replace(replace(request.querystring("sqlfrom"),"]",""),"[","")) and UCase(strIDXName) <> Ucase(rsFK.Fields("PK_TABLE_NAME").Value) then
					strIDXName = rsFK.Fields("PK_TABLE_NAME").Value
					Response.Write("<tr ")
					If intRowCount Mod 2 = 0 then
						Response.Write("class=GridOdd")
					Else
						Response.Write("class=GridEven")
					End If
					Response.Write("><td VALIGN=TOP>")
					If (connStructure.Properties("DBMS Name") = "MS Jet" OR connStructure.Properties("DBMS Name") = "ACCESS") and ocdAccessTableEdits Then
						Response.Write("<FONT SIZE=1><A HREF=DBDesignMSAccess.asp?action=deleterelation&relationname=" & Server.URLEncode(rsFK.Fields("FK_NAME").Value) & "&sqlfrom=" & Server.URLEncode(Request.Querystring("sqlfrom")) & "><IMG SRC=GRIDLNKDELETE.GIF border=0 ALT=Drop></a></FONT>")
					Elseif (connStructure.Provider = "SQLOLEDB.1" and ocdSQLTableEdits) Then
						Response.Write("<FONT SIZE=1><A HREF=DBDesignSQLServer.asp?action=deleterelation&amp;relationname=")
						Response.Write server.htmlencode(rsFK.Fields("FK_NAME").Value)
						Response.Write("&amp;sqlfrom=" & Server.URLEncode(Request.Querystring("sqlfrom")) & "><IMG SRC=GRIDLNKDELETE.GIF border=0 ALT=Drop></a></FONT>")
					else
						Response.Write("&nbsp;")
					End If
					Response.Write("</td><td ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("FK_NAME").Value)
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("PK_TABLE_NAME").Value)
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write("Many-to-One")
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("UPDATE_RULE").Value)
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("DELETE_RULE").Value)
					Response.Write("</td><td ALIGN=LEFT NOWRAP VALIGN=TOP>")
					Response.Write(rsFK.Fields("FK_COLUMN_NAME").Value & " &lt;&mdash; " & rsFK.Fields("PK_COLUMN_NAME").Value)
				Elseif UCase(rsFK.Fields("FK_TABLE_NAME").Value) = UCase(replace(replace(request.querystring("sqlfrom"),"]",""),"[","")) and UCase(strIDXName) = Ucase(rsFK.Fields("PK_TABLE_NAME").Value) then
					Response.Write("<BR>")
					Response.Write(rsFK.Fields("FK_COLUMN_NAME").Value & " &lt;&mdash; " & rsFK.Fields("PK_COLUMN_NAME").Value)
				Else
				End If
				rsFK.movenext
				If not rsFK.eof then
					If UCase(strIDXName) = UCase(rsFK.Fields("PK_TABLE_NAME").Value) and UCase(rsFK.FIelds("FK_TABLE_NAME").Value) = UCase(request.querystring("sqlfrom")) Then
					Elseif UCase(rsFK.Fields("FK_TABLE_NAME").Value) = UCase(request.querystring("sqlfrom")) and UCase(strIDXName) <> UCase(rsFK.Fields("PK_TABLE_NAME").Value) Then
					Response.Write("</TD></TR>")
					intRowCount = intRowCount + 1
					Else
					End If
				Else
					Response.Write("</TD></TR>")
					intRowCount = intRowCount + 1
				End If
				End If
			loop
			rsFK.movefirst
'			intRowCount = intRowCount + 1
			do while not rsFK.eof
				If ocdDatabaseType = "SQLServer" Then
				'**********************
								If (rsFK.Fields("PK_TABLE_SCHEMA").Value) = (strSQLObjectOwner) and (rsFK.Fields("PK_TABLE_NAME").Value) = (strSQLObjectName) and UCase(strIDXName) <>Ucase(rsFK.Fields("FK_TABLE_NAME").Value) then
					strIDXName = rsFK.Fields("FK_TABLE_NAME").Value
					Response.Write("<tr ")
					If intRowCount Mod 2 = 0 then
						Response.Write("class=GridEven")
					Else
						Response.Write("class=GridOdd")
					End If
					Response.Write("><td VALIGN=TOP>")
					If (connStructure.Properties("DBMS Name") = "MS Jet" OR connStructure.Properties("DBMS Name") = "ACCESS") and ocdAccessTableEdits Then
						Response.Write("<FONT SIZE=1><A HREF=DBDesignMSAccess.asp?action=deleterelation&relationname=" & Server.URLEncode(rsFK.Fields("FK_NAME").Value) & "&table=" & Server.URLEncode(rsFK.Fields("FK_TABLE_NAME").Value) & "><IMG SRC=GRIDLNKDELETE.GIF border=0 ALT=Drop></a></FONT>")
					Elseif (connStructure.Provider = "SQLOLEDB.1") and ocdSQLTableEdits Then
						Response.Write("<FONT SIZE=1><A HREF=DBDesignSQLServer.asp?action=deleterelation&amp;relationname=")
						Response.Write server.urlencode((rsFK.Fields("FK_NAME").Value))
						Response.Write("&amp;sqlfrom=" & Server.URLEncode(ocdQuotePrefix & rsFK.Fields("FK_TABLE_SCHEMA").Value & ocdQUoteSuffix & "." & ocdQUotePrefix & rsFK.Fields("FK_TABLE_NAME").Value & ocdQuoteSuffix) & "><IMG SRC=GRIDLNKDELETE.GIF border=0 ALT=Drop></a></FONT>")
					else
						Response.Write("&nbsp;")
					End If
					Response.Write("</td><td ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("FK_NAME").Value)
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("FK_TABLE_NAME").Value)
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write("One-To-Many")
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("UPDATE_RULE").Value)
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("DELETE_RULE").Value)
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("PK_COLUMN_NAME").Value & " &mdash;&gt; " & rsFK.Fields("FK_COLUMN_NAME").value)
				Elseif (rsFK.Fields("PK_TABLE_SCHEMA").Value) = (strSQLObjectOwner) and (rsFK.Fields("PK_TABLE_NAME").Value) = (strSQLObjectName) then
					Response.Write("<BR>")
					Response.Write(rsFK.Fields("PK_COLUMN_NAME").Value & " &mdash;&gt; " & rsFK.Fields("FK_COLUMN_NAME").Value) 			
				End If
				rsFK.movenext
				If not rsFK.eof then
					If UCase(strIDXName) = UCase(rsFK.Fields("FK_TABLE_NAME").Value) and (rsFK.Fields("PK_TABLE_SCHEMA").Value) = (strSQLObjectOwner) and (rsFK.Fields("PK_TABLE_NAME").Value) = (strSQLObjectName) Then
					Elseif (rsFK.Fields("PK_TABLE_SCHEMA").Value) = (strSQLObjectOwner) and (rsFK.Fields("PK_TABLE_NAME").Value) = (strSQLObjectName) Then
						Response.Write("</TD></TR>")
						intRowCount = intRowCount + 1
					End If
				Else
					Response.Write("</TD></TR>")
					intRowCount = intRowCount + 1
				End If
				'**************************
				
				Else
				
				If Ucase(rsFK.Fields("PK_TABLE_NAME").Value) = UCASE(replace(replace(request.querystring("sqlfrom"),"]",""),"[","")) and UCase(strIDXName) <> Ucase(rsFK.Fields("FK_TABLE_NAME").Value) then
					strIDXName = rsFK.Fields("FK_TABLE_NAME").Value
					Response.Write("<tr ")
					If intRowCount Mod 2 = 0 then
						Response.Write("class=GridEven")
					Else
						Response.Write("class=GridOdd")
					End If
					Response.Write("><td VALIGN=TOP>")
					If ocdDatabaseType = "Access" and ocdAccessTableEdits Then
						Response.Write("<FONT SIZE=1><A HREF=DBDesignMSAccess.asp?action=deleterelation&relationname=" & Server.URLEncode(rsFK.Fields("FK_NAME").Value) & "&table=" & Server.URLEncode(rsFK.Fields("FK_TABLE_NAME").Value) & "><IMG SRC=GRIDLNKDELETE.GIF border=0 ALT=Drop></a></FONT>")
					Elseif (connStructure.Provider = "SQLOLEDB.1") and ocdSQLTableEdits Then
						Response.Write("<FONT SIZE=1><A HREF=DBDesignSQLServer.asp?action=deleterelation&amp;relationname=")
						Response.Write Server.htmlencode(rsFK.Fields("FK_NAME").Value)
						Response.Write("&amp;sqlfrom=" & Server.URLEncode(rsFK.Fields("FK_TABLE_NAME").Value) & "><IMG SRC=""GRIDLNKDELETE.GIF"" border=0 ALT=Drop></a></FONT>")
					else
						Response.Write("&nbsp;")
					End If
					Response.Write("</td><td ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("FK_NAME").Value)
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("FK_TABLE_NAME").Value)
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write("One-To-Many")
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("UPDATE_RULE").Value)
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("DELETE_RULE").Value)
					Response.Write("</td><td NOWRAP ALIGN=LEFT VALIGN=TOP>")
					Response.Write(rsFK.Fields("PK_COLUMN_NAME").Value & " &mdash;&gt; " & rsFK.Fields("FK_COLUMN_NAME").value)
				Elseif UCase(rsFK.Fields("PK_TABLE_NAME").Value) = UCase(request.querystring("sqlfrom")) then
					Response.Write("<BR>")
					Response.Write(rsFK.Fields("PK_COLUMN_NAME").Value & " &mdash;&gt; " & rsFK.Fields("FK_COLUMN_NAME").Value) 			
				End If
				rsFK.movenext
				If not rsFK.eof then
					If UCase(strIDXName) = UCase(rsFK.Fields("FK_TABLE_NAME").Value) and UCase(rsFK.Fields("PK_TABLE_NAME").Value) = UCase(request.querystring("sqlfrom")) Then
					Elseif UCase(rsFK.Fields("PK_TABLE_NAME").Value) = UCase(request.querystring("sqlfrom")) Then
						Response.Write("</TD></TR>")
						intRowCount = intRowCount + 1
					End If
				Else
					Response.Write("</TD></TR>")
					intRowCount = intRowCount + 1
				End If
				End If
			loop
		End If
		Response.Write("</table>")
		rsFK.close
		set rsFK = nothing
	End If
	If ((connStructure.provider ="Microsoft.Jet.OLEDB.4.0" and ocdAccessTableEdits) OR (connStructure.provider  ="SQLOLEDB.1" and ocdSQLTableEdits)) Then ' OR objNDActiveDataForm.ADOConnection.provider ="SQLOLEDB.1"
		If connStructure.provider ="Microsoft.Jet.OLEDB.4.0" Then
			Response.Write("<FORM ACTION=DBDesignMSAccess.asp METHOD=GET>")
		Else
		If ocdDBMSVersion > CDbl(6.9) Then
			Response.Write("<FORM ACTION=DBDesignSQLServer.asp METHOD=GET>")
		End If
		End If
		
		Response.Write("<IMG SRC=appNew.Gif BORDER=0> Add <SELECT NAME=whatrelation><OPTION VALUE=""Many-To-One"" SELECTED>Many-To-One</OPTION><OPTION VALUE=""One-To-Many"">One-To-Many</OPTION></select> Relation to Table : ")

		for each qs in request.querystring
			If not UCASE(qs) = "ocdACTION" AND not UCASE(qs) = "ocdSUBMITACTION" AND not UCASE(qs) = "NDACTION"  AND not UCASE(qs) = "SQLFROM" Then
				Response.Write("<INPUT TYPE=Hidden NAME=""" & qs & """ VALUE=""" &server.htmlencode(Request.querystring(qs))  & """>")
			End If
		next
		Response.Write "<INPUT TYPE=HIDDEN NAME=SQLFROM VALUE="""
		Response.Write server.htmlencode(replace(replace(request.querystring("sqlfrom"),"]",""),"[",""))
		Response.Write """>"
		Response.Write("<SELECT NAME=""relatedto"" >")

		If ocdDatabaseType = "SQLServer" Then
				set rsNDTemp1 = server.createobject("ADODB.Recordset")
		set rsNDTemp1 = connStructure.OpenSchema(20) 'tables
		arrNDSchemaFields2 (0) = "TABLE_TYPE"
			arrNDSchemaFields2 (1) = "TABLE_NAME"
			arrNDSchemaFields2 (2) = "TABLE_SCHEMA"
			If rsNDTemp1.eof then
					redim arrNDSchema(1,1)
			Else
			arrNDSchema = rsNDTemp1.GetRows(,, arrNDSchemaFields2 )
			End If
		
		for intK = 0 to UBound( arrNDSchema ,2)
			If arrNDSchema (0, intK ) = "TABLE" and UCASE(left(arrNDSchema(1,intK),4)) <> "MSYS" then
				If UCase(arrNDSchema(1,intK)) = UCase(request.QueryString("SQLFrom")) Then
				Else
					Response.Write("<OPTION")
					Response.Write(" VALUE=""" & server.htmlencode(ocdQuotePrefix & arrNDSchema(2,intK) & ocdQuoteSuffix & "." & ocdQuotePrefix & arrNDSchema(1,intK) & ocdQuoteSuffix ) & """>" & server.htmlencode(ocdQuotePrefix & arrNDSchema(2,intK) & ocdQuoteSuffix & "." & ocdQuotePrefix & arrNDSchema(1,intK) & ocdQuoteSuffix )  & "</OPTION>")
				End If
			End If
		next

		Else
		set rsNDTemp1 = server.createobject("ADODB.Recordset")
		set rsNDTemp1 = connStructure.OpenSchema(20) 'tables
		arrNDSchemaFields (0) = "TABLE_TYPE"
			arrNDSchemaFields (1) = "TABLE_NAME"
			If rsNDTemp1.eof then
					redim arrNDSchema(1,1)
			Else
			arrNDSchema = rsNDTemp1.GetRows(,, arrNDSchemaFields )
			End If
		
		for intK = 0 to UBound( arrNDSchema ,2)
			If arrNDSchema (0, intK ) = "TABLE" and UCASE(left(arrNDSchema(1,intK),4)) <> "MSYS" then
				If UCase(arrNDSchema(1,intK)) = UCase(request.QueryString("SQLFrom")) Then
				Else
					Response.Write("<OPTION")
					Response.Write(" VALUE=""" & server.htmlencode(arrNDSchema(1,intK)) & """>" & arrNDSchema(1,intK) & "</OPTION>")
				End If
			End If
		next
		End If
		Response.Write("</select>")
		Response.Write("<INPUT NAME=ocdaction TYPE=HIDDEN Value=""AddRelation"">")
		Response.Write("<INPUT CLASS=Submit NAME=ocdsubmitaction TYPE=Submit Value=""Create Relation"">")
		Response.Write("</form>")
	End If
End If
Response.Write "</P>"
call writefooter("")
Private function NoSQLIdentifier (strSQLTableString) 
	dim tmpstrSQLTableString
	tmpstrSQLTableString = strSQLTableString
	Select Case ocdDatabaseType
	Case "Oracle"
'	If isOracle THen
		tmpstrSQLTableString = mid(strSQLTableString,instr(strSQLTableString,".") + 2)
		tmpstrSQLTableString = left(tmpstrSQLTableString, len(tmpstrSQLTableString)-1)
		NoSQLIdentifier = tmpstrSQLTableString
'	Else
'	If DatabaseType = "Access" Then
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
'		Response.Write tmpstrSQLTableString
'	Else
'		NoSQLIdentifier = tmpstrSQLTableString
'	End If
'	End If
	End Select
end function

Sub CleanUpObjects()
	On error resume next
	Err.clear
	Response.end
end sub
%>
