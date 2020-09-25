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
'On Error GoTo 0 'For debugging only
Dim strSQLFrom, strFieldName, rsSchema
strSQLFrom = CStr(FormatForSQL(Request.QueryString("SQLFROM"), ocddatabasetype, "RemoveSQLIdentifier"))
strFieldName = Request.QueryString("sqlfield")
Response.Write("<html><head><title>Describe Field</title>")
Response.Write("<link rel=stylesheet type=""text/css"" href=""" & ocdStyleSheet & """>")
Response.Write("</head><body onload=""javascript:self.focus();""><p><span class=""information"">")
Response.Write(Server.HTMLEncode(strSQLFrom) & " : " & Server.HTMLEncode(strFieldName))
Response.Write("</span></p>")
set rsSchema = (ocdTargetConn.OpenSchema(4, Array(empty,empty,strSQLFrom,strFieldName)))
'response.write rsSchema.recordcount
Do While Not rsSchema.eof
			Response.Write("<p><span class=""fieldname"">Type :</span> ")
			Select Case rsSchema("DATA_TYPE")
				Case 0 'adEmpty
					Response.Write("Empty")
				Case 16 'adTinyInt 
					Response.Write("TinyInt")
				Case 2 'adSmallInt
					Response.Write("SmallInt")
				Case 3 'adInteger 
					Response.Write("Integer")
				Case 20 'adBigInt 
					Response.Write("BigInt")
				Case 17 'adUnsignedTinyInt 
					Response.Write("UnsignedTinyInt")
				Case 18 'adUnsignedSmallInt
					Response.Write("UnsignedSmallInt")
				Case 19 'adUnsignedInt
					Response.Write("UnsignedInt")
				Case 21 'adUnsignedBigInt
					Response.Write("UnsignedBigInt")
				Case 4 'adSingle
					Response.Write("Single")
				Case 5 'adDouble
					Response.Write("Double")
				Case 6 'adCurrency
					Response.Write("Currency")
				Case 14 'adDecimal
					Response.Write("Decimal")
				Case 131 'adNumeric
					Response.Write("Numeric")
				Case 11 'adBoolean
					Response.Write("Boolean")
				Case 10 'adError
					Response.Write("Error")
				Case 132 'adUserDefined
					Response.Write("UserDefined")
				Case 12 'adVariant
					Response.Write("Variant")
				Case 9 'adIDispatch
					Response.Write("IDispatch")
				Case 13 'adIUnknown
					Response.Write("IUnknown")
				Case 72 'adGUID
					Response.Write("GUID")
				Case 7 'adDate
					Response.Write("Date")
				Case 133 'adDBDate
					Response.Write("DBDate")
				Case 134 'adDBTime
					Response.Write("DBTime")
				Case 135 'adDBTimeStamp
					Response.Write("DBTimeStamp")
				Case 8 'adBSTR
					Response.Write("BSTR")
				Case 129 'adChar
					Response.Write("Char")
				Case 200 'adVarChar
					Response.Write("VarChar")
				Case 201 'adLongVarChar
					Response.Write("LongVarChar")
				Case 130 'adWChar
					Response.Write("WChar")
				Case 202 'adVarWChar
					Response.Write("VarWChar")
				Case 203 'adLongVarWChar
					Response.Write("LongVarWChar")
				Case 128 'adBinary 
					Response.Write("Binary")
				Case 204 'adVarBinary
					Response.Write("VarBinary")
				Case 205 'adLongVarBinary
					Response.Write("LongVarBinary")
				Case Else
					Response.Write "Type" & fldTemp.Type
			End Select
			Response.Write("</p><p>")
			If IsNull(rsSchema("DESCRIPTION")) Then
				Response.Write("<span class=""warning"">No Description Available</span>")
			Else
				Response.Write("<span class=""fieldname"">Description :</span> ")
				Response.Write(Server.HTMLEncode(rsSchema("DESCRIPTION")))
			End If
	exit do
Loop
rsSchema.Close
Set rsSchema = Nothing
ocdTargetConn.Close
set ocdTargetConn = Nothing
Response.Write("</body></html>")
%>
