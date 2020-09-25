<%
'1 Click DB Pro copyright 1997-2004 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com/

        
'**Start Encode**

Class ocdManageSQLServer
	Public ADOConnection	'ADO connection currently in use for the object
	Public SQLConnect	'ADO connection string
	Public SQLUser	'ADO connection user name
	Public SQLPass	'ADO connection password
	Public SQLObject
	Public SQLObjectType
	Public QuoteSuffix
	Public QuotePrefix
	Public ErrorDescription
	Public ErrorNumber
	Private Sub Class_Initialize 'Set default values
		QuoteSuffix = """"
		QuotePrefix = """"
	end sub
Public Function GetSQLObjectOwner()
	GetSQLObjectOwner = GetSQLIDFPart(SQLObject,"SQLOBJECTOWNER",QuoteSuffix,QuotePrefix)

End Function

Public Function GetSQLObjectName()
GetSQLObjectName = GetSQLIDFPart(SQLObject,"SQLOBJECTNAME",QuoteSuffix,QuotePrefix)

end function

Private Function GetSQLIDFPart(strSQLIdf,strIdfPart, strPrefix,strSuffix)
	dim arrPieces, strTemp
	arrPieces = split(strSQLIdf,".")
	strTemp = ""
	select case UCASE(strIDfPart)
		Case "SQLOBJECTNAME"
			strTemp = arrPieces(UBound(arrPieces))
		Case "SQLOBJECTOWNER"
			if Ubound(arrPieces) > 0 Then
				strTemp = arrPieces(UBound(arrPieces)-1)
			end if
		Case Else
	End select
	strTemp = replace(strTemp,strPrefix,"")
	strTemp = replace(strTemp,strSuffix,"")	
	GetSQLIDFPart = strTemp
End Function

Public Sub Open()
	If not IsObject(ADOConnection) Then
		Set ADOConnection  = server.CreateObject("ADODB.Connection")
		ADOConnection.Mode = 1 'adModeRead
		Call ADOConnection.Open (SQLConnect, SQLUser, SQLPass)
	End If
End Sub	

Public Sub CopyObject(strCopyType, strNewObjectName)
	if strCopyType = "Data" Then
		 ADOConnection.execute "SELECT * INTO " & strNewObjectName & " FROM " & SQLObject
	Else
		 ADOConnection.execute "SELECT * INTO " & strNewObjectName & " FROM " & SQLObject & " WHERE 1 = 2"
	End if
end sub

Public Sub DropField(strFieldName)
		Dim dcrs, dcTableID, dcColID, dcName
		set dcrs = ADOConnection.execute ("Select sysobjects.id from sysobjects inner join sysusers on sysobjects.uid = sysusers.uid where sysobjects.name ='" & replace(GetSQLObjectName(),"'","''") & "' AND sysusers.name ='" & replace(GetSQLObjectOwner(),"'","''") & "'")
			if not dcrs.eof Then
				dcTableID = dcrs("ID")
				dcrs.close
				set dcrs = ADOConnection.Execute ("SELECT ""CDEFAULT"" from syscolumns where ""name""='" & replace(strFieldName,"'","''") & "'" & " and ""id"" = " & dcTableID )
				if not dcrs.eof then
					dcColID = dcrs("CDEFAULT")
					dcrs.close
					set dcrs = ADOConnection.Execute ("SELECT ""NAME"" from sysobjects where  ""id"" = " & dcColID )
					if not dcrs.eof then
						dcName = dcrs("NAME")
						dcrs.close
						response.write "ALTER TABLE " & SQLObject & " DROP Constraint " & QuotePrefix & dcName & QuoteSuffix & ""
						ADOConnection.Execute "ALTER TABLE " & SQLObject & " DROP Constraint " & QuotePrefix & dcName & QuoteSuffix & ""
					End if
				End if
			End if
			ADOConnection.Execute ("ALTER TABLE " & SQLObject & " DROP COLUMN " & quoteSuffix & strFieldName & QuotePrefix & "")
end Sub

sub DropIndex(strIndexName)
		on error resume next
		ADOConnection.Execute "DROP INDEX " & SQLObject & "." & QuoteSuffix &  strIndexName & QuotePrefix
		if err = -2147217900 Then 'maybe  primary key
			err.Clear
			ADOConnection.Execute ("ALTER TABLE " & SQLObject & " DROP Constraint " & QuoteSuffix &  strIndexName & QuotePrefix)
		End if	
end sub

Function GetHelpText(strObject)
	dim rstrtxt, vartrtxt
	vartrtxt = ""
	set rstrtxt = server.createobject("ADODB.Recordset")
	set rstrtxt = ADOConnection.execute ("sp_helptext '" & strObject & "'")
	do while not rstrtxt.eof
		vartrtxt = vartrtxt & (rstrtxt.Fields(0).Value & "")
		rstrtxt.movenext
	loop
	GetHelpText = vartrtxt
End function

End Class
%>