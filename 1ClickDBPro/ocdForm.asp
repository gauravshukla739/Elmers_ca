<%
'1 Click DB ASP Library - SQL to ASP Form Object
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

Class ocdForm
	Public ADOConnection	'ADO connection currently in use for the object
	Public ADORecordset	'ADO recordset containing either the record being edited or the empty structure of the recordset being added to
	Public SQLConnect	'ADO connection string
	Public SQLUser	'ADO connection user name
	Public SQLPass	'ADO connection password
	Public SQLSelect 'Comma delimited list of fields to be displayed for edit
	Public SQLFrom	'Name of table to be edited, using an SQL Join is possible but may sometimes produce unpredictable results, test first
	Public SQLWhereExtra 'Use to set extra security restrictions 
	Public SQLWhere 'When no autoincrement field is present, used to identify record
	Public AllowEdit 'if true, enable editing existing records
	Public AllowAdd	'if true, enable adding new records
	Public AllowView	'if true, enable viewing of records
	Public AllowDelete	'if true, enable deletion of records
	Public AllowMultiDelete
	Public MaxRelatedValues
	Public DatabaseType	'Determined automatically in code, should generally be used as a read only field
	Public HTMLCheckField	'HTML string containing the text displayed by a field that needs checking, default is red letter saying "Check Field"
	Public HTMLAttribSaveBtn 'TAG Attributes, include INPUT TYPE and VALUE for Save Button
	Public HTMLAttribCancelBtn 'TAG Attributes, include INPUT TYPE and VALUE for Cancel Button
	Public HTMLAttribNewBtn 'TAG Attributes, include INPUT TYPE and VALUE for New Button
	Public HTMLAttribDeleteBtn 'TAG Attributes, include INPUT TYPE and VALUE for Delete Button
	Public HTMLStatusFieldError 'HTML displayed in error message when field cannot be updated 
	Public HTMLNotFound 'HTML displayed in error message when record cannot be found
	Public HTMLOLEError 'HTML in place of ADO message "Multiple-step operation generated errors" for generic -2147217887 msg
	Public HTMLUpdateFieldFails
	Public HTMLAutoIncrement 'HTML displayed for autoincrement field value when adding a new record
	Public HTMLStatusStart	'HTML placed before error message display (if any)
	Public HTMLStatusEnd 'HTML placed before error message display (if any)
	Public FormStatus	'a string containing any form error messages
	Public HTMLAttribForm
	Public EnablePrimaryKeyLink
	Public FormMode
	Public EditMode
	Public SQLID
	Public FormEStringToken
	Public FormNullToken
	Public InvalidFields
	Public CallBeforeDelete
	Public CallAfterDelete
	Public CallBeforeUpdate
	Public CallAfterUpdate
	Public CallBeforeInsert
	Public CallAfterInsert
	Public CallOnCancel
	Public CallPreDelete
	Public MayBeReplica
	Public CallValidateBeforeUpdate
	Public CallValidateBeforeInsert
	'**Start Encode**
	'internal reserved
	Public QuoteSuffix
	Public QUotePrefix
	Private SQLSelectPK
	Public SQLSelectID
	Private FormEditConnect
	Private ADOMode 'can be used to set connection lock types 
	Public Debug

	Private Sub Class_Initialize
	
%><!--#INCLUDE FILE=ocdForm_Lang.asp--><%
		MaxRelatedValues = 1000
		MayBeReplica = True
		Debug = False
		FormNullToken = ""
		FormEditConnect = True
		SQLID = ""
		FormEStringToken = """"""
		AllowMultiDelete = False
		DatabaseType = ""
		SQLSelectID = ""
		SQLSelectPK = ""
		SQLWhereExtra = ""
		SQLWHERE = ""
		CallValidateBeforeUpdate = False
		CallValidateBeforeInsert = False
		EnablePrimaryKeyLink = True
		CallBeforeDelete = True
		CallAfterDelete = True
		CallBeforeUpdate = True
		CallAfterUpdate = True
		CallBeforeInsert = True
		CallAfterInsert = True
		CallOnCancel = False
		CallPreDelete = False
		QuoteSuffix = """"
		QuotePrefix = """"
		ADOMode = 3	   'adOpenReadWrite 0 'adOpenUnknown
	End Sub

	Public Sub DisplayFieldAsTextBox(ByVal strFieldName, ByVal strDefault, ByVal strAttributes)
		On Error Resume Next
		Dim tcf		  '- temp for CreateField return value
		tcf = ""
		Select Case EditMode
			Case "Add", "Edit"
				If strFIeldName = "s_Generation" And MayBeReplica Then
					If EditMode = "Edit" Then
						tcf = ADORecordset.Fields(strFieldName).Value
					End If
				ElseIf left(strFIeldName, 4) = "Gen_" And MayBeReplica Then
					If EditMode = "Edit" Then
						tcf = ADORecordset.Fields(strFieldName).Value
					End If
				ElseIf ADORecordset.FIelds(strFieldName).Type = 72 Then
					If EditMode = "Edit" Then
						tcf = ADORecordset.Fields(strFieldName).Value
					End If
				ElseIf UCase(strFieldName) = UCASE(SQLSelectID) Then
					If EditMode = "Edit" Then
						tcf = ADORecordset.Fields(strFieldName).Value & vbCRLF
					Else
						tcf = HTMLAutoIncrement
					End If
				Else
					tcf = tcf & "<input "
					tcf = tcf & "type=""text"" "
					tcf = tcf & "name=""ocdTF" & Server.HTMLEncode(strFieldName) & """ "
					tcf = tcf & "value="""
					If FormMode = "Save" Then
						tcf = tcf & Server.HTMLEncode(Request.Form("ocdTF" & strFieldName))
					Else
						If EditMode = "Add" Then
							If Request.Form("ocdTF" & strFieldName) <> "" Then
tcf = tcf & Server.HTMLEncode(Request.Form("ocdTF" & strFieldName))
Else
							tcf = tcf & Server.HTMLEncode(strDefault)
							End If
						Else
							If Not IsNull(ADORecordset.Fields(strFieldName).Value) Then
								Select Case ADORecordset.Fields(strFieldName).Type
									Case 6
									If ADORecordset.Fields(strFieldName).Value = CCur(FormatCurrency(ADORecordset.Fields(strFieldName).Value)) Then
											tcf = tcf & Server.HTMLEncode(FormatCurrency(ADORecordset.Fields(strFieldName).Value))
										Else
											tcf = tcf & Server.HTMLEncode(ADORecordset.Fields(strFieldName).Value)

										End If
									Case 202, 203, 200, 201
										If ADORecordset.Fields(strFieldName).Value = "" Then
											tcf = tcf & Server.HTMLEncode(FormEStringToken)
										Else
											tcf = tcf & Server.HTMLEncode(ADORecordset.Fields(strFieldName).Value)
										End If
									Case Else
										tcf = tcf & Server.HTMLEncode(ADORecordset.Fields(strFieldName).Value)
								End Select
							End If
						End If
					End If
					tcf = tcf & """ "
					tcf = tcf & strAttributes
					tcf = tcf & ">"
					If FormMode = "Save" Then
						If IsFieldDataInvalid(strFieldName) Then
							tCF = tCF & HTMLCheckField
						End If
					End If
				End If
			Case "ReadOnly"
				If Not IsNull(ADORecordset.Fields(strFieldName).Value) Then
					If ADORecordset.Fields(strFieldName).Type = 6 Then
						tcf = tcf & Server.HTMLEncode(FormatCurrency(ADORecordset.Fields(strFieldName).Value))
					Else
						tcf = tcf & Server.HTMLEncode(ADORecordset.Fields(strFieldName).Value)
					End If
				End If
			Case Else
				tCF = ""
		End Select
		Response.Write(tcf)
	End Sub

	Public Sub DisplayFieldAsCheckBox(ByVal strFieldName, ByVal varCheckValue, ByVal varUncheckValue, ByVal varDefault, ByVal strAttributes)
		Dim tcf
		tcf = ""
		Select Case EditMode
			Case "Add", "Edit"
				tcf = "<input type=""Checkbox"" name=""ocdCB" & Server.HTMLEncode(strFieldName) & """ "
				If FormMode = "Save" Then
					If Request.Form("ocdCB" & strFieldName) <> "" Then
						tcf = tcf & "checked"
					End If
				ElseIf EditMode = "Edit" Then
					If Not IsNull(ADORecordset.Fields(strFieldName).Value) Then
						If CStr(ADORecordset.Fields(StrFieldName).Value) = CStr(varCheckValue & "") Then
							tcf = tcf & "checked"
						End If
					End If
				Else
					If CStr(varDefault & "") = CStr(varCheckValue & "") Then
						tcf = tcf & "checked"
					End If
				End If
				tcf = tcf & " " & strAttributes & " >"
				tcf = tcf & "<input type=""Hidden"" name=""ocdCT" & Server.HTMLEncode(strFieldName) & """ value=""" & Server.HTMLEncode(CStr(varCheckValue & ""))
				tcf = tcf & """ " & strAttributes & " >"
				tcf = tcf & "<input type=""Hidden"" name=""ndchf" & Server.HTMLEncode(strFieldName) & """ value=""" & Server.HTMLEncode(CStr(varUnCheckValue & "")) & """>"
				If FormMode = "Save" Then
					If IsFieldDataInvalid(strFieldName) Then
						tCF = tCF & HTMLCheckField
					End If
				End If
			Case "ReadOnly"
				If Not IsNull(ADORecordset.Fields(strFieldName).Value) Then
					tcf = tcf & Server.HTMLEncode(ADORecordset.Fields(strFieldName).Value)
				End If
			Case Else
				tCF = ""
		End Select
		Response.Write(tcf)
	End Sub

	Public Sub DisplayFieldAsMemo(ByVal strFieldName, ByVal strDefault, ByVal strAttributes)
		Dim tcf		  '- temp for CreateField return value
		Dim tvarTemp
		Select Case EditMode
			Case "Add", "Edit"
				tcf = tcf & "<textarea "
				tcf = tcf & "name=""ocdTF" & Server.HTMLEncode(strFieldName) & """ "
				tcf = tcf & strAttributes & " "
				tcf = tcf & ">"
				If FormMode = "Save" Then
					tcf = tcf & Server.HTMLEncode(Request.Form("ocdTF" & strFieldName))
				Else
					If EditMode = "Add" Then
						If Request.Form("ocdTF" & strFieldName) = "" Then
							tcf = tcf & Server.HTMLEncode(strDefault)
						Else
							tcf = tcf & Server.HTMLEncode(Request.Form("ocdTF" & strFieldName))
						End If
					Else
						tvarTemp = ADORecordset.Fields(strFieldName).Value
						If Not IsNull(tvarTemp) Then
							If tvarTemp = "" Then
								tcf = tcf & Server.HTMLEncode(FormEStringToken)
							Else
								tcf = tcf & Server.HTMLEncode(tvarTemp)
							End If
						End If
					End If
				End If
				tcf = tcf & "</textarea>"
				If FormMode = "Save" Then
					If IsFieldDataInvalid(strFieldName) Then
						tCF = tCF & HTMLCheckField
					End If
				End If
			Case "ReadOnly"
				tvarTemp = ADORecordset.Fields(strFieldName).Value
				If Not IsNull(tvarTemp) Then
					tcf = tcf & Server.HTMLEncode(tvarTemp)
				End If
			Case Else
				tCF = ""
		End Select
		Response.Write(tcf)
	End Sub

	Public Sub DisplayFieldAsSelectList(ByVal strFieldName, ByVal strBoundValues, ByVal strDisplayValues, ByVal strDefaultValue, ByVal strAttributes)
		Dim tCDDF		  '- variable to hold return value of function
		Dim tarrBound, tarrDisplay, tintI, blnSL
		blnSL = False
		tcDDF = ""
		If AllowEdit Then
			tarrBound = split(strBoundValues, ";")
			tarrDisplay = split(strDisplayValues, ";")
			tCDDF = tCDDF & "<select "
			tCDDF = tCDDF & "name=""ocdTF" & Server.HTMLEncode(strFieldName) & """ "
			tCDDF = tCDDF & strAttributes & ">"
			For tintI = 0 To UBound(tarrBound)
				tCDDF = tCDDF & "<option value="""
				If tarrBound(tintI) <> "" Then
					tCDDF = tCDDF & Server.HTMLEncode(tarrBound(tintI))
				End If
				tCDDF = tCDDF & """ "
				If Request.Form("ocdTF" & strFieldName) <> "" Then
					If CStr(tarrBound(tintI)) = CStr(Request.Form("ocdTF" & strFieldName)) Then
						tCDDF = tCDDF & "selected "
						blnSL = True
					End If
				Else
					If (SQLID <> "" Or SQLWHERE <> "") And FormStatus = "" And InvalidFields = "" Then
						If Not ADORecordset.eof Then
							Select Case ADORecordset.Fields(strFieldName).Type
								Case 11								  'adBoolean
									If tarrBound(tintI) = "" Then
										If IsNull(ADORecordset.Fields(strFieldName).Value) Then
											tCDDF = tCDDF & "selected "
											blnSL = True
										End If
									ElseIf CInt(tarrBound(tintI)) = CInt(ADORecordset.Fields(strFieldName).Value) Then
										tCDDF = tCDDF & "selected "
										blnSL = True
									End If
								Case Else
									If CStr(tarrBound(tintI)) = CStr(ADORecordset.Fields(strFieldName).Value & "") Then
										tCDDF = tCDDF & "selected"
										blnSL = True
									End If
									If tarrBound(tintI) = "" And IsNull(ADORecordset.Fields(strFieldName).Value) Then
										tCDDF = tCDDF & "selected"
										blnSL = True
									End If
							End Select
						Else
							If Request.Form("ocdTF" & strFieldName) = "" Then
								If tarrBound(tintI) = strDefaultValue Then
									tCDDF = tCDDF & "selected"
									blnSL = True
								End If
							Else
								If Request.Form("ocdTF" & strFieldName) = strDefaultValue Then
									tCDDF = tCDDF & "selected"
									blnSL = True
								End If
							End If
						End If
					ElseIf (SQLID = "" And SQLWHERE = "") And FormStatus = "" And InvalidFields = "" Then
						If Request.Form("ocdTF" & strFieldName) = "" Then
							If tarrBound(tintI) = strDefaultValue Then
								tCDDF = tCDDF & "selected"
								blnSL = True
							End If
						Else
							If Request.Form("ocdTF" & strFieldName) = strDefaultValue Then
								tCDDF = tCDDF & "selected"
								blnSL = True
							End If
						End If
					End If
				End If
				tCDDF = tCDDF & ">"
				If strDisplayValues = "" Then
					If tarrBound(tintI) <> "" Then
						tCDDF = tCDDF & Server.HTMLEncode(tarrBound(tintI))
					End If
				Else
					If tarrDisplay(tintI) <> "" Then
						tCDDF = tCDDF & Server.HTMLEncode(tarrDisplay(tintI))
					End If
				End If
				Response.Write(tcddF)
				tcddf = ""
			Next
			If Not blnSL Then
				tCDDF = tCDDF & "<option value="""
				If Request.Form("ocdTF" & strFieldName) <> "" Then
					tCDDF = tCDDF & Server.HTMLEncode(Request.Form("ocdTF" & strFieldName))
				ElseIf (SQLID <> "" Or SQLWHERE <> "") And FormStatus = "" And InvalidFields = "" Then
					If Not IsNull(ADORecordset.Fields(strFieldName).Value) Then
						tCDDF = tCDDF & Server.HTMLEncode(ADORecordset.Fields(strFieldName).Value)
					End If
				End If
				tCDDF = tCDDF & """ SELECTED>"
				If (SQLID <> "" Or SQLWHERE <> "") And FormStatus = "" And InvalidFields = "" Then
					If Request.Form("ocdTF" & strFieldName) <> "" Then
						tCDDF = tCDDF & Server.HTMLEncode(Request.Form("ocdTF" & strFieldName))
					ElseIf (SQLID <> "" Or SQLWHERE <> "") And FormStatus = "" And InvalidFields = "" Then
						If Not IsNull(ADORecordset.Fields(strFieldName).Value) Then
							tCDDF = tCDDF & Server.HTMLEncode(ADORecordset.Fields(strFieldName).Value)
						End If
					End If
				Else
					tCDDF = tCDDF & Server.HTMLEncode(Request.Form("ocdTF" & strFieldName))
				End If
			End If
			tCDDF = tCDDF & "</select>"
			If IsFieldDataInvalid(strFieldName) Then
				tCDDF = tCDDF & HTMLCheckField
			End If
		Else
			If Not ADORecordset.eof Then
				tCDDF = ADORecordset.Fields(strFieldName).Value & vbCRLF
			End If
		End If
		Response.Write(tCDDF)
	End Sub

	Public Sub DisplayFieldAsButtons(ByVal strFieldName, ByVal strBoundValues, ByVal strDisplayValues, ByVal strDefaultValue, ByVal strOrientation)
		Dim tCDDF		  '- variable to hold return value of function
		Dim tarrBound, tarrDisplay, tintI
		tcDDF = ""
		tarrBound = split(strBoundValues, ";")
		tarrDisplay = split(strDisplayValues, ";")
		If UCASE(strOrientation) = "VERTICAL" Then
			tCDDF = tCDDF & "<table vorder=0>"
		End If
		For tintI = 0 To UBound(tarrBound)
			If UCASE(strOrientation) = "VERTICAL" Then
				tCDDF = tCDDF & "<tr><td>"
			End If
			tCDDF = tCDDF & "<input type=""radio"" name=""ocdTF" & Server.HTMLEncode(strFieldName) & """ "
			tCDDF = tCDDF & "value="""
			If tarrBound(tintI) <> "" Then
				tCDDF = tCDDF & Server.HTMLEncode(tarrBound(tintI))
			End If
			tCDDF = tCDDF & """ "
			If Request.Form("ocdTF" & strFieldName) <> "" Then
				If tarrBound(tintI) = Request.Form("ocdTF" & strFieldName) Then
					tCDDF = tCDDF & "CHECKED "
				End If
			Else
				If (SQLID <> "" Or SQLWHERE <> "") And FormStatus = "" And InvalidFields = "" Then
					If Not ADORecordset.eof Then
						If tarrBound(tintI) = "" And IsNull(ADORecordset.Fields(strFieldName).Value) Then
							tCDDF = tCDDF & "checked"
						ElseIf CStr(tarrBound(tintI)) = CStr(ADORecordset.Fields(strFieldName).Value) Then
							tCDDF = tCDDF & "checked"
						End If
					Else
						If tarrBound(tintI) = strDefaultValue Then
							tCDDF = tCDDF & "checked"
						End If
					End If
				ElseIf (SQLID = "" And SQLWHERE = "") And FormStatus = "" And InvalidFields = "" Then
					If tarrBound(tintI) = strDefaultValue Then
						tCDDF = tCDDF & "checked"
					End If
				End If
			End If
			tCDDF = tCDDF & ">"
			If strDisplayValues = "" Then
				If tarrBound(tintI) <> "" Then
					tCDDF = tCDDF & Server.HTMLEncode(tarrBound(tintI))
				End If
			Else
				If tarrDisplay(tintI) <> "" Then
					tCDDF = tCDDF & Server.HTMLEncode(tarrDisplay(tintI))
				End If
			End If
			If UCASE(strOrientation) = "VERTICAL" Then
				tCDDF = tCDDF & "</td></tr>"
			End If
		Next
		If UCASE(strOrientation) = "VERTICAL" Then
			tCDDF = tCDDF & "</table>"
		End If
		If IsFieldDataInvalid(strFieldName) Then
			tCDDF = tCDDF & HTMLCheckField
		End If
		Response.Write(tCDDF)
	End Sub

	Public Sub DisplayFieldAsRelatedValues(ByVal strFieldName, ByVal strConstraintSQL, ByVal strDefaultValue, ByVal strAttributes)
		
		Dim tCDDF, tvarvalue, tintI, trsConstraint, tarrGetRows, trsConstraintFieldCount
		If Not IsObject(ADOConnection) Then
			exit sub
		Else
			If ADOConnection.State <> &H00000001 Then
				exit sub
			End If
		End If
		Set trsConstraint = Server.CreateObject("ADODB.Recordset")
		If Debug Then
			Response.Write(strConstraintSQL)
		End If
		Call trsConstraint.open(strConstraintSQL, ADOConnection)
		dim inttrsCount
		If Not trsConstraint.eof Then
			tarrGetRows = trsConstraint.getrows
			trsConstraintFieldCount = trsConstraint.Fields.Count
			inttrsCount = UBound(tarrGetRows,2)
		Else
			trsConstraintFieldCount = 0
						inttrsCount = - 1
		End If
		trsConstraint.close()
		Set trsConstraint = Nothing
		If inttrsCount > MaxRelatedValues or inttrsCount = -1 Then
			Call DisplayFieldAsTextBox(strFieldName, strDefaultValue, strAttributes)
		Else
		tcDDF = ""
		Select Case EditMode
			Case "Add", "Edit"
				tCDDF = tCDDF & "<select "
				tCDDF = tCDDF & "name=""ocdTF" & Server.HTMLEncode(strFieldName) & """ "
				tCDDF = tCDDF & strAttributes & ">"
				If trsConstraintFieldCount <> 0 Then
					tCDDF = tCDDF & "<option "
					tCDDF = tCDDF & "value="""" "
					If FormMode = "Save" Then
						If Request.Form("ocdTF" & strFieldName) = "" Then
							tCDDF = tCDDF & "selected"
						End If
					ElseIf EditMode = "Add" Then
						If strDefaultValue = "" Then
							tCDDF = tCDDF & "selected"
						End If
					Else
						tvarvalue = ADORecordset.Fields(strFieldName).Value
						If IsNull(tvarvalue) Then
							tCDDF = tCDDF & "selected"
						End If
					End If
					tCDDF = tCDDF & "></option>"
					For tintI = 0 To UBound(tarrGetRows, 2)
						tCDDF = tCDDF & "<option "
						tCDDF = tCDDF & "value="""
						If Not IsNull(tarrGetRows(0, tIntI)) Then													
						If tarrGetRows(0, tIntI) = "" Then
								tCDDF = tCDDF & Server.HTMLEncode(FormEStringToken)
							Else
								tCDDF = tCDDF & Server.HTMLEncode(tarrGetRows(0, tIntI))
							End If
						End If
						tCDDF = tCDDF & """ "
						If FormMode = "Save" Then
							If isnull(tarrGetRows(0, tIntI)) Then
								If Request.Form("ocdTF" & strFieldName) = "" Then
									tCDDF = tCDDF & "selected "
								End If
							Else
								If tarrGetRows(0, tIntI) = "" And Request.Form("ocdTF" & strFieldName) = FormEStringToken Then
									tCDDF = tCDDF & "SELECTED "
								ElseIf CStr(tarrGetRows(0, tIntI)) = Request.Form("ocdTF" & strFieldName) Then
									tCDDF = tCDDF & "SELECTED "
								End If
							End If
						ElseIf EditMode = "Add" Then
							If IsNull(tarrGetRows(0, tIntI)) Then
								
									If strDefaultValue = "" Then
										tCDDF = tCDDF & "SELECTED "
									End If

							Else
							If Request.Form("ocdTF" & strFieldName) = "" Then
								If CStr(tarrGetRows(0, tIntI)) = strDefaultValue Then
									tCDDF = tCDDF & "SELECTED "
								End If
							Else
								If CStr(tarrGetRows(0, tIntI)) = Request.Form("ocdTF" & strFieldName) Then
									tCDDF = tCDDF & "SELECTED "
								End If
							
							End If
							End If
						ElseIf EditMode = "Edit" Then
							If IsNull(tarrGetRows(0, tIntI)) And IsNull(tvarvalue) Then
								tCDDF = tCDDF & "SELECTED "
							ElseIf IsNull(tarrGetRows(0, tIntI)) Or IsNull(tvarvalue) Then
								'no match
							Else
								If CStr(tarrGetRows(0, tIntI)) = CStr(tvarvalue) Then
									tCDDF = tCDDF & "SELECTED "
								End If
							End If
						End If
						tCDDF = tCDDF & ">"
						If trsConstraintFieldCount > 1 Then
							If Not IsNull(tarrGetRows(1, tIntI)) Then
								If CStr(tarrGetRows(1, tIntI)) = FormEstringToken Then
									tCDDF = tCDDF & Server.HTMLEncode(FormEstringToken)
								Else
									tCDDF = tCDDF & Server.HTMLEncode(tarrGetRows(1, tIntI))
								End If
							End If
						Else
							If Not IsNull(tarrGetRows(0, tIntI)) Then
								tCDDF = tCDDF & Server.HTMLEncode(tarrGetRows(0, tIntI))
							End If
						End If
						tCDDF = tCDDF & "</option>"
					Next
				End If
				tCDDF = tCDDF & "</select>"
				If FormMode = "Save" Then
					If IsFieldDataInvalid(strFieldName) Then
						tCDDF = tCDDF & HTMLCheckField
					End If
				End If
			Case "ReadOnly"
				If Not IsNull(ADORecordset.Fields(strFieldName).Value) Then
					tCDDF = tCDDF & Server.HTMLEncode(ADORecordset.Fields(strFieldName).Value)
				End If
			Case Else
				tCDDF = ""
		End Select
		Response.Write(tCDDF)
		End If
	End Sub
	Private Function IsFieldDataInvalid(ByVal strpFieldName)
		Dim isdiintI, isdiarrT, tmpInvalidFields
		IsFieldDataInvalid = False
		isdiarrT = split(InvalidFields, ";")
		For isdiintI = 0 To ubound(isdiarrT)
			If strpFieldName = isdiarrT(isdiintI) Then
				IsFieldDataInvalid = True
				Exit For
			End If
		Next
	End Function

	Public Sub Open()
	If Not Debug Then
		On Error Resume Next
	End If
		Dim fldTemp, fmRequest, strSQL, strURL, QS, strSQLIDX, rsIDX, strSQLWherePK, strCDp, arrSQLSelectPK, intTemp
		If Not EnablePrimaryKeyLink Then
			SQLWhere = ""
		Else
			If SQLWhere = "" Then
				SQLWhere = Request.QueryString("SQLWHERE")
			End If
		End If
		SQLWhere = FormatForSQL(SQLWhere,DataBaseType,"CleanUserSQL")
		If SQLID = "" And SQLWhere = "" Then
			SQLID = Request.QueryString("SQLID")
			If Not Allowmultidelete Then
				If Not IsNumeric(SQLID) Then
					SQLID = ""
				End If
			End If
		End If
		SQLID = FormatForSQL(SQLID,DataBaseType,"CleanUserSQL")

		If (SQLID <> "" Or SQLWHERE <> "") Then
			If AllowEdit Then
				EditMode = "Edit"
			Else
				EditMode = "ReadOnly"
			End If
		Else
			If AllowAdd Then
				EditMode = "Add"
			Else
				EditMode = "NotFound"
			End If
		End If
		If Request.Form("ocdEditSave") <> "" Then
			FormMode = "Save"
		Else
			FormMode = "View"
		End If
		SQLSelect = FormatForSQL(SQLSelect,DataBaseType,"CleanUserSQL")
		SQLFrom = FormatForSQL(SQLFrom,DataBaseType,"CleanUserSQL")
		If Request.Form("ocdEditNew") <> "" Then
			strURL = request.servervariables("SCRIPT_NAME") & "?sqlid="
			For Each QS In Request.QueryString
				If UCASE(QS) <> "SQLID" And UCASE(QS) <> "SQLWHERE" Then
					strURL = strURL & "&" & QS & "=" & Server.URLEncode(Request.QueryString(QS))
				End If
			Next
			Call Close()
			response.clear()
			Response.Redirect(strURL)
		ElseIf (request("ocdEditCancel") <> "") And (request("ocdEditCancelPage") <> "") Then
			If CallOnCancel Then
				Call ocdOnCancel()
			End If
			strURL = request("ocdEditCancelPage") & "?sqlid="
			For Each QS In Request.QueryString
				If UCASE(QS) <> "SQLID" And UCASE(QS) <> "OCDEDITDELETE" And UCASE(QS) <> "SQLWHERE" Then
					strURL = strURL & "&" & QS & "=" & Server.URLEncode(Request.QueryString(QS))
				End If
			Next
			response.clear()
			Response.Redirect(strURL)
			response.end()
		ElseIf (Request.Form("ocdEditCancel") <> "") Or ((Request.Form("ocdEditDelete") <> "" Or Request.QueryString("ocdEditDelete") <> "") And request("ocdEditConfirm") <> "" And Not AllowDelete) Then
			If CallOnCancel Then
				Call ocdOnCancel()
			End If
			strURL = request.servervariables("SCRIPT_NAME") & "?sqlid=" & SQLID & "&sqlwhere=" & server.urlencode(SQLWHERE)
			For Each QS In Request.QueryString
				If UCASE(QS) <> "SQLID" And UCASE(QS) <> "OCDEDITDELETE" And UCASE(QS) <> "SQLWHERE" Then
					strURL = strURL & "&" & QS & "=" & Server.URLEncode(Request.QueryString(QS))
				End If
			Next
			Call Close()
			response.clear()
			Response.Redirect(strURL)
		ElseIf (request("ocdEditDelete") <> "") And request("ocdEditConfirm") = "" And CallBeforeDelete Then
			Call ocdBeforeDelete()
		End If
		If Not IsObject(ADOConnection) Then
			Set ADOConnection = server.CreateObject("ADODB.Connection")
			If ADOMode = 0 Then
				If FormMode <> "Save" And Not ((Request.Form("ocdEditDelete") <> "" Or Request.QueryString("ocdEditDelete") <> "") And request("ocdEditconfirm") <> "") Then			 'nothing to update, use read only connection
					ADOConnection.mode = 1				'adModeRead
				End If
			Else
				ADOConnection.Mode = 0 'read write
			End If
			Call ADOConnection.Open(SQLConnect, SQLUser, SQLPass)
			If Err.Number <> 0 Then
				FormStatus = "Connect Error"
				exit sub
			End If
		End If
		If DatabaseType = "" Then
			DatabaseType = getDatabaseType(ADOConnection)
		End If
		If SQLSelectID = "" And SQLSelectPK = "" Then
			Set rsIDX = server.createobject("ADODB.Recordset")
			If UCASE(ADOConnection.provider) <> "MICROSOFT.JET.OLEDB.3.51" Then
				strSQLIDX = "Select * From " & SQLFrom & " WHERE 1 = 2"
				rsIDX.CursorLocation = 2
'			response.write strSQLIDX
				Call rsIDX.Open(strSQLIDX, ADOConnection, 0, 1)				 'adOpenForwardOnly, adLockReadOnly
				If err.number <> 0 Then
					Response.Write(err.description)
					response.end()
				End If
				For Each fldTemp In rsIDX.Fields
					If CBool(fldTemp.Properties("ISAUTOINCREMENT")) = True Then
						SQLSelectID = fldTemp.Name
						Exit For
					End If
				Next
				rsIDX.close()
				Set rsIDX = Nothing
			End If
		End If
		If SQLSelectID = "" And SQLSelectPK = "" And EnablePrimaryKeyLink Then			 'determine primary key fields dynamically
				SQLSelectPK = getPKFields(ADOConnection, DatabaseType, SQLFrom, quoteprefix, quotesuffix)
			End If
		If SQLSelectPK = "" And SQLSelectID = "" And EditMode = "Edit" Then
			EditMode = "View"			 'can't id record for update
		End If
		If (request("ocdEditDelete") <> "") And request("ocdEditConfirm") <> "" Then
			If instr(1, SQLFrom, ",") = 0 Then
				strSQL = "DELETE FROM " & FormatForSQL(SQLFrom, DatabaseType, "AddSQLIdentifier")
			Else
				strSQL = "DELETE FROM " & SQLFrom
			End If
			If SQLID <> "" Then
				If AllowMultiDelete And instr(SQLID, ",") <> 0 Then
					strSQL = strSQL & " WHERE (" & SQLSelectID & " IN ( " & SQLID & " ) "
				Else
					strSQL = strSQL & " WHERE (" & SQLSelectID & " = " & SQLID
				End If
			Else
				strSQL = strSQL & " WHERE (" & SQLWHERE
			End If
			If SQLWhereExtra <> "" Then
				strSQL = strSQL & ") AND " & SQLWhereExtra & ""
			Else
				strSQL = strSQL & ")"
			End If
'			If ((Not Isnumeric(SQLID) Or (EnablePrimaryKeyLink And SQLWHERE <> "")) Then
'				err.Raise(17)
'				Exit Sub
'			Else
				If CallPreDelete Then
					Call ocdPreDelete()
				End If
				ADOConnection.Execute(strSQL)
'			End If
			'have to redirect to a different screen after a delete
			If err.number = 0 Then
				If CallAfterDelete Then
					Call ocdAfterDelete()
				End If
				'if redirect not triggered in sub, goto new record
				strURL = request.servervariables("SCRIPT_NAME") & "?sqlid="
				For Each QS In Request.QueryString
					If UCASE(QS) <> "SQLID" And UCASE(QS) <> "OCDEDITDELETE" And UCASE(QS) <> "SQLWHERE" Then
						strURL = strURL & "&" & QS & "=" & Server.URLEncode(Request.QueryString(QS))
					End If
				Next
				Call Close()
				response.clear()
				Response.Redirect(strURL)
			Else
				FormStatus = FormStatus & err.description
				err.clear()
			End If
		End If
		If sqlSelect = "" Then
			sqlSelect = "*"
		End If
		Set ADORecordset = server.createobject("ADODB.Recordset")
		If FormMode = "View" Then
			strSQL = "Select " & SQLSelect & " from " & FormatForSQL(SQLFrom, DatabaseType, "AddSQLIdentifier")
			If EditMode = "Edit" Then
				If SQLID <> "" Then
					If DatabaseType = "Oracle" Then
						strSQL = strSQL & " Where (" & SQLSelectID & " = " & SQLID
					Else
						strSQL = strSQL & " Where (" & FormatForSQL(SQLSelectID, DatabaseType, "AddSQLIdentifier") & " = " & SQLID
					End If
				Else
					strSQL = strSQL & " Where (" & SQLWHERE
				End If
				If SQLWhereExtra <> "" Then
					strSQL = strSQL & ") AND " & SQLWhereExtra & ""
				Else
					strSQL = strSQL & ")"
				End If
				ADORecordset.CursorLocation = 3
				Call ADORecordset.open(strSQL, ADOConnection, 3, 3)				  'adOpenStatic , adLockReadOnly
				If SQLWHERE <> "" AND SQLWHEREEXtra <> "" Then
					ADORecordset.filter = SQLWhereExtra
				End if
				If Err.Number <> 0 Then
					err.clear()
					EditMode = "NotFound"
					If FormStatus = "" Then
						FormStatus =  HTMLNotFound
					End If
				ElseIf ADORecordset.eof Then
					EditMode = "NotFound"
					If FormStatus = "" Then
						FormStatus =  HTMLNotFound
					End If
				End If
				Set ADORecordset.activeconnection = Nothing
			Else
				ADORecordset.CursorLocation = 3
				strSQL = "SELECT " & SQLSelect & " FROM " & FormatForSQL(SQLFrom, DatabaseType, "AddSQLIdentifier") & " WHERE 1=2"
				Call ADORecordset.open(strSQL, ADOConnection, 3, 1)				'adOpenStatic , adLockReadOnly
				Set ADORecordset.activeconnection = Nothing
			End If
		ElseIf FormMode = "Save" Then
			If (AllowEdit And EditMode = "Edit") Then
				strSQL = "SELECT " & SQLSelect & " FROM " & FormatForSQL(SQLFrom, DatabaseType, "AddSQLIdentifier")
				If SQLSelectID <> "" And SQLWHERE = "" Then
					If DatabaseType = "Oracle" Then
						strSQL = strSQL & " Where (" & SQLSelectID & " = " & SQLID
					Else
						strSQL = strSQL & " WHERE (" & FormatForSQL(SQLSelectID, DatabaseType, "AddSQLIdentifier") & " = " & SQLID
					End If
				Else
					strSQL = strSQL & " WHERE (" & SQLWHERE
				End If
				If SQLWhereExtra <> "" Then
					strSQL = strSQL & ") AND " & SQLWhereExtra & ""
				Else
					strSQL = strSQL & ")"
				End If
				If FormEditConnect And DatabaseType <> "MySQL" And DatabaseType <> "Oracle" Then
					ADORecordset.CursorLocation = 2
				Else
					ADORecordset.CursorLocation = 3
				End If
				Call ADORecordset.open(strSQL, ADOConnection, 3, 3)				' adOpenStatic , adLockOptimistic
				If SQLWhere <> "" And SQLWhereExtra <> "" Then
					ADORecordset.Filter = SQLWhereExtra
				End If
				If Not FormEditConnect Then
					Set ADORecordset.activeconnection = Nothing
				End If
				If err.number <> 0 Then
					FormStatus = FormStatus & err.description
				Else
					If ADORecordset.eof Then
						EditMode = "NotFound"
						FormStatus = FormStatus & HTMLNotFound
					ElseIf CallValidateBeforeUpdate Then
						Call ocdValidateBeforeUpdate()
					ElseIf CallBeforeUpdate Then
						Call ocdBeforeUpdate()
					End If
				End If
			ElseIf (AllowAdd And EditMode = "Add") Then
				strSQL = "SELECT " & SQLSelect & " FROM " & FormatForSQL(SQLFrom, DatabaseType, "AddSQLIdentifier") & " WHERE 1=2"
				If DatabaseType = "MySQL" Or DatabasetYpe = "Oracle" Or Not FormEditConnect Then
					ADORecordset.CursorLocation = 3
				Else
					ADORecordset.CursorLocation = 2
				End If
				If FormEditConnect Then
					Call ADORecordset.open(strSQL, ADOConnection, 3, 3)					  ' adOpenStatic , adLockOptimistic
				Else
					Call ADORecordset.open(strSQL, ADOConnection, 3, 4)					  ' adOpenStatic , adLockOptimistic adLockBatchOptimistic = 4
				End If
				ADORecordset.AddNew()
				If Not FormEditConnect Then
					Set ADORecordset.activeconnection = Nothing
				End If
				If err.number <> 0 Then
					FormStatus = FormStatus & err.description
				ElseIf CallValidateBeforeInsert Then
					Call ocdValidateBeforeInsert()
				ElseIf CallBeforeInsert Then
					Call ocdBeforeInsert()
				End If
			End If
			Dim v1, v2
			If FormStatus = "" And InvalidFields = "" Then
				For Each fmRequest In Request.Form
					Select Case UCASE(fmRequest)
						Case "ocdTFS_GENERATION"
							'Don't process
						Case Else
							Select Case left(fmRequest, 5)
								Case "ocdCT"
									set fldTemp = ADORecordset.Fields(CStr(mid(fmRequest, 6)))
									v1 = fldTemp.Value
									If UCASE(fldTemp.Name) = UCase(SQLSelectID) Then
										'dont update
									ElseIf ((CBool(fldTemp.Attributes And &H20)) And Request.Form("ocdCB" & CStr(mid(fmRequest, 6))) = "") And fldTemp.Type <> 72 Then									'adFldIsNullable=&H00000020
										'null check
										If Not IsNull(v1) Then
											fldTemp.Value = null
										End If
									Else
										Select Case fldTemp.Type
											Case 72											  'adGUID
												'no update
											Case 11											  'adBoolean
												If Request.Form("ocdCB" & CStr(mid(fmRequest, 6))) <> "" Then
													If IsNull(V1) Then
														fldTemp.Value = True
													ElseIf Not v1 Then
														fldTemp.Value = True
													End If
												Else
													If IsNull(V1) Then
														fldTemp.Value = False
													ElseIf v1 Then
														fldTemp.Value = False
													ElseIf UCASE(EditMode) = "ADD" Then
													End If
												End If
											Case Else
												If CStr(Request.Form("ocdCB" & CStr(mid(fmRequest, 6)))) = "" And CStr(Request.Form("ndchf" & CStr(mid(fmRequest, 6)))) = "" Then
													'			if not IsNull(v1) Then
													fldTemp.Value = null
													'			End if
												Else
													If Request.Form("ocdCB" & CStr(mid(fmRequest, 6))) = "" Then
														fldTemp.Value = Request.Form("ndchf" & CStr(mid(fmRequest, 6)))
													Else
														fldTemp.Value = Request.Form("ocdCT" & CStr(mid(fmRequest, 6)))
													End If
													If err.number <> 0 Then
														FormStatus = FormStatus & HTMLUpdateFieldFails & mid(fmRequest, 6) & ": " & err.description & "<BR>"
														InvalidFields = InvalidFields & mid(fmRequest, 6) & ";"
														err.clear()
													End If
												End If
										End Select
									End If
								Case "ocdTF"
									If EditMode = "Add" And Request.Form(fmRequest) = "" Then
										'don't set value, lets default be used
									Else
										set fldTemp = ADORecordset.Fields(CStr(mid(fmRequest, 6)))
										v1 = fldTemp.Value
										If UCASE(fldTemp.Name) = UCase(SQLSelectID) Then
											'dont update
										ElseIf ((CBool(fldTemp.Attributes And &H20)) And (Request.Form(fmRequest) = FormNullToken Or Request.Form(fmRequest) = "")) And fldTemp.Type <> 72 Then										  'adFldIsNullable=&H00000020
											'null check
											If FormNullToken = Request.Form(fmRequest) Then
												If Not IsNull(v1) Then
													fldTemp.Value = null
												End If
											Else
												Select Case fldTemp.Type
													Case 129, 230, 200, 202
														If (v1) = FormNullToken Then
															fldTemp.Value = null
														Else
															fldTemp.Value = ""
														End If
													Case Else
														If Not IsNull(v1) Then
															fldTemp.Value = null
														End If
												End Select
											End If
										Else
											Select Case fldTemp.Type
												Case 72												 'adGUID
												Case 11												 'adBoolean
													Select Case Request.Form(fmRequest)
														Case "True", "Yes", "-1"
															If IsNull(v1) Then
																fldTemp.Value = True
															ElseIf Not v1 Then
																fldTemp.Value = True
															End If
														Case "False", "No", "0"
															If IsNull(v1) Then
																fldTemp.Value = False
															ElseIf v1 Then
																fldTemp.Value = False
															End If
														Case Else
															If Not IsNull(v1) Then
																fldTemp.Value = null
															End If
													End Select
												Case Else
													If CStr(Request.Form(fmRequest)) = "" Then
														If Not IsNull(v1) Then
															fldTemp.Value = null
														End If
													Else
														If fldTemp.Type = 135 And IsDate(Request.Form(fmRequest)) And DatabaseType = "MySQL" Then
															If IsNull(v1) Then
																fldTemp.Value = CDate(Request.Form(fmRequest))
															ElseIf Not v1 = CDate(Request.Form(fmRequest)) Then
																fldTemp.Value = CDate(Request.Form(fmRequest))
															End If
														ElseIf fldTemp.Type = 135 Then
															If IsNull(v1) Then
																fldTemp.Value = CDate(Request.Form(fmRequest))
															ElseIf Not v1 = CDate(Request.Form(fmRequest)) Then
																fldTemp.Value = CDate(Request.Form(fmRequest))
															End If
														ElseIf fldTemp.Type = 3 Then
															v1 = fldTemp.value
															If IsNull(v1) Then
																fldTemp.Value = CInt(Request.Form(fmRequest))
															ElseIf Not CStr((v1)) = CStr((Request.Form(fmRequest))) Then
																fldTemp.Value = CLng(Request.Form(fmRequest))
															End If
														ElseIf fldTemp.Type = 6 Then
															If IsNull(v1) Then
																fldTemp.Value = (Request.Form(fmRequest))
															ElseIf Not CStr(v1) = CStr(Request.Form(fmRequest)) Then
'Response.write (v1)
'response.write "z"
'response.write Request.Form(fmRequest)
																If IsNumeric(Request.Form(fmRequest)) Then
																	If Not CCur(v1) = CCur(Request.Form(fmRequest)) Then
																		fldTemp.Value = (Request.Form(fmRequest))
																	End If
																Else
																	fldTemp.Value = (Request.Form(fmRequest))
																End If
															End If
														ElseIf fldTemp.Type = 5 Then
															If IsNull(v1) Then
																fldTemp.Value = (Request.Form(fmRequest))
															ElseIf Not CStr(v1) = CStr(Request.Form(fmRequest)) Then
'Response.write CStr(v1)
'response.write Request.Form(fmRequest)
																If IsNumeric(Request.Form(fmRequest)) Then
																	If Not CDBl(v1) = CDBl(Request.Form(fmRequest)) Then
																		fldTemp.Value = (Request.Form(fmRequest))
																	End If
																Else
																	fldTemp.Value = (Request.Form(fmRequest))
																End If
															End If

														ElseIf fldTemp.Type = 202 Or fldTemp.Type = 203 Or fldTemp.Type = 201 Or fldTemp.Type = 200 Then
															If IsNull(v1) Then
																fldTemp.Value = CStr(Request.Form(fmRequest))
															ElseIf v1 <> "" And CStr(Request.Form(fmRequest)) = FormEStringToken Then
																fldTemp.Value = ""
															ElseIf v1 = "" And CStr(Request.Form(fmRequest)) = FormEStringToken Then
																fldTemp.Value = ""
															ElseIf Not CStr((v1)) = CStr(Request.Form(fmRequest)) Then
																fldTemp.Value = CStr(Request.Form(fmRequest))
															End If
														ElseIf fldTemp.Type = 131 Or fldTemp.Type = 130 Then
															v2 = (CStr(Request.Form(fmRequest)))
															If IsNull(v1) Then
																fldTemp.Value = Request.Form(fmRequest)
															ElseIf Not CStr((v1)) = (v2) Then
																fldTemp.Value = Request.Form(fmRequest)
															End If
														Else
'															response.write fldTemp.Type
															fldTemp.Value = Request.Form(fmRequest)
														End If
														Select Case Err.number
															Case 0
															Case -2147217887
																FormStatus = FormStatus & " " & HTMLUpdateFieldFails & " " & mid(fmRequest, 6) & ": " & HTMLOLEError & "<BR>"
																InvalidFields = InvalidFields & mid(fmRequest, 6) & ";"
																err.clear()
															Case Else
																FormStatus = FormStatus & " " & HTMLUpdateFieldFails & " " & mid(fmRequest, 6) & ": " & err.description & "<BR>"
																InvalidFields = InvalidFields & mid(fmRequest, 6) & ";"
																err.clear()
														End Select
													End If
											End Select
										End If
									End If
								Case Else								  'don't process
							End Select							'form field prefix
					End Select					  'form field name
					If err.number <> 0 Then
						FormStatus = FormStatus & HTMLStatusFieldError & mid(fmRequest, 6) & ": " & err.description & "<BR>"
						InvalidFields = InvalidFields & mid(fmRequest, 6) & ";"
						err.clear()
					End If
				Next				'Form Field to update
			End If			 'Update form field switch
			If CallValidateBeforeUpdate And CallBeforeUpdate And EditMode = "Edit" Then
				Call ocdBeforeUpdate()
			ElseIf CallValidateBeforeInsert And CallBeforeInsert And EditMode = "Add" Then
				Call ocdBeforeInsert()
			End If
			If InvalidFields = "" And FormStatus = "" Then
				If Not FormEditConnect Then
					ADORecordset.ActiveConnection = ADOConnection
					ADORecordset.UpdateBatch()
				Else
					ADORecordset.Update()
				End If
				'			response.end
				If err.number <> 0 Then
					FormStatus = FormStatus & err.description & "<BR>"
					ADORecordset.CancelUpdate()
					err.clear()
'					On Error GoTo 0
				Else
					If EditMode = "Add" Then
						If EditMode = "Add" And FormMode = "Save" And Not DatabaseType = "Oracle" And DatabaseType <> "MySQL" Then						  'redirect to newly added record
							ADORecordset.movelast()
						End If
						If CallAfterInsert Then
							Call ocdAfterInsert()
						End If
					Else
						If CallAfterUpdate Then
							Call ocdAfterUpdate()
						End If
					End If
				End If
			Else
				ADORecordset.CancelUpdate()
			End If
			If Not FormEditConnect Then
				Set ADORecordset.ActiveConnection = Nothing
			End If
			If FormStatus = "" And InvalidFields = "" Then			 'start record redirect
				If SQLSelectID <> "" Then
					strURL = request.servervariables("SCRIPT_NAME") & "?sqlwhere=&sqlid="
					On Error Resume Next
					Dim tmpCNewID
					tmpCNewID = ADORecordset.Fields(SQLSelectID).Value
					If err.number <> 0 Then
						err.clear()
						Call ADORecordset.Resync(1, 2)
						tmpCNewID = ADORecordset.Fields(SQLSelectID).Value
						If err.number <> 0 Then
							tmpCNewID = ""
						End If
						err.clear()
					End If
					strURL = strURL & tmpCNewID
				Else				'determine primary key to redirect
					strURL = request.servervariables("SCRIPT_NAME") & "?sqlid=&sqlwhere="
					strSQLWHEREPK = ""
					strCDp = ""
					arrSQLSelectPK = split(SQLSelectPK, ",")
					For intTemp = 0 To UBound(arrSQLSelectPK)
						strSQLWHEREPK = strSQLWHEREPK & FormatForSQL(CStr(arrSQLSelectPK(intTemp)), DatabaseType, "AddSQLIdentifier") & "="
						Select Case ADORecordset.Fields(arrSQLSelectPK(intTemp)).Type
							Case 2, 3, 4, 5, 14, 16, 17, 18, 19, 20, 21, 128, 131, 204, 6, 11							'adSmallInt, adInteger, adSingle, adDouble, adDecimal, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adBigInt, adUnsignedBigInt, adBinary, adNumeric, adVarBinary, adLongVarBinary, adCurrency, adBoolean
								strCDp = ""
							Case 135, 7, 133, 134							'adDBTimeStamp, adDate, adDBDate,  adDBTime
								If DatabaseType = "Access" Then
									strCDp = "#"
								Else
									strCDp = "'"
								End If
							Case 8, 129, 130, 200, 201, 202, 203							'adBSTR, adChar, adWChar, adVarChar, adLongVarChar, adVarWChar, adLongVarWChar						
								strCDp = "'"
						End Select
						Select Case ADORecordset.Fields(arrSQLSelectPK(intTemp)).Type
							Case 8, 129, 130, 200, 201, 202, 203							'adBSTR, adChar, adWChar, adVarChar, adLongVarChar, adVarWChar, adLongVarWChar		
								strSQLWHEREPK = strSQLWHEREPK & strCDp & Replace(ADORecordset.Fields(arrSQLSelectPK(intTemp)).Value, "'", "''") & strCDp & " AND "
							Case 135, 7, 133, 134							'adDBTimeStamp, adDate, adDBDate,  adDBTime
								strSQLWHEREPK = strSQLWHEREPK & strCDp & FormatForSQL(ADORecordset.Fields(arrSQLSelectPK(intTemp)).Value, DatabaseType, "SafeDate") & strCDp & " AND "
							Case Else
								strSQLWHEREPK = strSQLWHEREPK & strCDp & ADORecordset.Fields(arrSQLSelectPK(intTemp)).Value & strCDp & " AND "
						End Select
					Next
					If strSQLWHEREPK <> "" Then
						strSQLWHEREPK = left(strSQLWHEREPK, len(strSQLWHEREPK) - 5)
					End If
					strURL = strURL & Server.URLEncode(strSQLWHEREPK)
				End If
				For Each QS In Request.QueryString
					If UCASE(QS) <> "SQLID" And UCASE(QS) <> "SQLWHERE" Then
						strURL = strURL & "&" & QS & "=" & Server.URLEncode(Request.QueryString(QS))
					End If
				Next
				Call Close()
				response.clear()
				Response.Redirect(strURL)
			End If			 'End recordset redirect
		End If		  'End Form Mode Switch
	End Sub
	Public Sub Display(ByVal strTemplate)
		Dim strTemp, QS
		Select Case UCASE(strTemplate)
			Case "STATUS"
				If FormStatus <> "" Then
					Response.Write(HTMLStatusStart & FormStatus & HTMLStatusEnd)
				End If
			Case "START"
				strTemp = ""
				strTemp = strTemp & "<form method=""post"" action=""" & request.servervariables("SCRIPT_NAME") & "?"
				For Each QS In Request.QueryString
					If UCASE(QS) <> "OCDEDITDELETE" Then
						strTemp = strTemp & QS & "=" & Server.URLEncode(Request.QueryString(QS)) & "&"
					End If
				Next
				strTemp = strTemp & """ " & HTMLAttribForm & ">"
				Response.Write(strTemp)
			Case "END"
				Response.Write("</form>")
			Case "BUTTONS"
				strTemp = ""
				If EditMode <> "NotFound" Then
					If instr(SQLID, ",") = 0 Then
						If AllowAdd Or AllowEdit Then
							strTemp = strTemp & "<input type=""hidden"" name=""ocdCSSFix"">"
							strTemp = strTemp & "<input name=""ocdEditSave"" " & HTMLAttribSaveBtn & ">"
						End If
						If AllowAdd And (SQLID <> "" Or SQLWHERE <> "") Then
							strTemp = strTemp & "<input name=""ocdEditNew"" " & HTMLAttribNewBtn & ">"
						End If
						strTemp = strTemp & "<input name=""ocdEditCancel"" " & HTMLAttribCancelBtn & ">"
						If AllowDelete And (SQLID <> "" Or SQLWHERE <> "") Then
							strTemp = strTemp & "<input name=""ocdEditDelete"" " & HTMLAttribDeleteBtn & ">"
						End If
						Response.Write(strTemp)
					End If
				End If
		End Select
	End Sub

	Public Sub Close()
		On Error Resume Next
		ADORecordset.close()
		Set ADORecordset = Nothing
		ADOConnection.close()
		Set ADOConnection = Nothing
		err.clear()
	End Sub
End Class
%>