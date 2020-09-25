<%
'1 Click DB Pro copyright 1997-2004 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com/

        
'**Start Encode**


Class ocdCommand
	Public CommandID
	Public ADOCommand
	Public ADORecordset
	Public ADOConnection
	Public ADOCommandRecordsAffected
	Public FormatCurrencyAsNumber
	Public ShowOutputParamaters
	Public ShowInfoMessages
	Public ShowRecordsAffected
	Public ShowCommandText
	Public FormatNumberFields
	Public FormatNumberNumDigitsAfterDecimal 
	Public FormatNumberIncludeLeadingDigit 
	Public FormatNumberUseParensForNegativeNumbers 
	Public FormatNumberGroupDigits
	Public FormatNumberNegativePrefix
	Public FormatNumberNegativeSuffix
	Public ShowPercentileFields
	Public ShowSortLinks
	Public UseClientCursor
	Public QuotePrefix
	Public QuoteSuffix
	Public ocdCommConn
	Public HTMLMin
	Public HTMLMax
	Public HTMLAvg
	Public HTML25P
	Public HTML50P
	Public HTML75P
	Public HTLMSum
	Public MaxRecordsDisplay
	'**Start Encode**
	Public DatabaseType

Private sub Class_Initialize
MaxRecordsDisplay = 1000
	UseClientCursor = False
	FormatNumberUseParensForNegativeNumbers = -2
	FormatNumberIncludeLeadingDigit = -2
	FormatNumberFields = ""
	CommandID = "A"
	FormatNumberNumDigitsAfterDecimal = 0
	FormatNumberGroupDigits = -2
	ShowOutputParamaters = True
	FOrmatCurrencyAsNumber = False
	ShowInfoMessages = True
	ShowRecordsAffected = True
	ShowCommandText = True
	ShowPercentileFields = ""
	ShowSortLinks = False
	QuotePrefix = """"
	QuoteSuffix = """"
	HTML50P = "<SPAN CLASS=FieldName>50th Percentile:</SPAN>"
end sub
Public Sub Open()
'response.write "X"
	set ADORecordset = server.createobject("ADODB.Recordset")
	set ocdCommConn = ADOCommand.ActiveConnection
	DatabaseType = getDatabaseType(ocdCommConn)
	if ShowPercentileFields <> ""  THen
		useclientcursor = true
	End if
	if UseClientCursor Then

		ADORecordset.CursorLocation = 3

		ADORecordset.Open ADOCommand, , 2, 1
	Else
			
		set ADORecordset = ADOCommand.Execute ( ADOCommandRecordsAffected	)
	End if
	if err <> 0 Then
		response.write err.description
		err.clear
		exit sub
	end if
	if UseClientCursor  THen
		set ADORecordset.ActiveConnection = nothing
		if err <> 0 Then
			response.write err.description
			err.clear
			exit sub
		end if
	end if
	DatabaseType = getDataBaseType(ocdCommConn)
	end sub	
Public Sub Display (strTemplate)
	if not IsObject(ADORecordset) Then
		call open()
	End if
	on error resume next
	dim param,varretval,pren,perr,cfields,varval, qsTemp, strST, intcount, intDRCount

'	Response.write lngRecordsAffected
	if ShowRecordsAffected Then
		if ADOCommandRecordsAffected = -1 then
			'n/a
		Else
			IF ADOCommandRecordsAffected <> "" THen 
				Response.Write ("<P>" & Cstr( ADOCommandRecordsAffected ) & " Records Affected<P>")
			End if
		end if
	End if
	intCount = 0
	do until ((ADORecordset is nothing) ) 'while isobject(rsTemp)
		response.flush
		if ocdCommConn.errors.Count > 0 and DatabaseType = "SQLServer" Then 
			Response.write ("<P>")
			pren = 0
			for each perr in ocdCommConn.errors
				if perr.number = 0 and ShowInfoMessages Then
					If perr.number <> 0 and pren <> perr.number Then
						Response.write ("<IMG SRC=AppWarningSmall.gif height=11 width=12 border=0><FONT COLOR=RED><B>")
						response.write perr.number 
						Response.write ("</b></font>")
					End if
					response.write " " & perr.description
				End if
				pren = perr.number
				Response.write "<BR>"
			next
			if pren <> 0 then	
				exit sub
			End if
			Response.write ("</P>")
		End if
		intcount = intcount + 1

		if ADORecordset.state = 0 and DatabaseType = "SQLServer" and not UseClientCursor then 'adstateclosed
			set ADORecordset = ADORecordset.nextRecordset
		Else
			if ADORecordset.State <> 0 Then
				Response.write ("<table border=0 cellspacing=2 cellpadding=2 Class=Grid><tr CLASS=GridHeader>" )
   		 	  	'Header with name of fields
	   	   		For cFields=0 to ADORecordset.Fields.count-1
	       			Response.write ("<th VALIGN=Bottom> ")
					if ShowSortLinks Then
						Response.write "<A HREF="""
						Response.write request.servervariables("SCRIPT_NAME")
						Response.write "?"
						Response.write "sortorder=" 
						if request.querystring("sortorder") = "[" & ADORecordset.Fields(cFields).Name & "]" & " ASC" Then
							Response.write  Server.URLEncode("[" & ADORecordset.Fields(cFields).Name & "]" & " DESC")
						Else
							Response.write  Server.URLEncode("[" & ADORecordset.Fields(cFields).Name & "]" & " ASC")
						End if
						for each qsTemp in request.querystring
							Select Case qsTemp
							Case "sortorder"
							Case Else
								response.write "&amp;" & qstemp & "=" & server.urlencode(request.querystring(qsTemp))
							End Select							
						next
						Response.write """>"
						Response.write (Trim(ADORecordset.Fields(cFields).Name))
						Response.write "</A>"
					Else
						Response.write (Trim(ADORecordset.Fields(cFields).Name))
					End if
					Response.write ("</th>")
		   		Next
		   	   	Response.write ("</tr>")
				if ShowSortLinks <> "" Then
					'response.write request.querystring("sortorder")
					ADORecordset.sort = request.querystring("sortorder")
				End if
				intDRCount = 0
				do while not ADORecordset.eof
					intDRCount = intDRCount + 1
					if intDRCount mod 2 = 0 THen
						response.write ("<TR Class=GridEven>")
					Else
						response.write ("<TR Class=GridOdd>")
					End if
					For cFields=0 to ADORecordset.Fields.count-1
						varval = ""
						varval = ADORecordset.Fields(cFields).Value
						if strTemplate = "all" or strTemplate = "text" Then
							Response.write ("<TD NOWRAP VALIGN=TOP ALIGN=LEFT>")
						Else
							response.write ("<TD VALIGN=TOP ")
							Select Case ADORecordset.Fields(cFields).Type
								Case 11 'adBoolean
									response.write ("ALIGN=""CENTER""")
								Case 2, 3, 4, 5, 14, 16, 17, 18, 19, 20, 21, 128, 131, 204, 6 'adSmallInt, adInteger, adSingle, adDouble, adDecimal, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adBigInt, adUnsignedBigInt, adBinary, adNumeric, adVarBinary, adLongVarBinary, adCurrency, 
									response.write ("ALIGN=""RIGHT""")
							Case Else
									response.write ("ALIGN=""LEFT""")
							End Select
							response.write (" >")
						End if
						If not isnull(varval) Then
						Select Case ADORecordset.Fields(cFields).Type
						Case 5, 6
								if ADORecordset.Fields(cFields).Type = 6 and not FOrmatCurrencyAsNumber Then
								Response.write (server.htmlencode(FormatCurrency(varval)))
	
								Else
								if FormatNumberNegativePrefix <> "" THen
									if varval < 0 Then
										Response.write FormatNumberNegativePrefix
									End if
								End if
								Select Case FormatNumberFields
									Case "*"
										Response.write (server.htmlencode(formatnumber(varval,FormatNumberNumDigitsAfterDecimal)))
									Case ""
										Response.write (server.htmlencode((varval)))
									Case Else
										Response.write (server.htmlencode((varval)))
								End Select
								if FormatNumberNegativeSuffix <> "" Then
									if varVal < 0 THen
										Response.write FormatNumberNegativeSuffix
									end if
								End if
								End if
							Case Else
								Response.write (server.htmlencode(varval))
							End Select
						End if
						Response.write ("</TD>")
				   	Next
		   			
					on error resume next
					ADORecordset.movenext
					if err<> 0 Then
						response.write "<TR><TD COLSPAN=" & ADORecordset.FIelds.Count & ">"
						response.write "<SPAN CLASS=Warning>"
						response.write err.description
						Response.write "</SPAN>"
						Response.write "</TD></TR>"
						'response.end
						err.clear
					end if
					on error goto 0
					Response.write ("</TR>")
					if intDRCount > MaxRecordsDisplay - 1 Then
						do while not ADORecordset.eof
						intDRCount = intDRCount + 1
						ADORecordset.movenext
						loop
					end if

				loop
'				call DisplayStats("")
				Response.write ("</table>")
									if intDRCount > MaxRecordsDisplay - 1 Then
						response.write "<span class=""warning"">Display Limit Exceeded - First " & MaxRecordsDisplay & " of " & intDRCount & " Records Displayed </span>"
					end if

				response.write "<P>"
			End if
			if DatabaseType = "SQLServer" and not UseClientCursor Then
				set ADORecordset = ADORecordset.nextRecordset
				if err.number = 0 Then	
					if ocdCommConn.errors.Count > 0 Then 
						response.write ("<P>")
						pren = 0
						for each perr in ocdCommConn.errors
							if perr.number = 0 and Not ShowInfoMessages Then
								If perr.number <> 0 and pren <> perr.number Then
									Response.write ("<IMG SRC=AppWarningSmall.gif height=11 width=12 border=0><FONT COLOR=RED><B> " & perr.number)
								End if
								response.write  perr.description
								if perr.number <> 0 and pren <> perr.number Then
									Response.write ("</b></font>")
								End if
							End if
							pren = err.number
							Response.write "<BR>"
						next
						Response.write "</P>"
					end if
				End if
			Else
				set ADORecordset = nothing
			End if
		End if	
	loop
	set ADORecordset = Nothing
		If ShowOutputParamaters and DatabaseType = "SQLServer" Then
		For Each param In ADOCommand.Parameters
			Select Case param.direction
				Case &H0004
					Response.write "<P>"
					Response.write "<SPAN CLASS=FieldName>"
					Response.write server.htmlencode(param.name)
					Response.write " : </SPAN>"
					varretval = param.value
					if not isnull(varretval) then
						response.write Server.HTMLEncode(varretval)
					end if
					Response.write "</P>"
			End SElect
		next
	End if

end sub
public sub DisplayStats(strTemplate)
	dim fldCalc, intrsTempCount, intPLoopCount, intRunningCount, intfldCalc
	intrsTempCount = ADORecordset.Recordcount
	if intrsTempCOunt < 2 Then
		exit sub
	Else
	end if
	Select Case UCase(strTemplate)
		Case "TABLE"
			Response.write "<TABLE CLASS=Grid>"
			Response.write "<TR CLASS=GridHeader>"
			for each fldCalc in ADORecordset.Fields
				Response.write "<TH VALIGN=BOTTOM>"
	'			response.write fldCalc.type
				Select Case fldCalc.Type
				Case 5, 6, 3
					Response.write server.htmlencode(fldCalc.name)
				Case Else
					Response.write "&nbsp"
				End Select
				Response.write "</TH>"
			next
			Response.write "</TR>"
		Case Else
				if showPercentileFields <> "" Then
				Response.write "<TR CLASS=GridStatistics>"
				Response.write "<TD COLSPAN="""
				Response.write ADORecordset.Fields.Count

				Response.write """><SPAN CLASS=FieldName>Group Statistics:</SPAN></TD></TR>"
				End if
	End Select	
	Select Case ShowPercentileFields 
		Case "*"
			on error goto 0
			intRunningCount= 0
			if intrsTempCount > 0 THen
				
				Response.write "<TR CLASS=GridStatistics>"
				intfldCalc = 0
				Response.write "<TD VALIGN=TOP NOWRAP>"
				if intrsTempCount > 10 Then
					Select Case UCase(strTemplate)
		Case "TABLE"
				Response.write "<SPAN CLASS=FieldName>"

				Response.write "Minimum Value:"
				Response.write "</SPAN><BR>"
				Response.write "<SPAN CLASS=FieldName>"
				Response.write "25th Percentile:"
				Response.write "</SPAN><BR>"
				Response.write HTML50P
				Response.write "<BR>"
				Response.write "<SPAN CLASS=FieldName>"
				Response.write "75th Percentile:"
				Response.write "</SPAN><BR>"
				Response.write "<SPAN CLASS=FieldName>"
				Response.write "Maximum Value:"
				Response.write "</SPAN><BR>"
		Case Else
				Response.write "<SPAN CLASS=FieldName>"
				Response.write "MIN:"
				Response.write "</SPAN><BR>"
				Response.write "<SPAN CLASS=FieldName>"
				Response.write "25%:"
				Response.write "</SPAN><BR>"
				Response.write "<SPAN CLASS=FieldName>"
				Response.write "50%:"
				Response.write "</SPAN><BR>"
				Response.write "<SPAN CLASS=FieldName>"
				Response.write "75%:"
				Response.write "</SPAN><BR>"
				Response.write "<SPAN CLASS=FieldName>"
				Response.write "MAX:"
				Response.write "</SPAN><BR>"
		End select
				End if
				Response.write "<SPAN CLASS=FieldName>"
				Response.write "Average Value:"
				Response.write "</SPAN><BR>"
				Response.write "</TD>"
dim Rcrdnbr1, PercentBetween, Record_Holder, Prct,percentile, intPCL, lngRecordCount, varCalcVal
				intfldCalc = 0
				for each fldCalc in ADORecordset.Fields
					intfldCalc = intfldCalc + 1
					if intfldCalc > 1 Then
					Response.write "<TD VALIGN=TOP ALIGN=RIGHT>"
						select case fldCalc.type
						Case 5, 6, 3
							ADORecordset.Filter = ""
							ADORecordset.Filter = "[" & fldCalc.Name & "] <> Null"
							ADORecordset.sort = "[" & Cstr(fldCalc.name) & "] ASC"

						if intrsTempCount > 10 Then
							ADORecordset.MoveFirst
							response.write FormatNumber(fldCalc.Value,0)
							Response.write "<BR>"
dim varTargetPosition

'Percentiles
lngRecordCount = ADORecordset.RecordCount
for intPCL = 25 to 75 step 25
	ADORecordset.MoveFirst
	varTargetPosition = (lngRecordCount * (intPCL / 100))
	If Round(varTargetPosition) <> (varTargetPosition)  Then
		ADORecordset.Move CLng(Round(varTargetPosition))
		varCalcVal = fldCalc.value
		ADORecordset.MoveNext
		if 1=2 Then 
			varCalcVal = ((varCalcVal + fldCalc.value) / 2)'	
		Else
			varCalcVal = (varCalcVal + ((fldCalc.value - varCalcVal) * ((varTargetPosition) - Round(varTargetPosition))))'	
		End if
	Else
		ADORecordset.Move CLng(varTargetPosition)
		varCalcVal = fldCalc.value
	End If
	Response.write FormatNumber(varCalcVal,0)
	Response.write "<BR>"
next
'End Pct


'MAX
	ADORecordset.MoveLast
	Response.write FormatNumber(fldCalc.value,0)
	Response.write "<BR>"
'SUM + AVERAGE
End if
intRunningCount = 0
lngRecordCount = 0

IF NOT adorECORDSET.EOF THEN
ADORecordset.MoveFirst
do while not ADORecordset.eof
	if not isnull(fldCalc.Value) then
	lngRecordCount = lngRecordCount + 1

	intrunningcount = intrunningcount + fldCalc.Value
	End if
	ADORecordset.MoveNext
loop
'rESPONSE.WRITE "x"
eND IF
'AVG
if lngRecordCount > 0 Then
Response.write FormatNumber((CCur(intrunningcount) / CCur(lngRecordCount)),0)
End if
if ShowSum Then
Response.write "<BR>"
Response.write FormatNumber(intrunningcount,0)
end if

Case Else
	End Select	
					Response.write "</TD>"
					End if
				Next
End if
End Select	


		Select Case UCase(strTemplate)
		Case "TABLE"
			Response.write "</TABLE>"
	End Select		
ADORecordset.Filter = ""
ADORecordset.Sort = ""	
ADORecordset.MoveFirst
End Sub

End Class	
%>
