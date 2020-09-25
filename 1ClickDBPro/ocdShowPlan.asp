<%
'1 Click DB Pro copyright 1997-2003 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties. Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com
'Email inquiries to info@1ClickDB.com
    
'**Start Encode**


Class ocdShowPlan

	Public ADOCommand

	Private Sub Class_Initialize
	
	End Sub
	
	Public Sub Display (strTemplate)
		Dim rsTemp, intPrevErr, objErr, connTemp, strST, intCount, strOut, varCCount, varCountdown, varRealCost, varLastCount, varTotalCost, varTotalIO, varTotalCPU
		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		Set connTemp = ADOCommand.ActiveConnection
		Set rsTemp = ADOCommand.Execute	
		If Err.Number <> 0 Then
			Response.Write(Err.Description)
			Err.clear
			Exit Sub
		End If
		intCount = 0
		Do Until (rsTemp is nothing)
			Response.Flush
			If connTemp.Errors.Count > 0 Then 
				Response.Write("<P>")
				intPrevErr = 0
				For Each objErr In connTemp.Errors
					If objErr.number = 0 Then
						If objErr.number <> 0 And intPrevErr <> objErr.Number Then
							Response.Write("<IMG SRC=AppWarningSmall.gif height=11 width=12 border=0><FONT COLOR=RED><B>")
							Response.Write(objErr.Number)
							Response.Write("</b></font>")
						End If
						Response.Write(" " & objErr.Description)
					End If
					intPrevErr = objErr.number
					Response.Write("<BR>")
				Next
				If intPrevErr <> 0 then	
					Exit Sub
				End If
				Response.Write("</P>")
			End If
			intcount = intcount + 1
			If rsTemp.state = 0 Then 'adstateclosed
				Set rsTemp = rsTemp.nextRecordset
			Else
				Response.Write("<table border=0 cellspacing=2 cellpadding=2 Class=Grid>" )
				varTotalCost = 0
				varTotalIO = 0
				varTotalCPU = 0
				varRealCost = 0
				Do While Not rsTemp.EOF
					If varRealCost = 0 and Not isnull(rsTemp("TotalSubtreeCost")) Then
						varRealCost = rsTemp("TotalSubtreeCost")
					ElseIf varTotalCost <> 0 Then
						varRealCost = varRealCost + rsTemp("TotalSubtreeCost")
					End If
					If Not isnull(rsTemp("EstimateCPU")) Then
						varTotalCPU = varTotalCPU + rsTemp("EstimateCPU")
					End If
					If Not isnull(rstemp("EstimateIO")) Then
						varTotalIO = varTotalIO + rstemp("EstimateIO")
					End If
					rsTemp.MoveNext
				Loop
				rsTemp.MoveFirst
				varLastCount = 0 
				Do While Not rsTemp.EOF
					If varTotalCost =0 and Not isnull(rsTemp("TotalSubtreeCost")) Then
						varTotalCost = rsTemp("TotalSubtreeCost")
					ElseIf varTotalCost <> 0 Then
						varCountdown = varCountdown - rsTemp("TotalSubtreeCost")
					End If
					Response.Write("<TR>")
					Response.Write("<TD VALIGN=TOP ALIGN=RIGHT NOWRAP>")
					strOut = ""
					Select Case rsTemp("Type")
						Case "PLAN_ROW"
							Select Case rsTemp("PhysicalOp")
								Case "Compute Scalar"
									strOut = strOut & "<IMG SRC=appECComputeScalar.gif BORDER=0 ALT=""COMPUTE SCALAR" & vbCRLF & vbCRLF
								Case "Sort"
									strOut = strOut & "<IMG SRC=GRIDLNKASC.GIF BORDER=0 ALT=""SORT" & vbCRLF & vbCRLF
								Case "Filter"
									strOut = strOut & "<IMG SRC=GRIDLNKFILTER.GIF BORDER=0 ALT=""FILTER" & vbCRLF & vbCRLF
								Case "Stream Aggregate"
									strOut = strOut & "<IMG SRC=appECStreamAgg.gif BORDER=0 ALT=""STREAM AGGREGATE" & vbCRLF & vbCRLF
								Case "Hash Match"
									strOut = strOut & "<IMG SRC=appECHashMatch.gif BORDER=0 ALT=""HASH MATCH" & vbCRLF & vbCRLF
								Case "Top"
									strOut = strOut & "<IMG SRC=appECTop.gif BORDER=0 ALT=""TOP" & vbCRLF & vbCRLF
								Case "Clustered Index Scan"
									strOut = strOut & "<IMG SRC=appECClusteredISeek.gif BORDER=0 ALT=""CLUSTERED INDEX SCAN " & vbCRLF & vbCRLF
								Case "Bookmark Lookup"
									strOut = strOut & "<IMG SRC=appECClusteredISeek.gif BORDER=0 ALT=""BOOKMARK LOOKUP" & vbCRLF & vbCRLF
								Case "Clustered Index Seek"
									strOut = strOut & "<IMG SRC=appECClusteredISeek.gif BORDER=0 ALT=""CLUSTERED INDEX SEEK " & vbCRLF & vbCRLF
								Case "Index Seek"
									strOut = strOut & "<IMG SRC=appECClusteredISeek.gif BORDER=0 ALT=""INDEX SEEK" & vbCRLF & vbCRLF
								Case "Nested Loops"
									strOut = strOut & "<IMG SRC=appECNestedLoop.gif BORDER=0 ALT=""NESTED LOOPS " & vbCRLF & vbCRLF		
								Case Else
									strOut = strOut & "<IMG SRC=appSelect.gif BORDER=0 ALT="""
									strOut = strOut & UCase(rsTemp("PhysicalOp"))
									strOut = strOut & vbCRLF & vbCRLF
							End Select
							If Not IsNull(rsTemp("Warnings")) Then
								strOut = strOut & "Warnings : " & Server.HTMLEncode(rsTemp("Warnings")) & vbCRLF & vbCRLF
							End If
							strOut = strOut & "Physical Op : " & rsTemp("PhysicalOp") & vbCRLF
							strOut = strOut & "Logical Op : " & rsTemp("LogicalOp") & vbCRLF
							strOut = strOut & "Est Rows : " & rsTemp("EstimateRows") & vbCRLF
							strOut = strOut & "Avg Row Size : " & rsTemp("AvgRowSize") & vbCRLF
							strOut = strOut & "Est CPU : " & rsTemp("EstimateCPU") & vbCRLF
							strOut = strOut & "Est IO : " & rsTemp("EstimateIO") & vbCRLF
							strOut = strOut & "Est Execs : " & rsTemp("EstimateExecutions") & vbCRLF
							strOut = strOut & "Subtree Cost : " & rsTemp("TotalSubtreeCost") 
							strOut = strOut & """>"
						Case Else
							strOut = strOut & "<IMG BORDER=0 ALT=""" & UCase(rsTemp("Type")) & vbCRLF & vbCRLF
							strOut = strOut & "Est Rows : " & rsTemp("EstimateRows") & vbCRLF 
							strOut = strOut & "Subtree Cost : " & rsTemp("TotalSubtreeCost")
							strOut = strOut & """ SRC=appTable.gif>"
					End Select
					strST = rsTemp("StmtText")
					varLastCount = varTotalCost
					If Not IsNull(rsTemp("EstimateCPU")) And Not isnull(rsTemp("EstimateIO")) Then
						varCcount = Round((100 * ( (rsTemp("EstimateIO") + rsTemp("EstimateCPU"))/ (varTotalIO + varTotalCPU) )))
						Response.Write(varCCount & Server.HTMLEncode("%"))
					End If
					Response.Write("</TD>")
					Response.Write("<TD VALIGN=TOP NOWRAP>")
					strST = Replace(strST," ","&nbsp;")
					If InStr(strST,"|--") > 0 Then
						strST = Replace(strST,"|--", ("|--" & strOut & " "))
					Else
						strST = strOut & " " & strST
					End If
					If Not IsNull(rsTemp("Warnings")) Then
						Response.Write("<SPAN CLASS=Warning>")
						Response.Write(strST)
						Response.Write("</SPAN>")
					Else
						Response.Write(strST)
					End If
					Response.Write("</TD>")
	 				Response.Write("</TR>")
					varLastCount = rsTemp("TotalSubtreeCost")
					rsTemp.MoveNext
				Loop
				Response.Write("</table>")
				Set rsTemp = rsTemp.nextRecordset
				If Err.Number = 0 Then	
					If connTemp.errors.Count > 0 Then 
						Response.Write("<P>")
						intPrevErr = 0
						For Each objErr In connTemp.Errors
							intPrevErr = Err.Number
							If objErr.Number <> 0 and intPrevErr <> objErr.Number Then
								Response.Write("<IMG SRC=AppWarningSmall.gif height=11 width=12 border=0><FONT COLOR=RED><B> " & objErr.Number)
								Response.Write(objErr.description)
								Response.Write("<BR>")
							End If
						Next
						Response.Write("</P>")
					End If
				End If
			End If	
		Loop
		Set rsTemp = Nothing
	End Sub
End Class	
%>
