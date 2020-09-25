<%
'1 Click DB Pro copyright 1997-2004 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com
'Email inquiries to info@1ClickDB.com
        
'**Start Encode**

Class Loadit
Public TotalBytes
Private Sub Class_Initialize
	TotalBytes = 0
End Sub
Public Function GetFileInput()
	Dim varDelimiter, intOffset, intPosition, intCount, varValue, varTemp, intPositionStart, intPositionEnd, intPositionNew, varOutput, blnIsBinaryData, varAllData, varTotalBytes
varTotalBytes = Request.TotalBytes
	TotalBytes = varTotalBytes
		if  varTotalBytes = 0 Then
		GetFileInput = ""
		exit function
	End if
	if  varTotalBytes > 100000 Then
			GetFileInput = ""
		exit function
	End if
	varAllData = Request.BinaryRead(varTotalBytes)
	varDelimiter = MidB(varAllData, 1, InstrB(1, varAllData, ChrB(13)) - 1)
	intOffset = LenB(varDelimiter)
	intPosition = 1
	intCount = 1
	While intCount > 0
		intCount = InStrB(intPosition, varAllData, varDelimiter)
		varTemp = intCount - intPosition
		If varTemp > 1 Then
			varValue = MidB(varAllData, intPosition, varTemp)
			intPositionStart = 1 + InStrB(1, varValue, ChrB(34))
			intPositionEnd = InStrB(intPositionStart + 1, varValue, ChrB(34))
			intPositionNew = intPositionEnd
			If InStrB(1, varValue, ConvertStringToByte("Content-Type")) > 1 Then
				intPositionStart = 1 + InStrB(intPositionEnd + 1, varValue, ChrB(34))
				intPositionEnd = InStrB(intPositionStart + 1, varValue, ChrB(34))
				intPositionStart = 14 + InStrB(intPositionEnd + 1, varValue, ConvertStringToByte("Content-Type:"))
				intPositionEnd = InStrB(intPositionStart, varValue, ChrB(13))
				Select Case trim(ConvertByteToString(MidB(varValue, intPositionStart, intPositionEnd - intPositionStart)))
					Case "application/octet-stream"
						blnIsBinaryData = False
					Case Else
						blnIsBinaryData = True
				End select
				intPositionStart = intPositionEnd + 4
				intPositionEnd = LenB(varValue)
				varOutput = MidB(varValue, intPositionStart, intPositionEnd - intPositionStart)
			Else
				varOutput = trim(ConvertByteToString(MidB(varValue, intPositionNew + 2)))
			End If
		End If
		intPosition = intOffset + intCount
	Wend
	if blnIsBinaryData then
		getFileInput = ConvertByteToString(varOutput)
	else
		getFileInput = varOutput
	end if
End Function

Private Function ConvertStringToByte(ByteToConvert)
	Dim strChar, intI, varTemp
	For intI = 1 to Len(ByteToConvert)
	 	strChar = Mid(ByteToConvert, intI, 1)
		varTemp = varTemp & chrB(AscB(strChar))
	Next
	ConvertStringToByte = varTemp
End Function

Private Function ConvertByteToString(ByteToConvert)
	dim intI, strTemp
	For intI = 1 to LenB(ByteToConvert)
		strTemp = strTemp & chr(AscB(MidB(ByteToConvert,intI,1))) 
	Next
	ConvertByteToString = strTemp
End Function
		
End Class
%>