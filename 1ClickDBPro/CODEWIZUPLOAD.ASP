<% ' Except for @ Directives, there should be no ASP or HTML codes above this line
' Setting LANGUAGE = VBScript.Encode enables mixing of encoded and unencoded ASP scripts

'1 Click DB Pro copyright 1997-2004 David J. Kawliche

'1 Click DB Pro source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com/

        
'**Start Encode**
On Error Resume Next
Response.Buffer = true
Dim objUpload, frmItem, objFSO, strUploadFileName, strErrorMessage, connTest
If Trim(Cstr(Request.ServerVariables("LOCAL_ADDR"))) <> "207.21.247.253" Then
	strErrorMessage = "Not Enabled"
End If
If strErrorMessage = "" Then
	Server.Scripttimeout = 600
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objUpload = Server.CreateObject("SoftArtisans.FileUp") 
	Randomize    ' Initialize random-number genera
	strErrorMessage = ""
	objUpload.MaxBytes =  2097152 '2Mb
	objUpload.MaxBytes =  528576 '512k
	strUploadFileName = Session.SessionID 
	If Not IsNumeric(strUploadFileName) Then
		strErrorMessage = "Invalid Name"
	Else
		strUploadFileName = strUploadFileName & ".mdb"
	End If
End If
If strErrorMessage = "" Then
	objUpload.Path = "d:\webs\accesshelp.net\data\upload\" 
	For Each frmItem in objUpload.form
		If Not IsObject(objUpload.form(frmItem)) Then
		Else
			objUpload.MaxBytes =  528576 '.5 Mb
			Select Case objUpload.Form(frmItem).TotalBytes 
				Case 0 
					strErrorMessage = "A Zero Length File Was Uploaded"		
				Case 528576
					strErrorMessage = "Uploaded File is Too Big.<P>  There is a 512k limit on uploads without prior arrangement.  Contact <A HREF=mailto:info@1ClickDB.com>info@1ClickDB.com</A> for details."
			End Select
			If strErrorMessage = "" Then
				objUpload.form(frmItem).SaveAs strUploadFileName
			Else
				Exit For
			End If
		End If
	Next
	Set objUpload = nothing
End If
If strErrMessage = "" Then
	If Err.Number <> 0 Then
		strErrorMessage = "Unspecified Error"
	End If
End If
If strErrorMessage = "" Then
	Session("ocdSQLConnect") = "provider=Microsoft.Jet.OLEDB.4.0;data source=d:\webs\accesshelp.net\data\upload\" &  strUploadFileName
	Set connTest= server.createobject("ADODB.Connection")
	connTest.Open "provider=Microsoft.Jet.OLEDB.4.0;data source=d:\webs\accesshelp.net\data\upload\" &  strUploadFileName
	If Err.Number <> 0 Then
		objFSO.DeleteFile "d:\webs\accesshelp.net\data\upload\" &  strUploadFileName
		Session.Abandon
		strErrorMessage = "Unrecognized Format"
	End if
End If
If strErrorMessage = "" Then
	If Err.Number <> 0 Then
		strErrorMessage = "Unspecified Error"
	End if
End If
If strErrorMessage = "" Then
	Session("ocdSQLUser") = ""
	Session("ocdSQLPass") = ""
	Response.clear
	Response.redirect "schema.asp"
End If
Response.Write "<html><head><title>Upload Error</title><link rel=""stylesheet"" type=""text/css"" href=""ocdStyleSheet.css""></head><body>"
Response.Write("<img src=""appWarningSmall.gif"" alt=""warning""> <span class=""warning"">")
Response.Write(strErrorMessage)
Response.write "</span><P>Use your Browser's BACK button to continue<P>"
Response.Write "</body></html>"
%>