<%
'1 Click DB ASP Library - Pop Up Zoom Text for Input
'copyright 1997-2004 David J. Kawliche

'1 Click DB ASP Library source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'Use of this software and/or source code is strictly at your own risk.
'All warranties are specifically disclaimed except as required by law.

'For more information see : http://1ClickDB.com

'**Start Encode**

option explicit
dim strTextField, strCallingForm
strTextField = Server.HTMLEncode(request.querystring("textfield"))
strCallingForm = Server.HTMLEncode(request.querystring("callingform"))

%>
<html>
	<head>
		<title>Edit Text</title>
		<script language="JavaScript">
			<!--
			function writebacktext() {
				window.opener.document.forms[<%=strCallingForm%>].elements['<%=strTextField%>'].value = document.forms[0].zoomtext.value;
				self.close();
				}
			function getzoomtext() { 
				document.forms[0].zoomtext.value = window.opener.document.forms[<%=strCallingForm%>].elements['<%=strTextField%>'].value;
				self.focus();
				}
			function stopError() {
				return true;
				}
			window.onerror = stopError
			// -->
		</script>
		<style>
			<!--

a { 
   	font-size : 10pt;
	font-family : Tahoma, Arial, sans-serif;
	color : #330066;
}
a:hover { 
   	font-size : 10pt;
	font-family : Tahoma, Arial, sans-serif;
	color : #990000;
} 
body {
	font-size : 10pt;
	font-family : Tahoma, Arial, sans-serif;
	scrollbar-base-color : #300066;
	scrollbar-face-color : #666690;
	scrollbar-shadow-color : Silver;
	scrollbar-highlight-color : Silver;
	scrollbar-3dlight-color : #ffffff;
	scrollbar-darkshadow-color : Silver;
	scrollbar-track-color : #CCCCCC;
	scrollbar-arrow-color : #ffffff;
	background : #FFFFFF;
	margin : 10px;
}
p {
	font-size : 10pt;
	font-family : Tahoma, Arial, sans-serif;
	
}
textarea {
   	font-size : 10pt;
	font-family : Tahoma, Arial, sans-serif;
}

			// -->
		</style>
	</head>
	<body onload="javascript:getzoomtext();">
		<center>
		<form>
		<textarea cols="74" rows="20" name="zoomtext"></textarea>
		</form>
		<a href="#" onclick="javascript:writebacktext()">Update</a>
		<a href="#" onclick="javascript:self.close()">Cancel</a>
		</center>
	</body>
</html>