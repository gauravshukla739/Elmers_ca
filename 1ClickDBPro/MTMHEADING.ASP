<%
'1 Click DB copyright 1997-2001 David J. Kawliche

'All 1 Click DB technology is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'More info online at http://1ClickDB.com
'Email inquiries to info@1ClickDB.com
        
'**Start Encode**
%>
<!--#INCLUDE FILE=PageInit.asp-->
<html>
<head>
	<title><%=ocdBrandText%></title>
	<LINK rel=stylesheet type="text/css" href="<%=ocdStyleSheet%>">
</head>
<body topmargin="12"><%
call writetopmenu()
%></body>
</html>
<%
ocdTargetConn.Close
Set ocdTargetConn = Nothing
%>
