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
<html>
<head>
<title><%=ocdBrandText%></title>
<LINK rel=stylesheet type="text/css" href="<%=ocdStyleSheet%>">
</head>
<body >
<p>Your <%=ocdBrandText%> session has expired.</p>

<p><a href="Connect.asp" target="_parent">Click here to continue.</a></p>

</body>
</html>

