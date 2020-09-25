<%
'1 Click DB ASP Library - Pop Up Calendar for Input
'copyright 1997-2004 David J. Kawliche

'1 Click DB ASP Library source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'Use of this software and/or source code is strictly at your own risk.
'All warranties are specifically disclaimed except as required by law.

'For more information see : http://1ClickDB.com

'**Start Encode**

Option explicit
response.buffer=true
Dim strCallingForm, strDateField, strInitialDate, strInitialMonth, strInitialYear, datControl, intMonth, intMonthNext, intMonthPrev, intCount, intWeekday, intYearPrev, intYearNext
'If request.querystring("InitialMonth") = "" Then
strInitialDate = Request.QueryString("InitialDate")
'Else
'strInitialDate = ""
'End if
strCallingForm =  Request.QueryString("CallingForm")
strDateField = Request.QueryString("DateField") 
Response.write ("<html><head><title>Choose Date</title>")
Response.write (vbCRLF)
Response.write ("<script language=""JavaScript"">" & vbCRLF)
REsponse.write ("  <!--" & vbCRLF)
Response.write ("function writebackdate(selecteddate) { " & vbCRLF)

Response.write ("var tmp = selecteddate;" & vbCRLF)
Response.write ("if ('12:00:00 AM' != document.forms[0].elements['dTime'].value) {" & vbCRLF)
Response.write ("tmp = tmp + ' ' + document.forms[0].elements['dTime'].value;" & vbCRLF)
Response.write ("}") & vbCRLF
Response.write ("window.opener.document.forms[")
Response.write (strCallingForm & "].elements['" & strDateField)
Response.write ("'].value = tmp;" & vbCRLF)
Response.write ("self.close();" & vbCRLF)
Response.write ("}" & vbCRLF)
Response.write ("function acwstopError() {" & vbCRLF)
Response.write ("return true;" & vbCRLF)
Response.write ("}" & vbCRLF)
Response.write ("window.onError = acwstopError();" & vbCRLF)
Response.write ("  // -->" & vbCRLF)
Response.write ("</script>" & vbCRLF)
'Response.write ("<LINK rel=stylesheet type=""text/css"" href=""acwc.css"">" & vbCRLF )
%>
<STYLE>
A { 
   	font-size : 10pt;
	font-family : Tahoma, Arial, sans-serif;
	color : #330066;
}
A:hover { 
   	font-size : 10pt;
	font-family : Tahoma, Arial, sans-serif;
	color : #990000;
} 
A.menu { 
   	font-size : 10pt;
	font-family : Tahoma, Arial, sans-serif;
	color : #330066;
}
A.menu:hover { 
   	font-size : 10pt;
	font-family : Tahoma, Arial, sans-serif;
	color : #330000;
	background : #FFD700;
} 
A.menu:visited { 
   	font-size : 10pt;
	font-family : Tahoma, Arial, sans-serif;
	color : #330066;	
}
BODY {
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
P {
	font-size : 10pt;
	font-family : Tahoma, Arial, sans-serif;
}
TD {
   	font-size : 10pt;
	font-family : Tahoma, Arial, sans-serif;
}
TH {
   	font-size : 10pt;
	font-family : Tahoma, Arial, sans-serif;
}
</STYLE>
<%
Response.write ("</html><body onload=""javascript:self.focus();"">")
If IsDate(strInitialDate) Then 
	datControl = strInitialDate
	strInitialYear = year(CDate(datControl))
	strInitialMonth = cstr(Month(CDate(datControl)))
Else 
	If Request.Querystring("InitialMonth") = "" Then 
		strInitialMonth = Month(Now) 
	Else 
		strInitialMonth = Request.Querystring("InitialMonth") 
	End If 
	If Request.QueryString("initialYear") <> "" Then 
		datControl = strInitialMonth & "/" & "1" & "/" & Request.QueryString("InitialYear") 
		strInitialYear = request.querystring("initialyear")
	Else 
		datControl = strInitialMonth & "/" & "1" & "/" & Year(Now)
		strInitialYear = year(now)
	End If 
	strInitialDate = date
End If 
'strInitialDate = datControl 
intMonth = Month(datControl) 
Response.write("<center> <form method=""post"" action=""")
response.write(request.servervariables("SCRIPT_NAME"))
Response.write("?callingform=")
Response.write(Server.URLEncode(request.querystring("callingform")))
Response.write("&amp;DateField=")
Response.write(Server.URLEncode(request.querystring("datefield")))
Response.write("&amp;InitialDate=")
Response.write(Server.URLEncode(request.querystring("InitialDate")))

Response.write(""">")
Response.write ("<table Border=1> ")

Response.Write("<tr><td colspan=7 align=center cellpadding=10>")
Response.write(" <FONT SIZE=2>")
If intMonth < 12 then 
	intMonthNext = intMonth + 1 & "&InitialYear=" & Year(datControl) 
Else 
	intMonthNext = "1&InitialYear=" & Year(DateAdd("yyyy", 1, datControl)) 
End if 
If intMonth > 1 then 
	intMonthPrev = intMonth - 1 & "&InitialYear=" & Year(datControl) 
Else 
	intMonthPrev = "12&InitialYear=" & Year(DateAdd ("yyyy", -1, datControl)) 
End If 
Response.write ("<a href=""")
Response.write (request.servervariables("SCRIPT_NAME"))
Response.write ("?CallingForm=")
Response.write (server.urlencode(request.querystring("CallingForm")))
Response.write ("&DateField=")
Response.write (server.urlencode(request.querystring("DateField")))
Response.write ("&InitialMonth=")
Response.write (intMonthPrev)
Response.write (""">&lt;-</a></font> <B><FONT SIZE=4>")
Response.write (MonthName(intMonth))
REsponse.write ("</font></b> <FONT SIZE=2>")
Response.write ("<a href=""")
Response.write (request.servervariables("SCRIPT_NAME"))
Response.write ("?CallingForm=")
Response.write (server.urlencode(request.querystring("CallingForm")))
Response.write ("&DateField=")
Response.write (server.urlencode(request.querystring("DateField")))
Response.write ("&InitialMonth=")
Response.write (intMonthNext)
Response.write (""">-&gt;</a></FONT>")
Response.write ("<br>")
intyearNext = intMonth & "&InitialYear=" & Year(DateAdd ("yyyy", 1, datControl))
intyearPrev = intMonth & "&InitialYear=" & Year(DateAdd ("yyyy", -1, datControl))
Response.write ("<FONT SIZE=2><a href=""")
Response.write (request.servervariables("SCRIPT_NAME"))
Response.write ("?CallingForm=")
Response.write (server.urlencode(request.querystring("CallingForm")))
Response.write ("&DateField=")
Response.write (server.urlencode(request.querystring("DateField")))
Response.write ("&InitialMonth=")
Response.write (intYearPrev)
Response.write (""">&lt;-</a></font> <B><FONT SIZE=4>")
Response.write (Year(datControl))
Response.write ("</FONT></b> <FONT SIZE=2>")
Response.write ("<a href=""")
Response.write (request.servervariables("SCRIPT_NAME"))
Response.write ("?CallingForm=")
Response.write (server.urlencode(request.querystring("CallingForm")))
Response.write ("&DateField=")
Response.write (server.urlencode(request.querystring("DateField")))
Response.write ("&InitialMonth=")
Response.write (intYearNext)
Response.write (""">-&gt;</a></font>")
Response.write ("</TD></TR>")
Response.write ("<TR><TD COLSPAN=7 NOWRAP BGCOLOR=YELLOW><font size=2><b>Time : </b></FONT><INPUT SIZE=12 MaxLength=24 NAME=dTime ID=dTime VALUE=""" )
If request.form("dTime") <> "" Then
	Response.write (Server.HTMLEncode(request.form("dTime")))
Else
	Response.write (FormatDateTime((CDate(strInitialDate)),3))
End If
Response.write(""">")
Response.Write("</TD></TR>")
Response.write ("<TR><TD ALIGN=CENTER><font size=2><b>Su</b></FONT></TD><TD ALIGN=CENTER><font size=2><b>Mo</b></FONT></TD><TD ALIGN=CENTER><font size=2><b>Tu</b></FONT></TD><TD ALIGN=CENTER><font size=2><b>We</b></FONT></TD><TD ALIGN=CENTER><font size=2><b>Th</b></FONT></TD><TD ALIGN=CENTER><font size=2><b>Fr</b></FONT></TD><TD ALIGN=CENTER><font size=2><b>Sa</b></FONT></TD></TR>")
Response.write ("<TR>")
datControl =  CDATE(Cstr(strInitialMonth & "/" & "1" & "/" & strInitialYear ))
intWeekday = Weekday(datControl) 
For intCount = 1 to intWeekday - 1 
	response.write ("<TD><FONT SIZE=2>&nbsp;</FONT></TD>")
Next 
Do Until intMonth <> Month(datControl) 
	While intWeekday <> 8 
		Response.write ("<TD ALIGN=RIGHT VALIGN=TOP HEIGHT=20 WIDTH=20 ")
		If CDATE(datControl) = CDATE((FormatDateTime(CDate(strInitialDate),2))) Then 
			Response.Write ("BGCOLOR=""YELLOW""")
		ElseIf CDATE(datControl) = Date() Then 
			Response.Write ("BGCOLOR=""#999999""")
		
		End if
		Response.write (" >")
'		Response.write (FormatDateTime(CDate(strInitialDate),2))
		Response.write ("<FONT SIZE=2>")
		Response.write ("<A HREF="""" onClick=""")
		Response.Write ("writebackdate('" & Month(datControl) & "/" & Day(datControl) & "/" & Year(datControl)) 
		Response.write ("')"">")
		Response.write (Day(datControl))
		Response.write ("</A>")
		REsponse.write ("</FONT>")
		Response.write ("</TD>")
		intWeekday = intWeekday + 1 
		datControl=DateAdd("d", 1, datControl) 
		If intMonth <> Month(datControl) then 
			intWeekday = 8 
		End If 
	Wend 
	intWeekday = 1 
	Response.write ("</TR><TR> ")
Loop 
Response.write ("</TR>")

REsponse.write ("</TABLE><BR>")
Response.write ("</FORM></CENTER></BODY></HTML>")
%>
