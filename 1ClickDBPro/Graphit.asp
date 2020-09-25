<%@ LANGUAGE = VBScript.Encode %>
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
<%
'on error resume next
Class ndBarChart


Public Values
Public Labels
Public Colors
Public MinValue
Public MaxValue
Public BarHeight
Public TotalValue
Public ValueScale
Public ChartWidth
Public UseImageLink
Public ChartType
Private Sub class_initialize
	ChartWidth = 400
	BarHeight = 15
	Colors = Array("red" , "green" , "blue")
	MinValue = 0
	MaxValue = 0
	TotalValue = 0
end sub
Public Function Display()

dim strTemp
dim intI
dim varValue
dim colorcount
dim intcountit
intcountit = 1
For each varValue in Values
	if not isnumeric(varValue) Then
		Display = "This graphing function creates a bar graph using the first field in your query as legend values and the second field in your query as data values.  The second field must contain only numeric values."
		exit function
	End if
	IF Clng(varValue) > CLNG(MaxValue) THEN
		MaxValue =varValue
	end if
	If MinValue = 0 THEN
		MinValue = varValue
	ElseIf CLng(varValue) < CLng(MinValue) THEN
		MinValue = varValue
	End if
	TotalValue = TotalValue + varValue
NEXT

if MaxValue <> 0 Then
	ValueScale =(ChartWidth) / MaxValue
Else
	ValueScale = 1
END IF
'Response.write chrttype
dim strTempLabels
dim strTempValues
dim strTempValues2
Select Case Trim(ChartType)
Case Cstr("3,"),"3"
strTemp = "<table >"
strTemp = strTemp & "<tr><td>"
strTemp = strTemp & "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1""><tr>"
colorcount = 1
For intI = 0 to UBOUND(Values)
	strTemp = strTemp & "<tr><td valign=""middle"" align=""left"">"
	strTemp = strTemp & Labels(intI) & "</font></td>"
	strTemp = strTemp & "<td valign=""middle"" nowrap align=""left"">"
	strTemp = strTemp & "<img src=""AppPixel" 
	if colorcount =  1 Then
		strTemp = strTemp &  "blue"
		colorcount = colorcount + 1
	Elseif colorcount = 2 Then
		strTemp = strTemp & "green"
		colorcount = colorcount + 1
	Else
		strTemp = strTemp & "red"
		colorcount = 1
	End if
	strTemp = strTemp &   ".gif"" width=""" & Values(intI) * ValueScale & """ height=""" & BarHeight & """> "
	strTemp = strTemp & Values(intI)
	strTemp = strTemp & " ("
	strTemp = strTemp & CLNG((Values(intI) / TotalValue) * 100) & "%"
	strTemp = strTemp & ")"
	strTemp = strTemp & "</td></tr>"
Next
strTemp = strTemp & "</table>"
strTemp = strTemp & "</tr></td></table>"
Case "5","5,"

strTemp = strTemp & "<P>Line Chart from .Net Web Server at <A HREF=http://chart.aylo.com target=_blank>http://chart.aylo.com</a><P>"
strTemp = strTemp & "<A HREF=http://chart.aylo.com target=_blank><IMG BORDER=0 SRC=https://chart.aylo.com/ImageGenerator.aspx?type=Line&width=640&height=480&legend="


strTempLabels =""
strTempvalues = ""
strTempValues2 = ""

For intI = 0 to UBOUND(Values)
	strTempValues = strTempValues  & Server.URLEncode(Cstr(values(intI))) & ","
	strTempLabels = strTempLabels & Server.URLEncode(Cstr(Labels(intI))) & ","
Next
if len(strTempValues) <> 0 Then
	strTempValues = Left(strTempValues,len(strTempValues) -1)
	strTempLabels = Left(strTempLabels , len(strTempLabels)-1)
End if

strTemp = strTemp & strTempLabels &  "&values1=" & strTempValues & "></A>"
Case Else
strTemp = strTemp & "<P>Pie Chart from .Net Web Server at <A HREF=http://chart.aylo.com target=_blank>http://chart.aylo.com</a><P>"
strTemp = strTemp & "<A HREF=http://chart.aylo.com target=_blank><IMG BORDER=0 SRC=https://chart.aylo.com/ImageGenerator.aspx?type=Pie&width=480&height=480&legend="


strTempLabels =""
strTempvalues = ""
For intI = 0 to UBOUND(Values)
	strTempValues = strTempValues  & Server.URLEncode(Cstr(values(intI))) & ","
	strTempLabels = strTempLabels & Server.URLEncode(Cstr(Labels(intI))) & ","
Next
if len(strTempValues) <> 0 Then
	strTempValues = Left(strTempValues,len(strTempValues) -1)
	strTempLabels = Left(strTempLabels , len(strTempLabels)-1)
End if

strTemp = strTemp & strTempLabels &  "&values1=" & strTempValues & "></A>"
strTemp = strTemp & "<P><SMALL>No image will show if Internet service is unavailable</SMALL><P>"

End Select

display = strTemp
End Function
End Class
%>
<%
dim connGraphIt
dim rs
dim categories
dim values
dim strGraphItSQL
dim strGraphItGridID
strGraphItGridID = request.querystring("GRIDID")
set connGraphIt = server.createobject("ADODB.Connection")
connGraphIt.open ndnscSQLConnect, ndnscSQLUser, ndnscSQLPass
dim databasetype
	Select Case UCASE(connGraphIt.Properties("DBMS Name")) 
		Case "MS SQL SERVER"
			DatabaseType = "SQLServer"
		Case "MS JET", "ACCESS"
			DatabaseType = "Access"
		Case Else
			DatabaseType = "SQLServer"
	End Select

strGraphItSQL = "Select " & Request.Querystring("SQLSelect_" & strGraphItGridID )
strGraphItSQL = strGraphItSQL & " From "  
strGraphItSQL = strGraphItSQL & NiceSQLIdentifier(Request.Querystring("SQLFrom_" & strGraphItGridID ))
if not Request.Querystring("SQLWhere_" & strGraphItGridID ) = "" Then
	strGraphItSQL = strGraphItSQL & " WHERE " & Request.Querystring("SQLWhere_" & strGraphItGridID )
End If
if not Request.Querystring("SQLGroupBy_"& strGraphItGridID ) = "" then 
	strGraphItSQL = strGraphItSQL & " GROUP BY " & Request.Querystring("SQLGroupBy_" & strGraphItGridID )
End if
if not Request.Querystring("SQLHaving_" & strGraphItGridID ) = "" then 
	strGraphItSQL = strGraphItSQL & " HAVING " & Request.QueryString("SQLHaving_" & strGraphItGridID )
End if
if not Request.Querystring("SQLOrderBy_" & strGraphItGridID ) = "" then 
	strGraphItSQL = strGraphItSQL & " ORDER BY " & Request.Querystring("SQLOrderBy_" & strGraphItGridID )
End if





set rs = server.createobject("ADODB.Recordset")
'response.write strGRaphItSQL
rs.open strGraphItSQL, connGraphIt

call writeheader("")
if err<>0 Then 
	call writefooter("")
End if
categories = ""
values = ""
if rs.eof then
	writefooter("No Data to Graph")
End if
While Not rs.EOF
	categories = categories & rs.Fields(0).Value & Chr(9)
    values = values & rs.Fields(1).Value & ";"
    rs.MoveNext
Wend
categories = Left(categories, Len(categories) - 1)
values = Left(values, Len(values) - 1)
dim tmparrcategories
dim tmparrvalues
tmparrcategories = split(categories,chr(9))
tmparrvalues = split(values,";")
dim objChart 
set objChart = new ndBarChart
objChart.values = tmparrvalues
objChart.labels = tmparrcategories
objChart.ChartType =  request.querystring("GraphType")
response.write objChart.display
connGraphIt.close
set connGraphIt = nothing
dim tmpQS
Response.write "<P><A HREF=""Browse.asp?graphtype=&"
for each tmpQS in request.querystring
	if UCASE(tmpQS) <> "GRAPHTYPE" Then
	response.write tmpQS
	Response.write "="
	Response.write server.urlencode(request.Querystring(tmpQS))
	Response.write "&"
	End if
next
Response.write """>Return to Browse</a><P>"

%>

<%call writefooter("")
function NiceSQLIdentifier (strSQLTableString)
	if DatabaseType = "Access" Then
		strSQLTableString = Replace(strSQLTableString,"]","")
	strSQLTableString = Replace(strSQLTableString,"[","")
		NiceSQLIdentifier = "[" & strSQLTableString & "]"
	ElseIf DatabaseType = "SQLServer" Then
		strSQLTableString = Replace(strSQLTableString,"""","")
		NiceSQLIdentifier = """" & strSQLTableString & """"
	Else
		NiceSQLIdentifier = strSQLTableString
	End if
end function

%>