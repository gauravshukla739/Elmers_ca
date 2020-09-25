<%
'1 Click DB ASP Library - Double List Box
'copyright 1997-2004 David J. Kawliche
'with contributions from : Aymeric Grassart

'1 Click DB ASP Library source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'Use of this software and/or source code is strictly at your own risk.
'All warranties are specifically disclaimed except as required by law.

'For more information see : http://1ClickDB.com

'**Start Encode**

Class DoubleBox

Public doublebox_size
Public doublebox_delimiters
Public doublebox_fields
Public doublebox_LHeader
Public doublebox_RHeader
Public doublebox_Header
Public doublebox_Vfields
Public doublebox_Sfields
Public doublebox_VSfields
Public doublebox_form
Public doublebox_name

Private Sub Class_Initialize
	doublebox_Sfields = "" 
	doublebox_Header = ""
	doublebox_LHeader ="" 
	doublebox_fields = ""
	doublebox_size = 6
	doublebox_delimiters =","
	doublebox_RHeader = ""
	doublebox_form = "forms[0]"
	doublebox_name="DoubleBoxForm"
end sub

Public Sub DrawDoublebOX

doublebox_Vfields = doublebox_fields
doublebox_VSfields = doublebox_Sfields
dim arrFields
dim arrVFields
dim intI
dim arrVSFields
dim arrSFields
arrFields = split(doublebox_fields,doublebox_delimiters)
arrVFields = split(doublebox_Vfields,doublebox_delimiters)
arrVSFields = split(doublebox_VSfields,doublebox_delimiters)
arrSFields = split(doublebox_Sfields,doublebox_delimiters)
'should check all arrays have same ubound
'should check doublebox_form is okay
'should check doublebox_name is okay
%>
<script language="JavaScript" TYPE="text/javascript">
<!--
function unselectSelectedFields<%=doublebox_name%>()
	{
	len_sf = document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options.length				
	for(idx=0; idx<len_sf; idx++)
	document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options[idx].selected = false    				
	}
function unselectFields<%=doublebox_name%>()
	{
	len_f = document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options.length				
	for(idx=0; idx<len_f; idx++)
	document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options[idx].selected = false    				
	}
function AddtoSelectedFields<%=doublebox_name%>()
	{
		idx = document.<%=doublebox_form%>.Fields<%=doublebox_name%>.selectedIndex
		if(idx < 0 )
			return;
			deleteSpaces<%=doublebox_name%>();
			var len_f = document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options.length
			var text,value, i
			for(i = 0; i < len_f; i++)
				if (document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options[i].selected == true){
					text = document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options[i].text;
					value = document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options[i].value;
					optionName<%=doublebox_name%> = new Option(text, value)
					len_sf = document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options.length				
					document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options[len_sf]=optionName<%=doublebox_name%>
				}
			for(i = 0; i < len_f; i++)
				if(document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options[i].selected == true) {
					document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options[i] = null
					i--
					len_f = document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options.length
				}
				len_f = document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options.length
				if	(len_f > 0)
					document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options[0].selected = true;
					AddtoHidden<%=doublebox_name%>()
	}
function deleteSpaces<%=doublebox_name%>()
	{
		var i;
		with(document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>)
		{
			if(document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options.length > 0) {
				for(i=0; i<document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options.length; i++)
					if(document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options[i].value == "Spaces")
						document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options[i] = null;
			}		
				
		}
		var x;
		with(document.<%=doublebox_form%>.Fields<%=doublebox_name%>)
		{
			if(document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options.length > 0) {
				for(x=0; x<document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options.length; x++)
					if(document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options[x].value == "Spaces")
						document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options[x] = null;
			}		
		}
	}
function AddAlltoSelectedFields<%=doublebox_name%>()	
	{
		selectAllFields<%=doublebox_name%>();
		AddtoSelectedFields<%=doublebox_name%>();
	}
function selectAllFields<%=doublebox_name%>()
	{
		var i;
		var size = document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options.length
		for(i = 0 ; i < size; i++)
			document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options[i].selected = true;

	}
function AddtoHidden<%=doublebox_name%>()
	{
		document.<%=doublebox_form%>.<%=doublebox_name%>.value = ''
		size = document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options.length
		for(i = 0 ; i < size; i++) {
			 document.<%=doublebox_form%>.<%=doublebox_name%>.value = document.<%=doublebox_form%>.<%=doublebox_name%>.value + '<%=doublebox_delimiters%>'  +	document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options[i].value ;
		}
		newsize<%=doublebox_name%> = document.<%=doublebox_form%>.<%=doublebox_name%>.value.length
		newvalue<%=doublebox_name%> = document.<%=doublebox_form%>.<%=doublebox_name%>.value
		document.<%=doublebox_form%>.<%=doublebox_name%>.value = newvalue<%=doublebox_name%>.substring(1,newsize<%=doublebox_name%>)
	}
function RemovefromSelectedFields<%=doublebox_name%>()
	{
		deleteSpaces<%=doublebox_name%>()
		var len_f;
		var idx
		idx = document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.selectedIndex
		if(idx < 0 )
			return;
		var len_sf = document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options.length
		var text,value, i
		for(i = 0; i < len_sf; i++)
			if(document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options[i].selected == true) {
				text = document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options[i].text;
				value = document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options[i].value;
				optionName<%=doublebox_name%> = new Option(text, value)
				len_f = document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options.length				
				document.<%=doublebox_form%>.Fields<%=doublebox_name%>.options[len_f]=optionName<%=doublebox_name%>
			}
		for(i = 0; i < len_sf; i++)
			if(document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options[i].selected == true) {
				document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options[i] = null
				i--
				len_sf = document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options.length
			}
			len_sf = document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options.length
			if	(len_sf > 0)
				document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options[0].selected = true;
				AddtoHidden<%=doublebox_name%>()
	}
function RemoveAllfromSelectedFields<%=doublebox_name%>()
	{
		selectAllSelectedFields<%=doublebox_name%>();
		RemovefromSelectedFields<%=doublebox_name%>();
	}
function selectAllSelectedFields<%=doublebox_name%>()
	{
		var i;
		var size = document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options.length
		for(i = 0 ; i < size; i++)
			document.<%=doublebox_form%>.SelectedFields<%=doublebox_name%>.options[i].selected = true;
	}
//-->
</script>
<table border="0" cellspacing="3" cellpadding="2">
<%if trim(doublebox_Header) <>"" Then %>
<tr><td colspan=3 align=center><%=doublebox_header%></td></tr>
<%End if%>
<% if trim(doublebox_lHeader) <> "" or trim(doublebox_rHeader) <> "" Then%>
<tr><td align=center><%=doublebox_lheader%></td><td></td><td align=center><%=doublebox_rHeader%></td></tr>
<%end if%>
<tr>
<td align=center>
<select name="Fields<%=doublebox_name%>" size="<%=doublebox_size%>" onFocus="unselectSelectedFields<%=doublebox_name%>()" multiple Style="width:150">
<%
for intI = 0 to UBound(arrFields)
%>
<option value="<%=Server.HTMLEncode(arrVFields(intI))%>">
<%=arrFields(intI)%></OPTION>
<%next%>
<option value="Spaces">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</select>
</td>
<td ALIGN=CENTER>
<input TYPE="BUTTON" VALUE=" Select &gt; " onClick="javascript:AddtoSelectedFields<%=doublebox_name%>()" Class=Submit><br>
<input TYPE="BUTTON" VALUE="Select All" Class=Submit onClick="javascript:AddAlltoSelectedFields<%=doublebox_name%>()" ><P>
<BR>
<input TYPE="BUTTON" VALUE="Hide All" Class=Submit onClick="javascript:RemoveAllfromSelectedFields<%=doublebox_name%>()"><br>
<input TYPE="BUTTON" VALUE=" &lt; Hide" Class=Submit onClick="javascript:RemovefromSelectedFields<%=doublebox_name%>()">
</td>
<td align=center>
<select name="SelectedFields<%=doublebox_name%>" size="<%=doublebox_size%>" onFocus="unselectFields<%=doublebox_name%>()" multiple Style="width:150">
<%
for intI = 0 to UBound(arrSFields)
%>
<option value="<%=Server.HTMLEncode(arrVSFields(intI))%>"><%=arrSFields(intI)%>")
<%next%>
<option value="Spaces">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</select>
<input type="hidden" name="<%=doublebox_name%>" value="<%=Server.HTMLEncode(doublebox_VSfields)%>">
</td></tr>
</table>
<%
end sub

end class
%>
