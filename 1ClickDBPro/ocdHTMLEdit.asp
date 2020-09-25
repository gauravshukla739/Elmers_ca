<%
'1 Click DB ASP Library - Pop Up WYSIWYG HTML Editor 
'copyright 1997-2004 David J. Kawliche

'1 Click DB ASP Library source code is protected by international
'laws and treaties.  Never use, distribute, or redistribute
'any software and/or source code in violation of its licensing.

'Use of this software and/or source code is strictly at your own risk.
'All warranties are specifically disclaimed except as required by law.

'For more information see : http://1ClickDB.com

'**Start Encode**
%>
<HTML>
<HEAD>
<TITLE>HTML Editor</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
function writebacktext() { 
window.opener.document.forms[<%=request.querystring("callingform")%>].elements['<%=request.querystring("textfield")%>'].value =  cleanup(document.forms.HTMLEdit.elements['editbox'].documentHTML);
self.close();
}
function getzoomtext() { 
document.forms.HTMLEdit.elements['editbox'].documentHTML = window.opener.document.forms[<%=request.querystring("callingform")%>].elements['<%=request.querystring("textfield")%>'].value;
self.focus();
}
// -->
</SCRIPT>
<SCRIPT language=JavaScript>
<!--
// DHTML Editing Component Constants for JavaScript
// Copyright 1998 Microsoft Corporation.  All rights reserved.
//

//
// Command IDs
//
DECMD_BOLD =                      5000
DECMD_COPY =                      5002
DECMD_CUT =                       5003
DECMD_DELETE =                    5004
DECMD_DELETECELLS =               5005
DECMD_DELETECOLS =                5006
DECMD_DELETEROWS =                5007
DECMD_FINDTEXT =                  5008
DECMD_FONT =                      5009
DECMD_GETBACKCOLOR =              5010
DECMD_GETBLOCKFMT =               5011
DECMD_GETBLOCKFMTNAMES =          5012
DECMD_GETFONTNAME =               5013
DECMD_GETFONTSIZE =               5014
DECMD_GETFORECOLOR =              5015
DECMD_HYPERLINK =                 5016
DECMD_IMAGE =                     5017
DECMD_INDENT =                    5018
DECMD_INSERTCELL =                5019
DECMD_INSERTCOL =                 5020
DECMD_INSERTROW =                 5021
DECMD_INSERTTABLE =               5022
DECMD_ITALIC =                    5023
DECMD_JUSTIFYCENTER =             5024
DECMD_JUSTIFYLEFT =               5025
DECMD_JUSTIFYRIGHT =              5026
DECMD_LOCK_ELEMENT =              5027
DECMD_MAKE_ABSOLUTE =             5028
DECMD_MERGECELLS =                5029
DECMD_ORDERLIST =                 5030
DECMD_OUTDENT =                   5031
DECMD_PASTE =                     5032
DECMD_REDO =                      5033
DECMD_REMOVEFORMAT =              5034
DECMD_SELECTALL =                 5035
DECMD_SEND_BACKWARD =             5036
DECMD_BRING_FORWARD =             5037
DECMD_SEND_BELOW_TEXT =           5038
DECMD_BRING_ABOVE_TEXT =          5039
DECMD_SEND_TO_BACK =              5040
DECMD_BRING_TO_FRONT =            5041
DECMD_SETBACKCOLOR =              5042
DECMD_SETBLOCKFMT =               5043
DECMD_SETFONTNAME =               5044
DECMD_SETFONTSIZE =               5045
DECMD_SETFORECOLOR =              5046
DECMD_SPLITCELL =                 5047
DECMD_UNDERLINE =                 5048
DECMD_UNDO =                      5049
DECMD_UNLINK =                    5050
DECMD_UNORDERLIST =               5051
DECMD_PROPERTIES =                5052

//
// Enums
//

// OLECMDEXECOPT  
OLECMDEXECOPT_DODEFAULT =         0 
OLECMDEXECOPT_PROMPTUSER =        1
OLECMDEXECOPT_DONTPROMPTUSER =    2

// DHTMLEDITCMDF
DECMDF_NOTSUPPORTED =             0 
DECMDF_DISABLED =                 1 
DECMDF_ENABLED =                  3
DECMDF_LATCHED =                  7
DECMDF_NINCHED =                  11

// DHTMLEDITAPPEARANCE
DEAPPEARANCE_FLAT =               0
DEAPPEARANCE_3D =                 1 

// OLE_TRISTATE
OLE_TRISTATE_UNCHECKED =          0
OLE_TRISTATE_CHECKED =            1
OLE_TRISTATE_GRAY =               2
//-->
</SCRIPT>
<SCRIPT language=JavaScript>
<!--
function cleanup(htmlstuff) {
  tmpvar = htmlstuff
  var startbody = tmpvar.indexOf("<BODY>");
  var endbody = tmpvar.lastIndexOf("</BODY>");
  if(startbody > 0 && startbody < endbody) {
    tmpvar = tmpvar.substring(startbody + 6, endbody);
    //why the edit control likes to put 35 spaces in here is a mystery
	for(intI = 0; intI < 35; intI++){
      if (tmpvar.substring(intI, intI+1) == ' ') {
	  } else {
	    break;
      }
    }
    tmpvar = tmpvar.substring(intI);
  }
  if (tmpvar == "<P>&nbsp;</P>") {
    //with no other content, the edit control inserts the above markup
    tmpvar = ""; 
  }
  return tmpvar
}
function hedBOLD(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_BOLD,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedITALIC(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_ITALIC,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedUNDERLINE(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_UNDERLINE,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedFONT(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_FONT,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedCUT(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_CUT,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedCOPY(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_COPY,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedPASTE(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_PASTE,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedUNDO(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_UNDO,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedREDO(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_REDO,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}	
function hedIMAGE(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_IMAGE,OLECMDEXECOPT_PROMPTUSER);
  htmleditbox.focus();
}
function hedHYPERLINK(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_HYPERLINK,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedVISIBLEBORDERS(htmleditbox) {
  htmleditbox.ShowDetails = !htmleditbox.ShowDetails;
  htmleditbox.ShowBorders = !htmleditbox.ShowBorders;
  htmleditbox.focus();
}
function hedINSERTROW(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_INSERTROW,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedINSERTTABLE(htmleditbox) {
  var tabletxt = "<TABLE BORDER=1><TR><TD></TD></TR></TABLE>"
  varselection = htmleditbox.DOM.selection;
  if (varselection.type == "Control") {
    var selrange = varselection.createRange();
    selrange.item(0).insertAdjacentHTML("afterEnd", tabletxt);
    varselection.clear();
  } else {
    var selrange = varselection.createRange();  	
	selrange.pasteHTML(tabletxt);
  }
  htmleditbox.focus();
}
function hedJUSTIFYLEFT(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_JUSTIFYLEFT,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedJUSTIFYCENTER(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_JUSTIFYCENTER,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedJUSTIFYRIGHT(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_JUSTIFYRIGHT	,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedUNORDERLIST(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_UNORDERLIST,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedORDERLIST(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_ORDERLIST,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedOUTDENT(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_OUTDENT,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedINDENT(htmleditbox) {
  htmleditbox.ExecCommand(DECMD_INDENT,OLECMDEXECOPT_DODEFAULT);
  htmleditbox.focus();
}
function hedDELETECELL(htmleditbox) {
  if (htmleditbox.QueryStatus(DECMD_INSERTROW) != DECMDF_DISABLED) {
    htmleditbox.ExecCommand(DECMD_DELETECELLS,OLECMDEXECOPT_DODEFAULT);
  }
  htmleditbox.focus();
}
function hedDELETECOL(htmleditbox) {
  if (htmleditbox.QueryStatus(DECMD_INSERTROW) != DECMDF_DISABLED) {
    htmleditbox.ExecCommand(DECMD_DELETECOLS,OLECMDEXECOPT_DODEFAULT);
  }
  htmleditbox.focus();
}
function hedDELETEROW(htmleditbox) {
  if (htmleditbox.QueryStatus(DECMD_INSERTROW) != DECMDF_DISABLED) {
    htmleditbox.ExecCommand(DECMD_DELETEROWS,OLECMDEXECOPT_DODEFAULT);
  }
  htmleditbox.focus();
}
function hedINSERTCELL(htmleditbox) {
  if (htmleditbox.QueryStatus(DECMD_INSERTROW) != DECMDF_DISABLED) {
    htmleditbox.ExecCommand(DECMD_INSERTCELL,OLECMDEXECOPT_DODEFAULT);
  }
  htmleditbox.focus();
}
function hedINSERTCOL(htmleditbox) {
  if (htmleditbox.QueryStatus(DECMD_INSERTROW) != DECMDF_DISABLED) {
    htmleditbox.ExecCommand(DECMD_INSERTCOL,OLECMDEXECOPT_DODEFAULT);
  }
  htmleditbox.focus();
}
function hedINSERTROW(htmleditbox) {
  if (htmleditbox.QueryStatus(DECMD_INSERTROW) != DECMDF_DISABLED) {
    htmleditbox.ExecCommand(DECMD_INSERTROW,OLECMDEXECOPT_DODEFAULT);
  }
  htmleditbox.focus();
}
function hedMERGECELL(htmleditbox) {
  if (htmleditbox.QueryStatus(DECMD_INSERTROW) != DECMDF_DISABLED) {
    htmleditbox.ExecCommand(DECMD_MERGECELLS,OLECMDEXECOPT_DODEFAULT);
  }
  htmleditbox.focus();
}
function hedSPLITCELL(htmleditbox) {
  if (htmleditbox.QueryStatus(DECMD_INSERTROW) != DECMDF_DISABLED) {
	htmleditbox.ExecCommand(DECMD_SPLITCELL,OLECMDEXECOPT_DODEFAULT);
  }
  htmleditbox.focus();
}
//-->
</SCRIPT>
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
INPUT {
	color : #300066;
	background-color : silver;
	font-size : 8pt;
	font-family : Tahoma, Arial, sans-serif;
	font-weight : bold;
	padding-bottom : 0;
	padding-left : 0;
	padding-right : 0;
	padding-top : 0;
}
</STYLE>
</HEAD>
<BODY onLoad="javascript:getzoomtext();">
<CENTER>
<FORM name=HTMLEdit>

<INPUT TYPE=Button onclick="return hedCUT(document.forms.HTMLEdit.DHTMLEdit1);" name=btnCUT VALUE="Cut"><INPUT TYPE=Button onclick="return hedCOPY(document.forms.HTMLEdit.DHTMLEdit1);" name=btnCOPY VALUE="Copy"><INPUT TYPE=Button onclick="return hedPASTE(document.forms.HTMLEdit.DHTMLEdit1);" VALUE="Paste" name=btnPASTE> <INPUT TYPE=Button onclick="return hedUNDO(document.forms.HTMLEdit.DHTMLEdit1);" Value="Undo" name=btnUNDO><INPUT TYPE=Button onclick="return hedREDO(document.forms.HTMLEdit.DHTMLEdit1);" name=btnREDO VALUE="Redo"> <INPUT TYPE=Button onclick="return hedFONT(document.forms.HTMLEdit.DHTMLEdit1);" name=btnFONT VALUE="Font"><INPUT TYPE=Button onclick="return hedBOLD(document.forms.HTMLEdit.DHTMLEdit1);" name=btnBOLD VALUE="B"><INPUT TYPE=Button onclick="return hedITALIC(document.forms.HTMLEdit.DHTMLEdit1);" name=btnITALIC VALUE="I"><INPUT TYPE=Button onclick="return hedUNDERLINE(document.forms.HTMLEdit.DHTMLEdit1);" name=btnUNDERLINE VALUE="U"> <INPUT TYPE=Button onclick="return hedHYPERLINK(document.forms.HTMLEdit.DHTMLEdit1);" name=btnHYPERLINK VALUE="New Link"><BR><INPUT TYPE=Button onclick="return hedVISIBLEBORDERS(document.forms.HTMLEdit.DHTMLEdit1);" 
Value="Borders" name=btnVISIBLEBORDERS> <INPUT TYPE=Button onclick="return hedJUSTIFYLEFT(document.forms.HTMLEdit.DHTMLEdit1);" name=btnJUSTIFYLEFT Value="Left"><INPUT TYPE=Button onclick="return hedJUSTIFYCENTER(document.forms.HTMLEdit.DHTMLEdit1);" Value="Center" name=btnJUSTIFYCENTER><INPUT TYPE=Button onclick="return hedJUSTIFYRIGHT(document.forms.HTMLEdit.DHTMLEdit1);" Value="Right" name=btnJUSTIFYRIGHT> <INPUT TYPE=Button onclick="return hedOUTDENT(document.forms.HTMLEdit.DHTMLEdit1);" Value="Outdent" name=btnOUTDENT><INPUT TYPE=Button onclick="return hedINDENT(document.forms.HTMLEdit.DHTMLEdit1);" VALUE="Indent" name=btnINDENT> <INPUT TYPE=Button onclick="return hedORDERLIST(document.forms.HTMLEdit.DHTMLEdit1);" VALUE="OL" name=btnORDERLIST><INPUT TYPE=Button onclick="return hedUNORDERLIST(document.forms.HTMLEdit.DHTMLEdit1);" VALUE="UL" name=btnUNORDERLIST><BR><INPUT TYPE=Button onclick="return hedINSERTTABLE(document.forms.HTMLEdit.DHTMLEdit1);" VALUE="Table" name=ADDTABLE><INPUT TYPE=Button onclick="return hedINSERTROW(document.forms.HTMLEdit.DHTMLEdit1);" VALUE="Row" name=btnINSERTROW><INPUT TYPE=Button onclick="return hedINSERTCOL(document.forms.HTMLEdit.DHTMLEdit1);" VALUE="Col" name=btnINSERTCOL><INPUT TYPE=Button VALUE="Cell" onclick="return hedINSERTCELL(document.forms.HTMLEdit.DHTMLEdit1);" name=btnINSERTCELL><INPUT TYPE=Button onclick="return hedDELETEROW(document.forms.HTMLEdit.DHTMLEdit1);" VALUE="X Row" name=btnDELETEROW><INPUT TYPE=Button VALUE="X Col" onclick="return hedDELETECOL(document.forms.HTMLEdit.DHTMLEdit1);" name=btnDELETECOL><INPUT TYPE=Button VALUE="X Cell" onclick="return hedDELETECELL(document.forms.HTMLEdit.DHTMLEdit1);" name=btnDELETECELL><INPUT TYPE=Button VALUE="Merge" onclick="return hedMERGECELL(document.forms.HTMLEdit.DHTMLEdit1);" name=btnMERGECELLS><INPUT TYPE=Button VALUE="Split" onclick="return hedSPLITCELL(document.forms.HTMLEdit.DHTMLEdit1);" name=btnSPLITCELL><BR>
<P>
<OBJECT id=DHTMLEdit1 name=editbox classid=clsid:2D360201-FFF5-11D1-8D03-00A0C959BC0A width=450 height=250 VIEWASTEXT="YES" BORDER="0"></OBJECT>
<P>
<A HREF="" onClick="javascript:writebacktext();">Update</a>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<A HREF="" onClick="javascript:self.close();">Cancel</a>
</FORM>
</CENTER>
</BODY>
</HTML>
