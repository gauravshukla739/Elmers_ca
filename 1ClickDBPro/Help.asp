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

call writeheader("")
dim blnWalkThru

if request.querystring("Walkthru") <> "" Then
	blnWalkThru = True
Else
	blnWalkThru = False
End if
%>
 <IMG SRC="AppHelp.gif" ALT="Help"> <SPAN CLASS=Information><%
If blnWalkThru Then
%>1 Click DB Walk-Thru Tour<%
else
%>
1 Click DB Online Help

<%
end if
%></span>
<p>Visit <A HREF="http://1ClickDB.com/support/" target="_top">http://1ClickDB.com/support</a> or contact <A HREF="http://AccessHelp.net" target="_top">AccessHelp.net</a> for prompt attention and technical support.
<P><A NAME=Contents></A>
<B>Contents</B>
<OL>
<LI><IMG SRC=appConnect.gif ALT=Connect> <A HREF="#SQLConnect">Connect</A></LI>
<LI><IMG SRC=AppDB.Gif ALT=DB> <A HREF="#DBProperties">Database</A></LI>
<LI><IMG SRC=appCOmmand.gif ALT=Command> <A HREF="#SQLCommander">Command</A></LI>
<LI><IMG SRC=appSelect.gif ALT=Select> <A HREF="#SQLSelector">Select</A></LI>
<LI><IMG SRC=appTable.gif ALT=Browse> <A HREF="#Browse">Browse</A></LI>
<LI><IMG SRC=appSearch.gif ALT=Search> <A HREF="#SearchPages">Search</A></LI>
<LI><IMG SRC=Menu_Link_Default.gif ALT=Properties> <A HREF="#StructurePages">Properties</A></LI>
<LI><IMG SRC=AppNew.gif ALT=Add><IMG SRC=GRIDLNKEDIT.GIF ALT=Edit>
<IMG SRC=GRIDLNKDELETE.GIF ALT=Delete> <A HREF="#EditPages">Add/Edit/Delete</A></LI>
<LI><IMG SRC=appWizard.gif ALT=Wizard> <A HREF="#CodeWizard">Code Wizard</A></LI>
<LI><IMG SRC=appAudit.gif ALT=Audit> <A HREF="#AuditWizard">Audit Wizard</A></LI>
</OL>
<%If not blnWalkThru Then%>
<P><A HREF=<%=ocdPageName%>?walkthru=true class=menu>** Click Here to Activate 1 Click DB Walk-Thru Tour **</A></P>
<%end if%>
<P><B>Getting Started</B></P>After unzipping 1 Click DB .asp files into a folder on your web server (e.g. "1ClickDB"), start 1 Click DB by pointing your web browser to http://<I>myservername</I>/<I>1ClickDB</I>/Default.asp</P><P>Some tests are run on this page to check your web server's compatibility with 1 Click DB.  If there is a problem an error message is displayed, otherwise 1 Click DB automatically forwards you to the dynamic <A CLASS=Menu><IMG SRC=appConnect.gif Border=0> Connect</A> screen. <P>From this screen you may enter an ADO connect string directly on this screen or click on one of the listed Databases for a 1 Click DB Connection Wizard</P>
<HR>
<OL>
<LI>
<P><IMG SRC=appCOnnect.gif ALT=Connect>
<A NAME="#SQLConnect">Connect</A> </P>
<%If blnWalkThru Then%>

<P><IMG SRC=Screen_Connect.gif></P>
<%end if%>
<P>
Use the Connect screen to constructs the specially formatted information needed to link your web server to a database.
Click on a supported database type for a connection wizards or input the ADO information directly into the Connect String textbox.
The Connect String textbox can also be used to specify an ODBC Data Source Name ("DSN") or named shortcut to your database. These are typically created on your web server by an administrator using the Windows Control Panel interface.
</P>
<P>
<B>Having problems connecting?</B> Review the 1 Click DB Knowledge Base on <A HREF="http://1ClickDB.com/content/view.aspx?_@id=53416" target=_top>Resolving Common ADO Errors</A>.</P>

<P>
By default, connection information, including the user name and password are indirectly saved in special text files called session cookies.  These cookies do not store the actual information you are using in your browser or client computer, but a long cryptic name that refers to a set of variables maintained on the web server. These files are only active while you are logged on to the web site and expire if no activity is recorded for a timeout determined by your web server.  Administrators can also hard coding connection string information in the Config.asp configuration file.

<P>
If your web browser and web server support it, switch between Secure Socket Layer ("SSL") encryption and plain text communications using the "Enable SSL" and "Disable SSL" links. These links work toggling the URL of the current page from http:// to https:// If either your web server or your web browser does not support SSL encryption, attempts to display pages starting with https:// will fail.
<P>
Any database using an internal IP address (usually of the form 10.x.x.x or 192.x.x.x) or a Windows share name must be on the same local area network as the web server to connect.
<%if blnWalkThru Then %>
<P><IMG SRC=Screen_ConnectSQL.gif></P>
<%end if%>
<P>
1 Click DB uses the excellent free tree menu framework from http://treemenu.com.  This menu system supports most popular browsers, but will not be active when any of the 1 Click DB compatibility options are selected (e.g. "No Frames" or "No JavaScript".)  If the menu works for a user on some databases but not others, odds are there are special characters in the name of at least one of the database's table or view objects that have not been anticipated by 1 Click DB.  Please contact 1 Click DB support with a list of your database table and query view names for a quick remediation.  Since the menu relies on a frame system, bookmarks will generally return the user to the initial DB Properties screen.  The exception to this is when you use Save button present on a 1 Click DB Browse grid.  A "Restore Frameset" link is present at the bottom of all content pages.  This is useful when returning to a bookmarked Browse grid or any other time when the content frame has been "unframed" in the user's browser.
<P>

<A HREF="#Contents">Back To Contents</A>
<P>

</LI>
<LI><IMG SRC=appDB.gif ALT=DB> <A NAME="#DBProperties">Database</A>
<P>
The main database page displays detailed information about the Active Data Object provider used to connect to your database.  Different providers will return different information.  MS SQL Server users will also see links to perform common database maintenance operations.  Activate this screen by clicking on the "DB Properties" link or the database name and icon at the top of the tree menu.
<P>
<UL><LI><A Name=ADOSchemas>ADO Schemas</A><P>

All ADO connections can be queried for information about the structure ("schema") of your database.  The ADO Schemas screen provides a convenient way to execute these schema queries on your database.  Although there is a standard set of queries that can be performed, databases are not required to actually supply this information in response to a query.  Different databases will make available different information from this interface.  The exact version and type (OLEDB or ODBC) of database driver used also affect what information is available.  Because of this you may get the message "Schema not supported by this ADO Provider" when requesting this information.  This is normal behavior. 
<P>
</LI>
<LI><A Name=CGIVariables>CGI Variables</A><P>
Every web page contains information about the environment in which it is running.  These CGI Variables (also known as ServerVariables in ASP) include required information like the web server and page name, as well as optional information such as the browser being used. The CGI Variables page exposes all of this information in a list and also identifies the Server Type and Application Scripting environment.</LI>
</UL>
<P>
<A HREF="#Contents">Back To Contents</A>
<P>
</LI>
<LI><IMG SRC=appCommand.gif ALT=Command> <A Name=SQLCommander>Command</A><P>
<%if blnWalkThru Then %>
<P><IMG SRC=Screen_Command.gif></P>
<%end if%>
The Command page allows the user to submit and view the results of arbitrary SQL DDL and DML commands.  This can be used to test the syntax of your SQL or to access functionality not available through the graphic interface of 1 Click DB.  For SQL Server the Command page also features support for multiple recordsets and return values, full support for database messages generated by DBCC commands and stored procedures, as well as the ability to view your query execution path via SHOW PLAN.
<P>
<%if blnWalkThru Then %>
<%end if%>

<%if blnWalkThru Then %>
<P><IMG SRC=Screen_CommandShowPlan.gif></P>
<%end if%>
<A HREF="#Contents">Back To Contents</A>
<P></LI>
<LI><IMG SRC=appSelect.gif ALT=Select> <A Name=SQLSelector>Select</A><P>
<%if blnWalkThru Then %>
<P><IMG SRC=Screen_Select.gif></P>
<%end if%>

The Select screen allow the user to construct custom select queries for display in a full functioned 1 Click DB Browse grid.  SQL Select statements must be entered in their constituent parts on this page.  More specialized SQL statements and stored procedures can be run from the Command page.
<P>

<%if blnWalkThru Then %>
<P><IMG SRC=Screen_SelectField.gif></P>
<%end if%>

<A HREF="#Contents">Back To Contents</A>
<P></LI>
<LI><IMG SRC=appTable.gif ALT=Browse> <A Name=Browse>Browse</A><P>
<%if blnWalkThru Then %>
<P><IMG SRC=Screen_Browse.gif></P>
<%end if%>

<UL>

<LI> <B>Bookmark Your Browsing</b> 
<BR>
Since all the filter, sort, position and other user definable parameters are set in a browse page's querystring, you can effectively "save" your view by bookmarking it.  When using a frameset be sure to bookmark the frame containing the browse page.
<P>
</LI>

<LI>  <B>Drill Down</b> <IMG SRC="GRIDSMBTNDRILLDOWN.GIF" ALT="Drill Down" BORDER="0"><BR>
This button toggles drill down on and off for a grid.  When drill down is on the values for most fields turn into hyperlinks.  Clicking on these hyperlinks sets a filter on the grid for all other fields that match the current value.  If there is only one record in the view this button has no effect.
<P>

</LI>

<LI>  <B>First Prev Next Last</b> <IMG SRC="GRIDSMBTNFIRST.GIF" ALT="First" BORDER="0"> <IMG SRC="GRIDSMBTNPREV.GIF" ALT="Previous" BORDER="0"> <IMG SRC="GRIDSMBTNNEXT.GIF" BORDER="0" ALT="Next"> <IMG SRC="GRIDSMBTNLAST.GIF" BORDER="0" ALT="Last"><BR>
These controls let the user navigate through multiple pages on a grid.  From left to right the buttons will move to the first page, move to the previous page, move to the next page and move to the last page of a results set.  If you are on the first page of a results set, the move first and move previous buttons will display, but will not have active links.  If you are on the last page of a results set, or if there is only one page the move next and move last buttons will display but will not have active links. 
<P>
</LI>
<LI> <B>New</b> <IMG SRC="GRIDSMBTNNEW.GIF" ALT="New" BORDER="0"> <BR>
This button will take the user to the edit page defined for the recordset at the new record position.<P>
</LI>
<LI> <B>Print</b> <IMG SRC="GRIDSMBTNPRINT.GIF"  BORDER="0" ALT="Print"> <BR>
The print function allows one click access to your browser's print functionality.  A new window will open with ALL the data from the query displayed (no paging) and plain formatting (this may be a large file).  For Internet Explorer users the print dialog will also appear after the page has been fully loaded.
<P>
</LI>
<LI> <B>Export</b> <IMG SRC="GRIDSMBTNEXCEL.GIF" BORDER="0" ALT="Export to Excel"> <IMG SRC="GRIDSMBTNXML.GIF" BORDER="0" ALT="Export to XML"> <BR>
These buttons will stream the entire recordset as it is show on the screen as different export formats.  Defined filter and sort criteria are retained, however paging is ignored.  The entire recordset is always set for export.  For this reason, these buttons should be used with caution when dealing with large recordsets.  If the user attempts to export too much data at once, the transfer may fail.  Note that Excel exports are really HTML streams with the MIME type set for the Microsoft spreadsheet.
<P>
</LI>
<LI> <B>Sort Links</b> 
<IMG SRC="GRIDLNKASC.GIF" BORDER="0" ALT="Sort Ascending" WIDTH="11" HEIGHT="11"> <IMG SRC="GRIDLNKDESC.GIF" BORDER="0" ALT="Sort Descending" WIDTH="11" HEIGHT="11"> <BR>
These two buttons appear beneath the every field name that is sortable.  Binary and memo fields are not sortable and so these will not appear.
<P>
</LI>
<LI> <B>Filter Links</b> <IMG SRC="GRIDLNKFILTER.GIF" BORDER="0" ALT="Filter on This Field" WIDTH="11" HEIGHT="11"><BR>
These buttons appear below the names of every field except binary ones.
Clicking on them will bring up a filter screen.<P>
</LI>
<LI><B>Edit Links</b> <IMG SRC="GRIDLNKEDIT.GIF" border="0" HEIGHT="12" WIDTH="12" ALT="Edit Record"> <BR>
Clicking this button will take the user to the edit page for the current record.
<P>
</LI>

</ul>
<P>
<A HREF="#Contents">Back To Contents</A>
<P></LI>
<LI><IMG SRC=appSearch.gif ALT=Search> <A Name=SearchPages>Search</A><P>
<%if blnWalkThru Then %>
<P><IMG SRC=Screen_BrowseSearch.gif></P>
<%end if%>

Each table and view can be searched using the Query By Form functionality of the Search Pages. Keyword searches support boolean expressions like "and" "or" and "not". To search for all records containing either the word John or the word Smith anywhere in any field use the expression John Smith. To search for all records containing both the word John and the word Smith anywhere in any field, use the expression John AND Smith. To search for all records containing the exact phrase John Smith in any field, use the expression "John Smith". 
<p>
For Access and SQL Server, keyword searches for numeric expressions will examine all non-binary fields in the database.  For other databases and for all alphabetic criteria, only text fields are keyword searched.Page size and sort order can also be specified on this screen.  The check boxes next to each field name determine whether or not a field will be displayed when browsing results.
<P>
Search criteria are passed using query strings.  If your query string is too long, your browser may not display the results page.  This is most often a problem when doing keyword searches using numeric expressions on tables with a large number of fields.
<P>
<A HREF="#Contents">Back To Contents</A>
<P></LI>
<LI><IMG SRC=Menu_Link_Default.gif ALT=Properties> <A Name=StructurePages>Properties</A><P>
<%if blnWalkThru Then %>
<P><IMG SRC=Screen_Properties.gif></P>
<%end if%>

Property pages detail the various properties of your database objects.  Depending on the database and object type, different information will be returned.  For Access and SQL Server OLEDB connections these screens also contain a graphic interface to make changes to these objects.
<P>
<B>Related Tables</B><BR>
For Microsoft Access and SQL Server OLEDB connections, the structure page contains The Related to Table function.  This allows a user to create
referential links between fields in different tables.  From
the structure page for a table containing a field you wish to link, choose
the related table with field you wish to link to,
choose 'many to one' or 'one to many' relationship. The
next screen should show you the primary key fields on
the 'one' table.  Match these to the fields you wish to
link with on the 'many' table and you are done. Matching fields must be of the same data type and an Access autonumber or SQL Identity field may not be used in the key for the 'many' table. Currently
1 Click DB only supports creating links from primary key
fields.  This is not actually required by Access or SQL
Server databases, but is generally a good practice.  If you
need to link tables without using the primary key on
the 'one' side, you should generally reconsider your design.  Choosing Cascade Updates will automatically change the related fields in the 'many' table when the matching field in the 'one' table is updated.  Although there are techniques to automatically add records to the many table when a record is added to the 'one' table, creating a relationship between the two tables does not do this automatically. Choosing Cascade Deletes will automatically remove all related child records in the 'many' table when their parent record in the 'one' table is deleted.  
<P>
<A HREF="#Contents">Back To Contents</A>
<P></LI>
<LI><IMG SRC=appNew.gif ALT=Add> <IMG SRC=GRIDLNKEDIT.GIF ALT=Edit> <IMG SRC=GRIDLNKDELETE.GIF ALT=Delete> <A Name=EditPages>Add/Edit/Delete</A><P>
<%if blnWalkThru then%>

<P><IMG SRC=Screen_EditRelated.gif></P>
<%end if%>
Add/Edit/Delete pages enable the maintenance of database information via any HTML web browser.  When supported by the database provider, required fields are marked with a red asterisk, default values will be shown, and single field foreign keys displayed as drop down lists.  Microsoft Access and SQL Server databases using an OLEDB provider also display a master/detail view of all related records.  Most databases and ADO providers do not fully support updating Views.  
<%if blnWalkThru then%>
<P><IMG SRC=Screen_Edit.gif></P>
<%end if%>
<P>
<A HREF="#Contents">Back To Contents</A>
</P><P>&nbsp;</P></LI>

<LI><IMG SRC=appWizard.gif ALT=Wizard> <A Name=CodeWizard>Code Wizard</A><P>

<%if blnWalkThru then%>
<%end if%>
Administrators must set the variable ocdShowWizard = True in the ocdConfig.asp, Config.asp, or FreeConfig.asp (depending on version) configuration file to activate Code Wizard controls on Browse pages for MS Access and SQL Server.
<P>
Core Browse/Search/Export functions are contained in the file
ocdGrid.asp.  This file lists all built in public properties to customize DataGrid settings along with a short note describing its functions.  By default, properties to set all English language text for the 1 Click DB DataGrid are stored in the corresponding ocdGrid_Lang.asp include files.

<P>
Core Add/Edit/Delete functions are contained in the file
ocdForm.asp.  This file lists all built in public properties to customize Grid settings along with a short note describing its functions.  By default, properties to set all English language text in corresponding for the 1 Click DB DataForm are stored in the corresponding ocdForm_Lang.asp files.

<%if blnWalkThru then%>

<%end if%>
<P>

<A HREF="#Contents">Back To Contents</A>
<P>

<BLOCKQUOTE>
For more information on 1 Click DB including lots of tips for customizing Code Wizard applications, check out our online knowledge base at <A HREF="http://AccessHelp.net/support/" target="_top">http://AccessHelp.net/kb</a> or contact <A HREF="http://AccessHelp.net" target="_top">AccessHelp.net</a> for support.
</BLOCKQUOTE>
<P>
</LI>
<LI><IMG SRC=appAudit.gif ALT=Audit> <A Name=AuditWizard>Audit Wizard</A><P>
1 Click DB Wizard and SQL Server 7+ only.  See the article <A HREF="http://1clickdb.com/content/view.aspx?_@id=53431" target="_top">Time After Time</A> for details on the technique employed.
<P>
<A HREF="#Contents">Back To Contents</A>
<P></LI>
</OL>
<P><HR><P>
<%
call writefooter("")
%>
<P><IMG SRC=Screen_Execute.gif></P>
