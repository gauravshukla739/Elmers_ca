<%'WizardImport_AfterLoad.asp

'open ADO connection objects :
'ocdConnImport 
'ocdConnTarget

'transaction (if selected) is open on ocdConnTarget

'open ADO recordset objects
'rsImportSource
'rsImportTarget

'rsImportSource is read-only firehose cursor

'rsImportTarget starts empty and holds up to 10 inserted records 
'before resetting itself to avoid memory problems in ADO recordset

Response.write "<P>After Load</P>"

%>
