<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' Last updated 05/27/1999

' To prevent conflicting Session variables with multiple
' copies of this program, specify a unique data name for
' this copy:

strDataName = "thsNav"

' This opens the database webNav.mdb stored in the
' the directory /data under the main webNav files

dim db
Set db = Server.CreateObject("ADODB.Connection")
'DSN = "DRIVER={Microsoft Access Driver (*.mdb)}; " _
'	& "DBQ=" & Server.Mappath("data/" & dataName & ".mdb")
	
strDSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
	& Server.Mappath("../../../database/" & strDataName & ".mdb")
	
db.Open strDSN

' make sure the containing file does
' db.Close
' Set db = nothing

' ADO Constants-----------------------------------------------------------
	
' cursors
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

' cursor location
Const adUseServer = 2
Const adUseClient = 3

' locks
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

' commands
Const adCmdUnknown = &H0008
Const adCmdText = &H0001
Const adCmdTable = &H0002
Const adCmdStoredProc = &H0004
Const adCmdFile = &H0100
Const adCmdTableDirect = &H0200
%>