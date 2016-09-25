<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' Last updated 05/27/1999

dim dataName
dim strDelim
dim db

' To prevent conflicting Session variables with multiple
' copies of this program, specify a unique data name for
' this copy:

dataName = "thsCal"
strDelim = "#"			' use ' for SQL, # for Access

' This opens the database stored in the
' the directory /data under the main webCal files

Set db = Server.CreateObject("ADODB.Connection")

' OLEDB
DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
	& Server.Mappath("../../../database/" & dataName & ".mdb")

' ODBC
'DSN = "DRIVER={Microsoft Access Driver (*.mdb)}; " _
'	& "DBQ=" & Server.Mappath("data/" & dataName & ".mdb")

db.Open DSN

' make sure the containing file does
' db.Close
' Set db = nothing
%>