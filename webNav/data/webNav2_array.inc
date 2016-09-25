<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' Last updated 05/12/1999

' To prevent conflicting Session variables with multiple
' copies of this program, specify a unique data name for
' this copy:

strDataName = "thsNav"

' if the menu array doesn't already exist then
' create it

if Not IsArray(Session(strDataName & "Menu")) then
	dim strDSN, query, rs, count, x, depth, row, usedRow

	'DSN = "DRIVER={Microsoft Access Driver (*.mdb)}; " _
	'	& "DBQ=" & Server.Mappath("data/" & dataName & ".mdb")
		
	strDSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
		& Server.Mappath("../../../database/" & strDataName & ".mdb")

	Set rs = Server.CreateObject("ADODB.RecordSet")

	query = "SELECT * FROM menu_items ORDER BY item_order"

	' adOpenStatic = 3
	' adLockReadOnly = 1
	' adCmdText = &H0001

	rs.Open query, strDSN, 3, 1, &H0001

	count = CInt(rs.Recordcount - 1)
	Redim menuTemp(count,6)
	Redim menuItem(count,8)

	for x = 0 to count
		' go through each record and put the values
		' into a temporary two dimensional array
		menuTemp(x,0) = rs("item_id")
		menuTemp(x,1) = rs("item_name")
		menuTemp(x,2) = rs("item_description")
		menuTemp(x,3) = rs("item_parent")
		menuTemp(x,4) = rs("item_url")
		menuTemp(x,5) = rs("item_target")
		menuTemp(x,6) = rs("item_hide")
		rs.movenext
	next
	rs.Close
	Set rs = nothing
	
	depth = 0
	row = -1
	usedRow = ""
	
	' with the temporary array created, now go
	' through it to order the items and calculate
	' the items' depth in the tree and number
	' of children, producing these fields:
	' 0 id
	' 1 name
	' 2 description
	' 3 parent
	' 4 url
	' 5 target frame
	' 6 hidden
	' 7 depth in tree (added)
	' 8 number of children (added)

	
	' we prime the subroutine by displaying all
	' root items
	Call orderItems(0,"")
	Session(dataName & "Menu") = menuItem
else
	' if the Session array already exists then
	' assign it to the menuItem array
	menuItem = Session(dataName & "Menu")
end if

'-------------------------------------------------------------------------
Sub orderItems(parentID,pRow)
	' declare these values within the subroutine
	' so that they remain strictly local and can
	' be used recursively
	dim i, j, parentRow, oldRow
	parentRow = pRow
	for i = 0 to UBound(menuTemp,1)
		if parentID = menuTemp(i,3) then
			' row keeps global track of the current
			' position in the menu
			row = row + 1

			' oldRow keeps local track of the current
			' row even after the subroutine has been
			' called recursively to display the children
			' (which increments row and updates parentRow)
			oldRow = row		

			' transfer the contents of the temporary
			' array to the new one
			for j = 0 to 5
				menuItem(row,j) = menuTemp(i,j)
			next

			if menuTemp(i,6) = "True" then
				menuItem(row,6) = 1
			else
				menuItem(row,6) = 0
			end if
			
			menuItem(row,7) = depth
			menuItem(row,8) = 0

			' here we call this same subroutine, recursively,
			' adding to the tree depth each time and sub-
			' tracting that depth when we return
			depth = depth + 1
			Call orderItems(menuTemp(i,0),row)
			depth = depth - 1

			' each item must indicate not only the number of
			' immediate children but all progeny, so we add
			' one to the immediate parent row with each run
			' through and then we add the progeny of the current
			' row to the parent row, so it has a total
			if parentRow <> "" then
				menuItem(parentRow,8) = menuItem(parentRow,8) + menuItem(oldRow,8) + 1
			end if
		end if
	next
end Sub

%>