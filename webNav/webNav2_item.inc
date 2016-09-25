<%
' this generates the HTML for each menu
' item

' it depends upon
' 1) the showStatus() function having
'    already been defined
' 2) the existence of the menuItem array:
'    0 id
'    1 name
'    2 description
'    3 parent
'    4 url
'    5 target frame
'    6 hidden
'    7 depth in tree
'    8 number of children

' if a url was specified for the item then
' generate the link

if menuItem(x,4) <> "" then
	item = "<a href='" & menuItem(x,4) _
		& "' target='"
	if menuItem(x,5) <> "" then
		item = item & menuItem(x,5)
	else
	
' if no frame target was specified then
' make _top the target

		item = item & "_top"
	end if

	item = item & "' "

' if a description was entered for the
' item then generate the javascript to
' display it in the status bar

	if menuItem(x,2) <> "" then
		item = item & showStatus(menuItem(x,2))
	else
		item = item & showStatus(menuItem(x,1))
	end if
	item = item & ">" & menuItem(x,1) & "</a>"
else
	item = menuItem(x,1)
end if

item = item & "</nobr>"
%>