<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' Last updated 06/24/1999

' if the cancel button was hit then return the user
' to the administration page

if Request.Form("cancel") = "No" then
	response.redirect "webNav2_admin.asp"
end if
%>
<!--#include file="data/webNav2_data.inc"-->
<!--#include file="data/webNav2_array.inc"-->
<%
' otherwise delete menu item(s)

' 0 id
' 1 name
' 2 description
' 3 parent
' 4 url
' 5 target frame
' 6 hidden
' 7 depth in tree
' 8 number of children

dim itemIndex

itemIndex = CInt(Request.Form("index"))

for x = itemIndex to itemIndex + menuItem(itemIndex,8)

' delete the selected item and any sub-items
' adCmdText = &H0001

	query = "DELETE FROM menu_items WHERE " _
		& "(item_id)=" & menuItem(x,0)
	db.Execute query,,&H0001
next
db.Close
Set db = nothing

' we need to force regeneration of the menu
' so set the Session array to nothing

Application(dataName & "Menu") = ""

' now send the user back to the administration page

response.redirect "webNav2_admin.asp"
%>