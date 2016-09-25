<!-- 
Copyright 1999 Jason Abbott (jason@webott.com)
Last updated 07/30/1999
-->

<html>
<head>
<!--#include file="webNav2_themes.inc"-->
</head>
<body bgcolor="#<%=color(10)%>" text="#<%=color(7)%>" link="#<%=color(8)%>" vlink="#<%=color(8)%>" alink="#<%=color(9)%>">
<center><img src="../media/trojan1.jpg" width=100 height=109 alt="" border="0"></center>
<font face="Arial, Helvetica" size=2>


<!--#include file="data/webNav2_THSarray.inc"-->
<!--#include file="show_status.inc"-->
<%
' the above include provides the array menuItem
' with these contents:
' 0 id
' 1 name
' 2 description
' 3 parent
' 4 url
' 5 target frame
' 6 hidden
' 7 depth in tree
' 8 number of children

dim expand

' lists expanded folders by index number

if Request.QueryString("expand") <> "" then
	expand = Request.QueryString("expand")
else
	expand = ""
end if

for x = 0 to UBound(menuItem,1)
	' display item only if it is not hidden

	if menuItem(x,6) = 0 then
		' thisx keeps track of current item even after x is
		' possibly incremented to skip children
		thisx = x
		response.write "<nobr>"

		' insert a blank space for every unit of tree depth		
		for d = 1 to (menuItem(x,7))
			response.write "<img src='./images/blank_" _
				& style & ".gif'>"
		next
%>
<!--#include file="webNav2_item.inc"-->
<%		
		' if there are children then display plus/minus graphic
		if menuItem(x,8) > 0 then
			response.write "<a name='" & x & "'><a href='webNav2.asp?expand="
	 		if InStr(expand, "(" & x & ")") then
				' if the current item is selected for expansion then
				' display option to collapse
				response.write Replace(expand, "(" & x & ")", "") _
					& "' " & showStatus("collapse " & menuItem(x,1)) _
					& "><img src=""./images/minus_"
			else
				' otherwise show option to expand and skip all children
				response.write expand & "(" & x & ")" _
					& "#" & x & "' " & showStatus("expand " & menuItem(x,1)) _
					& "><img src=""./images/plus_"
				x = x + menuItem(x,8)	
			end if
			response.write style & ".gif"" border=0></a></a>"
		else
			response.write "<img src=""./images/blank_" _
				& style & ".gif"">"
		end if
		response.write VbCrLf & item & VbCrLf & "<br>" & VbCrLf

		' we're done displaying the current menu item
		' if items remain then adjust font size appropriately
		
		if x < UBound(menuItem,1) then
			' if the current item is root (depth=0) but the next is a
			' child then set font size to 1
	 		if menuItem(thisx, 7) = 0 AND menuItem(x + 1, 7) > 0 then
				response.write "<font size=1>"
				
				' if the current item is a child (depth>0) but the next
				' is at the root level then restore default font size
				
			elseif menuItem(thisx, 7) > 0 AND menuItem(x + 1, 7) = 0 then
				response.write "</font>"
			end if
		end if
	else
		' if the item should be hidden then also skip children
		x = x + menuItem(x,8)
	end if
next
%>
</table>
</body>
</html>