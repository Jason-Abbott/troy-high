<!-- 
Copyright 1999 Jason Abbott (jason@webott.com)
Last updated 06/24/1999
-->

<html>
<head>
<SCRIPT LANGUAGE="javascript">
<!--

function Move(direction) {
	var sel = document.orgForm.menu.selectedIndex;
	var menu = document.orgForm.menu.options;
	var length = document.orgForm.menu.length;
	var list = "";
	
	if (sel == -1) {
		alert("You must select an item to move");
		return false;
	}
	if (sel == 0 && direction == -1) {
		alert("Already at top of list");
		return false;
	}
	if (sel == length - 1 && direction == 1) {
		alert("Already at end of list");
		return false;
	}
	
	var oldValue = menu[sel+direction].value;
	var oldText = menu[sel+direction].text;
	
	menu[sel+direction].value = menu[sel].value;
	menu[sel+direction].text = menu[sel].text;

	menu[sel].value = oldValue;
	menu[sel].text = oldText;
	
	menu[sel+direction].selected = true;
	document.orgForm.menu.selectedIndex = sel+direction;
	
	for (var i=0; i < length; i++) {
		list = list + menu[i].value + ",";
	}

	document.orgForm.list.value = list;
	return true;
}

function Validate() {
	if (document.orgForm.list.value.length <= 0) {
		alert("You must organize the items before saving");
		return false;
	}
}

//-->
</SCRIPT>
<!--#include file="webNav2_themes.inc"-->
</head>
<body bgcolor="#<%=color(1)%>" link="#<%=color(7)%>" vlink="#<%=color(7)%>" alink="#<%=color(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=color(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=color(11)%>" border=0 cellpadding=3 cellspacing=0>
<form name="orgForm" action="webNav2_updated.asp" method="post">
<tr bgcolor="#<%=color(3)%>" valign="bottom">
	<td colspan=2><font face="Tahoma, Arial, Helvetica" size=4>
	<b>Menu Organization</b></font></td>
<tr>
	<td valign="center" align="right" rowspan=2>

<!--#include file="data/webNav2_array.inc"-->
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

itemIndex = CInt(Request.QueryString("index"))

dim itemParent

if itemIndex = -1 then
' if the "root" item was selected then
' the parent is 0

	itemParent = 0
elseif menuItem(itemIndex,8) > 1 then
' if the selected item has more than one child then
' treat it as the parent

	itemParent = menuItem(itemIndex,0)
else
' otherwise find the value of the selected item's
' parent

	itemParent = menuItem(itemIndex,3)
end if

count = 0
options = ""

for x = 0 to UBound(menuItem,1)
	if menuItem(x,3) = itemParent then
		count = count + 1
		options = options & "<option value=""" _
			& menuItem(x,0) & """>"
		if menuItem(x,6) then
			options = options & "(" & menuItem(x,1) & ")"
		else
			options = options & menuItem(x,1)
		end if
		options = options & VbCrLf
	end if
next
%>
<select name="menu" size=<%=count%> multiple>
<%=options%>
</select>
	</td>
	<td valign="top">
	<input type="button" value=" up " onClick="Move(-1);">
	</td>
<tr>
	<td valign="bottom">
	<input type="button" value="down" onClick="Move(1);">
	</td>
<tr>
	<td colspan=2 align="center" bgcolor="#<%=color(12)%>">
	<input type="submit" name="save" value="Save" onClick="return Validate();">
	<input type="submit" name="cancel" value="Cancel">
	</td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="list" value="">
</form>
</body>
</html>