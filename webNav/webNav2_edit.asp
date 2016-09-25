<!--#include file="data/webNav2_array.inc"-->
<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' Last updated 06/24/1999

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

if Request.Form("index") <> "" then
	itemIndex = CInt(Request.Form("index"))
	if Request.Form("organize") = "Organize" then
		response.redirect "webNav2_org.asp?index=" _
			& itemIndex
	end if
end if
%>

<html>
<head>
<!--#include file="webNav2_themes.inc"-->

<% if Request.Form("delete") = "Delete" then %>

</head>
<body bgcolor="#<%=color(1)%>" link="#<%=color(7)%>" vlink="#<%=color(7)%>" alink="#<%=color(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=color(5)%>" width="60%" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=color(11)%>" border=0 cellpadding=4 cellspacing=0 width="100%">
<form action="webNav2_deleted.asp" method="post">
<tr bgcolor="#<%=color(3)%>" valign="bottom">
	<td><font face="Tahoma, Arial, Helvetica" size=4>
	<b>Item Deletion</b></font></td>
<tr>
	<td align="center"><font face="Arial, Helvetica">
	Are you sure you want to erase the menu item "<%=menuItem(itemIndex,1)%>"
<% if menuItem(itemIndex,8) > 0 then %>
	and the <%=menuItem(itemIndex,8)%> item<%if menuItem(itemIndex,8) > 1 then%>s<%end if%> it contains
<% end if %>
	?</font></td>
<tr>
	<td align="center" bgcolor="#<%=color(12)%>">
		<input type="submit" name="delete" value="Yes">
		<input type="submit" name="cancel" value="No">
	</td>
<tr>
	<td align="center"><font face="Tahoma, Arial, Helvetica" size=2>
	<b><font color="#cc0000">Caution</font>: erased items cannot be restored</b></font>
	</td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="index" value="<%=itemIndex%>">
</form>

<%
else
' ----------------------------------
' EDIT FORM
' ----------------------------------
' if any button other than delete was hit, display the edit form
%>
<!--#include file="data/webNav2_data.inc"-->

<%
	dim itemID, itemName, itemDescription, itemParent, itemURL, itemTarget

	if Request.Form("edit") = "Edit" then
' edit the existing menu item

		itemID = menuItem(itemIndex,0)
		itemName = menuItem(itemIndex,1)
		itemDescription = menuItem(itemIndex,2)
		itemParent = menuItem(itemIndex,3)
		itemURL = menuItem(itemIndex,4)
		itemTarget = menuItem(itemIndex,5)
		
		if menuItem(itemIndex,6) then
			itemHidden = " checked"
		else
			itemHidden = ""
		end if
	else
' we must be adding a new event so the selected
' menu item will be the parent of the new item
' which is 0 if root (-1) was selected

		if Request.QueryString("parent") <> "" then
			itemParent = CInt(Request.QueryString("parent"))
		elseif itemIndex = -1 then
			itemParent = 0
		else
			itemParent = menuItem(itemIndex,0)
		end if
		
		itemID = 0
		itemName = ""
		itemURL = ""
		itemTarget = ""
		itemDescription = ""
		itemHidden = ""
	end if
%>

<SCRIPT LANGUAGE="javascript">
<!--
function Validate() {
	if (document.editform.item_name.value.length <= 0) {
		alert("You must enter a name for the item");
		document.editform.item_name.select();
		document.editform.item_name.focus();
		return false;
	}
}
//-->
</SCRIPT>
</head>
<body bgcolor="#<%=color(1)%>" link="#<%=color(7)%>" vlink="#<%=color(7)%>" alink="#<%=color(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=color(5)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table border=0 cellpadding=4 cellspacing=0>
<form name="editform" method="post" action="webNav2_updated.asp">
<tr>
	<td bgcolor="#<%=color(3)%>" colspan=2>
	<font face="Tahoma, Arial, Helvetica" size=4>
	<b>Menu Item Details</b></font></td>
<tr>
	<td align="right" bgcolor="#<%=color(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2 color=<%=color(14)%>>
	<b>Name:</b></font></td>
	<td bgcolor="#<%=color(11)%>">
	<input name="item_name" type="text" size="10" max="20" value="<%=itemName%>"></td>
<tr>
	<td align="right" bgcolor="#<%=color(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2 color=<%=color(14)%>>
	<b>Description:</b></font></td>
	<td bgcolor="#<%=color(11)%>">
	<input name="item_description" type="text" size="20" max="50" value="<%=itemDescription%>"></td>
<tr>
	<td align="right" bgcolor="#<%=color(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2 color=<%=color(14)%>>
	<b>URL:</b></font></td>
	<td bgcolor="#<%=color(11)%>">
	<input name="item_url" type="text" size="20" max="200" value="<%=itemURL%>"></td>
<tr>
	<td align="right" bgcolor="#<%=color(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2 color=<%=color(14)%>>
	<b>Target:</b></font></td>
	<td bgcolor="#<%=color(11)%>">
	<input name="item_target" type="text" size="20" max="20" value="<%=itemTarget%>"></td>
<tr>
	<td align="right" bgcolor="#<%=color(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2 color=<%=color(14)%>>
	<b>Parent:</b></font></td>
	<td bgcolor="#<%=color(11)%>">
	<select name="item_parent">
<option value=0>[root]
<%
	for x = 0 to UBound(menuItem,1)
	
' don't display the item being edited, or any of its
' children, in the selection list since none of them
' can be a valid parent
	
		if menuItem(x,0) = itemID then
			x = x + menuItem(x,8) + 1
			if x > UBound(menuItem,1) then
				Exit for
			end if
		end if
		
		response.write "<option value=" _
			& menuItem(x,0)
			
		if menuItem(x,0) = itemParent then
			response.write " selected"
		end if

		response.write ">"

		for d = 0 to menuItem(x,7)
			response.write "&nbsp;&nbsp;&nbsp;"
		next

		if menuItem(x,6) then
			response.write "(" & menuItem(x,1) & ")"
		else
			response.write menuItem(x,1)
		end if
	
		response.write VbCrLf
	next	
%>
	</select></td>
<tr>
	<td align="right" bgcolor="#<%=color(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2 color=<%=color(14)%>>
	<b>Hidden:</b></font></td>
	<td bgcolor="#<%=color(11)%>">
	<input name="item_hide" type="checkbox"<%=itemHidden%>></td>
<tr>
	<td colspan=2 bgcolor="#<%=color(12)%>" align="center"><nobr>
	<input type="submit" name="save" value="Save" onClick="return Validate();">
	<input type="submit" name="saveadd" value="Save & Add Another" onClick="return Validate();">
	<input type="submit" name="cancel" value="Cancel"></nobr>
	</td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="item_id" value=<%=itemID%>>
</form>

<table width="60%">
<tr>
	<td colspan=2><font face="Verdana, Arial, Helvetica" size=2>
	<dl>
	<dt><font color="#<%=color(6)%>"><b>Name</font>
	<dd>Required</b>.  The name that appears in the menu.
	<p>
	<dt><font color="#<%=color(6)%>"><b>Description</font>
	<dd>Optional</b>.  The text that appears in the status bar when the mouse is moved over the item in the menu.
	<p>
	<dt><font color="#<%=color(6)%>"><b>URL</font>
	<dd>Optional</b>.  The relative path to a page on the local server or a full URL (http://...) to a page on another server.  If you leave this blank the item's only practical use is as a menu category.  Note that a single item can be a category <b>and</b> a link to content, so you are not required to leave this blank if you intend the item to be a category.
	<p>
	<dt><font color="#<%=color(6)%>"><b>Target</font>
	<dd>Optional</b>.  The frame in which to display the specified URL.  If you leave this blank then it will be displayed in the "_top" frame, that is, the full browser window.
	<p>
	<dt><font color="#<%=color(6)%>"><b>Parent</b></font>
	<dd>The menu item under which this one should be listed.  Any existing menu item can be selected.
	<p>
	<dt><font color="#<%=color(6)%>"><b>Hidden</b></font>
	<dd>Select this if you want to disable this menu item along with any items it contains.  Hidden items are shown in (parantheses) in administration screens.
	<p>
	</dl>
	</td>
</table>

<%	end if %>
</center>
</body>
</html>