<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/ths1991_settings.inc"-->
<!--#include file="./people/people_data.inc"-->
<%
dim m_sBanner
dim m_sLink

m_sBanner = Request.QueryString("name")
if m_sBanner = "" then
	m_sBanner = "Here's some help"
else
	m_sBanner = "Hi " & m_sBanner & ", welcome to your past."
end if

m_sLink = "Add new items to the menu"

if Session(strDataName & "User") <> "" then
	m_sLink = "<a href='./webNav/webNav2_admin.asp'>" & m_sLink & "</a>"
end if
%>
<html>
<head>
<link rel='stylesheet' href='./style/ths1991.css' type='text/css'>
</head>
<body>

<div class='banner'><%=m_sBanner%></div></font>
<p>
<font size=3>There are a few extra things you can do here as a registered student.</font>
<dl>
<dt><b>Edit personal information</b>
<dd>Select yourself from "Students" in the menu.  You should see an "Edit" button.  Enter whatever information you like.  <b>Your home address and phone number(s) will only be visible to your classmates who have logged in</b>.  At present, you can also edit other people's information.  Try not to be mischevious.  Since only a handful of us are online, we may need to make some edits on behalf of others.
<p>
<dt><b>Edit picture information</b>
<dd>As you're looking through the pictures under any category, you should see an "Edit" button.  Add a description, keywords and date.  Right now, the biggest thing we're missing are dates.  Without dates, the pictures don't get sorted in the right order.  If you can enter dates for as many things as possible, it would be a great help.  Approximations are fine. 
<p>
<dt><b><%=m_sLink%></b>
<dd>If you have some pages you want to link to, you may add them to the menu.  Please don't delete any of the basic categories (Students, Years, Calendar).
<p>
<dt><b>Add or edit events on the calendar</b>
<dd>Click on an existing event to edit it or click on a day of the month to add a new event on that day.
</dl>

</body>
</html>