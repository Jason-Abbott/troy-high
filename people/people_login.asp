<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="../include/ths1991_settings.inc"-->
<!--#include file="people_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 5/4/2001

dim strError
dim strStatus
dim strQuery
dim oRS
dim oConn
dim dtLastLogin

strStatus = "You are not logged in"
strError = "The information you entered could not be validated. " _
	& "Please try again."

if Session(strDataName & "User") <> "" then
	' assume the user wants to logout
	Session(strDataName & "User") = ""
	' logout of other apps-------------------
	Session("thsNavUser") = ""
	Session("thsAlbumUser") = ""
	Session("thsCalUser") = ""
	Session("thsLastDate") = ""
	' --------------------------------------
	response.redirect "people_login.asp"
elseif Request.Form("login") <> "" then
	strQuery = "SELECT user_id, user_login, user_password, name_first, last_login FROM people WHERE " _
		& "user_login = '" & Request.Form("login") & "'"

	Set oConn = Server.CreateObject("ADODB.Connection")
	Set oRS = Server.CreateObject("ADODB.RecordSet")
	oConn.Open strDSN
	oRS.Open strQuery, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
	if oRS.EOF = -1 then
		strStatus = strError
	else
		if oRS("user_password") = Request.Form("password") then
			Session(strDataName & "User") = oRS("user_id")
			' login to other apps-------------------
			Session("thsNavUser") = oRS("user_id")
			Session("thsAlbumUser") = oRS("user_id")
			Session("thsCalUser") = oRS("user_id")
			' --------------------------------------
			
			' update login date
			dtLastLogin = oRS("last_login")
			strQuery = "UPDATE people SET last_login = #" & Now & "# WHERE user_id = " & oRS("user_id")
			oConn.Execute strQuery
			Session("thsLastDate") = dtLastLogin
			
			if dtLastLogin = "" or IsNull(dtLastLogin) then
				' first time student login
				response.redirect "../intro_student.asp?name=" & oRS("name_first")
				
			else
				response.redirect "people_list.asp"
			end if
		else
			strStatus = strError
		end if
	end if
	oRS.Close : Set oRS = nothing
	oConn.Close : Set oConn = nothing
end if
%>
<html>
<head>
<link rel='stylesheet' href='../style/ths1991.css' type='text/css'>
</head>
<body>
<center>

<table border='0' cellpadding='3' cellspacing='0' class='login'>
<form action="people_login.asp" method="post">
<tr valign="bottom">
	<td colspan='2' class='loginBanner'>Login</td>
<tr>
	<td colspan='2' align="center"><%=strStatus%></td>
<tr>
	<td class='label'>Username:</td>
	<td class='login'><input type="text" name="login" size='10'></td>
<tr>
	<td class='label'>Password:</td>
	<td class='login'><input type="password" name="password" size='10'></td>
<tr>
	<td colspan='2' align="center"><input type="submit" value="Sign in" class='button'></td>
</form>
</table>

If you're part of the class but can't seem to login, <a href="mailto:jason@webott.com">e-mail Jason</a> for help.
</center>
</body>
</html>