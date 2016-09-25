<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="../include/ths1991_settings.inc"-->
<!--#include file="people_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' updated 3/22/2000

dim strQuery
dim oRS
dim intUserID
dim strName, strLast, strFirst, strNewLast, strOldLast
dim strEmailSite1, strEmailSite2, strEmailUser1, strEmailUser2
dim strURL1, strURL2
dim strPhone1, strPhone2
dim strAddress, strCity, strState, strZIP
dim strEmployer, strJob
dim strBorn, strDied
dim strNotes
dim intStart, intLast

if Request.Form("id") <> "" then
	intUserID = CInt(Request.Form("id"))
	strQuery = "SELECT * FROM people WHERE (user_id)=" & intUserID
	Set oRS = Server.CreateObject("ADODB.RecordSet")
	oRS.Open strQuery, strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText
		intUserID = oRS("user_id")
		strFirst = oRS("name_first")
		strLast = oRS("name_last")
		strNewLast = oRS("name_newlast")
		strEmailSite1 = oRS("email_site1")
		strEmailSite2 = oRS("email_site2")
		strEmailUser1 = oRS("email_name1")
		strEmailUser2 = oRS("email_name2")
		strURL1 = oRS("user_url1")
		strURL2 = oRS("user_url2")
		strAddress = oRS("user_address")
		strCity = oRS("user_city")
		strState = oRS("user_state")
		strZIP = oRS("user_zip")
		strPhone1 = oRS("user_phone1")
		strPhone2 = oRS("user_phone2")
		strEmployer = oRS("user_employer")
		strJob = oRS("user_occupation")
		strBorn = oRS("user_born")
		strDied = oRS("user_died")
		intStart = oRS("grade_start")
		intLast = oRS("grade_last")
		strNotes = oRS("user_notes")
	oRS.Close
	Set oRS = nothing
	
	strOldLast = strLast
	
	' accomodate new last names
	if Trim(strNewLast) <> "" then
		strName = strFirst & " " & strNewLast & " (" & strLast & ")"
		strLast = strNewLast
	else
		strName = strFirst & " " & strLast
	end if
end if

function gradeList(intGrade)
	dim strHTML
	dim x
	for x = 0 to 12
		strHTML = strHTML & "<option value=" & x
		if intGrade = x then
			strHTML = strHTML & " selected"
		end if
		strHTML = strHTML & ">"
		if x = 0 then
			strHTML = strHTML & "K"
		else
			strHTML = strHTML & x
		end if
	next
	gradeList = strHTML
end function

%>
<html>
<head>
<link rel='stylesheet' href='../style/ths1991.css' type='text/css'>
</head>
<body>

<table cellspacing="0" cellpadding="0" border="0" width="100%" class='banner'>
<form action="people_updated.asp" method="post">
<tr>
	<td class='banner'><%=strName%></td>
	<td align='right'><input type="submit" value="Save" class='button'></td>
</table>

<table cellspacing="1" cellpadding="0" border="0">
<tr>
	<td align="center" colspan="2" class='section'>Personal Information</td>
<tr>
	<td class='label'>surname</td>
	<td class='detail'><input type="text" name="name_last" value="<%=strLast%>"></td>
<tr>
	<td class='label'>e-mail</td>
	<td class='detail'>
	<input type="text" name="email_user1" value="<%=strEmailUser1%>" size="10">@<input type="text" name="email_site1" value="<%=strEmailSite1%>" size="15"></td>
<tr>
	<td class='label'>web site</td>
	<td class='detail'>http://<input type="text" name="url1" value="<%=strURL1%>" size="30"></td>
<tr>
	<td valign="top" class='label'>address</td>
	<td class='detail'>
	<input type="text" name="address" value="<%=strAddress%>" size="30"><br>
<input type="text" name="city" value="<%=strCity%>" size="10">, 
<input type="text" name="state" value="<%=strState%>" size="2"> &nbsp;
<input type="text" name="zip" value="<%=strZIP%>" size="5"></td>
<tr>
	<td class='label'>phone</td>
	<td class='detail'><input type="text" name="phone1" value="<%=strPhone1%>"></td>
<tr>
	<td class='label'>born</td>
	<td class='detail'><input type="text" name="born" value="<%=strBorn%>"></td>
<tr>
	<td class='label'>died</td>
	<td class='detail'><input type="text" name="died" value="<%=strDied%>"></td>
<tr>
	<td class='label'>at Troy</td>
	<td class='detail'>from <select name="grade_start"><%=gradeList(intStart)%></select> to grade <select name="grade_last"><%=gradeList(intLast)%></select></td>
<tr>
	<td align="center" colspan="2" class='section'>What you've done since high school ...</td>
<tr>
	<td colspan='2' class='detail'>
<textarea cols="65" rows="10" name="notes"><%=strNotes%></textarea></td>
<tr>
	<td colspan="2" class='section'>Occupational Information</td>
<tr>
	<td class='label'>employer</td>
	<td class='detail'><input type="text" name="employer" value="<%=strEmployer%>"></td>
<tr>
	<td class='label'>occupation</td>
	<td class='detail'><input type="text" name="job" value="<%=strJob%>"></td>
<tr>
	<td class='label'>e-mail</td>
	<td class='detail'><input type="text" name="email_user2" value="<%=strEmailUser2%>" size="10">@<input type="text" name="email_site2" value="<%=strEmailSite2%>" size="15"></td>
<tr>
	<td class='label'>web site</td>
	<td class='detail'>http://<input type="text" name="url2" value="<%=strURL2%>" size="30"></td>
<tr>
	<td class='label'>phone</td>
	<td class='detail'><input type="text" name="phone2" value="<%=strPhone2%>"></td>

</table>
<input type="hidden" name="id" value="<%=intUserID%>">
<input type="hidden" name="first" value="<%=strFirst%>">
<input type="hidden" name="old_last" value="<%=strOldLast%>">
</form>

</body>
</html>