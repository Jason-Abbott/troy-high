<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="../include/ths1991_settings.inc"-->
<!--#include file="../webAlbum/show_status.inc"-->
<!--#include file="people_data.inc"-->
<%
' Copyright 2001 Jason Abbott (jason@webott.com)
' Last updated 4/27/2001

dim strQuery
dim oRS
dim intSubID		' webAlbum subcategory id
dim intUserID		' user id
dim strName
dim strLast
dim strFirst
dim strNewLast
dim strHTML
dim strFont
dim strTD
dim sMapLink

sMapLink = "http://maps.yahoo.com/py/maps.py?Pyt=Tmap&newFL=Use+Address+Below&country=us&Get%A0Map=Get+Map&addr="

strFirst = Request.QueryString("first")
strLast = Request.QueryString("last")
intSubID = Request.QueryString("subcat")
strTD = "<td align='right'>"

strQuery = "SELECT * FROM people WHERE name_first='" _
	& strFirst & "' AND name_last='" _
	& strLast & "'"
Set oRS = Server.CreateObject("ADODB.RecordSet")
oRS.Open strQuery, strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText
	intUserID = oRS("user_id")
	strNewLast = oRS("name_newlast")

	' personal e-mail
	if oRS("email_name1") <> "" AND oRS("email_site1") <> "" then
		strHTML = strHTML & "<tr><td class='label'>e-mail</td>" _
			& "<td class='detail'><a href='mailto:" _
			& oRS("email_name1") & "@" & oRS("email_site1") & "'>" _
			& oRS("email_name1") & "@" & oRS("email_site1") & "</a></td>"
	end if

	' personal web site
	if oRS("user_url1") <> "" then
		strHTML = strHTML & "<tr><td class='label'>site</td>" & vbCrLf _
			& "<td class='detail'><a href='http://" _
			& oRS("user_url1") & "' target='_top'>http://" & oRS("user_url1") & "</a></td>"
	end if

	' street address and home phone
	if Session(strDataName & "User") <> "" then
		' only show this to classmates
		if oRS("user_address") <> "" OR oRS("user_city") <> "" _
			OR oRS("user_state") <> "" then
			
			strHTML = strHTML & "<tr><td class='label'>address</td><td class='detail'>"
			
			if oRS("user_address") <> "" then
				strHTML = strHTML & "<a href=""javascript:showMap('" _
					& Server.URLEncode(Replace(oRS("user_address"),"#","")) _
					& "','" & oRS("user_zip") & "');"">" _
					& oRS("user_address") & "<br>"
			end if
			
			if oRS("user_city") <> "" then
				strHTML = strHTML & oRS("user_city")
				if oRS("user_state") <> "" then
					strHTML = strHTML & ", "
				end if
			end if
			
			if oRS("user_state") <> "" then
				strHTML = strHTML & oRS("user_state") & " &nbsp;"
			end if
			
			if oRS("user_zip") <> "" then
				strHTML = strHTML & oRS("user_zip")
			end if
			
			strHTML = strHTML & "</a></td>"
			
			if oRS("user_phone1") <> "" then
				strHTML = strHTML & "<tr><td class='label'>phone</td>" _
					& "<td class='detail'>" & oRS("user_phone1") & "</td>"
			end if
		end if			
	end if
	
	' born
	if oRS("user_born") <> "" then
		strHTML = strHTML & "<tr><td class='label'>born</td>" _
			& "<td class='detail'><a href='../webCal/webCal3_month.asp?date=" _
			& oRS("user_born") & "' " & showStatus("View in calendar") _
			& " target='body'>" & FormatDateTime(oRS("user_born"),1) & "</a></td>"
	end if
	
	' died
	if oRS("user_died") <> "" then
		strHTML = strHTML & "<tr><td class='label'>died</td>" _
			& "<td class='detail'><a href='../webCal/webCal3_month.asp?date=" _
			& oRS("user_died") & "' " & showStatus("View in calendar") _
			& " target='body'>" & FormatDateTime(oRS("user_died"),1) & "</a></td>"
	end if
	
	' time at THS
	if oRS("grade_start") <> "" AND oRS("grade_last") <> "" then
		strHTML = strHTML & "<tr><td class='label'>at Troy</td>" _
			& "<td class='detail'>from " & showGrade(oRS("grade_start")) _
			& " to " & showGrade(oRS("grade_last")) & "</td>"
	end if
	
	' last login
	if oRS("last_login") <> "" then
		strHTML = strHTML & "<tr><td class='label'>last login</td>" _
			& "<td class='detail'><a href='../webCal/webCal3_month.asp?date=" _
			& oRS("last_login") & "' " & showStatus("View in calendar") _
			& " target='body'>" & FormatDateTime(oRS("last_login"),1) _
			& "</a>, " & FormatDateTime(oRS("last_login"),3) & " (MST)</td>"
	end if

	' separation bar for employment
	if oRS("user_employer") <> "" OR oRS("user_occupation") <> "" _
		OR oRS("user_url2") <> "" then
		
		strHTML = strHTML & "<tr><td class='section' colspan='2'>employment information</td>"
	end if
	
	' employer
	if oRS("user_employer") <> "" then
		strHTML = strHTML & "<tr><td class='label'>employer</td>" _
			& "<td class='detail'>" & oRS("user_employer") & "</td>"
	end if
	
	' occupation
	if oRS("user_occupation") <> "" then
		strHTML = strHTML & "<tr><td class='label'>occupation</td>" _
			& "<td class='detail'>" & oRS("user_occupation") & "</td>"
	end if
	
	' work e-mail
	if oRS("email_name2") <> "" AND oRS("email_site2") <> "" then
		strHTML = strHTML & "<tr><td class='label'>e-mail</td>" _
			& "<td class='detail'><a href='mailto:" _
			& oRS("email_name2") & "@" & oRS("email_site2") & "'>" _
			& oRS("email_name2") & "@" & oRS("email_site2") & "</a></td>"
	end if
	
	' work web site
	if oRS("user_url2") <> "" then
		strHTML = strHTML & "<tr><td class='label'>site</td>" _
			& "<td class='detail'><a href='http://" _
			& oRS("user_url2") & "' target='_top'>http://" & oRS("user_url2") & "</a></td>"
	end if
	
	' work phone
	if oRS("user_phone2") <> "" AND Session(strDataName & "User") <> "" then
		' only show this to classmates
		strHTML = strHTML & "<tr><td class='label'>phone</td>" _
			& "<td class='detail'>" & oRS("user_phone2") & "</td>"
	end if
	
	' notes
	if Trim(oRS("user_notes")) <> "" then
		' replace cr with HTML breaks
		strHTML = strHTML & "<tr><td colspan='2' class='memo'>" _
			& Replace(Replace(oRS("user_notes"),vbCrLf & vbCrLf,"<p>"),vbCrLf,"<br>") & "</td>"
	end if

oRS.Close
Set oRS = nothing

' format grade level------------------------------------------------------
function showGrade(intGrade)
	dim strShow
	
	select case intGrade
		case 0
			strShow = "kindergarten"
		case 1
			strShow = "first grade"
		case 2
			strShow = "second grade"
		case 3
			strShow = "third grade"
		case 4
			strShow = "fourth grade"
		case 5
			strShow = "fifth grade"
		case 6
			strShow = "sixth grade"
		case 7
			strShow = "seventh grade"
		case 8
			strShow = "eighth grade"
		case 9
			strShow = "freshman year"
		case 10
			strShow = "sophomore year"
		case 11
			strShow = "junior year"
		case 12
			strShow = "senior year"
	end select
	
	showGrade = strShow	
end function


' accomodate new last names
if Trim(strNewLast) <> "" then
	strName = strFirst & " " & strNewLast & " (" & strLast & ")"
else
	strName = strFirst & " " & strLast
end if

%>
<html>
<head>
<link rel='stylesheet' href='../style/ths1991.css' type='text/css'>
<script language="javascript">
// pop window displaying map to store (jea:10/16/00)
function showMap(sAddress, sZip) {
	var url = "<%=sMapLink%>" + sAddress + "&csz=" + sZip;
	winMap = window.open(url,"map","height=725,width=660,scrollbars=yes,titlebar=no,resizable");
	//,screenX="+x+",left="+x+",screenY="+y+",top="+y
}
</script>
</head>
<body>

<table cellspacing="0" cellpadding="0" border="0" width="100%" class='banner'>
<form action="people_edit.asp" method="post">
<tr>
	<td class='banner'><%=strName%></td>
	<% if Session(strDataName & "User") <> "" then %>
	<td align="right">
	<input type="submit" value="Edit" class='button'>
	</td>
	<% end if %>
</table>

<table cellspacing="1" cellpadding="0" border="0">
<%=strHTML%>
</table>

<input type="hidden" name="id" value="<%=intUserID%>">
<input type="hidden" name="last" value="<%=strLast%>">
</form>

</body>
</html>