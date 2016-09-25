<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="../include/ths1991_settings.inc"-->
<!--#include file="people_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 4/27/2001

dim strQuery
dim oRS, oRS2
dim m_strClass
dim m_strLinkClass
dim strHTML
dim intSpan
dim strSpam
dim strEmail
dim intSubID
dim bUser
dim bNew
dim strNewText
dim m_strCountClass
dim x
dim sMapLink

sMapLink = "http://maps.yahoo.com/py/maps.py?Pyt=Tmap&newFL=Use+Address+Below&country=us&Get%A0Map=Get+Map&addr="

x = 1
	
if Session(strDataName & "User") <> "" then
	bUser = true
else	
	bUser = false
end if
bNew = false

strQuery = "SELECT * FROM people ORDER BY name_last, name_first"
Set oRS = Server.CreateObject("ADODB.RecordSet")
oRS.Open strQuery, strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText

' open a recordset of webAlbum subcategories so we can link to them
strDSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
	& Server.Mappath(g_sBaseData & "thsAlbum.mdb")
	
strQuery = "SELECT * FROM subcategory WHERE subcat_cat=1"
Set oRS2 = Server.CreateObject("ADODB.RecordSet")
oRS2.Open strQuery, strDSN, adOpenStatic, adLockReadOnly, adCmdText

do while not oRS.EOF
	' move the webAlbum cursor to the subcategory for this user
	strEmail = ""
	oRS2.Filter = "subcat_name='" & oRS("name_last") & ", " & oRS("name_first") & "'"
	if not oRS2.EOF AND not oRS2.BOF then
		' if there's a matching subcategory then get id for href
		intSubID = oRS2("subcat_id")
	end if
	'http://jason/webott/hosted/ths/webAlbum/webAlbum_view-frame.asp?subcat=24
	
	if IsNull(oRS("grade_last")) or oRS("grade_last") < 12 or oRS("grade_last") = "" then
		' student is non-graduate at Troy
		m_strClass = "non"
		m_strLinkClass = "OldPeopleList"
	else
		m_strLinkClass = "PeopleList"
		if x mod 2 = 0 then
			m_strClass = "even"
		else
			m_strClass = "odd"
		end if
	end if
	
	m_strCountClass = "count"
	'Session("theLastDate") <> "" and
	if Session(strDataName & "User") <> "" and oRS("last_updated") <> "" and oRS("user_id") <> Session(strDataName & "User") then
		if oRS("last_updated") > Session("thsLastDate") then
			' student has updated info since last login
			m_strCountClass = "updated"
			bNew = true
		end if
	end if
	
	strHTML = strHTML & "<tr><td class='" & m_strCountClass & "'>" & x & "&nbsp;</td>" _
		& "<td class='" & m_strClass & "One' align='right'>" _
		& "<a href='../webAlbum/webAlbum_view-frame.asp?subcat=" & intSubID _
		& "' class='" & m_strLinkClass & "'><b><nobr>" & oRS("name_last") & ", " _
		& oRS("name_first") & "</nobr></b></a>&nbsp;</td>" _
		& "<td class='" & m_strClass & "Two' valign='top'>"
	
	' personal contact information
	if oRS("email_name1") <> "" AND oRS("email_site1") <> "" then
		strHTML = strHTML & "<a href='mailto:" _
			& oRS("email_name1") & "@" & oRS("email_site1") & "' class='" & m_strLinkClass & "'>" _
			& oRS("email_name1") & "@" & oRS("email_site1") & "</a><br>"
		strEmail = oRS("email_name1") & "@" & oRS("email_site1")
	end if
	
	if oRS("user_url1") <> "" then
		strHTML = strHTML & "<a href='http://" & oRS("user_url1") _
			& "' target='_top' class='" & m_strLinkClass & "'>" & oRS("user_url1") & "</a><br>"
	end if
	
	if bUser then
		' hide from non-classmates
		if oRS("user_phone1") <> "" then
			strHTML = strHTML & oRS("user_phone1") & "<br>"
		end if
		
		if oRS("user_address") <> "" OR oRS("user_city") <> "" OR oRS("user_state") <> "" then
			if oRS("user_address") <> "" then
				strHTML = strHTML & "<a href=""javascript:showMap('" _
					& Server.URLEncode(Replace(oRS("user_address"),"#","")) _
					& "','" & oRS("user_zip") & "');"" class='address'>" _
					& oRS("user_address") & "<br>"
			end if
			
			if oRS("user_city") <> "" then
				strHTML = strHTML & oRS("user_city")
			end if
				
			if oRS("user_city") <> "" AND oRS("user_state") <> "" then
				strHTML = strHTML & ", " & oRS("user_state")
			end if
			
			if oRS("user_zip") <> "" then
				strHTML = strHTML & " &nbsp;" & oRS("user_zip")
			end if
			strHTML = strHTML & "</a>"
		end if
	end if
	
	if Right(strHTML, 6) = "'top'>" then strHTML = strHTML & "&nbsp;"
	
	' work contact information
	strHTML = strHTML & "</td><td class='" & m_strClass & "One' valign='top'>"
	
	if oRS("email_name2") <> "" AND oRS("email_site2") <> "" then
		strHTML = strHTML & "<a href='mailto:" _
			& oRS("email_name2") & "@" & oRS("email_site2") & "' class='" & m_strLinkClass & "'>" _
			& oRS("email_name2") & "@" & oRS("email_site2") & "</a><br>"
			
			if strEmail = "" then strEmail = oRS("email_name2") & "@" & oRS("email_site2")
	end if
	
	if oRS("user_url2") <> "" then
		strHTML = strHTML & "<a href='http://" & oRS("user_url2") _
			& "' target='_top' class='" & m_strLinkClass & "'>" & oRS("user_url2") & "</a><br>"
	end if

	if oRS("user_phone2") <> "" AND bUser then
		' hide from non-classmates
		strHTML = strHTML & oRS("user_phone2")
	end if
	
	if Right(strHTML, 6) = "'top'>" then strHTML = strHTML & "&nbsp;"
		
	strHTML = strHTML & "</td>"
	if strEmail <> "" and _
		oRS("name_last") <> "Kajava" and _
		oRS("name_last") <> "Christensen" and _
		oRS("name_last") <> "Rasmussen" then
		strSpam = strSpam & "; " & strEmail
	end if
	oRS.MoveNext
	x = x + 1
loop
oRS.Close : Set oRS = nothing
oRS2.Close : Set oRS2 = nothing

if bNew then
	strNewText = "<div class='new'>Information has been updated since your last login</div>"
end if

' the spammer doesn't work because the mailto link truncates at 255 characters
strSpam = Right(strSpam, Len(strSpam) - 2) & "; lmarone@yahoo.com"
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

<div class='banner'>All of us</div>

<% if not bUser then %><div class='new'>You must login to see home addresses and phone numbers</div><% end if %>
<%=strNewText%>

<center>
Use <a href="mailto:ths1991@yahoogroups.com">ths1991@yahoogroups.com</a> to send e-mail to the whole class.<br>
<table cellspacing="0" cellpadding="1" border="0" class="PeopleList">
<tr>
	<td></td>
	<td class="Head"><b>Name</b></td>
	<td class="Head"><b>Home</b></td>
	<td class="Head"><b>Work</b></td>
<%=strHTML%>
</table>
</center>

<%'if bUser then %>
<!-- <br><center><textarea cols="80" rows="7" class='email'><%=strSpam%></textarea></center> -->
<%'end if %>

</body>
</html>
