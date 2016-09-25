<% Option Explicit %>
<!--#include file="../include/ths1991_settings.inc"-->
<!--#include file="people_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' updated 5/4/2001

dim strQuery
dim oRS
dim intUserID

intUserID = CInt(Request.Form("id"))
strQuery = "SELECT * FROM people WHERE (user_id)=" & intUserID
Set oRS = Server.CreateObject("ADODB.RecordSet")
oRS.Open strQuery, strDSN, adOpenForwardOnly, adLockOptimistic, adCmdText
if Trim(Request.Form("old_last")) <> Trim(Request.Form("name_last")) then
	oRS("name_newlast") = Request.Form("name_last")
end if
oRS("email_site1") = Request.Form("email_site1")
oRS("email_site2") = Request.Form("email_site2")
oRS("email_name1") = Request.Form("email_user1")
oRS("email_name2") = Request.Form("email_user2")
oRS("user_url1") = Request.Form("url1")
oRS("user_url2") = Request.Form("url2")
oRS("user_address") = Request.Form("address")
oRS("user_city") = Request.Form("city")
oRS("user_state") = Request.Form("state")
oRS("user_zip") = Request.Form("zip")
oRS("user_phone1") = Request.Form("phone1")
oRS("user_phone2") = Request.Form("phone2")
oRS("user_employer") = Request.Form("employer")
oRS("user_occupation") = Request.Form("job")
if Trim(Request.Form("born")) <> "" then
	' Access doesn't like blank dates
	oRS("user_born") = Request.Form("born")
end if
if Trim(Request.Form("died")) <> "" then
	oRS("user_died") = Request.Form("died")
end if
oRS("user_notes") = Request.Form("notes")
oRS("grade_start") = Request.Form("grade_start")
oRS("grade_last") = Request.Form("grade_last")
oRS("last_updated") = Now
oRS.Update
oRS.Close
Set oRS = nothing

response.redirect "../people/people_detail.asp?subcat=1&first=" _
	& Request.Form("first") & "&last=" & Request.Form("old_last")
%>