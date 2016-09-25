<% Option Explicit %>
<% Response.Buffer = True %>
<html>
<head>
<!--#include file="./webAlbum/webAlbum_themes.inc"-->
<%
dim intPics		' number of album pictures
dim oFS			' file system object
dim oFolder		' folder object
dim oFiles		' files object

' get current picture count
Set oFS = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFS.GetFolder(Server.MapPath("webAlbum/pictures/hi-res/"))
Set oFiles = oFolder.Files
intPics = oFiles.Count
Set oFiles = nothing
Set oFolder = nothing
Set oFS = nothing
%>

</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(5)%>" vlink="#<%=arColor(5)%>" alink="#<%=arColor(6)%>">
<center>
<font face="Tahoma, Arial, Helvetica"><font size=5 color="ffffff"><b>Troy High School
Class of 1991</b></font>
<br>
<img src="media/senior-class-intro.jpg" width="500" height="352" alt="" border="1">
<br>
<font size=2>There are <%=intPics%> other pictures online, including recent (newer than 1991) photos of Tod, Brett, Erik, Aaron, Tiffany, Kelly, Cynthia, Susan, James, Randy and Jason—see how beautiful and handsome we've become.
<p>
Use <a href="mailto:ths1991@yahoogroups.com">ths1991@yahoogroups.com</a> to send e-mail to the whole class.
<p>
<!--#include file="./include/donate.inc"-->
<p>
If this isn't the class you were looking for, perhaps you want one of these:
<div style="text-align: left; width: 400px;">
<ul>
<li><a href="http://www.troyclassof81.com/" target="_top">Troy, Idaho Class of 1981</a>
<li><a href="http://groups.yahoo.com/group/THS-Alumni" target="_top">Troy, Idaho alumni message board</a>
<li><a href="http://groups.yahoo.com/group/TROYCLASS1991" target="_top">Troy, Texas Class of 1991</a>
</ul>
<div>

<!-- <form method="post" action="http://www.listbot.com/cgi-bin/subscriber" name="subscribe">
Enter your e-mail address to join the THS e-mail list<br>
<input type=text name="e_mail">
<input type=hidden name="list_id" value="ths1991">
<input type=hidden name="Act" value="subscribe_list">
<input type="submit" value="Join List">
</form> -->


</font>
</center>

</body>

<script language="javascript">
/*
if (confirm("The site host has been experiencing ongoing database problems.\n\n" +
	"I hope they'll have it fixed soon but until then you can visit the THS site at\n\n" +
	"    http://207.198.111.202/ths/default.html\n\n" +
	"Click OK to be redirected now.\n\n" +
	"Thanks for your patience\n    - Jason (1/5/2001)")) {
	
	top.location.href = "http://207.198.111.202/ths/default.html";
}
*/
</script>

</html>