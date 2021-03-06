<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="../include/validate.inc"-->
<!--#include file="../include/connect.inc"-->
<!--#include file="../include/locale.inc"-->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title>Dojo Fynctions Functions</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
</head>

<body>
	
	<%
	bitIsAdmin = Session("IsAdmin")
    bitIsSysAdmin = Session("IsSysAdmin")
	bitLoggedIn = Session("loggedin")
      
	%>
	<div id="header">
	<ul id="primary">
		<li><a href="Index.asp">Home..</a></li>
		<li><a href="Student.asp">Student</a></li>
        <li><span>Dojo</span></li>
		<li><a href="Class.asp">Class</a></li>
		<li><a href="Attendance.asp">Attendance</a></li>
        <li><a href="Event.asp">Event</a></li>
		<li><a href="Tournament.asp">Tournament</a></li>
		<li><a href="Reports.asp">Reports</a></li>
        <li><a href="Admin.asp">Admin</a></li>
		<% if bitIsSysAdmin = True then %>
             <li><a href="Sysadmin.asp">SysAdmin</a></li>
        <% end if %>
        
	</ul>
    
	</div>
	<div id="main">
		<div id="contents">
			<h2>Dojo Functions</h2>			
			<p class="note">Add or Maintain information about a dojo.</p>
			<h3><a href="Listdojo.asp">List Dojo Data</a></h3>
            <h3>Create printable class schedule</h3>

              <p>&nbsp;</p>
		</div>
        <!--#include file="../include/infooter.inc"-->  
	</div>
    <!--#include file="../include/footer.inc"-->
</body>
</html>
