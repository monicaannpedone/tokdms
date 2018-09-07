<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="../include/validateAdmin.inc"-->
<!--#include file="../include/connect.inc"-->
<!--#include file="../include/locale.inc"-->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title>SWP Admin</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
</head>

<body>
	<img src="../images/swiplogo.jpg" />
	
	<%
	bitIsAdmin = Session("IsAdmin")
    bitIsSysAdmin = Session("IsSysAdmin")
	bitLoggedIn = Session("loggedin")
      
	%>
	
      
	
	<div id="header">
	<ul id="primary">
		<li><a href="admin.asp">Admin</a></li>
		<li><span>Users</span></li>
		
		<li><a href="AdminApp.asp">Application</a></li>
		
        
		
	</ul>
	</div>
	<div id="main">
		<div id="contents">
			<h2>Admin Functions</h2>			
			<p>These are admin only functions  </p>
			<p><a href="Listusers.asp">List Users</a></p>
		</div>
        <!--#include file="../include/admininfooter.inc"-->
	</div>
     <!--#include file="../include/footer.inc"-->
</body>
</html>
