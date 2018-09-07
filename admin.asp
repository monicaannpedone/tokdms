<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="../include/validate.inc"-->
<!--#include file="../include/connect.inc"-->
<!--#include file="../include/locale.inc"-->
<!--#include file="../include/common.inc"-->

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title>Administrative Functions</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	<script defer src="../js/packs/regular.js"></script>
	<script defer src="../js/packs/solid.js"></script>
	<script defer src="../js/packs/light.js"></script>
    <script defer src="../js/fontawesome.js"></script>
	
</head>

<body>
<% strUserName = Session("Username") %>
  <div id="main">
   <% Call TopLine() %>
	<%
	bitIsAdmin = Session("IsAdmin")
    bitIsSysAdmin = Session("IsSysAdmin")
	bitLoggedIn = Session("loggedin")
	%>
		<div id="contents">
			<div id="HomeCol1">
	           <h3>Administrative Functions</h3>
			   <h4><a href="MaintConf.asp">Maintain Configuration Items</a></h4>
    		   <h4><a href="CodeMaint.asp">Maintain Codes</a></h4>
               <h4><a href="DisplayStatusCodes.asp">Display Status Codes</a></h4>
               <h4><a href="AssignGender.asp">Assign Gender Codes</a></h4>
            </div> 
           <div id="HomeCol2">
	            <h3>Dojo</h3>			
				<h4><a href="ListDojo.asp">List Dojo</a></h4>
           </div>
	       <div id="HomeCol3">
	           <h3>Class</h3>
	           <h4><a href="MaintClass.asp">Maintain Classes</a></h4>
			   <br>
			   <h3>Instructors</h3>
		<!--	   <h4 data-balloon="Appoint Instructor" data-balloon-pos="up"><a href="StudentFunction.asp?f=6">Appoint Instructor</a></h4>
				<h4 data-balloon="Appoint Senior Instructor " data-balloon-pos="up"><a href="StudentFunction.asp?f=7">Appoint Senior Instructor</a></h4>
				<h4 data-balloon="Appoint Head Instructor" data-balloon-pos="up"><a href="StudentFunction.asp?f=8">Appoint Head Instructor</a></h4>   -->

	           <br>
	       </div>   
        </div>
   </div>
</body>
</html>
