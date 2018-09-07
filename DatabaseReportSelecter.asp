<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<script SRC="calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!--#include file="../include/validateSysadmin.inc"-->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/common.inc" -->
<head>
	<title>Database Report Selector</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	<script defer src="../js/packs/solid.js"></script>
    <script defer src="../js/packs/regular.js"></script>
    <script defer src="../js/packs/light.js"></script>
    <script defer src="../js/fontawesome.js"></script>

</head>

<body>
<div id="main">
      <% Call TopLine() %>
	  <div id="contents">
	  <div id="OneCol">



		
        <TABLE  Border=0 Cellspacing=2 Cellpadding=2>
        <form action="DatabaseList.asp" method="post" name="DatabaseList.asp"  >
            <h1>Select Database</h1>
			 <tr>
			   <td  class="dLabel1">Database Local</td>			
			   <td class="entry">
                <input type="radio" name="dblocale" value="local" checked> Local<br>
                <input type="radio" name="dblocale" value="remotewinhost"> Remote Winhost<br>
                <input type="radio" name="dblocale" value="remotebrinkster"> Remote Brinkster
			   </td>
             </tr> 
			   <td  class="dLabel1">Database Stage</td>			
			   <td class="entry">
                <input type="radio" name="dbstage" value="local" checked>dev<br>
                <input type="radio" name="dbstage" value="remotewinhost">qa<br>
                <input type="radio" name="dbstage" value="remotebrinkster">prod
			   </td>
             </tr> 
             <tr>
			   <td colspan="3">&nbsp;</td>
             </tr>  
             <tr>
			   <td ><input type="submit" name="submit" value="Submit" /></td>
			   <td align="right"><input type='reset'  /></td>			   
               <td >&nbsp;</td>   
             </tr> 
             <tr>
			   <td colspan="3">&nbsp;</td>
             </tr>
             <tr>
			   <td colspan="3">&nbsp;</td>
             </tr>  
             <tr>
			   <td colspan="3">&nbsp;</td>
             </tr>
            <tr>
               <td>&nbsp;</td>
               <td">&nbsp;</td>
               <td>&nbsp;</td>
            </tr>
            </form>
      </TABLE>
		</div>
	</div>
   </div>
	</body>
</html>
