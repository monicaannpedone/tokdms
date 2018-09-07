<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<script SRC="calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/common.inc" -->
<head>
	<title>pullData</title>
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

<%
   strEM = Request.QueryString("em")
   
   datDate = Date()

%>
		
        <TABLE  Border=0 Cellspacing=2 Cellpadding=2>
        <form action="StartTestingSession1.asp" method="post" name="StartTestingSession1"  >
            <h1>Testing</h1>
            
            <br><br>
            Select the Month and Year
			 <tr>
			   <td >Select Sponsoring School</td>
			   <td> <% Call SelectDojo("SDojo",strDefaultDojo, 0) %> </td>
			 </tr>  
             <tr>
			   <td >Select Participating Schools</td>
			   <td > <% Call CheckboxDojo("PDojo",strDefaultDojo, 0) %></td>
			 </tr>
             <tr>
			   <td >Select Testing Location</td>
			   <td > <% Call SelectFromTable("TestLocation","") %></td>
			 </tr>
			 
			 <tr>  
			   <td class="dLabel1">Testing</td>
			   <td class="entry1"><Input Name="TestingDate" Prototype="MM/DD/YYYY" type="text" value="" size="12" maxlength="12"  />&nbsp;
               <a href="javascript:ShowCalendar('document.StartTestingSession1.TestingDate', document.StartTestingSession1.TestingDate.value);"><img src="calendar.gif"  align="absmiddle" width="18" height="18" border="0" alt="Click Here to Pick up a Date" /> &nbsp;Select a date</a></td>			   
               <td class="drowerror1">&nbsp;</td>   
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
			   <td colspan="3"><%=strEM%>;</td>
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
