<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/common.inc" -->

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title>Administrative Functions - Add  Code</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	<script defer src="../js/packs/regular.js"></script>
	<script defer src="../js/packs/solid.js"></script>
	<script defer src="../js/packs/light.js"></script>
    <script defer src="../js/fontawesome.js"></script>
	
</head>

<body>
	
	<%
	bitIsAdmin = Session("IsAdmin")
    bitIsSysAdmin = Session("IsSysAdmin")
	bitLoggedIn = Session("loggedin")
	strCodename  = Request.QueryString("cd")
	strErrMsg    = Request.QueryString("ErrMsg")
	'response.write("strCodeName=>"&strCodeName&"<<br>")
	'response.write("strErrMsg  =>"&strErrMsg&"<<br>")
   
	
	if len(trim(strCodeName)) > 0 then
	else
	  ' response.redirect("AddCode.asp?"&strCodeName)
    end if
	%>
	<div id="header">
	</div>
	<div id="main">
	    <% Call TopLine() %>
		<div id="contents">
	  <h2>Add a Code </h2>
            <Form action="AddCodeName1.asp" method="post" name="AddCodeName1" >
            <table class="dtab">
            <%
            
            %>
               <tr>                  
                  <td class="dtabh" width=32>Code</td>
                  <td class="dtabh" width=400>&nbsp;</td>
               </tr> 
               <tr>                  
                  <td class="drowname" width=32 align="left"><input name="CodeName" type="text"  class="entry" size="16"  width="16" maxlength="16" value="" /></td>
                  <td class="drowerror" width=400><%=strErrMsg%></td>
               </tr> 
               <tr>                  
                  <td class="dtabh" width=32>Code</td>
                  <td class="dtabh" width=400>&nbsp;</td>
               </tr> 
               
               <tr>                  
                  <td class="drowname" width=32 align="left"><input  type="Reset" value="Reset"></td>
                  <td class="drowname" width=400 align="right"><input  type="submit" value="Add"></td>
                  
               </tr> 
			</table>
            </Form>
			
			   
       <% 	   
	   
	   
	   %> 
        </div> 
	</div>
</body>
</html>
