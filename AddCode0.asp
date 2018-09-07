<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title>Administrative Functions - Add Physical Media Code</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
</head>

<body>
	
	<%
	bitIsAdmin = Session("IsAdmin")
    bitIsSysAdmin = Session("IsSysAdmin")
	bitLoggedIn = Session("loggedin")
	strCodename  = Request.QueryString("cd")
	strErrMsg    = Request.QueryString("ErrMsg")
	response.write("strCodeName=>"&strCodeName&"<<br>")
	response.write("strErrMsg  =>"&strErrMsg&"<<br>")
    response.end
	
	if len(trim(strCodeName)) > 0 then
	else
	  ' response.redirect("AddCode.asp?"&strCodeName)
    end if
	%>
	<div id="header">
	
  
	</div>
	<div id="main">
		<div id="contents">
	  <h2>Add a Code A</h2>
            <Form action="AddCode00.asp" method="post" name="AddCode00" >
            <table class="dtab">
            <%
            
            %>
               <tr>                  
                  <td class="dtabh" width=32>Code</td>
                  <td class="drownone" width=200>&nbsp;</td>
               </tr> 
               <tr>                  
                  <td class="drowname" width=32 align="left"><input name="CodeName" type="text"  class="entry" size="16"  width="16" maxlength="16" value="" /></td>
                  <td class="drowerror" width=200><%=strErrMsg%></td>
               </tr> 
               <tr>                  
                  <td class="drowname" width=32 align="left"><input  type="Reset" value="Reset"></td>
                  <td class="drowname" width=200 align="right"><input  type="submit" value="Add"></td>
                  
               </tr> 
			</table>
            </Form>
			
			   
       <% 	   
	   
	   
	   %> 
              
       
            
        </div> 
    <!--#include file="../include/infooter.inc"-->     
        
	</div>
    <!--#include file="../include/footer.inc"-->
</body>
</html>
