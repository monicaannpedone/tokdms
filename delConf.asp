<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/common.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<head>
	<title>Delete Config Item</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	<script defer src="../js/packs/solid.js"></script>
    <script defer src="../js/packs/regular.js"></script>
    <script defer src="../js/packs/light.js"></script>
    <script defer src="../js/fontawesome.js"></script>
</head>

<body>
	<img src="../images/tcslogo1.gif" />
	
	<%
	bitIsAdmin = Session("IsAdmin")
    bitIsSysAdmin = Session("IsSysAdmin")
	bitLoggedIn = Session("loggedin")
	strUser = Session("username")
	if len(Request.Form("n")) = 0 then
	   strType   = "I"
	   strName   = Request.QueryString("N")
	else
	   strType = "P"
	   strName = Request.Form("N")
	end if   	  
	
      
	%>
	<div id="header">
  
	</div>
	<div id="main">
		<div id="contents">
			<h2>Delete Configuration Item <%=strname%></h2>
            <h3></h3>
            <%			
			
            
			if strType = "I" then
			  strSQL = "SELECT * FROM dms_Config where cfg_Name = '"&strname&"';"
			  ors.Open strSQL
			  if not ors.EOF then
				  strDesc = trim(ors("cfg_desc"))
				  strVal  = trim(ors("cfg_value"))
			  end if	  
			  
			
			%> 
            
            <Form action="DelConf.asp" method="post" name="DelCode" >
                  <input type="hidden" name="N" value="<%=strname%>" /> 
                  
            <table class="dtab">
               <tr>                  
                  <td class="dLabel" width=50>Name</td>
                  <td class="dvalue" width=300><%=strName%></td>
                  <td class="drowerror" width=400>&nbsp;</td>
               </tr> 
               <tr>                  
                  <td class="dLabel" width=50>Description</td>
                  <td class="dvalue" width=300><%=strDesc%></td>
                  <td class="drowerror" width=400>&nbsp;</td>
               </tr> 
               <tr>                  
                  <td class="dLabel" width=50>Value</td>
                  <td class="dvalue" width=300><%=strVal%></td>
                  <td class="drowerror" width=400 >&nbsp;</td>
               </tr> 
               <tr>                  
                  
                  <td class="drowedit" width=350 colspan="2">Please Confirm Delete</td>
                  <td class="drowedit" width=400 align="left"><input  type="submit" value="Delete"></td>
               </tr> 
			</table>
            </Form>
			
			   
       <% 	   
	      else
		      strSQL = "SELECT * FROM dms_Config where cfg_Name = '"&strname&"';"
			   ors.Open strSQL
			   if NOT ors.EOF then
			      datNow = Now()
			      strSQL = "DELETE FROM dms_Config WHERE cfg_name = '"&strname&"';"
				  oconn.Execute strSQL
				  response.redirect("MaintConf.asp?msg=Configuration item "&strname&" was deleted")
			   else
			      response.redirect("ERROR Thing to delete "&strname&"Does not exist!<br>")

			   end if
		  
	      end if
	   %> 
              
       
            
        </div> 
        
	</div>
    <!--#include file="../include/footer.inc"-->
</body>
</html>
