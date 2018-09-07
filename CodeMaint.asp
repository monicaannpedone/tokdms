<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="../include/validate.inc"-->
<!--#include file="../include/connect.inc"-->
<!--#include file="../include/locale.inc"-->
<!--#include file="../include/common.inc"-->

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title>Administrative Functions - Maintain Codes</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	<script defer src="../js/packs/regular.js"></script>
	<script defer src="../js/packs/solid.js"></script>
	<script defer src="../js/packs/light.js"></script>
    <script defer src="../js/fontawesome.js"></script>
</head>

<body>
<div id="main">
      <% Call Topline() %>
	  <div id="contents">
	  <div id="OneCol">
	  <h1>Code Maintenance</h1>
	  <h3></h3>
      <a href="index.asp">Return to Main Menu</a><br><br>
	
	<%
	bitIsAdmin = Session("IsAdmin")
    bitIsSysAdmin = Session("IsSysAdmin")
	bitLoggedIn = Session("loggedin")
	strCodename = Request.QueryString("cd")
	strCodeValue = Request.QueryString("cv")
	strErrMsg    = Request.QueryString("ErrMsg")
	
	

	strSQL="SELECT DISTINCT CodeName FROM dms_Codes ORDER BY CodeName;"
			bitFlop = True
            icnt = 0
            oRS.Open strSQL
            %>
            <a href="AddCode.asp"><i class="far fa-plus"></i>&nbsp;&nbsp;Add a Code</a></br>
            <table>
            <tr>
               <td class="dtabh">Edit</td>
               <td class="dtabh">Code</td>
               <td class="dtabh">Delete</td>
            </td>   
            <%
            bitParity = 1
            do while NOT oRS.EOF
               inct = icnt + 1
               if bitParity = 1 then
                  strClass = "drowodd"
               else
                  strClass = "droweven"
               end if      
               strCodeName = ors("CodeName")
               %>
               <tr class="<%=strClass%>">
               <td class="dtabedit" align="center" ><a href="EditCode.asp?cd=<%=strCodename%>" > <i class="far fa-edit"></i></a></td> 
               <td><%=strCodeName%></td>
               <td class="dtabedit" align="center" ><a href="DeleteCode.asp?cd=<%=strCodename%>" > <i class="far fa-trash-alt"></i></a></td> 
               </tr>
               <%
               bitParity = NOT bitParity
			   ors.MoveNext
            loop  
            %>
            
            
            <tr>
             <td  colspan="3" class="dtabh">*** END ***</td>
            </tr> 
            </table>
            <%
            if icnt = 0 then
			      strMsg ="The Code Table has no entries" 
                  response.end
                  'response.redirect("Errorhandler.asp?ErrCode=NOTFound&Mod=EditCode&Msg="&strMsg)
            end if
			%> 
            
          
			
			   
       <% 	   
	   ors.Close
       set ors = Nothing
	   
	   %> 
              
       
            
        </div> 
      </div>   
        
	</div>
</body>
</html>
