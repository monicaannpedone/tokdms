<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--#include file="../include/validate.inc"-->
<!--#include file="../include/connect.inc"-->
<!--#include file="../include/locale.inc"-->
<!--#include file="../include/common.inc"-->

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<title>Edit Config Item</title>
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
	strUser = Session("username")
	if len(Request.Form("n")) = 0 then
	   strType   = "I"
	   strName   = Request.QueryString("N")
	   strErrMsg = Request.QueryString("ErrMsg")
	else
	   strType = "P"
	   strName = Request.Form("N")
	   strDesc = Request.Form("Desc")
	   strVal  = request.form("val")
	end if   	  
	
      
	%>
	<div id="header">
  
	</div>
	<div id="main">
		<div id="contents">
			<h2>Edit Configuration Item <%=strName%></h2>
            <h3></h3>
            <%			
			
            
			if strType = "I" then
			  strSQL = "SELECT * FROM dms_Config where cfg_Name = '"&strname&"';"
			  ors.Open strSQL
			  if not ors.EOF then
				  strDesc = trim(ors("cfg_desc"))
				  strVal  = trim(ors("cfg_value"))
			  end if	  
			  if len(strErrMsg) > 1 then
				 strE = left(strErrMsg,1) 
				 if strE = "N" then strErrMsgN = right(strErrMsg,len(strErrMsg)-1) end if
				 if strE = "D" then strErrMsgD = right(strErrMsg,len(strErrMsg)-1) end if
				 if strE = "V" then strErrMsgV = right(strErrMsg,len(strErrMsg)-1) end if
			  end if 
			
			%> 
            
            <Form action="ChgConf.asp" method="post" name="ChgCode" >
                  <input type="hidden" name="N" value="<%=strname%>" /> 
                  
            <table class="dtab">
               <tr>                  
                  <td class="dLabel" width=50>Name</td>
                  <td class="dvalue" width=300><%=strName%></td>
                  <td class="drowerror" width=400><%=strErrMsgN%></td>
               </tr> 
               <tr>                  
                  <td class="dLabel" width=50>Description</td>
                  <td class="drowedit" width=300><input name="Desc" type="text"  class="entry" size="50"  width="100" maxlength="100" value="<%=strDesc%>" /></td>
                  <td class="drowerror" width=400><%=strErrMsgD%></td>
               </tr> 
               <tr>                  
                  <td class="dLabel" width=50>Value</td>
                  <td class="drowedit" width=300><input name="Val" type="text"  class="entry" size="50"  width="100" maxlength="100" value="<%=strVal%>" /></td>
                  <td class="drowerror" width=400 ><%=strErrMsgV%></td>
               </tr> 
               <tr>                  
                  
                  <td class="drowedit" width=50>&nbsp;</td>
                  <td class="drowedit" width=300 align="left"><input  type="submit" value="Change"></td>
                  <td class="drownone" width=400 align="right"><input  type="Reset" value="Reset"></td>
               </tr> 
			</table>
            </Form>
			
			   
       <% 	   
	      else
			  if len(trim(strDESC)) = 0 then response.redirect("ChgConf.asp?N="&strName&"&ErrMsg=DName can not be blank") end if
			  if len(trim(strVal)) = 0 then response.redirect("ChgConf.asp?N="&strName&"&ErrMsg=VValue can not be blank") end if
		      strSQL = "SELECT * FROM dms_Config where cfg_Name = '"&strname&"';"
			   ors.Open strSQL
			   if NOT ors.EOF then
			      datNow = Now()
			      strSQL = "UPDATE dms_Config SET cfg_LastAccess='"&datNow&"' , cfg_User='"&strUser&"', cfg_LastAct='C', cfg_IsHistory=0, cfg_name='"&strname&"', "
				  strSQL = strSQL & "cfg_Desc ='"&strDesc&"', cfg_value='"&strval&"' WHERE cfg_name = '"&strname&"';"
				  response.write("strSQL=>"&strSQL&"< <br>")
				  oconn.Execute strSQL
				  response.redirect("MaintConf.asp?msg=Configuration item "&strname&" was successfully Changed")
			   else
			      response.redirect("ChgConf.asp?N="&strName&"&ErrMsg=NThe name "&strName&" Does Not Exist!")

			   end if
		  
	      end if
	   %> 
              
       
            
        </div> 
        
	</div>
</body>
</html>
