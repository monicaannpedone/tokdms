<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/common.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<head>
	<title>Add Config Item</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	<script defer src="../js/packs/solid.js"></script>
    <script defer src="../js/packs/regular.js"></script>
    <script defer src="../js/packs/light.js"></script>
    <script defer src="../js/fontawesome.js"></script>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
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
			<h2>Add Configuration Item</h2>
            <h3></h3>
            <%			
			
            
			if strType = "I" then
			  strDesc = ""
			  strVal  = ""
			  if len(strErrMsg) > 1 then
				 strE = left(strErrMsg,1) 
				 if strE = "N" then strErrMsgN = right(strErrMsg,len(strErrMsg)-1) end if
				 if strE = "D" then strErrMsgD = right(strErrMsg,len(strErrMsg)-1) end if
				 if strE = "V" then strErrMsgV = right(strErrMsg,len(strErrMsg)-1) end if
			  end if 
			
			%> 
            
            <Form action="AddConf.asp" method="post" name="AddCode" >
                  <input type="hidden" name="Name" value="<%=strname%>" /> 
                  
            <table class="dtab">
               <tr>                  
                  <td class="dLabel" width=50>Name</td>
                  <td class="drowedit" width=300><input name="N" type="text"  class="entry" size="50"  width="100" maxlength="100" value="<%=strName%>" /></td>
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
                  <td class="drowedit" width=300 align="left"><input  type="submit" value="Add"></td>
                  <td class="drownone" width=400 align="right"><input  type="Reset" value="Reset"></td>
               </tr> 
			</table>
            </Form>
			
			   
       <% 	   
	      else
		      if len(trim(strDESC)) = 0 then response.redirect("AddConf.asp?N="&strName&"&ErrMsg=DDescription can not be blank") end if
			  if len(trim(strname)) = 0 then response.redirect("AddConf.asp?N="&strName&"&ErrMsg=NName can not be blank") end if
			  if len(trim(strVal)) = 0 then response.redirect("AddConf.asp?N="&strName&"&ErrMsg=VValue can not be blank") end if
		      strSQL = "SELECT * FROM dms_Config where cfg_Name = '"&strname&"';"
			   ors.Open strSQL
			   if ors.EOF then
			      datNow = Now()
			      strSQL = "INSERT INTO dms_Config (cfg_LastAccess, cfg_User, cfg_LastAct, cfg_IsHistory, cfg_name, cfg_Desc, cfg_Value) VALUES("
				  strSQL = strSQL & "'"&datNow&"',  '"&strUser&"', 'A', 0, '"&strname&"', '"&strDesc&"', '"&strval&"') ;"
				  'response.write("strSQL=>"&strSQL&"< <br>")
				  oconn.Execute strSQL
				  response.redirect("MaintConf.asp?msg=Configuration item "&strname&" was successfully added")
			   else
			      response.redirect("AddConf.asp?N="&strName&"&ErrMsg=NThe name "&strName&" already exists!")

			   end if
		  
	      end if
	   %> 
              
       
            
        </div> 
        
	</div>
</body>
</html>
