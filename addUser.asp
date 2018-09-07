<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<script SRC="calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validateSysAdmin.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/common.inc" -->

<head>
	<title>Create a new user</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/balloon.css";</style>
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	<script defer src="../js/packs/solid.js"></script>
	<script defer src="../js/packs/light.js"></script>
	<script defer src="../js/packs/regular.js"></script>
    <script defer src="../js/fontawesome.js"></script>

</head>

<body>
<div >
   <div >	<%
	Call TopLine()
	bitIsAdmin = Session("IsAdmin")
    bitIsSysAdmin = Session("IsSysAdmin")
	bitLoggedIn = Session("loggedin")
	%>

	  <br/><br/>
		
        <TABLE  Border=0 Cellspacing=2 Cellpadding=2>
        <form action="AddUser1.asp" method="post" name="AddUser1"  >
			<% 
			Dim aryErrmsg(9)
			strUserName = Request.QueryString("u")
			strErrMsg   = Request.QueryString("err")
			for i = LBound(aryErrmsg) to UBound(aryErrmsg)
			   aryErrMsg(i) = ""
			next
			if Len(strErrMsg) > 3 Then
				if IsNumeric(left(strErrMsg,2)) then
				   intErrMsg = Cint(left(strErrmsg,2))
				   if intErrMsg < LBound(aryErrmsg) or intErrmsg > UBound(aryErrmsg) then 
					  intErrmsg = 0
					  aryErrMsg(0) ="Err Msg Number, "&intErrMsg&", default to zero, out of range, Lbound="&LBound(aryErrmsg)&"::UBound="&UBound(aryErrmsg)
				   else
				      intLenErrmsg = len(strErrMsg)
				      strErrmsg1 = Right(strErrMsg,intLenErrMsg-2)
				      aryErrMsg(intErrMsg) = strErrMsg1
				   end if
				else
				   intErrMsg = 00
				   aryErrMsg(0) = "Incorrect error number in message format, >"&strErrmsg&"<"
				end if
			else
			end if
			%>
            <h1>Create a new user</h1>
            <%
			
			strUsername = ""
        	bitIsActive = False
			bitIsBlocked  = False
			bitIsAdmin = False
			bitIsSysAdmin = False


			datCreatedate =Date()
			datEndDate = datCreateDate + 365 
			   
			strIsActiveC = ""
			if bitIsActive then strIsActiveC = "Checked" end if
			strIsAdminC       = ""
			if bitIsAdmin then strIsAdminC = "Checked" end if
			strIsSysAdminC    = ""
			if bitIsSysAdmin then strIsSysAdminC = "Checked" end if
			strIsBlockedC = ""
			if bitIsBlocked then strIsBlockedC = "Checked" end if
			strIsOp = ""
			if bitIsOp then strIsOp = "Checked" end if
			   
			if len(strPhone) > 0 then
			   strPhone = replace(strPhone," ","")
			   strPhone = replace(strPhone,"(","")
			   strPhone = replace(strPhone,")","")
			   strPhone = replace(strPhone,"-","")
			   strPhone = replace(strPhone,".","")
			   if len(strPhone) = 10 then
			      strPhoneD = "("&left(strPhone,3)&") "&mid(strPhone,4,3)&"-"&right(strPhone,4)
			   else
			      strPhoneD = strPhone
			   end if	 	 
			end if
			strDevC = ""
			if bitDev then strDevC = "Checked" end if
			strSandboxC = ""
			if bitSandbox then strsandboxC = "Checked" end if
			strTestC = ""
			if bitTest then strtestC = "Checked" end if
			strQAC = ""
			if bitQA then strQAC = "Checked" end if
			strProdimageC = ""
			if bitProdImage then strProdimageC = "Checked" end if
			strProdC = ""
			if bitProd then strProdC = "Checked" end if
			strComments = ""
			  	  
			%>
               <input type="hidden" name="u" value="<%=strUsername%>" />
			 <tr>
			   <td class="dLabel1">Username</td>
			   <td class="entry1"><Input Name="UserName"  type="text" value="<%=strUserName%>" size="32" maxlength="32"  /></td>			   
               <td class="drowerror1"><%=aryErrmsg(00)%></td>   
             </tr> 
             <tr>
			   <td class="dLabel1">First Name</td>
			   <td class="entry1"><Input Name="FName"  type="text" value="<%=strFName%>" size="32" maxlength="32"  /></td>			   
               <td class="drowerror1"><%=aryErrmsg(01)%></td>   
             </tr>
             <tr>
			   <td class="dLabel1">Last Name</td>
			   <td class="entry1"><Input Name="LName"  type="text" value="<%=strLName%>" size="32" maxlength="32"  /></td>			   
               <td class="drowerror1"><%=aryErrmsg(02)%></td>   
             </tr>              
             <tr>
			   <td class="dLabel1">Email</td>
			   <td class="entry1"><input name="Email"  type="text" value="<%=strEmail%>" size="50" maxlength="50" placeholder="yourname@gmail.com" /></td>
			   <td class="drowerror1"><%=aryErrmsg(03)%></td>   
             </tr> 
             <tr>
			   <td class="dLabel1">Phone</td>
			   <td class="entry1"><Input Name="Phone"  type="text" value="<%=strPhoneD%>" size="32" maxlength="32" placeholder="(212) 555-1212" /></td>			   
               <td class="drowerror1"><%=aryErrmsg(04)%></td>   
             </tr>
             <tr>
			   <td class="dLabel1">End Date</td>
			   <td class="entry1"><Input Name="ENDDATE"  type="text" value="<%=datEndDate%>" size="12" maxlength="12"  />&nbsp;
               <a href="javascript:ShowCalendar('document.AddUser1.ENDDATE', document.AddUser1.ENDDATE.value);"><img src="calendar.gif"  align="absmiddle" width="18" height="18" border="0" alt="Click Here to Pick up a Date" /> &nbsp;Select a date</a></td>			   
               <td class="drowerror1"><%=aryErrmsg(05)%></td>   
             </tr> 
             <tr>
			   <td class="dLabel1" valign="top">Attributes</td>
			   <td class="entry1">
                   <table class="dtab1">
                     <tr class="entry1">
                        <td class="entry1">Active</td>
                        <td class="entry1">Blocked</td>
                        <td class="entry1">Admin</td>
                        <td class="entry1">Sysadmin</td>
                     </tr>
                     
                     <tr>
                         <td class="entry1"><Input Name="IsActive"    type="Checkbox" <%=strIsActiveC%>   value="IsActive"  /></td>
                         <td class="entry1"><Input Name="IsBlocked"   type="Checkbox" <%=strIsBlockedC%>  value="IsBlocked"  /></td>
                         <td class="entry1"><Input Name="IsAdmin"     type="Checkbox" <%=strIsAdminC%>    value="IsAdmin"  /></td>
                         <td class="entry1"><Input Name="IsSysAdmin"  type="Checkbox" <%=strIsSysAdminC%> value="IsSysAdmin"  /></td>
                     </tr>
                   </table>
               </td>			   
               <td class="drowerror1"><%=aryErrmsg(06)%></td>   
             </tr>  
             <tr>
			   <td class="dLabel1" valign="top">Dojo</td>
			   <td class="entry1"><% Call SelectDojo("dojocd", "npk", 1) %></td>	
               <td class="drowerror1"><%=aryErrmsg(07)%></td>   
             </tr>  
             <tr>
			   <td class="dLabel1" valign="top">Can Change Dojo</td>
			   <td class="entry1"><Input Name="CanChangeDojo"    type="Checkbox"  value="1"  /></td>	
               <td class="drowerror1"><%=aryErrmsg(08)%></td>   
             </tr>  
          
             
  
            <%
'			loop
			%>
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
        <!--#include file="../include/infooter.inc"-->  
	</div>
    <!--#include file="../include/footer.inc"-->
</body>
</html>
