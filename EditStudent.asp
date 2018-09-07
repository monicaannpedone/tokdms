<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<script SRC="calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<head>
	<title>Maintain Users</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>

</head>

<body>

<div >
   <div >
	  <br/><br/>
		
        <TABLE  Border=0 Cellspacing=2 Cellpadding=2>
        <form action="EditStudent1.asp" method="post" name="EditStudent1"  >
			<% 
			Dim aryErrmsg(9)
			strMode = Request.Querystring("m")
			strErrMsg   = Request.QueryString("err")
			strStudentID = Request.Querystring("u")
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
			
			
			
			strSQL = "SELECT * FROM dms_student WHERE StudentID = "&strStudentID&";"
			oRS.Open strSQL
			bitNotFound = oRS.EOF
			if oRS.EOF then
			   strMsg="StudentID "&strStudentID&" Not Found"
			   response.redirect("MaintUsers.asp?msg="&strMsg)
			end if
			intCnt = 0
			strEnv = ""
			strKainum         = ors("KaiCD")
			strdojoCD  	      = ors("dojoCD")
			strStudentFNname  = ors("Fname")
			strStudentLName   = ors("Lname")
			strStudentMName	  = oRS("MName")
			
			
			ors.Close
		    set ors = Nothing
			   
			   
			   
			strStudentName=Trim(strStudentFName)&" "&Trim(strStudentLname)  	  
			%>
             <h1>Edit User <%=strStudentname%></h1>
               <input type="hidden" name="u" value="<%=strUsername%>" />
			 <tr>
			   <td class="dLabel1">Studentname</td>
			   <td class="entry1"><%=strStudentName%></td>			   
               <td class="drowerror1"><%=aryErrmsg(00)%></td>   
             </tr> 
             <tr>
			   <td class="dLabel1">Photo</td>
			   <td class="entry1"><img src=<%=strStudentID%>></td>			   
               <td class="drowerror1"><%=aryErrmsg(00)%></td>   
             </tr>
             <tr>
			   <td class="dLabel1">First Name</td>
			   <td class="entry1"><Input Name="FName"  type="text" value="<%=strStudentFName%>" size="32" maxlength="32"  /></td>			   
               <td class="drowerror1"><%=aryErrmsg(01)%></td>   
             </tr>
             <tr>
			   <td class="dLabel1">Last Name</td>
			   <td class="entry1"><Input Name="LName"  type="text" value="<%=strStudentLName%>" size="32" maxlength="32"  /></td>			   
               <td class="drowerror1"><%=aryErrmsg(02)%></td>   
             </tr>              
             <tr>
			   <td class="dLabel1">Email</td>
			   <td class="entry1"><input name="Email"  type="text" value="<%=strEmail%>" size="50" maxlength="50"  /></td>
			   <td class="drowerror1"><%=aryErrmsg(03)%></td>   
             </tr> 
             <tr>
			   <td class="dLabel1">Phone</td>
			   <td class="entry1"><Input Name="Phone"  type="text" value="<%=strPhoneD%>" size="32" maxlength="32"  /></td>			   
               <td class="drowerror1"><%=aryErrmsg(04)%></td>   
             </tr>
             <tr>
			   <td class="dLabel1">End Date</td>
			   <td class="entry1"><Input Name="ENDDATE"  type="text" value="<%=datEndDate%>" size="12" maxlength="12"  />&nbsp;
               <a href="javascript:ShowCalendar('document.EditUser1.ENDDATE', document.EditUser1.ENDDATE.value);"><img src="calendar.gif"  align="absmiddle" width="18" height="18" border="0" alt="Click Here to Pick up a Date" /> &nbsp;Select a date</a></td>			   
               <td class="drowerror1"><%=aryErrmsg(05)%></td>   
             </tr> 
             
            
             <tr>
			   <td class="dLabel1" valign="top">Comment</td>
			   <td class="entry1"><textarea name="Comment" rows="4"  wrap="soft" cols="64"  ><%=strComment%></textarea></td>			   
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
