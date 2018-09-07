<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->

<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<head>
	<title>List Users</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
    <style type="text/css" media="screen">@import "../css/font-awesome.min.css";</style>  
	<script src="https://use.fontawesome.com/2ad953023c.js"></script> 
	
</head>

<body>
<div id="main">
		<div id="contents">
	  <h2>List  Student</h2>
      <a href="Student.asp">Return to Student Menu</a>
      <br/><br/>
             <TABLE  Border=0 Cellspacing=2 Cellpadding=2>
               <TABLE Border=0 Cellspacing=2 Cellpadding=2>
               <tr>
                  
                  <form id="changedojocode" action="ChangeDojoCode.asp?mn=1", method="post">
                  <%
                   intOrder = Request.QueryString("o")
 			if NOT IsNumeric(intOrder) then 
			   intOrder = 0
			end if   
			if intOrder > 0 then
			   strOrderType = " ASC "
			   strOrderSign = "-"
			else
			   if intOrder = 0 then
			      strOrderType = ""
			      strorderSign = ""
			   else
			      strOrderType = " DESC "
			      strorderSign = ""
			      intOrder = Abs(intorder)
			   end if
			end if   
			     
			   
			select case intOrder
			       case 0    'ID
			            strOrder = "dojocd, Division, dms_dowCD.doworder,  lname, fname"
			       case 1          ' ID
			            strOrder = "StudentID "&strOrderType&", lname, fname, dojocd, Division, dms_dowCD.doworder"
			       Case 2          ' Dojo
			            strOrder = "dojocd "&strOrderType&", Division, lname , fname,  dms_dowCD.doworder"
			       Case 3          ' Last Name
			            strOrder = "Lname "&strOrderType&", fname, dojocd, Division, dms_dowCD.doworder"
			       Case 4          ' First name
			            strOrder = "FName "&strOrderType&", LName, dojocd, Division ,   dms_dowCD.doworder"
			       Case 5          ' Division
			            strOrder = "Division "&strOrderType&", lname, fname, dojocd, dms_dowCD.doworder "
			       Case 6          ' Anchor
			            strOrder = "dms_dowCD.doworder "&strOrderType&", lname,fname, dojocd, Division "
			            
			       case Else
			            strOrder = "dojocd, Division, dms_dowCD.doworder,  lname, fname"
			end select 
			                        
                   strusername = Session("Username")
                   strSQLA = "SELECT DojoCd, CanChangeDojo, SelActive, SelInactive  FROM dms_pref WHERE username='"&strUserName&"';"
			       oRSA.Open strSQLA
                   strDojoCdA     = trim(orsA("DojoCd"))
                   strSelActive   = orsA("SelActive")
                   strSelInactive = orsA("SelInActive")
                   bitCanChangeDojo = orsA("CanChangeDojo")
                   'response.write("Can Change Dojo=>"&bitcanchangedojo&"<<br>")
                   
                   strSQLB = "SELECT DojoShortName FROM dms_dojo WHERE dojoCd = '"&strDojoCdA&"';"
                   'response.write("strSQLB=>"&strsqlB&"<</br>")
                   oRSB.Open strSQLB
                   strDojoShortNameB = orsB("DojoShortName")
                  %> 
                   <td><b><%=StrDojoShortnameB%><td></td><td></td><td></td> <td></td><td></td><td></td> </b><td></td><td>&nbsp;</td><td>&nbsp;</td> <td>&nbsp;</td><td></td><td></td> <td> <% if bitCanChangeDojo = True then %> Hold to Select:<% end if %></td>
                   <td>
                   <%
                   
                  if bitCanChangeDojo = True then
                     %><select name="DojoCd" onchange="document.getElementById('changedojocode').submit();" ><%
                     
                   strSQL = "SELECT *  FROM dms_dojo WHERE kaicd='"&strKaiCd&"' AND DojoIsActive = 1 ORDER BY dojoOrder ;"
                  ' response.write("strSQL=>"&strsql&"<</br>")
                   
			       oRS.Open strSQL
			       
			       do while not oRS.EOF 
			           strDojoCd        = TRIM(ors("Dojocd"))
			           strDojoShortName = TRIM(ors("DojoShortName"))
			           if strDojoCd = strDojoCdA then
			              strSelected = " SELECTED "
			           else
			              strSelected = ""
			           end if      
			       %>        
                         <option value="<%=strDojoCd%>" <%=strSELECTED%>><%=strDojoShortName%></option>
                   <%
                       ors.Movenext
			       loop 
			       ors.close
			      
			       Set oRS = Server.CreateObject("ADODB.Recordset")
                   oRS.ActiveConnection = oConn
			       %> 
                       </select> 
                   <%  
                   else
                      strDojoCd = strDojoCdA  
                   end if 
			       strChoose = ""
			       strActiveChecked   = ""
			       strInactiveChecked = ""
			      if strSelActive and strSelInactive  then
			         strChoose = ""
			         strActiveChecked  = "Checked"
			         strInactiveChecked = "Checked"
			      else
			         if strSelactive  then   
			            strChoose = " AND IsActive = 1 "
			            strActiveChecked  = "Checked"

			         end if   
			        if strSelinActive then
			           strChoose = "AND IsActive = 0 "
			           strInactiveChecked = "Checked"
			        end if   
			end if
			'response.write("StrSelActive>"&strSelActive&"<</br>")
			'response.write("StrSelInActive>"&strSelInActive&"<</br>")
			
			'response.write("strChoose>"&strChoose&"<</br>")
			strSQL = "Select * FROM dms_Student WHERE kaicd = '"&strkaicd&"' AND dojocd='"&strdojocdA&"' "&strChoose&" ORDER BY dojocd, Division, AnchorDay,  lname, fname;"
			'response.write("strSQL=>"&strSQL&"<</br>")

strSQL = "SELECT StudentId, dojoCd, IsActive, LName, FName, Division, AnchorDay FROM  dbo.dms_Student"
StrSQL = strSQL&"  Left Join dms_DowCD on dms_student.AnchorDay = dms_dowcd.dowcd "
strSQL = strSQL&" WHERE kaicd = '"&strkaicd&"' AND dojocd='"&strdojocdA&"' "&strChoose&" " 
strSQL = strSQL &"  ORDER BY "&strOrder&" ;"
Session("DojoCd") = strDojoCdA			
                   %>   
                   </td>
                   <td>&nbsp;</td>
                   <td><input type="checkbox" name="selactive"   value="active"   onclick="document.getElementById('changedojocode').submit();"  <%=strActiveChecked%>  >Active </td>
                   <td><input type="checkbox" name="selinactive" value="inactive" onclick="document.getElementById('changedojocode').submit();"<%=strInActiveChecked%>>Inactive</td>
                   <td></td>        
                  </form>
               </tr> 
               </br>  
               </TABLE>
     
      
      <%=strMsg%>
      
<a href="MaintStudent.asp?m=a" onclick="window.open('MaintStudent.asp?m=a', 'newwindow', 'width=1000,height=800,scrollbars=yes,menubar=yes'); return false; "> <i class="fa fa-Plus"></i>Add Student</a>
      		

            <TABLE  Border=0 Cellspacing=2 Cellpadding=2>
            <tr style="background-color:navy; color:white;" >
               <td class="dtabh"><a href="ListStudent.asp?o=<%=strOrderSign%>0">Reset</a></td>
               <td class="dtabh" align="center"><a href="ListStudent.asp?o=<%=strOrderSign%>1">ID</a></td>
               <td class="dtabh" align="center"><a href="ListStudent.asp?o=<%=strOrderSign%>2">Dojo</a></td>
               <td class="dtabh"><a href="ListStudent.asp?o=<%=strOrderSign%>3">Last Name</a></td>
               <td class="dtabh"><a href="ListStudent.asp?o=<%=strOrderSign%>4">First Name</a></td>
               <td class="dtabh" align="center"><a href="ListStudent.asp?o=<%=strOrderSign%>5">Division</a></td>
               <td class="dtabh" align="center"><a href="ListStudent.asp?o=<%=strOrderSign%>6">Anchor</a></td>
             
            </tr>   		
			<% 
			strMsg = Request.QueryString("msg")
		'	strSQL = "Select StudentID, dojocd, lname, fname, division, AnchorDay, RankCd FROM dms_Student order by "&strOrder&" ;"
			
		
		
		    oRS.Open strSQL
			
			bitflip=0
			intCnt = 0
			do while not oRS.EOF
			   intCnt = intCnt + 1
			   if bitflip = 1 then
			      strClass = "droweven"
				  bitflip = 0
			   else
			      bitflip = 1
			      strClass = "drowodd"	  
			   end if  
			   strEnv = ""
			   strdojocd = ors("dojocd")
			   intStudentID = ors("StudentID")
			   strStudentID = CStr(intStudentID)
			   
			   strLName     = ors("LName")
			   strFName     = ors("FName")
			   strDivision  = ors("Division")
			   intAnchorDay = ors("AnchorDay")
			   
			   
			%>
			 <tr style="Background-color: <%=strBGColor%>; ">
			   <td class="dtabnone" align="center" ><a href="MaintStudent.asp?u=<%=intStudentID%>&m=e" onclick="window.open('MaintStudent.asp?u=<%=intStudentID%>&m=e', 'newwindow', 'width=1000,height=800,menubar=yes,scrollbars=yes'); return false;"> <i class="fa fa-edit"></i></a></td> 
			   <td class="<%=strClass%>"align="center"><%=strStudentID%></td>

			   <td class="<%=strClass%>" align="center"><%=strdojocd%></td>
			   <td class="<%=strClass%>"><%=strLName%></td>			   
               <td class="<%=strClass%>"><%=strFName%></td>               
               <td class="<%=strClass%>" align="center"><%=strDivision%></td>               
               <td class="<%=strClass%>" align="center"><%=intAnchorDay%></td>               

			 </tr>    
            <%
			ors.Movenext
			loop
			%>
            <tr>
               <td class="dtabnone">&nbsp;</td>
               <td colspan="6" class="dtabh">&nbsp;<%=intCnt%> Students</td>
               <td class="dtabnone">&nbsp;</td>
            </tr>
            <tr>
               <td colspan="8" class="dtabnone">&nbsp;</td>
            </tr>
            <tr>
               <td class="dtabnone">&nbsp;</td>
               <td colspan="8"  align="left" valign="middle" class="dtabnone"></td>
            
           </tr>    
				  </TABLE>
                  <br/><br/>
                  

			<%
			
			
			
			ors.Close
		    set ors = Nothing
			%> 
		</div>
	</div>
</body>
</html>
