<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->

<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/common.inc" -->

<head>
	<title>List Users</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>
	
</head>

<body>
<div id="main">
   <% Call TopLine() %>

		<div id="contents">
	  <h2>Assign Gender </h2>
      <a href="Admin.asp">Return to Admin Menu</a>
      <br/><br/>
               <TABLE Border=0 Cellspacing=2 Cellpadding=2>
               <tr>
                  
                  <form id="changedojocode" action="ChangeDojoCode.asp?mn=2" method="post">
                  <%
                   intOrder = Request.QueryString("o")
			                        
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
                   <td><b><%=StrDojoShortnameB%></td><td></td><td></td><td></td> <td></td><td></td><td></td></b><td></td><td>&nbsp;</td><td>&nbsp;</td> <td>&nbsp;</td><td></td><td></td> <td> <% if bitCanChangeDojo = True then %> Hold to Select:<% end if %></td>
                   <td>
                   <%
                   
                  if bitCanChangeDojo = True then
                     %><select name="DojoCd"  onchange="document.getElementById('changedojocode').submit();" ><%
                 
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

strSQL = "SELECT StudentId, dojoCd, IsActive, LName, FName, MName, Division, GenCd FROM  dbo.dms_Student"
strSQL = strSQL&" WHERE kaicd = '"&strkaicd&"' AND dojocd='"&strdojocdA&"' "&strChoose&" " 
strSQL = strSQL &"  ORDER BY DojoCd, Division, LName, FName, MName ;"
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
      
      		
<iframe id="MyFrame" name="MyFrame" style="display:none;"></iframe>
            <TABLE  Border='0' Cellspacing='2' Cellpadding='2'>
            <tr style="background-color:navy; color:white;" >
               <td class="dtabh">Dojo</td>
               <td class="dtabh">Division</td>
               <td class="dtabh" align="center">ID</td>
               <td class="dtabh">Last Name</td>
               <td class="dtabh">First Name</td>
               <td class="dtabh" align="center">MName</td>
               <td class="dtabh" align="center">Unk</td>
               <td class="dtabh" align="center">F</td>
               <td class="dtabh" align="center">M</td>
               <td class="dtabh" align="center">GNC</td>
            
            </tr>   		
			<% 
			strMsg = Request.QueryString("msg")
			
		
		
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
			   strdojocd    = ors("dojocd")
			   intStudentID = ors("StudentID")
			   strStudentID = CStr(intStudentID)
			   strLName     = ors("LName")
			   strFName     = ors("FName")
			   strMName     = ors("MName")
			   if len(strMName) = 0 then strMName = "                  " end if
			   strDivision  = ors("Division")
			   strGenCd     = ors("GenCd")
			   strGenCdCX = ""
			   strGenCdCF = ""
			   strGenCdCM = ""
			   strGenCdCN = ""
			   if strGenCd = "X" OR strGenCd = "" then strGenCdCX = "CHECKED" end if
			   if strGenCd = "F" then strGenCdCF = "CHECKED" end if
			   if strGenCd = "M" then strGenCdCM = "CHECKED" end if
			   if strGenCd = "N" then strGenCdCN = "CHECKED" end if
			   strButtonName="GenCd"&trim(strStudentID)
			   strID = "S"&strStudentID
			%>
			 <tr id="<%=strID%>" style="Background-color: <%=strBGColor%>; ">
			   <td class="<%=strClass%>" align="center"><%=strdojocd%></td>
			   <td class="<%=strClass%>" align="center"><%=strdivision%></td>
			   <td class="<%=strClass%>" align="center"> <%=strStudentID%></td>
			   <td class="<%=strClass%>"><%=strLName%></td>			   
               <td class="<%=strClass%>"><%=strFName%></td>  
               <td class="<%=strClass%>"><%=strMName%></td> 
            
               <form id="<%=strButtonName%>" action="ChangeGender.asp" method="post" target="MyFrame"> 
                   <input type="hidden" name="studentId" value="<%=strStudentID%>">
                  
         <!--          <td><input type="radio" name="<%=strButtonName%>" value="X" <%=strGenCdCX%>  onclick="window.open('ChangeGender.asp?u=<%=strStudentID%>&g=X', 'newwindow', 'width=400,height=200,scrollbars=no,menubar=no'); return false; "<%=strGendCd%>>Unk </td>  -->
                       <td><input type="radio" name="<%=strButtonName%>" value="X" <%=strGenCdCX%> onclick="document.getElementById('<%=strButtonName%>').submit();"<%=strGenCdX%>>Unk</td>          
                       <td><input type="radio" name="<%=strButtonName%>" value="F" <%=strGenCdCF%> onclick="document.getElementById('<%=strButtonName%>').submit();"<%=strGenCdF%>>F</td>          
                       <td><input type="radio" name="<%=strButtonName%>" value="M" <%=strGenCdCM%> onclick="document.getElementById('<%=strButtonName%>').submit();"<%=strGenCdM%>>M</td>          
                       <td><input type="radio" name="<%=strButtonName%>" value="N" <%=strGenCdCN%> onclick="document.getElementById('<%=strButtonName%>').submit();"<%=strGenCdN%>>GNC</td>          
          <!--         <td><input type="radio" name="<%=strButtonName%>" value="N" onclick="document.getElementById('<%=strButtonName%>').submit();"<%=strGenCdN%>>GNC</td>  -->
			   </form>
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
