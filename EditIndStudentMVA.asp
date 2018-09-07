<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<script SRC="calendar.js"></script>

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<head>
	<title>Edit Student</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<style type="text/css" media="screen">@import "../css/basic.css";</style>
	<style type="text/css" media="screen">@import "../css/tabs.css";</style>

</head>

<body>

<div >
   <div >
	  <br/><br/>
		    <h2>Edit Student Attendance Form</h2>
			<% 
			Dim aryErrmsg(100)
			
			strMode = LCase(Request.Querystring("m"))
			'response.write("Mode is >"&strMode&"<</br>")
			strErrMsg   = Request.QueryString("err")
			strStudentID = Request.Querystring("u")
			'response.write("Request.Querystring(u)=>"&strStudentID&"<</br>")
			
			'response.write("strErrMsg=>"&strErrMsg&"<<br>")
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
' *  
' **        Setup default variables
' *                 
	        strEnv              = ""
			strKaiCd           = "abk"
			strdojoCD  	        = Session("DojoCD")
			strFNname           = ""
			strLName            = ""
			strMName	        = ""
			bitIsInstructor     = False
			bitIsActive         = True
			bitIsHeadInstructor = False
			bitIsFistMember     = False
			bitIsHelper         = False
			bitIsAssistant      = False
			bitIsOnWeapons      = False
			datStartDate        = Now()
			datEndDate          = Null
			datEndDate          = Null
			strGiSizeCd         = ""
			strBeltSizeCd       = ""
			strDivision         = "1"
			strAnchorDay        = ""
			strRankCd           = "WB"
			strEmail            = ""
			strPhone            = ""
			
			
	'		If strMode = "a" then
	'		     'response.write("Add </br>")
	'			   strSQL = "SELECT Max(StudentId) AS MAXSTUDENTID FROM dms_student;"
     '             oRS.Open strSQL
	'		      bitNotFound = oRS.EOF
	'		      if oRS.EOF then
	'		         strMsg="StudentID "&strStudentID&" Not Found"
	'		      else
	'	             intMaxStudentID = ors("MAXSTUDENTID")
	'	             intNextStudentID = intMaxStudentID + 1
	'	             intStudentID = intNextStudentID
	'		         'response.write("Student Read Found "&intNextStudentID)   
	'		      end if
     '       end if      
              strSQL = "Select * from dms_Student WHERE StudentID = "&strStudentID&";"
              ors.open strSQL
              if ors.eof then
                 response.end("That Student ID is not found >"&strStudentID&"<<br>")
                 response.end
              end if    
	          strEnv              = ""
			  strKaiCd            = ors("KaiCd")
			  strdojoCD  	      = ors("DojoCD")
			  strFName            = TRIM(ors("FName"))
			  strLName            = Trim(ors("LName"))
			  strMName	          = Trim(ors("MName"))
			  bitIsInstructor     = ors("IsInstructor")
			  bitIsActive         = ors("IsActive")
			  bitIsHeadInstructor = ors("IsHeadInstructor")
			  bitIsFistMember     = ors("IsFistMember")
			  bitIsHelper         = ors("IsHelper")
			  bitIsAssistant      = ors("IsAssistant")
			  bitIsOnWeapons      = ors("IsOnWeapons")
			  datStartDate        = ors("StartDate")
			  datEndDate          = ors("EndDate")
			  datDOB              = ors("DOB")
			  strGiSizeCd         = ors("GiSizeCd")
			  strBeltSizeCd       = ors("BeltSizeCd")
			  strDivision         = ors("Division")
			  
			
			  strBeltSizeCode00Selected = ""  
			  strBeltSizeCode0Selected = ""  
			  strBeltSizeCode1Selected = ""  
			  strBeltSizeCode2Selected = ""  
			  strBeltSizeCode3Selected = ""  
			  strBeltSizeCode4Selected = ""  
			  strBeltSizeCode5Selected = ""  
			  strBeltSizeCode6Selected = ""  
			  strBeltSizeCode7Selected = ""  
			  strBeltSizeCode8Selected = ""  
			  strBeltSizeCode9Selected = ""  
			  strBeltSizeCodeUnkSelected = ""  
			  
		      strBeltSizeCd = Trim(strBeltSizeCd)
		     ' response.write("Belt Size Code=>"&strBeltSizeCd&"<<br>")
			  Select Case strBeltSizeCd 
			         Case "00"
			              strBeltSizeCode00Selected = "Selected"
			         Case "0"
			              strBeltSizeCode0Selected = "Selected"
			         Case "1"
			              strBeltSizeCode1Selected = "Selected"
			         Case "2"
			              strBeltSizeCode2Selected = "Selected"
			         Case "3"
			              strBeltSizeCode3Selected = "Selected"
			         Case "4"
			              strBeltSizeCode4Selected = "Selected"
			         Case "5"
			              strBeltSizeCode5Selected = "Selected"
			         Case "6"
			              strBeltSizeCode6Selected = "Selected"
			         Case "7"
			              strBeltSizeCode7Selected = "Selected"
			         Case "8"
			              strBeltSizeCode8Selected = "Selected"
			         Case "9"
			              strBeltSizeCode9Selected = "Selected"
			         Case Else
			              strBeltSizeCodeUnkSelected = "Selected"
			  End Select

             ' response.write("Belt Code Size Unk Selected=>"&strBeltSizeCodeUnkSelected&"<<br>")
             ' response.write("Belt Code Size 00 Selected=>"&strBeltSizeCode00Selected&"<<br>")
             ' response.write("Belt Code Size 0 Selected=>"&strBeltSizeCode0Selected&"<<br>")
             ' response.write("Belt Code Size 1 Selected=>"&strBeltSizeCode1Selected&"<<br>")
             ' response.write("Belt Code Size 2 Selected=>"&strBeltSizeCode2Selected&"<<br>")
             ' response.write("Belt Code Size 3 Selected=>"&strBeltSizeCode3Selected&"<<br>")
             ' response.write("Belt Code Size 4 Selected=>"&strBeltSizeCode4Selected&"<<br>")
             ' response.write("Belt Code Size 5 Selected=>"&strBeltSizeCode5Selected&"<<br>")
             ' response.write("Belt Code Size 6 Selected=>"&strBeltSizeCode6Selected&"<<br>")
             ' response.write("Belt Code Size 7 Selected=>"&strBeltSizeCode7Selected&"<<br>")
             ' response.write("Belt Code Size 8 Selected=>"&strBeltSizeCode8Selected&"<<br>")
             ' response.write("Belt Code Size 9 Selected=>"&strBeltSizeCode9Selected&"<<br>")
			  
			  if strDivision = "1" then strDivision1Selected = "Selected" else strDivision1Selected = "" end if
			  if strDivision = "2" then strDivision2Selected = "Selected" else strDivision2Selected = "" end if
			  if strDivision = "3" then strDivision3Selected = "Selected" else strDivision3Selected = "" end if
			  if strDivision = "4" then strDivision4Selected = "Selected" else strDivision4Selected = "" end if
			  strAnchorDay        = Trim(ors("AnchorDay"))
		'	  response.write("anchor day=>"&strAnchorDay&"<<br>")
			  if strAnchorDay = "S" then strAnchordaySSelected = "Selected" else strAnchorDaySSelected = "" end if
			  if strAnchorDay = "M" then strAnchordayMSelected = "Selected" else strAnchorDayMSelected = "" end if
			  if strAnchorDay = "T" then strAnchordayTSelected = "Selected" else strAnchorDayTSelected = "" end if
			  if strAnchorDay = "W" then strAnchordayWSelected = "Selected" else strAnchorDayWSelected = "" end if
			  if strAnchorDay = "H" then strAnchordayHSelected = "Selected" else strAnchorDayHSelected = "" end if
			  if strAnchorDay = "F" then strAnchordayFSelected = "Selected" else strAnchorDayFSelected = "" end if
			  if strAnchorDay = "R" then strAnchordayRSelected = "Selected" else strAnchorDayRSelected = "" end if
			  if strAnchorDay = "X" then strAnchordayXSelected = "Selected" else strAnchorDayXSelected = "" end if
		'	  response.write("strAnchordaySSelected=>"&strAnchordaySSelected&"<<br>")
		'	  response.write("strAnchordayMSelected=>"&strAnchordayMSelected&"<<br>")
		'	  response.write("strAnchordayTSelected=>"&strAnchordayTSelected&"<<br>")
		'	  response.write("strAnchordayWSelected=>"&strAnchordayWSelected&"<<br>")
		'	  response.write("strAnchordayHSelected=>"&strAnchordayHSelected&"<<br>")
		'	  response.write("strAnchordayFSelected=>"&strAnchordayFSelected&"<<br>")
		'	  response.write("strAnchordayRSelected=>"&strAnchordayRSelected&"<<br>")
		'	  response.write("strAnchordayXSelected=>"&strAnchordayXSelected&"<<br>")
			  
			  strAuthSpar         = ors("AuthSpar")
              if strAuthSpar = True then
                 strAuthSparChecked = "CHECKED" 
              else   
                 strAuthSparChecked = ""
              end if   
              strAuthGrapple         = ors("AuthGrapple")
              if strAuthGrapple = True then
                 strAuthGrappleChecked = "CHECKED" 
              else   
                 strAuthGrappleChecked = ""
              end if                 
			  strEmail            = ors("Email")
			  strPhone            = ors("Phone")
			  strAreaCode         = left(strPhone,3)
			  strExchange         = mid(strPhone,4,3)
			  strLineNo           = right(strPhone,4)
			  
			'  strPortrait         = "P"&right("000000"&strStudentID,7)&"-01-"&trim(LCase(strLName))&"-"&trim(LCase(strFName))&".jpg"
			'  strAltPortrait      = strLName&", "&strFName
			  RankCd              = ors("RankCd")
			 ' response.write("3 First name=>"&strFName&"<</br>")        
  			  'response.write("3 Last  name=>"&strLName&"<</br>") 
			 ' response.write("3 Portrait  =>"&strPortrait&"<</br>")
                           
   '           for each fld in ors.Fields
	'			 'response.write(fld.Name&"=>"&fld.Value&"<</br>")
	'			 if NOT fld.name="ChangeTimeStamp" then
	'			   ' response.write(fld.Name&"=>"&fld.Value&"<</br>")
	'			 else
	'			 end if
     '         next
			'ors.Close
		    'set ors = Nothing
		    'response.write(" 2 First name=>"&strFName&"<</br>")        
   
			strStudentName=Trim(strFName)&" "&Trim(strLname)  	  
			%>
 <Table>
   <form action="EditIndStudentMVA1.asp" method="post" >
               <tr >
                  <td class="dtabh">Active</td>
                  <td class="dtabh">Last Name</td>
                  <td class="dtabh">First Name</td>
                  <td class="dtabh">Division</td>
                  <td class="dtabh">Anchor Day</td>
                  <td class="dtabh">DOB</td>
                  <td class="dtabh">Belt Size</td>
                  <td class="dtabh">Spar</td>
                  <td class="dtabh">Grapple</td>
                 
                  <td class="dtabh" colspan="3">Best Phone</td>
                  
               </tr>   		
			<% 
			bitflip=0
			intCnt = 0
			strChoose = ""
			if strSelActive and strSelInactive  then
			   strChoose = ""
			else
			   if strSelactive  then   
			      strChoose = " AND IsActive = 1 "
			   end if   
			   if strSelinActive then
			      strChoose = "AND IsActive = 0 "
			   end if   
			end if
			'response.write("StrSelActive>"&strSelActive&"<</br>")
			'response.write("StrSelInActive>"&strSelInActive&"<</br>")
			
			'response.write("strChoose>"&strChoose&"<</br>")
	   
			%>
			 <tr class="<%=strClass%>">
			      <td class="dtabedit"><input name="DojoCd" type="hidden" value="<%=strDojoCD%>"><input name="StudentID" type="hidden" value="<%=strStudentID%>">Active</td>
                  <td class="dtabedit"> <input type="text" name="lname" value="<%=strLname%>"></td>
                  <td class="dtabedit"> <input type="text" name="fname" value="<%=strfname%>"></td>

                  <td class="dtabedit">
                    <select name="Division">
                        <option value="1" <%=strDivision1Selected%> >1</option>
                        <option value="2" <%=strDivision2Selected%> >2</option>
                        <option value="3" <%=strDivision3Selected%> >3</option>
                        <option value="4" <%=strDivision4Selected%> >4</option>
                    </select> 
                  </td>
                  <td>                   
                    <select name="AnchorDay">
                        <option value="S" <%=strAnchorDaySSelected%> >Sunday</option>
                        <option value="M" <%=strAnchorDayMSelected%> >Monday</option>
                        <option value="T" <%=strAnchorDayTSelected%> >Tuesday</option>
                        <option value="W" <%=strAnchorDayWSelected%> >Wednesday</option>
                        <option value="H" <%=strAnchorDayHSelected%> >Thursday</option>
                        <option value="F" <%=strAnchorDayFSelected%> >Friday</option>
                        <option value="R" <%=strAnchorDayRSelected%> >Saturday</option>
                        <option value="X" <%=strAnchorDayXSelected%> >Floater</option>
                    </select>   
                  <td class="dtabedit"><input type="text" name="DOB" value="<%=datDOB%>"</td>
                  <td class="dtabedit">
                    <select name="BeltSizeCd">
                        <option value="Unk" <%=strBeltSizeCodeUnkSelected%> >Unk</option>
                        <option value="00" <%=strBeltSizeCode00Selected%> >00</option>
                        <option value="0"  <%=strBeltSizeCode0Selected%> >0</option>
                        <option value="1"  <%=strBeltSizeCode1Selected%> >1</option>
                        <option value="2"  <%=strBeltSizeCode2Selected%> >2</option>
                        <option value="3"  <%=strBeltSizeCode3Selected%> >3</option>
                        <option value="4"  <%=strBeltSizeCode4Selected%> >4</option>
                        <option value="5"  <%=strBeltSizeCode5Selected%> >5</option>
                        <option value="6"  <%=strBeltSizeCode6Selected%> >6</option>
                        <option value="7"  <%=strBeltSizeCode7Selected%> >7</option>
                        <option value="8"  <%=strBeltSizeCode8Selected%> >8</option>
                        <option value="9"  <%=strBeltSizeCode9Selected%> >9</option>
                    </select>   
                  
                  </td>
                  <td class="dtabedit" align="center"><input type="checkbox" name="Authspar" Value = "Y" <%=strAuthsparchecked%>></td>
                  <td class="dtabedit" align="center"><input type="checkbox" name="Authgrapple" value="Y" <%=strAuthgrapplechecked%>></td>
                  
                  <td class="dtabedit">(<input type="text" name="AreaCode" value="<%=strAreaCode%>" size="3" maxlength="3">)&nbsp;</td> 
                  <td class="dtabedit"><input type="text"  name="Exchange" value="<%=strExchange%>" size="3" maxlength="3">-</td> 
                  <td class="dtabedit"><input type="text"  name="LineNo"   value="<%=strLineNo%>"   size="4" maxlength="4"></td> 
             <tr>
                  <td class="dcolerror" colspan=13">&nbsp;<%=strErrMsg%></td>

             </tr>
                
             <tr>
                 <td colspan="6" align="left">
                <input type="Reset" onclick="window.location.href="MonthViewAttendance.asp";")> 
                </td>

                <td colspan="6" align="right">
                <input type="submit" value="Change"  > 
                </td>
			 </tr>    
            <tr>
                </form>
			  </TABLE>
				  
				  
		</div>
	</div>
</body>
</html>
