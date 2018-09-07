<!-- #include file ="connect.inc" -->
<!-- #include file ="locale.inc" -->
<%
'   ***********************************
'  *****  Common Routines          *****
' 
Function GetHostName()
   strSQLJ = "select HOST_NAME() as Hostname;"
   orsJ.Open strSQLJ
   GetHostname = orsJ("HostName")
   orsJ.Close
   set orsJ = Nothing
   Set oRSJ = Server.CreateObject("ADODB.Recordset") 
   oRSJ.ActiveConnection = oConn                     
End Function

Function GetDBName()
   strSQLJ = "select db_name() AS DBNM;"
   orsJ.Open strSQLJ
   GetDBname = orsJ("DBNM")
   orsJ.Close
   set orsJ = Nothing
   Set oRSJ = Server.CreateObject("ADODB.Recordset") 
   oRSJ.ActiveConnection = oConn                     
End Function

Function GetDojoName(strDojo)' use this to get the dojo full name
strSQLD += "SELECT DojoName from dms_Dojo where dojoCd = '"&strDojo&"';"
orsD.Open strSQLD
if NOT orsD.EOF then
   GetDojoName = orsD("DojoName")
  
else
   Call RecordException("DojoNFnd", "Dojo Name is not found for Code >"&strDojo&"< ", "COMMON", "E")   
end if 
orsD.Close
set orsD = Nothing
Set oRSD = Server.CreateObject("ADODB.Recordset") 
oRSD.ActiveConnection = oConn 
End Function

Function GetDojoShortName(strDojo)' use this to get the dojo full name
strSQLD = "SELECT DojoShortName from dms_Dojo where dojoCd = '"&strDojo&"';"
orsD.Open strSQLD
if NOT orsD.EOF then
   GetDojoShortName = orsD("DojoShortName")
else
   Call RecordException("DojoNFnS", "Dojo ShortName is not found for Code >"&strDojo&"< ", "COMMON", "E")   
end if 
orsD.Close
set orsD = Nothing
Set oRSD = Server.CreateObject("ADODB.Recordset") 
oRSD.ActiveConnection = oConn 
End Function

'
'

Function mMod(intN, intMod)
   mMod = intN - int(intN/intMod)*intMod
'  mod(16,12) = int(16/12) (1)*12=12  `6-12 =4 
End Function

'  
'
'
Function TopLine()
%>
    <div id="titlecontents">
        <div id="TitleBox">
            <h2><a href="Admin.asp"><i class="fa fa-bars"></i></a>&nbsp;&nbsp;<img src="../images/lillilinvertedtomoe.gif" />&nbsp;Dojo Management System &nbsp;&nbsp;&nbsp;&nbsp;<a href="index.asp"><i class="fa fa-home"></i></a></h2>
        </div>   
        <div id="AcctBox">
            <h4><a href="admin.asp"><%=strUserName%>&nbsp;<i class="fa fa-user"></i></a>&nbsp;&nbsp;&nbsp;<a href="Sysadmin.asp"><i class="fa fa-cog"></i></a>&nbsp;&nbsp;<a href="Logout.asp"><i class="fa fa-sign-out"></i></a></h4>
        </div>
    </div>  

<%
End Function

Function CheckBoxValue(bitB)
    if bitB  then
       CheckBoxValue = 1  
    else 
       CheckBoxValue = 0 
    end if   
End Function

Function SelectCheckbox(strName,strValue)
    if strValue then
       strIsChecked    = "CHECKED"
    else 
       strIsChecked    = ""
    end if   
%>  <input type="checkbox" name="<%=strName%>" value="1" <%=strIsChecked%>>
<%
End Function

Function SelectRadio(strName,strValue)
    if strValue then
       strIsChecked    = "CHECKED"
       strNotIsChecked = "" 
    else 
       strNotIsChecked = "CHECKED"  
       strIsChecked    = ""
    end if   
%>  <input type = "radio" name="<%=strName%>" value = "1" <%=strIsChecked%> >Active? &nbsp;
    <input type = "radio" name="<%=strName%>" value = "0" <%=strNotIsChecked%> >InActive?
<%
End Function

Function CheckboxDojo(strName, strDefaultDojoCd, bitSelectAll)
strSQL = "SELECT DojoCd, DojoOrder, DojoShortName FROM dms_dojo WHERE DojoIsActive=1 ORDER BY DojoOrder;"
orsE.open strSQL
if orsE.EOF then
   strErrMsg = "CMSSDEOF-F-EOF Occurred on dms_dojo" 
   Response.redirect("ErrorHandler.asp?errCode=CMSSDEOF&Mod=SelectFromTable&ErrMsg="&strErrMsg)
end if

if bitSelectAll = 1 then
   %><input type="checkbox" name="<%=strname%>" value="all">All&nbsp;<%
end if
do while NOT orsE.EOF
   strDOJOCD        = Trim(orsE("DojoCd"))
   strDojoShortName = Trim(orsE("DojoShortname"))
   if trim(lCase(strDefaultDojoCd)) = trim(lcase(strDojoCd)) then
      strSelected = "CHECKED"
   else
      strSelected = ""
   end if      
   %><input type="checkbox" name="<%=strname&strDojoCd%>" <%=strSelected%> value="<%=strDojoCd%>"><%=strDojoCd%>&nbsp;<%
orsE.MoveNext
Loop
orsE.close
set orsE = Nothing
Set orsE = Server.CreateObject("ADODB.Recordset") 
orsE.ActiveConnection = oConn 
End Function


Function SelectDojo(strName, strDefaultDojoCd, bitSelectAll)
%><select name="<%=strName%>"><% 
strSQL = "SELECT DojoCd, DojoOrder, DojoShortName FROM dms_dojo WHERE DojoIsActive=1 ORDER BY DojoOrder;"
orsE.open strSQL
if orsE.EOF then
   strErrMsg = "CMSSDEOF-F-EOF Occurred on dms_dojo" 
   Response.redirect("ErrorHandler.asp?errCode=CMSSDEOF&Mod=SelectFromTable&ErrMsg="&strErrMsg)
end if

if bitSelectAll = 1 then
   %><option value="all" >All Dojos</option><%
end if
do while NOT orsE.EOF
   strDOJOCD        = Trim(orsE("DojoCd"))
   strDojoShortName = Trim(orsE("DojoShortname"))
   if trim(lCase(strDefaultDojoCd)) = trim(lcase(strDojoCd)) then
      strSelected = "SELECTED"
   else
      strSelected = ""
   end if      
   %><option value="<%=strDojoCd%>" <%=strSelected%> ><%=strDojoShortName%></option><%
orsE.MoveNext
Loop
orsE.close
set orsE = Nothing
  Set orsE = Server.CreateObject("ADODB.Recordset") 
  orsE.ActiveConnection = oConn 
%> </select><%
End Function

Function GetNextRank(strDivision, strRank)
	intDivisionID = CInt(strDivision)
	intRankId     = CInt(strRank)
	intRankIDNext = intRankID + 1
	if strDivision =1 or strDivision = 2 or strDivision = 3 then
	   if intRankIdNext  > 14 then intRankID = 14 end if
	else
	   if intRankIdNext  > 24 then intRankIDNext = 24 end if
	end if
	GetNextRank = intRankIDNext
End Function

Function GetLowerRank(strDivision, strRank)
	intDivisionID = CInt(strDivision)
	intRankId     = CInt(strRank)
	intRankIDLower = intRankID - 1
	
	if intRankIdLower  < 1 then intRankIDLower = 0 end if
    GetRankLower = intRankIdLower
End Function

Function GetKata(strDivision, strRank)
    intDivisionID = CInt(strDivision)
    intstrRank    = CInt(strRank)
    strSQL = "SELECT DivisionID, RankID, Active, KataID from dms_KataRankDivTable where DivisionID=intDivisionID AND RankID=intRankID;"
    orsR.Open strSQL
    if orsR.EOF then
       response.write("*******ERROR***** Can not find a kata for Division=>"&strDivision&"< and Rank=>"&strRank&"< <br>")
      response.flush
        response.end
    else
       intKataID = orsr("KataID")   
    end if
    orsr.Close
    set orsr = Nothing
    Set oRSr = Server.CreateObject("ADODB.Recordset") 
    oRSr.ActiveConnection = oConn 
    GetKata = intKataID   
End Function

Function GetKataName(strKataID)
    intKataID    = CInt(strKataID)
    strSQL = "SELECT KataID, Kata, Active from dms_Kata where Active = 1 AND KataID="&intKataID&" ;"
    orsR.Open strSQL
    if orsR.EOF then
       strKataName="*ERR* "&strKataID
    else
       strKataName= orsr("Kata") 
    end if
    orsr.Close
    set orsr = Nothing
    Set oRSr = Server.CreateObject("ADODB.Recordset") 
    oRSr.ActiveConnection = oConn 
    GetKataName = strKataName  
End Function
Function SelectFromTableInd(strCodeName,strSelectValue, strVarName)
%><select name="<%=strCodeName&strVarName%>"><% 
strSQL = "SELECT CodeValue, CodeValueText FROM dms_Codes WHERE CodeName='"&strCodeName&"' AND ValueIsActive=1 ORDER BY ValueOrder ;"
orsZ.open strSQL
if orsZ.EOF then
   strErrMsg = "CMSFTEOF-F-EOF Occurred on dms_Codes for >"&strCodeName&"<" 
   Response.redirect("ErrorHandler.asp?errCode=CMSFTEOF&Mod=SelectFromTable&ErrMsg="&strErrMsg)
end if
do while NOT orsZ.EOF
   strCodeValue    = Trim(orsZ("CodeValue"))
   strCodeValueText = Trim(orsZ("CodeValueText"))
   if trim(lCase(strSelectValue)) = trim(lcase(strCodeValue)) then
      strSelected = "SELECTED"
   else
      strSelected = ""
   end if      
   %><option value="<%=strCodeValue%>" <%=strSelected%> ><%=strCodeValueText%></option><%
orsZ.MoveNext
Loop
orsZ.close
set orsZ = Nothing
  Set oRSZ = Server.CreateObject("ADODB.Recordset") 
  oRSZ.ActiveConnection = oConn 


%> </select><%
End Function

Function SelectFromTable(strCodeName,strSelectValue)
%><select name="<%=strCodeName%>"><% 
strSQL = "SELECT CodeValue, CodeValueText FROM dms_Codes WHERE CodeName='"&strCodeName&"' AND ValueIsActive=1 ORDER BY ValueOrder ;"
orsG.open strSQL
if orsG.EOF then
   strErrMsg = "CMSFTEOF-F-EOF Occurred on dms_Codes for >"&strCodeName&"<" 
   Response.redirect("ErrorHandler.asp?errCode=CMSFTEOF&Mod=SelectFromTable&ErrMsg="&strErrMsg)
end if
do while NOT orsG.EOF
   strCodeValue     = Trim(orsG("CodeValue"))
   strCodeValueText = Trim(orsG("CodeValueText"))
   if trim(lCase(strSelectValue)) = trim(lcase(strCodeValue)) then
      strSelected = "SELECTED"
   else
      strSelected = ""
   end if      
   %><option value="<%=strCodeValue%>" <%=strSelected%> ><%=strCodeValueText%></option><%
orsG.MoveNext
Loop
orsG.close
set orsG = Nothing
Set oRSG = Server.CreateObject("ADODB.Recordset") 
oRSG.ActiveConnection = oConn 


%> </select><%
End Function

Function bit2str(bitV)
         if bitV = 1 then
            bit2str = "Y"
         else
            bit2str = "N"
         end if
end function
Function str2bit(strV)
         if ucase(strV) = "Y" then
            str2bit = 1
         else
            str2bit = 0
         end if
end function
Function bit2Checked(bitV)
         if bitV = 1 then
            bit2Checked = "CHECKED"
         else
            bit2Checked = ""
         end if
end function

Function RecordException(strErrCode, strErrMsg, strMod, strSev)
  strUser = Session("UserName")
  Select Case(strSev)
         Case("I")
             strSevName = "Info::"
         Case("W")
             strSevName = "Warning!:"
         Case("E")
             strSevName = "ERROR*!:"
         Case("F")
             strSevName = "!FATAL!:"
         Case Else
              strSevName = "!FATAL!:"
  end Select
  strErrMsg = strSevName&strErrMsg
  
  strSQL = "INSERT INTO dms_Incident (IncidentDateTime, IncidentModule,     IncidentCode,     IncidentText)"
  strSQL = strSQL &         " VALUES ('"&now()&"',        '"&strMod&"', '"&strErrCode&"', '"&strErrMsg&"' );"

  oConn.Execute(strSQL)
  
  if strSev = "E" or strSev = "F" then
     strTitle = strSevName&" Incident user "&strUser&" encountered error "&strErrCode&" in Module "&strMod&" at "&Now()
     strTo = "postmaster@tokdms.com"
     strCC = ""
%>
     <h1><%=strSystemNameLong%></h1>
     <h2> Error Handler</h2>
     <h3> Error <%=strErrCode%></h3>
     <p>An error has occurred in module <%=strMod%></p>
     <p><%=strErrMsg%></p>	
     <p>A copy of this notice has been logged as an incident in the database and sent to the Sysadmin</p>	
<%    

     if Trim(Lcase(strLocale)) = "local" then
        response.write("----------------------------------------------------------------------------<br>")
        response.write("Local - simulating email<br>")
        response.write("Title:"&strTitle&"<br>")
        response.write("To:"&strTo&"<br>")
        response.write("CC:"&strCC&"<br>")
        response.write("Text:"&strErrMsg&"<br>")
        response.write("----------------------------------------------------------------------------<br>")
     else   
        DMSMail strTo,strCC,strTitle, strErrMsg 
     end if 
     response.write("<br><br>Please contact your System Administrator<br>")
     response.flush()
     response.end()
  end if   
End Function  


Function GetClassDayID(intClassID)
  strSQL="SELECT DayID from dms_Class where ClassID = "&intClassID&";"
  response.write("strSQL=>"&strSQL &"<<br>")
  orsZ.open strSQL
 
  if orsZ.EOF then
     strEM = "Class ID >"&intClassID&"< does not exist;<br>"
     response.write(strEM)
     response.end
  end if
  GetClassID = orsz("DayID")
  response.write(">"&GetClassID&"<=GetClassID("&intClassID&") <br>")
  oRSZ.Close
  set oRSZ=Nothing
  Set oRSZ = Server.CreateObject("ADODB.Recordset") 
  oRSZ.ActiveConnection = oConn 
End Function

'  
'
'
Function GetDayAbbrev(intDayIDX)
  response.write("DayID=>"&intDayIDX&"<<br>")
  strSQL="SELECT DowCd from dms_DOW WHERE DayID = "&intDayIDX&";"
  response.write("strSQL=>"&strSQL &"<<br>")
  orsZ.open strSQL
  if orsZ.EOF then
     strEM = "DayID >"&intDayIDX&"< does not exist;<br>"
     response.write(strEM)
     response.end
  end if
  GetDayAbbrev = orsz("DOWCd")
  response.write(">"&GetDayAbbrev&"<=GetDayAbbrev("&intDayIDX&") <br>")
  oRSZ.Close
  set oRSZ=Nothing
  Set oRSZ = Server.CreateObject("ADODB.Recordset") 
  oRSZ.ActiveConnection = oConn 
End Function


Function GetRankName(strRankCd)
   GetRankName = GetCodeValueText("RankCd", strRankCd)
End Function

'  *
'  ** Get Condition Values Order
'  *
Function GetValueOrder(strCodeName, strCodeValue)
    strSQL = "SELECT CodeValueText, ValueOrder FROM dms_Codes WHERE CodeName='"&strCodeName&"' AND  CodeValue='"&strCodeValue&"'; "
    'response.write("strSQL=>"&strSQL&"<<br>")
   ' response.end
     oRSL.Open strSQL
    if oRSL.EOF then
	   strEM = "That CodeName and Value, >"&strCodeName&"<, >"&strCodeValue&"< does not exist; "
	   response.flush
	   response.end
	end if
    GetValueOrder= orsL("ValueOrder")
    orsL.close
    set orsL = Nothing
    Set oRSL = Server.CreateObject("ADODB.Recordset") 
    oRSL.ActiveConnection = oConn 
End Function

'  *
'  ** Get Condition Values
'  *
Function GetCodeValueText(strCodeName, strCodeValue)
    strSQL = "SELECT * FROM dms_Codes WHERE ValueIsActive = 1; "
    if NOT (Session("CodesLoaded") = 1) then
       oRSL.Open strSQL
       intRowCount = orsL.RecCount
       Dim CodesCodeName(intRowCount)
       Dim CodesCodeValue(intRowCount)
       Dim CodesCodeValueText(intRowCount)
       Dim CodesCodeOrder(intRowCount)
       orsL.close
       set orsL = Nothing
       Set oRSL = Server.CreateObject("ADODB.Recordset") 
       oRSL.ActiveConnection = oConn 
       strSQL = "SELECT CodeValueText FROM dms_Codes WHERE CodeName='"&strCodeName&"' AND  CodeValue='"&strCodeValue&"'; "
       inti = 0
       do while not orsL.EOF
          CodesCodeName(intRowCount)      = orsL("CodeName")       
          CodesCodeValue(intRowCount)     = orsL("CodeValue")       
          CodesCodeValueText(intRowCount) = orsL("CodeValueText")       
          CodesValueOrder(intRowCount)    = orsL("ValueOrder")       
       loop
       orsL.close
       set orsL = Nothing
       Set oRSL = Server.CreateObject("ADODB.Recordset") 
       oRSL.ActiveConnection = oConn 
       Session("CodesLoaded") = 1
       Session("NumberofCodesLoaded") = intRowCount
    else
       intRowCount = Session("NumberofCodesLoaded") 
       bitMatch = 0
       intRowCount = 1
       do while intJ <= intRowCount
          if CodesCodeName(intJ) = strCodeName then
             if CodesCodeValue(intJ) = strCodeValue then
                bitMatch  = 1
                strResult = CodesCodeValueText(intJ)
             then
             end if
          end if
          intJ = intJ + 1
       loop
       GetCodeValueText = strResult
    end if   
   ' strSQL = "SELECT CodeValueText FROM dms_Codes WHERE CodeName='"&strCodeName&"' AND  CodeValue='"&strCodeValue&"'; "
    'response.write("strSQL=>"&strSQL&"<<br>")
   ' response.end
   '  oRSL.Open strSQL
  '  if oRSL.EOF then
     if bitMatch = 0 then
	   strEM = "That CodeName and Value, >"&strCodeName&"<, >"&strCodeValue&"< does not exist; "
	   response.flush
	   response.end
	end if
   ' GetCodeValueText = orsL("CodeValueText")
  '  orsL.close
   ' set orsL = Nothing
  '  Set oRSL = Server.CreateObject("ADODB.Recordset") 
  '  oRSL.ActiveConnection = oConn 

End Function
'  *
'  ** Lookup the Last Testing Date
'  *
Function GetLastTestingDate(intStudentID)
    strSQL = "SELECT Max(TestDate) as LatestTestingDate from dms_Testing WHERE StudentID="&intStudentID&";"
   ' response.write("strSQL=>"&strSQL&"<<br>")
    
     oRSX.Open strSQL
    if orsX.EOF then 
       GetLastTestingDate = "First"
    else   
       GetLastTestingDate = orsX("LatestTestingDate")
    end if   
    orsX.close
    set orsX = Nothing
    Set oRSX = Server.CreateObject("ADODB.Recordset") 
    oRSX.ActiveConnection = oConn 

End Function

'
' Leap Year
'
Function IsLeapYearDate(dtmTestDate)
    IsLeapYearDate = False
    If IsDate(dtmTestDate) Then
      Dim dtmTempDate
      dtmTempDate = "1/31/" & Year(dtmTestDate)
      dtmTempDate = DateAdd("m", 1, dtmTempDate)
      If Day(dtmTempDate) = 29 Then IsLeapYearDate = True
    End If
  End Function
  
'  *
'  **  CALCAGE
'  *
Function CalcAge(FromDate, ToDate)
'
'  ***  
'  **   This calculates an age from two dates
' ***
'  
  CalcAge=DateDiff("d", FromDate, ToDate)
  CalcAge = iNT(CalcAge/365*10)/10
   
End Function

'  *
'  **  DMSMail
'  *

Function DMSMail(strTo,strCC,strTitle, strText)
'  *
'  **     This routine sends mail from a specific email address and domain which needs to be changed if other than this one
'  ***    Get the credentials from the email manager on the site
'  **     It uses a Persits Mail object which may not be on a given site (such as local)
'  *
 
Set Mailer = Server.CreateObject("Persits.MailSender") 
 'response.write("To:>"&strTo&"<</br>")
 'response.write("CC:>"&strCC&"<</br>")
 'response.write("Title:>"&strTitle&"<</br>")
 'response.write("Text:>"&strText&"<</br>")
Mailer.Host     = "mail.tokdms.com" 
Mailer.Username = "postmaster@tokdms.com" 
Mailer.Password = "Sanchin1153!" 
Mailer.FromName = "postmaster@tokdms.com" 
Mailer.From     = "postmaster@tokdms.com" 
Mailer.AddAddress strTo
Mailer.AddCC  strCC
Mailer.Subject = strTitle
Mailer.Body    = strText 
 
' catch errors 
On Error Resume Next 
 
Mailer.Send 
 
If Err <> 0 Then ' error occurred 
      set DMSMail = Err 
      response.write("There was an email error <br>")
      response.flush
      response.end
else 
     'Response.Write "Mail sent..." 
	 set DMSSMail = 0
End If 
Set Mailer = Nothing
 
End Function



' *******************  End of Common Routines **********************************************
%>