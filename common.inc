<!-- #include file ="connect.inc" -->
<!-- #include file ="locale.inc" -->
<%
'   ***********************************
'  *****  Common Routines          *****
'   ***********************************

Function Abend()
    forceabend =  1/0
End Function

Function GetDojoName(strDojo)' use this to get the dojo full name
strSQLD = "SELECT DojoName from dms_Dojo where dojoCd = '"&strDojo&"';"
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

Function GetDojoCd(intDojoNum)' use this to get the dojo full name
strSQLD = "SELECT DojoCd, DojoName from dms_Dojo where dojoorder = "&intDojoNum&";"
orsD.Open strSQLD
if NOT orsD.EOF then
   GetDojoCd = orsD("DojoCd")
else
   Call RecordException("DojoNFnd", "Dojo Code is not found for number >"&intDojoNum&"< ", "COMMON", "E")   
end if 
orsD.Close
set orsD = Nothing
Set oRSD = Server.CreateObject("ADODB.Recordset") 
oRSD.ActiveConnection = oConn 
End Function

Function GetDojoNum(strDojoCd)' use this to get the dojo full name
strSQLD = "SELECT dojoorder, DojoCd, DojoName from dms_Dojo where dojocd = '"&strDojocd&"';"
orsD.Open strSQLD
if NOT orsD.EOF then
   GetDojoNum = orsD("Dojoorder")
else
   Call RecordException("DojoNFnd", "Dojo Number is not found for DojoCd >"&strDojoCd&"< ", "COMMON", "E")   
end if 
orsD.Close
set orsD = Nothing
Set oRSD = Server.CreateObject("ADODB.Recordset") 
oRSD.ActiveConnection = oConn 
End Function

Function GetDojoID(strDojoCd)' use this to get the dojo full name
strSQLD = "SELECT dojoorder, DojoSchoolID, DojoCd, DojoName from dms_Dojo where dojocd = '"&strDojocd&"';"
orsD.Open strSQLD
if NOT orsD.EOF then
   GetDojoID = orsD("DojoSchoolID")
else
   Call RecordException("DojoIDFnd", "Dojo ID is not found for DojoCd >"&strDojoCd&"< ", "COMMON", "E")   
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
            <h2 data-balloon="Administrative Functions" data-balloon-pos="up"><a href="Admin.asp"><i class="fa fa-bars"></i></a>&nbsp;&nbsp;<img src="../images/lillilinvertedtomoe.gif" />&nbsp;Dojo Management System &nbsp;&nbsp;&nbsp;&nbsp;<a href="index.asp"><i class="fa fa-home"></i></a></h2>
        </div>   
        <div id="AcctBox">
            <h4  ><a href="EditUser.asp?u=<%=strUserName%>"><%=strUserName%>&nbsp;<i class="fa fa-user"></i></a>&nbsp;&nbsp;&nbsp;<a href="Sysadmin.asp"><i class="fa fa-cog"></i></a>&nbsp;&nbsp;<a href="Logout.asp"><i class="fa fa-sign-out"></i></a></h4>
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

Function SelectCheckboxLabel(strName,strValue,strLabel)
    if strValue = "1" or strValue = 1 or strvalue = True then
       strIsChecked    = "CHECKED"
    else 
       strIsChecked    = ""
    end if   
%>  <%=strLabel%> <input type="checkbox" name="<%=strName%>" value="1" <%=strIsChecked%>>
<%
End Function

Function SelectCheckbox(strName,strValue)
    if strValue = "1" or strValue = 1 or strvalue = True then
       strIsChecked    = "CHECKED"
    else 
       strIsChecked    = ""
    end if   
%>   <input type="checkbox" name="<%=strName%>" value="1" <%=strIsChecked%>>
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
    response.write("Div=>"&strDivision&"<; Rank=>"&strRank&"<<br>")
	intDivisionID = CInt(strDivision)
	intRankId     = CInt(strRank)
	strSQL = "SELECT RankID, KataID, DivisionID from dms_KataRankDivTable WHERE DivisionID = "&intDivisionID&" AND RankID = "&intRankID&" ORDER BY DivisionID, RankID;"
	orsR.Open STRSQL,oconn
    intRankIDNext = intRankID
    
    if orsR.Eof then
       If strRank = 0 then
          intRankIDNext = 1
       else   
          response.write("Can not find next rank for "&strRank&", for Division "&strDivision&"<br>")
          response.flush
	      response.end   
	   end if
	else
	   orsR.MoveNext
	   if NOT orsR.EOF then
	      intRankIDNext = orsR("RankID")
	   else
	      response.write("EOF after movenext<<br>")
	      
	   end if  
	   response.write("intRankIDNext=>"&intRankIDNext&"<<br>")
	end if
	GetNextRank = intRankIDNext
	orsR.Close
    set orsR = Nothing
    Set oRSR = Server.CreateObject("ADODB.Recordset") 
    oRSR.ActiveConnection = oConn 

End Function

Function GetLowerRank(strDivision, strRank)
	intDivisionID = CInt(strDivision)
	intRankId     = CInt(strRank)
	strSQL = "SELECT RankID, KataID, DivisionID from dms_KataRankDivTable WHERE DivisionID = "&inDivisionID&" AND RankID = "&intRankID&";" 
    orsR.Open STRSQL,Oconn,0
    intRankIDNext = intRankID
    if orsR.Eof then
       response.write("Can not find next rank for "&strRank&", for Division "&strDivision&"<br>")
       response.flush
	   response.end   
	else
	   orsR.MoveNext
	   if NOT orsR.EOF then
	      intRankIDNext = orsR("RankID")
	   end if   	
	end if
	GetLowerRank = intRankIDNext
	orsR.Close
    set orsR = Nothing
    Set oRSR = Server.CreateObject("ADODB.Recordset") 
    oRSR.ActiveConnection = oConn 
	
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

Function SelectFromTableIndJSRange(strCodeName,strSelectValue, strVarName,strJS, intSV, intEV)
%><select name="<%=strCodeName&strVarName%>" id="<%=strCodeName&strVarName%>" <%=strJS%>><% 
strSQL = "SELECT CodeValue, CodeValueText, ValueOrder FROM dms_Codes WHERE CodeName='"&strCodeName&"' AND ValueIsActive=1 AND (ValueOrder >= "&intSV&" AND ValueOrder <= "&intEV&" ) ORDER BY ValueOrder ;"
orsZ.open strSQL
if orsZ.EOF then
   strErrMsg = "CMSFTEOF-F-EOF Occurred on dms_Codes for >"&strCodeName&"<" 
   Response.redirect("ErrorHandler.asp?errCode=CMSFTEOF&Mod=SelectFromTable&ErrMsg="&strErrMsg)
end if
do while NOT orsZ.EOF
   strCodeValue                   = Trim(orsZ("CodeValue"))
   strCodeValueText               = Trim(orsZ("CodeValueText"))
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

Function SelectFromTableIndRange(strCodeName,strSelectValue, strVarName,intSV, intEV)
%><select name="<%=strCodeName&strVarName%>" id="<%=strCodeName&strVarName%>" ><% 
strSQL = "SELECT CodeValue, CodeValueText, ValueOrder FROM dms_Codes WHERE CodeName='"&strCodeName&"' AND ValueIsActive=1 AND (ValueOrder >= "&intSV&" AND ValueOrder <= "&intEV&" ) ORDER BY ValueOrder ;"
orsZ.open strSQL
if orsZ.EOF then
   strErrMsg = "CMSFTEOF-F-EOF Occurred on dms_Codes for >"&strCodeName&"<" 
   Response.redirect("ErrorHandler.asp?errCode=CMSFTEOF&Mod=SelectFromTable&ErrMsg="&strErrMsg)
end if
do while NOT orsZ.EOF
   strCodeValue                   = Trim(orsZ("CodeValue"))
   strCodeValueText               = Trim(orsZ("CodeValueText"))
   if trim(lCase(strSelectValue)) = trim(lcase(strCodeValue)) then
      strSelected = " SELECTED "
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

Function SelectFromTableIndJS(strCodeName,strSelectValue, strVarName,strJS)
%><select name="<%=strCodeName&strVarName%>" id="<%=strCodeName&strVarName%>" <%=strJS%>><% 
strSQL = "SELECT CodeValue, CodeValueText FROM dms_Codes WHERE CodeName='"&strCodeName&"' AND ValueIsActive=1 ORDER BY ValueOrder ;"
orsZ.open strSQL
if orsZ.EOF then
   strErrMsg = "CMSFTEOF-F-EOF Occurred on dms_Codes for >"&strCodeName&"<" 
   Response.redirect("ErrorHandler.asp?errCode=CMSFTEOF&Mod=SelectFromTable&ErrMsg="&strErrMsg)
end if
do while NOT orsZ.EOF
   strCodeValue                   = Trim(orsZ("CodeValue"))
   strCodeValueText               = Trim(orsZ("CodeValueText"))
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


'
Function SelectFromTableInd(strCodeName,strSelectValue, strVarName)
%><select name="<%=strCodeName&strVarName%>" id="<%=strCodeName&strVarName%>"><% 
strSQL = "SELECT CodeValue, CodeValueText FROM dms_Codes WHERE CodeName='"&strCodeName&"' AND ValueIsActive=1 ORDER BY ValueOrder ;"
orsZ.open strSQL
if orsZ.EOF then
   strErrMsg = "CMSFTEOF-F-EOF Occurred on dms_Codes for >"&strCodeName&"<" 
   Response.redirect("ErrorHandler.asp?errCode=CMSFTEOF&Mod=SelectFromTable&ErrMsg="&strErrMsg)
end if
do while NOT orsZ.EOF
   strCodeValue                   = Trim(orsZ("CodeValue"))
   strCodeValueText               = Trim(orsZ("CodeValueText"))
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
%><select name="<%=strCodeName%>" id="<%=strCodeName%>" ><% 
strSQL = "SELECT CodeValue, CodeValueText FROM dms_Codes WHERE CodeName='"&strCodeName&"'  AND ValueIsActive=1 ;"
orsG.open strSQL
if orsG.EOF then
   strErrMsg = "CMSFTEOF-F-EOF Occurred on dms_Codes for >"&strCodeName&"< Value=>"&strSelectValue&"  >>"&Right(strSQL,len(strSQL)-7)&"<<" 
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
         if bitV = 1 or bitV = "1" or bitV = True then
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
         if bitV = 1 or bitV = "1" or bitV = True then
            bit2Checked = "CHECKED"
         else
            bit2Checked = ""
         end if
end function

Function ChangeJournal(strTableName, strFieldname,  strModule, strAction)
  strUser = Session("UserName")
  strSQL = "INSERT INTO dms_ChangeJournal (JDateTime,        JUser,        JTableName,         JFieldName,          JModule,      JAction)"
  strSQL = strSQL &         " VALUES    ('"&now()&"','"&strUser&"', '"&strTablename&"', '"&strFieldName&"',  '"&strModule&"', '"&strAction&"' );"
  oConn.Execute(strSQL)
End Function  

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
     strSQL = "SELECT CodeValueText FROM dms_Codes WHERE CodeName='"&strCodeName&"' AND  CodeValue='"&strCodeValue&"'; "
    'response.write("strSQL=>"&strSQL&"<<br>")
   ' response.end
     oRSL.Open strSQL
     if orsL.EOF  then
	   strEM = "That CodeName and Value, >"&strCodeName&"<, >"&strCodeValue&"< does not exist; "
	   response.flush
	   response.end
	end if
    GetCodeValueText = orsL("CodeValueText")
    orsL.close
    set orsL = Nothing
    Set oRSL = Server.CreateObject("ADODB.Recordset") 
    oRSL.ActiveConnection = oConn 

End Function
'  *
'  ** Lookup the Last Testing Date
'  *
Function GetLastTestingDate(intStudentID)
    strSQL = "SELECT Max(TestDate) as LatestTestingDate from dms_Testing WHERE StudentID="&intStudentID&" AND StatusID = 1;"
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
  CalcAge = Round(CalcAge/365,1) 
   
End Function

Function DMSText(strTo, strCC, strText)
   intRC = DMSMail(strTo, strCC, "Text from DMS ", strText)
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
' response.write("To:>"&strTo&"<</br>")
' response.write("CC:>"&strCC&"<</br>")
' response.write("Title:>"&strTitle&"<</br>")
' response.write("Text:>"&strText&"<</br>")
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
      response.write("There was an email error " & Err.Description&" <br>")
      response.flush
      response.end
else 
     'Response.Write "Mail sent..." 
	 set DMSSMail = 0
End If 
Set Mailer = Nothing
 
End Function

Function SendEmail(strFromEmail, strFromName, strTo, strText, strHeading, strCC)
EmailSubject = strHeading
EmailBody    = strText

EmailFrom      = strFromEmail
EmailFromName  = strFromName
EmailTo        = strTo
Const SMTPServer = "smtp.gmail.com"
Const SMTPLogon = "tokdmssa@gmail.com"
Const SMTPPassword = "Sanchin1153!"
Const SMTPSSL = True
Const SMTPPort = 465

Const cdoSendUsingPickup = 1    'Send message using local SMTP service pickup directory.
Const cdoSendUsingPort = 2  'Send the message using SMTP over TCP/IP networking.

Const cdoAnonymous = 0  ' No authentication
Const cdoBasic = 1  ' BASIC clear text authentication
Const cdoNTLM = 2   ' NTLM, Microsoft proprietary authentication

' First, create the message

Set objMessage = CreateObject("CDO.Message")
objMessage.Subject = EmailSubject
objMessage.From = """" & EmailFromName & """ <" & EmailFrom & ">"
objMessage.To = EmailTo
objMessage.TextBody = EmailBody
objMessage.CC = strCC

' Second, configure the server

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPLogon

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPPassword

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = SMTPSSL

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

objMessage.Configuration.Fields.Update
'Now send the message!
On Error Resume Next
objMessage.Send

If Err.Number <> 0 Then
    MsgBox Err.Description,16,"Error Sending Mail"
Else 
    MsgBox "Mail was successfully sent !",64,"Information"
End If
End Function

' *******************  End of Common Routines **********************************************
Function TextSA(strMsg)
strMsg = date()&"-"&time()&" "&strLocale&"::"&strStage&"-"&strMsg
Call SendEmail("tokdmssa@gmail.com","tokdms alert", "9175795730@vtext.com", strMsg, strHeading, "")

End Function
%>