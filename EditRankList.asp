<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/common.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/validate.inc" -->
<style type="text/css" media="screen">@import "../css/balloon.css";</style>
<style type="text/css" media="screen">@import "../css/basic.css";</style>
<style type="text/css" media="screen">@import "../css/tabs.css";</style>

<head>
	<title>Attendance Report</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
	<script defer src="../js/packs/solid.js"></script>
    <script defer src="../js/fontawesome.js"></script>
    <script defer src="../js/balloon.js"></script>
	


</head>
<%  
    Function DoHeading(bitIsKyu, strDojoName,datTestingDate, strDivisionHead) 
    strTestingDate = FormatDateTime(datTestingDate,1)
    bitIsDan = NOT bitIsKyu
	if bitIsKyu then
	   strstr = "Kyu"
	else
	   strstr = "Dan"
	end if

%>
      
      <table width="100%" >
        <tr>
          <td class="coltitle"><%=strdojoname%></td>
        </tr>  
      </table>
      <br>
      <table width="100%" >
        <tr>
          <td class="colparam1">Test Date:<%=strTestingDate%></td>
          <td class="coldiv">Rank List - <%=strstr%> </td>
          <td class="colparam">Printed: <%=Now()%></td>
        </tr>
      </table>
      <br>
      <table width="100%" >
         <thead>
         <tr class="rowtitle">
           <td class="">&nbsp;</td>
           <td class="">&nbsp;</td>
           <td class="">Date</td>
           <td class="">Next</td>
           <td class="">&nbsp;</td>
           <td class="">&nbsp;</td>
           <td class="">&nbsp;</td>
           <td colspan = "2" class="">Belt Size</td>
           <% if bitIsKyu then %>
           <% else %>
           <td class="">Gi</td>
           <% end if %>    
              
         </tr>

         <tr class="rowtitleunderline">
           <td class="colnamet">Name</td>
           <td class="colrankt">Last Rank</td>
           <td class="coldobt">Last Tested</td>
           <td class="colranknt">Rank/Held</td>
           <td class="coldivt">Div</td>
           <td class="colanchort"><i class="fa fa-anchor"></td>
           <td class="colaget">Age</td>
           <td class="bsizet">Curr</td>
           <td class="bsizet">Next</td>
           <%  if  bitIsKyu then %>
           <% else %>
           <td class="gsizet">Size</td>
           <%  end if %>    
           
         </tr>
         </thead>
         <tbody>
         
<%  End Function %>

<%  Function PageBreakBefore() %>
     <div class="page-break-before"></div>
<%  End Function %>   
   
<%  Function PageBreakAfter()  %>
     <div class="page-break-after"></div> 
<%  End Function %>      

<%  Function DoBlankLine(bitIsKyu, intRN)
    bitIsDan = NOT bitIsKyu
%>
       
         <tr class="rowtitle">
           <td class="colnamet">&nbsp;</td>
           <td class="colrankt">&nbsp;</td>
           <td class="coldobt">&nbsp;</td>
           <td class="colrank">&nbsp;</td>
           <td class="coldobt">&nbsp;</td>
           <td class="colanchort">&nbsp;</td>
           <td class="colaget">&nbsp;</td>
           <td class="colnotest">&nbsp;</td>
           <%  if bitIsKyu then %>
           <% else %>
           <td class="colnotest">&nbsp;</td>
           <%  end if %>    
           

<%      
    End Function
%>
<% Function FillBlankLines(intRC, intMax) 
       'response.write("Blank Lines From="&intRC&" to intMax="&intMax&"<br>")
       for intJ = intRC to intMax  ' Finish off the page with blank lines
           Call DoBlankLine(bitRowParity, intJ) 
           'response.write("RN=>"&intJ&"<<br>")
           if intJ = intMAX then
              response.write("</tbody></table>")
           end if   
       next
   End Function
%>
<%
    Function DoDataLine(bitIsKyu, strName, strLastRank, strLastTesting, strDiv, strAnchorday,  strAge, strBeltSize, strGiSize)
    
    bitIsDan = NOT bitIsKyu
%>   
 
         <tr class="">
           <td class="colname"><%=strName%></td>
           <td class="colrank"><%=strLastRank%></td>
           <td class="coldob"><%=strLasttesting%></td>
           <td class="colrankn">&nbsp;</td>
           <td class="coldiv"><%=strDiv%></td>
           <td class="colanchor"><%=strAnchorDay%></td>
           <td class="colage"><%=strAge%></td>
           <td class="bsize"><%=strBeltSize%></td>
           <td class="gsize">&nbsp;</td>
           <%  if bitIsKyu then %>
           <% else %>
           <td class="gsize"><%=strGiSize%></td>
           <%  end if %>    
         </tr>
<%
    End Function
%>
<div id="main">
    <div id="titlecontents">
        <div id="TitleBox">
            <h2><a href="Admin.asp"><i class="fa fa-bars"></i></a>&nbsp;&nbsp;<img src="../images/lillilinvertedtomoe.gif" />&nbsp;Dojo Management System </h2>
        </div>   
        <div id="AcctBox">
            <h4><a href="admin.asp"><%=strUserName%>&nbsp;<i class="fa fa-user"></i></a>&nbsp;&nbsp;&nbsp;<i class="fa fa-cog"></i>&nbsp;&nbsp;<a href="Logout.asp"><i class="fa fa-sign-out"></i></a></h4>
        </div>
    </div>  
	  <div id="contents">
	  <div id="OneCol">
	  <h2>Edit Rank List</h2>
      <a href="index.asp">Return to Menu</a>
      <br/>
            <TABLE  Border=0 Cellspacing=2 Cellpadding=2>
               <TABLE Border=0 Cellspacing=2 Cellpadding=2>
               <tr>
                  <form id="changedojocode" action="changedojocode.asp", method="post">
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
                   if strDojoCdA = "all" then
                      strDojoShortNameB = "All Dojos"
                   else
                      strDojoShortNameB = orsB("DojoShortName")
                   end if
                  %> 
                   <td>Current School: <b><%=StrDojoShortnameB%><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td> <td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td> </b><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td> <td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td> <td>&nbsp;&nbsp; <% if bitCanChangeDojo = True then %> Hold to Select School:<% end if %></td>
                   <td>
                   <%
                   
                  if bitCanChangeDojo = True then
                     %><select name="DojoCd" onchange="document.getElementById('changedojocode').submit();" >
                       <option value="all">All Dojos</option><%
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
			     
			   

		</div>	
		</div>
	</div>
</body>

<%
bitIsKyu = Request.QueryString("k")
bitisDan = NOT bitIsKyu
datTestingDate = Request.Form("TestingDate")
if NOT ISDATE(datTestingDate) then
   response.redirect("RankListSelect.asp?e=1&k="&bitIsKyu)
end if

strReqDojoCd = Request.Form("selDojoCd")
'response.write("Rank List Report<br>")
'response.write("k=>"&Request.QueryString("k")&"<<br>")
'response.write("TestingDate=>"&Request.Form("TestingDate")&"<<br>")
'response.write("selDojoCd=>"&Request.Form("seldojocd")&"<<br>")

strDojoCd = strReqDojoCd
'response.write("strReqDojoCd=>"&strReqDojoCd&"<<br>")
if NOT strReqDojoCd = "all" then
   strDojoSel = "AND (dojoCd='"&strReqDojoCd&"') "
else
   strDojoSel = ""
end if 
if bitIsKyu then
   strRankKOD = "<="
else
   strRankKoD = ">"
end if      
'strSQL = "Select * FROM dms_Student  Left Join dms_Codes on dms_student.AnchorDay = dms_Codes.CodeValue  WHERE CodeName = 'AnchorDay' AND kaicd = '"&strkaicd&"' "&strDojoSel&"  "&strChoose&" ORDER BY "&strOrder&" ;"
   
strSQL = "SELECT StudentID, dojoCd, IsActive, LName, FName, RankCd, DOB, Division, AnchorDay, BeltSizeCd, GiSizeCd FROM  dbo.dms_Student"
StrSQL = strSQL&" LEFT JOIN dms_Codes ON dms_student.AnchorDay = dms_Codes.CodeValue "
'strSQL = strSQL &" WHERE  (IsActive = 1) AND (dojoCd = '"&strReqDojoCd&"') AND (IsInstructor = 0) ORDER BY dojoCd, Division, dms_dowCD.doworder, LName, FName"
strSQL = strSQL &" WHERE CodeName = 'AnchorDay' AND  (IsActive = 1)  AND RankCd "&strRankKoD&" 12 "&strDojoSel& " ORDER BY  dms_codes.valueorder, Division, LName, FName"
'response.write("strSQL=>"&strSQL&"<<br>")



orsZ.open strSQL
intcnt = 0

intRowCountMax  = 35
intRowCount     = 0

intPageCount= 0

if trim(strDojoCd) = "bk"  then strDojoName = "Traditional Okinawan Karate of Brooklyn" end if
if trim(strDojoCd) = "kin" then strDojoName = "Traditional Okinawan Karate of Kinnelon" end if
if trim(strDojoCd) = "pv"  then strDojoName = "Traditional Okinawan Karate of Pleasant Valley" end if
if trim(strDojoCd) = "ef"  then strDojoName = "Traditional Okinawan Karate of East Fishkill" end if
if trim(strDojoCd) = "npk" then strDojoName = "New Paltz Karate Academy" end if


intYearYYYY = Year(strReqMMDDYY)
'%  Function DoHeading(bitIsKyu, strDojoName,datTestingDate, strDivisionHead, strDOW, inWeekday, intMonthlen) 

Call DoHeading(bitIsKyu, strDojoName, datTestingDate, strDiv)  'First Header
bitDebug = 0
intRowCount = 0
intcnt = 0
do while NOT orsZ.eof
   intstudentID   = orsZ("StudentID")
   strDojoCd      = orsZ("DojoCd")
   strLName       = orsZ("LName")
   strFName       = orsZ("FName")
   strName        = left(trim(strFName)&" "&trim(strLName),26)
   strDivision    = orsZ("Division")
   intRankCd      = orsZ("RankCd")
   strRankCd      = Cstr(intRankCd)
   datDOB         = orsZ("DOB")
   strAnchorDay   = orsZ("AnchorDay")
   strBeltSizeCd  = orsZ("BeltSizeCd")
   strGiSizeCd    = orsZ("GiSizeCd")
   
   intRowCount = intRowCount + 1
   intcnt = intcnt + 1
   if intRowCount > intRowCountMax  then 
   response.write("</tbody></table>")
      Call PageBreakBefore()
      Call DoHeading(bitIsKyu, strDojoName, datTestingDate, strDiv)  'First Header
      intRowCount = 1
   end if   
   intAge = CalcAge(datDOB, datTestingDate)
   if intAge > 100 then
      strAge = ""
   else   
      strAge = FormatNumber(intAge,1)
   end if   
   strRank = GetCodeValueText("RankCd", intRankCd)
   if NOT intRankCd = 0  then
      datLastTesting = GetLastTestingDate(intStudentID)
      strLastTesting = FormatDateTime(datLastTesting,2)
   else 
      strLastTesting = "N/A"   
   end if
     
   Call DoDataLine(bitIsKyu, strName,strRank, strLastTesting, strDivision,strAnchorday,  strAge, strBeltSizeCd, strGiSizeCd)      
   orsZ.MoveNext
loop
' We are done, finish up the last page


%>

</html> 