<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<script LANGUAGE="JavaScript" SRC="calendar.js">
</script>
<style type="text/css" media="screen">@import "../css/basic.css";</style>
<style type="text/css" media="screen">@import "../css/tabs.css";</style>

<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/locale.inc" -->
<head>
	<title>Proof of Delivery</title>
	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />

	
	
	
</head>

<body>

     <%  
	   
	   if len(request.querystring("Storenumber")) > 1 then
		   strStore = trim(request.querystring("Storenumber"))
		   strIdate = ""
		else   
	       strStore=Request.Form("StoreNumber")
	       strIDate=Request.Form("IDate")
		end if
		if strStore = "*" or strStore = "?" then
		   response.redirect("pickstore.asp?act=pod")
		end if   
		if  len(strStore) > 0 then 
		    if not isnumeric(strStore) then
		      response.redirect("runact.asp?act=pod&err=Store >"&strStore&"< is blank or not numeric")	
		   end if
		else   
		   if len(strIDate) > 5 then
		      if IsDate(strIDate) then
		         datIDate = CDate(strIDate)
				 response.redirect("picknight.asp?act=pod&idate="&strIdate)
		      else
		         response.redirect("runact.asp?act=pod&err=Date >"&strIDate&"< is invalid")	 
		      end if	   
		   else 
		      if strIdate = "*" or strIdate = "?" then
			     response.redirect("picknight.asp?act=pod&idate=*")
		      else
			     response.redirect("runact.asp?act=pod&err=Date >"&strIDate&"< is invalid")
			  end if 	 
		   end if
		end if
	    strErrmsg = request.querystring("err")
	 
	 
	 
	    'strStore=Request.Form("StoreNumber")
	        
			'response.write("<h2>Results for Store Number "&strStore&"</h2>")
			'response.write(" store Number = >"&strStore&"< <br>")			
			strSQL = "SELECT * FROM "&strTablePrefix&"store where [Store Number] = '"&strStore&"';"
			'strSQL = "SELECT * FROM swp_store;"
			'response.write(" SQL = >"&strSQL&"< <br>")	
			oRS.Open strSQL
			
            if oRS.EOF then
			   strerrmsg="Store >"&strStore&"< was not found"
			   response.redirect("runact.asp?act=pod&err="&strerrmsg)
			end if
			   
 	   
strIsComplete               = ors("IsComplete")			   
strSWIPConversionNight      = ors("SWIP Conversion Night")
strDivision                 = ors("Division")
strSubDivision              = ors("Sub Division")
strAddress                  = ors("Address")
strSpaceNumber              = ors("Space Number")
strCity                     = ors("City")
strState                    = ors("State")
strZipcode                  = ors("Zip Code")
strPhoneNumber              = ors("Phone Number")
strRevisedSwitchCategory    = ors("Revised Switch Category")
strNumOfRegisters           = ors("Num of Registers")
strMODContacted             = ors("Delivery Confirm Mod")
strDeliveryConfirmationDate = ors("Delivery Confirmation Date")
strDeliveryConfirmMod       = ors("Delivery Confirm Mod")
strConversionDateNight1     = ors("Conversion Date Night 1") 
strConversionDateNight2     = ors("Conversion Date Night 2") 
strConversionDateNight3     = ors("Conversion Date Night 3") 
strConversionDateNight4     = ors("Conversion Date Night 4") 
strConversionDateNight5     = ors("Conversion Date Night 5") 
strConversionDateNight6     = ors("Conversion Date Night 6") 
strConversionDateNight7     = ors("Conversion Date Night 7") 
strConversionDateNight8     = ors("Conversion Date Night 8")
strConversionDateNight9     = ors("Conversion Date Night 9") 
strConversionDateNight10    = ors("Conversion Date Night 10") 
strConversionDateNight11    = ors("Conversion Date Night 11") 
strConversionDateNight12    = ors("Conversion Date Night 12") 
strEquipmentDeliveryDate    = ors("Equipment Delivery Date")
strRCSCRIISNumber           = ors("RCSC RIIS Number")
strRCSCShippedDate          = ors("RCSC Shipped Date")
strRCSCTrackingNumbers      = ors("RCSC Tracking Numbers")
strJabilPortalNumber        = ors("Jabil Portal Number")
strJabilShippedDate         = ors("Jabil Shipped Date")
strJabilTrackingNumbers     = ors("Jabil Tracking Numbers")
strCanadaScoreNumber        = ors("Canada SCORE Number")
strCanadaShippedDate        = ors("Canada Shipped Date")
strCanadaTrackingNumbers    = ors("Canada Tracking Numbers")
strSpencerWorcestershire    = ors("Spencer Worcestershire")
strJabilMemphis             = ors("Jabil Memphis")
strIBMRochester             = ors("IBM Rochester")
strTotalBoxes               = ors("Total Boxes")
strResch                    = ors("Resch")
strReschParts               = ors("Resch Parts")
strReschPartType            = ors("Resch Part Type")
strComments                 = ors("Comments")

strComments = trim(ors("Comments"))
strComments = replace(strComments,"/ ", "/")
strComments = replace(strComments,vbCRLF, "")
strCommentsDisp = ""

dim aComments
aComments = split(strComments," ",-1)
intBnd = UBound(aComments)
'Response.Write("IntBnd=>"&intBnd&"< <BR>")
for intI = 0 to intBnd
    'response.write("aComments("&intI&")=>"&aComments(intI)&"<<br>")
	strMid = aComments(IntI)
	strX = ""
	if len(strMid) > 3 and intI > 1 then
	   intJ = instr(1,strMid,"/")
	   if intJ > 1  and intJ < len(strMid) then 
	      strMonth = mid(strMid,1,len(strMid)-intJ-1)
		  strDay   = mid(strMid,intJ+1,len(strMid)-IntJ)
		  'response.write("Month=>"&strMonth&"<; Day=>"&strDay&"< <br>")
	      if isNumeric(strMonth) and isNumeric(strDay) then
		     strX = vbCRLF
		  end if
	   end if	  
	end if
	strCommentsDisp = strCommentsDisp & strX & aComments(intI)& " "
	'response.write("strCommentsDisp=>"&strCommentsDisp&"<<br>")	
next

strCompleteText=""
if strIsComplete = 1 then
   strCompleteText = " Completed "
end if   
strStoreDisplay = strStore&strCompleteText

%>

	

<div class="formouter">	
<div class="errmsg"><p><%=strErrMsg%></p></div>	
<h1>Action - Proof of Delivery</h1>
<form name="POD"	 method="post" action="actPOD1.asp"	>	
<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="white">
  <input type="hidden" name="StoreNumber" value="<%=strStore%>" />
  <input type="hidden" name="Idate" value="<%=strSWIPConversionNight %>" />
   <tr>
    <td width="100%">
      <table cellpadding="2" cellspacing="2" border="0" width="100%" bgcolor="white">
        <tr>
          <td colspan=6 class="style0"> Store Information</td>
        </tr>  
        <tr>
  	      <td width="4%" class="style1">Store:</td>
          <td width="25%" class="style2"><%=strStoreDisplay%></td>
          <td width="10%" class="style1">Division:</td>
          <td width="20%" class="style2"><%=strDivision%></td>
          <td width="18%" class="style1">Install Night:</td>
          <td width="23%" class="style2"><%=strSWIPConversionNight%></td>
          
        </tr>
        <tr>
          <td width="4%" class="style1">Address:</td>
          <td width="25%" class="style2"><%=strAddress%></td>
          <td width="10%" class="style1">Sub Division:</td>
          <td width="20%" class="style2"><%=strsubdivision%></td>
          <td width="18%" class="style1">Equipment Delivery Date:</td>        
          <td width="23%" class="style2"><%=strEquipmentDeliveryDate%></td>
        </tr>
        
        <tr>
          <td width="4%" class="style1">City:</td>
          <td width="25%" class="style2"><%=strCity%></td>
          <td width="20%" class="style1">Space Number:</td>
          <td width="20%" class="style2"><%=strSpaceNumber%></td>
          <td width="18%" class="style1">Delivery Confirm Date:<a href="javascript:ShowCalendar('document.POD.DeliveryConfirmationDate', document.POD.DeliveryConfirmationDate.value);">
  <img src="calendar.gif"  border="0" alt="Click Here to Pick a Date"/></a></td>        
          <td width="23%" class="style2"><input name="DeliveryConfirmationDate" value="<%=strDeliveryConfirmationDate%>" type="text" /></td>
          
        </tr>
        <tr>
          <td width="4%" class="style1">State:</td>
          <td width="25%" class="style2"><%=strState%></td>          
          <td width="10%" class="style1">Revised Switch Category:</td>
          <td width="20%" class="style2"><%=strRevisedSwitchCategory%></td>
          <td width="18%" class="style1">Delivery Confirm Mod:</td>
          <td width="23%" class="style2"><input name="Delivery Confirm Mod" value="<%=strDeliveryConfirmMod%>" type="text" /></td> 
        </tr>
        <tr>
          <td width="4%" class="style1">Zipcode:</td>
          <td width="25%" class="style2"><%=strZipCode%></td>
          <td width="10%" class="style1"># of Registers:</td>
          <td width="20%" class="style2"><%=strNumOfRegisters%></td>
          <td width="18%" class="style1"></td>
          <td width="23%" class="style2"></td> 
        </tr>
        <tr>
          <td colspan=6>&nbsp;</td>           
        </tr> 
        <tr>
          <td colspan=6 class="style0">Canada</td>           
        </tr>
        <tr>
          <td width="10%" class="style1">Score #:</td>
          <td width="20%" class="style2"><%=strCanadaScoreNumber%></td>
          <td width="10%" class="style1">Ship Date:</td>
          <td width="20%" class="style2"><%=strCanadaShippedDate%></td>
          <td width="20%" class="style1">Tracking Nums:</td>
          <td width="20%" class="style2"><%=strCanadaTrackingNumbers%></td> 
        </tr>
              
      </table>
    </td>
   </tr>            
</table> 
<table cellpadding="2" cellspacing="2" border="0" width="100%"  bgcolor="white" >
  <tr>
    <td colspan=8 class="style0">&nbsp;</td>
  </tr>  
  <tr>
    <td colspan=8 class="style0">Boxes</td>
  </tr>
  <tr>
    <td width="20%" class="style1">Spencer Worscestershire:</td>
    <td width="5%" class="style2"><%=strSpencerWorcestershire%></td>
    <td width="20%" class="style1">IBM Rochester:</td>
    <td width="5%" class="style2"><%=strIBMRochester%></td>
    <td width="20%" class="style1">Jabil Memphis:</td>
    <td width="5%" class="style2"><%=strJabilMemphis%></td>
    <td width="20%" class="style1">Total Boxes:</td>
    <td width="5%" class="style2"><%=strTotalBoxes%></td>
  </tr>
  <tr>
    <td colspan=8 class="style0">&nbsp;</td>
  </tr>  
</table>
<table cellpadding="2" cellspacing="2" border="0" width="100%"  bgcolor="white" >
  <tr>
    <td colspan=8 class="style0" >Conversion Night Dates</td>
  </tr>
  <tr>
  <td width="12%" class="style1" >Night 1:</td>
  <td width="13%" class="style2"><%=strConversionDateNight1%></td>
  <td width="12%" class="style1">Night 4:</td>
  <td width="13%" class="style2"><%=strConversionDateNight4%></td>
  <td width="12%" class="style1">Night 7:</td>
  <td width="13%" class="style2"><%=strConversionDateNight7%></td>
  <td width="12%" class="style1">Night 10:</td>
  <td width="13%" class="style2"><%=strConversionDateNight10%></td>
  </tr>
  <tr>
  <td width="12%" class="style1">Night 2:</td>
  <td width="13%" class="style2"><%=strConversionDateNight2%></td>
  <td width="12%" class="style1">Night 5:</td>
  <td width="13%" class="style2"><%=strConversionDateNight5%></td>
  <td width="12%" class="style1">Night 8:</td>
  <td width="13%" class="style2"><%=strConversionDateNight8%></td>
  <td width="12%" class="style1">Night 11:</td>
  <td width="13%" class="style2"><%=strConversionDateNight11%></td>
  </tr>
  <tr>
  <td width="12%" class="style1">Night 3: </td>
  <td width="13%" class="style2"><%=strConversionDateNight3%></td>
  <td width="12%" class="style1">Night 6:</td>
  <td width="13%" class="style2"><%=strConversionDateNight6%></td>
  <td width="12%" class="style1">Night 9:</td>
  <td width="13%" class="style2"><%=strConversionDateNight9%></td>
  <td width="12%" class="style1">Night 12:</td>
  <td width="13%" class="style2"><%=strConversionDateNight12%></td>
  </tr>
  <tr>
  <td colspan=8>&nbsp;</td>
  </tr>
</table> 
<table cellpadding="2" cellspacing="2" border="0"  width="100%" bgcolor="white" >
  <tr>
    <td width="34%">
      <table cellpadding="2" cellspacing="2" border="0" width="100%"  >
        <tr  valign="top">
          <td colspan=2 class="style0">RSCS</td>
        </tr>
        <tr>
          <td class="style1" >RIIS:</td>
          <td class="style2"><%=strRCSCRIISNumber%></td>
        </tr>
        <tr>
          <td class="style1">Shipped Date:</td>
          <td class="style2"><%=strRCSCShippedDate%></td>
        </tr>
        <tr>
          <td valign="top" class="style1">Tracking #:</td>
          <td class="style2"><%=strRCSCTrackingNumbers%></td>
        </tr>
        
      </table>  
    </td>
    <td width="33%">
    <table cellpadding="2" cellspacing="2" border="0" width="100%"  >
        <tr>
          <td colspan=2 class="style0">Jabil</td>
        </tr>
        <tr>
          <td class="style1">Portal Number:</td>
          <td class="style2"><%=strJabilPortalNumber%> </td>
        </tr>
        <tr>
          <td class="style1">Shipped Date:</td>
          <td class="style2"><%=strJabilShippedDate%></td>
        </tr>
        <tr>
          <td class="style1" valign="top">Tracking #:</td>
          <td class="style2"><%=strJabilTrackingNumbers%></td>
        </tr>
      </table> 
    </td>
    <td width="33%" valign="top">
     <table cellpadding="2" cellspacing="2" border="0" width="100%" >
        
        <tr>
          <td colspan=2 class="style0">Reschedule</td>
        </tr>
        <tr>
          <td class="style1">Is a Reschedule?:</td>
          <td class="style2"><%=strResch%></td>
          </tr>
        <tr>
          <td class="style1">Parts:</td>
          <td class="style2"><%=strReschParts%></td>
        </tr>
        <tr>
          <td class="style1" valign="top">Part Type:</td>
          <td class="style2"><%=strReschedPartType%></td>
        </tr>
        
        
       
        
      </table> 
      
    
    </td>
  </tr>
 
</table> 
<table cellpadding="2" cellspacing="2" border="0" width="104%"  bgcolor="white" >
  <tr>
    <td colspan=2>&nbsp;</td>
  </tr>  
  <tr >
   <td valign="top"  class="style0">Comments</td>
   <td class="style2"> <textarea name="Comments"  cols="140" rows="6" ><%=strCommentsDisp%></textarea></td>
  </tr>
  <tr>
    <td colspan=2>&nbsp;</td>
  </tr>  
</table> 
<table cellpadding="2" cellspacing="2" border="0" width="100%"  bgcolor="white">
  <tr>
   <td>      <input name="reset"   type="reset"/></td>
   <td align="right"> <input name="submit"  value="Update"  type="submit" /></td>
  </tr>
</table> 
</form>  
</div>
</body>
</html>
