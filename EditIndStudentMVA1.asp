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
	<script type="text/javascript">
		function refresh() {
            setTimeout(function () {
                location.reload()
            }, 100);
        }
        function refreshParent() {window.opener.location.reload();}
    </script>
</head>

<body>

<div >
   <div >
	  <br/><br/>
		
			<% 
			bitIsError = 0
			
			
			strMode = LCase(Request.Querystring("m"))
			' response.write("Mode is >"&strMode&"</br>")
			strErrMsg   = Request.QueryString("err")
			intStudentID = Request.Querystring("u")
			'response.write("intStudentID is ==>"&strStudentID&"<<br>")
		
			
			intStudentID   = Request.Form("StudentID")
			strDojoCd      = Request.Form("DojoCd")
			strLName       = Request.Form("LName")
			strFName       = Request.Form("FName")
			strDivision    = Request.Form("Division")
			strDOB         = Request.Form("DOB")
			strAnchorDay   = Request.Form("AnchorDay")
			strBeltSizeCd  = Request.Form("BeltSizeCd")
			strAuthSpar    = Request.Form("AuthSpar")
			strAuthGrapple = Request.Form("AuthGrapple")
			strAreaCode    = Request.Form("AreaCode")
			strExchange    = Request.Form("Exchange")
			strLineNo      = Request.Form("LineNo")
			strStudentID = inStudentID
			'response.write("intStudentID is ==>"&strStudentID&"<<br>")

			if len(trim(strLName)) = 0 then 
			   strEM = "01 Last Name must not be blank"  
               bitIsError = 1  
			end if
			if len(trim(strFName)) = 0 then 
			   strEM = "01 First Name must not be blank"  
               bitIsError = 1  
			end if
			Select Case strDivision
			       Case "1"
			       Case "2"
			       Case "3"
			       Case "4"
			       Case Else
			            StrEM = "Division=>"&str&"<, must be 1, 2, 3 or 4"
			            bitIsError = 1
			End Select  
			if NOT isDate(strDOB) then
			   strEM = "The Date must be in the form MM-DD-YYYY"
			   bitIsError = 1
			end if
  
            if strAuthSpar = "Y" then
                 bitAuthSpar = 1 
                
              else   
                 bitAuthSpar = 0
                 
              end if   
              if strAuthGrapple = "Y" then
                 bitAuthGrapple = 1
              else   
                 bitAuthGrapple = 0
              end if     
              if NOT IsNumeric(strAreaCode) or (NOT len(Trim(strAreaCode)) = 3)   then
                 strEM = "Area Code must be 3 numeric characters"
                 bitIsError = 1
              end if 
              if NOT IsNumeric(strExchange) or (NOT len(Trim(strExchange)) = 3)   then
                 strEM = "Exchange must be 3 numeric characters"
                 bitIsError = 1
              end if           
              if NOT IsNumeric(strLineNo) or (NOT len(Trim(strLineNo)) = 4)   then
                 strEM = "Line Number=>"&strLineNo&"<, must be 4 numeric characters"
                 bitIsError = 1
              end if           
              strPhone = Trim(strAreaCode)&trim(strExchange)&trim(strLineNo)
			  if bitIsError then
			     response.write("intStudentID=>"&intStudentID&"<<br>")
			     strRR="EditIndStudentMVA.asp?u="&intStudentID&"&err="&strEM
			     response.write("strRR=>"&strRR&"<<br>")
			     response.redirect(strRR)
			     
			  end if 
            '  for each fld in ors.Fields
		'		 'response.write(fld.Name&"=>"&fld.Value&"<</br>")
		'		 if NOT fld.name="ChangeTimeStamp" then
		'		   ' response.write(fld.Name&"=>"&fld.Value&"<</br>")
		'		 else
		'		 end if
         '     next
		 '   end if  
			'ors.Close
		    'set ors = Nothing
		    'response.write(" 2 First name=>"&strFName&"<</br>")        
            strSQL = "UPDATE dms_Student  SET LName='"&strLName&"', FName='"&strFName&"', Division='"&strDivision&"',"
            strsql = strSQL & " anchorday='"&strAnchorDay&"', AuthSpar="&bitAuthSpar&", AuthGrapple="&bitAuthGrapple&", "
            strSQL = strSQL & " Phone='"&strPhone&"', BeltSizeCd='"&strBeltSizeCd&"' "
            strSQL = strSQL & " WHERE kaicd='"&strKaiCd&"' AND dojoCD='"&strDojoCd&"' AND StudentID="&intStudentID&";"
            'response.write("strSQL=>"&strSQL&"<<br>")
           
             oConn.execute strSQL
            'response.end
             response.flush()
           Response.Write("<script> self.close(); window.onunload = refreshParent;</script>")
           'response.redirect("MonthViewAttendance.asp")
			%>
				  
				  
		</div>
	</div>
</body>
</html>
