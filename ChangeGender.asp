<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<!-- #include file ="../include/connect.inc" -->
<!-- #include file ="../include/validate.inc" -->
<!-- #include file ="../include/locale.inc" -->
<!-- #include file ="../include/common.inc" -->
<%

			
'  ----------------------------  Load Fields -----------------------------------------------------------------------------
   intStudentID    = Request.Form("StudentID")
   strV            = "GenCd"&intStudentID 
   strGenCd        = Request.Form(strV)
   'response.write("Request.QueryString=>"&Request.Form&"<<br>")
   strSQL = "UPDATE dms_Student SET   GenCd='"&strGenCd&"' WHERE StudentID="&intStudentID&";" 
   ' response.write("strSQL=>"&strSQL&"<<br>")
   oConn.Execute strSQL
   'response.end()            
   strEM="Success for Student "&intStudentID&", updated to "&strGenCd&"."
   'response.write(strEM)
     ' Response.Write("<script> self.close(); window.onunload = refreshParent;</script>")

   Response.Write("<script>confirm('Update Successful') ;  </script>")
   response.redirect("AssignGender.asp#S"&intStudentID)
%>