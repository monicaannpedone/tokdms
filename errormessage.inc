<%
'
''This include sets up the error message array of the form "nn<text>" where N is the error message
'
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
%>