<%
'============================================================================================
' PROGRAM INFO
'
' Developer: T.Kaihlaniemi 2001
' Email: tk@dragrace.org
' Date: 11.12.2001
' Description: ASP Calendar for showing day-links for selected days
'
' Modified 19.12.2001 to show english as default language
' Language versions: Eng  and finnish (calendar.asp?l=fi)
'============================================================================================
option explicit
Response.Expires=-240
' Session lcid for finnish date and time is 1035
Session.LCID = 1035

Response.Buffer = true
Response.Addheader "Pragma","no-cache"
'============================================================================================
' VARIABLES
'============================================================================================
Dim myToday ' this day
Dim i
Dim SelMonth, selDay, selYear
Dim tmpStr
Dim monthNamesFin, monthNamesEn, monthNames, dayNamesFin, dayNamesEn, dayNames
Dim lastDay
Dim tmpLang
dim firstWeekDay
Dim scriptName
scriptName = Request.ServerVariables("SCRIPT_NAME")
scriptName = replace(replace(scriptName, "default.asp", ""),"Default.asp", "")
Dim imgPath
imgPath = "/images/"
'============================================================================================
' HTML
'============================================================================================
%>
<html>
<head>
<title>ASPCalendar</title>
<style type="text/css">
.day{
	font-family:arial;
	font-size:11px;
	}
.day a {
	font-weight:bold;
	color: #000000;
	text-decoration:none;
}
.day a:visited {
	font-weight:bold;
	color: #000000;
	
}
.day a:hover {
	font-weight:bold;
	color: #2a2a2a;
}
.weekend {
	background-color: #f6f6f6;
}
.monthtableouter {
	background-color: #cccccc;
}
.monthtableinner {
	width: 200px;
}
.normalday {
	background-color: #ffffff;
}
.title {
	background-color: #f6f6f6;
	font-weight:bold;
	font-family:verdana;
	font-size:12px;
}
.month, .days {
	background-color: #f6f6f6;
	font-size:10px;
	font-weight:normal;
	font-family:arial;
}
.selectedday {
	background-color: #DBDBDB;
}
#today {
	border: 1px #000000 solid;
	background-color: #C4C4C4;
}
</style>
</head>
<body>
<form method="get" action="<%=Request.ServerVariables("SCRPIPT_NAME")%>">
<%
' call calendar
makeCalendar Request("l"),Request("d"),Request("m"),Request("y"),request("s")
%>
</form>
</body>
</html>
<%
'============================================================================================
' FUNCTIONS
'============================================================================================
function language(byval lang)

' finnish is default
' finnish months
monthNamesFin = array("", "Tammikuu", "Helmikuu", "Maaliskuu", "Huhtikuu", "Toukokuu", "Kesäkuu", "Heinäkuu", "Elokuu", "Syyskuu", "Lokakuu", "Marraskuu", "Joulukuu")
' english months
monthNamesEn = array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")

' finnish days
dayNamesFin = array("", "Ma", "Ti", "Ke", "To", "Pe", "La", "Su")
' english days 
dayNamesEn = array("", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun") 

tmpLang = lang
	if lang = "en" then
		dayNames = dayNamesEn
		monthNames = monthNamesEn
	elseIf lang = "fi" then
		dayNames = dayNamesFin
		monthNames = monthNamesFin
	else
		dayNames = dayNamesEn
		monthNames = monthNamesEn
	end if
end function
'============================================================================================

' renders monthselection
function monthSelect()
	with Response
		.Write "<a href=""" & scriptName & "?m=" & selMonth - 1 & "&y=" & selYear & "&d=" & selDay & """>"
		.write "<img src=""" & imgPath & "calPrev.gif"" border=""0"" alt=""" & monthNames(selMonth - 1)
		.Write """></a>" & chr(10)
		.Write "<select name=""m"" class=""month"" onChange=""showNew()"">" & chr(10)
	end with
	
	for i = 1 to Ubound(monthNames)
		if i = int(selMonth) then 
			tmpStr = " selected"
		else
			tmpStr = ""
		end if
		Response.Write "<option value=""" & i & """" & tmpStr & ">" & monthNames(i) & "&nbsp;</option>" & chr(10)
	next	
	Response.Write "</select>" & chr(10)
end function
'============================================================================================

' renders yearselection
function yearSelect()
	Response.Write "<select name=""y"" class=""month"" onChange=""showNew()"">" & chr(10)
	for i = year(now) - 5 to year(now) + 5
		if i = int(selYear) then 
			tmpStr = " selected"
		else
			tmpStr = ""
		end if
		Response.Write "<option value=""" & i & """" & tmpStr & ">" & i & "</option>" & chr(10)
	next
	
	' lets check if the next month value is more than 12
	' if, then show 1 month of next year
	dim selMonthShowNext, selYearShowNext
	selMonthShowNext = selMonth + 1
	selYearShowNext = selYear
	if selMonthShowNext < 1 then 
		selMonthShowNext = 12
		selYearShowNext = selYear - 1
	end if
	if selMonthShowNext > 12 then 
		selMonthShowNext = 1
		selYearShowNext = selYear + 1
	end if
	
	' write end for select, next month-link and hidden for language-selection
	with response
		.Write "</select>" & chr(10)
		.Write "<a href=""" & scriptName & "?m=" & selMonthShowNext & "&y=" & selYearShowNext & "&d=" & selDay & """>"
		.write "<img src=""" & imgPath & "calNext.gif"" border=""0"" alt=""" & monthNames(selMonthShowNext)
		.Write """></a>" & chr(10) 		
		.Write "<input type=""hidden"" name=""l"" value=""" & tmpLang & """>"
	end with
end function
'============================================================================================

' calculate dates
function dayToShow(byval selDayTmp, byval selMonthTmp, byval selYearTmp)
selDay = selDayTmp : selMonth = selMonthTmp : selYear = selYearTmp

' if day selection values are empty, use current date
if len(SelDay) = 0 then SelDay = day(now)
if len(SelYear) = 0 then SelYear = year(now)
if len(selMonth) = 0 then selMonth = month(now)

' if month is 0, then show month 12 of previous year
if selMonth < 1 then
	selMonth = 12
	selYear = selYear - 1
end if

' if month is over 12, then show 1 month of next year
If selMonth > 12 then
	selMonth = 1
	selyear = selYear + 1
end if

' temporary date for date calculations
dim tmpDate
tmpDate = CDate("1." & selMonth & "." & selYear)

' how many days are in this month
lastDay = day(DateSerial(Year(tmpDate), Month(tmpDate) + 1, 0))

' what is the weekday of first day in this month
firstWeekDay = weekday(DateSerial(Year(tmpDate), Month(tmpDate), 0)+1, 2)

' check if selected date is valuable
if selDay < 1 then selDay = 1
if int(selDay) > int(lastDay) then selDay = lastDay
end Function
'============================================================================================

' render calendar
function makeCalendar(byval lang, byval selDayTmp, byval selMonthTmp, byval selYearTmp, byval linkDays)
	
	Dim arrlinkDay
	arrLinkDay = split(linkDays, ",")	
	dayToShow selDayTmp, selMonthTmp, selYearTmp
	language(lang)	
	dim tmpDayInt, tmpDayInt2
	tmpDayint = 0	
	' render javascript 
	with response
		.Write "<script language=""javascript"">" & chr(10)
		.write "function showNew() {" & chr(10)
		.Write "document.forms[0].submit();" & chr(10)
		.Write "}" & chr(10) & "</script>"
		'.Write "<form method=""get"" action=""" & Request.ServerVariables("SCRIPT_NAME") & """>"
		.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" class=""monthtableouter""><tr><td>"
		.Write "<table border=""0"" cellpadding=""1"" cellspacing=""1"" class=""monthtableinner"">" & chr(10)
		.Write "<tr>" & chr(10) & "<td class=""title"" colspan=""7""><div align=""center"" class=""title"">"
		.Write monthNames(selMonth)& " " & selYear & "</div></td>" & chr(10) & "</tr>" & chr(10)
		.Write "<tr>"
	end with
	
	for i = 1 to UBound(dayNames) 
		with response
			.Write "<td class=""days"" width=""13%""><div align=center class=""days"">"
			.Write dayNames(i)
			.Write "</td>"
		end with		
	next
	Response.Write "</tr>"
	
	dim ib, thisDay
	'do while tmpDayInt2 < lastDay
	for ib = 1  to 6
	Response.Write "<tr>" & chr(10)
		for i = 1 to 7			
		
		if int(tmpDayInt2) = int(day(now)) - 1 and int(month(now)) = int(selMonth) and int(year(now)) = int(selYear) then 
			thisDay = " id=""today"""
		else
			thisDay = ""
		end if			
			
			if tmpDayInt2 = selDay - 1 and firstWeekday > tmpDayInt + 1 then
				Response.Write "<td class=""normalday"" width=""14%""><div align=""center"" class=""day"">"
			elseif tmpDayInt2 = selDay - 1 then
				Response.Write "<td class=""selectedday"" width=""14%""" & thisDay & "><div align=""center"" class=""day"">"
			else
				' check if day is weekend
				if i = 6 or i = 7 then
					Response.Write "<td class=""weekend"" width=""14%""" & thisDay & "><div align=""center"" class=""day"">" 
				else
					Response.Write "<td class=""normalday"" width=""14%""" & thisDay & "><div align=""center"" class=""day"">" 
				end if
			end if		
			
			
			if firstWeekday > tmpDayInt + 1 or lastDay < tmpDayInt2 + 1 then
				Response.Write "&nbsp;"
			else
				tmpDayInt2 = tmpDayInt2 + 1				
				dim dayFound, arrI
				dayFound = false
				for arrI = Lbound(arrLinkDay) to UBound(arrLinkDay)
					if int(tmpDayInt2) = int(arrLinkDay(arrI)) then dayFound = true
				next
				if dayFound = true then
					with response
						.Write "<a href=""" & scriptName & "?d=" & tmpDayInt2 & "&m=" 
						.write selMonth & "&y=" & selYear & """>" & tmpDayInt2 & "</a>"
					
					end with
					dayFound = false
				else
					Response.Write tmpDayInt2
				end if								
			end if
				
			Response.Write "</td>" & chr(10)
			tmpDayInt = tmpDayInt + 1
		next		
	Response.Write "</tr>" & chr(10)
	'loop
	next
	
	with Response
		.Write "<tr>" & chr(10) & "<td class=""title"" colspan=""7"" nowrap>" & chr(10)
		.Write "<div align=""center"" class=""title"">"
		.Write monthSelect()
		.Write "&nbsp;"
		.write yearSelect()
		.Write "</div>" & chr(10) & "</td>"
		.Write chr(10) & "</tr>" & chr(10) & "</table></td></tr></table>"
		'.Write "</form>"
	end with
end function
'============================================================================================
' that's it, the end
%>