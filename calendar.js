// JavaScript Document
<!--
function ShowCalendar(str_target, strDate) {
   var arr_months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
   var week_days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
   var n_weekstart = 1;   // 0 - Sun, 1 - Mon
   var dt_datetime = (strDate == null || strDate =="" ?  new Date() : Str2Date(strDate));
   
   var dt_prev_month = new Date(dt_datetime);
   dt_prev_month.setMonth(dt_datetime.getMonth()-1);
   
   var dt_prev_year = new Date(dt_datetime);
   dt_prev_year.setYear(dt_datetime.getYear()-1);
   
   var dt_next_month = new Date(dt_datetime);
   dt_next_month.setMonth(dt_datetime.getMonth()+1);
   
   var dt_next_year = new Date(dt_datetime);
   dt_next_year.setYear(dt_datetime.getYear()+1);
   
   var dt_firstday = new Date(dt_datetime);
   dt_firstday.setDate(1);
   dt_firstday.setDate(1-(7+dt_firstday.getDay()-n_weekstart)%7);
   var dt_lastday = new Date(dt_next_month);
   dt_lastday.setDate(0);
   
   // calendar header
   var sBuffer = new String (
      "<html>\n"+
      "<head>\n"+
      "   <title>Calendar</title>\n"+
      "</head>\n"+
      
      "<body style='margin-left:0;margin-right=0;margin-top=0;margin-bottom=0' bgcolor='white' onBlur='self.focus();'>\n"+
      "<table class='clsOTable' cellspacing='0' border='0' width='100%'>\n"+
      "<tr><td bgcolor='#4682B4'>\n"+
      "<table cellspacing='1' border='0' width='100%'>\n<tr>\n"+

      "<td bgcolor='#4682B4' align='center'></td>\n"+

      "<td bgcolor='#4682B4' align='center'><a href=\"javascript:window.opener.ShowCalendar('"+
      str_target+"', '"+ Date2Str(dt_prev_month)+"');\">"+
      "<img src='prev_month.gif' border='0' alt='Previous Month'></a></td>\n"+

      "   <td bgcolor='#4682B4' colspan='3' align='center'>"+
      "<font color='white' face='tahoma, verdana' size='2'>"+
      arr_months[dt_datetime.getMonth()]+" "+dt_datetime.getFullYear()+"</font></td>\n"+

      "   <td bgcolor='#4682B4' align='center'><a href=\"javascript:window.opener.ShowCalendar('"+
      str_target+"', '"+Date2Str(dt_next_month)+"');\">"+
      "<img src='next_month.gif' border='0' alt='Next Month'></a></td>\n"+
      
      "   <td bgcolor='#4682B4' align='center'></td>\n"+
      
      "</tr>\n"
   );

   // weekdays titles
   sBuffer += "<tr>\n";
   for (var n=0; n<7; n++)
      sBuffer += "   <td bgcolor='#87CEFA'  width=50 align='center'>" + "<font color='white' face='tahoma, verdana' size='2'>" + week_days[(n_weekstart+n)%7] + "</font></td>\n";
      
   // calendar table
   sBuffer += "</tr>\n";
   var dt_current_day = new Date(dt_firstday);
   while (dt_current_day.getMonth() == dt_datetime.getMonth() || dt_current_day.getMonth() == dt_firstday.getMonth()) {
      // row heder
      sBuffer += "<tr>\n";
      for (var n_current_wday=0; n_current_wday<7; n_current_wday++) {
            if (dt_current_day.getDate() == dt_datetime.getDate() && dt_current_day.getMonth() == dt_datetime.getMonth())
               // current date
               sBuffer += "   <td bgcolor='#FFB6C1' align='right'>";
            else if (dt_current_day.getDay() == 0 || dt_current_day.getDay() == 6)
               // weekend days
               sBuffer += "   <td bgcolor='#DBEAF5' align='right'>";
            else
               // working days of current month
               sBuffer += "   <td bgcolor='white' align='right'>";

            if (dt_current_day.getMonth() == dt_datetime.getMonth())
               // days of current month
               sBuffer += "<a href=\"javascript:window.opener." + str_target + ".value='" + Date2Str(dt_current_day) + "'; window.close();\">" + "<font color='black' face='tahoma, verdana' size='2'>";
            else
               // days of other months
               sBuffer += "<a href=\"javascript:window.opener." + str_target + ".value='" + Date2Str(dt_current_day) + "'; window.close();\">" + "<font color='gray' face='tahoma, verdana' size='2'>";
               
            sBuffer += dt_current_day.getDate() + "</font></a></td>\n";
            dt_current_day.setDate(dt_current_day.getDate()+1);
      }
      // row footer
      sBuffer += "</tr>\n";
   }
   // calendar footer
   sBuffer += "</table>\n" + "</tr>\n</td>\n</table>\n" + "</body>\n" + "</html>\n";

   var vWinCal = window.open("", "Calendar", "width=300,height=260,status=no,resizable=yes,top=200,left=200");
   vWinCal.opener = self;
   var calc_doc = vWinCal.document;
   calc_doc.write (sBuffer);
   calc_doc.close();
}

// datetime parsing and formatting routimes
function Str2Date(strDate) {
   var re_date = /^(\d+)\/(\d+)\/(\d+)/;
   if (!re_date.exec(strDate))
      return alert("Invalid Datetime format: "+ strDate);
   return (new Date (RegExp.$3, RegExp.$1-1, RegExp.$2));
}
function Date2Str(dt_datetime) {
   return (new String ((dt_datetime.getMonth()+1) + "/" + dt_datetime.getDate() + "/" + dt_datetime.getFullYear()));
} 
-->