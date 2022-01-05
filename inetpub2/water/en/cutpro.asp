<%

cutdate = date()
lastmonth=month(date())-2
lastyear = year(date())
spass = 0
set rs = conn.execute("select * from monthend where works = 0 ")
if not  rs.eof then


   cutdate = rs("cutdate")
   if date() > cutdate then
     lastmonth = month(rs("lastdate"))
     lastyear  = year(rs("lastdate"))
   end if
end if
rs.close


%>
