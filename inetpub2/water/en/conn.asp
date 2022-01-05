<%

lsDate = Request.Form("date")
newdate = year(lsDate) & "/" & month(lsDate) & "/" & Day(lsDate)
Set conn = Server.CreateObject("adodb.connection")
conn.Open "Driver={SQL Server};Server=(LOCAL);Database=wsdscu_new"
cuname=""
FMonthStart = 8
ArrMonth = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

%>
