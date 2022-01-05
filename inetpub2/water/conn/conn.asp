<%
Set conn = Server.CreateObject("adodb.connection")
conn.Open "Driver={SQL Server};Server=(local);Database=wsdscu"
CUName = "香港政府華員會系統"
FMonthStart = 8
ArrMonth = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

%>
