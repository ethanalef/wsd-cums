<!-- #include file="../conn.asp" -->
<!-- #include file="navigator.asp" -->
<%

key = request("key")
if key<>"" then
	stylefield = " memNo like '" & key & "%'"
end if
xyear = year(date())
xmonth = month(date())
xday  = day(date())
yyear = xyear-1
mdate = yyear&"/"&xmonth&"/"&xday

SQL = "select lnnum,date,code,date  from loan  where " & stylefield &  " and date >= '"&mdate&"' order by memNo"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 3

SQL1 = "select sum(amount)  from share  where " & stylefield & " and date < '"&mdate&"'group  by memNo"
Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.open sql1, conn, 4
opbal=rs1(0)
rs1.close
SQL1 = "select date,code,amount,tmpamt  from share  where " & stylefield & " and date >= '"&mdate&"'order by memNo"
Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.open sql1, conn, 4

do while not rs1.eof
   rs1(3) = rs1(3)+opbal+rs1(2)
   rs1.update
   rs1.movenext

   
loop
rs1.movetop

if not rs.eof then
	if request("page") <> "" then
		pageno = cint(request("page"))
	else
		pageno = 1
	end if
	rs.pagesize = 20
	pagesize=rs.pagesize
	rs.absolutepage = pageno
	recordcount=rs.recordcount
	pagecount = rs.pagecount
end if
%>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<title>Member Popup</title>
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function PostCode(memNo,memName,memGrade,employCond,firstAppointDate,memBday)
{
	window.opener.document.form1.memNo.value=memNo;
	window.opener.document.form1.memName.value=memName;
	window.opener.document.all.tags( "td" )['memName'].innerHTML=memName;
	window.opener.document.form1.memGrade.value=memGrade;
	window.opener.document.all.tags( "td" )['memGrade'].innerHTML=memGrade;
	window.opener.document.form1.employCond.value=employCond;
	window.opener.document.all.tags( "td" )['employCond'].innerHTML=employCond;
	window.opener.document.form1.firstAppointDate.value=firstAppointDate;
	window.opener.document.all.tags( "td" )['firstAppointDate'].innerHTML=firstAppointDate;

    D = parseInt(memBday.substr(0,2));
    M = parseInt(memBday.substr(3,2))-1;
    Y = parseInt(memBday.substr(6,4));
	today=new Date()
	dob=new Date(Y,M,D)
	age=Math.floor((today-dob)/(1000*60*60*24*365));

	window.opener.document.form1.age.value=age;
	window.opener.document.all.tags( "td" )['age'].innerHTML=age;
	window.opener.document.form1.netSalary.focus()
	window.close();
}
//-->
</script>
</head>
<body leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<div align="center">
<center>

<table border="0" cellpadding="0" cellspacing="0" width="100%">

	<tr>
		<td>
<%
if not rs.eof or not r1.eof then
	navigator(request.servervariables("script_name")&"?key="&key)
%>
		<table border="0" cellpadding="3" cellspacing="1" width="100%" bgcolor="#336699">
<%
	do while not rs.eof and rowcount < rs.pagesize
		rowcount = rowcount + 1
%>
			<tr>
				<td bgcolor="#ffffff"><a href="JavaScript:PostCode('<%=rs1(0)%>','<%=rs1(1)%>','<%=rs(2)%>','<%=rs(3)%>','<%=right("0"&day(rs(4)),2)&"/"&right("0"&month(rs(4)),2)&"/"&year(rs(4))%>','<%=right("0"&day(rs(5)),2)&"/"&right("0"&month(rs(5)),2)&"/"&year(rs(5))%>')"><% =rs("memNo") %></a></td>
				<td bgcolor="#ffffff"><a href="JavaScript:PostCode('<%=rs(0)%>','<%=rs(1)%>','<%=rs(2)%>','<%=rs(3)%>','<%=right("0"&day(rs(4)),2)&"/"&right("0"&month(rs(4)),2)&"/"&year(rs(4))%>','<%=right("0"&day(rs(5)),2)&"/"&right("0"&month(rs(5)),2)&"/"&year(rs(5))%>')"><% =rs("memName") %></a></td>
			</tr>
<%
		rs.movenext
                rs1.movenext
	loop
%>
		</table>
<%
	navigator(request.servervariables("script_name")&"?key="&key)
	response.write "<p>"
else
%>
		<br><p align="center"><font size="4">µL²Å¦X¬ö¿ý</font></p>
<%
end if
%>
		</td>
	</tr>
</table>
</center>
</div>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>