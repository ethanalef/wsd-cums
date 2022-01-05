<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"
end if

f = ""
a = split(request("TS"),",",-1,1)
if ubound(a) < 0 then
	response.redirect "memList.asp"
end if

for i = 0 to ubound(a)
	f = f & "'"&trim(A(i))&"',"
next
f = left(f,len(f)-1)

if request("mActive") = "Yes" then m=" and memSection<>'S/W'" end if
SQl = "select * from memMaster where deleted=0 and memSection In ("&f&")"&m&" order by memSection,memNo"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
%>
<html>
<head>
<title>Member List</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<style>
td {font-size:8pt}
</style>
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="1" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="15"><font size="4">EMSD Credit Union</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="4">Member List</font></td>
	</tr>
	<tr height="15" valign="bottom">
		<td><b>Section</b></td>
		<td><b>A/C No.</b></td>
		<td><b>Name</b></td>
		<td><b>Address</b></td>
		<td><b>Tel.</b></td>
		<td><b>Mobile</b></td>
		<td><b>HKID</b></td>
		<td><b>Sex</b></td>
		<td><b>Birthday</b></td>
		<td><b>Grade</b></td>
		<td><b>Guarantor's no.</b></td>
		<td><b>Guarantor for others</b></td>
		<td><b>Treasury ref. no.</b></td>
		<td><b>First appointment date</b></td>
		<td><b>Membership date</b></td>
	</tr>
<%
do while not rs.eof
%>
	<tr>
		<td><%=rs("memSection")%></td>
		<td><%=rs("memNo")%></td>
		<td><%=rs("memName")%></td>
		<td><%=rs("memAddr1")&" "&rs("memAddr2")&" "&rs("memAddr3")%></td>
		<td><%=rs("memContactTel")%></td>
		<td><%=rs("memMobile")%></td>
		<td><%=rs("memHKID")%></td>
		<td><%=rs("memGender")%></td>
		<td><%=right("0"&day(rs("memBday")),2)&"/"&right("0"&month(rs("memBday")),2)&"/"&year(rs("memBday"))%></td>
		<td><%=rs("memGrade")%></td>
		<td><%=rs("memGuarantorNo")%></td>
		<td><%=rs("memGuarantor4Others")%></td>
		<td><%=rs("treasRefNo")%></td>
		<td><%=right("0"&day(rs("firstAppointDate")),2)&"/"&right("0"&month(rs("firstAppointDate")),2)&"/"&year(rs("firstAppointDate"))%></td>
		<td><%=right("0"&day(rs("memDate")),2)&"/"&right("0"&month(rs("memDate")),2)&"/"&year(rs("memDate"))%></td>
	</tr>
<%
	rs.movenext
loop
%>
</table>
</center>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
