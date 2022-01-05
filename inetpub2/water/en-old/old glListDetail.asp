<%requiredLevel=3%>
<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "glList.asp"
end if
if request("id")<>"" then
	set rs = server.createobject("ADODB.Recordset")
	sql = "select * from glMaster where glId='" & request("id") & "'"
	rs.open sql, conn
	if rs.eof then
		response.redirect "glList.asp"
	end if
	For Each Field in rs.fields
		TheString = Field.name & "= rs(""" & Field.name & """)"
		Execute(TheString)
	Next
else
	response.redirect "glList.asp"
end if
%>
<html>
<head>
<title>Account</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" alink="#003399" link="#003399" vlink="#003399" bgcolor="#eeeef0" onload="form1.id.focus()">
<!-- #include file="menu.asp" -->
<center>
<br>
<form name="form1" method="post" action="glListDetail.asp">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td class="b8" align="right">Account no.</td>
		<td width=10></td>
		<td><input type="text" name="id" value="<%=glId%>" size="4"><input type="submit" value="Search" name="action" class="sbttn"></td>
		<td width=10></td>
		<td class="b8" align="right" width="100">Current Period</td>
		<td width=10></td>
		<td>11</td>
	</tr>
	<tr>
		<td class="b8" align="right">Account Name</td>
		<td width=10></td>
		<td><%=glName%></td>
		<td colspan=4 align=right><input type="submit" value="Back" name="back" class="sbttn"></td>
	</tr>
</table>
<br>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
  <tr bgcolor="#330000" align="center">
	<td><font size="2" color="#FFFFFF">Period</font></td>
	<td><font size="2" color="#FFFFFF">Y-T-D Balance</font></td>
	<td><font size="2" color="#FFFFFF">M-T-D Balance</font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">1 (Sep)</font></td>
	<td align=right><font size="2"><%=formatnumber(monthBal1,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(currMTD,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">2 (Oct)</font></td>
	<td align=right><font size="2"><%=formatnumber(monthBal2,2)%></font></td>
	<td></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">3 (Nov)</font></td>
	<td align=right><font size="2"><%=formatnumber(monthBal3,2)%></font></td>
	<td></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">4 (Dec)</font></td>
	<td align=right><font size="2"><%=formatnumber(monthBal4,2)%></font></td>
	<td></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">5 (Jan)</font></td>
	<td align=right><font size="2"><%=formatnumber(monthBal5,2)%></font></td>
	<td></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">6 (Feb)</font></td>
	<td align=right><font size="2"><%=formatnumber(monthBal6,2)%></font></td>
	<td></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">7 (Mar)</font></td>
	<td align=right><font size="2"><%=formatnumber(monthBal7,2)%></font></td>
	<td></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">8 (Apr)</font></td>
	<td align=right><font size="2"><%=formatnumber(monthBal8,2)%></font></td>
	<td></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">9 (May)</font></td>
	<td align=right><font size="2"><%=formatnumber(monthBal9,2)%></font></td>
	<td></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">10 (Jun)</font></td>
	<td align=right><font size="2"><%=formatnumber(monthBal10,2)%></font></td>
	<td></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">11 (Jul)</font></td>
	<td align=right><font size="2"><%=formatnumber(monthBal11,2)%></font></td>
	<td></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">12 (Aug)</font></td>
	<td align=right><font size="2"><%=formatnumber(monthBal12,2)%></font></td>
	<td></td>
  </tr>
</table>
</center>
</form>
</body>
</html>
