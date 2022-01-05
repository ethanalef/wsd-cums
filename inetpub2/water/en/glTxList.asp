<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
SQL = "select min(glId) as code from glMaster where deleted=0"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
mFrom = rs("code")
SQL = "select max(glId) as code from glMaster where deleted=0"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
mTo = rs("code")
rs.close

SQL = "select * from glControl"
rs.open sql, conn
acPeriod=rs("acPeriod")
acYear=rs("acYear")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>總賬細項列印</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.mPeriod.focus()">
<!-- #include file="menu.asp" -->
<div align="center"><center>
<b>總賬細項列印</b>
<form method="post" action="glTxListPrint.asp" name="form1">
<input type=hidden name="from" value="<% =Request.servervariables("script_name") %>">
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td width="100" align="right" class="b8">日期</td>
		<td width="10"></td>
		<td width="160">
			<select name="mPeriod" style="width:80px">
			<%
			for idx = acPeriod to 1 step -1
				if idx<=4 then
					m=idx+8
					y=acYear
				else
					m=idx-4
					y=acYear+1
				end if
				response.write "<option value="""&right("0"&m,2)&y&""">"&right("0"&m,2)&" "&y&"</option>"
			next
			for idx = 12 to 1 step -1
				if idx<=4 then
					m=idx+8
					y=acYear-1
				else
					m=idx-4
					y=acYear
				end if
				response.write "<option value="""&right("0"&m,2)&y&""">"&right("0"&m,2)&" "&y&"</option>"
			next
			%>
			</select>
		</td>
	</tr>
	<tr>
		<td align="right" class="b8">賬戶由</td>
		<td></td>
		<td>
			<input type="text" name="mFrom" value="<%=mFrom%>" maxlength=8 size=9>
		</td>
	</tr>
	<tr>
		<td align="right" class="b8">賬戶至</td>
		<td></td>
		<td>
			<input type="text" name="mTo" value="<%=mTo%>" maxlength=8 size=9>
		</td>
	</tr>
	<tr>
		<td align="right" class="b8">輸出</td>
		<td></td>
		<td>
			<select name="output" style="width:80px">
			<option value="html">Html
			<option value="text">Text
			<option value="word">Word
			<option value="excel">Excel
			</select>
			<input type="submit" value="確定" name="submit" class="sbttn">
		</td>
	</tr>
</table>
</form>
</center></div>
</body>
</html>