<!-- #include file="../conn.asp" -->
<!-- #include file="init.asp" -->
<%
	yr = year(date())
	dvdDay = dmy(date())
	rate = 2
%>

<html>
	<head>
		<title>股息全年結(PDF)</title>
		<meta http-equiv="Content-Type" content="text/html; charset=big5">
		<link href="../main.css" rel="stylesheet" type="text/css">
	</head>
	<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
		<div align="center">
		<center>
			<br><b5>股息全年結(PDF)</b>
			<form method="post" action="shfypprint.asp" name="form1">
				<table border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td align="right" class="b12">年份</td>
						<td><input type="text" value="<%=yr%>" name="yr" size="4"></td>
					</tr>
					<tr>
						<td align="right" class="b12">股息日期</td>
						<td><input type="text" name="dvdDay" value="<%=dvdDay%>" size="10" maxlength="10" <%=working%> onblur="if(!formatDate(this)){this.value=''};callage();">(dd/mm/yyyy)</td>
					</tr>
					<tr>
						<td align="right" class="b12">股息率</td>
						<td><input type="text" value="<%=rate%>" name="rate" size="4"></td>
					</tr>
					<tr>
						<td align="right" class="b12">輸出</td>
						<td>
							<select name="output" style="width:80px">
								<option value="html">Html
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