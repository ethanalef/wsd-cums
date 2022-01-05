
<html>
<head>
<title>社員生日名單</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.mMonth.focus()">
<!-- #include file="menu.asp" -->
<div align="center"><center>
<br><b>社員生日名單</b>
<form method="post" action="memberListPrintN.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td width="100" align="right" class="b8">月</td>
		<td width="10"></td>
		<td width="160">
			<select name="mMonth" style="width:80px">
			<%
			for idx = 1 to 12
				if idx=month(now) then
					response.write "<option selected>"&idx
				else
					response.write "<option>"&idx
				end if
			next%>
			</select>
		</td>
	</tr>
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