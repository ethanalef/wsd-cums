<%requiredLevel=3%>
<% divdate = day(date())&"/"&month(date())&"/"&year(date())
%>
<html>
<head>
<title>股金股息列印</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.output.focus()">
<!-- #include file="menu.asp" -->
<div align="center"><center>
<br><b>股息分配列印</b>
<form method="post" action="shdlstprt.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td align="right" class="b8">日期</td>
		<td width="10"></td>
		<td>

                      <input type="text" name="divdate" value="<%=divdate%>" size="10">
		</td>
	</tr>
	<tr>
		<td align="right" class="b8">輸出</td>
		<td width="10"></td>
		<td>
			<select name="output" style="width:88px">
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