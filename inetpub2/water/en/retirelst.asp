
<%
   noofday =100
%>
<html>
<head>
<title>退休社員報告</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.output.focus()">
<!-- #include file="menu.asp" -->
<div align="center"><center>
<br><b>退休社員報告</b>
<form method="post" action="retirelstprint.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">
 
        <tr>
              <td align="right" class="b8">股金結餘</td>
              <td width="10"></td>
              <td><input type="radio" checked="checked" name="idx" value="1"> > 0 
              <br>
              <input type="radio" name="idx" value="2">  = 0 
              <br>
              <input type="radio" name="idx" value="3">全選</td>
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