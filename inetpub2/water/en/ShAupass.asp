<!-- #include file="../CheckUserStatus.asp" -->
<%requiredLevel=3%>

<%
   mPeriod = year(date())&right("0"&month(date()),2)
%>
<html>
<head>
<title>�Ȧ欣���ϺЫإ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.mPeriod.focus()">
<!-- #include file="menu.asp" -->
<div align="center"><center>
<br><b12>�Ȧ欣���ϺЫإ�</b>
<form method="post" action="genshpay.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">

   
	<tr>
		<td align="right" class="b12">��X</td>
		<td width="10"></td>
		<td>

			<input type="submit" value="�T�w" name="submit" class="sbttn">
		</td>
	</tr>
</table>
</form>
</center></div>
</body>
</html>