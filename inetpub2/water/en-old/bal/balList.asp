<%requiredLevel=3%>

<%
   mPeriod = year(date())&right("0"&month(date()),2)
%>
<html>
<head>
<title>每月帳統計列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.mPeriod.focus()">
<!-- #include file="menu.asp" -->
<div align="center"><center>
<br><b12>每月帳統計列表</b>
<form method="post" action="BalListPrintn.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">

        <tr>
                <td align="right" class="b12">月份</td>
                <td width="10"></td>
                <td><input type="text"  name="mPeriod" value="<%=mPeriod%>" size="6" >(YYYYMM)</td>
        </tr>    
	<tr>
		<td align="right" class="b12">輸出</td>
		<td width="10"></td>
		<td>
			<select name="output" style="width:88px">
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