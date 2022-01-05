
<%
   noofday =100
   stdate1 = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
%>
<html>
<head>
<title>冷戶報告</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.stdate1.focus()">
<!-- #include file="menu.asp" -->
<div align="center"><center>
<br><b>冷戶報告</b>
<form method="post" action="dormantListPrint.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">
        
        <tr>
		<td align="right" class="b8">日期</td>
		<td width="10"></td>
		<td>
                <input type="text" name="stdate1" value="<%=stdate1%>" size="10" maxlength="10" onblur="if(!formatDate(this)){this.value=''};">
		(dd/mm/yyyy)	
                </td> 

        </tr>  
	<tr>
		<td align="right" class="b8">輸出</td>
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