<%requiredLevel=3%>

<html>
<head>
<%
     minpaid = 50
%>
<title>�Ȧ���b�ϺЫإ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.output.focus()">
<%' <!-- #include file="menu.asp" --> %>
<div align="center"><center>
<br><b>�Ȧ���b�ϺЫإ�</b>
<form method="post" action="atovListPrint.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">
        <tr>
                <td align="right" class="b16">�̧C���B(�Ѫ�)</td>  
                <td width="10"></td>
                <td><input type="text" name="minpaid" value="<%=minpaid%>" size="4" >   
        </tr> 
	<tr>
		<td align="right" class="b16">��X</td>
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