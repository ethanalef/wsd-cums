<%requiredLevel=3%>
<%
    styr = 1900
    edyr = year(date())
    myear = edyr - 1
%>   
<html>
<head>
<title>年結分析報告</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.output.focus()">

<div align="center"><center>
<br><b>年結分析報告</b>
<form method="post" action="yrendstatPrint.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td align="right" class="b8">年份</td>
		<td width="10"></td>
                <td> 
                <select name="myear">
<%
                 xx = 1900
                 do while xx <= edyr 
%>                      
		   <option<% if myear=xx then %> selected<% end if%>><%=xx%>
                   
                   
<%
            
                 xx = xx + 1
   
                  loop
%>   
                 
                 </select>
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