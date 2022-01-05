<!-- #include file="../CheckUserStatus.asp" -->
<%requiredLevel=3%>

<%
   mPeriod = year(date())&right("0"&month(date()),2)
%>
<html>
<head>
<title>股息列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.mPeriod.focus()">
<!-- #include file="menu.asp" -->
<div align="center"><center>
<br><b12>股息列表</b>
<form method="post" action="DivdListPrint.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td align="right" class="b8">社員現況</td>
		<td width="30"></td>
		<td>
			<select name="KIND" style="width:150px">
			<option value="N">正常
			<option value="L">呆帳
			<option value="D">冷戶
			<option value="V">IVA
			<option value="C">退社
 			<option value="x">凍結
			<option value="p">去世
			<option value="B">破產
			<option value="J">新戶
			<option value="T">庫房
			<option value="H">暫停銀行
			<option value="A">自動轉帳(ALL)
			<option value="0">自動轉帳(股金)
			<option value="1">自動轉帳(股金,利息)
			<option value="Z">自動轉帳(股金,本金)
			<option value="3">自動轉帳(利息,本金)
			<option value="M">庫房,銀行
			<option value="F">特別個案
                        <option value="8">終止社籍轉帳
                        <option value="9">終止社籍正常
			<option value="all">全選
			</select>

		</td>
	</tr>
	<tr>
		<td align="right" class="b8">派息狀況</td>
		<td width="30"></td>
		<td>
			<select name="bank">
                        <option></option>
			<option value="S" <%if bank="S" then response.write " selected" end if%>>股金</option>
			<option value="B" <%if bank="B" then response.write " selected" end if%>>銀行轉帳</option>
                        <option value="C" <%if bank="C" then response.write " selected" end if%>>支票</option>
			<option alue="H" <%if bank="H" then response.write " selected" end if%>>暫停派息</option>
			<option value="A" >全選
			</select>

		</td>
	</tr>
   
	<tr>
		<td align="right" class="b12">輸出</td>
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