<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
SQL = "select distinct memSection from memMaster order by memSection"
Set Rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
%>
<html>
<head>
<title>社員列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="javascript">
<!--
function CA(){
    for (var i=0;i<document.form1.elements.length;i++){
        var e = document.form1.elements[i];
        if ((e.name != 'allbox') && (e.type=='checkbox')){
            e.checked = document.form1.allbox.checked;
        }
    }
}
//-->
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.mActive.focus()">
<!-- #include file="menu.asp" -->
<div align="center"><center>
<br><b>社員列表</b>
<form method="post" action="memListPrint.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td width="150" align="right" class="b8">只列印活躍會員</td>
		<td width="10"></td>
		<td width="160">
			<select name="mActive" style="width:80px">
			<option>Yes
			<option>No
			</select>
		</td>
	</tr>
	<tr>
		<td align="right" class="b8">輸出</td>
		<td></td>
		<td>
			<select name="output" style="width:80px">
			<option value="html">Html
			<option value="word">Word
			<option value="excel">Excel
			</select>
		</td>
	</tr>
</table>
<br>
<table border="0" cellpadding="0" cellspacing="0">
    <tr>
        <td colspan="4" height="30" valign="top"><input name="allbox" type="checkbox" onClick="CA();"> <i>選擇全部</i> </td>
    </tr>
    <tr>
<%
locate=1
do while not rs.eof %>
        <td width="30"><input type="checkbox" name="TS" value="<% =rs("memSection") %>"></td>
        <td width="80"><%=rs("memSection")%></td>
<%
	if locate=1 then
		locate = 0
	else
		locate = 1
		response.write "</tr><tr>"
	end if
	rs.movenext
loop%>
	</tr>
    <tr>
        <td colspan="4" height="30" valign="bottom" align="right"><input type="submit" value="確定" name="submit" class="sbttn"></td>
    </tr>
</table>
</form>
</center></div>
</body>
</html>
<%
rs.close
set rs = nothing
conn.close
set conn = nothing
%>