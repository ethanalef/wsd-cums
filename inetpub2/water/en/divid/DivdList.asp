<!-- #include file="../CheckUserStatus.asp" -->
<%requiredLevel=3%>

<%
   mPeriod = year(date())&right("0"&month(date()),2)
%>
<html>
<head>
<title>�Ѯ��C��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.mPeriod.focus()">
<!-- #include file="menu.asp" -->
<div align="center"><center>
<br><b12>�Ѯ��C��</b>
<form method="post" action="DivdListPrint.asp" name="form1">
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td align="right" class="b8">�����{�p</td>
		<td width="30"></td>
		<td>
			<select name="KIND" style="width:150px">
			<option value="N">���`
			<option value="L">�b�b
			<option value="D">�N��
			<option value="V">IVA
			<option value="C">�h��
 			<option value="x">�ᵲ
			<option value="p">�h�@
			<option value="B">�}��
			<option value="J">�s��
			<option value="T">�w��
			<option value="H">�Ȱ��Ȧ�
			<option value="A">�۰���b(ALL)
			<option value="0">�۰���b(�Ѫ�)
			<option value="1">�۰���b(�Ѫ�,�Q��)
			<option value="Z">�۰���b(�Ѫ�,����)
			<option value="3">�۰���b(�Q��,����)
			<option value="M">�w��,�Ȧ�
			<option value="F">�S�O�Ӯ�
                        <option value="8">�פ���y��b
                        <option value="9">�פ���y���`
			<option value="all">����
			</select>

		</td>
	</tr>
	<tr>
		<td align="right" class="b8">�������p</td>
		<td width="30"></td>
		<td>
			<select name="bank">
                        <option></option>
			<option value="S" <%if bank="S" then response.write " selected" end if%>>�Ѫ�</option>
			<option value="B" <%if bank="B" then response.write " selected" end if%>>�Ȧ���b</option>
                        <option value="C" <%if bank="C" then response.write " selected" end if%>>�䲼</option>
			<option alue="H" <%if bank="H" then response.write " selected" end if%>>�Ȱ�����</option>
			<option value="A" >����
			</select>

		</td>
	</tr>
   
	<tr>
		<td align="right" class="b12">��X</td>
		<td width="10"></td>
		<td>
			<select name="output" style="width:88px">
			<option value="html">Html
			<option value="text">Text
			<option value="word">Word
			<option value="excel">Excel
			</select>
			<input type="submit" value="�T�w" name="submit" class="sbttn">
		</td>
	</tr>
</table>
</form>
</center></div>
</body>
</html>