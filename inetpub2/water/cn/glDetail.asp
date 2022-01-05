<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "gl.asp"
end if

glId = request("glId")

if request.form("action") = "Save" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
    set rs = server.createobject("ADODB.Recordset")
	msg = ""
	if id<>glId then id="" end if

	if id="" then
		sql = "select count(*) from glMaster where glId='"&glId&"'"
		rs.open sql, conn
		if rs(0) > 0 then
			msg = "編號已經存在. "
		end if
		rs.close
	end if

	if msg="" then
		conn.begintrans
		if id = "" then
			sql = "select * from glMaster where 0=1"
		else
			sql = "select * from glMaster where glId='"&glId&"'"
		end if
		rs.open sql, conn, 2, 2
		if id = "" then
			rs.addnew
			rs("glId") = glId
			rs("creationDate") = now
			addUserLog "Add G/L Account"
		else
			addUserLog "Modify G/L Account"
		end if
		rs("glName") = glName
		rs("glType") = glType
		rs.update
		conn.committrans
		msg = "紀錄已更新"
	end if
else
	if glId <> "" then
		SQL = "select * from glMaster where glId='"&glId&"'"
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if rs.eof then
			response.redirect "gl.asp"
		else
			For Each Field in rs.fields
				TheString = Field.name & "= rs(""" & Field.name & """)"
				Execute(TheString)
			Next
		end if
	end if
end if
%>
<html>
<head>
<title>總帳</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function formatNum(numform){
  if (isNaN(numform.value)||numform.value<0)
    return false;
  else
    return true;
}

function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.glId.value==""){
		reqField=reqField+", 編號";
		if (!placeFocus)
			placeFocus=formObj.glId;
	}

	if (formObj.glName.value==""){
		reqField=reqField+", 內容";
		if (!placeFocus)
			placeFocus=formObj.glName;
	}

    if (reqField){
        if (reqField.lastIndexOf(",")==0)
	        reqField = "請填入"+reqField.substring(2);
        else
	        reqField = "請填入"+reqField.substring(2,reqField.lastIndexOf(","))+'及'+reqField.substring(reqField.lastIndexOf(",")+2);
        alert(reqField);
        placeFocus.focus();
        return false;
    }else{
        return true;
    }
}
//  -->
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.glId.select();form1.glId.focus()">
<!-- #include file="menu.asp" -->
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form name="form1" method="post" action="glDetail.asp">
<input type="hidden" name="id" value="<%=glId%>">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td width="100" class="b8">編號</td>
		<td width="300">
			<input type="text" name="glId" value="<%=glId%>" size="10" maxlength="10">
			<input type="submit" value="Search" name="Search" class="sbttn">
		</td>
	</tr>
	<tr>
		<td><font size="2" class="b8">內容</td>
		<td><font size="2"><input type="text" name="glName" value="<%=glName%>" size="50" maxlength="50"></font></td>
	</tr>
	<tr>
		<td><font size="2" class="b8">分類</td>
		<td>
			<select name="glType">
			<option value="1"<% if glType="1" then response.write " selected" end if%>>Fixed Assets</option>
			<option value="2"<% if glType="2" then response.write " selected" end if%>>Loans</option>
			<option value="3"<% if glType="3" then response.write " selected" end if%>>Current Assets</option>
			<option value="4"<% if glType="4" then response.write " selected" end if%>>Liabilities</option>
			<option value="5"<% if glType="5" then response.write " selected" end if%>>Income</option>
			<option value="6"<% if glType="6" then response.write " selected" end if%>>Expenses</option>
		</td>
	</tr>
	<tr>
		<td align="right" colspan="2">
			<%if session("userLevel")<>5 then%>
			<input type="submit" value="儲存" onclick="return validating()" name="action" class="sbttn">
			<%end if%>
			<input type="submit" value="返回" name="back" class="sbttn">
		</td>
	</tr>
</table>
<br>
<br>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
  <tr bgcolor="#330000" align="center">
	<td><font size="2" color="#FFFFFF">會計月</font></td>
	<td width="110"><font size="2" color="#FFFFFF">去年</font></td>
	<td width="110"><font size="2" color="#FFFFFF">今年</font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">1 (Sep)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastBal1,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisBal1,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">2 (Oct)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastBal2,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisBal2,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">3 (Nov)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastBal3,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisBal3,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">4 (Dec)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastBal4,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisBal4,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">5 (Jan)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastBal5,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisBal5,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">6 (Feb)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastBal6,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisBal6,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">7 (Mar)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastBal7,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisBal7,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">8 (Apr)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastBal8,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisBal8,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">9 (May)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastBal9,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisBal9,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">10 (Jun)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastBal10,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisBal10,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">11 (Jul)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastBal11,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisBal11,2)%></font></td>
  </tr>
  <tr bgcolor="#FFFFFF">
	<td><font size="2">12 (Aug)</font></td>
	<td align=right><font size="2"><%=formatnumber(lastBal12,2)%></font></td>
	<td align=right><font size="2"><%=formatnumber(thisBal12,2)%></font></td>
  </tr>
</table>
</center>
</form>
</body>
</html>
