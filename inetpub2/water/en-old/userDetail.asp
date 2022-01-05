<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
arrLevel = Array("Inactive","Member","Operator","Supervisor","Administrator","Auditor","Preview")

if request.form("back") <> "" then
	response.redirect "user.asp"
end if

id = request("id")

if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
    set rs = server.createobject("ADODB.Recordset")
	msg = ""

	if id="" then
		sql = "select count(*) from loginUser where username='" & username & "'"
		rs.open sql, conn
		if rs(0) > 0 then
			msg = "用戶名稱已經存在 "
		end if
		rs.close
	end if

	if msg="" then
		conn.begintrans
		if id = "" then
			sql = "select top 1 * from loginUser order by uid desc"
		else
			sql = "select * from loginUser where uid=" & id
		end if
		rs.open sql, conn, 2, 2
		if id = "" then
			if rs.eof then
				id = 1
			else
				id = rs("uid") + 1
			end if
			rs.addnew
			rs("username") = username
			rs("uid") = id
			addUserLog "Add User"
		else
			addUserLog "Modify User"
		end if
		rs("userLevel") = cdbl(userLevel)
		if password<>"" then rs("password") = password end if
		rs.update
		conn.committrans
		msg = "紀錄已更新"
	end if
else
	if id <> "" then
		sql = "select * from loginUser where uid=" & id
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if rs.eof then
			response.redirect "user.asp"
		else
			username=rs("username")
			userLevel=rs("userLevel")
			if userLevel=5 then
				response.redirect "user.asp"
			end if
		end if
	end if
end if
%>
<html>
<head>
<title>用戶管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.username.value==""){
		reqField=reqField+", 用戶名稱";
		if (!placeFocus)
			placeFocus=formObj.memNo;
	}

<%if id="" then%>
	if (formObj.password.value==""){
		reqField=reqField+", 密碼";
		if (!placeFocus)
			placeFocus=formObj.password;
	}
<%end if%>

	if (formObj.password.value!=formObj.password1.value){
		reqField=reqField+", 相符的重入密碼";
		if (!placeFocus)
			placeFocus=formObj.password;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.username.focus()">
<!-- #include file="menu.asp" -->
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form name="form1" method="post" action="userDetail.asp">
<input type="hidden" name="id" value="<%=id%>">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td class="b8" align="right">用戶名稱</td>
		<td width=10></td>
		<td><input type="text" name="username" value="<%=username%>" size="50"<%if id<>"" then response.write " onfocus=""form1.password.focus();""" end if%>></td>
	</tr>
	<tr>
		<td class="b8" align="right">密碼</td>
		<td width=10></td>
		<td><input type="password" name="password" value="<%=password%>" size="50" maxlength="50"></td>
	</tr>
	<tr>
		<td class="b8" align="right">重入密碼</td>
		<td width=10></td>
		<td><input type="password" name="password1" value="<%=password1%>" size="50" maxlength="50"></td>
	</tr>
	<tr>
		<td class="b8" align="right">用戶級別</td>
		<td width=10></td>
		<td>
			<select name="userLevel">
<%for idx = 0 to 6%>
			<option value=<%=idx%> <%if idx=userLevel then response.write " selected" end if%>><%=arrLevel(idx)%></option>
<%next%>
			</select>
		</td>
	</tr>
	<tr>
		<td colspan="3" align="right">
			<input type="submit" value="確定" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
			<input type="submit" value="返回" name="back" class="sbttn">
		</td>
	</tr>
</table>
</center>
</form>
</body>
</html>
