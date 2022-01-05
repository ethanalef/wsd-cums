<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "handleParty.asp"
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
		sql = "select count(*) from handleParty where handleName='" & handleName & "'"
		rs.open sql, conn
		if rs(0) > 0 then
			msg = "委員名稱已經存在 "
		end if
		rs.close
	end if

	if msg="" then
		conn.begintrans
		if id = "" then
			sql = "select top 1 * from handleParty order by uid desc"
		else
			sql = "select * from handleParty where uid=" & id
		end if
		rs.open sql, conn, 2, 2
		if id = "" then
			if rs.eof then
				id = 1
			else
				id = rs("uid") + 1
			end if
			rs.addnew
			rs("handleName") = handleName
			rs("uid") = id
			addUserLog "Add Committee"
		else
			addUserLog "Modify Committee"
		end if
		rs("status") = cint(status)
		rs.update
		conn.committrans
		msg = "紀錄已更新"
	end if
else
	if id <> "" then
		sql = "select * from handleParty where uid=" & id
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if rs.eof then
			response.redirect "handleParty.asp"
		else
			handleName=rs("handleName")
			status=rs("status")
			if userLevel=5 then
				response.redirect "handleParty.asp"
			end if
		end if
	end if
end if
%>
<html>
<head>
<title>委員資料修正</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.username.value==""){
		reqField=reqField+", 委員名稱";
		if (!placeFocus)
			placeFocus=formObj.memNo;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.handleName.focus()">
<!-- #include file="menu.asp" -->
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form name="form1" method="post" action="handlePartyDetail.asp">
<input type="hidden" name="id" value="<%=id%>">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td class="b8" align="right">委員名稱</td>
		<td width=10></td>
		<td><input type="text" name="handleName" value="<%=handleName%>" size="50"<%if id<>"" then response.write " onfocus=""this.blur();""" end if%>></td>
	</tr>
	<tr>
		<td class="b8" align="right">狀況</td>
		<td width=10></td>
		<td>
			<select name="status">
			<option value=1<%if status=1 then response.write " selected" end if%>>Active</option>
			<option value=0<%if status=0 then response.write " selected" end if%>>Inactive</option>
			</select>
		</td>
	</tr>
	<tr>
		<td colspan="3" align="right">
			<%if session("userLevel")<>5 then%>
			<input type="submit" value="確定" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
			<%end if%>
			<input type="submit" value="返回" name="back" class="sbttn">
		</td>
	</tr>
</table>
</center>
</form>
</body>
</html>
