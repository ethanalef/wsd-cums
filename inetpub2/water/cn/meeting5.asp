<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
loan=request("loan")

if request("from") = Request.ServerVariables("script_name") then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
	bdNum=cint(bdNum)
	if request.form("submit")<>"" then
	    set rs = server.createobject("ADODB.Recordset")
		conn.begintrans
		for idx = 1 to bdNum
			Set rs = Server.CreateObject("ADODB.Recordset")
			if request.form("bd1c"&idx) = "" then
				sql = "select top 1 * from meetingNotes5 order by uid desc"
			else
				sql = "select * from meetingNotes5 where uid="&request.form("bd1c"&idx)
			end if
			rs.open sql, conn, 1, 3
			if request.form("bd1c"&idx) = "" then
				if rs.eof then
					id = 1
				else
					id = rs("uid") + 1
				end if
				rs.addnew
				rs("uid") = id
			end if
			rs("rpId") = loan
			rs("memNo") = request.form("bd2c"&idx)
			rs("memName") = request.form("bd3c"&idx)
			if request.form("bd4c"&idx)<>"" then rs("amount") = request.form("bd4c"&idx) else rs("amount")=0 end if
			rs("description") = request.form("bd5c"&idx)
			rs.update
			TheString="bd1c"&idx&"="""&rs("uid")&""""
			Execute(TheString)
			archive=archive&rs("uid")&","
			rs.close
		next
		if archive="" then
			conn.execute("delete from meetingNotes5 where rpId="&loan)
		else
			conn.execute("delete from meetingNotes5 where rpId="&loan&" and uid not in ("&left(archive,len(archive)-1)&")")
		end if
		conn.committrans
		addUserLog "Modify Meeting Notes"
		response.redirect "meetingNotesDetail.asp?id="&loan
	elseif request.form("back") <> "" then
		response.redirect "meetingNotesDetail.asp?id="&loan
	elseif request.form("add")<>"" then
		bdNum=bdNum+1
	else
		checkDel = false
		for idx = 1 to bdNum
			if request.form("del"&idx)<>"" then
				delNum=idx
				bdNum=bdNum-1
				checkDel = true
				exit for
			end if
		next
	end if
	focusmsg = "form1.bd2c"&bdNum+1&".focus()"
else
	SQL = "select * from meetingNotes5 where rpId=" & loan
	set rs = server.createobject("ADODB.Recordset")
	rs.open sql, conn
	if rs.eof then
		bdNum = 0
	else
		idx=1
		do while not rs.eof
			theString = "bd1c"&idx&"="&rs("uid")
			Execute(TheString)
			theString = "bd2c"&idx&"="&rs("memNo")
			Execute(TheString)
			theString = "bd3c"&idx&"="""&rs("memName")&""""
			Execute(TheString)
			theString = "bd4c"&idx&"="&rs("amount")
			Execute(TheString)
			theString = "bd5c"&idx&"="""&rs("description")&""""
			Execute(TheString)
			idx = idx + 1
			rs.movenext
		loop
		bdNum = idx-1
	end if
	focusmsg =  "form1.bd2c1.focus()"
end if
%>
<html>
<head>
<title>會議紀錄</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function popup(filename){
  window.open (filename,'pop','width=500,height=550,statusbar=no,toolbar=no,resizable,scrollbars,dependent')
}

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

	if (formObj.bd2c<%=bdNum+1%>.value==""){
		reqField=reqField+", 社員編號";
		if (!placeFocus)
			placeFocus=formObj.bd2c<%=bdNum+1%>;
	}

	if (formObj.bd4c<%=bdNum+1%>.value==""){
		reqField=reqField+", 金額";
		if (!placeFocus)
			placeFocus=formObj.bd4c<%=bdNum+1%>;
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

function validating1(){
	validcheck=true
<%for idx = 1 to bdNum%>
	if (document.form1.bd2c<%=idx%>.value=="")
		validcheck=false
<%next%>

	if (!validcheck){
		alert("請填入所有社員編號")
		return false;
	}else{
		return true;
	}
}
//  -->
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="<%=focusmsg%>">
<!-- #include file="menu.asp" -->
<center>
<form method="post" action="<%=request.servervariables("script_name")%>" name="form1">
<input type="hidden" name="From" value="<%=Request.servervariables("script_name")%>">
<input type="hidden" name="loan" value="<%=loan%>">
<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="bottom" bgcolor="#87CEEB" height="17" align="center">
		<td class="n8" width=30>#</td>
		<td class="n8">社員編號</td>
		<td class="n8" width=200>姓名</td>
		<td class="n8">金額</td>
		<td class="n8">理由</td>
		<td></td>
	</tr>
	<input type="hidden" name="bdNum" value="<% =bdNum %>">
<%
for idx = 1 to bdNum+1
	if checkDel and idx >= delNum then
		ii=idx+1
	else
		ii=idx
	end if
%>
	<input type="hidden" name="bd1c<%=idx%>" value="<% =eval("bd1c"&ii) %>">
	<input type="hidden" name="bd3c<%=idx%>" value="<% =eval("bd3c"&ii) %>">
	<tr>
		<td align="center"><%=idx%></td>
		<td><input type=text class="show" name="bd2c<%=idx%>" value="<% =eval("bd2c"&ii) %>" size=6 onchange="document.all.tags('td')['bd3c<%=idx%>'].innerHTML='';form1.bd3c<%=idx%>.value='';popup('pop_searchMem2.asp?key='+this.value+'&editNum=<%=idx%>')""><input type="button" value="選擇" onclick="popup('pop_searchMem2.asp?key='+form1.bd2c<%=idx%>.value+'&editNum=<%=idx%>')" class="xbttn"></td>
		<td class="show" id="bd3c<%=idx%>"><% =eval("bd3c"&ii) %></td>
		<td><input type=text class="show" name="bd4c<%=idx%>" value="<% =eval("bd4c"&ii) %>" size=20 onblur="if(!formatNum(this)){this.value=''};"></td>
		<td><input type=text class="show" name="bd5c<%=idx%>" value="<% =eval("bd5c"&ii) %>" size=50 maxlength=50></td>
        <td>
<%if idx>bdNum then%>
<input type="submit" value="新增" name="add" class="xbttn" onclick="return validating()">
<%else%>
<input type="submit" value="刪除" name="del<%=idx%>" class="xbttn" onclick="return confirm('刪除此紀錄?')" style="width:26">
<%end if%>
		</td>
    </tr>
<%
next
%>
	<tr>
        <td colspan="19" align="right" height="30">
        	<input type="submit" value="儲存" name="submit" class="sbttn" onclick="return validating1()">
        	<input type="submit" value="返回" name="back" class="sbttn">
        </td>
	</tr>
</table>
</center>
</body>
</html>
