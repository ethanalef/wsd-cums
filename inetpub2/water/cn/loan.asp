<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="navigator.asp" -->
<%
if request("del")<>"" then
	uid = request("del")
'	set rs = server.createobject("ADODB.Recordset")
	conn.execute("delete meetingnotes0 where appId="&uId)
	conn.execute("delete meetingnotes1 where appId="&uId)
	conn.execute("update loanApp set deleted=-1 where uId="&uId)
	msg = uId&" 已刪除"
'	rs.close
end if

For Each Field in Request.Form
	TheString = Field & "= Request.Form(""" & Field & """)"
	Execute(TheString)
Next
For Each Field in Request.querystring
	TheString = Field & "= Request.querystring(""" & Field & """)"
	Execute(TheString)
Next
if uid <> "" then
	sql_filter = sql_filter & " and uid = "&uid
end if
if memNo <> "" then
	sql_filter = sql_filter & " and memNo = '"&memNo&"'"
end if
if memName <> "" then
	sql_filter = sql_filter & " and memName like '%"&memName&"%'"
end if
if date1 <> "" then
	sql_filter = sql_filter & " and appDate >= '"&date1&"'"
end if
if date2 <> "" then
	sql_filter = sql_filter & " and appDate <= '"&date2&"'"
end if

sql = "select * from loanApp where deleted=0 " & sql_filter & " order by uid desc"
set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 3

'if rs.recordcount=0 then
'	response.redirect "loanDetail.asp"
'end if

if not rs.eof then
	if request("page") <> "" then
		pageno = cint(request("page"))
	else
		pageno = 1
	end if
	rs.pagesize = 20
	pagesize=rs.pagesize
	rs.absolutepage = pageno
	recordcount=rs.recordcount
	pagecount = rs.pagecount
end if
%>
<html>
<head>
<title>貸款申請</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function formatDate(dateform){
  cDate = dateform.value;
  dSize = cDate.length;
  if (dSize!=0){
    sCount= 0;
    for(var i=0; i < dSize; i++)
      (cDate.substr(i,1) == "/") ? sCount++ : sCount;
    if (sCount == 2){
		ySize = cDate.substring(cDate.lastIndexOf("/")+1,dSize).length;
		if (ySize<2 || ySize>4 || ySize == 3){
		  return false;
		 }
		idxBarI = cDate.indexOf("/");
		idxBarII = cDate.lastIndexOf("/");
		strD = cDate.substring(0,idxBarI);
		strM = cDate.substring(idxBarI+1,idxBarII);
		strY = cDate.substring(idxBarII+1,dSize);
		strM = (strM.length < 2 ? '0'+strM : strM);
		strD = (strD.length < 2 ? '0'+strD : strD);
		if(strY.length == 2)
		  strY = (strY > 50  ? '19'+strY : '20'+strY);
    }else{
    	if (dSize != 8)
			return false;
		strD = cDate.substring(0,2);
		strM = cDate.substring(2,4);
		strY = cDate.substring(4,8);
    }
    dateform.value = strD+'/'+strM+'/'+strY;
    if (!valDate(strM, strD, strY))
      return false;
    else
      return true;
  }
}
//  -->
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<form name="form1" method="post" action="loan.asp">
Application No. : <input type="text" name="uid" value="<%=uid%>" size="10">
Date : <input type="text" name="date1" value="<%=date1%>" size="10" onblur="if(!formatDate(this)){this.value=''};">
- <input type="text" name="date2" value="<%=date2%>" size="10" onblur="if(!formatDate(this)){this.value=''};">
Membership No. : <input type="text" name="memNo" value="<%=memNo%>" size="6">
Member Name : <input type="text" name="memName" value="<%=memName%>" size="10">
<input type="submit" value="Search" onclick="return validating()" class="sbttn">
</form>
<%if recordcount>pagesize then navigator("loan.asp?uid="&uid&"&date1="&date1&"&date2="&date2&"&memNo="&memNo&"&memName="&memName) end if%>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		<td><font size="2" color="#FFFFFF">申請序號</font></td>
		<td><font size="2" color="#FFFFFF">日期</font></td>
		<td><font size="2" color="#FFFFFF">社員編號</font></td>
		<td><font size="2" color="#FFFFFF">社員姓名</font></td>
		<td><font size="2" color="#FFFFFF">貸款額</font></td>
<%if session("userLevel")<>5 then%>
		<td bgcolor="#FFFFFF"><a href="loanDetail.asp"><font size="2">新增</font></a></td>
<%end if%>
	</tr>
<%
do while not rs.eof and rowcount < pagesize
	rowcount = rowcount + 1
%>
  <tr bgcolor="#FFFFFF">
	<td><a href="loanDetail.asp?uid=<%=rs("uid")%>"><font size="2"><%=rs("uid")%></font></a></td>
	<td><font size="2"><%=rs("appDate")%></font></td>
	<td><font size="2"><%=rs("memNo")%></font></td>
	<td><font size="2"><%=rs("memName")%></font></td>
	<td align="right"><font size="2"><%=formatNumber(rs("loanAmt"),2)%></font></td>
<%if session("userLevel")<>5 then%>
	<td><a href="loan.asp?del=<%=rs("uid")%>" onclick="return confirm('刪除此紀錄?')"><font size="2">刪除</font></a></td>
<%end if%>
  </tr>
<%
	rs.movenext
loop
%>
</table>
</center>
</body>
</html>
