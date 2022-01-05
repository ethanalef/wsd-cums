<!-- #include file="../conn.asp" -->

<!-- #include file="navigator.asp" -->
<%
if request("del")<>"" then
	lnnum = request("del")
	sql1 = "select * from memtx where lnnum="&lnnum
	Set rs1 = Server.CreateObject("ADODB.Recordset")
	rs1.open sql1, conn         
        if rs1.eof then
		conn.execute("delete loanrec  where lnnum="&lnnum)

		lnnum =""
        else
 		msg = lnnum&" 已借出貸款,不能刪除"
        end if	
end if
For Each Field in Request.Form
	TheString = Field & "= Request.Form(""" & Field & """)"
	Execute(TheString)
Next
For Each Field in Request.querystring
	TheString = Field & "= Request.querystring(""" & Field & """)"
	Execute(TheString)
Next
if lnnum <> "" then
	sql_filter = sql_filter & " and a.lnnum = "&lnnum
end if
if memNo <> "" then
	sql_filter = sql_filter & " and a.memNo = '"&memNo&"'"
end if
if memName <> "" then
	sql_filter = sql_filter & " and b.memName like '%"&memName&"%'"
end if
if memhkid <> "" then
	sql_filter = sql_filter & " and b.memhkid >= '"&memhkid&"'"
end if

chk="Approved"
sql = "select a.memno,a.lnnum,a.lndate,a.appamt,b.memname,b.memhkid from loanrec a,memmaster b where a.memno=b.memno and cleardate=""" & sql_filter & " order by lnnum desc"
set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 3


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
<title>新貸款數建立</title>
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
<form name="form1" method="post" action="lcloan.asp">
貸款編號< : <input type="text" name="lnnum" value="<%=lnnum%>" size="10">
身分證號碼 : <input type="text" name="memhkid" value="<%=memhkid%>" size="10">
社員姓名 : <input type="text" name="memNo" value="<%=memNo%>" size="6">
社員姓名 : <input type="text" name="memName" value="<%=memName%>" size="10">
<input type="submit" value="Search" onclick="return validating()" class="sbttn">
</form>
<%if recordcount>pagesize then navigator("Lcloan.asp?lnnum="&lnnum&"&memhkid="&memhkid&"&memNo="&memNo&"&memName="&memName) end if%>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		<td><font size="2" color="#FFFFFF">貸款編號</font></td>
		<td><font size="2" color="#FFFFFF">身分證號碼</font></td>
		<td><font size="2" color="#FFFFFF">社員編號</font></td>
		<td><font size="2" color="#FFFFFF">社員姓名</font></td>
		<td><font size="2" color="#FFFFFF">貸款額</font></td>
<%if session("userLevel")<>5 then%>
		<td bgcolor="#FFFFFF"><a href="lcloandetail.asp"><font size="2">新增</font></a></td>
<%end if%>
	</tr>
<%
do while not rs.eof and rowcount < pagesize
	rowcount = rowcount + 1
%>
  <tr bgcolor="#FFFFFF">
	<td><a href="lloandetail.asp?lnnum=<%=rs("lnnum")%>"><font size="2"><%=rs("lnnum")%></font></a></td>
	<td><font size="2"><%=rs("memhkid")%></font></td>
	<td><font size="2"><%=rs("memno")%></font></td>
	<td><font size="2"><%=rs("memName")%></font></td>
	<td align="right"><font size="2"><%=formatNumber(rs("appAmt"),2)%></font></td>
<%if session("userLevel")<>5 then%>
	<td><a href="lcloan.asp?del=<%=rs("lnnum")%>" onclick="return confirm('刪除此紀錄?')"><font size="2">刪除</font></a></td>
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
