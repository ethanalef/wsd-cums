<!-- #include file="../conn.asp" -->

<!-- #include file="navigator.asp" -->
<%
if request("del")<>"" then
	uid = request("del")
'	sql1 = "select * from memtx where lnnum="&uid
	Set rs1 = Server.CreateObject("ADODB.Recordset")
	rs1.open sql1, conn         
        if rs1.eof then
		conn.execute("delete loanrec  where lnnum="&uid)

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
IF REQUEST("NPAGE") <> "" OR REQUEST("UPAGE") <>"" THEN
   SQL_FILTER  =SESSION("STRSQL")
END IF


chk="Approved"
sql = "select * from loanApp where (SecondApproval='Approved' or firstApproval='Approved') and deleted=0  " & sql_filter & " order by uid desc"
set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 3

'if rs.recordcount=0 then
'	response.redirect "nloanDetail.asp"
'end if

if not rs.eof then
	if request("npage") <> "" then          
           pageno = session("cpageno")+1
           curpage = pageno
        
	else
		pageno = 1
                curpage = 0
	end if
	if request("upage") <> "" then          
           pageno = session("cpageno")-1
           curpage = pageno
        

	end if
	rs.pagesize = 20
	pagesize=rs.pagesize
	rs.absolutepage = pageno
	recordcount=rs.recordcount
	pagecount = rs.pagecount
        rowcount  = 0
        session("cpageno") = pageno
        SESSION("STRSQL")=sql_filter
end if
%>
<html>
<head>
<title>新貸款建立</title>
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
''		<td bgcolor="#FFFFFF"><a href="nloanDetail.asp"><font size="2">新增</font></a></td>
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<form name="form1" method="post" action="nloan.asp">
<div><center><font size="3">新貸款建立</font></center></div>
申請序號< : <input type="text" name="uid" value="<%=uid%>" size="10">
日期 : <input type="text" name="date1" value="<%=date1%>" size="10" onblur="if(!formatDate(this)){this.value=''};">
- <input type="text" name="date2" value="<%=date2%>" size="10" onblur="if(!formatDate(this)){this.value=''};">
社員姓名 : <input type="text" name="memNo" value="<%=memNo%>" size="6">
社員姓名 : <input type="text" name="memName" value="<%=memName%>" size="10">

<input type="submit" name="memSearch" value="搜尋" class="sbttn" >
</form>
<% if request.form("memSearch")<>"" or curpage > 0  Then %>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		<td><font size="2" color="#FFFFFF">申請序號</font></td>
		<td><font size="2" color="#FFFFFF">日期</font></td>
		<td><font size="2" color="#FFFFFF">社員編號</font></td>
		<td><font size="2" color="#FFFFFF">社員姓名</font></td>
		<td><font size="2" color="#FFFFFF">貸款額</font></td>

	</tr>
<%
do while not rs.eof and rowcount < pagesize
	rowcount = rowcount + 1
%>
  <tr bgcolor="#FFFFFF">
	<td><a href="nloanDetail.asp?uid=<%=rs("uid")%>"><font size="2"><%=rs("uid")%></font></a></td>
	<td><font size="2"><%=rs("appDate")%></font></td>
	<td><font size="2"><%=rs("memNo")%></font></td>
	<td><font size="2"><%=rs("memName")%></font></td>
	<td align="right"><font size="2"><%=formatNumber(rs("loanAmt"),2)%></font></td>
<%if session("userLevel")<>5 then%>
	<td><a href="nloan.asp?del=<%=rs("uid")%>" onclick="return confirm('刪除此紀錄?')"><font size="2">刪除</font></a></td>
<%end if%>
  </tr>
<%
	rs.movenext
loop
%>
</table>
<%if session("cpageno")>1 then%>
    <a href="ccloan.asp?upage=upage<font size="2">上一頁</font></a>
<%end if%>
<%if session("cpageno")< pagecount then%>
<a href="ccloan.asp?npage=npage<font size="2">下一頁</font></a>
<%end if %>
<%end if%>
</center>
</body>
</html>
