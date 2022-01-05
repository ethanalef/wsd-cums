<!-- #include file="../conn.asp" -->

<!-- #include file="navigator.asp" -->
<%
   username = session("username")
   userlevel = session("userlevel")
if request("del")<>"" then
	bncode = request("del")
	set rs = server.createobject("ADODB.Recordset")
	set rs =conn.execute("delete bank where bncode='"&bncode&"'  ")
	msg = uId&" 已刪除"
	
end if

For Each Field in Request.Form
	TheString = Field & "= Request.Form(""" & Field & """)"
	Execute(TheString)
Next
For Each Field in Request.querystring
	TheString = Field & "= Request.querystring(""" & Field & """)"
	Execute(TheString)
Next
if bncode <> "" then
	sql_filter = sql_filter & " where bncode = '"&bncode&"'  "
end if

if bank <> "" then
   if  sql_filter <>""  then          
	sql_filter = sql_filter & " and bank like '%"&bank&"%' "
   else
        sql_filter = sql_filter & " where bank like '%"&bank&"%'  "
   end if
end if



sql = "select * from bank   " & sql_filter 
set rs = server.createobject("ADODB.Recordset")
rs.open sql, conn, 3



if not rs.eof then
	if request("page") <> "" then
		pageno = cint(request("page"))
	else
		pageno = 1
	end if
	rs.pagesize = 10
	pagesize=rs.pagesize
	rs.absolutepage = pageno
	recordcount=rs.recordcount
	pagecount = rs.pagecount
end if
%>
<html>
<head>
<title>銀行資料操作</title>
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
<input type="hidden" name="Approval" value="<%=Approval%>">
<br>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<form name="form1" method="post" action="bank.asp">

銀行編號 : <input type="text" name="bncode" value="<%=bncode%>" size="6">
銀行名稱 : <input type="text" name="bank" value="<%=bank%>" size="40">

		
<input type="submit" value="搜尋" onclick="return validating()" class="sbttn">
</form>
<%if recordcount>pagesize then navigator("bank.asp?bncode="&bncode&"&bank="&bank) end if%>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">

		<td><font size="2" color="#FFFFFF">銀行編號</font></td>
		<td><font size="2" color="#FFFFFF">銀行名稱</font></td>

<%if session("userLevel")<>5 then%>
		<td bgcolor="#FFFFFF"><a href="bankDetail.asp"><font size="2">新增</font></a></td>
<%end if%>
	</tr>
<%
do while not rs.eof and rowcount < pagesize
    rowcount = rowcount + 1
%>
  <tr bgcolor="#FFFFFF">
	<td><a href="bankDetail.asp?bncode=<%=rs("bncode")%>"><font size="2"><%=rs("bncode")%></font></a></td>
	<td><font size="2"><%=rs("bank")%></font></td>
        <td><a href="bank.asp?del=<%=rs("bncode")%>" onclick="return confirm('確定刪除?')"><font size="2">刪除</font></a></td>

  </tr>
<%

	rs.movenext
loop
%>
</table>
</center>
</body>
</html>

