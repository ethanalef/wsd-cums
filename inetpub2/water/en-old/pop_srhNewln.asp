<!-- #include file="../conn.asp" -->
<!-- #include file="navigator.asp" -->
<%


key = request("key")

if key<>"" then
   if left(key,1)>="0" and left(key,1) <="9" then
	stylefield = "and a.memNo like '" & key & "%'"           
   else
	stylefield = "and a.memName like '" & key & "%'"
   end if 
end if

SQL = "select a.memNo,a.memname,b.uid,b.loanAmt from memmaster a,loanApp b where a.memNo=b.memno and b.deleted=0 and (b.firstApproval='Approved' or b.SecondApproval='Approved')  "& stylefield &" order by 1,3 desc " 
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 3
if rs.recordcount=1 then
%>
<script language="JavaScript">
<!--
	window.opener.document.form1.memNo.value='<% =rs("memNo") %>'
 
	window.opener.document.form1.memName.value='<% =rs("memName") %>'
        window.opener.document.form1.uid.value='<%=rs("uid")%>'
	window.opener.document.form1.Search.click()        
	
	window.close();
//-->
</script>
<%
	response.end
end if

if not rs.eof then
	if request("page") <> "" then
		pageno = cint(request("page"))
	else
		pageno = 1
	end if
	rs.pagesize = 15
	pagesize=rs.pagesize
	rs.absolutepage = pageno
	recordcount=rs.recordcount
	pagecount = rs.pagecount
end if
%>
<html>
<head>
<meta1 HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<title>Member Popup</title>
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function PostCode(memNo,memName,uid)
{
	window.opener.document.form1.memNo.value=memNo;  
	window.opener.document.form1.memName.value=memName;
	window.opener.document.form1.uid.value=uid;
        window.opener.document.form1.Search.click()        
	
	window.close();
}
//-->
</script>
</head>
<body leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="0" width="100%" bgcolor="#87CEEB">
	<form method="POST" action="<%=request.servervariables("script_name")%>" name="form">
    <tr>
        <td height="25">
			&nbsp;&nbsp;<b>搜尋</b>
			<input type="text" name="key" size="20" maxlength="50">
			<input type="submit" value="確定" class="sbttn" name="send">
        </td>
        <td align="right">
            <font face="arial" style="color: black; font: bold" size="3">搜尋社員</font>&nbsp;&nbsp;
        </td>
    </tr>
	</form>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<%
if key <> "" then
%>
	<tr>
		<td align="center">"<% =key %>" 的搜尋結果</td>
	</tr>
<%
end if
%>
	<tr>
		<td>
<%
if not rs.eof then
	navigator(request.servervariables("script_name")&"?key="&key)
%>
		<table border="0" cellpadding="3" cellspacing="1" width="100%" bgcolor="#336699">
<%
	do while not rs.eof and rowcount < rs.pagesize
		rowcount = rowcount + 1
%>
			<tr>
				<td bgcolor="#ffffff"><a href="JavaScript:PostCode('<%=rs(0)%>','<%=rs(1)%>' ,'<%=rs(2)%>'   )"><% =rs(0) %></a></td>
				<td bgcolor="#ffffff"><% =rs(1) %></a></td>
                                <td bgcolor="#ffffff"><% =rs(2) %></a></td>	
<%if rs(3)="C" then idx="Closed" else idx="     " end if %>			        
				<td bgcolor="#ffffff"><% =idx %></a></td>	
				</tr>
<%
		rs.movenext
	loop
%>
		</table>
<%
	navigator(request.servervariables("script_name")&"?key="&key)
	response.write "<p>"
else
%>
		<br><p align="center"><font size="4">無符合紀錄</font></p>
<%
end if
%>
		</td>
	</tr>
</table>
</center>
</div>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>