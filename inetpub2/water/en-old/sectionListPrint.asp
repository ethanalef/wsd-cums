<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%

   mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())

xname = request.form("accode")
pos = instr(xname,"-")
if pos > 0 then
accode = left(xname,pos-1)
mname =  mid(xname,pos+1,50)
else
accode=""
mname =""
end if
sql = "select memno,memname,memcname, memofficetel,memMobile from memMaster where memno='"&accode&"'   "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 3
if not rs.eof then
   mobile = rs("memMobile")
   offtel = rs("memofficetel")
end if
rs.close

sql = "select memno,memname,memcname,memofficetel,memGrade,memSection  from memMaster where accode='"&accode&"' and mstatus not in ('C','P','B') order by memno "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 3

if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"

end if
%>
<html>
<head>
<title>Section List</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<LINK href="../main.css" rel=STYLESHEET type=text/css>
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="3"  face="標楷體" >水務署員工儲蓄互助社<br>聯絡員列表<br><font size="2"  face="標楷體" >日期 : <%=mndate%><br><br><font size="3"  face="標楷體" >聯絡員 : <%=mname%>   辦公室電話 : <%=offtel%>  手提電話 : <%=mobile%></font></font></font></td></tr>
        

</table>

<br>
          

<table border="0" cellpadding="0" cellspacing="0">
	<tr height="20" valign="bottom">
		<td width=70 align="center"><font size="3"  face="標楷體" >社員編號</font></td>
                <td width=180 align="center"><font size="3"  face="標楷體" >英文姓名</font></td>
		<td width=70 align="center"><font size="3"  face="標楷體" >中文姓名</font></td>
                <td width=100 align="center"><font size="3"  face="標楷體" >職位</font></td>
		<td width=100 align="center"><font size="3"  face="標楷體" >部門</font></td>
                <td width=100 align="center"><font size="3"  face="標楷體" >辦公室電話</font></td>
	</tr>
	<tr><td colspan=7><hr></td></tr>
      
<%
thisSection=""
do while not rs.eof
   set ms = conn.execute("select * from share where memno='"&rs("memno")&"' order by memno,ldate,code ")
   ttlbal = 0
   if not ms.eof then
           do while not ms.eof
              select case ms("code")
                     case "0A","A1","A2","A3","C0","C1","C3" ,"B6" ,"A0","A7" ,"A4" 
                          ttlbal = ttlbal + ms("amount")
                      case "G0" ,"H0","B0","B1","B3","BE","BF","G3","H3" ,"MF"
                           ttlbal = ttlbal -ms("amount")
               end select
                      
           ms.movenext
           loop
   end if
   ms.close

    if ttlbal > 0 then
%>


	<tr>
		<td width=70 align="center"><%=rs("memNo")%></td>
                <td width=180 align="left"><%=rs("memname")%></td> 
		<td width=70 align="center"><font size="3"  face="標楷體" ><%=rs("memcname")%></font> </td>
                <td width=100 align="center"> <%=rs("memGrade")%></td>
                <td width=100 align="center"><%=rs("memSection")%></td>
                <td width=100 align="center"><%=rs("memofficetel")%></td>

	</tr>
<%
        end if
	rs.movenext
loop
%>
</table>
</center>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
