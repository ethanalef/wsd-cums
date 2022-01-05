<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%

memfrom = request.form("mFrom")
memto   = request.form("mTo")
Dstart  =request.form("mStart")
Dend    = request.form("mEnd")
mstart = year(Dstart)&"/"&right("0"&month(Dstart),2)&"/"&right("0"&day(Dstart),2)
mend   = year(Dend)&"/"&right("0"&month(Dend),2)&"/"&right("0"&day(Dend),2)
response.write(dstart)
response.write(dend)
response.end
set rs=conn.execute("select a.*,b.memname,b.memcname from loanrec a,memmaster b where a.memno=b.memno and (a.memno >='"&memFrom&"' and a.memno<='"&mrmTo&"' )  ")

if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"
elseif request.form("output")="text" then
	spaces=""
	for idx = 1 to 50
		spaces=spaces&" "
	next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.CreateTextFile(Server.MapPath("..\txt")&"\"&session("username")&".txt", True)
	objFile.Write  left(space2,20)
	objFile.Write "水務署員工儲蓄互助社"
	objFile.WriteLine ""
	objFile.Write  left(space2,21)
	objFile.Write "貸款細明表"
	objFile.WriteLine ""	
	objFile.Write left("社員編號"&spaces,10)
	objFile.Write left("社員名稱"&spaces,25)
	objFile.Write right(spaces&"貸款編號",12)
	objFile.Write right(spaces&"設定日期",12)
	objFile.Write right(spaces&"貸款金額",12) 
	objFile.Write right(spaces&"巳批期數",6)
	objFile.Write right(spaces&"每月本金",12)
	objFile.Write right(spaces&"本金結餘",12)
	objFile.Write right(spaces&"貸款情況",12)


	objFile.WriteLine ""
	for idx = 1 to 130
		objFile.Write "-"
	next
       
	objFile.WriteLine ""   
 
        do while not rs.eof
           if rs("lndate") >= dstart  and rs("lndate") <= dend then
              xlndate=right("0"& day(rs("lndate")),2)&"/"&right("0"&month(rs("lndate")),2)&"/"&year(rs("lndate"))
              objFile.Write right(spaces&rs("memno"),5)
              objfile.Write left(rs("memname")&" "&rs("memcname")&spaces ,25)
	      objFile.Write left(rs("lnnum")&spaces,12) 
              objfile.Write left(xlndate&spaces,12) 
	      objfile.Write right(spaces&formatnumber(rs("appamt"),2),13)
              objfile.Write right(spaces&rs("install"),4) 	
              objFile.Write right(spaces&formatnumber(rs("monthrepay"),2),12)
	      objFile.Write right(spaces&formatnumber(rs("bal"),2),13)
              if rs("repaystat")="C" then
                 objFile.Write "完"
               else
                 objFile.Write "未完"
               end if   
                objFile.WriteLine "" 
        end if    
           
		rs.movenext
	loop
	for idx = 1 to 100
		objFile.Write "-"
	next
	objFile.WriteLine ""
	objFile.Write space(10)&"Total Count : "
	objFile.Write right(spaces&formatnumber(ttlcnt,2),20)
	objFile.WriteLine ""


	objFile.Close

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.redirect "../txt/"&session("username")&".txt"
end if
%>
<html>
<head>
<title>貸款細明表</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
	<td colspan="15"><font size="4">水務署員工儲蓄互助社</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
	<td colspan="15"><font size="4">貸款細明表</font></td>
	</tr>
	<tr height="15" valign="bottom">
	<td width="80"><b>社員編號</b></td>
	<td width="200"><b>社員名稱</b></td>
	<td width="130" align="right"><b>貸款編號</b></td>
	<td width="130" align="right"><b>設定日期</b></td>
	<td width="130" align="right"><b>貸款金額</b></td>
	<td width="130" align="right"><b>巳批期數</b></td>
	<td width="130" align="right"><b>每月本金</b></td>
	<td width="130" align="right"><b>本金結餘</b></td>
	<td width="130" align="right"><b>貸款情況</b></td>

	</tr>
	<tr><td colspan=9><hr></td></tr>

<%
do while not rs.eof
   if rs("lndate") >= dstart and rs("lndate") <=dend then
   xlndate=right("0"& day(rs("lndate")),2)&"/"&right("0"&month(rs("lndate")),2)&"/"&year(rs("lndate"))
%>
	<tr>
		<td><%=rs("memNo")%></td>
		<td><%=rs1("memName")%><%=rs1("memcname")%></td>
		<td align="right"><%=rs("lnnum")%></td>
		<td align="right"><%=xlndate%></td>
		<td align="right"><%=formatnumber(rs("appamt"),2)%></td>
		<td align="right"><%=formatNumber(rs("install"),2)%></td>
		<td align="right"><%=formatnumber(rs("monthrepay"),2)%></td>
		<td align="right"><%=formatNumber(rs("bal"),2)%></td>
               if rs("repaystat")="C" then
                 <td width="20">完</td>
               else
                 <td width="20">未完</td>
               end if   
	</tr>

<%
      
        end if
	rs.movenext
loop
%>
	<tr><td colspan=4><hr></td></tr>

	
	<tr>
		
		<td>Total Count :</td>
		<td align="right"><%=formatNumber(ttlcnt,2)%></td>
		<td colspan="2"></td>
		
	</tr>

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
