<!-- #include file="../conn.asp" -->

<%
   server.scripttimeout = 1800

   yy = year(date())
   mm = month(date())
   dd = day(date())
   xyy = yy - 2 
   chkdate = xyy&"."&right("0"&mm,2)&"."&right("0"&dd,2)    

SQl = "Select memname,memcname,mstatus from memmaster  where  wdate is null and  mstatus='H'    order by memno "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 1,1

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
	objFile.Write "水務署員工儲蓄互助社"
	objFile.WriteLine ""
	objFile.Write "銀行暫停列表"
	objFile.WriteLine ""
 	objFile.Write "已過日期 : 兩年 "
	objFile.WriteLine ""
        objFile.WriteLine ""
	objFile.Write left("社員編號"&spaces,10)
	objFile.Write left("社員名稱"&spaces,40)
	objFile.Write right(spaces&"最後來往帳日期",10)
	objFile.Write right(spaces&"股金結餘",16)
	objFile.WriteLine ""
	for idx = 1 to 95
		objFile.Write "-"
	next
	objFile.WriteLine ""

        do while not rs.eof
           

              savettl = 0               
              sql1 = "Select memno,code,amount from share where memno='"&rs("memno")&"' order by memno,sdate,code "
              Set rs1 = Server.CreateObject("ADODB.Recordset")
              rs1.open sql1, conn, 2,2
              do while not rs1.eof   
       select case rs1("code")
              case "0A" ,"A1","A2","A3","C0","C1","C3"
                   savettl = savettl + rs1(2)
              case "B0","B1","G0","G1","G3","H0","H1","H3"
                  savettl = savettl - rs1(2)
       end select
                RS1.MOVENEXT
              LOOP
              RS1.CLOSE
           
             
		objFile.Write left(rs("memNo")&spaces,10)
		objFile.Write left("    "&rs("memName")&rs("memcname")&spaces,50)   
		
                objFile.Write right(spaces&formatnumber(savettl,2),18)
		objFile.WriteLine ""
            
          
       
	   rs.movenext
	loop
        
	for idx = 1 to 95
		objFile.Write "-"
	next

	objFile.Close
	rs.close
	set rs=nothing
        set rs1=nothing
	conn.close
	set conn=nothing
	response.redirect "../txt/"&session("username")&".txt"
end if
%>
<html>
<head>
<title>Delinquent Report</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<LINK href="../main.css" rel=STYLESHEET type=text/css>
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center">
		<td colspan="15"><font size="5">水務署員工儲蓄互助社</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="5">銀行暫停列表</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
		<td colspan="15"><font size="3">已過日期 : 兩年</font></td>
	</tr>
	<tr height="15" valign="bottom">
		<td width="80"><b>社員名稱</b></td>
		<td width="200"><b>社員名稱</b></td>		
		<td width="130" align="right"><b>最後來往帳日期</b></td>
		<td width="130" align="right"><b>股金結餘</b></td>
	</tr>
	<tr><td colspan=5><hr></td></tr>
<%


        do while not rs.eof
      
          
              savettl = 0               
              sql1 = "Select memno,code,amount from share where memno='"&rs("memno")&"' order by memno,sdate,code "
              Set rs1 = Server.CreateObject("ADODB.Recordset")
              rs1.open sql1, conn, 2,2
              do while not rs1.eof   
       select case rs1("code")
              case "0A" ,"A1","A2","A3","C0","C1","C3"
                   savettl = savettl + rs1(2)
              case "B0","B1","G0","G1","G3","H0","H1","H3"
                  savettl = savettl - rs1(2)
       end select
                RS1.MOVENEXT
              LOOP
              RS1.CLOSE
           
%>               
	<tr>
		<td><%=rs("memNo")%></td>
		<td><%=rs("memName")%><%=rs("memcname")%></td
		<td align="right"><%=formatNumber(savettl,2)%></td>
	</tr>
<%

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
