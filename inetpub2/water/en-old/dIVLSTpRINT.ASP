<!-- #include file="../conn.asp" -->

<%
server.scripttimeout = 1800


todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())

choice = request.form("kind")



    kind = request.form("KIND")

     select case request.form("KIND")
            case "S"
		stylefield =stylefield&" and a.bank = 'S' and a.deleted = 0 order by a.memno"
            case "B"

                 stylefield =stylefield&" and a.bank ='B'  and a.deleted = 0 order by a.memno"
            case "C"
                 stylefield =stylefield&" and a.bank ='N' and a.deleted = 0 order by a.memno"
            case  "N"
                 stylefield =stylefield&" and a.deleted = 1 order by a.memno"
            case "all"

                  stylefield =stylefield&"  order by a.memno"
     end select


SQl = "select a.*,b.memname,b.memcname,b.mstatus  from dividend a,memmaster b where a.memno=b.memno "&stylefield
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn



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
	objFile.Write "貸款帳細明表"
	objFile.WriteLine ""	
	objFile.Write "日期"&":"&right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
	objFile.WriteLine ""	
	objFile.WriteLine ""	
	objFile.Write " 社員編號 "
	objFile.Write " 社員名稱 "

	objFile.Write "    類別 "
	objFile.Write "       金額    "


	objFile.WriteLine ""
	for idx = 1 to 50
		objFile.Write "-"
	next
	objFile.WriteLine ""
        ttlcnt = 0
     	ttlapamt = 0
        ttlbal = 0
	do while not rs.eof

   
                select case rs("bank")
                       case "N"
                           status ="現金/支票"  
                       case "B"
  	 	 	   status ="銀行"  
                       case "S" 
                           status ="股金"
                       
                      
               end select
               if rs("deleted") = true then
                  status = "不派息"
               end if
               ttlname  = rs("memcname")
               ttlamt = ttlamt + rs("dividend")
               ttlcnt = ttlcnt + 1

  		objFile.Write left("   "&rs("memNo")&spaces,10) 
		objFile.Write left(ttlname&spaces,10)
		
                objFile.Write left(" "&status&spaces,8)
		objFile.Write right(spaces&formatnumber(rs("dividend"),2),10)
 

		objFile.WriteLine    

               

           

             
	rs.movenext
	loop
	for idx = 1 to 50
		objFile.Write "-"
	next
	objFile.WriteLine ""
        objFile.Write left(spaces,5)
        objFile.Write "總數 : "
        objFile.Write right(spaces&formatnumber(ttlcnt,0),8) 
        objFile.Write "  總金額 ：" 
        objFile.Write right(spaces&formatnumber(ttlamt,2),12) 

	objFile.WriteLine ""


	objFile.Close


	conn.close
	set conn=nothing
	response.redirect "../txt/"&session("username")&".txt"
end if
%>
<html>
<head>
<title>股息分配列表</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center">
	<td colspan="15"><font size="4">水務署員工儲蓄互助社</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
	<td colspan="15"><font size="4">股息分配細明表</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
	<td colspan="15"><font size="4">日期 : <%=todate%></font></td>
        </tr>
</center>
</table>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		
	<td><font size="2" color="#FFFFFF">社員編號</font></td>
	<td><font size="2" color="#FFFFFF">社員名稱</font></td>
	
	<td><font size="2" color="#FFFFFF">類別</font></td>
	<td><font size="2" color="#FFFFFF">金額</font></td>


	
	</tr>
	
<%
        ttlcnt = 0
     	ttlapamt = 0
        ttlbal = 0      
     
	do while not rs.eof
               
    
                select case rs("bank")
                       case "N"
                           status ="現金/支票"  
                       case "B"
  	 	 	   status ="銀行"  
                       case "S" 
                           status ="股金"
                       
                      
               end select
               if rs("deleted") = true then
                  status = "不派息"
               end if
               ttlname  = rs("memcname")
               ttlamt = ttlamt + rs("dividend")
                 ttlcnt = ttlcnt + 1
         
%>
   <tr bgcolor="#FFFFFF">
	
  	<td><font size="2"><%=rs("memno")%></font></td>
	<td><font size="2"><%=ttlname%></font></td>
        <td><font size="2"><%=status%></font></td>
        <td align="right"><font size="2"><%=formatnumber(rs("dividend") ,2)%></font></td>

   </tr> 
<%	

rs.movenext
loop
%>
	<tr>
		<td></td>
		<td>總數:<%=ttlcnt%></td>              		                 
		 <td>總金額 ：</td>
             
		<td width=100 align="right"><font size="2"><b><%=formatNumber(ttlamt,2)%></b></font></td>

				
              


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
