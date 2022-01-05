<!-- #include file="../conn.asp" -->

<%
mMonth = request("mMonth")
 mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())


if IsNumeric(mMonth) then
	if int(mMonth)<1 or int(mMonth)>12 then
		response.redirect "birthdayList.asp"
	end if
else
	response.redirect "birthdayList.asp"
end if

SQl =   "select m.memno, m.memname, m.memcname, CONVERT(varchar, m.membday, 107) as 'membday', " & _
        "	ISNULL(m.memaddr1, '') AS 'memaddr1', ISNULL(m.memaddr2, '') AS 'memaddr2', " & _
        "	ISNULL(m.memaddr3, '') AS 'memaddr3', " & _
        "	m.mstatus,m.memMobile,m.memEmail  ,m.accode " & _
        "from memmaster m,  (SELECT DISTINCT memNo FROM Share WHERE amount > 0) s " & _
        "WHERE m.memNo = s.memno " & _
        "AND m.mstatus NOT IN ('D', 'V','C','B','P','8','9') " & _
        "AND wdate is null " & _
        "order by m.memno "
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn

if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"

end if
%>
<html>
<head>
<title>Birthday List</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<font size="4"  face="標楷體" >
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="4"  face="標楷體" >水務署員工儲蓄互助社<br>
<%=monthname(mMonth)%>份 社員生日名單列表<br><font size="2"  face="標楷體" >日期 : <%=mndate%></font></font></td></tr>
        

</table>
<br>
<br>

              
<table border="0" cellpadding="0" cellspacing="0"  width="1150">

	<tr height="15" valign="bottom">
		<td width=70 align="center"><font size="3"  face="標楷體" >社員編號</font></td>
                <td width=150 align="center"><font size="3"  face="標楷體" >英文名稱</font></td> 
		
		<td width=70 align="center"><font size="3"  face="標楷體" >中文姓名</font></td>
                <td width=70 align="center"><font size="3"  face="標楷體" >出生日期</font></td>
		<td width=100 align="center"><font size="3"  face="標楷體" >狀況</font></td>
                <td width=230 align="center"><font size="3"  face="標楷體" >聯絡電話</font></center></td>
		<td width=230 align="center"><font size="3"  face="標楷體" >電郵</font></center></td>						
                <td width=230 align="center"><font size="3"  face="標楷體" >備註</font></center></td>
	</tr>
	<tr><td colspan=10><hr></td></tr>
<%
do while not rs.eof
                 select case rs("mstatus") 
                       	case  "L"
                              xstatus =  "呆帳"
                             
                        case  "D" 
                             xstatus =  "冷戶"  
		
                        case "V"
                              xstatus = "IVA"
                        
                        case "C"
                              xstatus = "退社" 
		
                        case  "B"
                              xstatus = "破產" 
			 
                        case  "P"
                              xstatus = "去世"
			       
                        case  "N" 
                              xstatus = "正常"
			    
                        case  "J" 
                              xstatus = "新戶"
                          
                        case  "T" 
                              xstatus = "庫房" 
                           
                        case  "H" 
                               xstatus = "暫停銀行"
			
                        case  "A"
                               xstatus =  "自動轉帳(ALL)"
			
                        case "0"
                              xstatus = "自動轉帳(股金)"
			 
                        case "1"
                              xstatus = "自動轉帳(股金,利息)"
			
                        case "2"
                              xstatus = "自動轉帳(股金,本金)"                         
			 
                        case "3"
                             xstatus = "自動轉帳(利息,本金)"                         
			 
                        case "M"
                             xstatus = "庫房,銀行"
			   
                        case "F"
                              xstatus = "特別個案"  
			  
                        case "8"
                             xstatus = "終止社籍轉帳"
                          
                        case "9"
                             xstatus = "終止社籍正常"
                          
                     
               end select  
                   idx = ""
                  if rs("accode")="9999" then idx="退休" end if
                            
%>
	<tr>
		<td width=70 align="center"><%=rs("memNo")%></td>
                <td width=150 align="center"><%=rs("memName")%></td>
		
		<td width=70 align="center"><font size="3"  face="標楷體" ><%=rs("memcname")%></font></td>		
		<td width=70 align="center"><%=right("0"&day(rs("memBday")),2)&"/"&right("0"&month(rs("memBday")),2)&"/"&year(rs("memBday"))%></td>
		<td width=100 align="center"><font size="2"  face="標楷體" ><%=xstatus%></font></td>
             
                <td width=230 align="left"><font size="2"  face="標楷體" ><%=rs("memMobile")%></font></td>
                <td width=230 align="left"><font size="2"  face="標楷體" ><%=rs("memEmail")%></font></td>		
                <td width=230 align="left"><font size="2"  face="標楷體" ><%=idx%></font></td>

	</tr>
<%
	rs.movenext
loop
%>
</table>
</font>
</center>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
