<!-- #include file="../conn.asp" -->

<%
server.scripttimeout = 1800

 mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
xstatus = request.form("KIND")
banks   = request.form("banks")
bnklmt  = request.form("banklmt")
tpay    = request.form("tpay")

   stylefield = ""
if xstatus="all" then

   stylefield = ""
else


  select case xstatus 
         case "A"

              if banks = "Y" and bnklmt="Y" then

                 stylefield = " where (mstatus='A' ) and (monthsave=0 or monthsave is null ) and (bnklmt = 0 or bnklmt is null )  "
                 
              else
              if banks = "Y" and bnklmt<>"Y"  then
                 stylefield = " where (mstatus='A' ) and (monthsave=0 or monthsave is null )  "
              else

             if bnklmt="Y" and banks<>"Y"  then
                stylefield = " where (mstatus='A') and (bnklmt = 0 or bnklmt is null ) " 
             else
                stylefield = " where ( mstatus='A' ) "
            end if  
            end if
            end if
           case "0"
             if banks = "Y" and bnklmt="Y" then

                 stylefield = " where ( mstatus='0' ) and (monthsave=0 or monthsave is null ) and (bnklmt = 0 or bnklmt is null )  "
                 
              else
              if banks = "Y" and bnklmt<>"Y"  then
                 stylefield = " where ( mstatus='0' ) and (monthsave=0 or monthsave is null )  "
              else

             if bnklmt="Y" and banks<>"Y"  then
                stylefield = " where ( mstatus='0' ) and (bnklmt = 0 or bnklmt is null ) " 
            else
                stylefield = " where ( mstatus='0' ) "
            end if  
            end if
            end if
          case "1"
             if banks = "Y" and bnklmt="Y" then

                 stylefield = " where ( mstatus='1' ) and (monthsave=0 or monthsave is null ) and (bnklmt = 0 or bnklmt is null )  "
                 
              else
              if banks = "Y" and bnklmt<>"Y"  then
                 stylefield = " where ( mstatus='1' ) and (monthsave=0 or monthsave is null )  "
              else

             if bnklmt="Y" and banks<>"Y"  then
                stylefield = " where (  mstatus='1') and (bnklmt = 0 or bnklmt is null ) " 
             else
                stylefield = " where ( mstatus='1' ) "
            end if  
            end if
            end if
         case "T","M"
            if tpay = "Y" then
                 stylefield = " where (mstatus='T'  or mstatus='M') and (bnklmt = 0 or bnklmt is null ) " 
            else
                 stylefield = " where (mstatus='T'  or mstatus='M') "
            end if                  
        case else
            stylefield = "where  mstatus = '"&xstatus&"' "
end select
end if


SQl = "select *  from memmaster "&stylefield&" and wdate is null  order by memno"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn,1,1

ttlamt = 0
ttlsamt = 0
ttlpamt = 0
ttlpint = 0
ttlisamt = 0
ttlipamt = 0
ttlipint = 0
xdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
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
	objFile.Write "社員轉帳資料細明表"
	objFile.WriteLine ""	

	objFile.Write "日期"&":"&xdate
	objFile.WriteLine ""	
	objFile.WriteLine ""	
	objFile.Write "社員編號   "
	objFile.Write "         社員名稱         "
	objFile.Write "   銀行帳號 "
	objFile.Write "   每月儲蓄(銀行) "
        objFile.Write " 轉帳上限 "
        objFile.Write " 每月儲蓄(庫房) "
	objFile.Write "  庫房過數 "
	objFile.Write "   現況 "

	objFile.WriteLine ""
	for idx = 1 to 120
		objFile.Write "-"
	next
	objFile.WriteLine ""

	do while not rs.eof
 
                select case rs("mstatus") 
                       	case  "L"
                              idx =  "呆帳"
                             
                        case  "D" 
                             idx =  "冷戶"  
		
                        case "V"
                              idx = "IVA"
                        
                        case "C"
                              idx = "退社" 
		
                        case  "P"
                              idx = "去世" 
			 
                        case  "B"
                              idx = "破產"
			       
                        case  "N" 
                              idx = "正常"
			    
                        case  "J" 
                              idx = "新戶"
                          
                        case  "T" 
                              idx = "庫房" 
                           
                        case  "H" 
                               idx = "暫停銀行"
			
                        case  "A"
                               idx =  "自動轉帳(ALL)"
			
                        case "0"
                              idx = "自動轉帳(股金)"
			 
                        case "1"
                              idx = "自動轉帳(股金,利息)"
			
                        case "2"
                              idx = "自動轉帳(股金,本金)"                         
			 
                        case "3"
                             idx = "自動轉帳(利息,本金)"                         
			 
                        case "M"
                             idx = "庫房,銀行"
			   
                        case "F"
                              idx = "特別個案"  
			  
                        case "8"
                             idx = "終止社籍轉帳"
                          
                        case "9"
                             idx = "終止社籍正常"
                          
                     
               end select 
            Bank = rs("bnk")&"-"&rs("bch")&"-"&rs("bacct")
            if Bank = "--"  then
               Bank = ""
            end if
            monthsave = rs("monthsave")
            if monthsave <>""  then
                monthsave = cint(monthsave)
            else 
               monthsave=0                          
            end if  
            bnklmt = rs("bnklmt")
            if bnklmt <>""  then
               bnklmt = cint(bnklmt) 
            else  
               bnklmt=0                          
            end if  
            monthssave = rs("monthssave")
            if monthssave <>""  then
               monthssave = cint(monthssave)
            else 
               monthssave= 0                           
            end if 
            tpayamt = rs("tpayamt")
            if tpayamt <>""   then
               tpayamt = cint(tpayamt) 
            else
               tpayamt=0                           
            end if 

  		objFile.Write left(" "&rs("memNo")&spaces,10) 
		objFile.Write left(rs("memname")&" "&rs("memcname")&spaces,22)
		objFile.Write left(bank&spaces,15)
		objFile.Write right(spaces&formatnumber(monthsave,2),13)
		objFile.Write right(spaces&formatnumber(bnklmt,2),13)
		objFile.Write right(spaces&formatnumber(monthssave,2),13)
		objFile.Write right(spaces&formatnumber(tpayamt,2),13)
		objFile.Write right(spaces&idx,10)              
		objFile.WriteLine    

                
 
	rs.movenext
	loop

	for idx = 1 to 120
		objFile.Write "-"
	next
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
<title>貸款帳細項列表</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="4"  face="標楷體" >水務署員工儲蓄互助社<br>社員轉帳資料細明表<br><font size="2"  face="標楷體" >日期 : <%=mndate%></font></font></td></tr>
        

</table>
<br>
<br>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="20" valign="bottom">
	<td width=70 align="center"><font size="3"  face="標楷體" >社員編號</font></td>
	<td width=70 align="center"><font size="3"  face="標楷體" >社員名稱</font></td>
	<td width=130 align="center"><font size="3"  face="標楷體" >銀行帳號</font></td>
	<td width=100 align="center"><font size="3"  face="標楷體" >每月儲蓄<br>(銀行)</font></td>
	<td width=100 align="center"><font size="3"  face="標楷體" >轉帳上限</font></td>
	<td width=100 align="center"><font size="3"  face="標楷體" >每月儲蓄<br>(庫房)</font></td>
	<td width=100 align="center"><font size="3"  face="標楷體" >庫房過數</font></td>
	<td width=130 align="center"><font size="3"  face="標楷體" >現況</font></td>
	</tr>
	<tr><td colspan=8><hr></td></tr>
<%
   if not rs.eof then

    do while not rs.eof
        
                select case rs("mstatus") 
                       	case  "L"
                              idx =  "呆帳"
                             
                        case  "D" 
                             idx =  "冷戶"  
		
                        case "V"
                              idx = "IVA"
                        
                        case "C"
                              idx = "退社" 
		
                        case  "P"
                              idx = "去世" 
			 
                        case  "B"
                              idx = "破產"
			       
                        case  "N" 
                              idx = "正常"
			    
                        case  "J" 
                              idx = "新戶"
                          
                        case  "T" 
                              idx = "庫房" 
                           
                        case  "H" 
                               idx = "暫停銀行"
			
                        case  "A"
                               idx =  "自動轉帳(ALL)"
			
                        case "0"
                              idx = "自動轉帳(股金)"
			 
                        case "1"
                              idx = "自動轉帳(股金,利息)"
			
                        case "2"
                              idx = "自動轉帳(股金,本金)"                         
			 
                        case "3"
                             idx = "自動轉帳(利息,本金)"                         
			 
                        case "M"
                             idx = "庫房,銀行"
			   
                        case "F"
                              idx = "特別個案"  
			  
                        case "8"
                             idx = "終止社籍轉帳"
                          
                        case "9"
                             idx = "終止社籍正常"
                          
                     
               end select 
            Bank = rs("bnk")&"-"&rs("bch")&"-"&rs("bacct")
            if Bank = "--"  then
               Bank = ""
            end if
            monthsave = rs("monthsave")
            if monthsave ="" or monthsave=0  then
               monthsave=""
            end if  
            bnklmt = rs("bnklmt")
            if bnklmt ="" or bnklmt=0  then
               bnklmt=""
            end if  
            monthssave = rs("monthssave")
            if monthssave ="" or monthssave=0  then
               monthssave=""
            end if 
            tpayamt = rs("tpayamt")
            if tpayamt ="" or tpayamt=0  then
               tpayamt=""
            end if 
%>
   <tr bgcolor="#FFFFFF">
	
  	<td width=70 align="center"><%=rs("memno")%></td>
	<td width=70 align="center"><font size="3"  face="標楷體" ><%=rs("memcname")%></font></td>
	<td width=130 align="center"><%=Bank%></td>	
	<td width=100 align="center"><%=monthsave%></td>
	<td width=100 align="center"><%=bnklmt%></td>
	<td width=100 align="center"><%=monthssave%></td>
	<td width=100 align="center"><%=tpayamt%></td>
	<td width=130 align="center"><font size="3"  face="標楷體" ><%=idx%></font></td>
   </tr> 
<%	

   rs.movenext
   loop
end if
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

