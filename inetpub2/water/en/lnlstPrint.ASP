<!-- #include file="../conn.asp" -->

<%
server.scripttimeout = 1800

stdate1 = request.form("stdate1")
stdate2 = request.form("stdate2")
yy = right(stdate1,4)
mm = mid(stdate1,4,2)
dd = left(stdate1,2)

todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())

stdate1 = yy&"/"&mm&"/"&dd

yy = right(stdate2,4)
mm = mid(stdate2,4,2)
dd = left(stdate2,2)

stdate2 = yy&"/"&mm&"/"&dd



mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
stylefield="and (lndate>='"&stdate1&"' and lndate<='"&stdate2&"' )  "

memno=request.form("memNo")
if memno <>"ALL" then
   pos = instr(memno,"-")
   xmemno= left(memno,pos-1)
   stylefield =stylefield&" and a.memno='"&xmemno&"' " 
end if
    kind = request.form("KIND")

     select case request.form("KIND")
            case "C"
		stylefield =stylefield&" and a.repaystat='C'  order by a.memno,a.lnnum"
            case "N"
                 stylefield =stylefield&" and a.repaystat ='N' order by a.memno,a.lnnum"
            case "all"

                  stylefield =stylefield&"  order by a.memno,a.lnnum"
     end select


SQl = "select a.*,b.memname,b.memcname,b.mstatus,b.monthssave,b.monthsave  from loanrec a,memmaster b where a.memno=b.memno "&stylefield
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
	objFile.Write "   貸款編號    "
	objFile.Write "   取票日期 "
	objFile.Write "　　　　尚欠　　"
	objFile.Write "　　新貸　　"
        objFile.Write "     貸款總額  "
	objFile.Write "  期數 "
	objFile.Write " 股金    "
         if kind ="N" then
	objFile.Write " 還款結餘 "
        else
	objFile.Write "清數日期    "
        end if

	objFile.WriteLine ""
	for idx = 1 to 115
		objFile.Write "-"
	next
	objFile.WriteLine ""
        ttlcnt = 0
     	ttlapamt = 0
        ttlbal = 0
	do while not rs.eof

                xdate = right("0"&day(rs("lndate")),2)&"/"&right("0"&month(rs("lndate")),2)&"/"&year(rs("lndate"))      
                ydate=""
                if rs("cleardate")<>"" then 
                ydate = right("0"&day(rs("cleardate")),2)&"/"&right("0"&month(rs("cleardate")),2)&"/"&year(rs("cleardate"))  
                end if   
                select case rs("repaystat")
                       case "C"
                           status ="巳清數"  
                       case "N"
  	 	 	   status ="未清數"  
               end select
               ttlname  = rs("memcname")
               ttlapamt = ttlapamt + round(rs("appamt"),0)
               if rs("lnflag") = "Y" then 
                  updamt = rs("appamt") - rs("chequeamt")
                  nwamt  = rs("chequeamt")
               else
                  updamt = 0
                  nwamt = rs("appamt")
               end if  
               ttlupdamt = ttlupdamt + updamt 
               ttlnamt = ttlnamt + rs("appamt")
               ttlcheqamt = ttlcheqamt + nwamt
               ttlbal   = ttlbal +   round(rs("bal"),0)

  		objFile.Write left("   "&rs("memNo")&spaces,10) 
		objFile.Write left(ttlname&spaces,10)
		objFile.Write left(rs("lnnum")&spaces,12) 
                objFile.Write left("  "&xdate&spaces,16)
		objFile.Write right(spaces&formatnumber(updamt,2),13)
                objFile.Write right(spaces&formatnumber(nwamt,2),13)
		objFile.Write right(spaces&formatnumber(rs("appamt"),2),15)
		objFile.Write right(spaces&formatnumber(rs("install"),0),6)
                paidamt = 0
                select case  rs("mstatus")
                       case "A","0","1","2"
 		            paidamt = rs("monthsave")
                       case "M","T"                            
		            
                            paidamt = rs("monthssave")
                end select 
                objFile.Write right(spaces&formatnumber(paidamt,0),6)
                if kind="N" then
		objFile.Write right(spaces&formatnumber(rs("bal"),2),16)
                else 
 		objFile.Write right(spaces&ydate,12)
                end if
		
		objFile.WriteLine    

                ttlcnt = ttlcnt + 1

           

             
	rs.movenext
	loop
	for idx = 1 to 115
		objFile.Write "-"
	next
	objFile.WriteLine ""
 
       objFile.Write space(51)
             objFile.Write right(spaces&formatnumber(ttlupdamt,2),13)     
       objFile.Write right(spaces&formatnumber(ttlcheqamt,2),13)      
        objFile.Write right(spaces&formatnumber(ttlapamt,2),15)      
     '   end if 
        objFile.Write space(22)
        if kind="N" then
        objFile.Write right(spaces&formatnumber(ttlbal,2),16)
        end if 
	objFile.WriteLine ""
	objFile.WriteLine ""
	objFile.Write "貸款數目總計"
	objFile.Write right(spaces&formatnumber(ttlcnt,0),10)
	objFile.WriteLine ""

	objFile.Close


	conn.close
	set conn=nothing
	response.redirect "../txt/"&session("username")&".txt"
end if
%>
<html>
<head>
<title>貸款帳列表</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="4"  face="標楷體" >水務署員工儲蓄互助社<br>貸款帳細明表<br><font size="2"  face="標楷體" >日期 : <%=mndate%></font></font></td></tr>
        <tr height="30" ><td colspan=9></td></tr>

        
	<tr height="15" valign="bottom">
        <font size="2"  face="標楷體" >
	<td width="70">社員編號</td>
	<td width="70">社員名稱</td>
	<td width="70" align="right">貸款編號</td>
	<td width="80" align="right">取票日期</td>
	<td width="60" align="right">尚欠</td>
	<td width="80" align="right">新貸</td>
	<td width="80" align="right">貸款總額</td>
	<td width="30" align="right">期數</td>
	<td width="80" align="right">股金</td>



<%if kind="N" then %>
	<<td width="80" align="right">還款結餘></td>
<%else%>
	<td width="80" align="right">清數日期</td>
<%end if%>
	<td width="50" align="right">現狀</td>
	
	</tr>
        <tr><td colspan=11><hr></td></tr>	
<%
        ttlcnt = 0
     	ttlapamt = 0
        ttlbal = 0      
     
	do while not rs.eof
                xdate = right("0"&day(rs("lndate")),2)&"/"&right("0"&month(rs("lndate")),2)&"/"&year(rs("lndate"))      
                ydate=""
                if rs("cleardate")<>"" then 
                ydate = right("0"&day(rs("cleardate")),2)&"/"&right("0"&month(rs("cleardate")),2)&"/"&year(rs("cleardate"))  
                end if   
                select case rs("repaystat")
                       case "C"
                           status ="巳清數"  
                       case "N"
  	 	 	   status ="未清數"  
               end select
               ttlname  = rs("memcname")
               ttlapamt = ttlapamt + round(rs("appamt"),0)
              if rs("lnflag") = "Y" then 
                  updamt = rs("appamt") - rs("chequeamt")
                  nwamt  = rs("chequeamt")
               else
                  updamt = 0
                  nwamt = rs("appamt")
               end if  
               ttlbal   = ttlbal +   round(rs("bal"),0)
               ttlnamt = ttlnamt + rs("appamt")
               ttlcheqamt = ttlcheqamt + nwamt
               ttlupd  = ttlupd + updamt
               ttlcnt = ttlcnt + 1
               paidamt = 0
               select case rs("mstatus")
                      case "0" ,"1","2","3","A"
                           paidamt = rs("monthsave")
                      case "M","T"
                           paidamt = rs("monthssave")
              end select    
%>
   <tr bgcolor="#FFFFFF">
	
  	<td width=70 align="center"><%=rs("memno")%></td>
	<td width=70 align="center"><%=ttlname%></td>
        <td width=80 align="right"><%=rs("lnnum")%></td>
	<td width=80 align="right"><%=xdate%></td>
	<td width=100 align="right"><%=formatnumber(updamt,2)%></td>
	<td width=100 align="right"><%=formatnumber(nwamt,2)%></td>
	<td width=100 align="right"><%=formatnumber(rs("appamt"),2)%></td>
	<td width=30 align="right"><%=formatnumber(rs("install"),0)%></td>
        <td width=100 align="right"><%=formatnumber(paidamt ,0)%></td>
<%if kind="N" then%>
	<td width=80 align="right"><%=formatnumber(rs("bal"),2)%></td>
<%else%>
	<td width=80 align="right"><%=ydate%></td>
<%end if%>
 	<td width=50 align="right"><%=status%></td>
   </tr> 
<%	

rs.movenext
loop
%>
        <tr><td colspan=11><hr></td></tr>	
	<tr>
		<td></td>
		<td>總數:<%=ttlcnt%></td>              		                 
		 <td>總金額 ：</td>
                <td></td>
		<td width=100 align="right"><%=formatNumber(ttlupd,2)%></td>
		<td width=100 align="right"><%=formatNumber(ttlcheqamt,2)%></td>
		<td width=100 align="right"><%=formatNumber(ttlapamt,2)%></td>
		<td></td>		
                <td></td>
<%if kind="N" then %> 
		<td width=100 align="right"><font size="2"><b><%=formatNumber(ttlbal,2)%></b></font></td>
<%end if %>
	</tr>
        <tr><td colspan=4></td>
            <td width=100 align="right">==========</td>   
            <td width=100 align="right">==========</td>   
            <td width=100 align="right">==========</td>   
            <td colspan=2></td>
        </tr>

</font>	
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
