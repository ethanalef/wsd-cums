<!-- #include file="../conn.asp" -->

<%
server.scripttimeout = 1800

mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
stdate1 = request.form("sdate1")
stdate2 = request.form("sdate2")
yy = right(stdate1,4)
mm = cint(mid(stdate1,4,2))
dd = cint(left(stdate1,2))
todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
stdate1 = yy&"/"&mm&"/"&dd

yy = right(stdate2,4)
mm = cint(mid(stdate2,4,2))
dd = cint(left(stdate2,2))

stdate2 = yy&"/"&mm&"/"&dd




     select case request.form("KIND")
            case "cash"
		stylefield =" and (a.code='E3' or a.code='F3') order by a.memno,a.ldate,a.code "
            case "bank"
                 stylefield =" and (a.code='E1' or a.code='F1') order by a.memno,a.ldate,a.code "
            case "Trea"
                stylefield =" and (a.code='E2' or a.code='F2') order by a.memno,a.ldate,a.code "
            case "Share"
                stylefield =" and (a.code='B0' )  order by a.memno,a.ldate,a.code "
            case "unpaid"
		 stylefield =" and (a.code='E0' or a.code='F0') and a.amount<0 order by a.memno,a.ldate,a.code "
            case "adjust"
		 stylefield =" and (a.code='E7' or a.code='F7') and a.amount<>0 order by a.memno,a.ldate,a.code "

     end select

SQl = "select a.memno,a.lnnum,a.code,a.ldate,a.amount,b.memname,b.memcname  from loan a,memmaster b where  a.memno=b.memno and (ldate>='"&stdate1&"' and ldate<='"&stdate2&"' and left(code,1)<>'D' and left(code,1)<>'0'  ) "&stylefield
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn,1,1

ttlamt = 0
ttlsamt = 0
ttlpamt = 0
ttlpint = 0
ttlisamt = 0
ttlipamt = 0
ttlipint = 0

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
	objFile.Write "���ȸp���u�x�W���U��"
	objFile.WriteLine ""
	objFile.Write "�U�ڲӶ��ө���"
	objFile.WriteLine ""	
	objFile.Write "���"&":"&right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
	objFile.WriteLine ""	
	objFile.WriteLine ""	
	objFile.Write "�����s��   "
	objFile.Write "         �����W��         "
	objFile.Write " �U�ڽs��    "
	objFile.Write "������   "
        objFile.Write "���O   "
	objFile.Write "     (�Q��)   "
	objFile.Write "    (����)   "
	objFile.Write "   (�Ѫ�)   "
	objFile.Write "   (�`���B)  "
	objFile.WriteLine ""
	for idx = 1 to 120
		objFile.Write "-"
	next
	objFile.WriteLine ""
       xmemno =rs("memno") 
        memcname = rs("memcname")
        memname = rs("memname")
        lnnum = rs("lnnum")
        xcode = ""
        saveamt = 0
	do while not rs.eof
         
 
         select case rs("code")
          case "E1"
               lcode = "�Ȧ����"
          case "E2"
		 lcode ="�w�����"
          case "E3"
		 lcode ="�{���ٴ�"
          case "ET"
		 lcode ="�Ѫ��ٴ�"
          case "ER"
		 lcode ="�h�٥���"
         case "fR"
		 lcode ="�h�٧Q��"
          CASE "EI"
               lcode ="���" 

            
    end select 
  if rs("memno") <> xmemno then
 if xcode <>"" then 
sqlstr = "select * from share where memno='"&xmemno&"' and ldate= '"&ydate&"' and code='"&xcode&"'  "
Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.open sqlstr, conn,2,2
if not rs1.eof then
   saveamt = rs1("amount")
else
   saveamt = 0.00
end if
xcode=""

end if
              samt = pint1 + pamt1+saveamt   
                ttlpamt = ttlpamt+pamt1
                ttlpint = ttlpint+pint1
                ttlTemp = ttlTemp +samt   
                ttlsamt = ttlsamt + saveamt 
                if ttlTemp <>"" then 
  		objFile.Write left(" "&xmemNo&spaces,10) 
		objFile.Write left(memname&" "&memcname&spaces,22)
		objFile.Write left(" "&lnnum&spaces,12) 
                objFile.Write left(" "&xdate&spaces,12)
		objFile.Write left(lcode&spaces,6)
		objFile.Write right(spaces&formatnumber(pint1,2),13)
		objFile.Write right(spaces&formatnumber(pamt1,2),13)
		objFile.Write right(spaces&formatnumber(saveamt,2),13)
		objFile.Write right(spaces&formatnumber(samt,2),13)
		objFile.WriteLine    

                
        xmemno =rs("memno") 
        memcname = rs("memcname")
        memname = rs("memname")
        lnnum = rs("lnnum")
                sttlamt = 0
                pamt1 = 0
                pint1 = 0
                samt = 0
                pint2 = 0
                pamt2 = 0
              end if                                  
           
           end if  
           ydate = rs("ldate")
              select case rs("code")
 
                     case "E1"
                           xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
                         
                           skind="�Ȧ�"
                           xx =  1
                           pamt1 = rs("amount")
                           xcode = "A1"                     
                     case "F1"
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
                           xx = 2  
			   skind="�Ȧ�" 
                           pint1 = rs("amount")
                           ttlpint1 = ttlpint1 + pint1
			   xcode = "A1"    
                    case   "E2"
                           xx = 3
                           pamt1 = rs("amount")
                         
			    xcode = "A2"  
                     case "F2"
			   xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
                           xx = 4 
			   skind="�w��" 	
                           pint1 = rs("amount")
                           xcode ="A2"
                    case   "E3"
			   xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
                           xx = 2
                           pamt1 = rs("amount")
                            xcode = "A3"  
                     case "F3"
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))	
			   xx = 2	
                           skind="�{��" 
                           pint1 = rs("amount")
                           xcode="A3"
                    case "A3"
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))	
			   xx = 2	
                           skind="�{��" 
                           saveamt = rs("amount")
                   case   "E7"
			   xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
                           xx = 2
                           pamt1 = rs("amount")
                            xcode = "A7"  
                     case "F7"
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))	
			   xx = 2	
                           skind="�վ�" 
                           pint1 = rs("amount")
                           xcode="A7"
                    case "A7"
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))	
			   xx = 2	
                           skind="�վ�" 
                           saveamt = rs("amount")                       
                     case "EI"
                           xx = 1 
			   skind="���" 	
                           pamt1 = rs("amount")
                            xcode = "AI"  
                     case "FI"
                           xx = 2
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
			   skind="���" 	
                           pint1 = rs("amount")	
                    case "ET","B0"
                            xx = 2
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
			   skind="�Ѫ��٥�" 	
                           pamt1 = rs("amount")	                               
                     case "FT" ,"F0"
                            xx = 2
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
			   skind="�Ѫ��ٮ�" 	
                           pint1 = rs("amount")	  
                     case "ER"
                            xx = 2
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
			   skind="�h�٥���" 	
                           pamt1 = rs("amount")	   
                     case "FR"
                            xx = 2
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
			   skind="�h�٧Q��" 	
                           pint1 = rs("amount")  
                     case "D1"
                            xx = 2
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
			   skind="" 	
                           pamt1 = rs("amount")                                    
               end select 
             
	rs.movenext
	loop
       if xcode <>"" then 
          sqlstr = "select * from share where memno='"&xmemno&"' and ldate= '"&ydate&"' and code='"&xcode&"'  "
          Set rs1 = Server.CreateObject("ADODB.Recordset")
          rs1.open sqlstr, conn,2,2
          if not rs1.eof then
             saveamt = rs1("amount")
          else
             saveamt = 0.00
          end if
          xcode=""

         end if
              samt = pint1 + pamt1+saveamt   
                ttlpamt = ttlpamt+pamt1
                ttlpint = ttlpint+pint1
                ttlTemp = ttlTemp +samt   
                ttlsamt = ttlsamt + saveamt 
                if ttlTemp <> "" then 
  		objFile.Write left(" "&xmemNo&spaces,10) 
		objFile.Write left(memname&" "&memcname&spaces,22)
		objFile.Write left(" "&lnnum&spaces,12) 
                objFile.Write left(" "&xdate&spaces,12)
		objFile.Write left(lcode&spaces,6)
		objFile.Write right(spaces&formatnumber(pint1,2),13)
		objFile.Write right(spaces&formatnumber(pamt1,2),13)
		objFile.Write right(spaces&formatnumber(saveamt,2),13)
		objFile.Write right(spaces&formatnumber(samt,2),13)
		objFile.WriteLine    
              end if
	for idx = 1 to 120
		objFile.Write "-"
	next
	objFile.WriteLine ""
	objFile.Write space(57)
    	objFile.Write "  �`���B �G "
	objFile.Write right(spaces&formatnumber(ttlpint,2),13)
	objFile.Write right(spaces&formatnumber(ttlpamt,2),13)
	objFile.Write right(spaces&formatnumber(ttlsamt,2),13)
	objFile.Write right(spaces&formatnumber(ttlTemp,2),13)
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
<title>�U�ڲӶ��C��</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="4"  face="�з���" >���ȸp���u�x�W���U��<br>�U�ڲӶ��ө���<br><font size="2"  face="�з���" >��� : <%=mndate%></font></font></td></tr>
        <tr height="30" ><td colspan=9></td></tr>

        
	<tr height="15" valign="bottom">
        <font size="2"  face="�з���" >
	<td width="80">�����s��</td>
	<td width="80">�����W��</td>
	<td width="90" align="center">�U�ڽs��</td>
	<td width="60" align="center"> ��� </td>
	<td width="60" align="center">���O</td>
	<td width="80" align="center"> �Q��</td>
	<td width="80" align="center">����</td>
	<td width="80" align="center">�Ѫ�</td>
	<td width="80" align="center">�`���B</td>
	</tr>
	<tr><td colspan=9><hr></td></tr>


		

	
<%
   if not rs.eof then
        xmemno =rs("memno") 
        memcname = rs("memcname")
        memname = rs("memname")
        lnnum = rs("lnnum")
        xcode  =""       
         saeamt = 0 
	do while not rs.eof
        
      
        select case rs("code")
          case "E1"
               lcode = "�Ȧ����"
          case "E2"
		 lcode ="�w�����"
          case "E3"
		 lcode ="�{���ٴ�"
          case "E7"
		 lcode ="�վ�"
          case "ET"
		 lcode ="�Ѫ��ٴ�"
          case "ER"
		 lcode ="�h�٥���"
         case "fR"
		 lcode ="�h�٧Q��"
          CASE "EI"
               lcode ="���" 

    end select 
 if rs("memno") <> xmemno then
if xcode <>"" then 
sqlstr = "select * from share where memno='"&xmemno&"' and ldate= '"&ydate&"' and code='"&xcode&"'  "
Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.open sqlstr, conn,2,2
if not rs1.eof then
   saveamt = rs1("amount")
else
   saveamt = 0.00
end if
xcode=""
rs1.close
end if
              samt = pint1 + pamt1+saveamt
	      ttlTemp = ttlTemp +samt
              ttlpamt1 = ttlpamt1+ pamt1
              ttlpint1 = ttlpint1 + pint1	
              ttlsamt = ttlsamt + saveamt
            if samt <> "" then 
%>
   <tr bgcolor="#FFFFFF">
	
  	<td width="80"><%=xmemno%></td>
	<td width="80"><%=memcname%></td>
	<td width="90"><%=rs("lnnum")%></td>	
	<td width="80"><%=xdate%></td>
	<td width="80" align="center"><%=lcode%></td>
	<td width="80" align="right"><%=formatnumber(pint1,2)%></td>
	<td width="100" align="right"><%=formatnumber(pamt1,2)%></td>
	<td width="80" align="right"><%=formatnumber(saveamt,2)%></td>
	<td width="100" align="right"><%=formatnumber(samt,2)%></td>
   </tr> 
<%	
                xmemno = rs("memno")
                memcname = rs("memcname")
                memname = rs("memname")
                lnnum = rs("lnnum")
                samt = 0
                pamt1 = 0
                pint1 = 0
              end if                                                                          
           end if  
              ydate = rs("ldate")
              select case rs("code")
                     case "E1"
                           xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
                         
                           skind="�Ȧ�"
                           xx =  1
                           pamt1 = rs("amount")
                           xcode = "A1"                     
                     case "F1"
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
                           xx = 2  
			   skind="�Ȧ�" 
                           pint1 = rs("amount")
                           ttlpint1 = ttlpint1 + pint1
			   xcode = "A1"    
                    case   "E2"
                           xx = 3
                           pamt1 = rs("amount")
                         
			    xcode = "A2"  
                     case "F2"
			   xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
                           xx = 4 
			   skind="�w��" 	
                           pint1 = rs("amount")
                        
                    case   "E3"
			   xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
                           xx = 2
                           pamt1 = rs("amount")
                            xcode = "A3"  
                     case "F3"
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))	
			   xx = 2	
                           skind="�{��" 
                           pint1 = rs("amount")
                    case "A3"
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))	
			   xx = 2	
                           skind="�{��" 
                           saveamt = rs("amount")
                       
                   case   "E7"
			   xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
                           xx = 2
                           pamt1 = rs("amount")
                            xcode = "A7"  
                     case "F7"
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))	
			   xx = 2	
                           skind="�վ�" 
                           pint1 = rs("amount")
                           xcode="A7"
                    case "A7"
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))	
			   xx = 2	
                           skind="�վ�" 
                           saveamt = rs("amount")
                                            
                     case "EI"
                           xx = 1 
			   skind="���" 	
                           pamt1 = rs("amount")
                            xcode = "AI"  
                     case "FI"
                           xx = 2
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
			   skind="���" 	
                           pint1 = rs("amount")	
                    case "ET","B0"
                            xx = 2
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
			   skind="�Ѫ��٥�" 	
                           pamt1 = rs("amount")	                               
                     case "FT""F0"
                            xx = 2
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
			   skind="�Ѫ��ٮ�" 	
                           pint1 = rs("amount")	  
                     case "ER"
                            xx = 2
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
			   skind="�h�٥���" 	
                           pamt1 = rs("amount")	   
                     case "FR"
                            xx = 2
			    xdate = right("0"&day(rs("ldate")),2)&"/"&right("0"&month(rs("ldate")),2)&"/"&year(rs("ldate"))
			   skind="�h�٧Q��" 	
                           pint1 = rs("amount")                
               end select 
rs.movenext
loop
end if
if xcode <>"" then 
sqlstr = "select * from share where memno='"&xmemno&"' and ldate= '"&ydate&"' and code='"&xcode&"'  "
Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.open sqlstr, conn,2,2
if not rs1.eof then
   saveamt = rs1("amount")
else
   saveamt = 0.00
end if
xcode=""
rs1.close
end if
              samt = pint1 + pamt1+saveamt
	      ttlTemp = ttlTemp +samt
              ttlpamt1 = ttlpamt1+ pamt1
              ttlpint1 = ttlpint1 + pint1	
              ttlsamt = ttlsamt + saveamt
            if samt <> "" then 
%>
   <tr bgcolor="#FFFFFF">
	
 	<td width="80"><%=xmemno%></td>
	<td width="80"><%=memcname%></td>
	<td width="90"><%=lnnum%></td>	
	<td width="80"><%=xdate%></td>
	<td width="80" align="center"><%=lcode%></td>
	<td width="80" align="right"><%=formatnumber(pint1,2)%></td>
	<td width="100" align="right"><%=formatnumber(pamt1,2)%></td>
	<td width="80" align="right"><%=formatnumber(saveamt,2)%></td>
	<td width="100" align="right"><%=formatnumber(samt,2)%></td>
   </tr> 
<%	
end if
%>
       	<tr><td colspan=9><hr></td></tr>
	<tr>
		<tr><td colspan=4></td>
		 <td>�`���B �G</td>
                
		<td width=80 align="right"><%=formatNumber(ttlpint1,2)%></td>
		<td width=100 align="right"><%=formatNumber(ttlpamt1,2)%></td>	
                <td width=80 align="right"><%=formatNumber(ttlsamt,2)%></td>		
		<td width=100 align="right"><%=formatNumber(ttlTemp,2)%></td>
	

	</tr>
        <tr><td colspan=5></td>
            <td width="80"  align="right">========</td>
         �@ <td width="100" align="right">=========</td>
            <td width="80"  align="right">========</td>
         �@ <td width="100" align="right">=========</td>
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
