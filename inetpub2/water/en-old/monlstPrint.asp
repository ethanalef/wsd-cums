<!-- #include file="../conn.asp" -->

<%
server.scripttimeout = 1800

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

choice = reauest.form("choice")
sql = ""

  conn.begintrans
  conn.execute( "create table #temp ( memno int ,lnnum char(10), ldate smalldatetime, code char(2) , amount money ) ")
  conn.execute( "insert into #temp (memno,ldate,code,amount ) select memno,ldate,code,amount from share where (right(code,1)='2' ) and  ldate>='"&stdate1&"' and ldate <='"&stdate2&"' "&sql )
  conn.execute( "insert into #temp (memno,lnnum,ldate,code,amount ) select memno,lnnum,ldate,code,amount from loan  where (right(code,1)='2' ) and ldate>='"&stdate1&"' and ldate <='"&stdate2&"' "&sql )
  conn.committrans

sql = "select a.*,b.memname,bmemcname,memHKID  from #temp a,memmaster b where a,memno=b.memno  order by memno,ldate,code"  
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
	objFile.Write "�I�бb�Ӷ��ө���"
	objFile.WriteLine ""	
	objFile.Write "���"&":"&right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
	objFile.WriteLine ""	
	objFile.WriteLine ""
        objFile.Write " �����Ҹ��X  "	
	objFile.Write "�����s��   "
	objFile.Write "         �����W��            "
	objFile.Write "�U�ڽs��   "
	objFile.Write "������   "
	objFile.Write "     (�Ѫ�)   "
	objFile.Write "     (�Q��)   "
	objFile.Write "   (����)   "
	objFile.Write "  (�`���B)  "
	objFile.WriteLine ""
	for idx = 1 to 122
		objFile.Write "-"
	next
	objFile.WriteLine ""
        memno =rs("memno") 
        memname = rs("memname")
        memcname=rs("memcname")   
        memhkid = rs("memhkid")     
        xcode = rs("ldate")
        saveamt = 0
	do while not rs.eof
           if rs("memno") <> memno and rs("ldate")<> xdate  then

              sttlamt = samt  + pint + pamt
              gttlamt = gttlamt + sttlamt

              objFile.Write left(" "&memhkid&spaces,10)
              objFile.Write left(memno&spaces,6)
              objFile.Write left(memname&" "&memcname&spaces,25)
              objFile.Write left(lnnum&spaces,12)
              objFile.Write right(spaces&formatnumber(samt,2),13)
              objFile.Write right(spaces&formatnumber(pint,2),13)
              objFile.Write right(spaces&formatnmber(pamt,2),13)
              objFile.Write right(spaces&formatnumber(sttlamt,2),13) 
              objFile.WriteLine ""             
              samt = 0          
              pint = 0 
              pamt =  0          
              lnnum = ""    
              memno =rs("memno") 
              memname = rs("memname")
              memcname=rs("memcname")        
              memhkid = rs("memhkid")
              xcode = rs("ldate")
           end if
           select case rs("code")
                  case "A2"
                       samt = rs("amount") 
                       ttlsamt = ttlsamt + samt
                  case "E2"
                       lnnum = rs("lnnum")
                       pamt =rs("amount")
                       ttlpamt = ttlpamt + pamt
                  case "F2"
                        pint = rs("amount")
                        ttlpint  = ttlpint + pint
          end select             
       	rs.movenext
	loop
	for idx = 1 to 122
		objFile.Write "-"
	next
	objFile.WriteLine ""
	objFile.Write space(57)
    	objFile.Write "  �`���B �G "
	objFile.Write right(spaces&formatnumber(ttlpint,2),13)
	objFile.Write right(spaces&formatnumber(ttlpamt,2),13)
	objFile.Write right(spaces&formatnumber(ttlsamt,2),13)
	objFile.Write right(spaces&formatnumber(gttlamt,2),13)
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
<title>�U�ڱb�Ӷ��C��</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height="30" valign="top" align="center">
	<td colspan="15"><font size="4">���ȸp���u�x�W���U��</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
	<td colspan="15"><font size="4">�U�ڱb�Ӷ��ө���</font></td>
	</tr>
	<tr height="35" valign="top" align="center">
	<td colspan="15"><font size="4">��� : <%=todate%></font></td>
        </tr>
</center>
</table>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="336699">
	<tr bgcolor="#330000" align="center">
		
	<td><font size="2" color="#FFFFFF">�����s��</font></td>
	<td><font size="2" color="#FFFFFF">�����W��</font></td>
	<td><font size="2" color="#FFFFFF">�U�ڽs��</font></td>
	<td><font size="2" color="#FFFFFF"> ��� </font></td>
	<td><font size="2" color="#FFFFFF"> ���O </font></td>
	<td><font size="2" color="#FFFFFF"> �Q�� </font></td>
	<td><font size="2" color="#FFFFFF"> ���� </font></td>
	<td><font size="2" color="#FFFFFF"> �Ѫ� </font></td>
	<td><font size="2" color="#FFFFFF"> �`���B </font></td>
	</tr>
	
<%
   if not rs.eof then
        xmemno =rs("memno") 
        xcode  =""       
         saeamt = 0 
	do while not rs.eof
        
           if rs("memno") <> xmemno then
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
          CASE "D1"
               lcode ="�s�U"  
          CASE "D0"
                lcode ="�U�ڲM��"
            
    end select 
if xcode <>"" then 
sqlstr = "select * from share where memno='"&xmemno&"' and sdate= '"&ydate&"' and code='"&xcode&"'  "
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
            if samt > 0 then 
%>
   <tr bgcolor="#FFFFFF">
	
  	<td><font size="2"><%=xmemno%></font></td>
	<td><font size="2"><%=rs("memname")%><%=rs("memcname")%></font></td>
	<td><font size="2"><%=rs("lnnum")%></font></td>	
	<td ><font size="2"><%=xdate%></font></td>
	<td align="right"><font size="2" ><center><%=lcode%></center></font></td>
	<td align="right"><font size="2"><%=formatnumber(pint1,2)%></font></td>
	<td align="right"><font size="2"><%=formatnumber(pamt1,2)%></font></td>
	<td align="right"><font size="2"><%=formatnumber(saveamt,2)%></font></td>
	<td align="right"><font size="2"><%=formatnumber(samt,2)%></font></td>
   </tr> 
<%	
                xmemno = rs("memno")
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
%>
	<tr>
		<td></td>
		<td></td>              		
                 <td></td>
                 <td></td>
		 <td>�`���B �G</td>
                
		<td width=100 align="right"><%=formatNumber(ttlpint1,2)%></td>
		<td width=100 align="right"><%=formatNumber(ttlpamt1,2)%></td>	
                <td width=100 align="right"><%=formatNumber(ttlsamt,2)%></td>		
		<td width=100 align="right"><%=formatNumber(ttlTemp,2)%></td>
	

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
