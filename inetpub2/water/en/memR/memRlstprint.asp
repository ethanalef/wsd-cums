<!-- #include file="../conn.asp" -->

<%
sqlstr=" and "
mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())


sxdate=request.form("stdate1")


mdate= dateserial(right(sxdate,4),mid(sxdate,4,2),left(sxdate,2))


set rs = server.createobject("ADODB.Recordset")

sql = "select a.memno,a.memname,a.memcname,a.membday,a.memhkid,a.memgender,a.mstatus,B.CODE,SUM(b.amount) from memmaster a,share b  where (a.wdate is null or a.wdate<='"&mdate&"' ) and  a.memno=b.memno and  b.ldate<='"&mdate&"'  GROUP BY a.memno,a.memcname,a.membday,a.memhkid,a.memgender,a.mstatus,B.CODE order by a.memno,a.memcname,a.membday,a.memhkid,a.memgender,a.mstatus,B.CODE"

rs.open sql, conn, 1, 1

if rs.eof then
   response.redirect "memstlst.asp"
end if

ttlamt = 0
    
if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"

end if
%>
<html>
<head>
<title>�������p�C�L(���U�x)</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table border="0" cellpadding="0" cellspacing="0">
	<tr height="30" valign="top" align="center"><td colspan="15"><font size="3" face="�з���"  >���ȸp���u�x�W���U��<br>�������i(���U�x)<br><font size="2"  face="�з���" >��� : <%=mndate%></font></font></td></tr>
        <tr height="30" ><td colspan=9></td></tr>


	<tr height="15" valign="bottom">
	<td width="80" align="center"><font size="3" face="�з���"  >�����s��</font></td>
	<td width=180 align="center"><font size="3"  face="�з���" >�^��m�W</font></td>
		<td width=70 align="center"><font size="3"  face="�з���" >����m�W</font></td>	
	<td width="130" align="center"><font size="3" face="�з���"  >�Ѫ����l</font></td>
        <td width="130" align="center"><font size="3" face="�з���"  >�U�ڵ��l</font></td>
        <td width="130" align="center"><font size="3" face="�з���"  >����</font></td> 
	</tr>
	<tr><td colspan=7><hr></td></tr>
<% 
   tlnamt = 0
 
   ttlamt = 0 
   ttl1   = 0
   ttl2   =0
   ttl3   =0
   ttl4   =0
   ttl5   = 0
   ttl6   =0
   ttl7   =0
   ttl8   =0
   ttl9   = 0
   ttl10   =0
   ttl11   =0
   ttl12   =0
   ttl13   = 0
   ttl14   =0
   ttl15   =0
   ttl16   =0
   ttl17   = 0
   ttl18   =0
   ttl19   =0
   cnt1   = 0
   cnt2   =0
   cnt3   =0
   cnt4   =0
   cnt5   = 0
   cnt6   =0
   cnt7   =0
   cnt8   =0
   cnt9   = 0
   cnt10   =0
   cnt11   =0
   cnt12   =0
   cnt13   = 0
   cnt14   =0
   cnt15   =0
   cnt16   =0
   cnt17   = 0
   cnt18   =0
   cnt19   =0
   xamt = 0
   lnamt = 0
   clsbal = 0
   ttlamt = 0
   ttllncnt = 0
   xmemno=rs("memno")
   memcname=rs("memcname")
   xmstatus=  rs("mstatus")
   xmembday = rs("membday")
   xmemhkid = rs("memhkid")
   xmemgender =rs("memgender")
   if xmembday<>"" then
         membday = right("0"&day(xmembday),2)&"/"&right("0"&month(xmembday),2)&"/"&year(xmembday)
                  age =year(date()) -year(xmembday)
                  yy = year(date())
                  adatae = dateserial(yy,1,1)
                  mm = month(xmembday)
                  dd = day(xmembday)
                   
                  bdate = dateserial(yy,mm,dd)
                  xday = ((bdate - adate)+z)/365
                  if xday >0.5 then
                     age = age + 1
                  end if
 else
       age = 0
       membday=""  
  end if
      if xmemGender="M" then
         
         sex="����"
      else
         
          sex="�k�h"
      end if
   do while not rs.eof


    
      if xmemno<> rs("memno") then
 
         xamt = 0
         lnamt = 0
         ylnnum=""
         set ms = server.createobject("ADODB.Recordset")
         mssql = "select lnnum,appamt,bal from loanrec where memno ='"&xmemno&"' and repaystat='N'   "
         ms.open mssql, conn, 1, 1
         if not ms.eof then
           
            lnamt  = ms("bal")
                           
         end if
       
         ms.close

         if lnamt > 0 then
            tlnamt = tlnamt + lnamt 
            ttllncnt = ttllncnt + 1
        
         end if 
         select case xmstatus
                case "A"
                     if clsbal > 0 then
                        ttl1 = ttl1 + clsbal
                        cnt1 = cnt1 + 1
                     end if
                     idx ="�۰���b(ALL)"
                case "B"
                     if clsbal > 0  or (clsbal=0 and lnamt >0)  then
                        ttl2 = ttl2 +clsbal
                        cnt2 = cnt2 + 1
                      end if
                     
                     idx ="�}��"
                case "C"
                   if clsbal > 0  or (clsbal=0 and lnamt >0)  then
                       ttl3 = ttl3+ clsbal
                       cnt3 = cnt3 + 1
                     end if 
                     idx ="�h��"
                case "D"
                    if clsbal > 0 then
                       ttl4 = ttl4 + clsbal
                       cnt4 = cnt4 + 1 
                    end if 
                     idx ="�N��"
                case "F"
                    ttl5 = ttl5 +  clsbal
                    cnt5 = cnt5 + 1
                     idx ="�S�O�Ӯ�"
                case "H"
                   ttl6 = ttl6 + clsbal
                   cnt6 = cnt6 + 1
                     idx ="�Ȱ��Ȧ�"
                case  "J"
                   ttl7 = ttl7 + clsbal
                   cnt7 = cnt7 + 1
                     idx ="�s��"
                case "L"
                 if clsbal > 0 or (clsbal=0 and lnamt>0) then
                   ttl8 = ttl8 + clsbal
                   cnt8 = cnt8 + 1  
                 end if
                     idx ="�b�b"
                case "M"
                  ttl9 = ttl9 + clsbal
                  cnt9 = cnt9 + 1
                     idx ="�w��,�Ȧ�"
                case "N"     
                 ttl10 = ttl10 + clsbal
                 cnt10 = cnt10 + 1
                     idx ="���`"            
                case "P"
                if clsbal > 0 then
                   ttl11 = ttl11 + clsbal
                   cnt11 = cnt11 + 1
                end if   
                     idx ="�h�@"
                case "T"
                 ttl12 = ttl12 + clsbal
                 cnt12 = cnt12 + 1
                     idx ="�w��"
                case "V"
                     if clsbal> 0 or (clsbal=0 and lnamt >0 ) then
                        ttl13 = ttl13 + clsbal
                        cnt13 = cnt13 + 1 
                     end if
                     idx ="IVA"
                    
                case "0"
                 ttl14 = ttl14 + clsbal
                 cnt14 = cnt14 + 1
                     idx =" �۰���b(�Ѫ�)"
                case "1"
                 ttl15 = ttl15 + clsbal
                 cnt15 = cnt15 + 1
                     idx ="�۰���b(�Ѫ�,�Q��)"
                case "2"
                 ttl16 = ttl16 + clsbal
                 cnt16 = cnt16 + 1
                     idx ="�۰���b(�Ѫ�,����)"
                case "3"
                 ttl17 = ttl17 + clsbal
                 cnt17 = cnt17 + 1
                     idx ="�۰���b(�Q��,����)"
                case "8"
                 ttl18 = ttl18 + clsbal
                 cnt18 = cnt18 + 1 
                     idx ="�פ���y��b"
                case "9"
                 ttl19 = ttl19 + clsbal
                 cnt19 = cnt19 + 1
                     idx ="�פ���y���`"
         end select    
    if clsbal > 0 or(clsbal=0 and lnamt > 0 ) then
   
%>
     <tr>
          <td width="80"><%=xmemno%></td>
          <td width=180 align="left"><%=rs("memname")%></td>  
          <td width="80"><font size="3" face="�з���"  ><%=memcname%></font></td>
          <td width="130" align="right"><%=formatnumber(clsbal,2)%></td>
          <td width="130" align="right"><%=formatnumber(lnamt,2)%></td>
          <td width="150" align="center"><font size="3" face="�з���"  ></font><%=idx%></font></td>
     </tr>

<%
 
        ttlamt = ttlamt + clsbal
        
        end if    
        xmemno=rs("memno")
          memcname=rs("memcname")
          xmstatus=rs("mstatus")
  xmembday = rs("membday")
   xmemhkid = rs("memhkid")
   xmemgender =rs("memgender")   
  if xmembday<>""   then
        membday = right("0"&day(xmembday),2)&"/"&right("0"&month(xmembday),2)&"/"&year(xmembday)
                  age =year(date()) -year(xmembday)
                  yy = year(date())
                  adatae = dateserial(yy,1,1)
                  mm = month(xmembday)
                  dd = day(xmembday)
                   
                  bdate = dateserial(yy,mm,dd)
                  xday = ((bdate - adate)+z)/365
                  if xday >0.5 then
                     age = age + 1
                  end if
  else
       age = 0
       membday=""  
  end if
      if xmemGender="M" then
         
         sex="����"
      else
         
          sex="�k�h"
      end if 
        clsbal = 0
        lnamt = 0
        xamt = 0
   end if  

   select case left(rs("code"),1)
          case "A","C","0"
               clsbal = clsbal + rs(7)
          case "B","G","H"
                clsbal = clsbal - rs(7)
         end select


 
    
     rs.movenext
    loop
 
         xamt = 0
         lnamt = 0
         ylnnum=""
         set ms = server.createobject("ADODB.Recordset")
         mssql = "select lnnum,appamt,bal from loanrec where memno ='"&xmemno&"' and repaystat='N'   "
         ms.open mssql, conn, 1, 1
         if not ms.eof then
           
            lnamt  = ms("bal")
                           
         end if
       
         ms.close

         if lnamt > 0 then
            tlnamt = tlnamt + lnamt 
            ttllncnt = ttllncnt + 1
         end if  
          select case xmstatus
                case "A"
                     if clsbal > 0 then
                        ttl1 = ttl1 + clsbal
                        cnt1 = cnt1 + 1
                     end if
                     idx ="�۰���b(ALL)"
                case "B"
                     if clsbal > 0  or (clsbal=0 and lnamt >0)  then
                        ttl2 = ttl2 +clsbal
                        cnt2 = cnt2 + 1
                      end if
                     
                     idx ="�}��"
                case "C"
                   if clsbal > 0  or (clsbal=0 and lnamt >0)  then
                       ttl3 = ttl3+ clsbal
                       cnt3 = cnt3 + 1
                     end if 
                     idx ="�h��"
                case "D"
                    if clsbal > 0 then
                       ttl4 = ttl4 + clsbal
                       cnt4 = cnt4 + 1 
                    end if 
                     idx ="�N��"
                case "F"
                    ttl5 = ttl5  + clsbal
                    cnt5 = cnt5 + 1
                     idx ="�S�O�Ӯ�"
                case "H"
                   ttl6 = ttl6 + clsbal
                   cnt6 = cnt6 + 1
                     idx ="�Ȱ��Ȧ�"
                case  "J"
                   ttl7 = ttl7 + clsbal
                   cnt7 = cnt7 + 1
                     idx ="�s��"
                case "L"
                  if clsbal > 0 or (clsbal=0 and lnamt>0) then
                   ttl8 = ttl8 + clsbal
                   cnt8 = cnt8 + 1  
                 end if
                     idx ="�b�b"
                case "M"
                  ttl9 = ttl9 + clsbal
                  cnt9 = cnt9 + 1
                     idx ="�w��,�Ȧ�"
                case "N"     
                 ttl10 = ttl10 + clsbal
                 cnt10 = cnt10 + 1
                     idx ="���`"            
                case "P"
                if clsbal > 0 then
                   ttl11 = ttl11 + clsbal
                   cnt11 = cnt11 + 1
                end if   
                     idx ="�h�@"
                case "T"
                 ttl12 = ttl12 + clsbal
                 cnt12 = cnt12 + 1
                     idx ="�w��"
                case "V"
                      if clsbal> 0 or (clsbal=0 and lnamt >0 ) then
                        ttl13 = ttl13 + clsbal
                        cnt13 = cnt13 + 1 
                     end if
                     idx ="IVA"
                case "0"
                 ttl14 = ttl14 + clsbal
                 cnt14 = cnt14 + 1
                     idx =" �۰���b(�Ѫ�)"
                case "1"
                 ttl15 = ttl15 + clsbal
                 cnt15 = cnt15 + 1
                     idx ="�۰���b(�Ѫ�,�Q��)"
                case "2"
                 ttl16 = ttl16 + clsbal
                 cnt16 = cnt16 + 1
                     idx ="�۰���b(�Ѫ�,����)"
                case "3"
                 ttl17 = ttl17 + clsbal
                 cnt17 = cnt17 + 1
                     idx ="�۰���b(�Q��,����)"
                case "8"
                 ttl18 = ttl18 + clsbal
                 cnt18 = cnt18 + 1 
                     idx ="�פ���y��b"
                case "9"
                 ttl19 = ttl19 + clsbal
                 cnt19 = cnt19 + 1
                     idx ="�פ���y���`"
         end select    
   if clsbal > 0 or(clsbal=0 and lnamt > 0 ) then
 %>
     <tr>
          <td width="80"><%=xmemno%></td>
          <td width="80"><font size="3" face="�з���"  ><%=memcname%></font></td>

          <td width="130" align="right"><%=formatnumber(clsbal,2)%></td>
          <td width="130" align="right"><%=formatnumber(lnamt,2)%></td>
          <td width="150" align="center"><font size="3" face="�з���"  ><%=idx%></font></td>
     </tr>

<%
        ttlamt = ttlamt + clsbal
       
 end if
 %>

     	<tr><td colspan=7><hr></td></tr>
        <tr>
            
              <td></td>
              <td></td>  
                 
             <td width="130" align="right"><%=formatnumber(ttlamt,2)%></td>
             <td width="130" align="right"><%=formatnumber(tlnamt,2)%></td>
         </tr>
        <tr><td></td>
 
         
             <td></td>             
             <td width="130" align="right">============</td>
              <td width="130" align="right">============</td>
         </tr>	


</table>
<BR>

<BR>


</center>
</body>
</html>

