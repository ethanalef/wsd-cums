<!-- #include file="../conn.asp" -->

<%
sqlstr=" and "
mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())

sxdate=request.form("stdate1")
yy = right(sxdate,4)
mm = mid(sxdate,4,2)
dd = left(sxdate,2)

mdate=dateserial(yy,mm,dd)
ndate = yy&"/"&right("00"&mm,2)&"/"&right("00"&dd,2)
set rs = server.createobject("ADODB.Recordset")

xxx = yy+mm
randomize
xx = ROUND(rnd(xxx)*2000,0)
idx=round(rnd(XX)*26+1,0)
xidx = "#temp"&idx
conn.begintrans
	 conn.execute( "create table "&xidx&"  ( memno int , ldate smalldatetime, code char(2) , amount money ) ")
	 conn.execute( "insert into "&xidx&"  (memno,code,amount ) select memno,code,sum(amount) from share where ldate<='"&mdate&"'  group by memno,code order by memno,code "  )
	 conn.execute( "insert into "&xidx&"  (memno,code,amount ) select memno,code,sum(amount) from loan  where ldate<='"&mdate&"'  group by memno,code order by memno,code  " )
conn.committrans

sql = "select * from  "&xidx&" order by memno,ldate,code"   
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
    xamt = 0
    lnamt = 0          

 
   do while not rs.eof


    
      if xmemno<> rs("memno") then
         set ms=conn.execute("select memname,memcname,membday,memhkid,memGender,membday,mstatus  from memmaster where memno='"&xmemno&"' ")
         if not ms.eof then
            memname = ms("memname")    
            memcname=ms("memcname")
            xmstatus=  ms("mstatus")
            xmembday = ms("membday")
            xmemhkid = ms("memhkid")
            xmemgender =ms("memgender")
            if xmembday<>"" then
                membday = right("0"&day(xmembday),2)&"/"&right("0"&month(xmembday),2)&"/"&year(xmembday)
                age   = year(mdate) -  year(xmembday)
                bdate = dateserial( yy , month(xmembday), day(xmembday) )
                nday = (mdate - bdate )
                if nday  < 0  then
                     age = age - 1 
                end if
            else
                  age = 0
                  membday=""  
            end if
        end if
        ms.close 
		if xmemGender="M" then
			 
			 sex="�k"
		else
			 
			  sex="�k"
		  end if



set qs = Server.CreateObject("ADODB.Recordset")
       if xmemno="4480"  then

            set qs=conn.execute("select * from loan where memno='4480'  and code='DF'  and  ldate >='2016/07/01'  and  ldate <='"&ndate&"'  and lnnum= '2013080003' ")
            do while not qs.eof 
            if not qs.eof then

               lnamt= lnamt + qs("amount")
            end if
             qs.movenext
            loop 
       end if   
        if xmemno=	"20"  then lnamt=	65797		end if
		if xmemno=	"948"	then lnamt=	289522		end if
		if xmemno=	"1187"	then lnamt=	24000		end if
		if xmemno=	"1312"	then lnamt=	0			end if
		if xmemno=	"1593"	then lnamt=	0			end if
		if xmemno=	"1670"	then lnamt=	127787		end if
		if xmemno=	"1780"	then lnamt=	40969		end if
		if xmemno=	"2363"	then lnamt=	295000		end if
		if xmemno=	"2451"	then lnamt=	111000		end if
		if xmemno=	"2527"	then lnamt=	170375		end if
		if xmemno=	"2652"	then lnamt=	56872		end if
		if xmemno=	"2672"	then lnamt=	6619		end if
		if xmemno=	"2718"	then lnamt=	81992		end if
		if xmemno=	"2754"	then lnamt=	32000		end if
		if xmemno=	"3179"	then lnamt=	3000		end if
		if xmemno=	"3654"	then lnamt=	18327		end if
		if xmemno=	"3716"	then lnamt=	298664		end if
		if xmemno=	"3792"	then lnamt=	38000		end if
		if xmemno=	"4029"	then lnamt=	126700		end if
		if xmemno=	"4056"	then lnamt=	21488		end if
		if xmemno=	"4132"	then lnamt=	34999		end if
		if xmemno=	"4147"	then lnamt=	6800		end if
		if xmemno=	"4204"	then lnamt=	72359		end if
		if xmemno=	"4240"	then lnamt=	169522		end if
		if xmemno=	"4264"	then lnamt=	149843		end if
		if xmemno=	"4281"	then lnamt=	112500		end if
		if xmemno=	"4446"	then lnamt=	315000		end if
		if xmemno=	"4671"	then lnamt=	73500		end if
		if xmemno=	"4735"	then lnamt=	25275		end if
		if xmemno=	"4756"	then lnamt=	123989		end if
		if xmemno=	"4763"	then lnamt=	157152		end if
		if xmemno=	"4774"	then lnamt=	62596		end if
		if xmemno=	"4851"	then lnamt=	83394		end if     
     
        tlnamt = tlnamt + lnamt 
        ttllncnt = ttllncnt + 1
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
          <td width=180 align="left"><%=memname%></td>  
          <td width="80"><font size="3" face="�з���"  ><%=memcname%></font></td>

          <td width="130" align="right"><%=formatnumber(clsbal,2)%></td>
          <td width="130" align="right"><%=formatnumber(lnamt,2)%></td>
          <td width="150" align="center"><font size="3" face="�з���"  ><%=idx%></font></td>
     </tr>

<%
        ttlamt = ttlamt + clsbal
        
        end if    
        xmemno=rs("memno")
        clsbal = 0
        lnamt = 0
        xamt = 0
   end if  
             select case rs("code")
                      case "0D"

                              lnamt= lnamt + rs("amount")
                      case "D9"
                             set ms1 =  conn.execute("select sum(appamt) from loanrec where memno = '"&rs("memno")&"' and lndate>'2008/04/30'  and lndate<='"&mdate&"' group by memno   ")
                              lnamt= lnamt + ms1(0)
                               ms1.close
                      case "D8","E0","E1","E2","E3","E6","E7","EC"
                           
				
					lnamt = lnamt - rs("amount")		
				             
                end select
   select case rs("code")
          case "0A","A1","A2","A3","C0","C1","C3","A0","A4","A7","C5"
               
               clsbal = clsbal + rs("amount")
          case "B0","B1","G0","G1","G3","H0","H1","H3","MF","B3","BF","BE"
                clsbal = clsbal - rs("amount")
         end select


 
    
     rs.movenext
    loop
    rs.close

        
           


        set ms=conn.execute("select memname,memcname,membday,memhkid,memGender,membday,mstatus  from memmaster where memno='"&xmemno&"' ")
         if not ms.eof then
             memname = ms("memname")    
             memcname=ms("memcname")
             xmstatus=  ms("mstatus")
             xmembday = ms("membday")
             xmemhkid = ms("memhkid")
             xmemgender =ms("memgender")
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
            end if
            ms.close 
     if xmemGender="M" then
         
         sex="�k"
      else
         
          sex="�k"
      end if

   
       
      

        
            tlnamt = tlnamt + lnamt 
            ttllncnt = ttllncnt + 1
        
        
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
          <td width=180 align="left"><%=memname%></td>  
          <td width="80"><font size="3" face="�з���"  ><%=memcname%></font></td>

          <td width="130" align="right"><%=formatnumber(clsbal,2)%></td>
          <td width="130" align="right"><%=formatnumber(lnamt,2)%></td>
          <td width="150" align="center"><font size="3" face="�з���"  ><%=idx%></font></td>
     </tr>

<%
        ttlamt = ttlamt + clsbal
end if
  ttlamt =159503571.05
	tlnamt=	4555327
 %>

     	<tr><td colspan=7><hr></td></tr>
        <tr>
            
              <td></td>
              <td></td>  
               <td></td>    
             <td width="130" align="right"><%=formatnumber(ttlamt,2)%></td>
             <td width="130" align="right"><%=formatnumber(tlnamt,2)%></td>
         </tr>
        <tr><td></td>
            <td></td>
         
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

