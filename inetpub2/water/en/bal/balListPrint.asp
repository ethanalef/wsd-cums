<!-- #include file="../conn.asp" -->

<%

mndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
xdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())

mPeriod = request.form("mPeriod")
yy  = cint(left(mperiod,4))
mm = cint(right(mPeriod,2))
xyy =  yy
xmm = mm + 1
if xmm= 13 then
   xyy = xyy + 1
   xmm = xmm - 12
end if
xdd= 1

nPeriod = left(mPeriod,4)&"."&right(mPeriod,2)&".01"
pPeriod = xyy&"."&right("0"&xmm,2)&".01"
xperiod = xyy&"/"&right("0"&xmm,2)&"/"&"01"
xxdate=DateSerial(yy,mm,1)
yydate=dateSerial(yy,mm+1,1-1)

ryy = yy
rmm = mm - 1
if rmm=0 then
    ryy = yy - 1
    rmm = 12
end if
RPeriod = ryy&rmm
        actln = 0
         glnamt = 0
server.scripttimeout = 1800
        set rs = conn.execute("select count(*) from memmaster where memdate>='"&nperiod&"' and memdate <'"&xperiod&"' ")
         if not rs.eof then
           
               nmemcnt = rs(0)
         end if 
         rs.close       

        set rs = server.createobject("ADODB.Recordset")
         sql = "Select memno,max(ldate) as xldate  from  share  group by memno order by memno "'  " 
         rs.open sql, conn,1,1
         if not rs.eof then 
         do while not rs.eof
              savettl = 0
            if rs("xldate")>=xxdate and rs("xldate")<=yydate then
               set ms = conn.execute("select * from share where memno ='"&rs("memno")&"' ")
               do while not ms.eof
                  select case left(ms("code"),1)
                        case "0","A","C"
                             if ms("code")<>"AI" then
                                savettl = savettl + ms("amount")
                             end if
                        case "B","G","H","M"
                             savettl = savettl - ms("amount")
                  end select 
                  sxldate = ms("ldate")
                  ms.movenext
               loop
               ms.close
               if savettl = 0 and sxldate>= xxdate and sldate<=yydate then 
                  set ls = conn.execute("select * from loanrec where memno='"&rs("memno")&"' and cleardate>='"&xxdate&"' and cleardate<='"&yydate&"' ")
                  if not rs.eof then 
 
                     omemcnt = omemcnt + 1
                  end if
                  ls.close
               end if           
               end if   
               rs.movenext
            loop

            end if
            rs.close
        set rs = server.createobject("ADODB.Recordset")
         set ms = server.createobject("ADODB.Recordset")  
         sql = "Select memno from  loanrec   where cleardate>='"&xxdate&"' and cleardate<='"&yydate&"' " 
         rs.open sql, conn,1,1
         if not rs.eof then 
         do while not rs.eof
           
               savettl = 0
               set ms = conn.execute("select * from share where memno ='"&rs("memno")&"' ")
               do while not ms.eof
                  select case left(ms("code"),1)
                        case "0","A","C"
                             if ms("code")<>"AI" then
                                savettl = savettl + ms("amount")
                             end if
                        case "B","G","H","M"
                             savettl = savettl - ms("amount")
                  end select 
                 
                  ms.movenext
               loop
               ms.close
               if savettl = 0  then 
 
                   omemcnt = omemcnt + 1
                   
               end if
          
               rs.movenext
            loop

            end if
            rs.close      
  
         memcnt = 0
         set rs = server.createobject("ADODB.Recordset")   
         savettl = 0
         sql = "select a.memno,a.code,a.amount as ammt,b.membday,b.mstatus "&_
                       " from share a,memmaster b where a.memno=b.memno and "&_                     
                       "a.ldate<='"&yydate&"' order by a.memno,a.code,b.membday ,b.mstatus "
         rs.open sql, conn, 1, 1 
        
         IF NOT RS.EOF  THEN
            xmemno=rs(0)
            xmstatus=rs("mstatus")
         do while  not rs.eof 
            if xmemno <> rs("memno") then 
               if savettl > 0  then
 
                  conn.execute("insert into shamt (memno,shamt) values ('"&xmemno&"' ,"&savettl&" ) ")
                  memcnt = memcnt + 1
  
               end if
               savettl = 0
               xmemno= rs(0)
               xmstatus=rs("mstatus") 
            end if
               select case rs("code")
                      case "0A","A1","A2","A3","C0","C1","C3","A0","A4","A7"
                           savettl = savettl + rs("ammt")
                      case "B0","B1","G0","G1","G3","H0","H1","H3","MF"
                           savettl = savettl - rs("ammt") 
               end select 
             
         rs.movenext
         loop
              if savettl > 0  then
                
                  memcnt = memcnt + 1
                   conn.execute("insert into shamt (memno,shamt) values ('"&xmemno&"' ,"&savettl&" ) ")
               end if
         END IF
         rs.close 
       
         vmemcnt = 0 
          llncnt = 0 
          ttllncnt = 0
         set rs = conn.execute("select a.*,b.mstatus  from loan a ,memmaster b where a.ldate < '"&nperiod&"' and a.memno=b.memno  order by a.memno,a.ldate,a.code ")
         if not rs.eof    then
            xmemno=rs("memno")
            lnamt = 0
            xmstatus = rs("mstatus")
         do while not rs.eof
                   if xmemno<>rs("memno") then
                      if lnamt >0   then
                         ttllncnt =  ttllncnt + 1
 
                       end if
                       xmemno=rs("memno")
                       lnamt = 0
                        xmstatus = rs("mstatus") 
                   end if 
            select case rs("code")

                   case "0D","D1"
                        lnamt = lnamt+ rs("amount")
                   case "D0","E0","E1","E2","E3","E4","E6","E7","EC"
                        lnamt = lnamt - rs("amount")
            end select 
            rs.movenext
         loop
                      if lnamt >0  then
                          ttllncnt =  ttllncnt + 1
                      end if
   
                       lnamt = 0
         END IF          
         rs.close

        set rs1 = conn.execute("select code,sum(amount) from loan where ldate < '"&nperiod&"' group by code  ")
        do while  not rs1.eof    
            select case rs1(0)
                   case "0D","D1"  
                        glnamt  =  glnamt + rs1(1)          
                   case "D0","E0","E1","E2","E3","E6","E7","EC"
                        glnamt = glnamt - rs1(1)
             end select
         rs1.movenext
         loop 
         rs1.close
         set rs = conn.execute("select chequeamt,appamt,lnflag from loanrec where lndate >= '"&nperiod&"' and lndate <'"&pperiod&"'   ")
         do while  not rs.eof    
            if rs(2)="Y" then
               actln  =  actln + rs(0)          
            else
               actln = actln + rs(1)
            end if
         rs.movenext
         loop
         rs.close  
 

         nwlncnt = 0
         set rs = conn.execute("select  COUNT(*) AS Expr1 from loanrec where lndate >= '"&xxdate&"' and lndate <='"&yydate&"' ")
         if not rs.eof    then
 
             nwlncnt =  rs(0)

         END IF 
         rs.close

         set rs = conn.execute("select appamt,bal,chequeamt,convert(char(10),lndate,102) as slndate   from loanrec where repaystat='N' and lndate< '"&xperiod&"'  ")
         do while not rs.eof
          
            if rs("slndate")>= nperiod  then   
               
               ttllnamt = ttllnamt + rs("appamt")
               ttlbal   = ttlbal + rs("bal")
           else
               ottllnamt = ottllnamt + rs("chequeamt")
               ottlbal = ittlbal + rs("bal")
           end if
           lncount  = lncount + 1 
           rs.movenext
        loop
        rs.close
       clncnt = 0
       set rs = conn.execute("select count(*)  from loanrec  where  repaystat='C' and cleardate>='"&xxdate&"' and  cleardate<='"&yydate&"' ")
       if not rs.eof then
          clncnt = rs(0)
       end if
       rs.close
       xclncnt = 0
       set rs = conn.execute("select count(*)  from loanrec  where   lndate>='"&xxdate&"' and  lndate<='"&yydate&"' and lnflag='Y' ")
       if not rs.eof then
          xclncnt = rs(0)
       end if
       rs.close

        ttlnwlnamt = 0
         set rs = conn.execute("select code,sum(amount) as samount from loan where  ldate >='"&nperiod&"' and ldate <'"&xperiod&"' group by code  ")
         do while not rs.eof
         

   
 
               select case rs("code")
                      case  "E1"
                           lbnkamt =  rs("samount")
                      case "E2"
                          lsadamt = rs("samount")
                      case  "E3" 
                          lchamt =  rs("samount")
                      case "E0","E6","E7","EC"
                          ajlnamt =ajlnamt + rs("samount") 
                      case  "F1"
                           ibnkamt =  rs("samount")
                      case "F2"
                          isadamt = rs("samount")
                      case  "F3" 
                          ichamt =  rs("samount")
                      case "F0","F6","F7"
                          ajintamt = ajintamt + rs("samount")

 
                      case "ET"
                           esavamt = rs("samount")
                      case "FT"
                           fsavamt =  rs("samount")
                     case "EC"
                          ajlnamt =ajlnamt + rs("samount")
             end select                    

               
          
             rs.movenext
             loop
         rs.close          
    
        ajshamt = 0
         if nperiod >"2008.04.30" then
            ttlamt = 1
         else
             chamt = 1
         end if
         set rs1 = conn.execute("select code,convert(char(10),ldate,102) as pydate  ,amount,lnflag from share where ldate<'"&xperiod&"'order by memno,ldate,code ")
         do while not rs1.eof
            curdate = rs1(1)

            if rs1("pydate")<  nperiod then
               select case rs1("code")
                      case "0A", "A1","A2","A3","C0","C1","C3","A0","A4","A7"
                           ttlamt = ttlamt + rs1("amount")
                      case "B0","B1","B2","B3","G0","G1","G3","H0","H1","H3","MF"
                           ttlamt = ttlamt - rs1("amount")
             end select
             else
              
               select case rs1("code")
                      case "MF"
                           ajwdamt = ajwdamt + rs1("amount")
                      case  "A1"
                           bnkamt = bnkamt + rs1("amount")
                      case "A2"
                           sadamt = sadamt + rs1("amount")
                      case  "A3" ,"0A"
                          chamt = chamt + rs1("amount")
                      case "A0","A4","A7"
                           ajshamt = ajshamt + rs1("amount")
                      case "C3"
                           divamt3 = divamt3 + rs1("amount")
                      
                      case "C1"
                           divamt1 = divamt1 + rs1("amount")
                      case "C0"
                           ajdivamt = ajdivamt + rs1("amount")
                      case "B1"
                          if rs1("lnflag")="Y" then
                              ajwdamt = ajwdamt + rs1("amount")
                          else
                              withdamt = withdamt + rs1("amount")
                          end if
                           
                      case "B0"
                            if rs1("amount") >= 0 then
                               ajwdamt = ajwdamt + rs1("amount")
                            else
                               ajshamt = ajshamt + rs1("amount")*-1
                            end if

                      case "B6"
                            ajshamt = ajshamt + rs1("amount")
                  case "G3"
			Gamt3 = Gamt3+rs1("amount")
                  case "H3"
			Hamt3 = Hamt3+rs1("amount")
                 case "G2"
			Gamt2 = Gamt2+rs("amount")
                  case "H2"
			Hamt2 = Hamt2+rs1("amount")
                 case "G1"
			Gamt1 = Gamt1+rs1("amount")
                  case "H1"
			Hamt1 = Hamt1+rs1("amount")
               
             end select                    
             end if
      
             rs1.movenext
             loop
         rs1.close
         
         csttlamt = bnkamt+sadamt+chamt-gamt3-hamt3+ ajshamt+0
         cwttlamt = withdamt+ajwdamt+0
         cajttlamt = ajshamt+ ajdivamt+ajlnamt+ajintamt+0
         cloanamt   = actln+0
         ttldiv = divamt1  + divamt3+ajdivamt+0
         ttlgamt = Gamt1 + Gamt2 + Gamt3+0
         ttlhamt = Hamt1 + Hamt2 + Hamt3+0
         cpayamt  = lbnkamt +lsadamt+lchamt+ ajlnamt+0
         cintamt  = ibnkamt +isadamt+ichamt+ajintamt+0
         ttlbnk =  lbnkamt+bnkamt + ibnkamt + divamt1+0
         ttlsad =  lsadamt+sadamt + isadamt+0
         ttlch  =  lchamt +chamt  + ichamt + divamt3 +0
         ttlrec = ttlbnk + ttlsad + ttlch + cajttlamt+0 
         gttlamt = csttlamt+ttlamt-cwttlamt+ttldiv+0
              
         payamt = actln + withdamt +0
	 ttlpay = payamt + ajwdamt +0
        
         actlnamt = glnamt + cloanamt - cpayamt +0
         actlncnt = ttllncnt +nwlncnt - clncnt  +0
         ttlmem = memcnt - nmemcnt+0+omemcnt
         oclncnt = clncnt - xclncnt 
         actmem = memcnt   +0
         gttlrate= round(actlnamt / gttlamt*100,0) +0 

if request.form("output")="word" then
	Response.ContentType = "application/msword"
elseif request.form("output")="excel" then
	Response.ContentType = "application/vnd.ms-excel"

end if
%>
<html>
<head>
<title>�C��b�έp�C��</title>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="10" topmargin="10" marginheight="0" marginwidth="0">
<center>
<table width="1012" border="0">

 <tr>
    <td width="99">&nbsp;</td>
    <td width="780">&nbsp;</td>
    <td width="142">&nbsp;</td>
  </tr>
	<tr>
        <td>&nbsp</td>
        <td align="center"><b><font size="4"  face="�з���" >���ȸp���u�x�W���U��<br>�C��b�έp��</font></b?</td>
        <td align="center"><font size="2"  face="�з���" >��� : <%=mndate%></font></td>
        </tr>
       

</table>
<table border="1" cellspacing="1" cellpadding="4" align="center"  bgcolor="336699">
	<tr bgcolor="#FFFFFF" align="center">
		<td><font size="2" ><%=yy%>/<%=right("0"&mm,2)%></font></td>
		<td><font size="2" >�ȡ@��</font></td>
		<td><font size="2" >�w�@��</font></td>
		<td><font size="2" >�{�@��</font></td>
		<td><font size="2" >�ա@��</font></td>
		<td><font size="2" ><%=mm%>��</font></td>
		<td><font size="2" ><%=rmm%>��</font></td>
		<td><font size="2" >�` �@��</font></td>

	</tr>
        <tr bgcolor="#FFFFFF">
             <td>  �Ѫ��@</td>
             <td align="right" ><%=formatnumber(bnkamt,2)%></td>
	     <td align="right" ><%=formatnumber(sadamt,2)%></td>
	     <td align="right" ><%=formatnumber(chamt-Gamt3-Hamt3,2)%></td>
             <td align="right" ><%=formatNUMBER(ajshamt,2)%></td>
	     <td align="right" ><%=formatnumber(csttlamt,2)%></td>
	     <td align="right" ><%=formatnumber(ttlamt,2)%></td>
	     <td align="right" ><%=formatnumber(gttlamt,2)%></td>
        </tr> 
         <tr bgcolor="#FFFFFF">
             <td>  �h�ѡ@</td>
             <td align="right" ><%=formatnumber(withdamt,2)%></td>
	     <td>�@</td>
	     <td>�@</td>
             <td align="right" ><%=formatNUMBER(ajwdamt,2)%></td>	     
	     <td align="right" ><%=formatnumber(cwttlamt,2)%></td>
             <td>�@</td>
	     <td>�@</td>
        </tr>  
         <tr bgcolor="#FFFFFF">
             <td>  �Ѯ��@</td>
             <%if divamt1 <> 0 then %>
             <td align="right" ><%=formatnumber(divamt1,2)%></td>
             <%else%>
             <td>�@</td>
             <%end if%>
             <%if divamt2 <> 0 then %>
	     <td align="right" ><%=formatnumber(divamt2,2)%></td>
             <%else%>
             <td>�@</td>
             <%end if%>
             <%if divamt3 <> 0 then %>
	     <td align="right" ><%=formatnumber(divamt3,2)%></td>
             <%else%>
             <td>�@</td>
             <%end if%>
             <%if ajdivamt <> 0 then %>
             <td align="right" ><%=formatNUMBER(ajdivamt,2)%></td>
             <%else%>
             <td>�@</td>
             <%end if%>
             <%if ttldiv <> 0 then %> 
	     <td align="right" ><%=formatnumber(ttldiv,2)%></td>
             <%else%>
             <td>�@</td>
             <%end if%>
             <td>�@</td>
	     <td>�@</td>
        </tr>   
         <tr bgcolor="#FFFFFF">
             <td>  �U�ڡ@</td>
             <td align="right" ><%=formatnumber(actln,2)%></td>
	     <td>�@</td>
	     <td>�@</td>
             <td>�@</td>	     
	     <td align="right" ><%=formatnumber(cloanamt,2)%></td>
             <td align="right" ><%=formatNUMBER(glnamt,2)%></td>
	     <td align="right" ><%=formatNUMBER(actlnamt,2)%></td>
        </tr>  
         <tr bgcolor="#FFFFFF">
             <td>  �ٴڡ@</td>
             <td align="right" ><%=formatnumber(lbnkamt,2)%></td>
	     <td align="right" ><%=formatnumber(lsadamt,2)%></td>
	     <td align="right" ><%=formatnumber(lchamt,2)%></td>
             <td align="right" ><%=formatnumber(ajlnamt,2)%></td>	     
	     <td align="right" ><%=formatnumber(cpayamt,2)%></td>
             <td>�@</td>
	     <td>�@</td>
        </tr>  
        <tr bgcolor="#FFFFFF">
             <td>  �Q���@</td>
             <td align="right" ><%=formatnumber(ibnkamt,2)%></td>
	     <td align="right" ><%=formatnumber(isadamt,2)%></td>
	     <td align="right" ><%=formatnumber(ichamt,2)%></td>
             <td align="right" ><%=formatnumber(ajintamt,2)%></td>	     
	     <td align="right" ><%=formatnumber(cintamt,2)%></td>
             <td>�@</td>
	     <td>�@</td>
        </tr> 
       <tr bgcolor="#FFFFFF">
             <td>  �J�|�O</td>
             <%if gamt1 <> 0 then %>
             <td align="right" ><%=formatnumber(gamt1,2)%></td>
             <%else%>
             <td>�@</td>
             <%end if%>
             <%if gamt2 <> 0 then %>
	     <td align="right" ><%=formatnumber(gamt2,2)%></td>
             <%else%>
             <td>�@</td>
             <%end if%>
             <%if gamt3 <> 0 then %>
	     <td align="right" ><%=formatnumber(gamt3,2)%></td>
             <%else%>
             <td>�@</td>
             <%end if%>
             <td>�@</td>	
             <%if ttlgamt <> 0 then %>     
	     <td align="right" ><%=formatnumber(ttlgamt,2)%></td>
             <%else%>
             <td>�@</td>
             <%end if%>
             <td>�@</td>
	     <td>�@</td>
        </tr> 
      <tr bgcolor="#FFFFFF">
             <td>  ��|�O</td>
             <%if hamt1 <> 0 then %>
             <td align="right" ><%=formatnumber(hamt1,2)%></td>
             <%else%>
             <td>�@</td>
             <%end if%>
             <%if hamt2 <> 0 then %>
	     <td align="right" ><%=formatnumber(hamt2,2)%></td>
             <%else%>
             <td>�@</td>
             <%end if%>
             <%if hamt3 <> 0 then %>
	     <td align="right" ><%=formatnumber(hamt3,2)%></td>
             <%else%>
             <td>�@</td>
             <%end if%>
             <td>�@</td>	     
             <%if ttlHamt <> 0 then %>
	     <td align="right" ><%=formatnumber(ttlHamt,2)%></td>
             <%else%>
             <td>�@</td>
             <%end if%>
             <td>�@</td>
	     <td align="center">�U��/�Ѫ�</td>
        </tr> 
      <tr bgcolor="#FFFFFF">
             <td>  ���@�J</td>
             <td align="right" ><%=formatnumber(ttlbnk,2)%></td>
	     <td align="right" ><%=formatnumber(ttlsad,2)%></td>
	     <td align="right" ><%=formatnumber(ttlch,2)%></td>
             <td align="right" ><%=formatnumber(cajttlamt,2)%></td>	     
	     <td align="right" ><%=formatnumber(ttlrec,2)%></td>
 
             <td align="center">�`�B</td>
             <td align="cener" ><%=formatnumber(gttlrate,2)%>%</td>
        </tr> 
        <BR>
     <tr bgcolor="#FFFFFF">
             <td>  ��@�X</td>
             <td align="right" ><%=formatnumber(payamt,2)%></td>
	     <td>�@</td>
	     <td>�@</td>
             <td align="right" ><%=formatnumber(ajwdamt,2)%></td>	     
	     <td align="right" ><%=formatnumber(ttlpay,2)%></td>
             <td align="right" >�@</td>
	     <td align="right" >�@</td>
        </tr> 
       

</table>
 <BR>
        <BR>
<table border="" cellpadding="0" cellspacing="0">
<tr>
<td>    �U���`�Ʃ�<%=mm%>��<%=yy%>�~�e  </td> 
<td>&nbsp</td>
<td align="right"><%=formatnumber(ttllncnt,0)%></td>
</tr>
<tr>
<td>    �s�U���`��</td> 
<td>&nbsp</td>
<td align="right">+<%=formatnumber(nwlncnt,0)%></td>
</tr>
<tr>
<td>    �w�M�ƴ`���U���`�� </td>
<td align="right">-<%=formatnumber(xclncnt,0)%></td> 
<td>&nbsp</td>
</tr>
<tr>
<td>    �w�M�ƨ�L�U���`�� </td>
<td align="right">-<%=formatnumber(oclncnt,0)%></td> 
<td>&nbsp</td>
</tr>
<tr>
<td>    �w�M�ƶU�ڦX�@�`�� </td>
<td>&nbsp</td>
<td align="right">-<%=formatnumber(clncnt,0)%></td> 
</tr>

<tr>
<td>    �U���`�ƦX�@   </td> 
<td>&nbsp</td>
<td align="right"><%=formatnumber(actlncnt,0)%></td>
</tr>
<tr></tr>
<tr></tr>
<tr>
<td>   �����`�Ʃ�<%=mm%>��<%=yy%>�~�e </td> 
<td>&nbsp</td>
<td align="right"><%=formatnumber(ttlmem,0)%></td>
</tr>
<tr>
<td>    �����h���`�� </td> 
<td>&nbsp</td>
<td align="right">-<%=formatnumber(omemcnt,0)%></td>

</tr>

<tr>
<td>    �s�����`��</td> 
<td>&nbsp</td>
<td align="right">+<%=formatnumber(nmemcnt,0)%></td>
</tr>
<tr>

<td>    �����`�ƦX�@</td> 
<td>&nbsp</td>
<td align="right"><%=formatnumber(actmem,0)%></td>
</tr>

</table>
</center>
</body>
</html>

