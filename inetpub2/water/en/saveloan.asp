<!-- #include file="../conn.asp" -->
<!-- #include file="cutpro.asp" -->
<!-- #include file="../addUserLog.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
if request.form("back") <> "" then
   response.redirect "main.asp"
   
end if
id = request("id")
if id <>"" then
   memno = id
end if
if request.form("bye") <> "" then
   id=""
	For Each Field in Request.Form
		TheString = Field & "= id"
		Execute(TheString)
	Next
     ttlpay = 0 
      pint = 0
      pamt = 0
	 repayamt = 0
         intamt = 0
	 saveamt = 0   
end if



if request.form("Search")<>"" or id <>""  then
        msg=""
      
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
        if id<>"" then 
		sql = "select * from loanrec where lnnum='"& xlnnum & "' "
        else
                set rs=conn.execute("select memno,memname,memcname,mstatus from memmaster where memno='"&memno&"' ")
                if not rs.eof then
                   For Each Field in rs.fields 
		   TheString = Field.name & "= rs(""" & Field.name & """)"
	           Execute(TheString)
		   Next
                rs.close
                id = memno  
                select case mstatus
                       case "L"
                           xstatus= "�b�b"
                       case  "D"
                           xstatus="�N��"
                       
                       case  "V"
                           xstatus= " IVA "
                         
                       case  "C"
                             xstatus= "�h��"
             
                       case  "P"
                             xstatus= "�h�@"
                         
                       case  "B"
                            xstatus="�}��"
                    
                       case  "N"
                            xstatus= "���`"
                        
                      case  "J"
                            xstatus= "�s��"
                       
                      case "H"
                          xstatus= "�Ȱ��Ȧ�"
                      
                       case  "A"
                            xstatus="�۰���b"

                       case  "0"
                            xstatus="�۰���b(�Ѫ�)"                       
                       case  "1"
                            xstatus="�۰���b(�Ѫ�,�Q��)"
                       case  "Z"
                            xstatus="�۰���b(�Ѫ�,����)"
                       case "3"
                             xstatus="�۰���b(�Q��,����)"
                       case  "M"
                           xstatus = "�w��,�Ȧ�"
                      
                      case  "T"
                            xstatus= "�w��"
                     case "F"
                          xstatus =  "���D�U��"
                end select
                         
              
                xlnnum =""
                set rs = conn.execute("select lnnum from loanrec where repaystat='N' and memno='"&memno&"' ")
                if not rs.eof then
                   xlnnum = rs(0)
                end if
                rs.close    
                saveamt = 0
		sql = "select amount,CODE  from share where memno='"&memno&"' "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
                do while not  rs.eof 
                   select case rs("code")
                          case "0A" ,"A1","A2","A3","C0","C1","C3","A0","A7" ,"A4" 
                               saveamt = saveamt + rs(0)
                          case "B0","B1","B3","G0","G1","G3","H0","H1","H3"
                               saveamt = saveamt - rs(0)
                    end select 
                rs.movenext
                loop
                rs.close  
   
                else
                    
                   msg ="�ɾڽs�����s�b "
                end if 
             
                 sql ="select * from loanrec where repaystat='N' and memno="&memno
        end if
      
        if msg="" then
        
		sql = "select * from loanrec where lnnum='"& xlnnum & "' "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if not rs.eof then

			For Each Field in rs.fields
			if Field.name="lndate"  then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
		
                repaydate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2) &"/"&year(date())
                id = memno
                rs.close 
                pint = 0
                pamt = 0
		sql = "select *  from loan where memno='"& memno & "'  and code='DE' and pflag= 1 "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
                do while  not rs.eof 
                      pamt = pamt + rs("bal")
                rs.movenext
                loop             
                rs.close
		sql = "select * from loan where memno ='"& memno & "'  and code= 'MF' and pflag = 0   "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
                do while not  rs.eof 
                      pint = pint + rs("bal")       
                rs.movenext
                loop
                rs.close


               end if
                else
                  msg = "�ɾڽs�����s�b "

                end if   	
end if



if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
        msg =""
	if msg="" then
		conn.begintrans
                xrepaydate = dateserial(right(repaydate,4),mid(repaydate,4,2),left(repaydate,2))

               yy = right(repaydate,4)
               mm = mid(repaydate,4,2)
               dd = left(repaydate,2)
               if yy/4=int(yy/4) and yy/100=int(yy/100) then
                  md=mid("3129303130313130313031",(mm-1)*2+1,2)
               else
                  md=mid("3128303130313130313031",(mm-1)*2+1,2)
               end if                 
               if ttlpay > 0 then
                  conn.execute("insert into share (memno,ldate,code,amount) values ('"&memno&"','"&xrepaydate&"','B0',"&ttlpay&" ) ")
                  adduserlog "���� : "&memno&" �Ѫ��ٴڪ��B $"&formatnumber(ttlpay,2)&" ��� "&repaydate
               end if
              if pamt > 0 then
                xx = pamt
 		sql = "select * from loan where lnnum='"& lnnum & "' and pflag=1 and code='DE'  "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn,1,1
                do while not rs.eof
                   xx = xx - rs("bal")
                   if xx = 0 then
                      rs.movelast
                   end if 
                   rs.movenext
                loop
                rs.close  
                  
                end if 
                if repayamt > 0 then
                   conn.execute("insert into loan (memno,lnnum,ldate,code,amount) values ('"&memno&"','"&lnnum&"','"&xrepaydate&"','E0',"&repayamt&" ) ")                                              
                  adduserlog "���� : "&memno&" �Ѫ��ٴ�  �������B $"&formatnumber(repayamt,2)&" ��� "&repaydate
                        
                conn.execute("update loanrec set bal = bal - "&repayamt&" where lnnum='"&lnnum&"' ")
                conn.execute("update loanrec set cleardate='"&xrepaydate&"' where lnnum='"&lnnum&"' and bal= 0 ")
                conn.execute("update loanrec set repaystat ='C' where lnnum='"&lnnum&"' and bal= 0 ") 
                set ms = conn.execute("select * from loanrec where memno='"&memno&"' and repaystat='N' ")
                if  ms.eof  then
                    if (saveamt - ttlpay )=0 then
                       conn.execute("update memmaster set wdate = '"&xrepaydate&"'  where memno='"&memno&"' ")
                    end if
                end if
                ms.close
               end if
               if pint > 0 then
                xx = pint
 		sql = "select * from loan where lnnum='"& lnnum & "' and pflag=1 and code='MF'  "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn,1,1
                do while not rs.eof
                   xx = xx - rs("bal")
                   if xx = 0 then
                      rs.movelast
                   end if 
                   rs.movenext
                loop
                rs.close  
                end if 
                if intamt > 0 then
                conn.execute("insert into loan (memno,lnnum,ldate,code,amount) values ('"&memno&"','"&lnnum&"','"&xrepaydate&"','F0',"&intamt&" ) ")                                              
               
                adduserlog "���� : "&memno&" �Ѫ��ٴ�  �Q�����B $"&formatnumber(intamt,2)&" ��� "&repaydate
                 end if
 

                id = ""
                
		conn.committrans
		msg = "�����w��s"
	end if
                id = ""
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next
      pint = 0
      pamt = 0
	 repayamt = 0
         intamt = 0
	 saveamt = 0    	
        ttlpay = 0
else
    chkdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1)           
    repaydate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
    ttlpay = 0 
      pint = 0
      pamt = 0
	 repayamt = 0
         intamt = 0
	 
end if

%>
<html>
<head>
<title>�Ѫ��ٴ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">

<script language="JavaScript">
<!--

function popup(filename){
  window.open (filename,'pop','width=500,height=550,statusbar=no,toolbar=no,resizable,scrollbars,dependent')
}

function formatNum(numform){
  if (isNaN(numform.value)||numform.value<0)
    return false;
  else
    return true;
}

function valDate(M, D, Y){
  Months= new Array(31,28,31,30,31,30,31,31,30,31,30,31);
  Leap  = false;
  if((Y % 4 == 0) && ((Y % 100 != 0) || (Y %400 == 0)))
    Leap = true;
  if((D < 1) || (D > 31) || (M < 1) || (M > 12) || (Y < 0))
    return false;
  if((D > Months[M-1]) && !((M == 2) && (D > 28)))
    return false;
  if(!(Leap) && (M == 2) && (D > 28))
    return false;
  if((Leap)  && (M == 2) && (D > 29))
    return false;
  return true;
};

function formatDate(dateform){

	    formObj=document.form1;    
            sMn = parseInt(formObj.lastmonth.value)
            sYr = parseInt(formObj.lastyear.value)
            spass   = parseInt(formObj.spass.value)

  cDate = dateform.value;
  dSize = cDate.length;

  var strValidChars = "0123456789/";
  var strChar = "";

 
   for (i = 0; i < dSize ; i++)
      {  
      strChar = cDate.substr(i,1);
      if ( strValidChars.indexOf(strChar) == -1)
         { 
         return false ;   
          }
      }

  if (dSize!=0){
    sCount= 0;

    for(var i=0; i < dSize; i++)
      (cDate.substr(i,1) == "/") ? sCount++ : sCount;
     
    if (sCount == 2){
		ySize = cDate.substring(cDate.lastIndexOf("/")+1,dSize).length;
		if (ySize<2 || ySize>4 || ySize == 3){
		  return false;
		 }
		idxBarI = cDate.indexOf("/");
		idxBarII = cDate.lastIndexOf("/");
		strD = cDate.substring(0,idxBarI);
		strM = cDate.substring(idxBarI+1,idxBarII);
		strY = cDate.substring(idxBarII+1,dSize);
		strM = (strM.length < 2 ? '0'+strM : strM);
		strD = (strD.length < 2 ? '0'+strD : strD);
		if(strY.length == 2)
		  strY = (strY > 50  ? '19'+strY : '20'+strY);
    }else{
    	if (dSize != 8)
			return false;
		strD = cDate.substring(0,2);
		strM = cDate.substring(2,4);
		strY = cDate.substring(4,8);
    }
    dateform.value = strD+'/'+strM+'/'+strY;
    if (!valDate(strM, strD, strY))
      return false;
    else

     Mn = parseInt(strM)
      Yr = parseInt(strY)
      if (((Mn<=sMn)&&(Yr==sYr))||(Yr<sYr)){
         return false ;
      }else{      
         return true;
      }

  }
}



function calculation(){
	formObj=document.form1;
        if (formObj.repayamt.value!=""&&formObj.repayamt.value!=0){
            ramt = parseFloat(formObj.repayamt.value)        
        }else{  
            ramt = 0
        }
        if  (formObj.intamt.value!=""&&formObj.intamt.value!=0){
            iamt = parseFloat(formObj.intamt.value)
        }else{
            iamt = 0
        }
      
        tamt = ramt + iamt
                   
        document.form1.ttlpay.value=tamt 
        document.all.tags( "td" )['ttlpay'].innerHTML=tamt;
       
        
} 


function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.memNo.value==""){
		reqField=reqField+", �����s��";
		if (!placeFocus)
			placeFocus=formObj.memNo;
	}


    if (reqField){
        if (reqField.lastIndexOf(",")==0)
	        reqField = "�ж�J"+reqField.substring(2);
        else
	        reqField = "�ж�J"+reqField.substring(2,reqField.lastIndexOf(","))+'��'+reqField.substring(reqField.lastIndexOf(",")+2);
        alert(reqField);
        placeFocus.focus();
        return false;
    }else{
        return true;
    }
}
//  -->
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.memNo.focus()">
<DIV>
<!-- #include file="menu.asp" -->


<%if msg<>"" then %>
<div><center><font size="3"><%=msg%></font></center></div>
<% end if%>

<br>
<form name="form1" method="post" action="saveloan.asp">

<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="xlnnum" value="<%=xlnnum%>">
<input type="hidden" name="mstatus" value="<%=mstatus%>">
<input type="hidden" name="saveamt" value="<%=saveamt%>">
<input type="hidden" name="chkdate" value="<%=chkdate%>">
<input type="hidden" name="ttlpay" value="<%=ttlpay%>">
<input type="hidden" name="cutdate" value="<%=cutdate%>">
<input type="hidden" name="lastmonth" value="<%=lastmonth%>">
<input type="hidden" name="lastyear" value="<%=lastyear%>">
<input type="hidden" name="spass" value="<%=spass%>">
<div><center><font size="3">�Ѫ��ٴ�</font></center></div>
<center>
<table border="0" cellspacing="0" cellpadding="0">
       <tr>
		<td width="500" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
			<tr>
               		<td width=30></td>
			<td class="b12" align="left">�������X</td>
			<td width=50></td>
			<td><input type="text" name="memNo" value="<%=memNo%>" size="10" <%if id<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>>
			<%if id = "" then %>
			<input type="button" value="���"  onclick="popup('pop_srhloan.asp?key='+document.form1.memNo.value)" class="sbttn"  >
			<input type="submit" value="�j�M" name="Search" class ="Sbttn">
			<% end if %>
                        </TD>
		
			</tr>

			<tr>
          		<td width=30></td>
			<td class="b12" align="left">�����W��</td>
			<td width=50></td>
			<td><input type="text" name="memName" value="<%=memName%>" size="30"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
			</tr>
			<tr>
          		<td width=30></td>
			<td class="b12" align="left"></td>
			<td width=50></td>
			<td><input type="text" name="memcName" value="<%=memcName%>" size="30"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
			</tr>

			<tr>
               		<td width=30></td>
			<td class="b12" align="left">�Ѫ����s</td>
			<td width=50></td>
			<td id ="saveamt"><%=formatnumber(saveamt,2)%></td> 
			</tr>

			<tr>
               		<td width=30></td>
			<td class="b12" align="left">�ٴڤ��</td>
			<td width=50></td>
			<td><input type="text" name="repaydate" value="<%=repaydate%>" size="10" onblur="if(!formatDate(this)){this.value=''};form1.repaydate.value=this.value">
			</tr>

			<tr>
               		<td width=30></td>
			<td class="b12" align="left">�ٴڥ���</td>
			<td width=50></td>
			<td><input type="text" name="repayamt" value="<%=repayamt%>" size="10" onblur="if(!formatNum(this)){this.value=''};calculation();"></td>
			</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">�ٴڧQ��</td>
		<td width=50></td>
		<td><input type="text" name="intamt" value="<%=intamt%>" size="10" onblur="if(!formatNum(this)){this.value=''};calculation();"></td>
		
	</tr>
 	<tr>
               <td width=30></td>
		<td class="b12" align="left">�ٴڦX�@���B</td>
		<td width=50></td>
		<td id="ttlpay"></td>
		
	</tr>   
	<tr>
		<td width=30></td>
		<td colspan="3" align="right">
		<% if id <> "" then %>
                <%if xlnnum <> "" then %>
			<%if session("userLevel")<>2 and session("userLevel")<>1 and session("userLevel")<>4 then%>
			<input type="submit" value="�x�s" onclick="return validating()&&confirm('�T�w�x�s?')" name="action" class="sbttn">
                        <input type="button" value="�d�߭ӤH�b" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.chkdate.value )" class="sbttn">					
			<%end if%>
		<% end if %>
                <%end if %>  
			<input type="submit" value="����" name="bye" class="sbttn">
			<input type="submit" value="��^" name="back" class="sbttn">
		</td>
	</tr>

</table>
</td>
	<td width="400" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
			<tr>
                	<td width=30></td>
			<td class="b12" align="left">�ɾڽs��</td>
			<td width=50></td>
			<td><input type="text" name="lnnum" value="<%=lnnum%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
			</tr>

			<tr>
 		        <td width=30></td>
			<td class="b12" align="left">�������</td>
			<td width=50></td>
			<td><input type="text" name="lndate" value="<%=lndate%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
			</tr>

			<tr>
 	                <td width=30></td>
			<td class="b12" align="left">�ɾڪ��B</td>
			<td width=50></td>
			<td><input type="text" name="appamt" value="<%=appamt%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
			</tr>

			<tr>
        	       <td width=30></td>
			<td class="b12" align="left">�ɾڵ��l</td>
			<td width=50></td>
			<td><input type="text" name="bal" value="<%=bal%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
			</tr>

			<tr>
            		<td width=30></td>
			<td class="b12" align="left">�������</td>
			<td width=50></td>
			<td><input type="text" name="pamt" value="<%=pamt%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
			</tr>
			<tr>
            		<td width=30></td>
			<td class="b12" align="left">����Q��</td>
			<td width=50></td>
			<td><input type="text" name="pint" value="<%=pint%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.repaydate.focus();""" end if%>></td> 
			</tr>
</table>
</td>
</tr>
</table>
</CENTER>
</form>
</body>
</html>
