<!-- #include file="../conn.asp" -->
<!-- #include file="cutpro.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
   response.redirect "main.asp"
   
end if
    todate  = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
    repaydate=todate
if id <>"" then
   memno = id
   pint = 0
   pamt  = 0 
   intamt = 0
   mstatus=""
end if
if request.form("bye") <> "" then
   id=""
	For Each Field in Request.Form
		TheString = Field & "= id"
		Execute(TheString)
	Next
    pint = 0
   pamt  = 0 
   intamt = 0
   cashamt = 0
   lostamt = 0
   todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
   
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
                id = memno  

                else
                    
                   msg ="�����s�����s�b "
                end if 
                rs.close
                 sql ="select * from loanrec where  repaystat='N'  and  memno="&memno 
        end if
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
                            xstatus="�۰���b(ALL)"

                       case  "0"
                            xstatus="�۰���b(�Ѫ�)"                       
                       case  "1"
                            xstatus="�۰���b(�Ѫ�,�Q��)"
                       case  "Z"
                            xstatus="�۰���b(�Ѫ�,����)"
                      case  "3"
                            xstatus="�۰���b(�Q��,����)"
                       case  "T"
                           xstatus = "�w��,�Ȧ�"
                      
                      case  "M"
                            xstatus= "�w��"
                       case "F"
                              xstatus = "�S�O�Ӯ�"  
			  
                        case "8"
                             xstatus = "�פ���y��b"
                          
                        case "9"
                             xstatus = "�פ���y���`"
                                            
                end select
        if msg="" then
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
                lnnum=rs("lnnum")
                id = memno
                rs.close 
		pint = 0
		pamt  = 0 
                yy = year(date())
                mm  =month(date())-1
              
		sql = "select lnnum,memno,ldate,bal,max(ldate)  from loan where memno='"& memno & "' and lnnum='"&lnnum&"' and code='E1' group by  lnnum,memno,ldate,bal"
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
                do while  not rs.eof 
                      pamt = pamt + rs("bal")
                rs.movenext
                loop             
                rs.close
		sql = "select * from loan where memno ='"& memno & "'  and code= 'F1' and lnnum='"&lnnum&"'   "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
                do while not  rs.eof 
                      pint = pint + rs("bal")       
                rs.movenext
                loop
                rs.close
                ttlamt = pint+pamt          
                repaydate = todate 
		chkdate =right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1)
                else
                  id = ""
                  msg = "�ɾڽs�����s�b "
                end if
                end if   
                cashamt = 0
                intamt = 0	
end if



if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
 

	if msg="" then
		conn.begintrans

                 xrepaydate = right(repaydate,4)&"/"&mid(repaydate,4,2)&"/"&left(repaydate,2)
                
                if cashamt <>"" and cashamt >0 then
                   
                   conn.execute("insert into loan (memno,lnnum,ldate,code,amount) values ('"&memno&"','"&lnnum&"','"&xrepaydate&"','E0',"&cashamt&"*-1 ) ")                  
                   conn.execute("update loanrec set bal = bal + ("&cashamt&") where lnnum ='"&lnnum&"' ")
                   conn.execute("update loanrec set repaystat = 'C' where lnnum ='"&lnnum&"' and bal = 0  ")
                end if  
                
                if intamt <>"" and intamt>0 then               
                   conn.execute("insert into loan (memno,lnnum,ldate,code,amount) values ('"&memno&"','"&lnnum&"','"&xrepaydate&"','F0',"&intamt&"*-1 ) ")
                end if                      
               
                   conn.execute("insert into share (memno,ldate,code,amount) values ('"&memno&"' ,'"&xrepaydate&"','A0',"&ttlpay&" )   ")               
                                        
		conn.committrans
		msg = "�����w��s"
	end if
        id = ""
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next

 
       
else
   mstatus=""
end if
%>
<html>
<head>
<title>�U�ڰh�ڦܪѪ��ާ@</title>
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
        if (formObj.cashamt.value!=""&&formObj.cashamt.value!=0){
            ramt = parseFloat(formObj.cashamt.value)        
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



	if (!formatDate(formObj.repaydate)){
		reqField=reqField+", �ٴڤ��";
		if (!placeFocus)
			placeFocus=formObj.memBday;
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
<form name="form1" method="post" action="repbackloan.asp">

<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="xlnnum" value="<%=xlnnum%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="chkdate" value="<%=chkdate%>">
<input type="hidden" name="xstatus" value="<%=xstatus%>">
<input type="hidden" name="mstatus" value="<%=mstatus%>">
<input type="hidden" name="ttlpay" value="<%=ttlpa%>">
<input type="hidden" name="cutdate" value="<%=cutdate%>">
<input type="hidden" name="lastmonth" value="<%=lastmonth%>">
<input type="hidden" name="lastyear" value="<%=lastyear%>">
<input type="hidden" name="spass" value="<%=spass%>">
<div><center><font size="3">�U�ڰh�ڦܪѪ��ާ@</font></center></div>
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
			<input type="button" value="���"  onclick="popup('pop_srhLoan.asp?key='+document.form1.memNo.value)" class="sbttn"  >
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
		<td class="b12" align="left">�����{��</td>
		<td width=45></td>
		<td id="xstatus" ><%=xstatus%></td> 
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">�h�ڤ��</td>
		<td width=50></td>
		<td><input type="text" name="repaydate" value="<%=repaydate%>" size="10" onblur="if(!formatDate(this)){this.value=''};form1.repaydate.value=this.value">
	</tr>


	<tr>
               <td width=30></td>
		<td class="b12" align="left">�h�٥���</td>
		<td width=50></td>
		<td><input type="text" name="cashamt" value="<%=cashamt%>" size="10" onblur="if(!formatNum(this)){this.value=''};calculation();"></td>
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">�h�٧Q��</td>
		<td width=50></td>
		<td><input type="text" name="intamt" value="<%=intamt%>" size="10"  onblur="if(!formatNum(this)){this.value=''};calculation();"></td>
	</tr>
 	<tr>
               <td width=30></td>
		<td class="b12" align="left">�h�٦X�@���B</td>
		<td width=50></td>
		<td id="ttlpay"></td>
		
	</tr>   


	<tr>
		<td width=30></td>
		<td colspan="3" align="right">
		<% if id <> "" then %>
               
		<%if session("userLevel")<>2 and session("userLevel")<>1 and session("userLevel")<>4 then%>
		<input type="submit" value="�x�s" onclick="return validating()&&confirm('�T�w�x�s?')" name="action" class="sbttn">
		<input type="button" value="�d�߭ӤH�b" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.chkdate.value )" class="sbttn">
		<%end if%>			
		<% end if %>
               
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



        </table>
        </td>  
   </tr>
</table>       
</center>
</form>
</body>
</html>
