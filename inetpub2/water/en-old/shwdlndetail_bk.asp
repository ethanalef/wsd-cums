<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "shwdloan.asp"'
end if

xlnnum = request("lnnum")

if request.form("action") <> "" then
        addloan = 0
	For Each Field in Request.Form
 	 
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)

	
	Next
  	set rs = server.createobject("ADODB.Recordset")
	msg = ""

	if cleardate="" then
           	msg = "�M�Ƥ������J "
        else
        if cashamt  = "" then
        
		msg = "�M�ƪ��B����J "
        end if
        end if




	if msg="" then
		conn.begintrans                
                xcleardate = right(cleardate,4)&"/"&mid(cleardate,4,2)&"/"&left(cleardate,2)

                  rpayamt = cashamt * -1
                  conn.execute("update loanrec set bal = bal - "&cashamt&"  where lnnum='"&lnnum&"' ")
                  conn.execute("update loanrec set cleardate='"&xcleardte&"'  where  lnnum='"&lnnum&"' and bal=0  ")
                  conn.execute("update loanrec set repaystat='C'  where  lnnum='"&lnnum&"' and bal=0  ")                                                     
		  conn.execute("insert into share (memno,code,ldate,amount) values ('"&memno&"','B0','"&xcleardate&"',"&ttlpamt&" )  ")	
		  conn.execute("insert into loan (memno,lnnum,code,ldate,amount) values ('"&memno&"','"&lnnum&"','E0','"&xcleardate&"',"&cashamt&")  ")   

                  if  pamt > 0  then  
                       conn.execute("update loan set pdate='"&xcleardate&"' where lnnum='"&lnnum&"' and pflag=0 and code='DE'")  
                      conn.execute("update loan set pflag=0  where lnnum='"&lnnum&"' and pflag=1 and code='DE'")             

                  end if      
 
                  rintamt = cashint *-1    
                  if cashint > 0 then
                  conn.execute("insert into loan (memno,lnnum,code,ldate,amount) values ('"&memno&"','"&lnnum&"','F0','"&xcleardate&"',"&cashint&" )  ")  
                  if  pint > 0 then
                        conn.execute("update loan set  pdate='"&xcleardate&"'  where lnnum='"&lnnum&"' and pflag=0 and code='DF' ")
                       conn.execute("update loan set pflag=0   where lnnum='"&lnnum&"' and pflag=1 and code='DF' ")
  
                  end if 
                 end if       
		conn.committrans
		msg = "�����w��s"
             end if
           
	     response.redirect "shwdloan.asp"'
else
	if xlnnum <> "" then
		sql = "select * from loanrec where lnnum= " & xlnnum
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if rs.eof then
			response.redirect "shwdloan.asp"
		else
			For Each Field in rs.fields
			if Field.name="lndate"  then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
		end if
                rs.close 
               id = memno
              
		pint = 0
		pamt  = 0 
              
              
		sql = "select *  from loan where memno='"& memno & "'  and code='DE' and pflag= 1 "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
                do while  not rs.eof 
                      pamt = pamt + rs("bal")
                rs.movenext
                loop             
                rs.close
		sql = "select * from loan where memno ='"& memno & "'  and code= 'DF' and pflag = 1   "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
                do while not  rs.eof 
                      pint = pint + rs("bal")       
                rs.movenext
                loop
                rs.close
	 end if
	cleardate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
	cchkdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1)
        todate = cleardate 
end if
%>
<html>
<head>
<title>�Ѫ��M�ƫإ�</title>
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

function formatyn(numform){
  if (numform.value!='Y'&&numform.value!='N')
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
      return true;
  }
}

function calculation(){
	formObj=document.form1;

 
	if (formObj.cleardate.value!=""){
           
            if (formObj.pint.value!=""){
               ppint1 = parseFloat(formObj.pint.value)
            }else{
               ppint1 = 0
            }        

	   lnbal  = parseFloat(formObj.bal.value)  
  

           ssdate = formObj.cleardate.value          
           ttdate = formObj.todate.value 

	   Months= new Array(31,28,31,30,31,30,31,31,30,31,30,31)

	   Y=ssdate.substr(6,4)           
           M=ssdate.substr(3,2)
	   D=ssdate.substr(0,2)
            
           
           mD = Months[M -1] 
           ppint2 = 0.00

           if (m = 2){
              if((Y % 4 == 0) && ((Y % 100 != 0) || (Y %400 == 0))){
                mD = 29              
              }
           }   
   
             
             ppint2 = Math.round(lnbal*0.01*D/mD*100)/100
 
             
             ttlint =ppint1 + ppint2
        

              document.form1.cashamt.value = lnbal
              document.form1.cashint.value = ttlint
              document.all.tags( "td" )['cashint'].innerHTML= ttlint;        
              document.form1.ttlpamt.value = lnbal+ttlint
              document.all.tags( "td" )['ttlpamt'].innerHTML= lnbal+ttlint;
	}

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

	if (formObj.reqamt.value==""){
		reqField=reqField+", �U�ڪ��B";
		if (!placeFocus)
			placeFocus=formObj.reqamt;
	}

	if (formObj.appamt.value==""){
		reqField=reqField+", ��֪��B";
		if (!placeFocus)
			placeFocus=formObj.appamt;
	}

	if (formObj.cleardate.value==""){
		reqField=reqField+", ���M�Ƥ��";
		if (!placeFocus)
			placeFocus=formObj.cleardate;
	}

	if (formObj.newlnnum.value==""){
		reqField=reqField+", ���s�U�ڸ��X";
		if (!placeFocus)
			placeFocus=formObj.newlnnum;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.lnnum.focus()">
<DIV>
<!-- #include file="menu.asp" -->

<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>

<br>
<form name="form1" method="post" action="shwdlndetail.asp">

<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="cchkdate" value="<%=cchkdate%>">
<input type="hidden" name="cashint" value="<%=cashint%>">
<input type="hidden" name="ttlpamt" value="<%=ttlpamt%>">
<div><center><font size="3">�Ѫ��M�ƫإ�</font></center></div>

<center>

		<td width="700" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
	<tr>
               <td width=30></td>
		<td class="b12" align="left">�U�ڸ��X</td>
		<td width=50></td>
		<td><input type="text" name="lnnum" value="<%=lnnum%>" size="10" maxlength="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td>
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">�������X</td>
		<td width=50></td>
		<td><input type="text" name="memNo" value="<%=memNo%>" size="10" maxlength="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
		
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left"><input type="hidden" name="xstatus" value="<%=xstatus%>">���</td>
		<td width=50></td>
		<td><input type="text" name="lndate" value="<%=lndate%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">�U�ڪ��B</td>
		<td width=50></td>
		<td><input type="text" name="appamt" value="<%=appamt%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">�u������</td>
		<td width=50></td>
		<td><input type="text" name="install" value="<%=install%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">�C���ٴ�</td>
		<td width=50></td>
		<td><input type="text" name="monthrepay" value="<%=monthrepay%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">�������</td>
		<td width=50></td>
		<td><input type="text" name="pamt" value="<%=pamt%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">����Q��</td>
		<td width=50></td>
		<td><input type="text" name="pint" value="<%=pint%>" size="10"></td> 
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">�������l</td>
		<td width=50></td>
		<td><input type="text" name="bal" value="<%=bal%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">�M�Ƥ��</td>
		<td width=50></td>
		<td><input type="text" name="cleardate" value="<%=cleardate%>" size="10"onblur="if(!formatDate(this)){this.value=''};calculation();"></td>
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">�M�ƥ������l</td>
		<td width=24></td>
		<td><input type="text" name="cashamt" value="<%=cashamt%>" size="10" <%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">�M�ƧQ��</td>
		<td width=50></td>
		<td id="cashint" ><%=formatnumber(cashint,2)%></td> 
	</tr>
    
	<tr>
               <td width=30></td>
		<td class="b12" align="left">�M�ƪ��B</td>
		<td width=50></td>
		<td id="ttlpamt"><%=formatNumber(ttlpamt,2)%></td>
	</tr>
	<tr>
		<td width=30></td>
		<td colspan="3" align="right">
			<%if session("userLevel")<>5 and session("userLevel")<>1 and session("userLevel")<>6 then%>
			<input type="submit" value="�x�s" onclick="return validating()&&confirm('�T�w�x�s?')" name="action" class="sbttn">
			<%end if%>
		        <input type="button" value="�d�߶U��" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.cchkdate.value )" class="sbttn">			
			<input type="submit" value="��^" name="back" class="sbttn">
		</td>
	</tr>
</table>
</center>

</form>
</body>
</html>
