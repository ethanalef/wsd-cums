<!-- #include file="../conn.asp" -->
<!-- #include file="../addUserLog.asp" -->
<!-- #include file="cutpro.asp" -->

<%
if request.form("back") <> "" then
   response.redirect "main.asp"
   
end if
  todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())


if request.form("bye") <> "" then
   id=""
	For Each Field in Request.Form
		TheString = Field & "= id"
		Execute(TheString)
	Next
 
  sdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
end if


if request.form("Search") <> "" then 
 
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
    msg=""
    
   if memno <>"" then          
        set rs = conn.execute("select memno,memName,memcName,mstatus from memmaster where memno='"&memno&"' and mstatus='B' ")

        if not rs.eof then

           id = memno           
           memname = rs("memname")
           memcname = rs("memcname")
           mstatus = rs("mstatus")
        else
 
          msg = "���O�}������"
       end if
        rs.close
   end if 
   
   if msg=""  then
    yy = right(sdate,4)
    mm = mid(sdate,4,2)
    dd = left(sdate,2)
    xxdate=dateserial(yy,mm,dd)
    set rs=conn.execute("select * from loanrec  where memno='"&memno&"' and repaystat='N'   ")
     if not rs.eof then
        lnamt = rs("bal")
     
     end if     
     rs.close

     end if           
   
  end if   

else


if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
      
        sxdate=dateserial(right(sdate,4),mid(sdate,4,2),left(sdate,2))
      
      
       conn.begintrans
       conn.execute("insert into share  (memno,ldate ,code,amount ) values ('"&memno&"','"&sxdate&"','MF',"&cfee&" ) ")                                
       conn.committrans 
	msg = "�����w��s"
        id = ""

       
else
     sdate=RIGHT("0"&day(date()),2)&"/"&RIGHT("0"&month(date()),2)&"/"&year(date())          
     chkdate=RIGHT("0"&day(date()),2)&"/"&RIGHT("0"&month(date()),2)&"/"&(year(date())-1)


end if
end if
%>
<html>
<head>
<title>�����U�ګإ�</title>
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

 
         return true;
  

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

	if (formObj.cfee.value==""){
		reqField=reqField+", �N��O���B";
		if (!placeFocus)
			placeFocus=formObj.cfee;
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
<form name="form1" method="post" action="coldfee.asp">

<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="xlnnum" value="<%=xlnnum%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="xstatus" value="<%=xstatus%>">
<input type="hidden" name="mstatus" value="<%=mstatus%>">
<input type="hidden" name="chkdate" value="<%=chkdate%>">
<input type="hidden" name="cutdate" value="<%=cutdate%>">
<input type="hidden" name="lastmonth" value="<%=lastmonth%>">
<input type="hidden" name="lastyear" value="<%=lastyear%>">
<input type="hidden" name="spass" value="<%=spass%>">
<div><center><font size="3">�����U�ګإ�</font></center></div>
<center>
<table border="0" cellspacing="0" cellpadding="0">
				<tr height="22">
					<td class="b8" align="right">�����s��</td>
					<td width=10></td>
					<td><input type="text" name="memNo" value="<%=memNo%>" size="10" maxlength="10" "<%if id<>"" then response.write " onfocus=""form1.refyr.focus();""" end if%>>
                                         <input type="submit" value="�j�M" name="Search" class ="Sbttn">     
                                        </td>
				</tr>
		

				<tr height="22">
					<td class="b8" align="right">�m�W</td>
					<td width=10></td>
					<td id="memname"><%=memname%></td>
                                
				</tr>
				<tr height="22">
					<td class="b8" align="right"></td>
					<td width=10></td>
					
                                        <td id="memcname"><%=memcname%></td>
				</tr>
				<tr height="22">
					<td class="b8" align="right">���</td>
					<td width=10></td>
                                                                          
					<td><input type="text" name="sdate"  value="<%=sdate%>" size="10" maxlength="10"  onblur="if(!formatDate(this)){this.value=''};"></td>
                                        
				</tr>
                                <tr height="22">
                                <td class="b8" align="right">�������B</td>
                               
                                <td width="10"></td>
    		                <td><input type="text" name="lnamt"  value="<%=lnamt%>" size="10" maxlength="10" ></td>   
                                </tr>  

 
		<tr>
					<td colspan="3" align="right">
					<%if id<>"" then %>						
					<input type="submit" value="�x�s" onclick="return validating()&&confirm('�T�w�x�s?')" name="action" class="sbttn">
					<input type="button" value="�d�߭ӤH�b" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.chkdate.value  )" class="sbttn">
                                        <%end if%>
					<input type="submit" value="����" name="bye" class="sbttn">
					<input type="submit" value="��^" name="back" class="sbttn">
				</td>
				</tr>

</table>
</center>
</form>
</body>
</html>
