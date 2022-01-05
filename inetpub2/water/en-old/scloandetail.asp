<!-- #include file="../conn.asp" -->
<!-- #include file="cutpro.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "scloan.asp"'
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
           	msg = "清數日期未填入 "
        else
        if cashamt  = "" then
        
		msg = "清數金額未填入 "
        end if
        end if




	if msg="" then
		conn.begintrans                

                 xcleardate = right(cleardate,4)&"/"&mid(cleardate,4,2)&"/"&left(cleardate,2)
                
                  if cashamt > 0 then        
                  conn.execute("update loanrec set bal = bal - "&cashamt&"    where lnnum='"&lnnum&"' ")		 	
                  conn.execute("update loanrec set  repaystat='C' where lnnum='"&lnnum&"' ")		 		
                  conn.execute("update loanrec set cleardate='"&xcleardate&"' where lnnum='"&lnnum&"' ")
		  conn.execute("insert into loan (memno,lnnum,code,ldate,amount) values ('"&memno&"','"&lnnum&"','E3','"&xcleardate&"',"&cashamt&")  ")   
                  
                  if  pamt <> "" then
                       conn.execute("update loan set pdate ='"&xcleardate&"'  where lnnum='"&lnnum&"' and pflag=1 and code='DE' ")
                       conn.execute("update loan set pflag = 0   where lnnum='"&lnnum&"' and pflag=1 and code='DE' ")

                  end if 
                  end if
               
                                                 
                  
                  
                  if  pint > 0 then
                       conn.execute("update loan set pdate ='"&xcleardate&"'  where lnnum='"&lnnum&"' and pflag=1 and code='DF' ")
                       conn.execute("update loan set pflag = 0   where lnnum='"&lnnum&"' and pflag=1 and code='DF' ")

                  end if 
                   
        
		conn.committrans
		msg = "紀錄已更新"
             end if
             
	     response.redirect("ccloan.asp")
else
	if xlnnum <> "" then
		sql = "select * from loanrec where lnnum= " & xlnnum
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if rs.eof then
			response.redirect "ccloan.asp"
		else
			For Each Field in rs.fields
			
					TheString = Field.name & "= rs(""" & Field.name & """)"
			
				Execute(TheString)
			Next
                        yy = year(rs("lndate"))
                        mm = month(rs("lndate"))
                        dd = day(rs("lndate"))
                        xday = cint(mid("312831303130313130313031",(mm-1)*2+1,2)) 
		end if

                rs.close
                id = memno 
                 
                pint =0	
                 minterest  = 0
                pass = 0
                set rs = conn.execute("select * from loan where lnnum = '"&lnnum&"' ")
                if not rs.eof then
                do while not rs.eof  and pass = 0
                   select case rs("code")
                          case "0D","E1","E2","E3","DE"
                               pass = 1
                              
                   end select
                   rs.movenext
                 loop
                 end if
                               
                                
                          

                 
                if bal = appamt and pass = 0 then
                   minterest = bal * .01 *(xday - dd+1)/xday
    
                end if


		sql = "select * from loan where memno ='"& memno & "'  and  pflag = 1   "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
                do while not  rs.eof 
                    select case rs("code")
                           case  "ME"
                                 pamt = pamt = rs("amount")
                           case "MF"
                                 pint = pint + rs("bal")       
                    end select 
                rs.movenext
                loop
                rs.close                                  
               
                   
	 end if
         
	cleardate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
	cchkdate  = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1)
        todate    = cleardate

end if
%>
<html>
<head>
<title>現金清數(本金)建立</title>
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

 
	if (formObj.cleardate.value!=""){

            mint = Math.round(parseFloat(formObj.minterest.value)*100)/100           
            if (formObj.pint.value!=""){
               ppint1 = parseFloat(formObj.pint.value)
            
            }else{
               ppint1 = 0
           
            }      
	    
	   lnbal  = parseFloat(formObj.bal.value) 
           appamt  = parseFloat(formObj.appamt.value) 
           ssdate = formObj.cleardate.value
           lldate = formObj.lndate.value          
           ttdate = formObj.todate.value 
           
	   Months= new Array(31,28,31,30,31,30,31,31,30,31,30,31)

	   Y= ssdate.substr(6,4)
           M= ssdate.substr(3,2)
	   D= ssdate.substr(0,2)
           mD = Months[M -1] 

	   XY=  lldate.substr(6,4)
           XM=  lldate.substr(3,2)
	   XD=  lldate.substr(0,2)
           XmD = Months[M -1] 
           ppint2 = 0.00
           
            
 

           if (M == 2){

              if((Y % 4 == 0) && ((Y % 100 != 0) || (Y %400 == 0)))
              {
                 mD = 29              
              }
              }   
             DD = 0 
             
             if (( Y == XY ) && ( M == XM)){
                mint = 0
                DD =  D -  XD + 1
                 ppint2 = Math.round(lnbal*0.01*DD/mD*100)/100
             }else{
                if ( ppint1 !=0 ){
                   mint = 0
                }  
                if ( appamt == lnbal) {
                     ppint2 = Math.round(lnbal*0.01*(XmD-XD+1)/XmD*100)/100 + Math.round(lnbal*0.01*D/mD*100)/100     
                 }else{
                ppint2 = Math.round(lnbal*0.01*D/mD*100)/100
                }    
             }
 
       
             
             ttlint =Math.round((ppint1 + ppint2  +mint)*100)/100
        

              document.form1.cashamt.value = lnbal
              document.form1.cashint.value =0
             
              document.form1.ttlpamt.value = lnbal
             
           
	}

}

function caladd(){
	formObj=document.form1;
	cashint   = parseFloat(formObj.cashint.value) 
        bal        = parseFloat(formObj.bal.value) 
        tllamt   = cashint + bal
         document.form1.ttlpamt.value =tllamt 


}


function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.memNo.value==""){
		reqField=reqField+", 社員編號";
		if (!placeFocus)
			placeFocus=formObj.memNo;
	}



	if (formObj.appamt.value==""){
		reqField=reqField+", 批核金額";
		if (!placeFocus)
			placeFocus=formObj.appamt;
	}

	if (!formatDate(formObj.cleardate)){
		reqField=reqField+", ’清數日期";
		if (!placeFocus)
			placeFocus=formObj.cleardate;
	}

	if (formObj.cashint.value==""){
		reqField=reqField+", ’清數利息";
		if (!placeFocus)
			placeFocus=formObj.cashint;
	}
	if (formObj.cashamt.value==""){
		reqField=reqField+", ’清數本金結餘";
		if (!placeFocus)
			placeFocus=formObj.cashamt;
	}
    if (reqField){
        if (reqField.lastIndexOf(",")==0)
	        reqField = "請填入"+reqField.substring(2);
        else
	        reqField = "請填入"+reqField.substring(2,reqField.lastIndexOf(","))+'及'+reqField.substring(reqField.lastIndexOf(",")+2);
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.cleardate.focus()">
<DIV>

<!-- #include file="menu.asp" -->

<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>

<br>
<form name="form1" method="post" action="scloanDetail.asp">

<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="cchkdate" value="<%=cchkdate%>">

<input type="hidden" name="pamt" value="<%=pamt%>">
<input type="hidden" name="minterest" value="<%=minterest%>">
<input type="hidden" name="cutdate" value="<%=cutdate%>">
<input type="hidden" name="lastmonth" value="<%=lastmonth%>">
<input type="hidden" name="lastyear" value="<%=lastyear%>">
<input type="hidden" name="spass" value="<%=spass%>">
<div><center><font size="3">現金清數(本金)建立</font></center></div>
<center>

	<td width="700" valign="top">
	 	<table border="0" cellspacing="0" cellpadding="0">
	<tr>
        <td width=30></td>
	<td class="b12" align="left">貸款號碼</td>
	<td width=50></td>
	<td><input type="text" name="lnnum" value="<%=lnnum%>" size="10" maxlength="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td>
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">社員號碼</td>
		<td width=50></td>
		<td><input type="text" name="memNo" value="<%=memNo%>" size="10" maxlength="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
		
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">取票日期</td>
		<td width=50></td>
		<td><input type="text" name="lndate" value="<%=lndate%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">貸款金額</td>
		<td width=50></td>
		<td><input type="text" name="appamt" value="<%=appamt%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">攤分期數</td>
		<td width=50></td>
		<td><input type="text" name="install" value="<%=install%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">每月還款</td>
		<td width=50></td>
		<td><input type="text" name="monthrepay" value="<%=monthrepay%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">本金結餘</td>
		<td width=50></td>
		<td><input type="text" name="bal" value="<%=bal%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">脫期利息</td>
		<td width=50></td>
		<td><input type="text" name="pint" value="<%=pint%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">清數日期</td>
		<td width=50></td>
		<td><input type="text" name="cleardate" value="<%=cleardate%>" size="10" onblur="if(!formatDate(this)){this.value=''};calculation();"></td>
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">清數本金</td>
		<td width=50></td>
		<td><input type="text" name="cashamt" value="<%=cashamt%>" size="10"  maxlength="10" <%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">清數利息</td>
		<td width=50></td>
		<td><input type="text" name="cashint" value="<%=cashint%>" size="10" onblur="if(!formatNum(this)){this.value=''};caladd();"></td>
	</tr>
    
	<tr>
               <td width=30></td>
		<td class="b12" align="left">清數金額</td>
		<td width=50></td>
		<td><input type="text" name="ttlpamt" value="<%=ttlpamt%>" size="10" </TD>
	</tr>
	<tr>
		<td width=30></td>
		<td colspan="3" align="right">
<% if id <> "" then %>
			<%if session("userLevel")<>2 and session("userLevel")<>1 and session("userLevel")<>6 then%>
			<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
			<%end if%>
			<input type="button" value="查詢個人帳" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.cchkdate.value )" class="sbttn">					
<%end if%>
			<input type="submit" value="返回" name="back" class="sbttn">
		</td>
	</tr>
</table>

</center>
</form>
</body>
</html>
