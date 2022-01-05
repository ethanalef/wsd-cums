<!-- #include file="../conn.asp" -->
<!-- #include file="init.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "lcloan.asp"'
end if

xln = request("lnnum")
if xln <> "" then
pos  = instr(xln,"*")
if pos > 0 then
xmemno = left(xln,pos -1)
xlnnum = mid(xln,pos+1,10)
xlndate = mid(xln,pos+11,10)
end if
end if



if request.form("action") <> "" then
	msg = ""
        addloan = 0
	For Each Field in Request.Form
 	 
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)

	
	Next

	if msg="" then
           if newlnnum <> "" then
		conn.begintrans  
                xcldate =right(cleardate,4)&"/"&mid(cleardate,4,2)&"/"&left(cleardate,2)   
                yy = year(cleardate)
                mm = month(cleardate)
                dd = day(cleardate)
                                
                if ((yy/4)= int(yy/4) and (yy/100)=int(yy/100)) then
                   daylist="312931303130313130313031"
                   mD = mid(daylist,(mm-1)*2+1,2)
                else
                   daylist="312831303130313130313031"
                   mD = mid(daylist,(mm-1)*2+1,2)
                end if
                
      
          
                conn.execute("update loanrec set cleardate='"&xcldate&"' ,bal=0,repaystat='C' where lnnum='"&lnnum&"' ")	                
                set rs=conn.execute("select amount from loan where lnnum='"&newlnnum&"' and code='D9'  ")
                newappamt = rs(0) 
                rs.close
                set rs=nothing

              
   
                conn.execute("insert into loan (memno,lnnum,code,ldate,amount) values ( '"&memno&"','"&lnnum&"','D8','"&xcldate&"',"&bal&") ")
	    
                xlnnum=""   
                conn.committrans
		msg = "紀錄已更新"             
             end if
	end if
else

	if xlnnum <> "" then
                set ms=conn.execute("select oldlnnum,lndate from loanrec where lnnum='"&xlnnum&"' ")
                if not ms.eof then
                     lnnum=  ms(0)
                     xdate = dmy(ms(1))
                 end if
                 ms.close


		sql = "select * from loanrec where repaystat='N' and lnnum='"&lnnum&"'  "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn ,1,1

		if rs.eof then
			response.redirect "lcloan.asp"
		else
			For Each Field in rs.fields
			
					TheString = Field.name & "= rs(""" & Field.name & """)"
			
				Execute(TheString)
			Next
		end if

         end if  
         newlnnum = xlnnum
         cleardate = xdate
         lndate = dmy(rs("lndate"))   
         
         cchkdate  = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1)    
end if
%>
<html>
<head>
<title>循環貸款建立</title>
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

      Mn = strM
      Yr = strY
      if (((Mn<=sMn)&&(Yr=sYr))||(Yr<sYr)){
         return dalse ;
      }else{      
         return true;
      }

  }
}




function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;


	if (!formatDate(formObj.cleardate)){
		reqField=reqField+", 清數日期";
		if (!placeFocus)
			placeFocus=formObj.cleardate;
	}

	if (formObj.newlnnum.value==""){
		reqField=reqField+", 新貸款號碼";
		if (!placeFocus)
			placeFocus=formObj.newlnnum;
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
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
</center>
<br>
<form name="form1" method="post" action="lcloanDetail.asp">

<input type="hidden" name="xlnnum" value="<%=xlnnum%>">
<input type="hidden" name="cchkdate" value="<%=cchkdate%>">
<input type="hidden" name="cutdate" value="<%=cutdate%>">
<input type="hidden" name="lastmonth" value="<%=lastmonth%>">
<input type="hidden" name="lastyear" value="<%=lastyear%>">
<input type="hidden" name="spass" value="<%=spass%>">
<div><center><font size="3">循環貸款建立</font></center></div>

<center>

	<td width="700" valign="top">
	<table border="0" cellspacing="0" cellpadding="0">
		<tr>
                <td width=30></td>
		<td class="b12" align="left">貸款號碼</td>
		<td width=50></td>
		<td><input type="text" name="lnnum" value="<%=lnnum%>" size="10" maxlength="10"<%if xlnnum<>"" then response.write " onfocus=""form1.action.focus();""" end if%>></td>
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">社員號碼</td>
		<td width=50></td>
		<td><input type="text" name="memNo" value="<%=memNo%>" size="10" maxlength="10"<%if xlnnum<>"" then response.write " onfocus=""form1.action.focus();""" end if%>></td> 
		
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left"><input type="hidden" name="xstatus" value="<%=xstatus%>">日期</td>
		<td width=50></td>
		<td><input type="text" name="lndate" value="<%=lndate%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.action.focus();""" end if%>></td> 
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">貸款金額</td>
		<td width=50></td>
		<td><input type="text" name="appamt" value="<%=appamt%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.action.focus();""" end if%>></td> 
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">攤分期數</td>
		<td width=50></td>
		<td><input type="text" name="install" value="<%=install%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.action.focus();""" end if%>></td> 
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">每月還款</td>
		<td width=50></td>
		<td><input type="text" name="monthrepay" value="<%=monthrepay%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.action.focus();""" end if%>></td> 
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">本金結餘</td>
		<td width=50></td>
		<td><input type="text" name="bal" value="<%=bal%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.action.focus();""" end if%>></td> 
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">清數日期</td>
		<td width=50></td>
		<td><input type="text" name="cleardate" value="<%=cleardate%>" size="10" onblur="if(!formatDate(this)){this.value=''}">(dd/mm/yyyy)</td>
	</tr>
	<tr>
               <td width=30></td>
		<td class="b12" align="left">新貸款號碼</td>
		<td width=37></td>
		<td><input type="text" name="newlnnum" value="<%=newlnnum%>" size="10" onblur="if(!formatDate(this)){this.value=''};">
	</tr>   
	<tr>
		<td width=30></td>
		<td colspan="3" align="right">
		<% if xlnnum <> "" then %>
		<%if session("userLevel")<>2 and session("userLevel")<>1 and session("userLevel")<>6 then%>
		<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
		<%end if%>
		<input type="button" value="查詢貸款" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.cchkdate.value )" class="sbttn">
		<%end if%>
		<input type="submit" value="返回" name="back" class="sbttn">
		</td>
	</tr>
</table>
</center>
</form>
</body>
</html>
