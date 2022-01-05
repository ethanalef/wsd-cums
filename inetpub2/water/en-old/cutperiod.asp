<!-- #include file="../conn.asp" -->
<!-- #include file="cutpro.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%





if request("action")<>"" then

	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
       yr =  lastyear
       mm = lastmonth -1
       if mm <1 then 
          yr =  lyr - 1
       end if

       if yr/4=int(ry/4) and yr/100=int(yr/100) then
          dd  =mid("312931303130313130313031",(mm-1)*2+1,2)
       else
          dd  =mid("312831303130313130313031",(mm-1)*2+1,2)
       end if
       xdate = yr&"/"&mm&"/"&dd
       ydate = lastyear&"/"&lastmonth&"/"&mday
     
       addUserLog "截數設定        "
       
       conn.execute("update monthend set works= 1 where works = 0 ")
      
       conn.execute(" insert into monthend (lastdate,cutdate,works) values ('"&xdate&"','"&ydate&"',0 ) ")
       
    
       
       response.redirect("completed.asp")

else
  set rs = conn.execute("select * from monthend where works = 0 ")
  if not  rs.eof then
      mday   = day(rs("cutdate"))
     lastyear  = year(rs("cutdate"))
     lastmonth = month(rs("cutdate"))+1
     if lastmonth = 13 then
        lastmonth = 1
        lastyear = lastyear + 1
     end if
    
    
 

  end if
  rs.close

end if
%>
<html>
<head>
<title>>截數設定建立</title>

<script language="JavaScript">
<!--
function formatNum(numform){
  if (isNaN(numform.value)||numform.value<0)
    return false;
  else
    return true;
}
function checkDay(mDay){
  D=mDay.value;
  M=<%=lastmonth%>;
  Y=<%=lastmonth%>;
  if(isNaN(D) || D=="")
    return false;
  Months= new Array(31,28,31,30,31,30,31,31,30,31,30,31);
  Leap  = false;
  if((Y % 4 == 0) && ((Y % 100 != 0) || (Y %400 == 0)))
    Leap = true;
  if((D < 1) || (D > 31))
    return false;
  if((D > Months[M-1]) && !((M == 2) && (D > 28)))
    return false;
  if(!(Leap) && (M == 2) && (D > 28))
    return false;
  if((Leap)  && (M == 2) && (D > 29))
    return false;
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
      if (((Mn<=sMn)&&(Yr=sYr))||(Yr<sYr)){
         return false ;
      }else{      
         return true;
      }

  }
}
function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.period.value==""){
		reqField=reqField+", 截數日期";
		if (!placeFocus)
			placeFocus=formObj.mDay;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.mDay.focus()">

<!-- #include file="menu.asp" -->
<br>
<center>
<h3>截數設定建立</h3>
<form name="form1" method="post" action="<% =Request.servervariables("script_name") %>" onsubmit="return validating()">
<input type="hidden" name="cutdate" value="<%=cutdate%>">
<input type="hidden" name="lastmonth" value="<%=lastmonth%>">
<input type="hidden" name="lastyear" value="<%=lastyear%>">
<input type="hidden" name="spass" value="<%=spass%>">
<%if msg<>"" then%>
<div align=center><font color="red"><%=msg%></font></div>
<%end if%>
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td class="b8">日期</td>
		<td width="10"></td>
		<td>
			<input type="text" name="mDay" value="<%=mDay%>" size="2" maxlength="2" onblur="if(!checkDay(this)){this.value=''};">/<%=lastmonth%>/<%=lastyear%>
			<input type="submit" value="確定" name="action" class="sbttn">
		</td>
	</tr>
</table>
</form>
</center>
</body>
</html>
