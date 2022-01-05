<!-- #include file="../conn.asp" -->
<!-- #include file="../addUserLog.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="cutpro.asp" -->

<%
if request.form("bye") <> "" then
   response.redirect "main.asp"
end if

if request.form("clrScr") <> "" then
     memno=""
     memname=""
     memcname=""
     savamt =""
     ttlbal = 0
     memName ="" 
     memcName =""
     id =""
     
end if
if request.form("Search")<>"" then
   	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
   if id ="" then 
        set rs=conn.execute("select memno,memname,memcname,mstatus from memmaster where memno='"&memno&"'  ")
        if not rs.eof then
	   For Each Field in rs.fields	   
	   TheString = Field.name & "= rs(""" & Field.name & """)"				
	   Execute(TheString)
	   Next	
           if mstatus = "J" then
           id = memno
           yy = year(date())-1
	   ldate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
           todate= "01/"&right("0"&month(date()),2)&"/"&yy 
           cfee = 25.00
           bfee = 1.00
           saveamt = 0	
           ttlbal=cfee+bfee
           else
              msg = "社員不是新戶"	
 	      memno=""
	      memname=""
	      memcname=""
	      savamt =""
              id = ""
           end if           
        else
           msg = "社員不存在"	
	      memno=""
	      memname=""
	      memcname=""
	      savamt =""
              id = ""
        end if          
        rs.close
    else
           yy = year(date())-1
	   ldate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
           todate= "01/"&right("0"&month(date()),2)&"/"&yy 
           cfee = 25.00
           bfee = 1.00
           saveamt = 0	
           ttlbal=cfee+bfee+saveamt 
    end if
end if




if request.form("action") <>"" then

	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next

        pdate = right(ldate,4)&"/"&mid(ldate,4,2)&"/"&left(ldate,2)

 	conn.begintrans
       
           conn.execute("insert into share (memno,code,ldate,amount) values ('"&memno&"','A3','"&pdate&"',"&ttlbal&") ")
      
        conn.execute("insert into share (memno,code,ldate,amount) values ('"&memno&"','G3','"&pdate&"',"&Bfee&") ")
        conn.execute("insert into share (memno,code,ldate,amount) values ('"&memno&"','H3','"&pdate&"',"&cfee&") ")
        conn.execute("update memmaster set mstatus='J' where memno='"&memno&"' ")
        addUserLog "Open Share account member : "&memno&" Share amount "&ttlal
        addUserLog "deduct member fee and group fee "&cfee&" and "&bfee
        addUserLog "Memnber "&memno&" of status from New r to Normal "
	conn.committrans
        amount = 0
        memno=""
        memname = ""
        memcname =""
	addUserLog "add new share account open "
	msg = "紀錄已更新"
end if

%>
<html>
<head>
<title>新社員開戶建立</title>
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
  yy = parseInt(Y); 
  if ( yy < 1900)
      return false;
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
      Mn = strM
      Yr = strY
      if (((Mn<=sMn)&&(Yr==sYr))||(Yr<sYr)){
         return false ;
      }else{      
         return true;
      }
  }
}






function checkId(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (!formatNum(formObj.id)){
        alert("Please fill correct account No.");
		form1.id.select();form1.id.focus();
        return false;
    }else{
        return true;
    }
}



function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (!formatNum(formObjttlbal)){
		reqField=reqField+", 股金";
		if (!placeFocus)
			placeFocus=formObj.ttlbal;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="calculation();form1.memNo.focus();">
<!-- #include file="menu.asp" -->
<BR>
<div><center><font size="3">新社員開戶建立</font></center></div>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form name="form1" method="post" action="newacc.asp">
<table border="0" cellspacing="0" cellpadding="0">
<input type="hidden" name="uid" value="<%=uid%>">
<input type="hidden" name="memName" value="<%=memName%>">
<input type="hidden" name="memcName" value="<%=memcName%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="cutdate" value="<%=cutdate%>">
<input type="hidden" name="lastmonth" value="<%=lastmonth%>">
<input type="hidden" name="lastyear" value="<%=lastyear%>">
<input type="hidden" name="spass" value="<%=spass%>">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td width="400" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">

				<tr>
					<td class="b12" align="right">社員編號</td>
					<td width=10></td>
					<td>
						<input type="text" name="memNo" value="<%=memNo%>" size="10" maxlength="10" >
						<%if id ="" then %>
						<input type="button" value="選擇" onclick="popup('pop_srhNewacc.asp?key='+document.form1.memNo.value)" class="sbttn">
						<input type="submit" value="搜尋" name="Search" class ="Sbttn">
						<%end if%>
						<input type="submit" value="取消" name="clrSrc" class="sbttn">
						<input type="submit" value="返回" name="bye" class="sbttn">
					</td>
				</tr>
				<tr height="22">
					<td class="b12" align="right">姓名</td>
					<td width=10></td>
					<td id="memName"><%=memName%></td>
					<td width="10"></td>
					<td id="memcName"><%=memcName%></td>
				</tr>
<%if id<>"" then %>
				<tr height="22">
					<td class="b12" align="right">日期</td>
					<td width=10></td>
					<td><input type="text" name="ldate" value="<%=ldate%>" size="10" maxlength="10"  onblur="if(!formatDate(this)){this.value=''};"></TD>
				</tr>
				<tr height="22">
					<td class="b12" align="right">存入金額</td>
					<td width=10></td>
					<td><input type="text" name="ttlbal" value="<%=ttlbal%>" size="10" maxlength="10" onblur="if(!formatNum(this)){this.value=''};" ></td>
				</tr>
				<tr height="22">
					<td class="b12" align="right">協會費</td>
					<td width=10></td>
					<td><input type="text" name="cfee" value="<%=cfee%>" size="10" maxlength="10"onblur="if(!formatNum(this)){this.value=''};" ></td>
				</tr>
				<tr height="22">
					<td class="b12" align="right">入社費</td>
					<td width=10></td>
					<td><input type="text" name="bfee" value="<%=bfee%>" size="10" maxlength="10"onblur="if(!formatNum(this)){this.value=''};" ></td>
				</tr>

<%end if %>
		</td>
	</tr>
</table>
				<tr>
					<td colspan="3" align="right">
<%if id <>"" then %>
						<%if session("userLevel")=5 OR  session("userLevel")<>3 OR  session("userLevel")<>4 then%>
						<input type="submit" value="確定" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
						<%end if%>
<%end if%>

				</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<br>

</center>
</form>
</body>
</html>
