<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
y=year(now)
m=month(now)-1
m = 2 
if m = 0 then
   m = 12
   y = y - 1
end if

server.scripttimeout = 1800

if request("action")<>"" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
        yy  = right(paydate,4)
        mm  = mid(paydate,4,2)
        dd  = left(paydate,2)  
 
        xxdate= dateSerial(yy,mm,dd)
        set rs = server.createobject("ADODB.Recordset")
        conn.begintrans
        sqlstr = "select a.*,b.mstatus  from dividend a ,memmaster b where  a.deleted= 0  and a.dividend > 0 and a.memno=b.memno  "


        rs.open sqlstr, conn ,1,1
        do while not rs.eof
        
           if rs("bank")="H" then
              conn.execute("insert into share (memno , ldate , code , amount ) values ( "&rs(0)&" ,'"&xxdate&"' ,'CH' , "&rs(1)&" ) ")
           else
              conn.execute("insert into share (memno , ldate , code , amount ) values ( "&rs(0)&" ,'"&xxdate&"' ,'C0' , "&rs(1)&" ) ")
            end if
          
           select case rs("bank")
                  case "C"
                       conn.execute("insert into share (memno , ldate , code , amount ) values ( "&rs(0)&" ,'"&xxdate&"' ,'C3' , "&rs(1)&"*-1 ) ")  
                  case "B"
                       conn.execute("insert into share (memno , ldate , code , amount ) values ( "&rs(0)&" ,'"&xxdate&"' ,'C1' , "&rs(1)&"*-1 ) ")

           end select
           
    
        rs.movenext
        loop
        set rs = nothing
      

        conn.execute("update dividend  set deleted= 1  where deleted=0 and bank<>'H'  ")
	conn.committrans
       

        msg = "派息完成!"
	response.redirect "completed.asp"
else
      id = ""
      set rs=conn.execute("select * from dividend  where deleted = 0 and bank<>'H'  ")
       if rs.eof then
           msg = "派息巳完成"
           id = "1"
        end if
        rs.close
end if
%>
<html>
<head>
<title>派息過數</title>

<script language="JavaScript">
<!--
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

	if (formObj.paydate.value==""){
		reqField=reqField+", 派息日期";
		if (!placeFocus)
			placeFocus=formObj.paydate;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.paydate.focus()">
<!-- #include file="menu.asp" -->
<br>
<center><font size="3">派息過數</font>

<form name="form1" method="post" action="<% =Request.servervariables("script_name") %>" onsubmit="return validating()">
<%if msg<>"" then%>
<div align=center><font color="red"><%=msg%></font></div>
<%end if%>
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
               <td width=30></td>
		<td class="b12" align="left"<font size ="2">派息日期</font></td>
		<td width=50></td>
		<td><input type="text" name="paydate" value="<%=paydate%>" size="10" onblur="if(!formatDate(this)){this.value=''};">
                    <input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
                </td>                 
 	</tr>

</table>
</form>
</center>
</body>
</html>
