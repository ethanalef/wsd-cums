<!-- #include file="../conn.asp" -->
<!-- #include file="../addUserLog.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->

<%

if request.form("Bye") <> "" then
   response.redirect "main.asp"

end if

id = request("id")
if id <>"" then
   memno = id
end if

if  request.form("clrScn")<>"" then
                id = ""
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next
end if

if request.form("Search")<>"" or id <>""  then
     
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
        idx  = accode
		sql = "select * from memMaster where status = '*' and memNo='"& memNo & "' "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if not rs.eof then

			For Each Field in rs.fields
			if Field.name="memBday" or Field.name="firstAppointDate" or Field.name="memDate" or Field.name="Wdate" then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next

                id = memno
                

                rs.close
                todate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
              
                if bnk <> "" then
  		sql =  "select * from bank where bncode='"& bnk & "' "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
                if not rs.eof then
                   bnkname = rs("bank")
                end if
                rs.close
               
                end if
                else
                  msg = memno& "不是委員"

                end if
         
end if
if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
 
	msg = ""
        if idx <>"" then
           pos = instr(idx,"-")
           xaccode = left(idx,pos-1)
        else
           xaccode="    "
        end if
 
        set rs1 = conn.execute("select memno,memname,memcname from memmaster where status='*' and memno= '"&memno&"' ")       
        if  rs1.eof then
            msg =memno& "不是委員"
        end if 
	if msg="" then
		conn.begintrans
                 addUserLog "Change Group from "&xaccode&" To "&memno
                conn.execute("update memmaster set accode='"&memno&"' where accode= '"&xaccode&"' " )
       
		conn.committrans
		msg = "紀錄已更新"
	end if
        
                id = ""
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next
end if
%>
<html>
<head>
<title>轉換聯絡人建立</title>
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


function callage(){
	 formObj=document.form1;
        
            ssdate= formObj.memBday.value
            ttdate= formObj.todate.value
            YY=parseInt(ssdate.substr(6,4)) 
	    YYY=parseInt(ttdate.substr(6,4)) 
            xyy = YYY - YY
            document.all.tags( "td" )['age'].innerHTML=xyy; 
}

function formatHKID(dateform){
  HKID = dateform.value
  fchar = HKID.substr(0,1) 
  Uchar ='ABCDEFGHIJKLMNOPQRSTUVWXYZ'
  dSize = HKID.length-1
  if (dSize==7){
     
     sCount = 0
     for(var i=1; i < 28; i++)
     (Uchar.substr(i-1,1) == fchar) ? sCount=i : sCount
     ttl = 8*sCount
     i = 1
     while ( i < 7 ) {

              ttl = ttl + (8-i)*(HKID.substr(i,1))
 
        i++
     }    
   
     a1 = ttl % 11
     if (HKID.substr(7,1)=='A'){
        if (a1==1) 
           return true
     }  
     if (HKID.substr(7,1)=='0'){
        if (a1==0) 
           return true
     }  
     a2 = 11 - a1
     if (HKID.substr(7,1)==a2){
           return true
     }       
     return false;
  }else{
   return false;
  }

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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="callage();form1.memno.focus()">
<form method="post" action="chgroup.asp" name="form1">
<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="age3" value="<%=age%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="idx" value="<%=idx%>">

<!-- #include file="menu.asp" -->
<div align="center"><center>
<br><font size="3">轉換聯絡人建立</font>
<br>
<%if msg<>"" then%>
<div align=center><font color="red"><%=msg%></font></div>
<%end if%>
<br>
<table border="0" cellpadding="0" cellspacing="0">
                <tr>
                    <td align="right" ><font size="2" >分組/組別</formt></td>
                    <td width="10"></td>
<%if id="" then %>
 		    <td>
		    <select name="accode">
                    <option>
		    <option<% if accode="9999" then %> selected<% end if%>>9999 - 工作人員
<%
                     set rs=nothing 
                     set rs=conn.execute("select  memno,memcname,memname,status from memmaster where  status='*'   order by memno  "    )
                         do while not rs.eof
                            if  rs(3)="*" then
                            idx = rs(0)&"-"&rs(2)&" "&rs(1)
                      
%> 
                     		
			<option<% if accode=rs(0) then %> selected<% end if%>><%=idx%>
<%
                           end if            
                        rs.movenext
                        loop
                        rs.close 
			
%>                  
		    </select>
                   </td>
<%else%>
           <td id ="idx"><%=idx%></td>
                            
<%end if %>
		
        <tr>
               	<td align="right" ><font size="2" >委員編號</font></td>
		<td width="10"></td>

		<td><input type="text" name="memNo" value="<%=memNo%>" size="4" >
<%if id="" then %>
		
		<input type="submit" value="確定" onclick="return validating()&&confirm('確定轉組?')" name="Search" class ="Sbttn">               
<%end if%>
                </td>
	 </td> 
         </tr>
<% if id <>"" then %>	
                <tr>
                    <td align="right" ><font size="2" >英文名稱</formt></td>
                    <td width="10"></td>
                    <td><input type="text" name="memName" value="<%=memName%>" size="35" ></td>
                    <td><font size="2" >中文名稱</formt></td>
                    <td width="10"></td>
                    <td><input type="text" name="memcName" value="<%=memcName%>" size="10" ></td>
               </tr>
               <tr>
                    <td align="right" ><font size="2" >身份證號碼</formt></td>
                    <td width="10"></td>
                    <td><input type="text" name="memHKID" value="<%=memHKID%>" size="20" onblur="if(!formatHKID(this)){this.value=''}"></td> 
                    <td><font size="2" >性別</formt></td>
                    <td width="10"></td>
		    <td>
			<select name="memGender">
			<option<%if memGender="M" then response.write " selected" end if%>>男</option>
			<option<%if memGender="F" then response.write " selected" end if%>>女</option>
			</select>
		    </td>                         
               </tr> 
                <tr>
                <td align="right" ><font size="2" >出生日期</td>  
                <td width="10"></td>
                <td><input type="text" name="memBday" value="<%=memBday%>" size="10" maxlength="10" onblur="if(!formatDate(this)){this.value=''};callage();">(dd/mm/yyyy)</td>
                <td align="left" ><font size="2" >年齡</font></td>  
                <td width="10"></td>  
	        <td id="age"><%=age%></td>
                </tr>        
              <tr>
                    <td align="right" ><font size="2" >入職日期</formt></td>
                    <td width="10"></td>
                    <td><input type="text" name="AppointDate" value="<%=AppointDate%>" size="10" onblur="if(!formatDate(this)){this.value=''};callage();">(dd/mm/yyyy)</td>
                    <td><font size="2" >受顧條件</formt></td>
                    <td width="10"></td>
		    <td><input type="text" name="employCond" value="<%=employCond%>" size="20" maxlength="20"></td>
	                             
               </tr>             
               <tr>
                  <td align="right" ><font size="2" >職位</formt></td>
                    <td width="10"></td>
                    <td><input type="text" name="memGrade" value="<%=memGrade%>" size="20" ></td>
                  <td><font size="2" >部門</formt></td>
                    <td width="10"></td>
		    <td><input type="text" name="memSection" value="<%=memSection%>" size="20" maxlength="20"></td>
		                             
               </tr>
  
<%end if %>
</table>
<table border="0" cellpadding="0" cellspacing="0">
	<tr >

		<td>
<%if id<>"" then %>
			<input type="submit" name="action" value="確定儲存" onclick="return validating()&&confirm('確定儲存?')"  class="sbttn">
<%end if%>
                        <input type="submit" name="clrScn" value=" 取  消 "   class="sbttn" > 
                        <input type="submit" name="Bye" value=" 返  回 "   class="sbttn" ></td> 
		</td>
	</tr>
</table>
</form>
</center></div>
</body>
</html>