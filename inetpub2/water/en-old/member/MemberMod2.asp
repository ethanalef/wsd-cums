<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
 userlevel = session("userlevel")

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
       BnkName = ""
        xbnk = ""
        xbch = ""
        xbacct =""
        xmonthsave = ""
        xmonthssave = ""
        xbnklmt = ""
        xtpayamt = ""        
end if

if request.form("Search")<>"" or id <>""  then
     
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next

		sql = "select * from memMaster where memNo='"& memNo & "' "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if not rs.eof then

			For Each Field in rs.fields
			if Field.name="memBday" or Field.name="appointdate" or Field.name="memDate" or Field.name="Wdate" then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
            if userlevel <> 4 then
                select case rs("mstatus") 
                       	case  "L"
                              xstatus =  "呆帳"
                             
                        case  "D" 
                             xstatus =  "冷戶"  
		
                        case "V"
                              xstatus = "IVA"
                        
                        case "C"
                              xstatus = "退社" 
		
                        case  "B"
                              xstatus = "破產" 
			 
                        case  "P"
                              xstatus = "去世"
			       
                        case  "N" 
                              xstatus = "正常"
			    
                        case  "J" 
                              xstatus = "新戶"
                          
                        case  "T" 
                              xstatus = "庫房" 
                           
                        case  "H" 
                               xstatus = "暫停銀行"
			
                        case  "A"
                               xstatus =  "自動轉帳(ALL)"
			
                        case "0"
                              xstatus = "自動轉帳(股金)"
			 
                        case "1"
                              xstatus = "自動轉帳(股金,利息)"
			
                        case "2"
                              xstatus = "自動轉帳(股金,本金)"                         
			 
                        case "3"
                             xstatus = "自動轉帳(利息,本金)"                         
			 
                        case "M"
                             xstatus = "庫房,銀行"
			   
                        case "F"
                              xstatus = "特別個案"  
			  
                        case "8"
                             xstatus = "終止社籍轉帳"
                          
                        case "9"
                             xstatus = "終止社籍正常"
                          
                     
               end select 
               xbnk = bnk
               xbch = bch
               xbacct = bacct
               xmonthsave = monthsave
               xmonthssave = monthssave
               xtpayamt = tpayamt
               xbnklmt = bnklmt
               end if
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
                  msg = "社員號碼不存在 "

                end if
     
end if



if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
        set rs = server.createobject("ADODB.Recordset")
	msg = ""

 

	if msg="" then
		conn.begintrans
		if id = "" then
			sql = "select * from memMaster where 0 = 1"

		else
			sql = "select * from memMaster where memNo=" & id
		end if
		rs.open sql, conn, 2, 2
		if id = "" then

			rs.addnew
			rs("memNo") = memNo
			rs("deleted") = 0
			id = rs("memNo")
			addUserLog "Add Member"
		else
			addUserLog "Modify Member Detail"
		end if

		rs("memName") = memName
                rs("memcname")= memcname
		rs("memAddr1") = memAddr1
		rs("memAddr2") = memAddr2
		rs("memAddr3") = memAddr3
		rs("memContactTel") = memContactTel
                rs("memofficetel") = memofficetel

		rs("memMobile") = memMobile
                rs("mememail") =  mememail
                rs("employCond") = employCond
		if EDCOMM<>"" then rs("EDCOMM") = EDCOMM else rs("EDCOMM")=0 end if
                rs("treasRefNo") = treasRefNo
		rs("memHKID") = memHKID
		  if memGender="男" then 
		   rs("memGender") = "M" 
                else
                  rs("memGender") = "F" 
                end if
                select case status 
                       case "委員"
                          rs("status")= "*"
                       case "社員"
                          rs("status")= "0"
                end select 
                if userlevel = 4 then
                rs("bnk") = bnk
                rs("bch") = bch
                rs("bacct") = bacct

                
                select case mstatus 
                       	case "呆帳"
                             rs("mstatus")= "L"
                        case  "冷戶"  
			   rs("mstatus")="D" 
                        case "IVA"
                           rs("mstatus")="V"  
                        case "退社" 
			   rs("mstatus")="C" 
                        case  "去世" 
			   rs("mstatus")="P" 
                        case "破產"
			   rs("mstatus")="B" 
                        case "正常"
			    rs("mstatus")="N" 
                        case  "新戶"
                           rs("mstatus")="J" 
                        case  "庫房" 
                           rs("mstatus")="T" 
                        case "暫停銀行"
			   rs("mstatus")="H" 
                        case "自動轉帳(ALL)"
			   rs("mstatus")="A" 
                        case "自動轉帳(股金)"
			   rs("mstatus")="0" 
                        case "自動轉帳(股金,利息)"
			   rs("mstatus")="1" 
                        case "自動轉帳(股金,本金)"                         
			   rs("mstatus")="Z" 
                        case "自動轉帳(利息,本金)"                         
			   rs("mstatus")="3" 
                        case "庫房,銀行"
			    rs("mstatus")="M" 
                        case "特別個案"  
			    rs("mstatus")="F"   
                        case "終止社籍轉帳"
                           rs("mstatus")="8"
                        case "終止社籍正常"
                           rs("mstatus")="9"                             
                     
               end select 
                if tpayamt <>"" then rs("tpayamt")  = tpayamt else tpayamt=null end if
                if monthsave <> "" then rs("monthsave") = monthsave  else rs("monthsave") = null end if
                if monthssave <>"" then rs("monthssave") = monthssave else rs("monthssave")= null end if         
                if bnklmt<>"" then  rs("bnklmt") = bnklmt  else rs("bnklmt")=null end if
                rs("bacct") = bacct            
                end if 
                if AppointDate<>"" then rs("AppointDate") = right(AppointDate,4)&"/"&mid(AppointDate,4,2)&"/"&left(AppointDate,2) else rs("AppointDate")=NULL end if
                RS("REMARK") = REMARK
                RS("B1") = B1
                RS("B1ID") = B1ID
                RS("B1RELATION") = B1RELATION
                RS("B1ADD") = B1ADD
                rs("B1Add2") = B1Add2 
                RS("B2") = B2
                RS("B2ID") = B2ID
                RS("B2RELATION") = B2RELATION
                RS("B2ADD") = B2ADD
                Rs("B2Add2") = B2Add2

		if memBday<>"" then rs("memBday") = right(memBday,4)&"/"&mid(memBday,4,2)&"/"&left(memBday,2) else rs("memBday")=NULL end if
                if wdate<>"" then rs("wdate") = right(wdate,4)&"/"&mid(wdate,4,2)&"/"&left(wdate,2) else rs("wdate")=NULL end if
		rs("memGrade") = memGrade
		rs("memSection") = memSection
		
		rs("treasRefNo") = treasRefNo
		rs("employCond") = employCond
		if firstAppointDate<>"" then rs("firstAppointDate") = right(firstAppointDate,4)&"/"&mid(firstAppointDate,4,2)&"/"&left(firstAppointDate,2) else rs("firstAppointDate") = NULL  end if
		if memDate<>"" then rs("memDate") = right(memDate,4)&"/"&mid(memDate,4,2)&"/"&left(memDate,2) else rs("memdate")=NULL end if
                if memincome<>"" then rs("memincome") = memincome else rs("memincome")=null end if
		rs("introd1") = introd1
                rs("introd2") = introd2
                rs("memoadd1") = memoadd1
                rs("memoAdd2") = memoadd2
                rs("memoadd3") = memoadd3
                pos = instr(accode,"-")
                if pos > 0 then
                   xaccode =cint( left(accode,pos-1) ) 
                else
                   xaccode = null
                end if
                rs("accode")=xaccode
                rs.update
		conn.committrans
		msg = "紀錄已更新"
	end if
                id = ""
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next
       
        BnkName = ""
        xbnk = ""
        xbch = ""
        xbacct =""
        xmonthsave = ""
        xmonthssave = ""
        xbnklmt = ""
        xtpayamt = ""
end if
%>
<html>
<head>
<title>社員資料修正</title>
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


function callage(){
	 formObj=document.form1;

         
            ssdate= formObj.memBday.value
            ttdate= formObj.todate.value
            DD=parseInt(ssdate.substr(0	,2))
            if (DD == 0 ) {
               DD=parseInt(ssdate.substr(1,1))
            }
            MM=parseInt(ssdate.substr(3	,2))   
            if (MM ==0){
               MM=parseInt(ssdate.substr(4,1))   
            }
            YY=parseInt(ssdate.substr(6	,4))
            xD=parseInt(ttdate.substr(0	,2))
            if (xD == 0 ) {
               xD=parseInt(ttdate.substr(1,1))
            }
            xM=parseInt(ttdate.substr(3	,2))   
            if (xM ==0){
               xM=parseInt(ttdate.substr(4,1))   
            } 
            xY=parseInt(ttdate.substr(6	,4))            

               ndate1 = xY+'/'+parseInt(ssdate.substr(3	,2))  
           if (xM>=10) {
               ndate2 = xY+'/'+parseInt(ttdate.substr(3,2))  
           }else{
                ndate2 = xY+'/0'+parseInt(ttdate.substr(3,2))  
           } 


            if (ndate1 > ndate2){
               nday =1- parseInt(((MM - xM)*30+(30-DD))/365.25*10)/10
               a  =  xY - YY - 1+ nday;
            }else{
              a = xY - YY; 
            }   
        
            document.all.tags( "td" )['age'].innerHTML= a ; 
}


function formatHKID(dateform){
  HKID = dateform.value
 
  fchar = HKID.substr(0,1) 
  schar = HKID.substr(1,1) 
   
  Uchar ='ABCDEFGHIJKLMNOPQRSTUVWXYZ'
  y    = HKID.length
  dSize = HKID.length - 1
  lchar = HKID.substr(y,1) 
  pass = 1
  if (dSize == 7) 
  {
     sCount = 0
     for(var i=1; i < 28; i++)
     (Uchar.substr(i-1,1) == fchar) ? sCount=i : sCount
     if (sCount==0){
         alert('輸入錯誤,身份證第一位置是大楷英文字母');
         return false;
     }    
     ttl = 36 * 9 + ( 64 + sCount - 55) * y
     i = 1
     while ( i < 7 ) {
          
              ttl = ttl + (8-i)*(HKID.substr(i,1))
 
        i++
     }    
   
     a1 = 11 - ttl % 11
     if (HKID.substr(7,1)=='A'){
        if (a1==10) 
          
           return true
     }  
     if (HKID.substr(7,1)=='0'){
        if (a1==11) 
           
           return true
     }  
  
     if (HKID.substr(7,1)==a1){
           return true
     }       
     alert('輸入錯誤,身份證第八位置的數值不正確');    
     return false;     
  }
  if (dSize == 8) 
  {
     sCount = 0
     for(var i=1; i < 26; i++)
     (Uchar.substr(i-1,1) == fchar) ? sCount=i : sCount
     if (sCount==0){
         alert('輸入錯誤,身份證第一位置是大楷英文字母');
         return false;
     }    
     ttl =  ( 64 + sCount - 55) * y
     sCount = 0
     for(var i=1; i < 26; i++)
     (Uchar.substr(i-1,1) == schar) ? sCount=i : sCount
     if (sCount==0){
         alert('輸入錯誤,身份證第二位置是大楷英文字母');
         return false;
     }    
     ttl =ttl + (( 64 + sCount - 55) * (y -1))
     i = 2
     while ( i < 8 ) {
          
              ttl = ttl + (9-i)*(HKID.substr(i,1))
 
        i++
     }    
   
     a1 = 11 - ttl % 11
     if (HKID.substr(8,1)=='A'){
        if (a1==10) 
          
           return true
     }  
     if (HKID.substr(8,1)=='0'){
        if (a1==11) 
           
           return true
     }  
  
     if (HKID.substr(8,1)==a1){
           return true
     }       
     alert('輸入錯誤,身份證第九位置的數值不正確');    
     return false;     
  }
 
}


function BankChange(){
	if (document.form1.bnk.value==''){
		document.form1.bnk.value=''
		document.all.tags( "td" )['bnkName'].innerHTML=''
	}else{
	popup('pop_srhBank.asp?key='+document.form1.bnk.value)
        
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

	if (formObj.memName.value==""){
		reqField=reqField+", 姓名";
		if (!placeFocus)
			placeFocus=formObj.memName;
	}

	if (formObj.memHKID.value==""){
		reqField=reqField+", 身分證號碼";
		if (!placeFocus)
			placeFocus=formObj.memHKID;
	}

	if (!formatDate(formObj.memBday)){
		reqField=reqField+", 出生日期";
		if (!placeFocus)
			placeFocus=formObj.memBday;
	}



	if (!formatDate(formObj.memDate)){
		reqField=reqField+", 入社日期";
		if (!placeFocus)
			placeFocus=formObj.memDate;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="callage();form1.memNo.focus()">
<DIV>

<!-- #include file="menu.asp" -->
<form name="form1" method="post" action="memberMod2.asp">
<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="age3" value="<%=age%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="xstatus" value="<%=xstatus%>">
<input type="hidden" name="xbnk" value="<%=xbnk%>">
<input type="hidden" name="xbch" value="<%=xbch%>">
<input type="hidden" name="xbacct" value="<%=xbacct%>">

<input type="hidden" name="xbnklmt" value="<%=xbnklmt%>">
<input type="hidden" name="xmonthsave" value="<%=xmonthsave%>">
<input type="hidden" name="xmonthssave" value="<%=xmonthssave%>">
<input type="hidden" name="xtpayamt" value="<%=xtpayamt%>">

<table border="0" cellpadding="0" cellspacing="0" width="100%" bgcolor="#87CEEB">
<tr >
     <td><font size="4">社員資料修正</font></td>
     <td align="right">
     <% if id <>"" then %>
     <%if session("userLevel")<>5 and session("userLevel")<>1 and session("userLevel")<>6 then%>
     <input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
     <%end if%>
     <%else%>

     <%end if%>
     <input type="submit" name="clrScn" value=" 取  消 "   class="sbttn" > 
     <input type="submit" name="Bye" value=" 返  回 "   class="sbttn" ></td> 
</tr>

</table>
<%if msg<>"" then%>
<div align=center><font color="red"><%=msg%></font></div>
<%end if%>
<table border="0" cellpadding="0" cellspacing="0" >
<tr>
                <td width=12"></td>
     		<td><font size="2" >社員號碼</formt></td>                
		<td  width="23"></td>
		<td><input type="integer" name="memNo" value="<%=memNo%>" size="10" <%if id<>"" then response.write " onfocus=""form1.memName.focus();""" end if%>>
		<%if id = "" then %>
		<input type="button" value="選擇"  onclick="popup('pop_srhMemnoM.asp')" class="sbttn"  >
		<input type="submit" value="搜尋" name="Search" class ="Sbttn">
		<%end if%>
                </td>
 </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" >
<tr>

                <td width=12"></td>
     		<td><font size="2" >社員狀況</formt></td>
                <td width=21></td>
		<td>
                <% if userlevel=4 then %>        
			<select name="mstatus">
                        <option></option>
			<option<%if mstatus="L" then response.write " selected" end if%>>呆帳</option>
			<option<%if mstatus="D" then response.write " selected" end if%>>冷戶</option>
                        <option<%if mstatus="V" then response.write " selected" end if%>> IVA </option>
			<option<%if mstatus="C" then response.write " selected" end if%>>退社</option>
                        <option<%if mstatus="x" then response.write " selected" end if%>>凍結</option>
			<option<%if mstatus="P" then response.write " selected" end if%>>去世</option>
			<option<%if mstatus="B" then response.write " selected" end if%>>破產</option>
			<option<%if mstatus="N" then response.write " selected" end if%>>正常</option>
                        <option<%if mstatus="J" then response.write " selected" end if%>>新戶</option>
                        <option<%if mstatus="T" then response.write " selected" end if%>>庫房</option>
			<option<%if mstatus="H" then response.write " selected" end if%>>暫停銀行</option>
			<option<%if mstatus="A" then response.write " selected" end if%>>自動轉帳(ALL)</option>
			<option<%if mstatus="0" then response.write " selected" end if%>>自動轉帳(股金)</option>
			<option<%if mstatus="1" then response.write " selected" end if%>>自動轉帳(股金,利息)</option>
			<option<%if mstatus="2" then response.write " selected" end if%>>自動轉帳(股金,本金)</option>
			<option<%if mstatus="3" then response.write " selected" end if%>>自動轉帳(利息,本金)</option>
			<option<%if mstatus="M" then response.write " selected" end if%>>庫房,銀行</option>
			<option<%if mstatus="F" then response.write " selected" end if%>>特別個案</option>
                        <option<%if mstatus="8" then response.write " selected" end if%>>終止社籍轉帳</option>
                        <option<%if mstatus="9" then response.write " selected" end if%>>終止社籍正常</option>
			</select>
                <%else%>
                        <%=xstatus%>
                <%end if%>
		</td>
		<td width=111></td>
                <td><font size="2" >代表身份</formt></td>
		<td width=13</td>
		<td>
			<select name="status">

                       
			<option<%if status="*" then response.write " selected" end if%>>委員   </option>
                        <option<%if status="0" then response.write " selected" end if%>>社員   </option>
			</select>
		</td>
</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%"  >
<tr>
      
        <td width="1%" valign="top">
         </td>
      	<td width="60%" valign="top">
 		<table border="0" cellspacing="0" cellpadding="0" >
                <tr>
                    <td><font size="2" >英文姓名</formt></td>
                    <td width="10"></td>
                    <td><input type="text" name="memName" value="<%=memName%>" size="35" ></td>
                    <td><font size="2" >中文姓名</formt></td>
                    <td width="10"></td>
                    <td><input type="text" name="memcName" value="<%=memcName%>" size="10" ></td>
               </tr>
               <tr>
                    <td><font size="2" >身份證號碼</formt></td>
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
                <td align="left" ><font size="2" >出生日期</td>  
                <td width="10"></td>
                <td><input type="text" name="memBday" value="<%=memBday%>" size="10" maxlength="10" onblur="if(!formatDate(this)){this.value=''};callage();">(dd/mm/yyyy)</td>
                <td align="left" ><font size="2" >年齡</font></td>  
                <td width="10"></td>  
	        <td id="age"><%=age%></td>
                </tr>        
              <tr>
                    <td><font size="2" >入職日期</formt></td>
                    <td width="10"></td>
                    <td><input type="text" name="AppointDate" value="<%=AppointDate%>" size="10" onblur="if(!formatDate(this)){this.value=''}">(dd/mm/yyyy)</td>
                    <td><font size="2" >受顧條件</formt></td>
                    <td width="10"></td>
		    <td><input type="text" name="employCond" value="<%=employCond%>" size="20" maxlength="20"></td>
		                             
               </tr>             
               <tr>
                    <td><font size="2" >職位</formt></td>
                    <td width="10"></td>
                    <td><input type="text" name="memGrade" value="<%=memGrade%>" size="20" ></td>
                  <td><font size="2" >部門</formt></td>
                    <td width="10"></td>
		    <td><input type="text" name="memSection" value="<%=memSection%>" size="20" maxlength="20"></td>
		                             
               </tr>
  
                <tr>
                    <td><font size="2" >聯絡員</formt></td>
                    <td width="10"></td>
 		    <td>
		    <select name="accode">
                    <option>
		    <option<% if accode="9999" then %> selected<% end if%>>9999 - 工作人員
<%
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
                    <td><font size="2" >月薪</formt></td>
                    <td width="10"></td>
		<td><input type="text" name="memincome" value="<%=memincome%>" size="20" maxlength="20"></td>
		                             
               </tr>
               <tr>
                    <td><font size="2" >入社日期</formt></td>
                    <td width="10"></td>
                    <td><input type="text" name="memDate" value="<%=memDate%>" size="10" onblur="if(!formatDate(this)){this.value=''}"></td>
                    <td><font size="2" >退社日期</formt></td>
                    <td width="10"></td>
		    <td><input type="text" name="wdate"   value="<%=wdate%>"   size="10" onblur="if(!formatDate(this)){this.value=''}"></td>
		                             
               </tr> 
              </table>
 	       <table border="0" cellspacing="0" cellpadding="0" >
               <tr>
		<td><font size="2" >銀行</formt></td>
                <td width=48></td>	
                <%if userlevel=4 then%>	
                     <td><input type="text" name="bnk" value="<%=bnk%>" size="3"  onchange="BankChange()"></td>
                <%else%>
                     <td><%=xbnk%>
                 <%end if%> 
	
		<td width=5></td>
		<td><input type="text" name="bnkName" value="<%=bnkName%>" size="60"<%if memHKID<>"" then response.write " onfocus=""form1.bch.focus();""" end if%>></td> 
                </tr>  
                </table>
	        <table border="0" cellspacing="0" cellpadding="0" >
       	 	<tr>         
		<td><font size="2" >分行</formt></td>
		<td width=18></td>
                <%if userlevel = 4 then %>  
		    <td><input type="text" name="bch" value="<%=bch%>" size="3" ></td>
                <%else%> 
                    <td><%=xbch%></td>
                <%end if%> 
		<td width=10></td>
		<td><font size="2" >銀行帳號</formt></td>
		<td width=10></td>
                <%if userlevel = 4 then %>
		     <td><input type="text" name="bacct" value="<%=bacct%>" size="15" ></td>
                <%else%> 
                    <td><%=xbacct%></td>
                <%end if%> 
		<td width=10></td>
		<td><font size="2" >轉帳上限</formt></td>
		<td width=1></td>
                <%if userlevel = 4 then %>
		    <td><input name="bnklmt" value="<%=bnklmt%>" size="10" ></td>                                 
                <%else%> 
                    <td align="right"><%=xbnklmt%></td>
                <%end if%> 
	</tr>
        <tr>
		<td><font size="2" >每月儲蓄(銀行)</formt></td>
		<td width=5></td>
                <%if userlevel = 4 then %>
		     <td><input type="text" name="monthsave" value="<%=monthsave%>" size="10" maxlength="10"></td>	
                <%else%> 
                    <td align="right"><%=xmonthsave%></td>
                <%end if%> 
		<td width=5></td>
		<td><font size="2" >每月儲蓄(庫房)</formt></td>
		<td width=5></td>
                <%if userlevel = 4 then %>  
		     <td><input type="text" name="monthssave" value="<%=monthssave%>" size="10" maxlength="10"></td>	
                <%else%> 
                    <td align="right"><%=xmonthssave%></td>
                <%end if%> 
		<td width=5></td>
		<td><font size="2" >庫房過數</formt></td>
		<td width=5></td>
                <%if userlevel = 4 then %>
		     <td><input type="text" name="tpayamt" value="<%=tpayamt%>" size="10" maxlength="10"></td>
                <%else%> 
                    <td align="right" ><%=xtpayamt%></td>
                <%end if%> 
        </tr> 
        </table>
	<table border="0" cellspacing="0" cellpadding="0" >
        <tr>
		<td><font size="2" >電郵地址</formt></td>
		<td width=54></td>
		<td><input type="text" name="mememail" value="<%=mememail%>" size="50" maxlength="50"></td>
        </tr> 

        <TR>
	     
		<td><font size="2" >備註欄</formt></td>
		<td width=54></td>

		<td><textarea rows="5" name="remark" cols="65"><%=remark%></textarea></td>
        </TR> 
        </table>
        </td>
	<td width="1%" valign="top">
        <td> 
      	<td width="40%" valign="top">
 		<table border="0" cellspacing="0" cellpadding="0" >
                <tr> 
                     <td><font size="2" >居住地址</formt></td>             
                    
                     <td><input type=" text" name="memAddr1" value="<%=memAddr1%>" size="40" maxlength="40"></td>
                </tr>

                <tr> 
                     <td><font size="2" ></formt></td>             
		  
                     <td><input type=" text" name="memAddr2" value="<%=memAddr2%>" size="40" maxlength="40"></td>
                </tr>
                <tr> 
                     <td><font size="2" ></formt></td>             
		    
                     <td><input type=" text" name="memAddr3" value="<%=memAddr3%>" size="40" maxlength="40"></td>
                </tr>
          
                <tr>
                     <td><font size="2" >住址電話</formt></td>  
                     
                      <td><input type="text" name="memContactTel" value="<%=memContactTel%>" size="10" ></td> 
                     
          
                </tr>
               <tr> 
                     <td><font size="2" >辦公地址</formt></td>             
                    
                     <td><input type=" text" name="memoadd1" value="<%=memOadd1%>" size="40" maxlength="40"></td>
                </tr>
                <tr> 
                     <td><font size="2" ></formt></td>             
		  
                     <td><input type=" text" name="memoadd2" value="<%=memoadd2%>" size="40" maxlength="40"></td>
                </tr>
                <tr> 
                     <td><font size="2" ></formt></td>             
		    
                     <td><input type=" text" name="memoadd3" value="<%=memoadd3%>" size="40" maxlength="40"></td>
                </tr>
          
                <tr>
                     <td><font size="2" >辦公電話</formt></td>  
                     
                      <td><input type="text" name="memofficetel" value="<%=memofficetel%>" size="10" ></td> 
                     
          
                </tr>
                <tr>
                     <td><font size="2" >聯絡電話</formt></td>  
                     
                      <td><input type="text" name="memMobile" value="<%=memMobile%>" size="10" ></td> 
                     
               </tr>
                     <tr>
                     <td><font size="2" >介紹人 1.</formt></td>                       
                     <td><input type="text" name="introd1" value="<%=introd1%>" size="35" ></td>                                
                     </tr>
                     <tr>
                     <td><font size="2" >介紹人 2.</formt></td>                       
                     <td><input type="text" name="introd2" value="<%=introd2%>" size="35" ></td>                                
                     </tr>          
                 
                </table>  
        </td>  
</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" >
        <td width="1%" valign="top">
         </td>  
	<td width="48%" valign="top">
 		<table border="0" cellspacing="0" cellpadding="0" >
       		 <TR>
 	
		<td><font size="2" >承繼人 1.</formt></td>  
		<td width=10</td>
		<td><input type="text" name="B1" value="<%=B1%>" size="20" maxlength="20"></td>
                </tr>
                <tr>
		<td><font size="2" >身分證號碼</formt></td>  
		<td width=10></td>
		<td><input type="text" name="B1ID" value="<%=B1ID%>" size="20" maxlength="20"></td>    
                </tr>
                <tr>  
		<td><font size="2" >關係</formt></td>  
		<td width=10></td>
		<td><input type="text" name="B1relation" value="<%=B1relation%>" size="20" maxlength="20"></td>    
                </TR>
                <tr>
		<td><font size="2" >地址</formt></td>  
		<td width=10></td>
		<td><input type="text" name="B1ADD" value="<%=B1ADD%>" size="35" maxlength="35"></td>
                </tr>
                <tr>
		<td><font size="2" ></formt></td>  
		<td width=75></td>
		<td><input type="text" name="B1ADD2" value="<%=B1ADD2%>" size="35" maxlength="35"></td>
                </tr>  
                </table>
        </td>
        <td width="2%" valign="top">
        </td>
	<td width="48%" valign="top">
 		<table border="0" cellspacing="0" cellpadding="0" >
      		 <TR>
 	
		<td><font size="2" >承繼人 2.</formt></td>  
		<td width=18></td>
		<td><input type="text" name="B2" value="<%=B2%>" size="20" maxlength="20"></td>
                </tr>
                <tr>
		<td><font size="2" >身分證號碼</formt></td>  
		<td width=20></td>
		<td><input type="text" name="B2ID" value="<%=B2ID%>" size="20" maxlength="20"></td>    
                </tr>
                <tr>  
		<td><font size="2" >關係</formt></td>  
		<td width=20></td>
		<td><input type="text" name="B2relation" value="<%=B2relation%>" size="20" maxlength="20"></td>    
                </TR>
                <tr>
		<td><font size="2" >地址</formt></td>  
		<td width=75></td>
		<td><input type="text" name="B2ADD" value="<%=B2ADD%>" size="35" maxlength="35"></td>
                </tr>
                <tr>
		<td><font size="2" ></formt></td>  
		<td width=75></td>
		<td><input type="text" name="B2ADD2" value="<%=B2ADD2%>" size="35" maxlength="35"></td>
                </tr>
                </table>
        </td>
</table>

</form>
</body>
</html>
