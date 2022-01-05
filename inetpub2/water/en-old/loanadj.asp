<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
  xyr = year(date()) - 1 
  todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
  YYdate  = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&xyr


   check=""
  
   key = 0
if request.form("bye") <> "" then
   response.redirect "main.asp"   
end if
if request.form("cancel") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next

	xdate = right(chkdate,4)&"/"&mid(chkdate,4,2)&"/"&left(chkdate,2)
        ydate = right(adjdate,4)&"/"&mid(adjdate,4,2)&"/"&left(adjdate,2) 
	if msg="" then
                
		select case xkind
                       case "Cash"
 
                             if adjpamt <> "" then 
			          addUserLog "刪除 現金帳 貸款編號 : "&lnnum&" 日期 : "&chkdate&" (本金) $ "&pamt 
                                  conn.begintrans  
                                  conn.execute("delete loan  where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='E3' ")
                                  conn.execute("update loanrec set bal = bal + "&pamt&"  where lnnum='"&lnnum&"' ")
                                  conn.execute("update loanrec set cleardate=null  where  lnnum='"&lnnum&"' and bal>0  ")
                                  conn.execute("update loanrec set repaystat='N'  where  lnnum='"&lnnum&"' and bal>0  ")                                     
                                  conn.execute("update loan set pflag=1  where lnnum='"&lnnum&"' and pflag=0 and code='ME' and pdate='"&xdate&"'  ")             
                                  conn.committrans
                            end if
                            if adjpint <> "" then
                               addUserLog "刪除 現金帳 貸款編號 : "&lnnum&" 日期 : "&chkdate&" (利息) $ "&pint
                               conn.begintrans
                               conn.execute("update loan set pflag=1  where lnnum='"&lnnum&"' and pflag=0 and code='MF' and pdate='"&xdate&"'  ")             
                               conn.execute("delete loan  where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='F3' ") 
                               conn.committrans 
                            end if  
 
                            
                            if adjsamt <>"" then
                                
                            addUserLog "刪除 股金退款帳  日期 : "&chkdate&" (金額) $ "&samt&"  至 "&adjsamt
                            conn.begintrans
                            conn.execute("delete  share where memno='"&memno&"' and code='A3' and ldate='"&xdate&"' ") 
                            conn.execute("update share set pflag=1  where memno='"&memno&"' and pflag=0 and code='AI' and pdate='"&xdate&"'  ")               
                            conn.committrans  
                           end if
		       case "Share"
                            if ADJpamt <> "" then   
			        addUserLog "刪除 股金帳 貸款編號 : "&lnnum&" 日期 : "&chkdate&"(本金) $ "&pamt
                                 conn.begintrans
                                conn.execute("delete loan  where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='ET' ")
                                conn.execute("update loanrec set bal = bal + "&pamt&" where lnnum='"&lnnum&"' ")
                                conn.execute("update loanrec set cleardate=null  where  lnnum='"&lnnum&"' and bal>0  ")
                                conn.execute("update loanrec set repaystat='N'  where  lnnum='"&lnnum&"' and bal>0  ")                                     
                                conn.execute("update loan set pflag=1  where lnnum='"&lnnum&"' and pflag=0 and code='ME' and pdate='"&xdate&"'  ")             
                                conn.committrans
                            end if
                            if ADJpint <> "" then 
                               addUserLog "刪除 股金帳 貸款編號 :"&lnnum&" 日期 : "&chkdate&" (利息) $ "&pint

                               conn.begintrans 
                               conn.execute("delete  loan  where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='FT' ") 
                               conn.execute("update loan set pflag=1  where lnnum='"&lnnum&"' and pflag=0 and code='ME' and pdate='"&xdate&"'  ")             
                               conn.committrans 
                            end if 
    '                       if adjdate  <>"" then
    '                           addUserLog "刪除 股金帳 貸款編號 :"&lnnum&" 日期 : "&chkdate
    '                           conn.execute("delete loan  where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='F3' ") 
    '                        end if  
                           savettl = adjpamt+adjpint
                           addUserLog "刪除 股金退款帳  日期 : "&chkdate&" (金額) $ "&actttl 
                           conn.begintrans
                            
                            conn.execute("delete  share  where memno='"&memno&"' and code='B0' and ldate='"&xdate&"' ")
                            conn.committrans
                       case "Back"
                            if ADJPAMT <> "" then
			      addUserLog "刪除 退款帳 貸款編號 : "&lnnum&" 日期 : "&chkdate&" (本金) $ "&pamt
                              conn.begintrans
                               conn.execute("delete loan where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='E0' ")
                               conn.execute("update loanrec set bal = bal - "&pamt&" where lnnum='"&lnnum&"' ")
                               conn.execute("update loanrec set cleardate='"&xdate&"'  where  lnnum='"&lnnum&"' and bal=0  ")
                               conn.execute("update loanrec set repaystat='C'  where  lnnum='"&lnnum&"' and bal=0  ")                                     
                               conn.committrans
                                  
                            end if
                            if ADJpint <> ""  then 
                                adduserlog "刪除 退款帳 貸款編號 :"&lnnum&" 日期 : "&chkdate&"(利息) $ "&pint
                                conn.begintrans
                                conn.execute("delete loan where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='F0' ")
                                conn.committrans 
                             end if 

                            savettl = adjpamt+adjpint
                            actttl  = pamt + pint
                            addUserLog "刪除 股金退款帳  日期 : "&chkdate&"(金額) $ "&savettl                        
                            conn.begintrans  
                            conn.execute("delete share  where memno='"&memno&"' and code='B1' and ldate='"&chkdate&"' ") 
                            conn.committrans 
                 end select
		
                id =""

		msg = "紀錄已更新"
	end if
        id = ""
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next
  xyr = year(date()) - 1 
  todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
  YYdate  = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&xyr
    id = ""
   memno=""
   memcname=""
   memname=""
   key = 0
           

end if
if request.form("clrScr") <> "" then
   id = ""
   memno=""
   memcname=""
   memname=""
   key = 0
   
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
     
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next

		sql = "select * from memMaster where memNo='"& memNo & "' "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if not rs.eof then

			For Each Field in rs.fields
			
					TheString = Field.name & "= rs(""" & Field.name & """)"
			
				Execute(TheString)
			Next

                id = memno
                rs.close
                todate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
              

   memno = id
   pint = 0
   pamt  = 0 
   intamt = 0
   mstatus=""

                else
                  msg = "社員號碼不存在 "

                end if
         
end if





if request.form("Srh2")<>"" then

        msg=""
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next

        xdate = right(chkdate,4)&"/"&mid(chkdate,4,2)&"/"&left(chkdate,2)
  
        select case kind
               case "現金"
                     KEY =1
                     xkind = "Cash"
                     cond=" and code='A3' "
               case "股金還款"
                     KEY = 2 
                     xkind = "Share" 
                     cond = "  and code='B0' " 
               case "退款"
                     KEY = 3 
                     xkind = "Back"
                     cond = " and code = 'B1' "	
               case "調整"
                     KEY = 4 
                     xkind = "Adj"
                     cond = " and code = 'E7' "		    
         end select   
         samt = 0
           sql  = "select * from share where memno='"&memno&"' and ldate ='"&xdate&"' "
           set rs = conn.execute(sql) 
           do while  not rs.eof  
              select case rs("code")     
                     case "A3"                       
                           samt = samt + rs("amount")
                           if key = 1  then
                               adjsamt = samt
                           end if
                     case "B0"                       
                           samt = samt + rs("amount")
                           if key = 2  then
                               adjsamt = samt
                           end if               
                     case "A7"                       
                           samt = samt + rs("amount")
                           if key = 4  then
                               adjsamt = samt
                           end if             
                    case "B0"
                          if rs("amount") < 0 then
                             samt = samt + rs("amohnt")
                             if key = 3 then
                                adjamt = samt
                             end if
                          end if
                   end select 
           rs.movenext
           loop
          rs.close
     
        adjpamt = 0
        adjpint = 0
        sql =  "select * from loan where memno='"&memno&"' and ldate ='"&xdate&"'  "
        set rs=conn.execute(SQL)
        LNNUM = ""
       if not rs.eof  then   
           lnnum =rs("lnnum")
           do while not rs.eof 
              
              select case rs("code")
                     case "E3"
                          IF KEY = 1 THEN
                          pamt = rs("amount")
                          adjpamt = pamt
                          END IF
                     case "F3"
                          IF KEY = 1 THEN
                          pint = rs("amount")
                          adjpint = pint  
                          END IF 
                     case "E7"
                          IF KEY = 4 THEN
                          pamt = rs("amount")
                          adjpamt = pamt
                          END IF
                     case "F7"
                          IF KEY = 4 THEN
                          pint = rs("amount")
                          adjpint = pint  
                          END IF 
                     case "E0" 
                           
                           if( rs("amount") >  0 and key =2) then
                             pamt = rs("amount")
                             adjpamt = pamt
                          END IF
                          if  (rs("amount")<0 and key=3 )  then
                             pamt = rs("amount")*-1
                             adjpamt = pamt
                          END IF                          
                     case "F0" 
                           if( rs("amount") >  0 and key =2) then
                             pint = rs("amount")
                             adjpint = pint
                          END IF
                          if  (rs("amount")<0 and key=3 )  then
                             pint = rs("amount")*-1
                             adjpint = pint
                          END IF                        
              end select
           rs.movenext  
           loop
           rs.close
           samt = 0


           IF PAMT = 0 and samt=0 THEN
               LNNUM = ""
               msg = "貸款帳項找不到1 "
            END IF
         
            adjdate = chkdate
          
                  
        else

               msg = "貸款帳項找不到2 "
	
        end if   	
 
end if



if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next

	xdate = right(chkdate,4)&"/"&mid(chkdate,4,2)&"/"&left(chkdate,2)

        ydate = right(adjdate,4)&"/"&mid(adjdate,4,2)&"/"&left(adjdate,2) 
	if msg="" then
                
		select case xkind
                       case "Cash"
                             if adjpamt <> "" then 
			          addUserLog "修改 現金帳 貸款編號 : "&lnnum&" 日期 : "&chkdate&" (本金) $ "&pamt&" 至 "&adjpamt 
                                  conn.begintrans
                                  conn.execute("update loan set amount="&adjpamt&"  where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='E3' ")
                                  conn.execute("update loanrec set bal = bal + "&pamt&" - "&adjpamt&" where lnnum='"&lnnum&"' ")
                                  conn.committrans                                   
                            end if
                            if adjpint <> "" then

                               addUserLog "修改 現金帳 貸款編號 : "&lnnum&" 日期 : "&chkdate&" (利息) $ "&pint&" 至 "&adjpint
                               conn.begintrans 
                               conn.execute("update loan set amount="&adjpint&"  where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='F3' ") 
                               conn.committrans
                            end if  

                            if adjsamt <>"" then
                            
                            addUserLog "修改 股金現金帳  日期 : "&chkdate&" (金額) $ "&samt&"  至 "&adjamt
                            conn.begintrans
                            conn.execute("update share set amount = "&adjsamt&" where memno='"&memno&"' and code='A3' and ldate='"&xdate&"' ") 
                            conn.committrans
                            end if
		       case "Share"
                            if ADJpamt <> "" then   
			        addUserLog "修改 股金還款帳 貸款編號 : "&lnnum&" 日期 : "&chkdate&"(本金) $ "&pamt&" 至 "&adjpamt
                                conn.begintrans
                                conn.execute("update loan set amount="&ADJpamt&"  where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='E0' ")
                                conn.execute("update loanrec set bal = bal + "&pamt&" - "&adjpamt&" where lnnum='"&lnnum&"' ")
                                conn.committrans  
                            end if
                            if ADJpint <> "" then 
                               addUserLog "修改 股金還款帳 貸款編號 :"&lnnum&" 日期 : "&chkdate&" (利息) $ "&pint&" 至 "&adjpint
                               conn.begintrans
                               conn.execute("update loan set amount="&ADJpint&"  where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='F0' ") 
                               conn.committrans  
                            end if 
    '                       if adjdate  <>"" then
    '                           addUserLog "修改 股金還款帳 貸款編號 :"&lnnum&" 日期 : "&chkdate&" 至 "&adjdate

    '                           conn.execute("update loan set ldate="&ydate&"  where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='F3' ") 
    '                        end if  
                           savettl = cint(adjpamt)+(adjpint)
                           actttl  = cint(pamt) + (pint)
                            addUserLog "修改 股金還款帳 日期 : "&chkdate&" (金額) $ "&actTtl&"  至 "&savettl
                            conn.begintrans 
                            conn.execute("update share set amount = "&savettl&" where memno='"&memno&"' and code='B0' and ldate='"&xdate&"' ")
                            conn.committrans
                       case "Back"
                            if ADJPAMT <> "" then
			      addUserLog "修改 退款帳 貸款編號 : "&lnnum&" 日期 : "&chkdate&" (本金) $ "&pamt&" 至 "&adjpamt
                               conn.begintrans                              
                               conn.execute("update loan set amount="&ADJpamt*-1&" where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='E0' ")
                               conn.execute("update loanrec set bal = bal + "&pamt&" - "&adjpamt&" where lnnum='"&lnnum&"' ")
                               conn.committrans   
                            end if
                            if ADJpint <> ""  then 
                                adduserlog "修改 退款帳 貸款編號 :"&lnnum&" 日期 : "&chkdate&"(利息) $ "&pint&" 至 "&adjpint
                                conn.begintrans
                                conn.execute("update loan set amount="&ADJpint*-1&" where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='F0' ")
                                conn.committrans
                             end if 
   '                        if adjdate  <>"" then
   '                            addUserLog "修改 退款帳 貸款編號 : "&lnnum&" 日期 : "&chkdate&" 至 "&adjdate
   '                            conn.execute("update loan set ldate="&ydate&"  where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='F3' ") 
   '                         end if 
                            savettl = cint(adjpamt)+cint(adjpint)
                            actttl  = cint(pamt) + cint(pint)
                            addUserLog "修改 股金退款帳  日期 : "&chkdate&"(金額) $ "&actTtl*-1&"  至 "&savettl*-1                        
                            conn.execute("update share set amount = "&savettl*-1&" where memno='"&memno&"' and code='B1' and ldate='"&xdate&"' ") 
                            conn.committrans
                      case "Adj"
                            if ADJPAMT <> "" then
			      addUserLog "修改 調整帳 貸款編號 : "&lnnum&" 日期 : "&chkdate&" (本金) $ "&pamt&" 至 "&adjpamt
                              conn.begintrans                               
                              conn.execute("update loan set amount="&ADJpamt&" where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='E7' ")
                              conn.execute("update loanrec set bal = bal + "&pamt&" - "&adjpamt&" where lnnum='"&lnnum&"' ")
                              conn.committrans    
                            end if
                            if ADJpint <> ""  then 
                                adduserlog "修改 調整帳 貸款編號 :"&lnnum&" 日期 : "&chkdate&"(利息) $ "&pint&" 至 "&adjpint
                                conn.begintrans
                                conn.execute("update loan set amount="&ADJpint&" where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='F7' ")
                                conn.committrans  
                             end if 

                            savettl = cint(adjpamt)+cint(adjpint)
                            actttl  = cint(pamt) + cint(pint)
                            addUserLog "修改 股金調整帳  日期 : "&chkdate&"(金額) $ "&actTtl&"  至 "&savettl
                           conn.begintrans
                           conn.execute("update share set amount = "&savettl&" where memno='"&memno&"' and code='A7' and ldate='"&xdate&"' ") 
                           conn.committrans  
                end select
		
                id =""

		msg = "紀錄已更新"
	end if
        id = ""
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next
  xyr = year(date()) - 1 
  todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
  YYdate  = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&xyr
    id = ""
   memno=""
   memcname=""
   memname=""
   key = 0
           

 
       

end if
%>
<html>
<head>
<title>貸款細項修正</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">

<script language="JavaScript">
<!--
function popup(filename){
  window.open (filename,'pop','width=500,height=550,statusbar=no,toolbar=no,resizable,scrollbars,dependent')
}

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

      Mn = strM
      Yr = strY
      if (((Mn<=sMn)&&(Yr==sYr))||(Yr<sYr)){
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

	if (formObj.adjpamt.value==""){
		reqField=reqField+",  修正貸款本金 ";
		if (!placeFocus)
			placeFocus=formObj.adjpamt;
	}

	if (formObj.adjpint.value==""){
		reqField=reqField+",  修正貸款利息 ";
		if (!placeFocus)
			placeFocus=formObj.adjpint;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.memNo.focus()">
<!-- #include file="menu.asp" -->
<%if msg<>"" then %>
<div><center><font size="3" color="red"><%=msg%></font></center></div>
<% end if%>

<br>
<form name="form1" method="post" action="loanadj.asp">

<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="check" value="<%=check%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="YYdate" value="<%=YYdate%>">
<input type="hidden" name="xstatus" value="<%=xstatus%>">
<input type="hidden" name="memName" value="<%=memName%>">
<input type="hidden" name="memcName" value="<%=memcName%>">
<input type="hidden" name="pamt" value="<%=pamt%>">
<input type="hidden" name="pint" value="<%=pint%>">
<input type="hidden" name="key" value="<%=key%>">
<input type="hidden" name="xkind" value="<%=xkind%>">
<input type="hidden" name="cutdate" value="<%=cutdate%>">
<input type="hidden" name="lastmonth" value="<%=lastmonth%>">
<input type="hidden" name="lastyear" value="<%=lastyear%>">
<input type="hidden" name="spass" value="<%=spass%>">
<div><center><font size="3" >貸款細項修正</font></center></div>
<BR>
<center>
<table border="0" cellspacing="0" cellpadding="0">

			<tr>
               		<td width=30></td>
			<td class="b12" align="left">社員號碼</td>
			<td width=50></td>
			<td><input type="text" name="memNo" value="<%=memNo%>" size="10" <%if id<>"" then response.write " onfocus=""form1.chkdate.focus();""" end if%>> 			
                        <%if id = "" then %>
		        <input type="button" value="選擇"  onclick="popup('pop_srhMemnoM.asp')" class="sbttn"  >
		        <input type="submit" value="搜尋" name="Search" class ="Sbttn">
 		        <%end if%>            
                        </TD>
                        </tr>                        			
			<tr>
               		<td width=30></td>
			<td class="b12" align="left">社員名稱</td>
			<td width=50></td>
               		<td id="memName"><%=memName%></td>
 	        	         
                        </tr>
               		<td width=30></td>
			<td class="b12" align="left"></td>
			<td width=50></td>
               		
                        <td id="memcName"><%=memcName%></td>
                        </tr>
                        
                        <tr>
			<td width=30></td>
			<td class="b12" align="left">修改日期</td>
			<td width=50></td>
                   
			<td><input type="text" name="chkdate" value="<%=chkdate%>" size="10" onblur="if(!formatDate(this)){this.value=''};form1.repaydate.value=this.value"></TD>     
                   
                        </tr>   
                            
			<tr>
			<td width=30></td>
			<td align="right" class="b12">貸款項目</td>
			<td width="50"></td>
			
                        <%if lnnum="" then %>
                        <td>  
			<select name="KIND" style="width:88px">		
			<option<% if KIND="Cash"  then response.write " selected" end if%>>現金
			<option<% if KIND="Share" then response.write " selected" end if%>>股金還款
                        <option<% if KIND="Back"  then response.write " selected" end if%>>退款	
                        <option<% if KIND="Adj"  then response.write " selected" end if%>>調整			
			</select>
                        <%if id <>"" then %>
                         
                        <input type="submit" value="搜尋" name="Srh2" class ="Sbttn">
 			<input type="submit" value="取消" name="clrScr" class ="Sbttn">
			<input type="submit" value="返回" name="bye" class ="Sbttn">
                        <input type="button" value="查詢貸款" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.YYdate.value )" class="sbttn">                       
                        <%end if %>
                        </td>  
                        <%else%>
                        <td  id = "kind"><%=kind%></td>
                        <%end if %> 

		        </tr>
                       
			<tr>
			<td width=30></td>
			<td align="right" class="b12">貸款編號</td>
			<td width="50"></td>
			<td><input type="text" name="lnnum" value="<%=lnnum%>" size="10" <%if id<>"" then response.write " onfocus=""form1.adjdate.focus();""" end if%>>		
		        </td>

	                </tr>
			<tr>
			<td width=30></td>
			<td align="right" class="b12">貸款本金</td>
			<td width="50"></td>
			<td id="pamt"><%=pamt%>			
		        </td>
	                </tr>

			<tr>
			<td width=30></td>
			<td align="right" class="b12">貸款利息</td>
			<td width="50"></td>
			<td id="pint"><%=pint%>			
		        </td>
	                </tr>
                        <%if samt > 0 then%>
			<tr>
			<td width=30></td>
			<td align="right" class="b12">貸款股金</td>
			<td width="50"></td>
			<td id="samt"><%=samt%>			
		        </td>
	                </tr>
                        <%end if %>
			<%if lnnum<> ""  or samt > 0  then %>

			<tr>
			<td width=30></td>
			<td align="right" class="b12">更正本金</td>
			<td width="50"></td>
			<td><input type="text"  name="adjpamt"  value="<%=adjpamt%>" size="10" maxlength="10" >			
		        </td>
	                </tr>
			<tr>
			<td width=30></td>
			<td align="right" class="b12">更正利息</td>
			<td width="50"></td>
			<td><input type="text" name="adjpint" value="<%=adjpint%>" size="10" maxlength="10" ></td>

                        </tr>
                        <%if key = 1 then %>
			<tr>
			<td width=30></td>
			<td align="right" class="b12">更正股金</td>
			<td width="50"></td>
			<td><input type="text" name="adjsamt" value="<%=adjsamt%>" size="10" maxlength="10" ></td>

                        </tr>  
                        <%end if %>
                        <tr>
                        <td></td>
                        <td></td>
                        <td></td>   
                        <td> 
        		<input type="submit" value="刪除" onclick="confirm('確定刪除?')" name="cancel" class="sbttn">
			<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
                        
			<input type="submit" value="取消" name="clrScr" class ="Sbttn">
			<input type="submit" value="返回" name="bye" class ="Sbttn">
		        </td>
	                </tr>
                        <%end if %>
</table>       
</center>
</form>
</body>
</html>
                                                                                    