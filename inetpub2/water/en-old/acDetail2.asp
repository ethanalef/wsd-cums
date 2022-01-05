<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->

<%

 userlevel = session("userlevel")
 username  = session("username")
if request.form("back") <> "" then
   response.redirect "main.asp"
   
end if

if request.form("Search")<>"" then
      

	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
     
        id = memno
        set rs = conn.execute("select memno,memname,memhkid,memcname,memGrade,employCond,membday,bnklmt,monthsave,monthssave,accode,tpayamt,mstatus,remark from memmaster where memno='"&memno&"' ")
        if not rs.eof then
	   		For Each Field in rs.fields
			if Field.name="memBday" or Field.name="firstAppointDate" or Field.name="memDate" or Field.name="Wdate" then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
                membday = right("0"&day(rs("membday")),2)&"/"&right("0"&month(rs("membday")),2)&"/"&year(rs("membday"))
                select case mstatus
                       case "L"
                           xstatus= "呆帳"
                           repayamt = 0
                       case  "D"
                           xstatus="冷戶"
                       	repayamt = 0
                       case  "V"
                           xstatus= " IVA "
                          repayamt = 0
                       case  "C"
                             xstatus= "退社"
             		   repayamt = 0	
                       case  "P"
                             xstatus= "去世"
                           repayamt = 0
                       case  "B"
                            xstatus="破產"
                          repayamt = 0
                       case  "N"
                            xstatus= "正常"
                            repayamt = monthsave    
                      case  "J"
                            xstatus= "新戶"
                            repayamt = monthsave
                      case "H"
                          xstatus= "暫停銀行"
                          repayamt = monthsave
                       case  "A"
                            xstatus="自動轉帳(ALL)"
			    repayamt = monthsave
                       case  "0"
                            xstatus="自動轉帳(股金)"                       
			    repayamt = monthsave	

                       case  "1"
                            xstatus="自動轉帳(股金,利息)"
                            repayamt = monthsave
                       case  "Z"
                            xstatus="自動轉帳(股金,本金)"
                            repayamt = monthsave
                       case  "3"
                            xstatus="自動轉帳(利息,本金)"
                            repayamt = monthsave

                       case  "M"
                           xstatus = "庫房,銀行"
                           repayamt = monthssave
                      case  "T"
                            xstatus= "庫房"
                            repayamt = monthssave   
                       case "F"
                              xstatus = "特別個案"  
			       repayamt = monthsave
                        case "8"
                             xstatus = "終止社籍轉帳"
                              repayamt = monthsave
                        case "9"
                             xstatus = "終止社籍正常"
                             repayamt = monthsave
                end select
 
           chkdate =right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1)
           if day(date()) = 29 and month(date())=2  then
              if (year(date())-1)/4 <> int((year(date())-1)/4) then
                 chkdate =  right("0"&(day(date())-1),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1)
              end if
           end if
            if accode<>"9999" then   
            set rs1=conn.execute("select  memcname,memname,memofficetel from memmaster where memno='"&accode&"' ")
            if not rs1.eof then
               xcname= rs1(0)
               xname =rs1(1)
                xcontel=rs1(2)
              end if
            rs1.close
           else
               xcname="工作人員" 
               xname =  ""
               xcontel = "27879222"
           end if
         if isdate(membday) then 
           xyr = right(membday,4)
           xmm = mid(membday,4,2)
           xdd = left(membday,2)
           mn  = month(date())
           yr  = year(date()) 
           bday = dateserial(yr,xmm,xdd)
           if bday >date() then
              pdate = dateserial(yr-1,xmm,xdd) 
              dd = int(datediff("d", pdate,date())/365.25*10)/10   
              age =year(date())  - year(membday)-1+dd
           else
                dd =  int(datediff("d", bday ,date())/365.25*10)/10    
                 age =year(date())  - year(membday)+dd
           end if 
       
          end if     

               
           set rs1 = conn.execute("select * from loanrec where memno='"&memno&"' and repaystat='N' ")
           if not rs1.eof then   
 	   		For Each Field in rs1.fields
			if Field.name="lndate" then
					TheString = "if rs1(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs1(""" & Field.name & """)),2)&""/""&right(""0""&month(rs1(""" & Field.name & """)),2)&""/""&year(rs1(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs1(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next 
                        if monthrepay="" then
                           monthrepay =iif( int(appamt/install)=appamt/install,int(appamt/install),int(appamt/install)+1)
                        end if
            xlndate=rs1("lndate")
           if mstatus="T" or mstatus="M" then            
              xmess = "每月儲蓄(庫房)"
              monthsamt = monthssave
              
           else
              xmess = "每月儲蓄(銀行)"
              monthsamt = monthsave
           end if                      
           else
                lnnum=""          
           end if 
          
           rs1.close
           ttlbal = 0
 

SQl = "select a.memno,a.memname,a.memcname, b.dttl ,b.cttl     from memmaster a "&_
      " right join ( select memno ,sum( case when left(code,1)  in  ( '0','A','C' ) and code<>'AI' and code<>'CH' "&_
      "  then amount else 0 end ) as dttl , sum( case when left(code,1)  in  ( 'G','H','B' ) and code<>'MF' "&_
      "  then amount else 0 end ) as cttl from share where memno= '"&memno&"' group by memno ) b on a.memno=b.memno "&_
     "  where a.memno = '"&memno&"'  and  a.mstatus not in ('C','P','B','I' ) and a.wdate is null  order by a.memno   "
          
           set rs1 = conn.execute(sql)
           if  not rs1.eof  then      
              ttlbal = rs1(3) - rs1(4) 
           end if 
           rs1.close

              
           if lnnum <> "" then
           xintamt = 0
           yy = year(xlndate)
           mm = month(xlndate)
           dd = day(xlndate)
           xlndate=dateserial(yy,mm,dd)
           md = day(dateserial(yy,mm+1,1-1))
             if bal > monthrepay  then 
                 pamt1 = monthrepay
              else
                 pamt1 = bal
              end if
                pamt2 = 0
                pint2 = 0       
                pint1  = 0 
		set rs1  = conn.execute("select *  from loan where memno='"& memno & "'  and left(code,1)='D' and pflag= 1 ")				
                do while  not rs1.eof 
                   select case rs1("code")  
                          case "DE"  
                               pamt2 = pamt2 + rs1("amount")
                          case "DF"
                              
                                  pint2  = pint2 + rs1("amount")
                       
                                        
                   end select    
                rs1.movenext
                loop             
                rs1.close
                pamt = pamt2
                pint = pint2
                pint1 = 0
                xmd = day(dateserial(yy,mm+1,1-1))
                xmon = month(date()) - mm
                samt  = 0   
 
                select case xmon
                       case 1

                            if lnflag="Y" then
                              if appamt=bal then                                
                                 a1 = round((bal-chequeamt)*.01,2)
                                 a2 = round(chequeamt*.01*(xmD-dd+1)/xmD,2)
                                 a3 = round(bal *.01,2)
                                 pint2 = a1 + a2 + a3
                                   
                               else
                               a2 = 0
                               set ms = conn.execute("select * from loan where memno='"&memno&"' and ldate>='"&xlndate&"' and code='E1' ")
                               if not ms.eof then
                                  if ms("amount")<> monthrepay then
                                     a2 = round(chequeamt*.01*(xmD-dd+1)/xmD,2)
                                     xintamt = a2
                                  end if
                              end if
                              ms.close
                                  a3 = round(bal *.01,2)
                                  pint2 = a3 + a2
                  
                               end if
                            else
                               if appamt=bal then
                                 a1 = round(bal*.01*(xmd-dd+1)/xmd,2)
                                 a2 = round(bal *.01,2)
                                 pint2 = a1 + a2
                               else                                 
                                 a2 = round(bal *.01,2)
                                 pint2 = a2
                               end if
                            end if
                         case 0
                               if lnflag = "Y" then                                                                  
                                  a1 = round((bal-chequeamt)*.01,2)
                                  a2 = round(chequeamt*.01*(xmD-dd+1)/xmD,2)                                 
                                  xintamt = a2
                                  if xlndate > dateserial(yy,mm,24) then
                                     pint2  = a2   
                                  else
                                     pint2 = a1 + a2 
                                  end if
                                 
                              else
                                  a2 = round(chequeamt*.01*(xmD-dd+1)/xmD,2)                                 
                                 pint2 =  a2                                     
                               end if
                         case else
                              pint2 = round(bal*0.01,2)
                end select
               ttlpamt = pamt1 + pamt2
               ttlpint = pint1 + pint2
                ttlamt = ttlpint+ttlpamt
              


                select case mstatus
                       case "A","N","0","1","2"
                             intamt = ttlpint
                             pincamt = ttlpamt
                             monthsamt = monthsave +samt 
                              ttlamt = intamt + pincamt + monthsamt
                             if ttlamt >=bnklmt and bnklmt > 0 then
                                ttlamt = bnklmt
                                if ttlamt > intamt  then
                                   xamt1 = ttlamt - intamt 
                                   if xamt1 > pincamt then
                                      xamt2 = xamt1 - pincamt
                                      if xamt2 < monthsave then
                                         monthsamt = xamt2
                                       else
                                           monthamt = monthsave
                                       end if
                                  else
                                        pincamt = xamt1
                                        monthsamt = 0
                                  end if  
                             end if 
                          end if
                      case "T"
                             ttlamt = tpayamt
                             intamt = ttlpint
                             pincamt = tpayamt - ttlpint
                      case "M"
                            intamt = ttlpint
                             pincamt = ttlpamt
                             monthsamt = monthssave
                              ttlamt = intamt + pincamt + monthsamt
                      case else
                           intamt =  0
                             pincamt = 0
                             monthsamt = 0     
                            ttlamt = 0                      
                end select      
              set rs1=conn.execute("select a.* from guarantor a ,loanrec b where a.lnnum='"&lnnum&"' and a.lnnum=b.lnnum and b.repaystat='N' ")
              xx = 1
              do while not rs1.eof
                 select case xx
                        case  1 
                             guid1 = rs1("guarantorID")
                             guname1 = rs1("guarantorname")
                             gucname1 = rs1("guarantorcname")
                             gattlbal = 0
                             set rs2 = conn.execute("select * from share where memno='"&guid1&"' and code not in ('CH','AI')  ")
			     do while not rs2.eof
		              select case left(rs2("code"),1)
                		     case "A","0","C"
                         		 gattlbal = gattlbal + rs2("amount")
		                     case "B","G","H","M"
                		           gattlbal = gattlbal - rs2("amount")
		               end select
                      
           			rs2.movenext
		            loop
		            rs2.close
                            gusave1 = gattlbal
                       case  2
                             guid2 = rs1("guarantorID")
                             guname2 = rs1("guarantorname")
                             gucname2 = rs1("guarantorcname")
                             gattlbal = 0
                             set rs2 = conn.execute("select * from share where memno='"&guid2&"' and code not in ('CH','AI')  ")
			     do while not rs2.eof
		              select case left(rs2("code"),1)
                		     case "A","0","C","M"
                         		 gattlbal = gattlbal + rs2("amount")
		                     case "B","G","H","M"
                		           gattlbal = gattlbal - rs2("amount")
		               end select
                      
           			rs2.movenext
		            loop
		            rs2.close
                            gusave2 = gattlbal
                       case 3             
                             guid3 = rs1("guarantorID")
                             guname3 = rs1("guarantorname")
                             gucname3 = rs1("guarantorcname")
                             gattlbal = 0
                             set rs2 = conn.execute("select * from share where memno='"&guid3&"' and code not in ('CH','AI')  ")
			     do while not rs2.eof
		              select case left(rs2("code"),1)
                		     case "A","0","C","M"
                         		 gattlbal = gattlbal + rs2("amount")
		                     case "B","G","H"
                		           gattlbal = gattlbal - rs2("amount")
		               end select
                      
           			rs2.movenext
		            loop
		            rs2.close
                            gusave3 = gattlbal
                   end select
                    xx = xx + 1
              rs1.movenext
              loop 
              rs1.close   
          else
             
             

          end if
          set rs1=conn.execute("select a.* from guarantor a,loanrec b where a.guarantorid='"&memno&"' and a.lnnum=b.lnnum and b.repaystat='N' ")
          xx = 1
          do while not rs1.eof
             select case xx
                    case 1          
                        guoid1 = rs1("memno")
                        guoln1 = rs1("lnnum")
                        set rs2 = conn.execute("select memno,memname,memcname from memmaster where memno='"&guoid1&"' ")
                        if not rs2.eof then
                           guoname1 = rs2("memname")
                           guocname1 = rs2("memcname")
                           
                        end if
                        rs2.close
                    case 2
                        guoid2 = rs1("memno")
                        guoln2 = rs1("lnnum")
                        set rs2 = conn.execute("select memno,memname,memcname from memmaster where memno='"&guoid2&"' ")
                        if not rs2.eof then
                           guoname2 = rs2("memname")
                           guocname2 = rs2("memcname")
			   
                        end if
                        rs2.close
                    case 3    
			 guoid3 = rs1("memno")
		         guoln3 = rs1("lnnum")
                        set rs2 = conn.execute("select memno,memname,memcname from memmaster where memno='"&guoid3&"' ")
                        if not rs2.eof then
			  	
                           guoname3 = rs2("memname")
                           guocname3 = rs2("memcname")
			   
                        end if
                        rs2.close
             end select
                        guoid1 = rs1("memno")
                        guoln1 = rs1("lnnum")
             xx = xx + 1   
          rs1.movenext
          loop
          rs1.close      

                cleardate = ""     
                todate    = cleardate
                cashint = 0
                cashamt = 0
                ttlpamt = 0
        else
             msg ="社員編號不存在!"  
             memno = ""  
             tpayamt = 0
             repayamt = 0               
        end if
        rs.close

else
   memno = ""
   memcName =""
   memName  = ""
   chkdate = ""
   saveamt  = 0.00
   id = ""   
end if
%>
<html>
<head>
<title>個人帳查詢</title>
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

function calculation(){
	formObj=document.form1;

 
	if (formObj.cleardate.value!=""){
           
            if (formObj.pint.value!=""){
               ppint1 = parseFloat(formObj.pint.value)
            }else{
               ppint1 = 0
            }      
            if (formObj.pamt.value!=""){
               ppamt1 = parseFloat(formObj.pamt.value)
            }else{
               ppamt1 = 0
            }      
	   Payamt = parseFloat(formObj.bal.value)    
           Appamt = parseFloat(formObj.appamt.value)  
           xintamt = parseFloat(formObj.xintamt.value)    
           chequeamt  = parseFloat(formObj.chequeamt.value)   
           lnflag  = formObj.lnflag.value 
           ssdate = formObj.cleardate.value
           ttdate = formObj.todate.value 
           xxdate = formObj.lndate.value 

	   Months= new Array(31,28,31,30,31,30,31,31,30,31,30,31)

	   Y=ssdate.substr(6,4)           
           M=ssdate.substr(3,2)
	   D=ssdate.substr(0,2)

	   XY=xxdate.substr(6,4)           
           XM=xxdate.substr(3,2)
	   XD=xxdate.substr(0,2)
           if (Y==XY){
              Mon = M - XM - 1
           }else{
              DY =parseInt( (Y- XY)*12)
              Mon = parseInt(M) + parseInt(DY)-   parseInt(XM) - 1   
           }   
           mD=Months[M-1]
           xmD=Months[XM-1]

           if ((Y==XY)&&(M==XM)){
              if (lnflag=='Y'){
                  a1 = parseInt((Payamt - chequeamt)*.01*D/mD*100)/100
                  a2 = parseInt(chequeamt*.01*(D-XD+1)/mD*100)/100
                  if (XD>=24){
                     ppint3 = a2
                  }else{
                  ppint3 = parseInt((a1 + a2 )*100)/100
                  }
                  
              }else{
                 a2 = parseInt( Payamt*0.01*(D-XD+1)/mD*100)/100
                  ppint3 = parseInt((a2 )*100)/100                  
              }                              

               ttlint = parseInt((ppint1+ppint3)*100)/100          	
           }else{
             if (lnflag=='Y'){
                   a1 = 0
                   a2 = 0
                  if (Appamt==Payamt){
                     a1 = parseInt((Payamt - chequeamt)*.01*100)/100
                     a2 = parseInt(chequeamt*.01*(xmD-XD+1)/mD*100)/100
                  }else{
                     a2 = xintamt
                  }
                  a3 = parseInt(Payamt*.01*D/mD*100)/100 + parseInt(Payamt*.01*Mon*100)/100
                  ppint3 = parseInt((a1 + a2+ a3 )*100)/100
                  
              }else{
                  a1 = 0
                  if (Appamt==Payamt){
                     a1 = parseInt( Payamt*0.01*(xmD-XD+1)/mD*100)/100
                  }
                  a2 = parseInt(Payamt*.01*D/mD*100)/100 + parseInt(Payamt*.01*Mon*100)/100
                  ppint3 = parseInt((a2 )*100)/100                  
              }                                         
               ttlint = parseInt((ppint1+ppint3)*100)/100  
           }
          
           document.all.tags( "td" )['cashamt'].innerHTML  =Payamt
           document.all.tags( "td" )['cashint'].innerHTML = ttlint
           document.all.tags( "td" )['ttlpamt'].innerHTML =  Payamt +ttlint

         }

}


function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (!formatNum(formObj.salaryDedut)){
		reqField=reqField+", 庫房扣薪";
		if (!placeFocus)
			placeFocus=formObj.salaryDedut;
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
<DIV>
<!-- #include file="menu.asp" -->
<div align="right"><a href="membermod2.asp?id=<%=request("id")%>">社員資料修正</a>&nbsp;&nbsp;</div>
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
</center>
<br>

<div><center><font size="3">個人帳查詢</font></center></div>
<form name="form1" method="post" action="acDetail2.asp">
<table border="0" cellspacing="0" cellpadding="0">
<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="memName" value="<%=memName%>">
<input type="hidden" name="memcName" value="<%=memcName%>">
<input type="hidden" name="memGrade" value="<%=memGrade%>">
<input type="hidden" name="chequeamt" value="<%=chequeamt%>">
<input type="hidden" name="lnflag" value="<%=lnflag%>">
<input type="hidden" name="membday" value="<%=membday%>">
<input type="hidden" name="memhkid" value="<%=memhkid%>">
<input type="hidden" name="age" value="<%=age%>">
<input type="hidden" name="employCond" value="<%=employCond%>">
<input type="hidden" name="firstAppointDate" value="<%=firstAppointDate%>">
<input type="hidden" name="lnnum" value="<%=lnnum%>">
<input type="hidden" name="bal" value="<%=bal%>">
<input type="hidden" name="appamt" value="<%=appamt%>">
<input type="hidden" name="lndate" value="<%=lndate%>">
<input type="hidden" name="monthsave" value="<%=monthsave%>">
<input type="hidden" name="monthssave" value="<%=monthssave%>">
<input type="hidden" name="tpayamt" value="<%=tpayamt%>">
<input type="hidden" name="monthrepay" value="<%=monthrepay%>">
<input type="hidden" name="xstatus" value="<%=xstatus%>">
<input type="hidden" name="pamt" value="<%=pamt%>">
<input type="hidden" name="pint" value="<%=pint%>">
<input type="hidden" name="cashamt" value="<%=cashamt%>">
<input type="hidden" name="cashint" value="<%=cashint%>">
<input type="hidden" name="ttlpamt" value="<%=ttlpamt%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="repayamt" value="<%=repayamt%>">
<input type="hidden" name="remark" value="<%=remark%>">
<input type="hidden" name="repaystatus" value="<%=repaystatus%>">
<input type="hidden" name="intamt" value="<%=intamt%>">
<input type="hidden" name="pincamt" value="<%=pincamt%>">
<input type="hidden" name="ttlamt" value="<%=ttlamt%>">
<input type="hidden" name="monthsmt" value="<%=monthsamt%>">
<input type="hidden" name="xintamt" value="<%=xintamt%>">

	<tr>
		<td width="450" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
                                 
				<tr>
                                        <td width=10></td>
					<td width="100" class="b8" align="right">社員編號</td>
					<td width=10></td>
					<td width="250">
					<input type="text" name="memNo" value="<%=memNo%>" size="10" maxlength="10"  <%if id<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>>
					<%if id = "" then %>
					<input type="button" value="選擇"  onclick="popup('pop_srhMemnoM.asp')" class="sbttn"  >
					<input type="submit" value="搜尋" name="Search" class ="Sbttn">
                                        <input type="submit" value="返回" name="back" class="sbttn">
					<%else%>
			                <input type="submit" value="取消" name="bye" class="sbttn">
					<input type="submit" value="返回" name="back" class="sbttn">
					<% end if %>
					</td>
					<td></td>

	
				</tr>
				<tr height="22">
					<td width=10></td>
					<td class="b8" align="right">姓名</td>
					<td width=10></td>
					<td id="memName"><%=memName%><%=memcName%></td>
					<td ></td>
				</tr>
				<tr height="22">
					<td width=10></td>
					<td class="b8" align="right">身分證編號</td>
					<td width=10></td>
					<td id="memhkid%"><%=memhkid%></td>
				</tr>
				<tr height="22">
					<td width=10></td>
					<td class="b8" align="right">職位</td>
					<td width=10></td>
					<td id="memGrade"><%=memGrade%></td>
				</tr>
				<tr height="22">
					<td width=10></td>
					<td class="b8" align="right">出生日期</td>
					<td width=10></td>
					<td id="membday"><%=membday%></td>
				</tr>
				<tr height="22">
					<td width=10></td>
					<td class="b8" align="right">年齡</td>
					<td width=10></td>
					<td id="age"><%=age%></td>
				</tr>
				<tr height="22">
					<td width=10></td>
					<td class="b8" align="right">招聘條款</td>
					<td width=10></td>
					<td id="employCond"><%=employCond%></td>
                                </TR>
				<tr height="22">
					<td width=10></td>
					<td class="b8" align="right">狀況</td>
					<td width=10></td>
					<td id="xstatus"><%=xstatus%></td>
                                </TR>

                                </TR>
				<tr height="22">
					<td width=10></td>
					<td class="b8" align="right">聯絡人</td>
					<td width=10></td>
					<td id="xcname"><%=xname%>,<%=xcname%></td>
                                
                              </tr>
				<tr height="22">
					<td width=10></td>
					<td class="b8" align="right">備註</td>
					<td width=10></td>
					<td id="remark"><font size="2" color="red"><%=remark%></font></td>
                                </TR>
			</table>	
		    </td>
  		<td width="350" valign="top">
		<table border="0" cellspacing="0" cellpadding="0">                  
 		<tr height="22">
			<td width=10></td>
			<td class="b8" align="right">1.擔保人</td>
			<td width=10></td>			
			<td id="guid1"><%=guid1%></td>
			<td width=10></td>
                        <td id="guname1"><%=guname1%></td>
			<td id="gucname1"><%=gucname1%></td>
		</tr>
                       <tr>
 			<td width=10></td>
			<td class="b8" align="right">儲蓄結餘</td>
			<td width=10></td>  
 			<% if gusave1 > 0 then  %>                     
                        <td id="gusave1"><%=formatNumber(gusave1,2)%></td>                        
                        <%end if %> 
                        </tr>
			<tr height="22">
			<td width=10></td>
			<td class="b8" align="right">2.擔保人</td>
			<td width=10></td>
			<td id="guid2"><%=guid2%></td>
			<td width=10></td>
                        <td id="guname2"><%=guname2%></td>
			<td id="gucname2"><%=gucname2%></td>
			</tr>
                       <tr>
 			<td width=10></td>
			<td class="b8" align="right">儲蓄結餘</td>
			<td width=10></td>   
                        <%if gusave2 > 0 then %>                     
                        <td id="gusave2"><%=formatNumber(gusave2,2)%></td>  
                        <% end if %> 
                        </tr>
			<tr height="22">
			<td width=10></td>
			<td class="b8" align="right">3.擔保人</td>
			<td width=10></td>
			<td id="guid3"><%=guid3%></td>
                       <td width=10></td> 
                        <td id="guname3"><%=guname3%></td>
			<td id="gucname3"><%=gucname3%></td>
			</tr>
                       <tr>
 			<td width=10></td>
			<td class="b8" align="right">儲蓄結餘</td>
			<td width=10></td>   
                        <% if gusave3 > 0 then %>                   
                        <td id="gusave3"><%=formatNumber(gusave3,2)%></td>   
                        <% end if %>
                        </tr> 

                </table>
                </td>
  		<td width="400" valign="top">
		<table border="0" cellspacing="0" cellpadding="0">    
<%if guoid1<>"" then %>              
 		<tr height="22">
			<td width=10></td>
			<td class="b8" align="right">1.擔保其他人</td>
			<td width=5></td>			
			<td id="guoid1"><%=guoid1%></td>
                </tr>
                <tr >
			<td width=10></td>
			<td class="b8" align="right">姓名</td>
			<td width=5></td>
                        <td id="guoname1"><%=guoname1%></td>
                    
		</tr>

                 <tr height="22">
			
			<td width=10></td>	
			<td class="b8" align="right">借據編號</td>
			<td width=10></td>  
 			<% if guoid1 <> "" then  %>                     
                        <td id="guln1"><%=guoln1%></td>                        
                        <%end if %> 
                        </tr>
<%end if%>
<%if guoid2<>"" then %>
			<tr height="22" >
			<td width=10></td>
			<td class="b8" align="right">2.擔保其他人</td>
			<td width=10></td>
			<td id="guoid2"><%=guoid2%></td>
                </tr>
                <tr >
			<td width=10></td>
			<td class="b8" align="right">姓名</td>
			<td width=5></td>
                        <td id="guoname2"><%=guoname2%></td>
                    
		</tr>

                       <tr>
 			<td width=10></td>
			<td class="b8" align="right">借據編號</td>
			<td width=10></td>   
                        <%if guoid2 <> "" then %>                     
                        <td id="guln2"><%=guoln2%></td>  
                        <% end if %> 
                        </tr>
<%end if %>
<%if guoid3 <> "" then %>
			<tr height="22">
			<td width=10></td>
			<td class="b8" align="right">3.擔保其他人</td>
			<td width=10></td>
			<td id="guoid3"><%=guoid3%></td>
			                </tr>
                <tr >
			<td width=10></td>
			<td class="b8" align="right">姓名</td>
			<td width=5></td>
                        <td id="guoname3"><%=guoname3%></td>
                    
		</tr>

                       <tr>
 			<td width=10></td>
			<td class="b8" align="right">貸款編號</td>
			<td width=10></td>   
                        <% if guoid3 <> "" then %>                   
                        <td id="guln3"><%=guoln3%></td>   
                        <% end if %>
                        </tr> 
<%end if %>
                </table>
                </td>
          </tr>

</table>	
</tr>
<table border="0" cellspacing="0" cellpadding="0">
<tr>
		<td width="300" valign="top">
                 <table border="0" cellspacing="0" cellpadding="0">	
		<%if memno <> "" then %>
			<tr height="22">	
			<td class="b8" align="right">股金結餘</td>
			<td width=30></td>
			<td id="ttlbal"><b><%=formatNumber(ttlbal,2)%></b></td>
                        </tr>
                        <tr height="22">			
			<%if mstatus="A" or mstatus="M" then %>
			<td class="b8" align="right">自動轉帳</td>
			<td width=10></td>
			<td id="monthsave"><%=monthsave%></td>
			<td width=10></td>
                        </tr> 
		<%end if %>
         
		<%if mstatus="T" then %>
                        <tr>
         	        <td class="b8" align="right">庫房扣薪(股金)</td>
			<td width=30></td>
			<td  id="monthssave" ><%=formatnumber(monthssave,2)%></td>
			<%end if %>
		        </tr>
                <%end if %>
		<%if tpayamt <>0 then %>
                       <tr>
         	        <td class="b8" align="right">庫房扣薪</td>
			<td width=30></td>
                         
			<td id="tpayamt"><%=formatnumber(tpayamt,2)%></td>
			
		        </tr>
		<%end if %>
               </table>
               </td>

		<%if lnnum <> "" then %>	
		<td width="300" valign="top">
                        <table border="0" cellspacing="0" cellpadding="0"> 
			<tr height="22">
			<td width=10></td>
			<td class="b8" align="right">借據編號</td>
			<td width=10></td>
			<td id="lnnum"><%=lnnum%></td>
                        <td width=10></td>
                        </tr>

			<tr height="22">
			<td width=10></td>
			<td class="b8" align="right">期數</td>
			<td width=10></td>
			<td id="install"><%=install%></td>
                        </tr>
                        <tr height="22">
                        <td width=10></td> 
                        <td class="b8" align="right">每月還款</td>
			<td width=10></td> 
			<td id="monthrepay"><%=formatNumber(monthrepay,2)%></td>
                        </tr>

			
                        </table>
                        </td>
              
		<td width="300" valign="top">
                 <table border="0" cellspacing="0" cellpadding="0">                                 
			<tr height="22">
			<td width=10></td>                 
                        <td class="b8" align="right">取票日期</td>
			<td width=10></td> 
			<td id="lndate"><%=lndate%></td>
                        </tr>
		<tr height="22">
			<td width=10></td> 
                        <td class="b8" align="right">借據金額</td>
			<td width=10></td> 
			<td id="appamt"><%=formatNumber(appamt,2)%></td>
			</tr>    

 			<tr height="22">	
			<td width=10></td> 
                        <td class="b8" align="right">借據結餘</td>
			<td width=10></td> 
			<td id="bal"><font color=red><%=formatNumber(bal,2)%></font></td>
			</tr>       
                   </table>
                   </td>    
	
<% end if %>
</tr>
</table>
<table

<%if repaystat="N" then %>
				<tr>
					<td width=10></td>
					<td class="b8" align="right">清數日期</td>
					<td width=10></td>
		<td><input type="text" name="cleardate" value="<%=cleardate%>" size="10" onblur="if(!formatDate(this)){this.value=''};calculation();"></td>
				</tr>
	<tr>
               <td width=10></td>
		<td class="b8" align="right">清數本金</td>
		<td width=10></td>
		<td id="cashamt"><%=formatNumber(cashamt,2)%></td>	


               <td width=10></td>
		<td class="b8" align="right">清數利息</td>
		<td width=10></td>
		<td id="cashint"><%=formatNumber(cashint,2)%></td>	

               <td width=10></td>
		<td class="b8" align="right">清數金額合共</td>
		<td width=10></td>
		<td id="ttlpamt"><font color=red><b><%=formatNumber(ttlpamt,2)%></b></font></td>
	
	</tr>  
	<tr>
               <td width=10></td>

		<td class="b8" align="right"><%=xmess%></td>



		<td width=10></td>
		<td id="monthsamt"><%=formatNumber(monthsamt,2)%></td>	

              <td width=10></td>
		<td class="b8" align="right">還款利息</td>
		<td width=10></td>
		<td id="intamt"><%=formatNumber(intamt,2)%></td>	
               <td width=10></td>
		<td class="b8" align="right">還款本金</td>
		<td width=10></td>
		<td id="pincamt"><%=formatNumber(pincamt,2)%></td>	

               <td width=10></td>
		<td class="b8" align="right">還款金額合共</td>
		<td width=10></td>
		<td id="ttlamt  "><font color=red><b><%=formatNumber(ttlamt,2)%></b></font></td>
	
	</tr>  
<%end if%>
				<tr>
					<td width=10></td>
					<td class="b8" align="right">瀏覽日期</td>
					<td width=10></td>
					<td><input type="text" name="chkdate" value="<%=chkdate%>" size="10" onblur="if(!formatDate(this)){this.value=''};">
				</tr>

</table>
<td colspan="3" align="right">


<% if id <> "" then %>
		<input type="button" value="查詢個人帳" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.chkdate.value )" class="sbttn">					
<% end if %>
		</td>
	</tr>
</table>




</form>
</body>
</html>

 