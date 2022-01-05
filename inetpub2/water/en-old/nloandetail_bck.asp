<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "main.asp"'
end if
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next



if request.form("Search")<>""   then
	For Each Field in Request.Form
 	 
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)

	
	Next
	if uid <> "" then
		sql = "select * from loanapp where  DELETED= 0  AND uid= " & uid
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn

		For Each Field in rs.fields
			if  Field.name="lndate"  then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
                rs.close
                bal=chequeamt
                appamt=loanAmt
                lndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())  
		install= installment
                if int(appamt/install)=(appamt/install) then
                   monthrepay = int(appamt/install)
                else
                   monthrepay = int(appamt/install)+1
                end if
                 bal = 0
                intamt =0 
                lnflag = "N"
                if oldlnnum<>"" then
                   lnflag ="Y"
                   set rs = conn.execute("select bal from loanrec where lnnum ='"&oldlnnum&"' and repaystat='N' ")
                   if not rs.eof then
                      bal =bal +  rs(0)
                   end if
                   rs.close
                  
                end if
                                       
                
	  else 	
         if id ="" then
         
            set rs = conn.execute("select memname,memcname,memno from memmaster where memno = '"&memno&"' " )
            if not rs.eof then
               memname = rs(0)
               memcname = rs(1)
               id = memno
            end if
             rs.close
         end if  
         if id <> "" then
          set rs = conn.execute("select * from loanapp where DELETED = 0 AND memno='"&memno&"' and (firstapproval='Approved' or secondapproval='Approved') ")    
          if not rs.eof then
             		For Each Field in rs.fields
			if  Field.name="lndate"  then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next		
                     rs.close       
                        
                bal = loanamt 
                appamt=loanAmt
                lndate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())  
		install= installment
                if int(appamt/install)=(appamt/install) then
                   monthrepay = int(appamt/install)
                else
                   monthrepay = int(appamt/install)+1
                end if
                 
                intamt =0 
                lnflag = "N"
                lnflag = "N"
                if oldlnnum<>"" then
                   lnflag ="Y"
                   set rs = conn.execute("select bal from loanrec where lnnum ='"&oldlnnum&"' and repaystat='N' ")
                   if not rs.eof then
                      bal = chequeamt + rs(0)
                   end if
                   rs.close
                     
                end if
            
          else
              intamt = 0
              monthrepay = 0 
              msg ="無申請貸款"
              id = "" 
          end if      
          else
            
             intamt = 0
              monthrepay = 0 
              msg ="社員不存在"
         end if               
         end if  
  
          xlnnum = "" 

else

if request.form("action") <> "" then
         addloan = 0
	For Each Field in Request.Form
 	 
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)

	
	Next



    set rs1 = server.createobject("ADODB.Recordset")
	msg = ""

	if Lnnum="" then
                xyear=year(date())
		xmon = right("0"&month(date()),2)
                chkdate=xyear&xmon 
		sql1 = "select max(lnnum) from loanrec "
		rs1.open sql1, conn,1
		if not rs1.eof then

                   if mid(rs1(0),1,6) = chkdate then  
                   num = mid(rs1(0),7,4)
                   num = num + 1
                   lnnum = mid(rs1(0),1,6)&right("0000"&num,4)
                else
                   lnnum = xyear&xmon&"0001"      
		end if
                end if
		rs1.close
 
		
	end if

		conn.begintrans

		conn.execute("update loanapp set deleted = 1 where uid='"&uid&"'   " )

	        sql1 = "select * from loanrec where 0=1"
		rs1.open sql1, conn, 2, 2
		rs1.addnew
                rs1("lnnum")=lnnum
                rs1("memno")=memno
             	if lndate<>"" then rs1("lndate") = right(lndate,4)&"/"&mid(lndate,4,2)&"/"&left(lndate,2) else rs1("lndate")="" end if
                rs1("appamt")=cdbl(appamt)
                rs1("monthrepay")=cdbl(monthrepay)
		rs1("install")=cdbl(install)
                
                rs1("lnflag")=lnflag
                rs1("bal")=bal
		rs1("repaystat") = "N"

        
               if oldlnnum <>"" then
                lnflag="Y"
                rs1("chequeamt") = chequeamt
                rs1("oldlnnum") = oldlnnum
    
                end if
                
		rs1.update
                rs1.close
                xlndate = right(lndate,4)&"/"&mid(lndate,4,2)&"/"&left(lndate,2)
                if  oldlnnum <>"" then
                    conn.execute("insert into loan ( memno,lnnum,code,ldate,amount) values ( '"&memno&"','"&lnnum&"','D1','"&xlndate&"',"&bal&") ")  
               else
                  conn.execute("insert into loan ( memno,lnnum,code,ldate,amount) values ( '"&memno&"','"&lnnum&"','D0','"&xlndate&"',"&appamt&") ")                                                  
               end if 

               
               addUserLog  "社員編號 "&memno&" 的新貸款帳號 "&lnnum

                if guarantorID<>"" then
                   conn.execute("insert into guarantor  ( memno,lnnum,date,guarantorID,guarantorName ) values ( '"&memno&"','"&lnnum&"','"&xlndate&"',"&guarantorID&",'"&guarantorName&"') ") 
               end if
               if guarantor2ID<>"" then
                   conn.execute("insert into guarantor  ( memno,lnnum,date,guarantorID,guarantorName ) values ( '"&memno&"','"&lnnum&"','"&xlndate&"',"&guarantor2ID&",'"&guarantor2Name&"') ") 
               end if
               if guarantor3ID<>"" then
                   conn.execute("insert into guarantor  ( memno,lnnum,date,guarantorID,guarantorName ) values ( '"&memno&"','"&lnnum&"','"&xlndate&"',"&guarantor3ID&",'"&guarantor3Name&"') ") 
               end if	
              
            
               xlnnum = lnnum               
               
		conn.committrans
		msg = "紀錄已更新"


end if
end if
if request.form("bye")<> "" then
        id =""                
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next
        uid = ""  
        xlnnum = ""
        autopamt = 0
        intamt = 0
        salarydeduct = 0                        		
        lnnum = ""
end if
%>
<html>
<head>
<title>貸款建立</title>
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


function calculation(){
	formObj=document.form1;

 
	if (formObj.lndate.value!=""){
           ssdate = formObj.lndate.value
	   Months= new Array(31,28,31,30,31,30,31,31,30,31,30,31)	  
	   Y=ssdate.substr(6,4)           
           M=ssdate.substr(3,2)
	   D=ssdate.substr(0,2)
	   YY=parseInt(ssdate.substr(6,4)) 
           MM=parseInt(ssdate.substr(3,2))
	   DD=parseInt(ssdate.substr(0,2))
           mD=Months[M-1]
           if((Y % 4 == 0) && ((Y % 100 != 0) || (Y %400 == 0))){                 
               if(M == 2){
               mD = 29
           }
           }
               
           olnbal =parseInt(formObj.bal.value)
           lnamt =parseInt(formObj.appamt.value)
	   chqamt = parseInt(formObj.chequeamt.value)l
					
           if (formObj.oldlnnum.value!=""){
              difamt=lnamt-chqamt
              int1 =(difamt*0.01)
              int2 =(chqamt*0.01*(mD-D+1)/mD)
	      int3 = int1+int2
              document.form1.intamt.value = int3
	      document.form1.chequeamt.value = chqamt
	      document.all.tags( "td" )['intamt'].innerHTML=Math.round(int3,3)
 
           }else{

             int2 =Math.round((chqamt*0.01*(mD-D+1)/mD),2) 
	     document.form1.intamt.value = int2
            
	     document.all.tags( "td" )['intamt'].innerHTML=int2	
           }
        }
} 


function clearln(){        
     if (document.form1.lnflag.value =='Y'){
        document.form1.lnflag.value = 'N'

}else{
       
         document.form1.lnflag.value = 'Y'
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





	if (!formatDate(formObj.lndate)){
		reqField=reqField+", 設定日期";
		if (!placeFocus)
			placeFocus=formObj.lndate;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="calculation();form1.memNo.focus()">
<DIV>
<!-- #include file="menu.asp" -->
<%if msg<>"" then %>
<div><center><font color="red"><%=msg%></font></center></div>
<br>
<% end if%>


<form name="form1" method="post" action="nloanDetail.asp">

<input type="hidden" name="uid" value="<%=uid%>">
<input type="hidden" name="xlnnum" value="<%=xlnnum%>">
<input type="hidden" name="loanAmt" value="<%=loanAmt%>">
<input type="hidden" name="installment" value="<%=installment%>">
<input type="hidden" name="guarantorName" value="<%=guarantorName%>">
<input type="hidden" name="guarantor2Name" value="<%=guarantor3Name%>">
<input type="hidden" name="guarantor3name" value="<%=guarantor3Name%>">
<input type="hidden" name="intamt" value="<%=intamt%>">
<input type="hidden" name="chequeamt" value="<%=chequeamt%>">
<input type="hidden" name="bal" value="<%=bal%>">
<input type="hidden" name="intamt" value="<%=intamt%>"> 
<input type="hidden" name="monthrepay" value="<%=monthrepay%>">
<input type="hidden" name="oldlnnum" value="<%=oldlnnum%>">

<div><center><font size="3">貸款建立</font></center></div>


<center>

		<td width="700" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
			<tr>
             		 <td width=30></td>
			<td class="b12" align="left">社員號碼</td>
			<td width=50></td>
			<td><input type="text" name="memNo" value="<%=memNo%>" size="10" maxlength="10"<%if id<>"" then response.write " onfocus=""form1.lndate.focus();""" end if%>>
			<%if id = "" then %>
			<input type="button" value="選擇"  onclick="popup('pop_srhnewln.asp?key='+document.form1.memNo.value)" class="sbttn"  >          
			<input type="submit" value="搜尋" name="Search" class ="Sbttn">
			<% end if %>
                        </TD>
			</tr>

			<tr>
 	               <td width=30></td>
			<td class="b12" align="left">社員名稱</td>
			<td width=50></td>
			<td><input type="text" name="memName" value="<%=memName%>" size="30" maxlength="10"<%if id<>"" then response.write " onfocus=""form1.lndate.focus();""" end if%>></td>		
                        
			</tr>
			<td width=30></td>
			<td class="b12" align="left"></td>
			<td width=50></td>
			<td><input type="text" name="memcName" value="<%=memcName%>" size="30" maxlength="10"<%if id<>"" then response.write " onfocus=""form1.lndate.focus();""" end if%>></td>		
                        
			</tr>
			<tr>
	               <td width=30></td>	
			<td class="b12" align="left">貸款號碼</td>
			<td width=50></td>
			<td><input type="text" name="lnnum" value="<%=lnnum%>" size="10" maxlength="10"<%if id<>"" then response.write " onfocus=""form1.lndate.focus();""" end if%>></td>		
			</tr>

			<tr>
  		        <td width=30></td>
			<td class="b12" align="left">取票日期</td>
			<td width=50></td>
			<td><input type="text" name="lndate" value="<%=lndate%>" size="10"  onblur="if(!formatDate(this)){this.value=''};calculation();"></td> 

	               <td width=30></td>
			<td class="b12" align="left">1.擔保人 </td>
			<td width=20></td>
			<td><input type="text" name="guarantorID" value="<%=guarantorID%>" size="10"<%if id<>"" then response.write " onfocus=""form1.lndate.focus();""" end if%>></td> 
			<td width=5></td> 
	                <td id="guarantName"><%=guarantorName%></td>
			</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">貸款金額</td>
		<td width=50></td>
		<td><input type="text" name="appamt" value="<%=appamt%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.lndate.focus();""" end if%>></td> 
	               <td width=30></td>
			<td class="b12" align="left">2.擔保人 </td>
			<td width=20></td>
			<td><input type="text" name="guarantor2ID" value="<%=guarantor2ID%>" size="10"<%if id<>"" then response.write " onfocus=""form1.lndate.focus();""" end if%>></td> 
			<td width=5></td> 
			<td id="guarant2Name"><%=guarantor2Name%></td>
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">攤還期數</td>
		<td width=50></td>
		<td><input type="text" name="install" value="<%=install%>" size="10"<%if id<>"" then response.write " onfocus=""form1.lndate.focus();""" end if%>></td> 
	               <td width=30></td>
		<td class="b12" align="left">3.擔保人 </td>
		<td width=20></td>
		<td><input type="text" name="guarantor3ID" value="<%=guarantor3ID%>" size="10"<%if id<>"" then response.write " onfocus=""form1.lndate.focus();""" end if%>></td> 
                <td width=5></td> 
		<td id="guarant3Name"><%=guarantor3Name%></td>
	</tr>



	<% if id <>"" then  %>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">每月還款</td>
		<td width=50></td>
		<td id= "monthrepay"><%=formatNumber(monthrepay,2)%></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">貸款結餘</td>
		<td width=50></td>
		<td id="bal"><%=formatnumber(bal,2)%></td> 
	</tr>

	<tr>
               <td width=30></td>
               <td class="b12" align="left">續約貸款</td>
               <td width=50></td> 
               <td><input type="checkbox" name="lnflag"   value="<% =lnflag %>"<%if lnflag="Y" then response.write " checked" end if%> onclick="if (this.checked){clearln()}"></td>
	</tr>


<% end if %>


  
	<tr>
		<td width=30></td>
		<td colspan="3" align="right">
		<%if id<>"" then %>
		<%if session("userLevel")<>5 and session("userLevel")<>1 and session("userLevel")<>6 then%>
		<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
		<%end if%>
		<%end if %>
		<input type="submit" value="取消" name="bye" class="sbttn">

			<input type="submit" value="返回" name="back" class="sbttn">
		</td>
	</tr>
</table>

</CENTER>

</form>
</body>
</html>
