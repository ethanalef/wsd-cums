<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "ncloan.asp"'
end if
if request.form("calc") <> "" then
    
   monthrepay = request.form("appamt")/request.form("install")
end if

if request.form("action") <> "" then
        addloan = 0
	For Each Field in Request.Form
 	 
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)

	
	Next
    set rs = server.createobject("ADODB.Recordset")
	msg = ""

	if Lnnum<>"" then
		sql = "select * from loanrec where lnnum=" & lnnum
		rs.open sql, conn,1
		if rs.eof then
			msg = "貸款號碼不存在 "
		end if
		rs.close
	end if



	if msg="" then
		conn.begintrans

             
	        sql = "select * from loanrec where LNNUM="&XLNNUM
		rs.open sql, conn, 3, 3
            
  	        if cleardate<>"" then rs("cleardate") = right(cleardate,4)&"/"&mid(cleardate,4,2)&"/"&left(cleardate,2) end if
                 rs("repaystat")=repaystat
		if autopamt <> 0 then rs("autopamt")=autopamt  else rs("autopamt") = 0 end if
		if salarydeduct<>"" then  rs("salarydeduct")=salarydeduct else  rs("salarydeduct")=0 end if
                if  intchoice = "銀行"  then
                    rs("intchoice") = "A"
                else
                    rs("intchoice") = "S"                   
                end if
		rs.update
		
              


		rs.close 	
         
        
		conn.committrans
		msg = "紀錄已更新"

	end if
else
	if xlnnum <> "" then
		sql = "select * from loanrec where lnnum= " & xlnnum
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if rs.eof then
			response.redirect "ncloan.asp"
		else
		For Each Field in rs.fields
			if Field.name="cleardate" or Field.name="lndate"  then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
		end if

                rs.close 
         end if  
         if intchoice = "A" then
            intchoice = "銀行"
         else
             intchoice = "庫房"
          end if    
                        
		
	
end if
%>
<html>
<head>
<title>貸款資料修正</title>
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
	mInterest=0;loanAmt=0;installment=0
	if (formObj.repayamt.value!=""&&formObj.loanAmt.value!=0){
           loanAmt=parseInt(formObj.repayamt.value)
			chequeDate=formObj.repaydate.value
			Months= new Array(31,28,31,30,31,30,31,31,30,31,30,31);
			Y=parseInt(chequeDate.substr(7,4))
			M=pseInt(chequeDate.substr(4,1))
			D=parseInt(chequeDate.substr(1,2))
			mD=Months[M-1]
			if((Y % 4 == 0) && ((Y % 100 != 0) || (Y %400 == 0)) && M==2)
				mD=29
			mInterest=Math.floor(loanAmt*.01/mD*D)
		}else{
			mInterest=0
		}

		document.all.tags( "td" )['intamt'].innerHTML=mInterest;

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

	if (formObj.lnnum.value==""){
		reqField=reqField+", 貸款號碼";
		if (!placeFocus)
			placeFocus=formObj.lnnum;
	}

	if (formObj.salarydeduct.value==""){
		reqField=reqField+", 庫房扣薪";
		if (!placeFocus)
			placeFocus=formObj.salarydeduct;
	}

	if (formObj.autopamt.value==""){
		reqField=reqField+", 自動轉賬";
		if (!placeFocus)
			placeFocus=formObj.autopamt;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.memNo.focus()">
<DIV>

<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>

<br>
<form name="form1" method="post" action="loanrepay.asp">

<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="xlnnum" value="<%=xlnnum%>">
<div><center><font size="3">貸款還款</font></center></div>


<div style="z-index: 97; left: 350px; width: 600px; position: absolute; top: 116px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">社員號碼</td>
		<td width=50></td>
		<td><input type="text" name="memNo" value="<%=memNo%>" size="10" <%if id<>"" then response.write " onfocus=""form1.memName.focus();""" end if%>></td>
<%if id = "" then %>
		<td><input type="button" value="選擇"  onclick="popup('pop_srhloan.asp?key='+document.form1.memNo.value)" class="sbttn"  ></td>          
		<td><input type="submit" value="搜尋" name="Search" class ="Sbttn"></td>
<% end if %>

	</tr>
	</tr>
</TABLE>
</div>
<div style="z-index: 98; left: 350px; width: 236px; position: absolute; top: 140px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
 
               <td width=30></td>
		<td class="b12" align="left">貸款號碼</td>
		<td width=50></td>
		<td><input type="text" name="lnnum" value="<%=lnnum%>" size="10" maxlength="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td>		
	</tr>
</TABLE>
</div>
<div style="z-index: 99; left: 350px; width: 236px; position: absolute; top: 163px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">設定日期</td>
		<td width=50></td>
		<td><input type="text" name="lndate" value="<%=lndate%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>
</TABLE>
</div>
<div style="z-index: 101; left: 350px; width: 236px; position: absolute; top: 186px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">貸款金額</td>
		<td width=50></td>
		<td><input type="text" name="appamt" value="<%=appamt%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>
</TABLE>
</div>

<div style="z-index: 102; left: 350px; width: 236px; position: absolute; top: 209px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">攤分期數</td>
		<td width=50></td>
		<td><input type="text" name="install" value="<%=install%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>
</TABLE>
</div>

<div style="z-index: 103; left: 350px; width: 236px; position: absolute; top: 232px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">每月還款</td>
		<td width=50></td>
		<td><input type="text" name="monthrepay" value="<%=monthrepay%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	</tr>
</TABLE>
</div>

</div>

<div style="z-index: 104; left: 350px; width: 236px; position: absolute; top: 265px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">還款日期</td>
		<td width=50></td>
		<td><input type="text" name="repaydate" value="<%=repaydate%>" size="10"></td> 
	</tr>
</TABLE>
</div>
<div style="z-index: 105; left: 350px; width: 236px; position: absolute; top: 288px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">還款金額</td>
		<td width=50></td>
		<td><input type="text" name="repayamt" value="<%=repayamt%>" size="10" onblur="if(!formatNum(this)){this.value=''};calculation();"></td>
	</tr>
</TABLE>
</div>

<div style="z-index: 106; left: 350px; width: 236px; position: absolute; top: 311px;  height: 68px" >
 <table id="table1" border="0" cellspacing="0" cellpadding="0" >
	<tr>
               <td width=30></td>
		<td class="b12" align="left">利息</td>
		<td width=75></td>
		<td>
		<td><input type="text" name="intamt" value="<%=intamt%>" size="10"></td> 
		</td>
	</tr>
</TABLE>
</div>


<div style="z-index: 118; left: 350px; width: 1000px; position: absolute; top: 380px;  height: 68px">
<table id-"table20" border="0" cellspacing="0" cellpadding="0">       
	<tr>
		<td width=30></td>
		<td colspan="3" align="right">
			<%if session("userLevel")<>5 and session("userLevel")<>1 and session("userLevel")<>6 then%>
			<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
			<%end if%>
			<input type="submit" value="返回" name="back" class="sbttn">
		</td>
	</tr>
</table>
</DIV>
</div>
</form>
</body>
</html>
