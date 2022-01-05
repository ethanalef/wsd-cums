<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<!-- #include file="cutpro.asp" -->
<%
if request.form("back") <> "" then
	response.redirect "main.asp"'
end if


if request.form("calc") <> "" then

   monthrepay = request.form("appamt")/request.form("install")
end if


if request.form("Search")<>""  or id <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
 
        msg = ""  
         if id ="" then
            set rs = conn.execute("select memname,memcname from memmaster where memno = '"&memno&"' " )
            if not rs.eof then
               memname = rs(0)
               memcname = rs(1)
            end if
             rs.close
         end if  
         
	if xlnnum = "" then
           set rs=conn.execute("select a.memno,a.memcname,a.memname,b.lnnum from memmaster a,loanrec b where a.memno='"&memno&"' and a.memno=b.memno and b.repaystat='N' ")
           if rs.eof then
              msg = "社員沒有貸款記錄 !"
           else
              xlnnum = rs("lnnum")	
              memname = rs("memname")
	      memcname = rs("memcname")
              memno=rs("memno")
              id= memno
           end if
           rs.close
       end if               
 
       if msg ="" then
		sql = "select * from loanrec where lnnum = '"&xlnnum&"'  "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn

		For Each Field in rs.fields
			if Field.name="cleardate" or Field.name="lndate"  then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
		
           
                rs.close 
         work = "0"
         if appamt<>bal then
            work = "1"
         end if

         end if  
     

		sql = "select * from guarantor where lnnum= '" & xlnnum &"' "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
            x=1
            do while not rs.eof
               select case x
                 case 1   
                  guarantorID=rs("guarantorID")
                  guarantorName=rs("guarantorName") 
                  idx1= 0              
                case 2
                 guarantor2ID=rs("guarantorID")
                 guarantor2Name=rs("guarantorName")     
                 idx2 = 0 
                case 3
                 guarantor3ID=rs("guarantorID")
                 guarantor3Name=rs("guarantorName")     
                 idx3= 0
               end select
               x=x+1
               rs.movenext
               loop

end if



if request.form("action") <> "" then
        addloan = 0
	For Each Field in Request.Form
 	 
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)

	
	Next

        set rs = server.createobject("ADODB.Recordset")
	msg = ""
        nlndate = DateSerial(right(lndate,4),mid(lndate,4,2),left(lndate,2))
        if cleardate <>"" then
        ncleardate= DateSerial(right(cleardate,4),mid(cleardate,4,2),left(cleardate,2))
        end if
	if Lnnum<>"" then

		sql = "select * from loanrec where lnnum=" & lnnum
		rs.open sql, conn,1
		if not rs.eof then
                   xbal = rs("bal")
                   xclrdate = rs("cleardate")
                   xrepaystat = rs("repaystat")
                    xrepay = rs("monthrepay")   
                    xlndate = rs("lndate")               
                    xxlndate= right("0"&+day(xlndate),2)&"/"&right("0"&month(xlndate),2)&"/"&year(xlndate)  
                    lnflag=rs("lnflag")
                    oldlnnum = rs("oldlnnum")
               end if
               rs.close  
               if idx1 = 1 then
                  conn.execute("delete guarantor where guarantorID='"&guarantorID&"' and lnnum='"&lnnum&"' ")
              end if
              if idx2 = 1 then
                 conn.execute("delete guarantor where guarantorID='"&guarantor2ID&"' and lnnum='"&lnnum&"' ")
              end if
              if idx3 = 1 then
                 conn.execute("delete guarantor where guarantorID='"&guarantor3ID&"' and lnnum='"&lnnum&"' ")
              end if		
                 IF CLEARDATE <> "" THEN
		         	                   
	
                       mess = "社員編號 "&memno&" 的貸款帳號 "&lnnum&" 修改清數日期由 "&xclrdate&" 至 "&xcleardate&" 清數現狀由 ”N” 至 ”C” "
                       addUserLog  mess
                       conn.begintrans
                       CONN.EXECUTE("UPDATE LOANREC SET CLEARDATE='"&nCLEARDATE&"',REPAYSTAT='"&REPAYSTAT&"' WHERE LNNUM = '"&LNNUM&"' "  )  
                       conn.committrans 
                    
                else 
                     
		   
                       mess = "社員編號 "&memno&" 的貸款帳號 "&lnnum&" 修改清數日期由 "&xclrdate&" 至  空白  清數現狀由 ”C” 至 ”N” "
                       addUserLog  mess
                       conn.begintrans
                       CONN.EXECUTE("UPDATE LOANREC SET CLEARDATE=NULL ,REPAYSTAT='N' WHERE LNNUM = '"&LNNUM&"' ")
                       conn.committrans 
                  
                end if    
                if xrepay <> monthrepay then
                      conn.begintrans
  		      CONN.EXECUTE("UPDATE LOANREC SET monthrepay="&monthrepay&"  WHERE LNNUM = '"&LNNUM&"' ")                                              
                      conn.committrans     
                end if 
                if bal <> xbal then
                      mess = "社員編號 "&memno&" 的貸款帳號 "&lnnum&" 修改本金結餘期由 "&xbal&" 至  &bal&"
                      addUserLog  mess
                      conn.begintrans
  		      CONN.EXECUTE("UPDATE LOANREC SET bal="&bal&"  WHERE LNNUM = '"&LNNUM&"' ")                                              
                      conn.committrans
                end if   
                if lndate <> xxlndate then
                     
                      mess = "社員編號  "&memno&" 的貸款帳號 "&lnnum&" 修改貸款日期由 "&xlndate&" 至  &lndate&"
                      addUserLog  mess
                      conn.begintrans
                      if lnflag="Y" then
                         conn.execute("update loanrec set cleardate='"&nlndate&"' where lnnum='"&oldlnnum&"' ")
                         conn.execute("update loan set ldate='"&nlndate&"' where lnnum='"&ooldlnnum&"' and code='D8' ")
                      end if
                      conn.execute("update loan  set ldate = '"&nlndate&"' where LNNUM = '"&LNNUM&"' and ldate='"&xlndate&"' ")  
  		      CONN.EXECUTE("UPDATE LOANREC SET lndate ='"&nlndate&"'  WHERE LNNUM = '"&LNNUM&"' ")                                              
                      conn.committrans
                end if   
		
                id = ""
	        For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	        Next
        
		

		msg = "紀錄已更新"

	end if
end if
if request.form("bye")<> "" then
        id = ""
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next
                       		
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
     Mn = parseInt(strM)
      Yr = parseInt(strY)
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
<%if msg<>"" then %>
<div><center><font color="red"><%=msg%></font></center></div>
<br>
<% end if%>


<form name="form1" method="post" action="ncloanDetail.asp">

<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="xlnnum" value="<%=xlnnum%>">
<input type="hidden" name="guarantorName" value="<%=guarantorName%>">
<input type="hidden" name="guarantor2Name" value="<%=guarantor2Name%>">
<input type="hidden" name="guarantor3Name" value="<%=guarantor3Name%>">
<input type="hidden" name="mstatus" value="<%=mstatus%>">
<input type="hidden" name="cutdate" value="<%=cutdate%>">
<input type="hidden" name="lastmonth" value="<%=lastmonth%>">
<input type="hidden" name="lastyear" value="<%=lastyear%>">
<input type="hidden" name="spass" value="<%=spass%>">
<input type="hidden" name="work" value="<%=work%>">
<div><center><font size="3">貸款修正</font></center></div>

<center>

		<td width="700" valign="top">
			<table border="0" cellspacing="0" cellpadding="0">
			<tr>
             		 <td width=30></td>
			<td class="b12" align="left">社員號碼</td>
			<td width=50></td>
			<td><input type="text" name="memNo" value="<%=memNo%>" size="10" maxlength="10"<%if id<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>>
			<%if id = "" then %>
			<input type="button" value="選擇"  onclick="popup('pop_srhLoan.asp?key='+document.form1.memNo.value)" class="sbttn"  >          
			<input type="submit" value="搜尋" name="Search" class ="Sbttn">
			<% end if %>
                        </TD>
			</tr>

			<tr>
 	               <td width=30></td>
			<td class="b12" align="left">社員名稱</td>
			<td width=50></td>
			<td><input type="text" name="memName" value="<%=memName%>" size="30" maxlength="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td>		
			</tr>
			<tr>
 	               <td width=30></td>
			<td class="b12" align="left"></td>
			<td width=50></td>
			<td><input type="text" name="memcName" value="<%=memcName%>" size="30" maxlength="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td>		
			</tr>
			<tr>
	               <td width=30></td>	
			<td class="b12" align="left">貸款號碼</td>
			<td width=50></td>
			<td><input type="text" name="lnnum" value="<%=lnnum%>" size="10" maxlength="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td>		
			</tr>

			<tr>
  		        <td width=30></td>
			<td class="b12" align="left">設定日期</td>
			<td width=50></td>
			<td><input type="text" name="lndate" value="<%=lndate%>" size="10"<%if work="1" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 

	               <td width=30></td>
			<td class="b12" align="left">1.擔保人 </td>
			<td width=20></td>
			<td><input type="text" name="guarantorID" value="<%=guarantorID%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
			<td width=5></td> 
	                <td id="guarantName"><%=guarantorName%>
                        <% if guarantorID<>"" then %>
                        <input type="radio"  name="idx1" value="1">取消
                        <%end if%> 
                        </td>
			</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">貸款金額</td>
		<td width=50></td>
		<td><input type="text" name="appamt" value="<%=appamt%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	               <td width=30></td>
			<td class="b12" align="left">2.擔保人 </td>
			<td width=20></td>
			<td><input type="text" name="guarantor2ID" value="<%=guarantor2ID%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
			<td width=5></td> 
			<td id="guarant2Name"><%=guarantor2Name%>
                        <% if guarantor2ID<>"" then %>
                        <input type="radio"  name="idx2" value="1">取消
                        <%end if%>
                        </td>
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">攤分期數</td>
		<td width=50></td>
		<td><input type="text" name="install" value="<%=install%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
	               <td width=30></td>
		<td class="b12" align="left">3.擔保人 </td>
		<td width=20></td>
		<td><input type="text" name="guarantor3ID" value="<%=guarantor3ID%>" size="10"<%if xlnnum<>"" then response.write " onfocus=""form1.cleardate.focus();""" end if%>></td> 
                <td width=5></td> 
		<td id="guarant3Name"><%=guarantor3Name%>
                <% if guarantor3ID<>"" then %>
                <input type="radio"  name="idx3" value="1">取消
                <%end if%>
                </td>
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">每月還款</td>
		<td width=50></td>
		<td><input type="text" name="monthrepay" value="<%=monthrepay%>" size="10"></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">本金結餘</td>
		<td width=50></td>
		<td><input type="text" name="bal" value="<%=bal%>" size="10" ></td> 
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">清數日期</td>
		<td width=50></td>
		<td><input type="text" name="cleardate" value="<%=cleardate%>" size="10" onblur="if(!formatDate(this)){this.value=''};"></td>
	</tr>

	<tr>
               <td width=30></td>
		<td class="b12" align="left">清數現狀</td>
		<td width=50></td>
		<td><input type="text" name="repaystat" value="<%=repaystat%>" maxlength="1" size="1" ></td>
	</tr>
  
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
