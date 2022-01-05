<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
For Each Field in Request.Form
	TheString = Field & "= Request.Form(""" & Field & """)"
	Execute(TheString)
Next

bdNum=cdbl(bdNum)
sql = "select * from glControl"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn
acPeriod=rs("acPeriod")
acYear=rs("acYear")
rs.close
if acPeriod<=4 then
	m=acPeriod+8
	y=acYear
else
	m=acPeriod-4
	y=acYear+1
end if

if request("from") = Request.ServerVariables("script_name") then
	if request.form("Back")<>"" then
		response.redirect "glTx.asp"
	end if

	msg = ""
	if msg = "" then
		if request.form("Add")<>"" then
			mDate = y&"/"&m&"/"&mDay
			conn.begintrans
			sql = "select top 1 * from glTx order by glTxNo desc"
			rs.open sql, conn, 2, 2
			if rs.eof then
				id = 1
			else
				id = rs("glTxNo")+1
			end if
			rs.addnew
			For Each Field in rs.fields
				if Field.name="glTxNo" then
					rs("glTxNo")=id
				elseif Field.name="txDate" then
					rs("txDate") = mDate
				elseif Field.name="deleted" then
					rs("deleted") = 0
				else
					TheString = "rs(""" & Field.name & """) = " & Field.name
					Execute(TheString)
				end if
			Next
			rs.update
			rs.close
			if txType="D" then
				conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" + "&txAmt&" where glId='"&glId&"'")
			else
				conn.execute("update glMaster set thisBal"&acPeriod&"=thisBal"&acPeriod&" - "&txAmt&" where glId='"&glId&"'")
			end if
			addUserLog "Add G/L Transaction"
			conn.committrans

			bdNum=bdNum+1
			thisString="bd1c"&bdNum&"="""&right("0"&mDay,2)&"/"&right("0"&m,2)&"/"&y&""""
			execute(thisString)
			thisString="bd2c"&bdNum&"="""&glId&""""
			execute(thisString)
			thisString="bd3c"&bdNum&"="""&glDes&""""
			execute(thisString)
			thisString="bd4c"&bdNum&"="""&txItem&""""
			execute(thisString)
			thisString="bd5c"&bdNum&"="""&txAmt&""""
			execute(thisString)
			thisString="bd6c"&bdNum&"="""&txType&""""
			execute(thisString)
			glId=""
			glDes="&nbsp;"
			txItem=""
			txAmt=""
			txType=""
		end if
	end if
else
    bdNum = 0
    txDate=today
end if
%>
<html>
<head>
<title>總賬入數</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function popup(filename){
  window.open (filename,'pop','width=500,height=550,statusbar=no,toolbar=no,resizable,scrollbars,dependent')
}

function checkDay(mDay){
  D=mDay.value;
  M=<%=m%>;
  Y=<%=y%>;
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

function formatNum(numform){
  if (isNaN(numform.value)||numform.value<0)
    return false;
  else
    return true;
}

function checkGlId(idform){
<%
sql = "select glId,glName from glMaster"
rs.open sql, conn, 3
%>
	codeArray = new Array(<%=rs.recordcount%>);
	desArray = new Array(<%=rs.recordcount%>);
<%
idx = 0
do while not rs.eof
	response.write "codeArray["&idx&"] = """&rs("glId")&""";"&vbcr
	response.write "desArray["&idx&"] = """&rs("glName")&""";"&vbcr
	rs.movenext
	idx=idx+1
loop
%>

	for (var i = 0; i < <%=rs.recordcount%>; i++) {
		if (idform.value==codeArray[i]){
			document.form1.glDes.value=desArray[i];
			document.all.tags( "td" )['glDes'].innerHTML=desArray[i];
			return true;
		}
	}
	alert('A/C Code not found');
	idform.value='';
	document.form1.glDes.value='';
	document.all.tags( "td" )['glDes'].innerHTML='';
	return false;
}

function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.mDay.value==""){
		reqField=reqField+", 日期";
		if (!placeFocus)
			placeFocus=formObj.mDay;
	}

	if (formObj.glId.value==""){
		reqField=reqField+", 賬戶號碼";
		if (!placeFocus)
			placeFocus=formObj.glId;
	}

	if (formObj.txAmt.value==""||!formatNum(formObj.txAmt)){
		reqField=reqField+", 金額";
		if (!placeFocus)
			placeFocus=formObj.txAmt;
	}

	if (formObj.txType.value!="D"&&formObj.txType.value!="C"){
		reqField=reqField+", 借貸";
		if (!placeFocus)
			placeFocus=formObj.txType;
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
<body leftMargin="0" topMargin="0" marginheight="0" marginwidth="0" bgcolor="#eeeef0" onload="form1.mDay.focus()">
<!-- #include file="menu.asp" -->
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form method="post" action="<%=request.servervariables("script_name")%>" name="form1">
<input type="hidden" name="from" value="<%=request.servervariables("script_name")%>">
<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="bottom" bgcolor="#87CEEB" height="17" align="center">
		<td class="n8" width=30>#</td>
		<td class="n8">日期</td>
		<td class="n8">賬戶號碼</td>
		<td class="n8">賬戶名稱</td>
		<td class="n8">內容</td>
		<td class="n8">金額</td>
		<td class="n8">D/C</td>
		<td class="n8"></td>
	</tr>
<%for idx = 1 to bdNum%>
	<input type="hidden" name="bd1c<%=idx%>" value="<%=eval("bd1c"&idx)%>">
	<input type="hidden" name="bd2c<%=idx%>" value="<%=eval("bd2c"&idx)%>">
	<input type="hidden" name="bd3c<%=idx%>" value="<%=eval("bd3c"&idx)%>">
	<input type="hidden" name="bd4c<%=idx%>" value="<%=eval("bd4c"&idx)%>">
	<input type="hidden" name="bd5c<%=idx%>" value="<%=eval("bd5c"&idx)%>">
	<input type="hidden" name="bd6c<%=idx%>" value="<%=eval("bd6c"&idx)%>">
	<tr>
		<td class="n10" align="center"><%=idx%></td>
		<td class="show" align="center"><%=eval("bd1c"&idx)%></td>
		<td class="show" align="center"><%=eval("bd2c"&idx)%></td>
		<td class="show"><%=eval("bd3c"&idx)%></td>
		<td class="show"><%if eval("bd4c"&idx)="" then response.write "&nbsp;" else response.write eval("bd4c"&idx) end if%></td>
		<td class="show" align="right"><%=formatNumber(eval("bd5c"&idx),2)%></td>
		<td class="show" align="center" style="background-color='<%if eval("bd6c"&idx)="C" then response.write "#ff4500" end if%>'"><%=eval("bd6c"&idx)%></td>
		<td></td>
    </tr>
<%next%>
	<input type="hidden" name="bdNum" value="<%=bdNum%>">
	<input type="hidden" name="glDes" value="<%=glDes%>">
	<tr>
		<td align="center" class="n10"><%=idx%></td>
		<td><input type="text" name="mDay" value="<%=mDay%>" size="2" maxlength="2" onblur="if(!checkDay(this)){this.value=''};">/<%=m%>/<%=y%>&nbsp;</td>
		<td><input type="text" name="glId" value="<%=glId%>" maxlength=4 size=4 onchange="return checkGlId(this)"></td>
		<td id="glDes" width="300"><%=glDes%></td>
		<td><input type="text" name="txItem" value="<%=txItem%>" size=40></td>
		<td><input type="text" name="txAmt" value="<%=txAmt%>" size=11 onblur="if(!formatNum(this)){this.value=''};"></td>
		<td><input type="text" name="txType" value="<%=txType%>" size=1 onclick="this.value=(this.value=='D')?'C':'D';if (this.value=='C'){this.style.backgroundColor='#ff4500'}else{this.style.backgroundColor=''};" onchange="this.value=this.value.toUpperCase();if (this.value=='C'){this.style.backgroundColor='#ff4500'}else{this.style.backgroundColor=''};if (this.value!='D'&&this.value!='C'){this.value=''};"></td>
        <td><input type="submit" value="新增" name="add" class="xbttn" onclick="return validating()"></td>
    </tr>
	<tr>
        <td colspan="8" align="right"><input type="submit" value="返回" name="back" class="sbttn"></td>
    </tr>
</table>
</form>
</center></div>
</body>
</html>
<%
conn.close
set conn=nothing
%>