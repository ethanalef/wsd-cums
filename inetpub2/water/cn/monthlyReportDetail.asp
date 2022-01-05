<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
id = request("id")

if request.form("back") <> "" then
	response.redirect "monthlyReport.asp"
elseif request.form("print") <> "" then
	response.redirect "monthlyReportPrint.asp?id="&id
elseif request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
    set rs = server.createobject("ADODB.Recordset")
	msg = ""

	if msg="" then
		conn.begintrans
		if id = "" then
			sql = "select top 1 * from monthlyReport order by uid desc"
		else
			sql = "select * from monthlyReport where uid=" & id
		end if
		rs.open sql, conn, 2, 2
		if id = "" then
			if rs.eof then
				id = 1
			else
				id = rs("uid") + 1
			end if
			rs.addnew
			rs("uid") = id
			addUserLog "Add Monthly Report"
		else
			addUserLog "Modify Monthly Report"
		end if
		For Each Field in rs.fields
			if Field.name="uid" or Field.name="deleted" then
			elseif Field.name="rpDate" or Field.name="StartDate" or Field.name="EndDate" then
				TheString = "if " & Field.name & "<>"""" then rs(""" & Field.name & """) = right(" & Field.name & ",4)&""/""&mid(" & Field.name & ",4,2)&""/""&left(" & Field.name & ",2) else rs(""" & Field.name & """)=null end if"
				Execute(TheString)
			else
				TheString = "rs(""" & Field.name & """) = " & Field.name
				Execute(TheString)
			end if
		Next
		rs.update
		conn.committrans
		msg = "資料已更新"
	end if
else
	if id <> "" then
		sql = "select * from monthlyReport where uid=" & id
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if rs.eof then
			response.redirect "monthlyReport.asp"
		else
			For Each Field in rs.fields
				if Field.name="rpDate" or Field.name="StartDate" or Field.name="EndDate" then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
		end if
	else
		set rs = conn.execute("select min(rpDate) from meetingNotes where deleted=0 and rpDate>(select max(endDate) from monthlyReport where deleted=0)")
		if not isnull(rs(0)) then StartDate=right("0"&day(rs(0)),2)&"/"&right("0"&month(rs(0)),2)&"/"&year(rs(0)) end if
		set rs = conn.execute("select max(rpDate) from meetingNotes where deleted=0")
		if not isnull(rs(0)) then EndDate=right("0"&day(rs(0)),2)&"/"&right("0"&month(rs(0)),2)&"/"&year(rs(0)) end if
		rpDate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
		rs.close
		set rs=nothing
	end if
end if
%>
<html>
<head>
<title>董事會報告書</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
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

function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (!formatDate(formObj.rpDate)){
		reqField=reqField+", 報告書日期";
		if (!placeFocus)
			placeFocus=formObj.rpDate;
	}

	if (!formatDate(formObj.startDate)){
		reqField=reqField+", 開始日期";
		if (!placeFocus)
			placeFocus=formObj.startDate;
	}

	if (!formatDate(formObj.endDate)){
		reqField=reqField+", 終止日期";
		if (!placeFocus)
			placeFocus=formObj.endDate;
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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.rpDate.focus()">
<!-- #include file="menu.asp" -->
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form name="form1" method="post" action="monthlyReportDetail.asp">
<input type="hidden" name="id" value="<%=id%>">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td class="b8" align="right">報告書日期</td>
		<td width=10></td>
		<td><input type="text" name="rpDate" value="<%=rpDate%>" size="10" maxlength="10" onblur="if(!formatDate(this)){this.value=''};"></td>
	</tr>
	<tr>
		<td class="b8" align="right">開始日期</td>
		<td width=10></td>
		<td><input type="text" name="startDate" value="<%=startDate%>" size="10" maxlength="10" onblur="if(!formatDate(this)){this.value=''};"></td>
	</tr>
	<tr>
		<td class="b8" align="right">終止日期</td>
		<td width=10></td>
		<td><input type="text" name="endDate" value="<%=endDate%>" size="10" maxlength="10" onblur="if(!formatDate(this)){this.value=''};"></td>
	</tr>
	<tr>
		<td class="b8" align="right">連續三次不出席</td>
		<td width=10></td>
		<td>
			<select name="absent">
			<option<%if absent="0" then response.write " selected" end if%> value="0">無</option>
			<option<%if absent="1" then response.write " selected" end if%> value="1">有</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="b8" align="right">已採取行動</td>
		<td width=10></td>
		<td><input type="text" name="actions" value="<%=actions%>" size="50" maxlength="500"></td>
	</tr>
	<tr>
		<td class="b8" align="right">拒絕貸款理由</td>
		<td width=10></td>
		<td><input type="text" name="rejectReason" value="<%=rejectReason%>" size="50" maxlength="500"></td>
	</tr>
	<tr>
		<td class="b8" align="right">其他事項</td>
		<td width=10></td>
		<td><input type="text" name="others" value="<%=others%>" size="50" maxlength="500"></td>
	</tr>
	<tr>
		<td colspan="3" align="right" height="30">
			<%if session("userLevel")<>5 then%>
			<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
			<%end if%>
			<input type="submit" value="返回" name="back" class="sbttn">
<%if id <> "" then %>
			&nbsp;&nbsp;<input type="submit" value="列印" name="print" class="sbttn">
<% end if %>
		</td>
	</tr>
</table>
</form>
</center>
</body>
</html>
