<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<!-- #include file="../addUserLog.asp" -->
<%
id = request("id")

if request.form("back") <> "" then
	response.redirect "meetingNotes.asp"
elseif request.form("print") <> "" then
	response.redirect "meetingNotesPrint.asp?id="&id
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
			sql = "select top 1 * from meetingNotes order by uid desc"
		else
			sql = "select * from meetingNotes where uid=" & id
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
			addUserLog "Add Meeting Notes"
		else
			addUserLog "Modify Meeting Notes"
		end if
		For Each Field in rs.fields
			if Field.name="uid" or Field.name="deleted" then
			elseif Field.name="rpDate" or Field.name="lastRpDate" then
				TheString = "if " & Field.name & "<>"""" then rs(""" & Field.name & """) = right(" & Field.name & ",4)&""/""&mid(" & Field.name & ",4,2)&""/""&left(" & Field.name & ",2) else rs(""" & Field.name & """)=null end if"
				Execute(TheString)
			elseif Field.name="interview" then
				TheString = "if " & Field.name & "<>"""" then rs(""" & Field.name & """) = cdbl(" & Field.name & ") else rs(""" & Field.name & """)=0 end if"
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
		sql = "select * from meetingNotes where uid=" & id
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if rs.eof then
			response.redirect "meetingNotes.asp"
		else
			For Each Field in rs.fields
				if Field.name="rpDate" or Field.name="lastRpDate" or Field.name="nextRpDate" then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next
		end if
	else
		present=""
		set rs = conn.execute("select * from handleParty where status=1")
		do while not rs.eof
			present=present&","&rs("handleName")
			rs.movenext
		loop
		set rs = conn.execute("select max(rpDate) from meetingNotes where deleted=0")
		if not isnull(rs(0)) then lastRpDate=right("0"&day(rs(0)),2)&"/"&right("0"&month(rs(0)),2)&"/"&year(rs(0)) end if
		set rs = conn.execute("select max(rpNo) from meetingNotes where deleted=0 and year(rpDate)="&year(date()))
		if isnull(rs(0)) then rpNo=1 else rpNo=rs(0)+1 end if
		present=right(present,len(present)-1)
		rpDate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
		rs.close
		set rs=nothing
	end if
end if
%>
<html>
<head>
<title>會議紀錄</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function moveCol(fromCol,toCol){
    for (ii=0; ii<document.form1[fromCol].length; ii++) {
        if (document.form1[fromCol].options[ii].selected) {
            var oOption = document.createElement("OPTION");
            oOption.text=document.form1[fromCol].options[ii].text;
            oOption.value=document.form1[fromCol].options[ii].value;
            document.form1[toCol].add(oOption);
        }
    }
    for (ii=document.form1[fromCol].length-1; ii>=0; ii--) {
        if (document.form1[fromCol].options[ii].selected) {
            document.form1[fromCol].remove(ii);
        }
    }
}

function selectCol(){
    if (document.form1.present.length>0) {
        for (ii=0; ii<document.form1.present.length; ii++) {
            document.form1.present.options[ii].selected=true;
        }
    }
    if (document.form1.absent.length>0) {
        for (ii=0; ii<document.form1.absent.length; ii++) {
            document.form1.absent.options[ii].selected=true;
        }
    }
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
		reqField=reqField+", 開會日期";
		if (!placeFocus)
			placeFocus=formObj.rpDate;
	}

	if (formObj.rpNo.value==""){
		reqField=reqField+", 會議次數";
		if (!placeFocus)
			placeFocus=formObj.rpNo;
	}

	if (!formatDate(formObj.lastRpDate)){
		reqField=reqField+", 前次開會日期";
		if (!placeFocus)
			placeFocus=formObj.lastRpDate;
	}

	if (!formatDate(formObj.nextRpDate)){
		reqField=reqField+", 下次開會日期";
		if (!placeFocus)
			placeFocus=formObj.nextRpDate;
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
<form name="form1" method="post" action="meetingNotesDetail.asp">
<input type="hidden" name="id" value="<%=id%>">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td class="b8" align="right">開會日期</td>
		<td width=10></td>
		<td><input type="text" name="rpDate" value="<%=rpDate%>" size="10" maxlength="10" onblur="if(!formatDate(this)){this.value=''};"></td>
	</tr>
	<tr>
		<td class="b8" align="right">會議次數</td>
		<td width=10></td>
		<td><input type="text" name="rpNo" value="<%=rpNo%>" size="3" maxlength="3" onblur="if(!formatNum(this)){this.value=''};"></td>
	</tr>
	<tr>
		<td class="b8" align="right">會議類別</td>
		<td width=10></td>
		<td>
			<select name="rpType">
			<option<%if rpType="常會" then response.write " selected" end if%>>常會</option>
			<option<%if rpType="特別會議" then response.write " selected" end if%>>特別會議</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="b8" align="right">開會時間</td>
		<td width=10></td>
		<td>
			<select name="rpTime">
			<option<%if rpTime="下午" then response.write " selected" end if%>>下午</option>
			<option<%if rpTime="上午" then response.write " selected" end if%>>上午</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="b8" align="right">前次開會日期</td>
		<td width=10></td>
		<td><input type="text" name="lastRpDate" value="<%=lastRpDate%>" size="10" maxlength="10" onblur="if(!formatDate(this)){this.value=''};"></td>
	</tr>
	<tr>
		<td class="b8" align="right">修正</td>
		<td width=10></td>
		<td><input type="text" name="amenment" value="<%=amenment%>" size="50" maxlength="100"></td>
	</tr>
	<tr>
		<td class="b8" align="right">出席委員</td>
		<td width=10></td>
		<td>
			<table border="0" cellspacing="0" cellpadding="0">
				<tr valign="top">
					<td class="b8">
						出席<br>
						<select name="present" size="3" style="width:120px" multiple>
<%
if present<>"" then
	A = split(present,",",-1,1)
	if isarray(A) then
		if (ubound(A) >= 0) then
			for i = 0 to ubound(A)
%>
			            <option><%=A(i)%>
<%
			next
		end if
	end if
end if
%>
						</select>
					</td>
					<td align="center" valign="middle" width="30">
						<br><input type="button" value=" < " class="sbttn" name="toLeft" onclick="moveCol('absent','present')"><br>
						<input type="button" value=" > " class="sbttn" name="toRight" onclick="moveCol('present','absent')">
					</td>
					<td class="b8">
						缺席<br>
						<select name="absent" size="3" style="width:120px" multiple>
<%
if absent<>"" then
	A = split(absent,",",-1,1)
	if isarray(A) then
		if (ubound(A) >= 0) then
			for i = 0 to ubound(A)
%>
			            <option><%=A(i)%>
<%
			next
		end if
	end if
end if
%>
					</select>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td class="b8" align="right">出席之其他人仕</td>
		<td width=10></td>
		<td><input type="text" name="attendee" value="<%=attendee%>" size="50" maxlength="50"></td>
	</tr>
	<tr>
		<td class="b8" align="right">接見人數</td>
		<td width=10></td>
		<td><input type="text" name="interview" value="<%=interview%>" size="3" maxlength="3" onblur="if(!formatNum(this)){this.value=''};"></td>
	</tr>
	<tr>
		<td class="b8" align="right">溫習章程</td>
		<td width=10></td>
		<td><input type="text" name="overview" value="<%=overview%>" size="50" maxlength="50"></td>
	</tr>
	<tr>
		<td class="b8" align="right">下次開會日期</td>
		<td width=10></td>
		<td><input type="text" name="nextRpDate" value="<%=nextRpDate%>" size="10" maxlength="10" onblur="if(!formatDate(this)){this.value=''};"></td>
	</tr>
	<tr>
		<td class="b8" align="right">下次開會時間</td>
		<td width=10></td>
		<td>
			<select name="nextRpTime">
			<option<%if nextRpTime="下午" then response.write " selected" end if%>>下午</option>
			<option<%if nextRpTime="上午" then response.write " selected" end if%>>上午</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="b8" align="right">其他行動</td>
		<td width=10></td>
		<td><textarea name="otherAction" rows="4" cols="50"><%=otherAction%></textarea></td>
	</tr>
	<tr>
		<td colspan="3" align="right" height="30">
			<%if session("userLevel")<>5 then%>
			<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')&&selectCol()" name="action" class="sbttn">
			<%end if%>
			<input type="submit" value="返回" name="back" class="sbttn">
<%if id <> "" then %>
			&nbsp;&nbsp;<input type="submit" value="列印" name="print" class="sbttn">
<% end if %>
		</td>
	</tr>
</table>
</form>
<%if id<>"" then%>
<a href="meeting0.asp?loan=<%=id%>">批準之貸款</a> |
<a href="meeting1.asp?loan=<%=id%>">否決之貸款</a> |
<a href="meeting2.asp?loan=<%=id%>">撤回資金</a> |
<a href="meeting3.asp?loan=<%=id%>">放棄期滿副抵抽品</a> |
<a href="meeting4.asp?loan=<%=id%>">批準延期</a> |
<a href="meeting5.asp?loan=<%=id%>">否決延期</a>
<%end if%>
</center>
</body>
</html>
