<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
loan=request("loan")
if request.form("back") <> "" then
	response.redirect "meetingNotesDetail.asp?id="&loan
end if

if request.form("action") <> "" then
	conn.begintrans
	conn.execute("delete from meetingNotes1 where rpId="&loan)
	if request("thisApp")<>"" then
		A = split(request("thisApp"),",",-1,1)
		if isarray(A) then
			if (ubound(A) >= 0) then
				for i = 0 to ubound(A)
					conn.execute("insert into meetingNotes1 (rpId,appId) values ("&loan&","&A(i)&")")
				next
			end if
		end if
	end if
	conn.committrans
	response.redirect "meetingNotesDetail.asp?id="&loan
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

function moveAllCol(fromCol,toCol){
	for (ii=0; ii<document.form1[fromCol].length; ii++) {
		var oOption = document.createElement("OPTION");
		oOption.text=document.form1[fromCol].options[ii].text;
		oOption.value=document.form1[fromCol].options[ii].value;
		document.form1[toCol].add(oOption);
	}
	for (ii=document.form1[fromCol].length-1; ii>=0; ii--) {
		document.form1[fromCol].remove(ii);
	}
}

function selectCol(){
    if (document.form1.thisApp.length>0) {
        for (ii=0; ii<document.form1.thisApp.length; ii++) {
            document.form1.thisApp.options[ii].selected=true;
        }
    }
	return true;
}
//  -->
</script>
</head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0">
<!-- #include file="menu.asp" -->
<br>
<center>
<form name="form1" method="post" action="<%=Request.servervariables("script_name")%>">
<input type="hidden" name="loan" value="<%=loan%>">
<table border="0" cellspacing="0" cellpadding="0">
	<tr valign="top">
		<td class="b8">
			本會議否決之貸款<br>
			<select name="thisApp" size="20" style="width:300px" multiple>
<%
set rs=conn.execute("select a.* from loanApp a,meetingNotes1 b where a.uid=b.appId and b.rpId="&loan&" order by a.uid")
do while not rs.eof
	response.write "<option value="&rs("uid")&">"&rs("memNo")&" - "&rs("memName")&" - "&formatnumber(rs("loanAmt"),2)
	rs.movenext
loop
%>
			</select>
		</td>
		<td align="center" valign="middle" width="30">
			<input type="button" value="<<" class="sbttn" name="allToLeft" onclick="moveAllCol('otherApp','thisApp')"><br><br>
			<input type="button" value=" < " class="sbttn" name="toLeft" onclick="moveCol('otherApp','thisApp')"><br><br>
			<input type="button" value=" > " class="sbttn" name="toRight" onclick="moveCol('thisApp','otherApp')"><br><br>
			<input type="button" value=">>" class="sbttn" name="allToRight" onclick="moveAllCol('thisApp','otherApp')">
		</td>
		<td class="b8">
			所有己否決之貸款<br>
			<select name="otherApp" size="20" style="width:300px" multiple>
<%
set rs=conn.execute("select * from loanApp where (firstApproval='Rejected' or secondApproval='Rejected') and deleted=0 and uid not in (select appId from meetingNotes1) order by uid")
do while not rs.eof
	response.write "<option value="&rs("uid")&">"&rs("memNo")&" - "&rs("memName")&" - "&formatnumber(rs("loanAmt"),2)
	rs.movenext
loop
%>
			</select>
		</td>
	</tr>
	<tr>
        <td colspan="19" align="right" height="30">
			<input type="submit" value="儲存" onclick="return confirm('確定儲存?')&&selectCol()" name="action" class="sbttn">
			<input type="submit" value="返回" name="back" class="sbttn">
        </td>
	</tr>
</table>
</center>
</body>
</html>
