<%requiredLevel=3%>
<!-- #include file="../conn.asp" -->
<!-- #include file="../CheckUserStatus.asp" -->
<%
id=request("id")
For Each Field in Request.Form
	TheString = Field & "= Request.Form(""" & Field & """)"
	Execute(TheString)
Next

bdNum=cdbl(bdNum)
today=right("0"&day(date),2)&"/"&right("0"&month(date),2)&"/"&year(date)

if request("from") = Request.ServerVariables("script_name") then
	msg = ""
	if msg = "" then
		if request.form("Submit")<>"" then
			conn.begintrans
			Set bdRs = Server.CreateObject("ADODB.Recordset")
			for idx = 1 to bdNum
				sql = "select * from glTx where 1=0"
				bdRs.open sql, conn, 1, 3
				bdRs.addnew
				bdRs("TxDate") = request.form("bd2c"&idx)
				bdRs("glId") = request.form("bd3c"&idx)
				bdRs("txItem") = request.form("bd4c"&idx)
				bdRs("txAmt") = request.form("bd5c"&idx)
				bdRs("txType") = request.form("bd6c"&idx)
				bdRs.update
				Execute(TheString)
				bdRs.close

				'' *** update glMaster ***

				sql = "select * from glControl"
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.open sql, conn
				if mYear=acYear then y="this" else y="last" end if
				thisBal=y&"Bal"&rs("acPeriod")
				rs.close
				sql = "select "&thisBal&" from glMaster where glId='"&request.form("bd3c"&idx)&"'"
				rs.open sql, conn, 1, 3
				if request.form("bd6c"&idx)="D" then
					rs(0) = rs(0) + request.form("bd5c"&idx)
				else
					rs(0) = rs(0) - request.form("bd5c"&idx)
				end if
				rs.update
				rs.close
			next
			conn.committrans
			conn.close
			set conn=nothing
			response.redirect "glTx.asp"
		elseif request.form("add")<>"" then
			bdNum=bdNum+1
			TheString="bd2c"&bdNum+1&"=today"
			Execute(TheString)
		else
			checkDel = false
			for idx = 1 to bdNum
				if request.form("del"&idx)<>"" and msg="" then
					delNum=idx
					bdNum=bdNum-1
					checkDel = true
					exit for
				end if
			next
		end if
	end if
else
    bdNum = 0
    bd2c1=today
end if
focusmsg = "form1.bd2c"&bdNum+1&".focus()"
%>
<html>
<head>
<title>G/L Transaction</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<style>
<!--
.show0 {font-size: 10px; font-family: ms sans serif; border-left: #777777 1px solid; border-top: #777777 1px solid; border-bottom: #777777 1px solid; border-right:none; background-color:#eeeeee}
.show1 {font-size: 10px; font-family: ms sans serif; border-left: #777777 1px solid; border-top: #777777 1px solid; border-bottom: #777777 1px solid; border-right:none; background-color:#dddddd}
//-->
</style>
<script language="JavaScript">
<!--
function popup(filename){
  window.open (filename,'pop','width=500,height=550,statusbar=no,toolbar=no,resizable,scrollbars,dependent')
}

function selectPop(item,filename){
  newValue=showModalDialog(filename, '', 'resizable: no; help: no; status: no; scroll: no;');
  if (newValue != null) { item.value=newValue; }
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
    if (sCount != 2){
      return false;
    }
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
    dateform.value = strD+'/'+strM+'/'+strY;
    if (!valDate(strM, strD, strY))
      return false;
    else
      return true;
  }
}

function formatNum(numform){
  if (isNaN(numform.value)||numform.value<0)
    return false;
  else
    return true;
}

function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.bd2c<%=bdNum+1%>.value==""){
		reqField=reqField+", date";
		if (!placeFocus)
			placeFocus=formObj.bd2c<%=bdNum+1%>;
	}else{
		if (!formatDate(formObj.bd2c<%=bdNum+1%>)){
			reqField=reqField+", date";
			if (!placeFocus)
				placeFocus=formObj.bd2c<%=bdNum+1%>;
		}
	}

	if (formObj.bd3c<%=bdNum+1%>.value==""){
		reqField=reqField+", A/C no.";
		if (!placeFocus)
			placeFocus=formObj.bd3c<%=bdNum+1%>;
	}

	if (formObj.bd5c<%=bdNum+1%>.value==""||!formatNum(formObj.bd5c<%=bdNum+1%>)){
		reqField=reqField+", amount";
		if (!placeFocus)
			placeFocus=formObj.bd5c<%=bdNum+1%>;
	}

	if (formObj.bd6c<%=bdNum+1%>.value==""){
		reqField=reqField+", type";
		if (!placeFocus)
			placeFocus=formObj.bd6c<%=bdNum+1%>;
	}

    if (reqField){
        if (reqField.lastIndexOf(",")==0)
	        reqField = "Please fill in "+reqField.substring(2);
        else
	        reqField = "Please fill in "+reqField.substring(2,reqField.lastIndexOf(","))+' and '+reqField.substring(reqField.lastIndexOf(",")+2);
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
<body leftMargin="0" topMargin="0" marginheight="0" marginwidth="0" alink="#003399" link="#003399" vlink="#003399" bgcolor="#eeeef0" onload="<%=focusmsg%>">
<!-- #include file="menu.asp" -->
<center>
<%if msg<>"" then response.write "<font color=red>"&msg&"</font>" end if%>
<br>
<form method="post" action="glTxDetail.asp" name="form1">
<input type="hidden" name="From" value="<% =Request.servervariables("script_name") %>">
<table border="0" cellpadding="0" cellspacing="0">
	<tr valign="bottom" bgcolor="#87CEEB" height="17" align="center">
		<td class="n8" width=30>#</td>
		<td class="n8">Date</td>
		<td class="n8">Account No.</td>
		<td class="n8">Item</td>
		<td class="n8">Amount</td>
		<td class="n8">D/C</td>
		<td class="n8"></td>
	</tr>
	<input type="hidden" name="bdNum" value="<% =bdNum %>">
<%
for idx = 1 to bdNum+1
	if checkDel and idx >= delNum then
		ii=idx+1
	else
		ii=idx
	end if
%>
	<input type="hidden" name="bd1c<%=idx%>" value="<% =eval("bd1c"&ii) %>">
	<tr>
		<td align="center" class="n8"><%=idx%></td>
		<td><input type=text name="bd2c<%=idx%>" value="<% =eval("bd2c"&ii) %>" class="show<%=idx mod 2%>" maxlength=10 size=11 ondblclick="if(!this.value)this.value='<%=today%>'" onblur="if(!formatDate(this)){this.value=''};"></td>
		<td><input type=text name="bd3c<%=idx%>" value="<% =eval("bd3c"&ii) %>" class="show<%=idx mod 2%>" maxlength=12 size=12 ondblclick="popup('pop_searchAc.asp?editNum=<%=idx%>');"></td>
		<td><input type=text name="bd4c<%=idx%>" value="<% =eval("bd4c"&ii) %>" class="show<%=idx mod 2%>" size=40 onfocus="form1.bd5c<%=idx%>.focus();"></td>
		<td><input type=text name="bd5c<%=idx%>" value="<% =eval("bd5c"&ii) %>" class="show<%=idx mod 2%>" maxlength=25 size=25 onblur="if(!formatNum(this)){this.value=''};"></td>
		<td><input type=text name="bd6c<%=idx%>" value="<% =eval("bd6c"&ii) %>" class="show<%=idx mod 2%>" size=4 onclick="this.value=(this.value=='D')?'C':'D';"></td>
        <td>
<%if idx>bdNum then%>
<input type="submit" value="Add" name="add" class="xbttn" onclick="return validating()">
<%else%>
<input type="submit" value="Del" name="del<%=idx%>" class="xbttn" onclick="return confirm('Delete this record?')" style="width:26">
<%end if%>
		</td>
    </tr>
<%
next
%>
	<tr>
        <td colspan="6" align="right"><input type="submit" value="Save" name="submit" class="sbttn"></td>
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