<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
if request.form("back") <> "" then
   response.redirect "main.asp"
   
end if

  todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())


if request.form("bye") <> "" then
   id=""
	For Each Field in Request.Form
		TheString = Field & "= id"
		Execute(TheString)
	Next

  todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())   
 
   chkdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1) 
end if


if request.form("Search")<>"" or id <>""  then
 cashamt=0
                      intamt=0
                      saveamt=0
                     
        msg=""
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next

                set rs=conn.execute("select a.memno,a.memname,a.memcname,a.mstatus,b.* from memmaster a,dividend b where a.memno='"&memno&"'  and b.memno=a.memno  ")
                if not rs.eof then
                   For Each Field in rs.fields 
		   TheString = Field.name & "= rs(""" & Field.name & """)"
	           Execute(TheString)
		   Next
                   id = memno  
                   xbank = bank  
                   todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())   
 
                   chkdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1)                         
                else
           
                   id = "" 
                   msg =" ���O������� "
                   memno = ""                           
                end if
                rs.close
    
             

 
                opt = 0

else 




if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next

        select case bank
               case "�Ѫ�"
                    mbank = "S"
               case "�Ȧ���b"
                    mbank = "B"
               case "�䲼"
                    mbank = "C"
               case "�Ȱ�����"
                    mbank = "H"
        end select
        if xbank <> mbank then
           conn.execute("update dividend set bank ='"&mbank&"' where memno='"&id&"' ")
        end if
   id=""
	For Each Field in Request.Form
		TheString = Field & "= id"
		Execute(TheString)
	Next
 
        	conn.begintrans
                
                 
 
                     
                                             
		conn.committrans
		msg = "�����w��s"

        id = ""
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next

  todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())   
 
   chkdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1) 
       
else
   id = ""
   todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())   
   
   chkdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date())-1) 
end if
end if
%>
<html>
<head>
<title>�������t�إ�</title>
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






function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.memNo.value==""){
		reqField=reqField+", �����s��";
		if (!placeFocus)
			placeFocus=formObj.memNo;
	}








    if (reqField){
        if (reqField.lastIndexOf(",")==0)
	        reqField = "�ж�J"+reqField.substring(2);
        else
	        reqField = "�ж�J"+reqField.substring(2,reqField.lastIndexOf(","))+'��'+reqField.substring(reqField.lastIndexOf(",")+2);
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
<br>
<form name="form1" method="post" action="shpaypro.asp">

<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="chkdate" value="<%=chkdate%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="xbank" value="<%=xbank%>">


<div><center><font size="3">�������t�إ�</font></center></div>
<br>

<center>
<%if msg<>"" then %>
<div><center><font size="3"><%=msg%></font></center></div>
<BR>
<% end if%>

<table border="0" cellspacing="0" cellpadding="0">
 
			<tr>
               		<td width=30></td>
			<td class="b12" align="left">�������X</td>
			<td width=50></td>
			<td><input type="text" name="memNo" value="<%=memNo%>" size="10" <%if id<>"" then response.write " onfocus=""form1.bank.focus();""" end if%>>
			<%if id = "" then %>	
		         <input type="submit"  value="�j�M" name="Search" class ="Sbttn">
			<% end if %>
                        </TD>
                        </tr>
			<tr>
                		<td width=30></td>
				<td class="b12" align="left">�����W��</td>
				<td width=50></td>
				<td><input type="text" name="memName" value="<%=memName%>" size="30"<%if id<>"" then response.write " onfocus=""form1.bank.focus();""" end if%>></td> 
			</tr>
                       </tr>
			<tr>
                		<td width=30></td>
				<td class="b12" align="left"></td>
				<td width=50></td>
				<td><input type="text" name="memcName" value="<%=memcName%>" size="30"<%if id<>"" then response.write " onfocus=""form1.bank.focus();""" end if%>></td> 
			</tr>

			<tr>
                		<td width=30></td>
				<td class="b12" align="left">�Ѯ����B</td>
				<td width=50></td>
				<td><input type="text" name="dividend" value="<%=dividend%>" size="30"<%if id<>"" then response.write " onfocus=""form1.bank.focus();""" end if%>></td> 
			</tr>
<tr>

                <td width=12"></td>
     		<td><font size="2" >�������p</formt></td>
                <td width=21></td>
		<td>
                     
			<select name="bank">
                        <option></option>
			<option<%if bank="S" then response.write " selected" end if%>>�Ѫ�</option>
			<option<%if bank="B" then response.write " selected" end if%>>�Ȧ���b</option>
                        <option<%if bank="C" then response.write " selected" end if%>>�䲼</option>
			<option<%if bank="H" then response.write " selected" end if%>>�Ȱ�����</option>

			</select>
                 
		</td>
       </tr>
	<tr>
		<td width=30></td>
		<td colspan="3" align="right">
		<% if id <> "" then %>
               
				<input type="submit" value="�x�s" onclick="return validating()&&confirm('�T�w�x�s?')" name="action" class="sbttn">
		<input type="button" value="�d�߶U��" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.chkdate.value )" class="sbttn">					
               
				
		<% end if %>
               
		<input type="submit" value="����" name="bye" class="sbttn">
		<input type="submit" value="��^" name="back" class="sbttn">
		</td>
	</tr>
 
</table>       
</center>
</form>
</body>
</html>
