<!-- #include file="../conn.asp" -->
<%
 

   msg=""
   memno= ""
   cutdate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&(year(date()))
   ts1  = 0 
   ts2  =0
   ts3  = 0 
   ts4  =0   
   ts5  = 0 
   ts6  =0
   ts7  = 0 
   ts8  =0   
   ts9  = 0 
   ts10  =0   
   ts11  = 0 
   ts12  =0
   ts13  = 0 
   ts14  =0   
   ts15  = 0 
   ts16  =0
   ts17  = 0 
   ts18  =0   
   ts19  = 0 
   ts20  =0   
   
%>
<html>
<head>
<title>�������p�C�L</title>
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

function clearall(){
       formObj=document.form1;
          formObj.TS20.value = 0 ;
         formObj.TS20.checked=false;

}
function clearother(mark){
     formObj=document.form1;

         formObj.TS1.checked=false;
         formObj.TS2.checked=false; 
         formObj.TS3.checked=false;  
         formObj.TS4.checked=false; 
         formObj.TS5.checked=false; 
         formObj.TS6.checked=false; 
         formObj.TS7.checked=false; 
         formObj.TS8.checked=false; 
         formObj.TS9.checked=false; 
         formObj.TS10.checked=false; 
         formObj.TS11.checked=false;
         formObj.TS12.checked=false; 
         formObj.TS13.checked=false;  
         formObj.TS14.checked=false; 
         formObj.TS15.checked=false; 
         formObj.TS16.checked=false; 
         formObj.TS17.checked=false; 
         formObj.TS18.checked=false; 
         
     
}


function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;



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
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" bgcolor="#eeeef0" onload="form1.action.focus()">
<input type="hidden" name="id" value="<%=id%>">
<!-- #include file="menu.asp" -->


<input type="hidden" name="id" value="<%=id%>">
<form method="post" action="memstlstPrint.asp" name="form1">

<div align="center"><font size="3">�������p�C�L</font></center></div>
<center>
<br>
<font size="3"  face="�з���" >
<%if msg<>"" then %>
<div><center><font size="3" colour="red" ><%=msg%></font></center></div>
<% end if%>
<table border="0" cellpadding="0" cellspacing="0">
<tr>
		<td width="150" class="n8"><input type="checkbox" name="TS1" value="<%=ts1%>" <%if ts1<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts1=1;clearall();}")%> �۰���b(ALL)</td>
                <td width="20"> 
		<td width="150" class="n8"><input type="checkbox" name="TS2" value="<%=ts2%>" <%if ts2<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts2=1;clearall();}}")%> �۰���b(�Ѫ�)</td>
                <td width="20"> 
		<td width="150" class="n8"><input type="checkbox" name="TS3" value="<%=ts3%>" <%if ts3<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts3=1;clearall();}")%>�۰���b(�Ѫ�,�Q��)</td>
</tr>
<tr>
                <td width="150" class="n8"><input type="checkbox" name="TS4" value="<%=ts4%>" <%if ts4<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts4=1;clearall();}")%>�۰���b(�Ѫ�,����)</td>
                <td width="20"> 
		<td width="150" class="n8"><input type="checkbox" name="TS5" value="<%=ts5%>" <%if ts5<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts5=1;clearall();}")%>�۰���b(�Q��,����)</td>
                <td width="20"> 
                <td width="150" class="n8"><input type="checkbox" name="TS6" value="<%=ts6%>" <%if ts6>0 then response.write " checked" end if%>    onclick="if (this.checked){ts6=1;clearall();}")%>�w��,�Ȧ�</td>
</tr>
<tr>
                <td width="150" class="n8"><input type="checkbox" name="TS7" value="<%=ts7%>" <%if ts7<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts7=1;clearall();}")%>�b�b</td>
                <td width="20"> 
		<td width="150" class="n8"><input type="checkbox" name="TS8" value="<%=ts8%>" <%if ts8<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts8=1;clearall();}")%>�N��</td>
                <td width="20"> 
                <td width="150" class="n8"><input type="checkbox" name="TS9" value="<%=ts9%>" <%if ts9<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts9=1;clearall();}")%> IVA </td>
</tr>
<tr>
                <td width="150" class="n8"><input type="checkbox" name="TS10" value="<%=ts10%>" <%if ts10<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts10=1;clearall();}")%>�h��</td>
                <td width="20"> 
                <td width="150" class="n8"><input type="checkbox" name="TS11" value="<%=ts11%>" <%if ts11<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts11=1;clearall();}")%>�w��</td>               
                  <td width="20"> 
                <td width="150" class="n8"><input type="checkbox" name="TS12" value="<%=ts12%>" <%if ts12<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts12=1;clearall();}")%>�h�@</td>
                <td width="20"> 
</tr>
<tr> 
                <td width="150" class="n8"><input type="checkbox" name="TS13" value="<%=ts13%>" <%if ts13<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts13=1;clearall();}")%>�}��</td>

                <td width="20"> 
		<td width="150" class="n8"><input type="checkbox" name="TS14" value="<%=ts14%>" <%if ts14<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts14=1;clearall();}")%>���`</td>
                <td width="20"> 
                <td width="150" class="n8"><input type="checkbox" name="TS15" value="<%=ts15%>" <%if ts15<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts15=1;clearall();}")%>�s��</td>
               
</tr>
<tr>

                <td width="150" class="n8"><input type="checkbox" name="TS16" value="<%=ts16%>" <%if ts16<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts16=1;clearall();}")%>�Ȱ��Ȧ�</td>              
                <td width="20"> 
                <td width="150" class="n8"><input type="checkbox" name="TS17" value="<%=ts17%>" <%if ts17<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts17=1;clearall();}")%>�S�O�Ӯ�</td>
                  <td width="20"> 
                 <td width="150" class="n8"><input type="checkbox" name="TS18" value="<%=ts18%>" <%if ts18<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts18=1;clearall();}")%>�פ���y��b</td>              
</tr>
<tr>             
		<td width="150" class="n8"><input type="checkbox" name="TS19" value="<%=ts19%>" <%if ts19<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts19=1;clearall();}")%>�פ���y���`</td>
                <td width="20"> 
                <td width="150" class="n8"><input type="checkbox" name="TS20" value="<%=ts20%>" <%if ts20<>0 then response.write " checked" end if%>    onclick="if (this.checked){ts20=1;clearother(1);}")%>����</td>
                <td></td>
                <td></td>
</tr>
</table>  
<br>
<br>


<table border="0" cellpadding="0" cellspacing="0">
       <tr>
               	<td align="right" >�I�Ƥ��</td>
		<td width="10"></td>
		<td><input type="text" name="cutdate" value="<%=cutdate%>" size="10" onblur="if(!formatDate(this)){this.value=''};"></td>
        </tr>
	<tr>
		<td align="right" >��X</td>
		<td width="10"></td>
		<td>
			<select name="output" style="width:88px">
			<option value="html">Html
			<option value="text">Text
			<option value="word">Word
			<option value="excel">Excel
			</select>               
			<input type="submit" value="��X" onclick="return validating()&&confirm('�T�w��X?')"  name="action" class="sbttn">
		</td>
	</tr>
</table>
</form>
</center></div>
</body>
</html>