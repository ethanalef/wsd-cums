<!-- #include file="../conn.asp" -->

<!-- #include file="../addUserLog.asp" -->
<%
if request.form("bye") <> "" then
   response.redirect "main.asp"
   
end if
if request.form("cancel") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next	
        response.write(xkind)
        response.end
	xdate = right(chkdate,4)&"/"&mid(chkdate,4,2)&"/"&left(chkdate,2)
        ydate = right(adjdate,4)&"/"&mid(adjdate,4,2)&"/"&left(adjdate,2) 
	if msg="" then
                conn.begintrans
		select case xkind
                       case "Cash"
 

                           if adjdate  <>"" then
                               addUserLog "�R�� �{���b  ��� : "&chkdate&" �� "&adjdate
                               conn.execute("update share set ldate="&adjdate&"  where memno='"&memno&"'  and ldate='"&xdate&"' and code='A3' ") 
                            end if  
                            
                            if adjsamt <>"" then
                                
                            addUserLog "�R�� �Ѫ��h�ڱb  ��� : "&chkdate&" (���B) $ "&samt&"  �� "&adjsamt
                            conn.execute("delete  share where memno='"&memno&"' and code='A3' and ldate='"&xdate&"' ") 
                            end if
		       case "Share"

                           if adjdate  <>"" then
                               addUserLog "�R�� �Ѫ��b  ��� : "&chkdate
                               conn.execute("delete share where memno='"&memno&"' and ldate='"&xdate&"' and code='B1' ") 
                            end if  
                           savettl = adjpamt+adjpint
        
                            addUserLog "�R�� �Ѫ��h�ڱb  ��� : "&chkdate&" (���B) $ "&actttl
                            conn.execute("delete  share  where memno='"&memno&"' and code='B0' and ldate='"&chkdate&"' ")

 
                 end select
		conn.committrans
                id =""

		msg = "�����w��s"
	end if
        id = ""
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next

    id = ""
   memno=""
   memcname=""
   memname=""
   key = 0
           
end if
if request.form("clrScr") <> "" then
   id = ""
   memno=""
   memcname=""
   memname=""
   
   
end if
  todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())

if id <>"" then
   memno = id
   pint = 0
   pamt  = 0 
   intamt = 0
   mstatus=""
end if
if request.form("bye") <> "" then
   id=""
	For Each Field in Request.Form
		TheString = Field & "= id"
		Execute(TheString)
	Next
    pint = 0
   pamt  = 0 
   intamt = 0
   cashamt = 0
   lostamt = 0
   todate = right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
end if

if request.form("Search")<>"" or id <>""  then
     
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next

		sql = "select * from memMaster where memNo='"& memNo & "' "
		set rs = server.createobject("ADODB.Recordset")
		rs.open sql, conn
		if not rs.eof then

			For Each Field in rs.fields
			if Field.name="memBday" or Field.name="firstAppointDate" or Field.name="memDate" or Field.name="Wdate" then
					TheString = "if rs(""" & Field.name & """)<>"""" then " & Field.name & " = right(""0""&day(rs(""" & Field.name & """)),2)&""/""&right(""0""&month(rs(""" & Field.name & """)),2)&""/""&year(rs(""" & Field.name & """)) end if"
				else
					TheString = Field.name & "= rs(""" & Field.name & """)"
				end if
				Execute(TheString)
			Next

                id = memno
                rs.close
                todate=right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
              

   
   pint = 0
   pamt  = 0 
   intamt = 0
   mstatus=""

                else
                  msg = "�������X���s�b "

                end if
         
end if

if request.form("Srh2")<>""   then
        msg=""
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
     
        xdate = right(chkdate,4)&"/"&mid(chkdate,4,2)&"/"&left(chkdate,2)
                       case "Back"
                            if ADJPAMT <> "" then
			      addUserLog "�R�� �h�ڱb �U�ڽs�� : "&lnnum&" ��� : "&chkdate&" (����) $ "&pamt
                              
                               conn.execute("delete loan where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='E0' ")
                                conn.execute("update loanrec set bal = bal - "&pamt&" where lnnum='"&lnnum&"' ")
                                  
                            end if
                            if ADJpint <> ""  then 
                                adduserlog "�R�� �h�ڱb �U�ڽs�� :"&lnnum&" ��� : "&chkdate&"(�Q��) $ "&pint
                                conn.execute("delete loan where lnnum='"&lnnum&"' and ldate='"&xdate&"' and code='F0' ")
                             end if 

                            savettl = adjpamt+adjpint
                            actttl  = pamt + pint
                            addUserLog "�R�� �Ѫ��h�ڱb  ��� : "&chkdate&"(���B) $ "&savettl                        
                            conn.execute("delete share  where memno='"&memno&"' and code='B1' and ldate='"&chkdate&"' ") 
        select case kind
               case "�{��"
                    stylefield =" code='A3' "
                    xkind = "Cash" 
               case "�h��"
 		    stylefield =" code='B0' "
                    xkind = "Share" 
        end select   
 
           set rs=conn.execute("select * from share where memno='"&id&"' and ldate ='"&xdate&"'  ")
       if not rs.eof then   
           do while not rs.eof 
              
              select case rs("code")
                     case "A3"
                          samt = rs("amount")
                     case "B0"
                          samt = rs("amount")
                     
              end select
           rs.movenext  
           loop
           adjpamt=""
           adjpint=""
                  
        else
                  
                  msg = "�b���䤣�� "
	
        end if   	
 	
end if



if request.form("action") <> "" then
	For Each Field in Request.Form
		TheString = Field & "= Request.Form(""" & Field & """)"
		Execute(TheString)
	Next
         
	xdate = right(chkdate,4)&"/"&mid(chkdate,4,2)&"/"&left(chkdate,2)
	if msg="" then
                conn.begintrans
		select case kind
                       case "cash"
			     addUserLog "�ק� �{���b  ��� : "&chkdate&"(�Ѫ�) $ "&samt&" �� "&adjsamt
			 
                            conn.execute("update share set amount="&adjsamt&" where memno='"&memno&"' and sdate='"&xdate&"' and code='A3' ")
                        
		       case "back"
			     addUserLog "�ק� �h�ѱb  ��� : "&chkdate&"(�Ѫ�) $ "&samt&" �� "&adjsamt
                            conn.execute("update share set amount="&adjsamt&" where memno='"&memno&"' and sdate='"&xdate&"' and code='B0' ")
                         

                 end select
		conn.committrans
                id =""

		msg = "�����w��s"
	end if
        id = ""
	For Each Field in Request.Form
		TheString = Field &"= id  "
		Execute(TheString)
	Next

 
       
else
   check=""
   chkdate =right("0"&day(date()),2)&"/"&right("0"&month(date()),2)&"/"&year(date())
end if
%>
<html>
<head>
<title>�Ѫ��Ӷ��ץ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">

<script language="JavaScript">
<!--
function popup(filename){
  window.open (filename,'pop','width=500,height=550,statusbar=no,toolbar=no,resizable,scrollbars,dependent')
}

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


function memberChange(){
	if (document.form1.memNo.value==''){
		document.form1.memName.value=''
		document.all.tags( "td" )['memName'].innerHTML=''
		document.form1.memGrade.value=''
		document.all.tags( "td" )['memGrade'].innerHTML=''
		document.form1.employCond.value=''
		document.all.tags( "td" )['employCond'].innerHTML=''
		document.form1.age.value=''
		document.all.tags( "td" )['age'].innerHTML=''
		document.form1.firstAppointDate.value=''
		document.all.tags( "td" )['firstAppointDate'].innerHTML=''
	}
	popup('pop_srhMemno.asp?key='+document.form1.memNo.value)
}



function validating(){
	formObj=document.form1;
	reqField="";
	placeFocus=false;

	if (formObj.adjsamt.value==""){
		reqField=reqField+",  �󥿪Ѫ� ";
		if (!placeFocus)
			placeFocus=formObj.adjsamt;
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
<% '<!-- #include file="menu.asp" -->  %>
<%if msg<>"" then %>
<div><center><font size="3" color="red"><%=msg%></font></center></div>
<% end if%>

<br>
<form name="form1" method="post" action="saveadj.asp">

<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="check" value="<%=check%>">
<input type="hidden" name="todate" value="<%=todate%>">
<input type="hidden" name="YYdate" value="<%=YYdate%>">
<input type="hidden" name="xstatus" value="<%=xstatus%>">
<input type="hidden" name="memName" value="<%=memName%>">
<input type="hidden" name="memcName" value="<%=memcName%>">
<input type="hidden" name="pamt" value="<%=pamt%>">
<input type="hidden" name="pint" value="<%=pint%>">
<input type="hidden" name="key" value="<%=key%>">
<input type="hidden" name="xkind" value="<%=xkind%>">
<div><center><font size="3" >�Ѫ��Ӷ��ץ�</font></center></div>
<BR>
<center>
<table border="0" cellspacing="0" cellpadding="0">

			<tr>
               		<td width=30></td>
			<td class="b12" align="left">�������X</td>
			<td width=50></td>
			<td><input type="text" name="memNo" value="<%=memNo%>" size="10" <%if id<>"" then response.write " onfocus=""form1.chkdate.focus();""" end if%>> 			
                        <%if id = "" then %>
		        <input type="button" value="���"  onclick="popup('pop_srhMemnoM.asp')" class="sbttn"  >
		        <input type="submit" value="�j�M" name="Search" class ="Sbttn">
 		        <%end if%>            
                        </TD>             			
			<tr>
               		<td width=30></td>
			<td class="b12" align="left">�����W��</td>
			<td width=50></td>
               		<td id="memName"><%=memName%></td>
                        <td id="memcName"><%=memcName%></td>
                        </tr>
                       <tr>
			<td width=30></td>
			<td class="b12" align="left">�ק���</td>
			<td width=50></td>
                        <%if samt = 0 then %>
			<td><input type="text" name="chkdate" value="<%=chkdate%>" size="10" onblur="if(!formatDate(this)){this.value=''};form1.repaydate.value=this.value"></TD>     
                        <%else%>
                        <td id="chkdate"><%=chkdate%></td>
                        <%end if%>
                        </tr>   
			<tr>
			<td width=30></td>
			<td align="right" class="b12">����</td>
			<td width="50"></td>
			
                        <%if samt=0  then %>
                        <td>  
			<select name="KIND" style="width:88px">		
			<option<% if KIND="Cash"  then response.write " selected" end if%>>�{��
                        <option<% if KIND="Back"  then response.write " selected" end if%>>�h��			
			</select>
                        <input type="submit" value="�j�M" name="Srh2" class ="Sbttn">
 			<input type="submit" value="����" name="clrScr" class ="Sbttn">
			<input type="submit" value="��^" name="bye" class ="Sbttn">
                        <input type="button" value="�d�߭ӤH�b" onclick="popup('viewlninfo.asp?key='+document.form1.memNo.value+'*'+document.form1.YYdate.value )" class="sbttn">                       
                        </td>  
                        <%else%>
                        <td  id = "kind"><%=kind%></td>
                        <%end if %> 

		        </tr>
                       

			<tr>
			<td width=30></td>
			<td align="right" class="b12">�Ѫ�</td>
			<td width="50"></td>
			<td id="samt"><%=samt%>			
		        </td>
	                </tr>

                        <%if key = 1 then %>
			<tr>
			<td width=30></td>
			<td align="right" class="b12">�󥿪Ѫ�</td>
			<td width="50"></td>
			<td><input type="text" name="adjsamt" value="<%=adjsamt%>" size="10" maxlength="10" ></td>

                        </tr>  
                        <%end if %>			<%if key <> ""  then %>
                        <tr>
                        <td></td>
                        <td></td>
                        <td></td>   
                        <td> 
        		<input type="submit" value="�R��" onclick="confirm('�T�w�R��?')" name="cancel" class="sbttn">
			<input type="submit" value="�x�s" onclick="return validating()&&confirm('�T�w�x�s?')" name="action" class="sbttn">
                        
			<input type="submit" value="����" name="clrScr" class ="Sbttn">
			<input type="submit" value="��^" name="bye" class ="Sbttn">
		        </td>
	                </tr>
                        <%end if%>
</table>       
</center>
</form>
</body>
</html>
                                                                                    