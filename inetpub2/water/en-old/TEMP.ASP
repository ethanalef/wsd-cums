
<div style="z-index: 99; left: 0px; width: 1100px; position: absolute; top: 100px;  height: 68px">
<table id="table1" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td width=10></TD>            
		<td class="b8" align="left">英文名稱</td>
		<td width=50></td>
		<td><input type="text" name="memName" value="<%=memName%>" size="50" maxlength="50"></td>
		<td width=10></TD>            
		<td class="b8" align="left">中文名稱</td>
		<td width=35></td>
		<td><input type="text" name="memcName" value="<%=memcName%>" size="50" maxlength="50"></td>   
        </tr>
</table>
</DIV>
<div style="z-index: 100; left: 0px; width: 1100px; position: absolute; top: 120px;  height: 68px">
<table id="table2" border="0" cellspacing="0" cellpadding="0">	
         <tr> 
		<td width=10></td>
		<td class="b8" align="left">身分證號碼</td>
		<td width=37></td>
		<td><input type="text" name="memHKID" value="<%=memHKID%>" size="20" maxlength="15"></td>
		<td width=10></td>
		<td class="b8" align="right">出生日期 (dd/mm/yy)</td>
		<td width=20></td>
		<td><input type="text" name="memBday" value="<%=memBday%>" size="15" maxlength="10"></td>
                <td width=10></td>
		<td class="b8" align="right">性別</td>
		<td width=20></td>
		<td>
			<select name="memGender">
			<option<%if memGender="M" then response.write " selected" end if%>>M</option>
			<option<%if memGender="F" then response.write " selected" end if%>>F</option>
			</select>
		</td>
        </tr>
</TABLE>
</DIV>
<div style="z-index: 101; left: 0px; width: 136px; position: absolute; top: 140px;  height: 68px">
<table id="table31" border="0" cellspacing="0" cellpadding="0">	
       <tr>
		<td width=10></td>  
               <td></td>
       </tr>
       <tr>	                 
         	<td width=10></td>
		<td class="b8" align="left">地址</td>
       </tr>
       <tr>
		<td width=10></td>  
               <td></td>
       </tr>
</table>
</div>
<div style="z-index: 102; left: 137px; width: 158px; position: absolute; top: 140px;   height: 64px">
<table id="table32" border="0" cellspacing="0" cellpadding="0">	
      <tr>
	       <td><input type=" text" name="memAddr1" value="<%=memAddr1%>" size="50" maxlength="50"></td>
      </tr>
      <tr>
 	      <td><input type="text" name="memAddr2" value="<%=memAddr2%>" size="50" maxlength="50"></td>
      </tr>
      <tr>
	      <td><input type="text" name="memAddr3" value="<%=memAddr3%>" size="50" maxlength="50"></td>
      </tr>		
</tanle>   
</div>
<div style="z-index: 103; left: 295px; width: 100px; position: absolute; top: 140px;  height: 65px; text-align: center">
<table id="table33" border="0" cellspacing="0" cellpadding="0">	
	<TR>
		
     		<td width=10></td>
		<td class="b8" align="left">家居電話號碼</td>
		<td width=10></td>
	</tr>
        <tr>
                <td width=10></td>
		<td class="b8" align="left">辦公室電話號碼</td>
		<td width=10></td>
	</tr>

		<td width=10></TD>
		<td class="b8" align="left">手提電話號碼</td>
		<td width=10></td>
         </tr>
</table>
</div>
<div style="z-index: 104; left: 396px; width: 349px; position: absolute; top: 140px;height: 63px">
<table id="table34" border="0" cellspacing="0" cellpadding="0">	
	<tr>
		<td width=10></td>
		<td><input type="text" name="memContactTel" value="<%=memContactTel%>" size="20" maxlength="20"></td>	
        <TR>
		<td width=10></td>
		<td><input type="text" name="memOfficeTel" value="<%=memofficetel%>" size="20" maxlength="20"></td>
	</tr>
	<TR>
		<td width=10></td>
		<td><input type="text" name="memMobile" value="<%=memMobile%>" size="20" maxlength="20"></td>	
	</TR>
</TABLE>
</DIV>
<div style="z-index: 105; left: 746px; width: 95px; position: absolute; top: 140px;height: 72px">
<table id="table35" border="0" cellspacing="0" cellpadding="0">	


	<tr>
	        <td width=10></td>
		<td class="b8" align="left">銀行</td>	
	</TR>
        <TR>
		<td width=10></td>
		<td class="b8" align="left">分行</td>

	</TR>
		<td width=10></td>	
                <TD></TD>
		<td width=40></td>
	<TR>	
		<td width=10></td>	
		<td class="b8" align="left">銀行賬號</td>
	</TR>
		
</table>
</DIV>
<div style="z-index: 106; left: 843px; width: 166px; position: absolute; top: 140px;height: 69px">
<table id="table36" border="0" cellspacing="0" cellpadding="0">	
         <TR>
		<td width=10></td>
		<td><input type="text" name="bnk" value="<%=bnk%>" size="20" maxlength="20"></td>
        </TR>
	<TR> 
		<td width=10></td>
		<td><input type="text" name="bch" value="<%=bch%>" size="20" maxlength="20"></td>	
        </TR>
        <TR>
		<td width=10></td>
		<td><input type="text" name="bacct" value="<%=bacct%>" size="20" maxlength="20"></td>	
        </TR>
</TABLE>
</DIV>
<div style="z-index: 107; left: 0px; width: 136px; position: absolute; top: 160px;  height: 68px">
<table id-"table5" border="0" cellspacing="0" cellpadding="0">
	<tr>

		<td width=10></td>
		<td class="b8" align="left">職級</td>
		<td width=75></td>
		<td><input type="text" name="memGrade" value="<%=memGrade%>" size="20" maxlength="8"></td>
		<td width=20></td>
		<td class="b8" align="right">部門</td>
		<td width=30></td>
		<td><input type="text" name="memSection" value="<%=memSection%>" size="20" maxlength="10"></td>
		<td width=10></td>
		<td class="b8" align="left">POS</td>
		<td width=20></td>
		<td><input type="text" name="mempos" value="<%=memPOS%>" size="20" maxlength="10"></td>
		<td width=10></td>
		<td class="b8" align="right">RANK</td>
		<td width=33></td>
		<td><input type="text" name="memRANK" value="<%=memRANK%>" size="20" maxlength="10"></td>
	</tr>
</table>
</DIV>
<div style="z-index: 108; left: 0px; width: 136px; position: absolute; top: 180px;  height: 68px">
<table id-"table6" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td width=10></td>
		<td class="b8" align="right">CU Designation</td>
		<td width=18></td>
		<td><input type="text" name="cuDesignation" value="<%=cuDesignation%>" size="20" maxlength="20"></td>
		<td width=20></td>
		<td class="b8" align="right">教育委員</td>
		<td width=10></td>
		<td><input type="text" name="EdComm" value="<%=EdComm%>" size="20" maxlength="20"></td>
	
	
</tr>
</table>
</DIV>
<div style="z-index: 109; left: 0px; width: 136px; position: absolute; top: 200Px;  height: 68px">
<table id-"table8" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td width=10></td>
		<td class="b8" align="left">Treas Ref No</td>
		<td width=30></td>
		<td><input type="text" name="treasRefNo" value="<%=treasRefNo%>" size="20" maxlength="20"></td>
		<td width=20></td>
		<td class="b8" align="right">受顧條件</td>
		<td width=10></td>
		<td><input type="text" name="employCond" value="<%=employCond%>" size="20" maxlength="20"></td>
	
	</tr>
</table>
</DIV>
<div style="z-index: 110; left: 0px; width: 136px; position: absolute; top: 220px;  height: 68px">
<table id-"table9" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td width=10></td>
		<td class="b8" align="left">入職日期 </td>
		<td width=48></td>
		<td><input type="text" name="firstAppointDate" value="<%=firstAppointDate%>" size="20" maxlength="20"></td>
		<td width=20></td>
		<td class="b8" align="right">入社日期</td>
		<td width=10></td>
		<td><input type="text" name="memDate" value="<%=memDate%>" size="20" maxlength="20"></td>
		<td width=10></td>
		<td class="b8" align="right">退社日期</td>
		<td width=10></td>
		<td><input type="text" name="WDate" value="<%=WDate%>" size="20" maxlength="20"></td>	
	</tr>
</table>
</DIV>
<div style="z-index: 111; left: 0px; width: 136px; position: absolute; top: 240px;  height: 68px">
<table id-"table7" border="0" cellspacing="0" cellpadding="0">

	<tr>
                <td width=10></td> 
		<td class="b8" align="left">擔保人社員編號</td>
		<td width=20></td>
		<td><input type="text" name="memG1" value="<%=memG1%>" size="20"></td>
		<td width=10></td>
		<td><input type="text" name="memG2" value="<%=memG2%>" size="20"></td>
		<td width=10></td>
		<td><input type="text" name="memG3" value="<%=memG3%>" size="20"></td>
		<td width=10></td>
	</tr>


	<tr>
  		<td width=10></td> 
		<td class="b8" align="left">擔保其他社員之編號</td>
		<td width=10></td>
		<td><input type="text" name="memG4O1" value="<%=memG4O1%>" size="15"></td>
		<td width=10></td>
		<td><input type="text" name="memG4O2" value="<%=memG4O2%>" size="15"></td>
		<td width=10></td>
		<td><input type="text" name="memG4O3" value="<%=memG4O3%>" size="15"></td>
	</tr>
</table>
</DIV>
<div style="z-index: 112; left: 0px; width: 136px; position: absolute; top: 260px;  height: 68px">
<table id-"table4" border="0" cellspacing="0" cellpadding="0">
	        <td width=10></td>
		<td class="b8" align="left">電郵地址</td>
		<td width=20></td>
		<td><input type="text" name="mememail" value="<%=mememail%>" size="100" maxlength="100"></td>
		
	
</TABLE>
</DIV>
<div style="z-index: 113; left: 0px; width: 136px; position: absolute; top: 280px;  height: 68px">
<table id-"table8" border="0" cellspacing="0" cellpadding="0">
            <TR>
	        <td width=10></td>
		<td class="b8" align="left">備註欄</td>
		<td width=33></td>
		<td><input type="text" name="remark" value="<%=remark%>" size="100" maxlength="100"></td>
		
             </TR> 
	
</TABLE>
</DIV>
<div style="z-index: 114; left: 0px; width: 136px; position: absolute; top: 300px;  height: 68px">
<table id-"table9" border="0" cellspacing="0" cellpadding="0" >
         <TR>
 		<td width=10></td>
		<td class="b8" align="left">受益人 1.</td>
		<td width=20></td>
		<td><input type="text" name="B1" value="<%=B1%>" size="20" maxlength="20"></td>
  		<td width=30></td>
		<td class="b8" align="right">身分證號碼</td>
		<td width=20></td>
		<td><input type="text" name="B1ID" value="<%=B1ID%>" size="20" maxlength="20"></td>    
  		<td width=20></td>
		<td class="b8" align="right">關係</td>
		<td width=20></td>
		<td><input type="text" name="B1relation" value="<%=B1relation%>" size="20" maxlength="20"></td>    
        </TR>

</table>
</DIV>
<div style="z-index: 115; left: 0px; width: 136px; position: absolute; top: 320px;  height: 68px">
<table id-"table19" border="0" cellspacing="0" cellpadding="0" >
		<td width=10></td>
		<td class="b8" align="left">地址</td>
		<td width=45></td>
		<td><input type="text" name="B1ADD" value="<%=B1ADD%>" size="100" maxlength="100"></td>
</TABLE>
</DIV>
<div style="z-index: 116; left: 0px; width: 136px; position: absolute; top: 340px;  height: 68px">
<table id-"table9" border="0" cellspacing="0" cellpadding="0" >
         <TR>
 		<td width=10></td>
		<td class="b8" align="left">受益人 2.</td>
		<td width=20></td>
		<td><input type="text" name="B2" value="<%=B2%>" size="20" maxlength="20"></td>
  		<td width=30></td>
		<td class="b8" align="right">身分證號碼</td>
		<td width=20></td>
		<td><input type="text" name="B2ID" value="<%=B2ID%>" size="20" maxlength="20"></td>    
  		<td width=20></td>
		<td class="b8" align="right">關係</td>
		<td width=20></td>
		<td><input type="text" name="B2relation" value="<%=B2relation%>" size="20" maxlength="20"></td>    
        </TR>

</table>
</DIV>
<div style="z-index: 117; left: 0px; width: 136px; position: absolute; top: 360px;  height: 68px">
<table id-"table19" border="0" cellspacing="0" cellpadding="0" >
		<td width=10></td>
		<td class="b8" align="left">地址</td>
		<td width=45></td>
		<td><input type="text" name="B2ADD" value="<%=B2ADD%>" size="100" maxlength="100"></td>
</TABLE>
</DIV>
<div style="z-index: 118; left: 0px; width: 136px; position: absolute; top: 380px;  height: 68px">
<table id-"table20" border="0" cellspacing="0" cellpadding="0">       
	<tr>
		
		<td colspan="3" align="right">
			<%if session("userLevel")<>5 and session("userLevel")<>1 and session("userLevel")<>6 then%>
			<input type="submit" value="儲存" onclick="return validating()&&confirm('確定儲存?')" name="action" class="sbttn">
			<%end if%>
			<input type="submit" value="返回" name="back" class="sbttn">
		</td>
	</tr>
</table>
</DIV>