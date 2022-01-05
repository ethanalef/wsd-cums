<%@ CodePage=950 %>

<!--#include file="db.asp"-->
<%

' Table Level SQL
Const ewSqlSelect = "SELECT * FROM [userRights]"
Const ewSqlWhere = ""
Const ewSqlGroupBy = ""
Const ewSqlHaving = ""
Const ewSqlOrderBy = ""
Const ewSqlOrderBySessions = ""
Const ewSqlKeyWhere = "[PID] = @PID"
Const ewSqlMasterSelect = "SELECT * FROM [loginUser]"
Const ewSqlMasterWhere = ""
Const ewSqlMasterGroupBy = ""
Const ewSqlMasterHaving = ""
Const ewSqlMasterOrderBy = ""
Const ewSqlMasterFilter = "[uid] = @User_Fk"
Const ewSqlDetailFilter = "[User_Fk] = @User_Fk"
Const ewSqlUserIDFilter = ""
%>

<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private, no-cache, no-store, must-revalidate"
%>
<%

' Initialize common variables
x_PID = Null: ox_PID = Null: z_PID = Null
x_Member1 = Null: ox_Member1 = Null: z_Member1 = Null
x_Member2 = Null: ox_Member2 = Null: z_Member2 = Null
x_Member3 = Null: ox_Member3 = Null: z_Member3 = Null
x_Member4 = Null: ox_Member4 = Null: z_Member4 = Null
x_Loan1 = Null: ox_Loan1 = Null: z_Loan1 = Null
x_Loan2 = Null: ox_Loan2 = Null: z_Loan2 = Null
x_Loan3 = Null: ox_Loan3 = Null: z_Loan3 = Null
x_Loan4 = Null: ox_Loan4 = Null: z_Loan4 = Null
x_Loan5 = Null: ox_Loan5 = Null: z_Loan5 = Null
x_Loan6 = Null: ox_Loan6 = Null: z_Loan6 = Null
x_Loan7 = Null: ox_Loan7 = Null: z_Loan7 = Null
x_Loan8 = Null: ox_Loan8 = Null: z_Loan8 = Null
x_Loan9 = Null: ox_Loan9 = Null: z_Loan9 = Null
x_cLoan1 = Null: ox_cLoan1 = Null: z_cLoan1 = Null
x_cLoan2 = Null: ox_cLoan2 = Null: z_cLoan2 = Null
x_cLoan3 = Null: ox_cLoan3 = Null: z_cLoan3 = Null
x_AutoPay1 = Null: ox_AutoPay1 = Null: z_AutoPay1 = Null
x_AutoPay2 = Null: ox_AutoPay2 = Null: z_AutoPay2 = Null
x_AutoPay3 = Null: ox_AutoPay3 = Null: z_AutoPay3 = Null
x_AutoPay4 = Null: ox_AutoPay4 = Null: z_AutoPay4 = Null
x_AutoPay5 = Null: ox_AutoPay5 = Null: z_AutoPay5 = Null
x_AutoPay6 = Null: ox_AutoPay6 = Null: z_AutoPay6 = Null
x_AutoPay7 = Null: ox_AutoPay7 = Null: z_AutoPay7 = Null
x_AutoPay8 = Null: ox_AutoPay8 = Null: z_AutoPay8 = Null
x_AutoPay9 = Null: ox_AutoPay9 = Null: z_AutoPay9 = Null
x_AutoPay11 = Null: ox_AutoPay11 = Null: z_AutoPay11 = Null
x_AutoPay12 = Null: ox_AutoPay12 = Null: z_AutoPay12 = Null
x_Saving1 = Null: ox_Saving1 = Null: z_Saving1 = Null
x_Saving2 = Null: ox_Saving2 = Null: z_Saving2 = Null
x_Saving3 = Null: ox_Saving3 = Null: z_Saving3 = Null
x_Saving4 = Null: ox_Saving4 = Null: z_Saving4 = Null
x_Saving5 = Null: ox_Saving5 = Null: z_Saving5 = Null
x_Saving6 = Null: ox_Saving6 = Null: z_Saving6 = Null
x_Saving7 = Null: ox_Saving7 = Null: z_Saving7 = Null
x_Saving8 = Null: ox_Saving8 = Null: z_Saving8 = Null
x_Saving9 = Null: ox_Saving9 = Null: z_Saving9 = Null
x_MemAcct1 = Null: ox_MemAcct1 = Null: z_MemAcct1 = Null
x_Reporting1 = Null: ox_Reporting1 = Null: z_Reporting1 = Null
x_Reporting2 = Null: ox_Reporting2 = Null: z_Reporting2 = Null
x_Reporting3 = Null: ox_Reporting3 = Null: z_Reporting3 = Null
x_Reporting4 = Null: ox_Reporting4 = Null: z_Reporting4 = Null
x_Reporting5 = Null: ox_Reporting5 = Null: z_Reporting5 = Null
x_Reporting6 = Null: ox_Reporting6 = Null: z_Reporting6 = Null
x_Reporting7 = Null: ox_Reporting7 = Null: z_Reporting7 = Null
x_Reporting8 = Null: ox_Reporting8 = Null: z_Reporting8 = Null
x_Reporting9 = Null: ox_Reporting9 = Null: z_Reporting9 = Null
x_Reporting10 = Null: ox_Reporting10 = Null: z_Reporting10 = Null
x_Reporting11 = Null: ox_Reporting11 = Null: z_Reporting11 = Null
x_Other4 = Null: ox_Other4 = Null: z_Other4 = Null
x_Other3 = Null: ox_Other3 = Null: z_Other3 = Null
x_Other2 = Null: ox_Other2 = Null: z_Other2 = Null
x_Other1 = Null: ox_Other1 = Null: z_Other1 = Null
x_User_Fk = Null: ox_User_Fk = Null: z_User_Fk = Null
%>
<%
Response.Buffer = True

' Load key from QueryString
x_PID = Request.QueryString("PID")

' Get action
sAction = Request.Form("a_edit")
If sAction = "" Or IsNull(sAction) Then
	sAction = "I"	' Display with input box
Else

	' Get fields from form
	x_PID = Request.Form("x_PID")
	x_Member1 = Request.Form("x_Member1")
	x_Member2 = Request.Form("x_Member2")
	x_Member3 = Request.Form("x_Member3")
	x_Member4 = Request.Form("x_Member4")
	x_Loan1 = Request.Form("x_Loan1")
	x_Loan2 = Request.Form("x_Loan2")
	x_Loan3 = Request.Form("x_Loan3")
	x_Loan4 = Request.Form("x_Loan4")
	x_Loan5 = Request.Form("x_Loan5")
	x_Loan6 = Request.Form("x_Loan6")
	x_Loan7 = Request.Form("x_Loan7")
	x_Loan8 = Request.Form("x_Loan8")
	x_Loan9 = Request.Form("x_Loan9")
	x_cLoan1 = Request.Form("x_cLoan1")
	x_cLoan2 = Request.Form("x_cLoan2")
	x_cLoan3 = Request.Form("x_cLoan3")
	x_AutoPay1 = Request.Form("x_AutoPay1")
	x_AutoPay2 = Request.Form("x_AutoPay2")
	x_AutoPay3 = Request.Form("x_AutoPay3")
	x_AutoPay4 = Request.Form("x_AutoPay4")
	x_AutoPay5 = Request.Form("x_AutoPay5")
	x_AutoPay6 = Request.Form("x_AutoPay6")
	x_AutoPay7 = Request.Form("x_AutoPay7")
	x_AutoPay8 = Request.Form("x_AutoPay8")
	x_AutoPay9 = Request.Form("x_AutoPay9")
	x_AutoPay11 = Request.Form("x_AutoPay11")
	x_AutoPay12 = Request.Form("x_AutoPay12")
	x_Saving1 = Request.Form("x_Saving1")
	x_Saving2 = Request.Form("x_Saving2")
	x_Saving3 = Request.Form("x_Saving3")
	x_Saving4 = Request.Form("x_Saving4")
	x_Saving5 = Request.Form("x_Saving5")
	x_Saving6 = Request.Form("x_Saving6")
	x_Saving7 = Request.Form("x_Saving7")
	x_Saving8 = Request.Form("x_Saving8")
	x_Saving9 = Request.Form("x_Saving9")
	x_MemAcct1 = Request.Form("x_MemAcct1")
	x_Reporting1 = Request.Form("x_Reporting1")
	x_Reporting2 = Request.Form("x_Reporting2")
	x_Reporting3 = Request.Form("x_Reporting3")
	x_Reporting4 = Request.Form("x_Reporting4")
	x_Reporting5 = Request.Form("x_Reporting5")
	x_Reporting6 = Request.Form("x_Reporting6")
	x_Reporting7 = Request.Form("x_Reporting7")
	x_Reporting8 = Request.Form("x_Reporting8")
	x_Reporting9 = Request.Form("x_Reporting9")
	x_Reporting10 = Request.Form("x_Reporting10")
	x_Reporting11 = Request.Form("x_Reporting11")
	x_Other4 = Request.Form("x_Other4")
	x_Other3 = Request.Form("x_Other3")
	x_Other2 = Request.Form("x_Other2")
	x_Other1 = Request.Form("x_Other1")
	x_User_Fk = Request.Form("x_User_Fk")
End If

' Check if valid key
If x_PID = "" Or IsNull(x_PID) Then Response.Redirect "userRightslist.asp"

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case sAction
	Case "I": ' Get a record to display
		If Not LoadData() Then ' Load Record based on key
			Response.Clear
			Response.write("No records found")
			response.End()
		End If
	Case "U": ' Update

		If EditData() Then ' Update Record based on key
				End If
		
		Response.Write("Update Record Successful<br>")
		Response.Write("<span><a href=""userRightsEdit.asp?PID="& Server.URLEncode(x_PID) &""">Back To User Rights Edit</a></span><br>")
			Response.Write("<span><a href=loginUserList.asp>Back To User List</a></span>")
	
			response.End()
End Select
%>

<script type="text/javascript">
<!--
function EW_checkMyForm(EW_this) {
return true;
}
//-->
</script>

<link href="wsdscu.css" rel="stylesheet" type="text/css" />

<p><strong>Edit : User Rights</strong><br>
  <br><a href="loginUserlist.asp">Back to Main</a></span></p>
<%
If Session(ewSessionMessage) <> "" Then
%>
<p><span class="ewmsg"><%= Session(ewSessionMessage) %></span></p>
<%
	Session(ewSessionMessage) = "" ' Clear message
End If
%>
<form name="fuserRightsedit" id="fuserRightsedit" action="userRightsedit.asp" method="post" onSubmit="return EW_checkMyForm(this);">
  <p> 
    <input type="hidden" name="a_edit" value="U">
  <table class="ewTable">
    <tr> 
      <td class="ewTableHeader"><span>User Fk</span></td>
      <td class="ewTableAltRow"><span id="cb_x_User_Fk"> 
        <% Response.Write x_User_Fk %>
        <input type="hidden" id="x_User_Fk" name="x_User_Fk" value="<%= x_User_Fk %>">
        </span></td>
    </tr>
  </table>
  <p>
    <input type="submit" name="btnAction" id="btnAction3" value="EDIT">
  <table width="100%" border="1">
    <tr> 
      <td valign="top"><p>社員資料</p>
        <table class="ewTable">
          <tr> 
            <td class="ewTableHeader"><span>PID<span class='ewmsg'>&nbsp;*</span></span></td>
            <td class="ewTableAltRow"><span id="cb_x_PID"> 
              <% Response.Write x_PID %>
              <input type="hidden" id="x_PID" name="x_PID" value="<%= x_PID %>">
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>加入新社員</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Member1"> 
              <% If x_Member1 = True Then %>
              <input type="checkbox" name="x_Member1" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Member1" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>社員資料修正</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Member2"> 
              <% If x_Member2 = True Then %>
              <input type="checkbox" name="x_Member2" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Member2" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>轉換聯絡人建立 </span></td>
            <td class="ewTableAltRow"><span id="cb_x_Member3"> 
              <% If x_Member3 = True Then %>
              <input type="checkbox" name="x_Member3" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Member3" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>新社員開戶建立</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Member4"> 
              <% If x_Member4 = True Then %>
              <input type="checkbox" name="x_Member4" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Member4" value="1">
              <% End If %>
              </span></td>
          </tr>
        </table>
        <p>&nbsp; </p>
      </td>
      <td valign="top"><p>貸款</p>
        <table class="ewTable">
          <tr> 
            <td class="ewTableHeader"><span>貸款申請</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Loan1"> 
              <% If x_Loan1 = True Then %>
              <input type="checkbox" name="x_Loan1" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Loan1" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>新貸款建立</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Loan2"> 
              <% If x_Loan2 = True Then %>
              <input type="checkbox" name="x_Loan2" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Loan2" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>貸款修正</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Loan3"> 
              <% If x_Loan3 = True Then %>
              <input type="checkbox" name="x_Loan3" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Loan3" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>貸款列印</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Loan4"> 
              <% If x_Loan4 = True Then %>
              <input type="checkbox" name="x_Loan4" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Loan4" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>現金還款</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Loan5"> 
              <% If x_Loan5 = True Then %>
              <input type="checkbox" name="x_Loan5" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Loan5" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>股金還款</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Loan6"> 
              <% If x_Loan6 = True Then %>
              <input type="checkbox" name="x_Loan6" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Loan6" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>貸款退款至股金操作 </span></td>
            <td class="ewTableAltRow"><span id="cb_x_Loan7"> 
              <% If x_Loan7 = True Then %>
              <input type="checkbox" name="x_Loan7" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Loan7" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>貸款細項列印</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Loan8"> 
              <% If x_Loan8 = True Then %>
              <input type="checkbox" name="x_Loan8" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Loan8" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>貸款細項修正</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Loan9"> 
              <% If x_Loan9 = True Then %>
              <input type="checkbox" name="x_Loan9" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Loan9" value="1">
              <% End If %>
              </span></td>
          </tr>
        </table></td>
      <td valign="top"><p>清數操作</p>
        <table class="ewTable">
          <tr> 
            <td class="ewTableHeader"><span>循環貸款</span></td>
            <td class="ewTableAltRow"><span id="cb_x_cLoan1"> 
              <% If x_cLoan1 = True Then %>
              <input type="checkbox" name="x_cLoan1" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_cLoan1" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>現金清數</span></td>
            <td class="ewTableAltRow"><span id="cb_x_cLoan2"> 
              <% If x_cLoan2 = True Then %>
              <input type="checkbox" name="x_cLoan2" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_cLoan2" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>股金清數</span></td>
            <td class="ewTableAltRow"><span id="cb_x_cLoan3"> 
              <% If x_cLoan3 = True Then %>
              <input type="checkbox" name="x_cLoan3" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_cLoan3" value="1">
              <% End If %>
              </span></td>
          </tr>
        </table></td>
      <td valign="top"><p>自動轉帳</p>
        <table class="ewTable">
          <tr> 
            <td class="ewTableHeader"><span>轉帳建立</span></td>
            <td class="ewTableAltRow"><span id="cb_x_AutoPay1"> 
              <% If x_AutoPay1 = True Then %>
              <input type="checkbox" name="x_AutoPay1" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_AutoPay1" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>特別個案轉帳輸入操作</span></td>
            <td class="ewTableAltRow"><span id="cb_x_AutoPay2"> 
              <% If x_AutoPay2 = True Then %>
              <input type="checkbox" name="x_AutoPay2" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_AutoPay2" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>銀行轉帳試算</span></td>
            <td class="ewTableAltRow"><span id="cb_x_AutoPay3"> 
              <% If x_AutoPay3 = True Then %>
              <input type="checkbox" name="x_AutoPay3" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_AutoPay3" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>銀行轉帳磁碟建立</span></td>
            <td class="ewTableAltRow"><span id="cb_x_AutoPay4"> 
              <% If x_AutoPay4 = True Then %>
              <input type="checkbox" name="x_AutoPay4" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_AutoPay4" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>銀行脫期建立</span></td>
            <td class="ewTableAltRow"><span id="cb_x_AutoPay5"> 
              <% If x_AutoPay5 = True Then %>
              <input type="checkbox" name="x_AutoPay5" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_AutoPay5" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>銀行轉帳超額細明表</span></td>
            <td class="ewTableAltRow"><span id="cb_x_AutoPay6"> 
              <% If x_AutoPay6 = True Then %>
              <input type="checkbox" name="x_AutoPay6" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_AutoPay6" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>銀行轉帳過數</span></td>
            <td class="ewTableAltRow"><span id="cb_x_AutoPay7"> 
              <% If x_AutoPay7 = True Then %>
              <input type="checkbox" name="x_AutoPay7" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_AutoPay7" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>庫房脫期建立</span></td>
            <td class="ewTableAltRow"><span id="cb_x_AutoPay8"> 
              <% If x_AutoPay8 = True Then %>
              <input type="checkbox" name="x_AutoPay8" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_AutoPay8" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>庫房轉帳試算</span></td>
            <td class="ewTableAltRow"><span id="cb_x_AutoPay9"> 
              <% If x_AutoPay9 = True Then %>
              <input type="checkbox" name="x_AutoPay9" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_AutoPay9" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>庫房過數</span></td>
            <td class="ewTableAltRow"><span id="cb_x_AutoPay11"> 
              <% If x_AutoPay11 = True Then %>
              <input type="checkbox" name="x_AutoPay11" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_AutoPay11" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>特別個案轉帳試算</span></td>
            <td class="ewTableAltRow"><span id="cb_x_AutoPay12"> 
              <% If x_AutoPay12 = True Then %>
              <input type="checkbox" name="x_AutoPay12" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_AutoPay12" value="1">
              <% End If %>
              </span></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td valign="top"><p>股金</p>
        <table class="ewTable">
          <tr> 
            <td class="ewTableHeader"><span>股息計算操作</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Saving1"> 
              <% If x_Saving1 = True Then %>
              <input type="checkbox" name="x_Saving1" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Saving1" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>股息列印</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Saving2"> 
              <% If x_Saving2 = True Then %>
              <input type="checkbox" name="x_Saving2" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Saving2" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>股息分配建立</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Saving3"> 
              <% If x_Saving3 = True Then %>
              <input type="checkbox" name="x_Saving3" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Saving3" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>股息分配列印</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Saving4"> 
              <% If x_Saving4 = True Then %>
              <input type="checkbox" name="x_Saving4" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Saving4" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>股息過數</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Saving5"> 
              <% If x_Saving5 = True Then %>
              <input type="checkbox" name="x_Saving5" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Saving5" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>退股建立</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Saving6"> 
              <% If x_Saving6 = True Then %>
              <input type="checkbox" name="x_Saving6" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Saving6" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>現金存款建立</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Saving7"> 
              <% If x_Saving7 = True Then %>
              <input type="checkbox" name="x_Saving7" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Saving7" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>股金列印</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Saving8"> 
              <% If x_Saving8 = True Then %>
              <input type="checkbox" name="x_Saving8" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Saving8" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>股金細項修正</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Saving9"> 
              <% If x_Saving9 = True Then %>
              <input type="checkbox" name="x_Saving9" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Saving9" value="1">
              <% End If %>
              </span></td>
          </tr>
        </table></td>
      <td valign="top"><p>個人戶口</p>
        <table class="ewTable">
          <tr> 
            <td class="ewTableHeader"><span>社員資料查詢</span></td>
            <td class="ewTableAltRow"><span id="cb_x_MemAcct1"> 
              <% If x_MemAcct1 = True Then %>
              <input type="checkbox" name="x_MemAcct1" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_MemAcct1" value="1">
              <% End If %>
              </span></td>
          </tr>
        </table></td>
      <td valign="top"><p>報表</p>
        <table class="ewTable">
          <tr> 
            <td class="ewTableHeader"><span>個人資料列表</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Reporting1"> 
              <% If x_Reporting1 = True Then %>
              <input type="checkbox" name="x_Reporting1" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Reporting1" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>呆賬報告</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Reporting2"> 
              <% If x_Reporting2 = True Then %>
              <input type="checkbox" name="x_Reporting2" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Reporting2" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>冷戶報告</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Reporting3"> 
              <% If x_Reporting3 = True Then %>
              <input type="checkbox" name="x_Reporting3" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Reporting3" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>社員分組/組員列表 </span></td>
            <td class="ewTableAltRow"><span id="cb_x_Reporting4"> 
              <% If x_Reporting4 = True Then %>
              <input type="checkbox" name="x_Reporting4" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Reporting4" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>社員轉帳資料列表 </span></td>
            <td class="ewTableAltRow"><span id="cb_x_Reporting5"> 
              <% If x_Reporting5 = True Then %>
              <input type="checkbox" name="x_Reporting5" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Reporting5" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>社員生日名單</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Reporting6"> 
              <% If x_Reporting6 = True Then %>
              <input type="checkbox" name="x_Reporting6" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Reporting6" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>現金帳列表 </span></td>
            <td class="ewTableAltRow"><span id="cb_x_Reporting7"> 
              <% If x_Reporting7 = True Then %>
              <input type="checkbox" name="x_Reporting7" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Reporting7" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>庫房帳列表 </span></td>
            <td class="ewTableAltRow"><span id="cb_x_Reporting8"> 
              <% If x_Reporting8 = True Then %>
              <input type="checkbox" name="x_Reporting8" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Reporting8" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>銀行帳列表</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Reporting9"> 
              <% If x_Reporting9 = True Then %>
              <input type="checkbox" name="x_Reporting9" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Reporting9" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>其他帳列</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Reporting10"> 
              <% If x_Reporting10 = True Then %>
              <input type="checkbox" name="x_Reporting10" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Reporting10" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>每月帳統計列表</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Reporting11"> 
              <% If x_Reporting11 = True Then %>
              <input type="checkbox" name="x_Reporting11" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Reporting11" value="1">
              <% End If %>
              </span></td>
          </tr>
        </table></td>
      <td valign="top"><p>其他</p>
        <table class="ewTable">
          <tr> 
            <td class="ewTableHeader"><span>用戶使用紀錄</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Other4"> 
              <% If x_Other4 = True Then %>
              <input type="checkbox" name="x_Other4" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Other4" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>用戶管理</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Other3"> 
              <% If x_Other3 = True Then %>
              <input type="checkbox" name="x_Other3" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Other3" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>資料庫輸入</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Other2"> 
              <% If x_Other2 = True Then %>
              <input type="checkbox" name="x_Other2" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Other2" value="1">
              <% End If %>
              </span></td>
          </tr>
          <tr> 
            <td class="ewTableHeader"><span>資料庫輸出</span></td>
            <td class="ewTableAltRow"><span id="cb_x_Other1"> 
              <% If x_Other1 = True Then %>
              <input type="checkbox" name="x_Other1" value="1" checked>
              <% Else %>
              <input type="checkbox" name="x_Other1" value="1">
              <% End If %>
              </span></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <p>&nbsp; 
</form>

<%
conn.Close ' Close Connection
Set conn = Nothing
%>
<%

'-------------------------------------------------------------------------------
' Function LoadData
' - Load Data based on Key Value
' - Variables setup: field variables

Function LoadData()
	Dim rs, sSql, sFilter
	sFilter = ewSqlKeyWhere
	If Not IsNumeric(x_PID) Then
		LoadData = False
		Exit Function
	End If
	sFilter = Replace(sFilter, "@PID", AdjustSql(x_PID)) ' Replace key value
	sSql = ewBuildSql(ewSqlSelect, ewSqlWhere, ewSqlGroupBy, ewSqlHaving, ewSqlOrderBy, sFilter, "")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSql, conn
	If rs.Eof Then
		LoadData = False
	Else
		LoadData = True
		rs.MoveFirst

		' Get the field contents
		x_PID = rs("PID")
		x_Member1 = rs("Member1")
		x_Member2 = rs("Member2")
		x_Member3 = rs("Member3")
		x_Member4 = rs("Member4")
		x_Loan1 = rs("Loan1")
		x_Loan2 = rs("Loan2")
		x_Loan3 = rs("Loan3")
		x_Loan4 = rs("Loan4")
		x_Loan5 = rs("Loan5")
		x_Loan6 = rs("Loan6")
		x_Loan7 = rs("Loan7")
		x_Loan8 = rs("Loan8")
		x_Loan9 = rs("Loan9")
		x_cLoan1 = rs("cLoan1")
		x_cLoan2 = rs("cLoan2")
		x_cLoan3 = rs("cLoan3")
		x_AutoPay1 = rs("AutoPay1")
		x_AutoPay2 = rs("AutoPay2")
		x_AutoPay3 = rs("AutoPay3")
		x_AutoPay4 = rs("AutoPay4")
		x_AutoPay5 = rs("AutoPay5")
		x_AutoPay6 = rs("AutoPay6")
		x_AutoPay7 = rs("AutoPay7")
		x_AutoPay8 = rs("AutoPay8")
		x_AutoPay9 = rs("AutoPay9")
		x_AutoPay11 = rs("AutoPay11")
		x_AutoPay12 = rs("AutoPay12")
		x_Saving1 = rs("Saving1")
		x_Saving2 = rs("Saving2")
		x_Saving3 = rs("Saving3")
		x_Saving4 = rs("Saving4")
		x_Saving5 = rs("Saving5")
		x_Saving6 = rs("Saving6")
		x_Saving7 = rs("Saving7")
		x_Saving8 = rs("Saving8")
		x_Saving9 = rs("Saving9")
		x_MemAcct1 = rs("MemAcct1")
		x_Reporting1 = rs("Reporting1")
		x_Reporting2 = rs("Reporting2")
		x_Reporting3 = rs("Reporting3")
		x_Reporting4 = rs("Reporting4")
		x_Reporting5 = rs("Reporting5")
		x_Reporting6 = rs("Reporting6")
		x_Reporting7 = rs("Reporting7")
		x_Reporting8 = rs("Reporting8")
		x_Reporting9 = rs("Reporting9")
		x_Reporting10 = rs("Reporting10")
		x_Reporting11 = rs("Reporting11")
		x_Other4 = rs("Other4")
		x_Other3 = rs("Other3")
		x_Other2 = rs("Other2")
		x_Other1 = rs("Other1")
		x_User_Fk = rs("User_Fk")
	End If
	rs.Close
	Set rs = Nothing
End Function
%>
<%

'-------------------------------------------------------------------------------
' Function EditData
' - Edit Data based on Key Value
' - Variables used: field variables

Function EditData()
	
	On Error Resume Next
	Dim rs, sSql, sFilter
	Dim rsold, rsnew
	sFilter = ewSqlKeyWhere
	If Not IsNumeric(x_PID) Then
		EditData = False
		Exit Function
	End If
	sFilter = Replace(sFilter, "@PID", AdjustSql(x_PID)) ' Replace key value
	sSql = ewBuildSql(ewSqlSelect, ewSqlWhere, ewSqlGroupBy, ewSqlHaving, ewSqlOrderBy, sFilter, "")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open sSql, conn, 1, 2
	If Err.Number <> 0 Then
		Session(ewSessionMessage) = Err.Description
		rs.Close
		Set rs = Nothing
		EditData = False
		Exit Function
	End If

	' clone old rs object
	Set rsold = CloneRs(rs)
	If rs.Eof Then
		EditData = False ' Update Failed
	Else

		' Field PID
		' Field Member1

		sTmp = x_Member1
		If sTmp <> "" Then
			rs("Member1") = True
		Else
			rs("Member1") = False
		End If

		' Field Member2
		sTmp = x_Member2
		If sTmp <> "" Then
			rs("Member2") = True
		Else
			rs("Member2") = False
		End If

		' Field Member3
		sTmp = x_Member3
		If sTmp <> "" Then
			rs("Member3") = True
		Else
			rs("Member3") = False
		End If

		' Field Member4
		sTmp = x_Member4
		If sTmp <> "" Then
			rs("Member4") = True
		Else
			rs("Member4") = False
		End If

		' Field Loan1
		sTmp = x_Loan1
		If sTmp <> "" Then
			rs("Loan1") = True
		Else
			rs("Loan1") = False
		End If

		' Field Loan2
		sTmp = x_Loan2
		If sTmp <> "" Then
			rs("Loan2") = True
		Else
			rs("Loan2") = False
		End If

		' Field Loan3
		sTmp = x_Loan3
		If sTmp <> "" Then
			rs("Loan3") = True
		Else
			rs("Loan3") = False
		End If

		' Field Loan4
		sTmp = x_Loan4
		If sTmp <> "" Then
			rs("Loan4") = True
		Else
			rs("Loan4") = False
		End If

		' Field Loan5
		sTmp = x_Loan5
		If sTmp <> "" Then
			rs("Loan5") = True
		Else
			rs("Loan5") = False
		End If

		' Field Loan6
		sTmp = x_Loan6
		If sTmp <> "" Then
			rs("Loan6") = True
		Else
			rs("Loan6") = False
		End If

		' Field Loan7
		sTmp = x_Loan7
		If sTmp <> "" Then
			rs("Loan7") = True
		Else
			rs("Loan7") = False
		End If

		' Field Loan8
		sTmp = x_Loan8
		If sTmp <> "" Then
			rs("Loan8") = True
		Else
			rs("Loan8") = False
		End If

		' Field Loan9
		sTmp = x_Loan9
		If sTmp <> "" Then
			rs("Loan9") = True
		Else
			rs("Loan9") = False
		End If

		' Field cLoan1
		sTmp = x_cLoan1
		If sTmp <> "" Then
			rs("cLoan1") = True
		Else
			rs("cLoan1") = False
		End If

		' Field cLoan2
		sTmp = x_cLoan2
		If sTmp <> "" Then
			rs("cLoan2") = True
		Else
			rs("cLoan2") = False
		End If

		' Field cLoan3
		sTmp = x_cLoan3
		If sTmp <> "" Then
			rs("cLoan3") = True
		Else
			rs("cLoan3") = False
		End If

		' Field AutoPay1
		sTmp = x_AutoPay1
		If sTmp <> "" Then
			rs("AutoPay1") = True
		Else
			rs("AutoPay1") = False
		End If

		' Field AutoPay2
		sTmp = x_AutoPay2
		If sTmp <> "" Then
			rs("AutoPay2") = True
		Else
			rs("AutoPay2") = False
		End If

		' Field AutoPay3
		sTmp = x_AutoPay3
		If sTmp <> "" Then
			rs("AutoPay3") = True
		Else
			rs("AutoPay3") = False
		End If

		' Field AutoPay4
		sTmp = x_AutoPay4
		If sTmp <> "" Then
			rs("AutoPay4") = True
		Else
			rs("AutoPay4") = False
		End If

		' Field AutoPay5
		sTmp = x_AutoPay5
		If sTmp <> "" Then
			rs("AutoPay5") = True
		Else
			rs("AutoPay5") = False
		End If

		' Field AutoPay6
		sTmp = x_AutoPay6
		If sTmp <> "" Then
			rs("AutoPay6") = True
		Else
			rs("AutoPay6") = False
		End If

		' Field AutoPay7
		sTmp = x_AutoPay7
		If sTmp <> "" Then
			rs("AutoPay7") = True
		Else
			rs("AutoPay7") = False
		End If

		' Field AutoPay8
		sTmp = x_AutoPay8
		If sTmp <> "" Then
			rs("AutoPay8") = True
		Else
			rs("AutoPay8") = False
		End If

		' Field AutoPay9
		sTmp = x_AutoPay9
		If sTmp <> "" Then
			rs("AutoPay9") = True
		Else
			rs("AutoPay9") = False
		End If

		' Field AutoPay11
		sTmp = x_AutoPay11
		If sTmp <> "" Then
			rs("AutoPay11") = True
		Else
			rs("AutoPay11") = False
		End If

		' Field AutoPay12
		sTmp = x_AutoPay12
		If sTmp <> "" Then
			rs("AutoPay12") = True
		Else
			rs("AutoPay12") = False
		End If

		' Field Saving1
		sTmp = x_Saving1
		If sTmp <> "" Then
			rs("Saving1") = True
		Else
			rs("Saving1") = False
		End If

		' Field Saving2
		sTmp = x_Saving2
		If sTmp <> "" Then
			rs("Saving2") = True
		Else
			rs("Saving2") = False
		End If

		' Field Saving3
		sTmp = x_Saving3
		If sTmp <> "" Then
			rs("Saving3") = True
		Else
			rs("Saving3") = False
		End If

		' Field Saving4
		sTmp = x_Saving4
		If sTmp <> "" Then
			rs("Saving4") = True
		Else
			rs("Saving4") = False
		End If

		' Field Saving5
		sTmp = x_Saving5
		If sTmp <> "" Then
			rs("Saving5") = True
		Else
			rs("Saving5") = False
		End If

		' Field Saving6
		sTmp = x_Saving6
		If sTmp <> "" Then
			rs("Saving6") = True
		Else
			rs("Saving6") = False
		End If

		' Field Saving7
		sTmp = x_Saving7
		If sTmp <> "" Then
			rs("Saving7") = True
		Else
			rs("Saving7") = False
		End If

		' Field Saving8
		sTmp = x_Saving8
		If sTmp <> "" Then
			rs("Saving8") = True
		Else
			rs("Saving8") = False
		End If

		' Field Saving9
		sTmp = x_Saving9
		If sTmp <> "" Then
			rs("Saving9") = True
		Else
			rs("Saving9") = False
		End If

		' Field MemAcct1
		sTmp = x_MemAcct1
		If sTmp <> "" Then
			rs("MemAcct1") = True
		Else
			rs("MemAcct1") = False
		End If

		' Field Reporting1
		sTmp = x_Reporting1
		If sTmp <> "" Then
			rs("Reporting1") = True
		Else
			rs("Reporting1") = False
		End If

		' Field Reporting2
		sTmp = x_Reporting2
		If sTmp <> "" Then
			rs("Reporting2") = True
		Else
			rs("Reporting2") = False
		End If

		' Field Reporting3
		sTmp = x_Reporting3
		If sTmp <> "" Then
			rs("Reporting3") = True
		Else
			rs("Reporting3") = False
		End If

		' Field Reporting4
		sTmp = x_Reporting4
		If sTmp <> "" Then
			rs("Reporting4") = True
		Else
			rs("Reporting4") = False
		End If

		' Field Reporting5
		sTmp = x_Reporting5
		If sTmp <> "" Then
			rs("Reporting5") = True
		Else
			rs("Reporting5") = False
		End If

		' Field Reporting6
		sTmp = x_Reporting6
		If sTmp <> "" Then
			rs("Reporting6") = True
		Else
			rs("Reporting6") = False
		End If

		' Field Reporting7
		sTmp = x_Reporting7
		If sTmp <> "" Then
			rs("Reporting7") = True
		Else
			rs("Reporting7") = False
		End If

		' Field Reporting8
		sTmp = x_Reporting8
		If sTmp <> "" Then
			rs("Reporting8") = True
		Else
			rs("Reporting8") = False
		End If

		' Field Reporting9
		sTmp = x_Reporting9
		If sTmp <> "" Then
			rs("Reporting9") = True
		Else
			rs("Reporting9") = False
		End If

		' Field Reporting10
		sTmp = x_Reporting10
		If sTmp <> "" Then
			rs("Reporting10") = True
		Else
			rs("Reporting10") = False
		End If

		' Field Reporting11
		sTmp = x_Reporting11
		If sTmp <> "" Then
			rs("Reporting11") = True
		Else
			rs("Reporting11") = False
		End If

		' Field Other4
		sTmp = x_Other4
		If sTmp <> "" Then
			rs("Other4") = True
		Else
			rs("Other4") = False
		End If

		' Field Other3
		sTmp = x_Other3
		If sTmp <> "" Then
			rs("Other3") = True
		Else
			rs("Other3") = False
		End If

		' Field Other2
		sTmp = x_Other2
		If sTmp <> "" Then
			rs("Other2") = True
		Else
			rs("Other2") = False
		End If

		' Field Other1
		sTmp = x_Other1
		If sTmp <> "" Then
			rs("Other1") = True
		Else
			rs("Other1") = False
		End If

		' Call updating event
		If Recordset_Updating(rsold, rs) Then

			' clone new rs object
			Set rsnew = CloneRs(rs)
			rs.Update
			If Err.Number <> 0 Then
				Session(ewSessionMessage) = Err.Description
				EditData = False
			Else
				EditData = True
			End If
		Else
			rs.CancelUpdate
			EditData = False
		End If
	End If

	' Call updated event
	If EditData Then
		Call Recordset_Updated(rsold, rsnew)
	End If
	rs.Close
	Set rs = Nothing
	rsold.Close
	Set rsold = Nothing
	rsnew.Close
	Set rsnew = Nothing
	
	
End Function

'-------------------------------------------------------------------------------
' Recordset updating event

Function Recordset_Updating(rsold, rsnew)
	On Error Resume Next

	' Please enter your customized codes here
	Recordset_Updating = True
End Function

'-------------------------------------------------------------------------------
' Recordset updated event

Sub Recordset_Updated(rsold, rsnew)
	On Error Resume Next
	Dim table
	table = "userRights"
End Sub
%>

