<link rel="stylesheet" href="template1.css" type="text/css">
<link rel="stylesheet" href="main.css" type="text/css">

<SCRIPT language="javascript">
<!--
var activeMenu = 0;
document.onmouseover = killMenu;
document.onmouseout = killMenu;


function activateMenu(menuLayerRef) {
  if (activeMenu != menuLayerRef) {
    if (activeMenu) hideMenu("menu" + activeMenu);
      activeMenu = menuLayerRef;
      menuID = "menu" + menuLayerRef;
      menutopID = "menutop" + menuLayerRef;
      document.all[menuID].style.pixelTop =  110; //53; //107;
      document.all[menuID].style.pixelLeft = document.all[menutopID].offsetLeft-4;
      showMenu(menuID)
  }
  window.event.cancelBubble = true;
}

function showMenu(layerID) {
  eval('document.all["' + layerID + '"].style.visibility = "visible"');
  ToolbarMenu = document.all(layerID)
  hideElement("SELECT");
  hideElement("OBJECT");
}

function hideMenu(layerID) {
  eval('document.all["' + layerID + '"].style.visibility = "hidden"');
  showElement("SELECT");
  showElement("OBJECT");
}

function killMenu(e) {
  if (activeMenu) {
    menuID = "menu" + activeMenu;
    hideMenu(menuID);
    activeMenu = 0;
  }
}

function hideElement(elmID)
{
  for (i = 0; i < document.all.tags(elmID).length; i++)
  {
    obj = document.all.tags(elmID)[i];
    if (! obj || ! obj.offsetParent)
      continue;

      // Find the element's offsetTop and offsetLeft relative to the BODY tag.
      objLeft   = obj.offsetLeft;
      objTop    = obj.offsetTop;
      objParent = obj.offsetParent;
      while (objParent.tagName.toUpperCase() != "BODY")
      {
        objLeft  += objParent.offsetLeft;
        objTop   += objParent.offsetTop;
        objParent = objParent.offsetParent;
      }
      // Adjust the element's offsetTop relative to the dropdown menu
      objTop = objTop - 69;

      if (ToolbarMenu.offsetLeft > (objLeft + obj.offsetWidth) || objLeft > (ToolbarMenu.offsetLeft + ToolbarMenu.offsetWidth));
      else if (objTop > ToolbarMenu.offsetHeight);
      else
        obj.style.visibility = "hidden";
  	}
}

function showElement(elmID)
{
  for (i = 0; i < document.all.tags(elmID).length; i++)
  {
    obj = document.all.tags(elmID)[i];
    if (! obj || ! obj.offsetParent)
      continue;
    obj.style.visibility = "";
  }
}

<%
thisURL=request.servervariables("script_name")
thisURL=mid(thisURL,InstrRev(thisURL,"/")+1,InstrRev(thisURL,".")-InstrRev(thisURL,"/")-1)
%>
function helppopup(){
  window.open ('../help/<%=thisURL%>.htm','pop','width=700,height=550,statusbar=no,toolbar=no,resizable,scrollbars,dependent')
}
//-->
</script>
<%


sub draw_sub_menu()

	jo_sub_menu_array = split(jo_sub_menu_text,",")
	jo_sub_menu_link_array = split(jo_sub_menu_link,",")

	Response.Write	"<DIV id='menu" & id & "' class='menu' onMouseover='activateMenu(" & id & ");'>"
	Response.Write	"<table cellpadding='0' cellspacing='0'>"

	For i = 0 To UBound(jo_sub_menu_array)


		If jo_sub_menu_array(i) = "hr" Then	
	
			Response.Write	"<tr><td colspan='2'><HR STYLE='color: #CCCCCC' SIZE=1 width='100%'></td></tr>"
		Else
			Response.Write	"<tr><td colspan='2'><A href='" & jo_sub_menu_link_array(i) & "'>" & jo_sub_menu_array(i) & "</A></td></tr>"

		End If	

	Next

	Response.Write	"</table>"
	Response.Write	"</DIV>"

end sub




select case session("userLevel")
case 5 ''Auditor
%>
<DIV id="menu1" class="menu" onMouseover="activateMenu(1);">
<A href="gl.asp">G/L Maintenance</A>
<BR>
<A href="glList.asp">G/L List</A>
<BR>
<A href="glTx.asp">Transaction Maintenance</A>
<BR>
<A href="glTxList.asp">Transaction List</A>
<HR STYLE='color: white' SIZE=1 width='180'>
<A href="trialBalance.asp">Trial Balance</A>
<BR>
<A href="pl.asp">Profit & Lost Statement</A>
<BR>
<A href="balanceSheet.asp">Balance Sheet</A>
</DIV>

<DIV id="menu2" class="menu" onMouseover="activateMenu(2);">
<A href="ac.asp">Account Maintenance</A>
<BR>
<A href="acTx.asp">Daily Transaction Input</A>
<BR>
<A href="acBal.asp">Print Balance Statement</A>
<BR>
<A href="acTxList.asp">Transaction List</A>
<HR STYLE='color: white' SIZE=1 width='280'>
<A href="atList.asp">Account Check list for Auto-pay</A>
<BR>
<A href="sdList.asp">Account Check list for Salary deduction</A>
<HR STYLE='color: white' SIZE=1 width='280'>
<A href="delinquentReport.asp">Delinquent Loan report</A>
<BR>
<A href="dormantList.asp">Dormant Account List</A>
<BR>
<A href="sectionList.asp">Section report</A>
<HR STYLE='color: white' SIZE=1 width='280'>
<A href="yearEndReport.asp">Year End Testing Report</A>
<HR STYLE='color: white' SIZE=1 width='280'>
<A href="loan.asp">Loan application Assessment</A>
<BR>
<A href="loanReport.asp">Loan Application Assessment Reports</A>
<BR>
<A href="meetingNotes.asp">Meeting Notes</A>
<BR>
<A href="monthlyReport.asp">Monthly Report</A>
<BR>
<A href="handleParty.asp">Commitee</A>
<HR STYLE='color: white' SIZE=1 width='280'>
<A href="cheque.asp">Cheque Reconciliation</A>
</DIV>




<DIV id="menu3" class="menu" onMouseover="activateMenu(3);">
<A href="member.asp">Member Maintenance</A>
<BR>
<A href="memList.asp">Member list</A>
<BR>
<A href="birthdayList.asp">Birthday list</A>
</DIV>

<DIV id="menu4" class="menu" onMouseover="activateMenu(4);">
<A href="acBal.asp">Print Balance Statement</A>
<BR>
<A href="acTxList.asp">A/C Transaction List</A>
<BR>
<A href="atList.asp">Account Check list for Auto-pay</A>
<BR>
<A href="sdList.asp">Account Check list for Salary deduction</A>
<BR>
<A href="delinquentReport.asp">Delinquent Loan report</A>
<BR>
<A href="dormantList.asp">Dormant Account List</A>
<BR>
<A href="sectionList.asp">Section report</A>
<BR>
<A href="yearEndReport.asp">Year End Testing Report</A>
<BR>
<A href="loanReport.asp">Loan Application Assessment Reports</A>
<BR>
<A href="memList.asp">Member list</A>
<BR>
<A href="birthdayList.asp">Birthday list</A>
</DIV>

<DIV id="menu5" class="menu" onMouseover="activateMenu(5);">
<A href="changePassword.asp">Change Password</A>
<BR>
<A href="userLog.asp">User Log Report</A>
</DIV>





<%
case 4 ''Administrator

	'jo_sub_menu_text = "Member Maintenance,Member list,Birthday list,hr,Delinquent Loan report,Dormant Account List,Section report"
	'jo_sub_menu_link = "member.asp,memList.asp,birthdayList.asp,hr,delinquentReport.asp,dormantList.asp,sectionList.asp"

	jo_sub_menu_text = "加入新社員,社員資料修正,社員列表,hr,呆賬報告,冷戶報告,社員分組報告"
	jo_sub_menu_link =  "memberDetail.asp, member.asp,memList.asp,hr,delinquentReport.asp,dormantList.asp,sectionList.asp"

	id = 1

	draw_sub_menu()



	'jo_sub_menu_text = "Loan application Assessment,Loan Application Assessment Reports" 
	jo_sub_menu_text = "貸款申請,貸款申請列表"
	jo_sub_menu_link = "loan.asp,loanReport.asp"

	id = 2


	draw_sub_menu()




	'jo_sub_menu_text = "Account Enquiry,Account Maintenance,Daily Transaction Input,Print Balance Statement,Transaction List,hr,League-Due Process,Auto-processing for Auto-pay,Auto-processing for Salary deduction,Account Check list for Auto-pay,Account Check list for Salary deduction,hr,Year End Testing Report,Year End Process,hr,Meeting Notes,Monthly Report,Committee,hr,Cheque Reconciliation"
	'jo_sub_menu_link = "ac1.asp,ac.asp,acTx.asp,acBal.asp,acTxList.asp,hr,leagueDue.asp,autoplay.asp,salaryDeduction.asp,atList.asp,sdList.asp,hr,yearEndReport.asp,yearEnd.asp,hr,meetingNotes.asp,monthlyReport.asp,handleParty.asp,hr,cheque.asp"

	jo_sub_menu_text = "社員資料查詢,個人賬修正,個人賬入數,社員個人結算書,個人賬細項列印,hr,目動扣除協會費,銀行轉賬自動過數,庫房轉賬自動過數,銀行轉賬試算,庫房轉賬試算,hr,年結股息試算,年結股息計算" ',hr,會議紀錄,董事會報告書,委員資料修正,hr,支票對數"
	jo_sub_menu_link = "ac1.asp,ac.asp,acTx.asp,acBal.asp,acTxList.asp,hr,leagueDue.asp,autopay.asp,salaryDeduction.asp,atList.asp,sdList.asp,hr,yearEndReport.asp,yearEnd.asp,hr,meetingNotes.asp,monthlyReport.asp,handleParty.asp,hr,cheque.asp"

	id = 3

	draw_sub_menu()


	'jo_sub_menu_text = "G/L Maintenance,G/L List,Transaction Maintenance,Transaction List,Post Surplus or Deficit,Post Delinquent,hr,Period End,hr,Trial Balance,Profit & Lost Statement,Balance Sheet"
	'jo_sub_menu_link = "gl.asp,glList.asp,glTx.asp,glTxList.asp,postSurplus.asp,postDelinquent.asp,hr,periodEnd.asp,hr,trialBalance.asp,pl.asp,balanceSheet.asp"

	jo_sub_menu_text = "總賬修正,總賬表,總賬入數,總賬細項列印,歸納盈利,歸納呆賬,hr,每月完結,hr,總賬試算表,損益表,財務報告表,社員生日名單,總賬修正,總賬表,總賬入數,總賬細項列印,總賬試算表,損益表,財務報告表"
	jo_sub_menu_link = "gl.asp,glList.asp,glTx.asp,glTxList.asp,postSurplus.asp,postDelinquent.asp,hr,periodEnd.asp,hr,trialBalance.asp,pl.asp,balanceSheet.asp,birthdayList.asp.gl.asp,glList.asp,glTx.asp,glTxList.asp,hr,trialBalance.asp,pl.asp,balanceSheet.asp"

	id = 4

	draw_sub_menu()


	'jo_sub_menu_text = "Print Balance Statement,A/C Transaction List,Account Check list for Auto-pay,Account Check list for Salary deduction,Delinquent Loan report,Dormant Account List,Section report,Year End Testing Report,Loan Application Assessment Reports,Member list,Birthday list"
	'jo_sub_menu_link = "acBal.asp,acTxList.asp,atList.asp,sdList.asp,delinquentReport.asp,dormantList.asp,sectionList.asp,yearEndReport.asp,loanReport.asp,memList.asp,birthdayList.asp"

	jo_sub_menu_text = "社員個人結算書,個人賬細項列印,銀行轉賬試算,庫房轉賬試算,呆賬報告,冷戶報告,社員分組報告,年結股息試算,貸款申請列表,社員列表,社員生日名單"
	jo_sub_menu_link = "acBal.asp,acTxList.asp,atList.asp,sdList.asp,delinquentReport.asp,dormantList.asp,sectionList.asp,yearEndReport.asp,loanReport.asp,memList.asp,birthdayList.asp"

	id = 5

	draw_sub_menu()



	'jo_sub_menu_text = "Database Exporting,Database Importing,User Administration,User Log Report"
	'jo_sub_menu_link = "dataExport.asp,dataImport.asp,user.asp,userLog.asp"

	jo_sub_menu_text = "資料庫輸出,資料庫輸入,用戶管理,用戶使用紀錄"
	jo_sub_menu_link = "dataExport.asp,dataImport.asp,user.asp,userLog.asp"


	id = 6

	draw_sub_menu()

%>









<%
case 3 ''Supervisor
%>
<DIV id="menu1" class="menu" onMouseover="activateMenu(1);">
<A href="gl.asp">G/L Maintenance</A>
<BR>
<A href="glList.asp">G/L List</A>
<BR>
<A href="glTx.asp">Transaction Maintenance</A>
<BR>
<A href="glTxList.asp">Transaction List</A>
<HR STYLE='color: white' SIZE=1 width='180'>
<A href="periodEnd.asp">Period End</A>
<HR STYLE='color: white' SIZE=1 width='180'>
<A href="trialBalance.asp">Trial Balance</A>
<BR>
<A href="pl.asp">Profit & Lost Statement</A>
<BR>
<A href="balanceSheet.asp">Balance Sheet</A>
</DIV>

<DIV id="menu2" class="menu" onMouseover="activateMenu(2);">
<A href="ac.asp">Account Maintenance</A>
<BR>
<A href="acTx.asp">Daily Transaction Input</A>
<BR>
<A href="acBal.asp">Print Balance Statement</A>
<BR>
<A href="acTxList.asp">Transaction List</A>
<HR STYLE='color: white' SIZE=1 width='280'>
<A href="leagueDue.asp">League-Due Process</A>
<BR>
<A href="autopay.asp">Auto-processing for Auto-pay</A>
<BR>
<A href="salaryDeduction.asp">Auto-processing for Salary deduction</A>
<BR>
<A href="atList.asp">Account Check list for Auto-pay</A>
<BR>
<A href="sdList.asp">Account Check list for Salary deduction</A>
<HR STYLE='color: white' SIZE=1 width='280'>
<A href="delinquentReport.asp">Delinquent Loan report</A>
<BR>
<A href="dormantList.asp">Dormant Account List</A>
<BR>
<A href="sectionList.asp">Section report</A>
<HR STYLE='color: white' SIZE=1 width='280'>
<A href="yearEndReport.asp">Year End Testing Report</A>
<BR>
<A href="yearEnd.asp">Year End Process</A>
<HR STYLE='color: white' SIZE=1 width='280'>
<A href="loan.asp">Loan application Assessment</A>
<BR>
<A href="loanReport.asp">Loan Application Assessment Reports</A>
<BR>
<A href="meetingNotes.asp">Meeting Notes</A>
<BR>
<A href="monthlyReport.asp">Monthly Report</A>
<HR STYLE='color: white' SIZE=1 width='280'>
<A href="cheque.asp">Cheque Reconciliation</A>
</DIV>

<DIV id="menu3" class="menu" onMouseover="activateMenu(3);">
<A href="member.asp">Member Maintenance</A>
<BR>
<A href="memList.asp">Member list</A>
<BR>
<A href="birthdayList.asp">Birthday list</A>
</DIV>

<DIV id="menu4" class="menu" onMouseover="activateMenu(4);">
<A href="acBal.asp">Print Balance Statement</A>
<BR>
<A href="acTxList.asp">A/C Transaction List</A>
<BR>
<A href="atList.asp">Account Check list for Auto-pay</A>
<BR>
<A href="sdList.asp">Account Check list for Salary deduction</A>
<BR>
<A href="delinquentReport.asp">Delinquent Loan report</A>
<BR>
<A href="dormantList.asp">Dormant Account List</A>
<BR>
<A href="sectionList.asp">Section report</A>
<BR>
<A href="yearEndReport.asp">Year End Testing Report</A>
<BR>
<A href="loanReport.asp">Loan Application Assessment Reports</A>
<BR>
<A href="memList.asp">Member list</A>
<BR>
<A href="birthdayList.asp">Birthday list</A>
</DIV>
<%
case 2 ''Operator
%>
<DIV id="menu1" class="menu" onMouseover="activateMenu(1);">
<A href="gl.asp">G/L Maintenance</A>
<BR>
<A href="glTx.asp">Transaction Maintenance</A>
<BR>
<A href="periodEnd.asp">Period End</A>
</DIV>

<DIV id="menu2" class="menu" onMouseover="activateMenu(2);">
<A href="ac.asp">Account Maintenance</A>
<BR>
<A href="acTx.asp">Daily Transaction Input</A>
<BR>
<A href="cheque.asp">Cheque Reconciliation</A>
</DIV>
<%
case 1 ''Member
%>
<DIV id="menu1" class="menu" onMouseover="activateMenu(1);">
<A href="loan.asp">Loan application Assessment</A>
<BR>
<A href="loanReport.asp">Loan Application Assessment Reports</A>
<BR>
<A href="meetingNotes.asp">Meeting Notes</A>
<BR>
<A href="monthlyReport.asp">Monthly Report</A>
</DIV>

<DIV id="menu2" class="menu" onMouseover="activateMenu(2);">
<A href="ac.asp">Account Maintenance</A>
<BR>
<A href="member.asp">Member Maintenance</A>
</DIV>
<%
end select
%>





<%


select case session("userLevel")

	case 4 ''Administrator

	'jo_menu_text  = "Membership,Loan,Account System,General Ledger,Reports shortcut,Others"
	jo_menu_text = "社員資料,貸款,個人戶口,總賬,報表,其他,登出 "
	jo_menu_array = split(jo_menu_text, ",")


end select

%>




<table>
<tr>
<td><img src="images/logo.gif"></td>
<td>水務署員工儲蓄互助社系統<br>
Water Supplies Department Staff Credit Union Membership, Accounting, Savings and Loans Software </td>
</tr>
</table>


<!-- Main Menu -->
<div id="topmenutabs">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
<td><img border="0" src="images/blank.gif" /></td>


<!--th class="mtab-ls"><img border="0" src="images/blank.gif" /></th-->
<!--th class="mtab-rs"><img border="0" src="images/blank.gif" /></th-->

<%

For i = 0 To (UBound(jo_menu_array)) 
	Response.Write 	"<td class='mtab-l'><img border='0' src='images/blank.gif' /></td>"
  	Response.Write	"<td class='mtab-r'><img border='0' src='images/blank.gif' /></td>"
Next

%>


<td style="width:100%; background: #FFF;"><img border="0" src="images/blank.gif" /></td>
</tr>
<tr>
<td class="menuBackground">&nbsp;</td>
<%

For i = 0 To (UBound(jo_menu_array))

	if (i = UBound(jo_menu_array)) then
		Response.Write	"<td id='menutop" & (i+1) & "' class='menutabs-td' colspan='2' nowrap>&nbsp;&nbsp;<a href='../logout.asp'>" & jo_menu_array(i) & "</a>&nbsp;&nbsp;</td>"
	else 
		Response.Write	"<td id='menutop" & (i+1) & "' class='menutabs-td' colspan='2' nowrap>&nbsp;&nbsp;<a href='#' onmouseover='activateMenu(" & (i+1) & ");' id='amenutop" & (i+1) & "'>" & jo_menu_array(i) & "</a>&nbsp;&nbsp;</td>"
	end if
Next

%>

<td width="100%" class="menuBackground"><img border="0" src="images/blank.gif" /></td>
</tr>
</table>
</div>