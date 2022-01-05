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
      document.all[menuID].style.pixelTop = 48;
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
      objTop = objTop - 48;

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
//-->
</script>
<%
select case session("userLevel")
case 4
%>
<DIV id="menu1" class="menu" onMouseover="activateMenu(1);">
<A href="gl.asp">G/L Maintenance</A>
<BR>
<A href="glList.asp">G/L List</A>
<BR>
<A href="glTx.asp">G/L Transaction Maintenance</A>
<BR>
<A href="glTxList.asp">G/L Transaction List</A>
<BR>
<A href="#">Day End</A>
<BR>
<A href="#">Period End</A>
<BR>
<A href="#">Trial Balance</A>
<BR>
<A href="#">Profit & Loss Statment</A>
<BR>
<A href="#">Balance Sheet</A>
</DIV>

<DIV id="menu2" class="menu" onMouseover="activateMenu(2);">
<A href="ac.asp">Account Details</A>
<BR>
<A href="acTx.asp">Transaction List</A>
<BR>
<A href="#">Auto-processing</A>
<BR>
<A href="#">League-Due Process</A>
<BR>
<A href="#">Year end Bonus Update</A>
<BR>
<A href="#">Account Check list report</A>
<BR>
<A href="#">Year end operation</A>
<BR>
<A href="#">Report</A>
</DIV>

<DIV id="menu3" class="menu" onMouseover="activateMenu(3);">
<A href="member.asp">Maintenance</A>
<BR>
<A href="#">Member List</A>
<BR>
<A href="#">Lucky Draw</A>
</DIV>

<DIV id="menu4" class="menu" onMouseover="activateMenu(4);">
<A href="#">Statistical Reports</A>
<BR>
<A href="#">Print Balance Statement</A>
<BR>
<A href="#">Transaction List</A>
<BR>
<A href="#">Account Check list report</A>
<BR>
<A href="#">Delinquent Loan report</A>
<BR>
<A href="#">Dormant Account List</A>
<BR>
<A href="#">Section report</A>
<BR>
<A href="#">Year end Testing report</A>
<BR>
<A href="#">Member Details</A>
<BR>
<A href="#">Birthday list</A>
<BR>
<A href="#">Name list</A>
<BR>
<A href="#">Member account list</A>
<BR>
<A href="#">Loan application Accessment</A>
</DIV>

<DIV id="menu5" class="menu" onMouseover="activateMenu(5);">
<A href="#">Database Exporting</A>
<BR>
<A href="#">User Administration</A>
</DIV>
<%
case else%>
<DIV id="menu1" class="menu" onMouseover="activateMenu(1);">
<A href="#">Account Details</A>
<BR>
<A href="#">Transaction List</A>
<BR>
<A href="#">Auto-processing</A>
<BR>
<A href="#">Auto-processing</A>
<BR>
<A href="#">Auto-processing</A>
<BR>
<A href="#">Auto-processing</A>
<BR>
<A href="#">Auto-processing</A>
</DIV>
<%
end select%>


<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr bgcolor="#ffffff">
		<td><a href="main.asp"><img src="../images/logo.gif" border="0"></a></td>
		<td align="right"><b>Login user : </b><%=ucase(session("username"))%> &nbsp;&nbsp;</td>
	</tr>
	<tr>
		<td bgcolor="#336699" class="menutop">
<%
select case session("userLevel")
case 4
%>
			<a href="#" onmouseover="activateMenu(1);" class="menutop" id="menutop1">General Ledger</a>
			|
			<a href="#" onmouseover="activateMenu(2)" class="menutop" id="menutop2">Account System</a>
			|
			<a href="#" onmouseover="activateMenu(3)" class="menutop" id="menutop3">Membership</a>
			|
			<a href="#" onmouseover="activateMenu(4)" class="menutop" id="menutop4">Reports shortcut</a>
			|
			<a href="#" onmouseover="activateMenu(5)" class="menutop" id="menutop5">Others</a>
			|
<%
case else%>
			<a href="#" onmouseover="activateMenu(1);" class="menutop" id="menutop1">General Ledger</a>
			|
<%
end select%>
			<a href="..\logout.asp" class="menutop">Logout</a>
		</td>
		<td bgcolor="#336699" class="menutop" align="right">
<%
thisFileName=Request.ServerVariables("script_name")
thisFileName=mid(thisFileName,InstrRev(thisFileName,"/")+1)
if request.QueryString <> "" then thisFileName=thisFileName&"?"&request.QueryString
%>
			<a href="../cn/<%=thisFileName%>" class="menutop">¤¤¤å</a>
			|
			<a href="../en/<%=thisFileName%>" class="menutop">English</a> &nbsp;&nbsp;
		</td>
	</tr>
</table>