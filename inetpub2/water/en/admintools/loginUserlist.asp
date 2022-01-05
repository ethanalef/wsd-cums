<%@ CodePage=950 %>

<!--#include file="db.asp"-->
<%
' Table Level Constants

Const ewTblVar = "loginUser"
Const ewTblRecPerPage = "RecPerPage"
Const ewSessionTblRecPerPage = "loginUser_RecPerPage"
Const ewTblStartRec = "start"
Const ewSessionTblStartRec = "loginUser_start"
Const ewTblShowMaster = "showmaster"
Const ewSessionTblMasterKey = "loginUser_MasterKey"
Const ewSessionTblMasterWhere = "loginUser_MasterWhere"
Const ewSessionTblDetailWhere = "loginUser_DetailWhere"
Const ewSessionTblAdvSrch = "loginUser_AdvSrch"
Const ewTblBasicSrch = "psearch"
Const ewSessionTblBasicSrch = "loginUser_psearch"
Const ewTblBasicSrchType = "psearchtype"
Const ewSessionTblBasicSrchType = "loginUser_psearchtype"
Const ewSessionTblSearchWhere = "loginUser_SearchWhere"
Const ewSessionTblSort = "loginUser_Sort"
Const ewSessionTblOrderBy = "loginUser_OrderBy"
Const ewSessionTblKey = "loginUser_Key"

' Table Level SQL
Const ewSqlSelect = "SELECT * FROM [loginUser]"
Const ewSqlWhere = ""
Const ewSqlGroupBy = ""
Const ewSqlHaving = ""
Const ewSqlOrderBy = ""
Const ewSqlOrderBySessions = ""
Const ewSqlKeyWhere = "[uid] = @uid"
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
x_uid = Null: ox_uid = Null: z_uid = Null
x_username = Null: ox_username = Null: z_username = Null
x_password = Null: ox_password = Null: z_password = Null
x_userLevel = Null: ox_userLevel = Null: z_userLevel = Null
x_lastLoginTime = Null: ox_lastLoginTime = Null: z_lastLoginTime = Null
%>
<%
nStartRec = 0
nStopRec = 0
nTotalRecs = 0
nRecCount = 0
nRecActual = 0
sDbWhereMaster = ""
sDbWhereDetail = ""
sSrchAdvanced = ""
psearch = ""
psearchtype = ""
sSrchBasic = ""
sSrchWhere = ""
sDbWhere = ""
sOrderBy = ""
sSqlMaster = ""
nDisplayRecs = 1000
nRecRange = 10

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

' Handle Reset Command
ResetCmd()

' Get Search Criteria for Basic Search
SetUpBasicSearch()

' Build Search Criteria
If sSrchAdvanced <> "" Then
	If sSrchWhere <> "" Then sSrchWhere = sSrchWhere & " AND "
	sSrchWhere = sSrchWhere & "(" & sSrchAdvanced & ")"
End If
If sSrchBasic <> "" Then
	If sSrchWhere <> "" Then sSrchWhere = sSrchWhere & " AND "
	sSrchWhere = sSrchWhere & "(" & sSrchBasic & ")"
End If

' Save Search Criteria
If sSrchWhere <> "" Then
	Session(ewSessionTblSearchWhere) = sSrchWhere
	nStartRec = 1 ' reset start record counter
	Session(ewSessionTblStartRec) = nStartRec
Else
	sSrchWhere = Session(ewSessionTblSearchWhere)
	Call RestoreSearch()
End If

' Build Filter condition
sDbWhere = ""
If sDbWhereDetail <> "" Then
	If sDbWhere <> "" Then sDbWhere = sDbWhere & " AND "
	sDbWhere = sDbWhere & "(" & sDbWhereDetail & ")"
End If
If sSrchWhere <> "" Then
	If sDbWhere <> "" Then sDbWhere = sDbWhere & " AND "
	sDbWhere = sDbWhere & "(" & sSrchWhere & ")"
End If

' Set Up Sorting Order
sOrderBy = ""
SetUpSortOrder()

' Set up SQL
sSql = ewBuildSql(ewSqlSelect, ewSqlWhere, ewSqlGroupBy, ewSqlHaving, ewSqlOrderBy, sDbWhere, sOrderBy)

'Response.Write sSql ' Uncomment to show SQL for debugging
%>



<script type="text/javascript">
<!--
var EW_dateSep; // default date separator
if (EW_dateSep == '') EW_dateSep = '/';

var EW_fieldSep = ', '; // default separator between display fields
var EW_ver55 = true;

function EW_DHTMLEditor(name) {
	this.name = name;
	this.create = function() { this.active = true; }
	this.editor = null;
	this.active = false;
}

function EW_createEditor(name) {
	if (typeof EW_DHTMLEditors == 'undefined')
		return;
	if (name && name.substring(0,2) == 'r_')
		name = name.replace(/r_/, 'x_');
	for (var i = 0; i < EW_DHTMLEditors.length; i++) {
    var ed = EW_DHTMLEditors[i];
		var cr = !ed.active;
		if (name) cr = cr && ed.name == name;
		if (cr) {
			if (typeof ed.create == 'function')
				ed.create();
			if (name)
				break;
		}
	}
}

function EW_submitForm(EW_this) {	
	if (typeof EW_UpdateTextArea == 'function')
		EW_UpdateTextArea();
	if (EW_checkMyForm(EW_this))
		EW_this.submit();
}

function EW_RemoveSpaces(value) {
	str = value.replace(/^\s*|\s*$/g, "");
	str = str.toLowerCase();
	if (str == "<p />" || str == "<p/>" || str == "<p>" ||
		str == "<br />" || str == "<br/>" || str == "<br>" ||
		str == "&nbsp;" || str == "<p>&nbsp;</p>")
		return ""
	else
		return value;
}

function ew_IsHiddenTextArea(input_object) {
	return (input_object && input_object.type && input_object.type == "textarea" &&
		input_object.style && input_object.style.display &&
		input_object.style.display == "none");
}

function ew_SetFocus(input_object) {
	if (!input_object || !input_object.type)
		return;
	var type = input_object.type;	 			
	if (type == "radio" || type == "checkbox") {
		if (input_object[0])
			input_object[0].focus();
		else
			input_object.focus();
	}	else if (!ew_IsHiddenTextArea(input_object)) { 
		input_object.focus();  
	}  
	if (type == "text" || type == "password" || type == "textarea" || type == "file") {
		if (!ew_IsHiddenTextArea(input_object))
			input_object.select();
	}
}

function EW_onError(form_object, input_object, object_type, error_message) {
	alert(error_message);
	if (typeof ew_GotoPageByElement != 'undefined') // check if multi-page
		ew_GotoPageByElement(input_object);									
	ew_SetFocus(input_object);
	return false;	
}

function EW_hasValue(obj, obj_type) {
	if (obj_type == "TEXT" || obj_type == "PASSWORD" || obj_type == "TEXTAREA" || obj_type == "FILE")	{
		if (obj.value.length == 0) 
			return false;		
		else 
			return true;
	}	else if (obj_type == "SELECT") {
		if (obj.type != "select-multiple" && obj.selectedIndex == 0)
			return false;
		else if (obj.type == "select-multiple" && obj.selectedIndex == -1)
			return false;
		else
			return true;
	}	else if (obj_type == "RADIO" || obj_type == "CHECKBOX")	{
		if (obj[0]) {
			for (i=0; i < obj.length; i++) {
				if (obj[i].checked)
					return true;
			}
		} else {
			return (obj.checked);
		}
		return false;	
	}
}

// Date (mm/dd/yyyy)
function EW_checkusdate(object_value) {
	if (object_value.length == 0)
		return true;
	
	isplit = object_value.indexOf(EW_dateSep);
	
	if (isplit == -1 || isplit == object_value.length)
		return false;
	
	sMonth = object_value.substring(0, isplit);
	
	if (sMonth.length == 0)
		return false;
	
	isplit = object_value.indexOf(EW_dateSep, isplit + 1);
	
	if (isplit == -1 || (isplit + 1 ) == object_value.length)
		return false;
	
	sDay = object_value.substring((sMonth.length + 1), isplit);
	
	if (sDay.length == 0)
		return false;
	
	isep = object_value.indexOf(' ', isplit + 1); 
	if (isep == -1) {
		sYear = object_value.substring(isplit + 1);
	} else {
		sYear = object_value.substring(isplit + 1, isep);
		sTime = object_value.substring(isep + 1);
		if (!EW_checktime(sTime))
			return false; 
	}
	
	if (!EW_checkinteger(sMonth)) 
		return false;
	else if (!EW_checkrange(sMonth, 1, 12)) 
		return false;
	else if (!EW_checkinteger(sYear)) 
		return false;
	else if (!EW_checkrange(sYear, 0, 9999)) 
		return false;
	else if (!EW_checkinteger(sDay)) 
		return false;
	else if (!EW_checkday(sYear, sMonth, sDay))
		return false;
	else
		return true;
}

// Date (yyyy/mm/dd, )
function EW_checkdate(object_value) {
	if (object_value.length == 0)
		return true;
	
	isplit = object_value.indexOf(EW_dateSep);
	
	if (isplit == -1 || isplit == object_value.length)
		return false;
	
	sYear = object_value.substring(0, isplit);
	
	isplit = object_value.indexOf(EW_dateSep, isplit + 1);
	
	if (isplit == -1 || (isplit + 1 ) == object_value.length)
		return false;
	
	sMonth = object_value.substring((sYear.length + 1), isplit);
	
	if (sMonth.length == 0)
		return false;
	
	isep = object_value.indexOf(' ', isplit + 1); 
	if (isep == -1) {
		sDay = object_value.substring(isplit + 1);
	} else {
		sDay = object_value.substring(isplit + 1, isep);
		sTime = object_value.substring(isep + 1);
		if (!EW_checktime(sTime))
			return false; 
	}
	
	if (sDay.length == 0)
		return false;
	
	if (!EW_checkinteger(sMonth)) 
		return false;
	else if (!EW_checkrange(sMonth, 1, 12)) 
		return false;
	else if (!EW_checkinteger(sYear)) 
		return false;
	else if (!EW_checkrange(sYear, 0, 9999)) 
		return false;
	else if (!EW_checkinteger(sDay)) 
		return false;
	else if (!EW_checkday(sYear, sMonth, sDay))
		return false;
	else
		return true;
}

// Date (dd/mm/yyyy)
function EW_checkeurodate(object_value) {
	if (object_value.length == 0)
	  return true;
	
	isplit = object_value.indexOf(EW_dateSep);
	
	if (isplit == -1 || isplit == object_value.length)
		return false;
	
	sDay = object_value.substring(0, isplit);
	
	monthSplit = isplit + 1;
	
	isplit = object_value.indexOf(EW_dateSep, monthSplit);
	
	if (isplit == -1 ||  (isplit + 1 )  == object_value.length)
		return false;
	
	sMonth = object_value.substring((sDay.length + 1), isplit);
	
	isep = object_value.indexOf(' ', isplit + 1); 
	if (isep == -1) {
		sYear = object_value.substring(isplit + 1);
	} else {
		sYear = object_value.substring(isplit + 1, isep);
		sTime = object_value.substring(isep + 1);
		if (!EW_checktime(sTime))
			return false; 
	}
	
	if (!EW_checkinteger(sMonth)) 
		return false;
	else if (!EW_checkrange(sMonth, 1, 12)) 
		return false;
	else if (!EW_checkinteger(sYear)) 
		return false;
	else if (!EW_checkrange(sYear, 0, null)) 
		return false;
	else if (!EW_checkinteger(sDay)) 
		return false;
	else if (!EW_checkday(sYear, sMonth, sDay)) 
		return false;
	else
		return true;
}

function EW_checkday(checkYear, checkMonth, checkDay) {
	maxDay = 31;
	
	if (checkMonth == 4 || checkMonth == 6 ||	checkMonth == 9 || checkMonth == 11) {
		maxDay = 30;
	} else if (checkMonth == 2)	{
		if (checkYear % 4 > 0)
			maxDay =28;
		else if (checkYear % 100 == 0 && checkYear % 400 > 0)
			maxDay = 28;
		else
			maxDay = 29;
	}
	
	return EW_checkrange(checkDay, 1, maxDay); 
}

function EW_checkinteger(object_value) {
	if (object_value.length == 0)
		return true;
	
	var decimal_format = ".";
	var check_char;
	
	check_char = object_value.indexOf(decimal_format);
	if (check_char < 1)
		return EW_checknumber(object_value);
	else
		return false;
}

function EW_numberrange(object_value, min_value, max_value) {
	if (min_value != null) {
		if (object_value < min_value)
			return false;
	}
	
	if (max_value != null) {
		if (object_value > max_value)
			return false;
	}
	
	return true;
}

function EW_checknumber(object_value) {
	if (object_value.length == 0)
		return true;
	
	var start_format = " .+-0123456789";
	var number_format = " .0123456789";
	var check_char;
	var decimal = false;
	var trailing_blank = false;
	var digits = false;
	
	check_char = start_format.indexOf(object_value.charAt(0));
	if (check_char == 1)
		decimal = true;
	else if (check_char < 1)
		return false;
	 
	for (var i = 1; i < object_value.length; i++)	{
		check_char = number_format.indexOf(object_value.charAt(i))
		if (check_char < 0) {
			return false;
		} else if (check_char == 1)	{
			if (decimal)
				return false;
			else
				decimal = true;
		} else if (check_char == 0) {
			if (decimal || digits)	
			trailing_blank = true;
		}	else if (trailing_blank) { 
			return false;
		} else {
			digits = true;
		}
	}	
	
	return true;
}

function EW_checkrange(object_value, min_value, max_value) {
	if (object_value.length == 0)
		return true;
	
	if (!EW_checknumber(object_value))
		return false;
	else
		return (EW_numberrange((eval(object_value)), min_value, max_value));	
	
	return true;
}

function EW_checktime(object_value) {
	if (object_value.length == 0)
		return true;
	
	isplit = object_value.indexOf(':');
	
	if (isplit == -1 || isplit == object_value.length)
		return false;
	
	sHour = object_value.substring(0, isplit);
	iminute = object_value.indexOf(':', isplit + 1);
	
	if (iminute == -1 || iminute == object_value.length)
		sMin = object_value.substring((sHour.length + 1));
	else
		sMin = object_value.substring((sHour.length + 1), iminute);
	
	if (!EW_checkinteger(sHour))
		return false;
	else if (!EW_checkrange(sHour, 0, 23)) 
		return false;
	
	if (!EW_checkinteger(sMin))
		return false;
	else if (!EW_checkrange(sMin, 0, 59))
		return false;
	
	if (iminute != -1) {
		sSec = object_value.substring(iminute + 1);		
		if (!EW_checkinteger(sSec))
			return false;
		else if (!EW_checkrange(sSec, 0, 59))
			return false;	
	}
	
	return true;
}

function EW_checkphone(object_value) {
	if (object_value.length == 0)
		return true;
	
	if (object_value.length != 12)
		return false;
	
	if (!EW_checknumber(object_value.substring(0,3)))
		return false;
	else if (!EW_numberrange((eval(object_value.substring(0,3))), 100, 1000))
		return false;
	
	if (object_value.charAt(3) != "-" && object_value.charAt(3) != " ")
		return false
	
	if (!EW_checknumber(object_value.substring(4,7)))
		return false;
	else if (!EW_numberrange((eval(object_value.substring(4,7))), 100, 1000))
		return false;
	
	if (object_value.charAt(7) != "-" && object_value.charAt(7) != " ")
		return false;
	
	if (object_value.charAt(8) == "-" || object_value.charAt(8) == "+")
		return false;
	else
		return (EW_checkinteger(object_value.substring(8,12)));
}


function EW_checkzip(object_value) {
	if (object_value.length == 0)
		return true;
	
	if (object_value.length != 5 && object_value.length != 10)
		return false;
	
	if (object_value.charAt(0) == "-" || object_value.charAt(0) == "+")
		return false;
	
	if (!EW_checkinteger(object_value.substring(0,5)))
		return false;
	
	if (object_value.length == 5)
		return true;
	
	if (object_value.charAt(5) != "-" && object_value.charAt(5) != " ")
		return false;
	
	if (object_value.charAt(6) == "-" || object_value.charAt(6) == "+")
		return false;
	
	return (EW_checkinteger(object_value.substring(6,10)));
}


function EW_checkcreditcard(object_value) {
	var white_space = " -";
	var creditcard_string = "";
	var check_char;
	
	if (object_value.length == 0)
		return true;
	
	for (var i = 0; i < object_value.length; i++) {
		check_char = white_space.indexOf(object_value.charAt(i));
		if (check_char < 0)
			creditcard_string += object_value.substring(i, (i + 1));
	}	
	
	if (creditcard_string.length == 0)
		return false;	 
	
	if (creditcard_string.charAt(0) == "+")
		return false;
	
	if (!EW_checkinteger(creditcard_string))
		return false;
	
	var doubledigit = creditcard_string.length % 2 == 1 ? false : true;
	var checkdigit = 0;
	var tempdigit;
	
	for (var i = 0; i < creditcard_string.length; i++) {
		tempdigit = eval(creditcard_string.charAt(i));		
		if (doubledigit) {
			tempdigit *= 2;
			checkdigit += (tempdigit % 10);			
			if ((tempdigit / 10) >= 1.0)
				checkdigit++;			
			doubledigit = false;
		}	else {
			checkdigit += tempdigit;
			doubledigit = true;
		}
	}
		
	return (checkdigit % 10) == 0 ? true : false;
}


function EW_checkssc(object_value) {
	var white_space = " -+.";
	var ssc_string="";
	var check_char;
	
	if (object_value.length == 0)
		return true;
	
	if (object_value.length != 11)
		return false;
	
	if (object_value.charAt(3) != "-" && object_value.charAt(3) != " ")
		return false;
	
	if (object_value.charAt(6) != "-" && object_value.charAt(6) != " ")
		return false;
	
	for (var i = 0; i < object_value.length; i++) {
		check_char = white_space.indexOf(object_value.charAt(i));
		if (check_char < 0)
			ssc_string += object_value.substring(i, (i + 1));
	}	
	
	if (ssc_string.length != 9)
		return false;	 
	
	if (!EW_checkinteger(ssc_string))
		return false;
	
	return true;
}
	

function EW_checkemail(object_value) {
	if (object_value.length == 0)
		return true;
	
	if (!(object_value.indexOf("@") > -1 && object_value.indexOf(".") > -1))
		return false;    
	
	return true;
}
	
// GUID {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}	
function EW_checkGUID(object_value)	{
	if (object_value.length == 0)
		return true;
	if (object_value.length != 38)
		return false;
	if (object_value.charAt(0)!="{")
		return false;
	if (object_value.charAt(37)!="}")
		return false;	
	
	var hex_format = "0123456789abcdefABCDEF";
	var check_char;	
	
	for (var i = 1; i < 37; i++) {		
		if ((i==9) || (i==14) || (i==19) || (i==24)) {
			if (object_value.charAt(i)!="-")
				return false;
		} else {
			check_char = hex_format.indexOf(object_value.charAt(i));
			if (check_char < 0)
				return false;
		}
	}
	return true;
}
	

// Update a combobox with filter value
// object_value_array format
// object_value_array[n] = option value
// object_value_array[n+1] = option text 1
// object_value_array[n+2] = option text 2
// object_value_array[n+3] = option filter value
function EW_updatecombo(obj, object_value_array, filter_value) {	
	var value = (obj.selectedIndex > -1) ? obj.options[obj.selectedIndex].value : null;
	for (var i = obj.length-1; i > 0; i--) {
		obj.options[i] = null;
	}	
	for (var j=0; j<object_value_array.length; j=j+4) {
		if (object_value_array[j+3].toUpperCase() == filter_value.toUpperCase()) {
			EW_newopt(obj, object_value_array[j], object_value_array[j+1], object_value_array[j+2]);			
		}	
	}
	EW_selectopt(obj, value);
}

// Create combobox option 
function EW_newopt(obj, value, text1, text2) {
	var text = text1;
	if (text2 != "")
		text += EW_fieldSep + text2;
	var optionName = new Option(text, value, false, false)
	var length = obj.length;
	obj.options[length] = optionName;
}

// Select combobox option
function EW_selectopt(obj, value) {
	if (value != null) {
		for (var i = obj.length-1; i>=0; i--) {
			if (obj.options[i].value.toUpperCase() == value.toUpperCase()) {
				obj.selectedIndex = i;
				break;
			}
		}
	}
}

// Get image width/height
function EW_getimagesize(file_object, width_object, height_object) {
	if (navigator.appVersion.indexOf("MSIE") != -1)	{
		myimage = new Image();
		myimage.onload = function () {
			width_object.value = myimage.width; height_object.value = myimage.height;
		}		
		myimage.src = file_object.value;
	}
}

// Get Netscape Version
function getNNVersionNumber() {
	if (navigator.appName == "Netscape") {
		var appVer = parseFloat(navigator.appVersion);
		if (appVer < 5) {
			return appVer;
		} else {
			if (typeof navigator.vendorSub != "undefined")
				return parseFloat(navigator.vendorSub);
		}
	}
	return 0;
}

// Get Ctrl key for multiple column sort
function ewsort(e, url) {	
	var ctrlPressed = 0;	
	if (parseInt(navigator.appVersion) > 3) {
		if (navigator.appName == "Netscape") {
			var ua = navigator.userAgent;
    	var isFirefox = (ua != null && ua.indexOf("Firefox/") != -1);
			if ((!isFirefox && getNNVersionNumber() >= 6) || isFirefox)				
				ctrlPressed = e.ctrlKey;
			else
				ctrlPressed = ((e.modifiers+32).toString(2).substring(3,6).charAt(1)=="1");			
		} else {
		 ctrlPressed = event.ctrlKey;
		}
		if (ctrlPressed) {
			var newPage = "<scr" + "ipt language=\"JavaScript\">setTimeout('window.location.href=\"" + url + "&ctrl=1\"', 10);</scr" + "ipt>";
			document.write(newPage);
			document.close();			
			return false;
		}
	}
	return true;
}

// Confirm Message
function ew_confirm(msg)
{
	var agree=confirm(msg);
	if (agree)
		return true ;
	else {
		return false ;
	}
}

// Confirm Delete Message
function ew_confirmdelete(msg)
{
	var agree = confirm(msg);
	if (agree)
		return true ;
	else {
		ew_cleardelete(); // clear delete status
		return false ;
	}
}

// Set mouse over color
function ew_mouseover(row) {
	row.mover = true; // mouse over
	if (!row.selected) {
		if (usecss)
			row.className = rowmoverclass;
		else
			row.style.backgroundColor = rowmovercolor;
	}
}

// Set mouse out color
function ew_mouseout(row) {
	row.mover = false; // mouse out
	if (!row.selected) {
		ew_setcolor(row);
	}
}

// Set row color
function ew_setcolor(row) {
	if (row.selected) {
		if (usecss)
			row.className = rowselectedclass;
		else
			row.style.backgroundColor = rowselectedcolor;
	}
	else if (row.edit) {
		if (usecss)
			row.className = roweditclass;
		else
			row.style.backgroundColor = roweditcolor;
	}
	else if ((row.rowIndex-firstrowoffset)%2) {
		if (usecss)
			row.className = rowaltclass;
		else
			row.style.backgroundColor = rowaltcolor;
	}
	else {
		if (usecss)
			row.className = rowclass;
		else
			row.style.backgroundColor = rowcolor;
	}
}

// Set selected row color
function ew_click(row) {
	if (row.deleteclicked)
		row.deleteclicked = false; // reset delete button/checkbox clicked
	else {
		var bselected = row.selected;
		ew_clearselected(); // clear all other selected rows
		if (!row.deleterow) row.selected = !bselected; // toggle
		ew_setcolor(row);		
	}
}

// Clear selected rows color
function ew_clearselected() {
	var table = document.getElementById(tablename);
	for (var i = firstrowoffset; i < table.rows.length; i++) {
		var thisrow = table.rows[i];
		if (thisrow.selected && !thisrow.deleterow) {
			thisrow.selected = false;
			ew_setcolor(thisrow);
		}
	}
}

// Clear all row delete status
function ew_cleardelete() {
	var table = document.getElementById(tablename);
	for (var i = firstrowoffset; i < table.rows.length; i++) {
		var thisrow = table.rows[i];
		thisrow.deleterow = false;
	}
}

// Click all delete button
function ew_clickall(chkbox) {
	var table = document.getElementById(tablename);
	for (var i = firstrowoffset; i < table.rows.length; i++) {
		var thisrow = table.rows[i];
		thisrow.selected = chkbox.checked;
		thisrow.deleterow = chkbox.checked;
		ew_setcolor(thisrow);
	}
}

// Click single delete link
function ew_clickdelete() {
	ew_clearselected();
	var table = document.getElementById(tablename);
	for (var i = firstrowoffset; i < table.rows.length; i++) {
		var thisrow = table.rows[i];
		if (thisrow.mover) {
			thisrow.deleteclicked = true;
			thisrow.deleterow = true;
			thisrow.selected = true;
			ew_setcolor(thisrow);
			break;
		}
	}
}

// Click multi delete checkbox
function ew_clickmultidelete(chkbox) {
	ew_clearselected();
	var table = document.getElementById(tablename);
	for (var i = firstrowoffset; i < table.rows.length; i++) {
		var thisrow = table.rows[i];
		if (thisrow.mover) {
			thisrow.deleteclicked = true;
			thisrow.deleterow = chkbox.checked;
			thisrow.selected = chkbox.checked;
			ew_setcolor(thisrow);
			break;
		}
	}
}

// Create XMLHTTP
// Note: AJAX feature requires IE5.5+, FF1+, and NS6.2+
function EW_createXMLHttp() {
	if (!(document.getElementsByTagName || document.all))
		return;		
	var ret = null;
	try {
		ret = new ActiveXObject('Msxml2.XMLHTTP');
	}	catch (e) {
	    try {
	        ret = new ActiveXObject('Microsoft.XMLHTTP');
	    } catch (ee) {
	        ret = null;
	    }
	}
	if (!ret && typeof XMLHttpRequest != 'undefined')
	    ret = new XMLHttpRequest();	
	return ret;
}

// Update a combobox with filter value by AJAX
// object_value_array format
// object_value_array[n] = option value
// object_value_array[n+1] = option text 1
// object_value_array[n+2] = option text 2
function EW_ajaxupdatecombo(obj, filter_value) {
	if (!(document.getElementsByTagName || document.all))
		return;
	try {
		var value = (obj.selectedIndex > -1) ? obj.options[obj.selectedIndex].value : null;
		for (var i = obj.length-1; i > 0; i--) {
			obj.options[i] = null;
		}
		var s = eval('obj.form.s_' + obj.name + '.value');
		//var s = eval('s_' + obj.name);
		//if (!s || s == '' || filter_value == '') return;
		if (!s || s == '') return;
		var lc = eval('obj.form.lc_' + obj.name + '.value');
		if (!lc || lc == '') return;
		var ld1 = eval('obj.form.ld1_' + obj.name + '.value');
		if (!ld1 || ld1 == '') return;
		var ld2 = eval('obj.form.ld2_' + obj.name + '.value');
		if (!ld2 || ld2 == '') return;
		var xmlHttp = EW_createXMLHttp();
		if (!xmlHttp) return;		
		xmlHttp.open('get', EW_LookupFn + '?s=' + s + '&q=' + encodeURIComponent(filter_value) +
			'&lc=' + encodeURIComponent(lc) +
			'&ld1=' + encodeURIComponent(ld1) +
			'&ld2=' + encodeURIComponent(ld2));
		xmlHttp.onreadystatechange = function() {
			//alert(xmlHttp.responseText);					
			if (xmlHttp.readyState == 4 && xmlHttp.status == 200 &&
				xmlHttp.responseText) {
				//alert(xmlHttp.responseText);
				var object_value_array = xmlHttp.responseText.split('\r');
				for (var j=0; j<object_value_array.length-2; j=j+3) {
					EW_newopt(obj, object_value_array[j], object_value_array[j+1],
						object_value_array[j+2]);
				}
				EW_selectopt(obj, value);
			}
		}		
		xmlHttp.send(null);
	}	catch (e) {}
}

function EW_HtmlEncode(text) {
	var str = text;
	str = str.replace(/&/g, '&amp');
	str = str.replace(/\"/g, '&quot;');
	str = str.replace(/</g, '&lt;');
	str = str.replace(/>/g, '&gt;'); 
	return str;
}

// Google Suggest for textbox by AJAX
// object_value_array format
// object_value_array[n] = display value
// object_value_array[n+1] = display value 2
function EW_ajaxupdatetextbox(object_name) {
	var obj, as;	
	if (document.all) {
		obj = document.all(object_name);
		if (obj) as = document.all('as_' + object_name);		
	} else if (document.getElementById) {
		obj = document.getElementById(object_name);
		if (obj) as = document.getElementById('as_' + object_name);
	}	
	if (!obj || !as) return false;
	try {
		var s = eval('obj.form.s_' + obj.name + '.value');
		//var s = eval('s_' + obj.name);
		var q = obj.value;
		q = q.replace(/^\s*/, ''); // left trim				
		if (!s || s == '' || q.length == 0) return false;					
		var lt = eval('obj.form.lt_' + obj.name + '.value');
		if (!lt || lt == '') return;
		var xmlHttp = EW_createXMLHttp();
		if (!xmlHttp) return;				
		xmlHttp.open('get', EW_LookupFn + '?s=' + s + '&q=' + encodeURIComponent(q) +
			'&lt=' + encodeURIComponent(lt));
		xmlHttp.onreadystatechange = function() {
			//if (xmlHttp.readyState == 4) alert(xmlHttp.responseText);
			if (xmlHttp.readyState == 4 && xmlHttp.status == 200 &&
				xmlHttp.responseText) {										
				var object_value_array = xmlHttp.responseText.split('\r');
				var sHtml = '';
				for (var j=0; j<object_value_array.length-2; j=j+2) {
					var value = object_value_array[j];
					var text = object_value_array[j];
					if (object_value_array[j+1] != "")
						text += EW_fieldSep + object_value_array[j+1];
					var i = j/2 + 1;
					sCtrlID = object_name + "_mi_" + i;
					sFunc1 = "EW_astOnMouseClick(" + i + ", \"" + object_name + "\", \"" + as.id + "\")";
					sFunc2 = "EW_astOnMouseOver(" + i + ", \"" + object_name + "\")";
					sHtml += "<div class=\"ewAstListItem\" id=\"" + sCtrlID + "\" name=\"" + sCtrlID + "\" onclick='" + sFunc1 + "' + onmouseover='" + sFunc2 + "'>" + text + "</div>";
					// add hidden field to store the value of current item
					sMenuItemValueID = sCtrlID + "_value";
					sHtml += "\n\r";
					sHtml += "<input type=\"hidden\" id=\"" + sMenuItemValueID + "\" name=\"" + sMenuItemValueID + "\" value=\"" + EW_HtmlEncode(text) + "\">";
				}
				//alert(sHtml);	
				EW_astShowDiv(as.id, sHtml);
			} else {
				EW_astHideDiv(as.id);
			}
		}
		xmlHttp.send(null);
	}	catch (e) {}
	return false;
}

// Extended basic search clear form
function EW_clearForm(objForm){
	with (objForm) {
		for (var i=0; i<elements.length; i++){
			var tmpObj = eval(elements[i]);
			if (tmpObj.type == "checkbox" || tmpObj.type == "radio"){
				tmpObj.checked = false;
			} else if (tmpObj.type == "select-one"){
				tmpObj.selectedIndex = 0;
			} else if (tmpObj.type == "select-multiple") {
				for (var j=0; j<tmpObj.options.length; j++)
					tmpObj.options[j].selected = false;
			} else if (tmpObj.type == "text"){
				tmpObj.value = "";
			}
		}
	}
}

// Functions for adding new option dynamically

function EW_ShowAddOption(id) {
	if (!document.getElementById) return;
	var elem;
	elem = document.getElementById("ao_" + id);
	if (elem) elem.style.display = "block"; 
	elem = document.getElementById("cb_" + id);
	if (elem)	elem.style.display = "none";	
}

function EW_HideAddOption(id) {
	var elem;
	elem = document.getElementById("cb_" + id);
	if (elem)	elem.style.display = "inline"; 
	elem = document.getElementById("ao_" + id);
	if (elem) elem.style.display = "none"; 
}

function EW_PostNewOption(id) {
	var elem;
	var url = EW_AddOptFn + "?";
	elem = document.getElementById("ltn_" + id);
	url += "ltn=" + encodeURIComponent(elem.value);
	elem = document.getElementById("dfn_" + id);
	if (elem) url += "&dfn=" + encodeURIComponent(elem.value);
	elem = document.getElementById("dfq_" + id);
	if (elem) url += "&dfq=" + encodeURIComponent(elem.value);
	elem = document.getElementById("lfn_" + id);
	if (elem) url += "&lfn=" + encodeURIComponent(elem.value);
	elem = document.getElementById("lfq_" + id);
	if (elem) url += "&lfq=" + encodeURIComponent(elem.value);
	elem = document.getElementById("df2n_" + id);
	if (elem) url += "&df2n=" + encodeURIComponent(elem.value);
	elem = document.getElementById("df2q_" + id);
	if (elem) url += "&df2q=" + encodeURIComponent(elem.value);	
	
	var lf = document.getElementById("lf_" + id);
	var lfm = document.getElementById("lfm_" + id);
	if (lf) {
		if (EW_hasValue(lf, "TEXT")) {
			url += "&lf=" + encodeURIComponent(lf.value); 
		} else {
			if (!EW_onError(lf.form, lf, "TEXT", (lfm?lfm.value:"Missing link field value")))
				return false;		
		}
	}
	
	var df = document.getElementById("df_" + id);
	var dfm = document.getElementById("dfm_" + id);
	if (df) {
		if (EW_hasValue(df, "TEXT")) {
			url += "&df=" + encodeURIComponent(df.value); 
		} else {
			if (!EW_onError(df.form, df, "TEXT", (dfm?dfm.value:"Missing display field value")))
				return false;		
		}
	}
	
	var df2 = document.getElementById("df2_" + id);
	var df2m = document.getElementById("df2m_" + id);
	if (df2) {
		if (EW_hasValue(df2, "TEXT")) {
			url += "&df2=" + encodeURIComponent(df2.value); 
		} else {
			if (!EW_onError(df2.form, df2, "TEXT", (df2m?df2m.value:"Missing display field #2 value")))
				return false;		
		}
	}
	
	try {			
		var xmlHttp = EW_createXMLHttp();
		if (!xmlHttp) return;		
		xmlHttp.open('get', url, true); // not async					
		xmlHttp.onreadystatechange = function() {
			//alert(xmlHttp.responseText);					
			if (xmlHttp.readyState == 4 && xmlHttp.status == 200 &&
				xmlHttp.responseText) {				
				var opt = xmlHttp.responseText.split('\r');
				if (opt.length > 3 && opt[0]== 'OK') {
					var elem = document.getElementById(id);			
					if (elem) {																					
						EW_newopt(elem, opt[1], opt[2], opt[3]);								
						EW_HideAddOption(id);
						elem.options[elem.options.length-1].selected = true;
						elem.focus();
					}
				} else {
					alert(xmlHttp.responseText);
				}				
			}
		}		
		xmlHttp.send(null);
	}	catch (e) {}

}	
//-->
</script>
<script type="text/javascript">
<!--
var firstrowoffset = 1; // first data row start at
var tablename = 'ewlistmain'; // table name
var usecss = true; // use css
//var usecss = false; // use css
var rowclass = 'ewTableRow'; // row class
var rowaltclass = 'ewTableAltRow'; // row alternate class
var rowmoverclass = 'ewTableHighlightRow'; // row mouse over class
var rowselectedclass = 'ewTableSelectRow'; // row selected class
var roweditclass = 'ewTableEditRow'; // row edit class
var rowcolor = '#FFFFFF'; // row color
var rowaltcolor = '#F5F5F5'; // row alternate color
var rowmovercolor = '#FFFFFF'; // row mouse over color
var rowselectedcolor = '#A2A2A2'; // row selected color
var roweditcolor = '#FFFF99'; // row edit color
//-->
</script>
<script type="text/javascript">
<!--
var EW_DHTMLEditors = [];
//-->
</script>
<%

' Set up Record Set
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.Open sSql, conn, 1, 2
nTotalRecs = rs.RecordCount
If nDisplayRecs <= 0 Then ' Display All Records
	nDisplayRecs = nTotalRecs
End If
nStartRec = 1
SetUpStartRec() ' Set Up Start Record Position
%>
<link href="wsdscu.css" rel="stylesheet" type="text/css" />

<p><strong>Login User List</strong></p>
<%
If Session(ewSessionMessage) <> "" Then
%>
<p><span class="ewmsg"><%= Session(ewSessionMessage) %></span></p>
<%
	Session(ewSessionMessage) = "" ' Clear message
End If
%>
<% If nTotalRecs > 0 Then %>
<form method="post">
  <table id="ewlistmain" class="ewTable">
    <!-- Table header -->
    <tr class="ewTableHeader"> 
      <td>Edit</td>
      <td valign="top">UID</td>
      <td valign="top">UserName</td>
      <td valign="top">Level</td>
      <td valign="top">Last Login Time</td>
    </tr>
    <%

' Avoid starting record > total records
If CLng(nStartRec) > CLng(nTotalRecs) Then
	nStartRec = nTotalRecs
End If

' Set the last record to display
nStopRec = nStartRec + nDisplayRecs - 1

' Move to first record directly for performance reason
nRecCount = nStartRec - 1
If Not rs.Eof Then
	rs.MoveFirst
	rs.Move nStartRec - 1
End If
nRecActual = 0
Do While (Not rs.Eof) And (nRecCount < nStopRec)
	nRecCount = nRecCount + 1
	If CLng(nRecCount) >= CLng(nStartRec) Then
		nRecActual = nRecActual + 1

	' Set row color
	sItemRowClass = " class=""ewTableRow"""
	sListTrJs = " onmouseover='ew_mouseover(this);' onmouseout='ew_mouseout(this);' onclick='ew_click(this);'"

	' Display alternate color for rows
	If nRecCount Mod 2 <> 1 Then
		sItemRowClass = " class=""ewTableAltRow"""
	End If
	x_uid = rs("uid")
	x_username = rs("username")
	x_password = rs("password")
	x_userLevel = rs("userLevel")
	x_lastLoginTime = rs("lastLoginTime")
%>
    <!-- Table body -->
    <tr<%=sItemRowClass%><%=sListTrJs%>> 
      <td><a href="userRightsEdit.asp?PID=<%= Server.URLEncode(x_uid)%>">user 
        Rights Details</a></td>
      <!-- uid -->
      <td><span> 
        <% Response.Write x_uid %>
        </span></td>
      <!-- username -->
      <td><span> 
        <% Response.Write x_username %>
        </span></td>
      <!-- password -->
 	
      <!-- userLevel -->
      <td><span> 
        <% Response.Write x_userLevel %>
        </span></td>
      <!-- lastLoginTime -->
      <td><span> 
        <% Response.Write x_lastLoginTime %>
        </span></td>
    </tr>
    <%
	End If
	rs.MoveNext
Loop
%>
  </table>
</form>
<% End If %>
<%

' Close recordset and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
<form action="loginUserlist.asp" name="ewpagerform" id="ewpagerform">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td nowrap>
<span>
<%

' Display page numbers
If nTotalRecs > 0 Then
	rsEof = (nTotalRecs < (nStartRec + nDisplayRecs))
	If CLng(nTotalRecs) > CLng(nDisplayRecs) Then

		' Find out if there should be Backward or Forward Buttons on the TABLE.
		If 	nStartRec = 1 Then
			isPrev = False
		Else
			isPrev = True
			PrevStart = nStartRec - nDisplayRecs
			If PrevStart < 1 Then PrevStart = 1 %>
		<a href="loginUserlist.asp?start=<%=PrevStart%>"><b>Prev</b></a>
		<%
		End If
		If (isPrev Or (Not rsEof)) Then
			x = 1
			y = 1
			dx1 = ((nStartRec-1)\(nDisplayRecs*nRecRange))*nDisplayRecs*nRecRange+1
			dy1 = ((nStartRec-1)\(nDisplayRecs*nRecRange))*nRecRange+1
			If (dx1+nDisplayRecs*nRecRange-1) > nTotalRecs Then
				dx2 = (nTotalRecs\nDisplayRecs)*nDisplayRecs+1
				dy2 = (nTotalRecs\nDisplayRecs)+1
			Else
				dx2 = dx1+nDisplayRecs*nRecRange-1
				dy2 = dy1+nRecRange-1
			End If
			While x <= nTotalRecs
				If x >= dx1 And x <= dx2 Then
					If CLng(nStartRec) = CLng(x) Then %>
		<b><%=y%></b>
					<%	Else %>
		<a href="loginUserlist.asp?start=<%=x%>"><b><%=y%></b></a>
					<%	End If
					x = x + nDisplayRecs
					y = y + 1
				ElseIf x >= (dx1-nDisplayRecs*nRecRange) And x <= (dx2+nDisplayRecs*nRecRange) Then
					If x+nRecRange*nDisplayRecs < nTotalRecs Then %>
		<a href="loginUserlist.asp?start=<%=x%>"><b><%=y%>-<%=y+nRecRange-1%></b></a>
					<% Else
						ny=(nTotalRecs-1)\nDisplayRecs+1
							If ny = y Then %>
		<a href="loginUserlist.asp?start=<%=x%>"><b><%=y%></b></a>
							<% Else %>
		<a href="loginUserlist.asp?start=<%=x%>"><b><%=y%>-<%=ny%></b></a>
							<%	End If
					End If
					x=x+nRecRange*nDisplayRecs
					y=y+nRecRange
				Else
					x=x+nRecRange*nDisplayRecs
					y=y+nRecRange
				End If
			Wend
		End If

		' Next link
		If NOT rsEof Then
			NextStart = nStartRec + nDisplayRecs
			isMore = True %>
		<a href="loginUserlist.asp?start=<%=NextStart%>"><b>Next</b></a>
		<% Else
			isMore = False
		End If %>
		<br>
<%	End If
	If CLng(nStartRec) > CLng(nTotalRecs) Then nStartRec = nTotalRecs
	nStopRec = nStartRec + nDisplayRecs - 1
	nRecCount = nTotalRecs - 1
	If rsEof Then nRecCount = nTotalRecs
	If nStopRec > nRecCount Then nStopRec = nRecCount %>
	Records <%= nStartRec %> to <%= nStopRec %> of <%= nTotalRecs %>
<% Else %>
	<% If sSrchWhere = "0=101" Then %>
	<% Else %>
	No records found
	<% End If %>
<% End If %>
</span>
		</td>
	</tr>
</table>
</form>

<%

'-------------------------------------------------------------------------------
' Function BasicSearchSQL
' - Build WHERE clause for a keyword

Function BasicSearchSQL(Keyword)
	Dim sKeyword
	sKeyword = AdjustSql(Keyword)
	BasicSearchSQL = ""
	BasicSearchSQL = BasicSearchSQL & "[username] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[password] LIKE '%" & sKeyword & "%' OR "
	If Right(BasicSearchSQL, 4) = " OR " Then BasicSearchSQL = Left(BasicSearchSQL, Len(BasicSearchSQL)-4)
End Function

'-------------------------------------------------------------------------------
' Function SetUpBasicSearch
' - Set up Basic Search parameter based on form elements pSearch & pSearchType
' - Variables setup: sSrchBasic

Sub SetUpBasicSearch()
	Dim arKeyword, sKeyword
	psearch = Request.QueryString(ewTblBasicSrch)
	psearchtype = Request.QueryString(ewTblBasicSrchType)
	If psearch <> "" Then
		If psearchtype <> "" Then
			While InStr(psearch, "  ") > 0
				sSearch = Replace(psearch, "  ", " ")
			Wend
			arKeyword = Split(Trim(psearch), " ")
			For Each sKeyword In arKeyword
				sSrchBasic = sSrchBasic & "(" & BasicSearchSQL(sKeyword) & ") " & psearchtype & " "
			Next
		Else
			sSrchBasic = BasicSearchSQL(psearch)
		End If
	End If
	If Right(sSrchBasic, 4) = " OR " Then sSrchBasic = Left(sSrchBasic, Len(sSrchBasic)-4)
	If Right(sSrchBasic, 5) = " AND " Then sSrchBasic = Left(sSrchBasic, Len(sSrchBasic)-5)
	If psearch <> "" then
		Session(ewSessionTblBasicSrch) = psearch
		Session(ewSessionTblBasicSrchType) = psearchtype
	End If
End Sub

'-------------------------------------------------------------------------------
' Function ResetSearch
' - Clear all search parameters
'

Sub ResetSearch()

	' Clear search where
	sSrchWhere = ""
	Session(ewSessionTblSearchWhere) = sSrchWhere

	' Clear advanced search parameters
	Session(ewSessionTblAdvSrch & "_x_uid") = ""
	Session(ewSessionTblAdvSrch & "_x_username") = ""
	Session(ewSessionTblAdvSrch & "_x_password") = ""
	Session(ewSessionTblAdvSrch & "_x_userLevel") = ""
	Session(ewSessionTblAdvSrch & "_x_lastLoginTime") = ""
	Session(ewSessionTblBasicSrch) = ""
	Session(ewSessionTblBasicSrchType) = ""
End Sub

'-------------------------------------------------------------------------------
' Function RestoreSearch
' - Restore all search parameters
'

Sub RestoreSearch()

	' Restore advanced search settings
	x_uid = Session(ewSessionTblAdvSrch & "_x_uid")
	x_username = Session(ewSessionTblAdvSrch & "_x_username")
	x_password = Session(ewSessionTblAdvSrch & "_x_password")
	x_userLevel = Session(ewSessionTblAdvSrch & "_x_userLevel")
	x_lastLoginTime = Session(ewSessionTblAdvSrch & "_x_lastLoginTime")
	psearch = Session(ewSessionTblBasicSrch)
	psearchtype = Session(ewSessionTblBasicSrchType)
End Sub

'-------------------------------------------------------------------------------
' Function SetUpSortOrder
' - Set up Sort parameters based on Sort Links clicked
' - Variables setup: sOrderBy, Session(TblOrderBy), Session(Tbl_Field_Sort)

Sub SetUpSortOrder()
	Dim sOrder, sSortField, sLastSort, sThisSort
	Dim bCtrl

	' Check for an Order parameter
	If Request.QueryString("order").Count > 0 Then
		sOrder = Request.QueryString("order")

		' Field [uid]
		If sOrder = "uid" Then
			sSortField = "[uid]"
			sLastSort = Session(ewSessionTblSort & "_x_uid")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_uid") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_uid") <> "" Then Session(ewSessionTblSort & "_x_uid") = ""
		End If

		' Field [username]
		If sOrder = "username" Then
			sSortField = "[username]"
			sLastSort = Session(ewSessionTblSort & "_x_username")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_username") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_username") <> "" Then Session(ewSessionTblSort & "_x_username") = ""
		End If

		' Field [password]
		If sOrder = "password" Then
			sSortField = "[password]"
			sLastSort = Session(ewSessionTblSort & "_x_password")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_password") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_password") <> "" Then Session(ewSessionTblSort & "_x_password") = ""
		End If

		' Field [userLevel]
		If sOrder = "userLevel" Then
			sSortField = "[userLevel]"
			sLastSort = Session(ewSessionTblSort & "_x_userLevel")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_userLevel") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_userLevel") <> "" Then Session(ewSessionTblSort & "_x_userLevel") = ""
		End If

		' Field [lastLoginTime]
		If sOrder = "lastLoginTime" Then
			sSortField = "[lastLoginTime]"
			sLastSort = Session(ewSessionTblSort & "_x_lastLoginTime")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_lastLoginTime") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_lastLoginTime") <> "" Then Session(ewSessionTblSort & "_x_lastLoginTime") = ""
		End If
		Session(ewSessionTblOrderBy) = sSortField & " " & sThisSort
		Session(ewSessionTblStartRec) = 1
	End If
	sOrderBy = Session(ewSessionTblOrderBy)
	If sOrderBy = "" Then
		sOrderBy = ewSqlOrderBy
		Session(ewSessionTblOrderBy) = sOrderBy
		If sOrderBy <> "" Then
			Dim arOrderBy, i
			arOrderBy = Split(ewSqlOrderBySessions, ",")
			For i = 0 to UBound(arOrderBy)\2
				Session(ewSessionTblSort & "_" & arOrderBy(i*2)) = arOrderBy(i*2+1)
			Next
		End If
	End If
End Sub

'-------------------------------------------------------------------------------
' Function SetUpStartRec
' - Set up Starting Record parameters based on Pager Navigation
' - Variables setup: nStartRec

Sub SetUpStartRec()
	Dim nPageNo

	' Check for a START parameter
	If Request.QueryString(ewTblStartRec).Count > 0 Then
		nStartRec = Request.QueryString(ewTblStartRec)
		Session(ewSessionTblStartRec) = nStartRec
	ElseIf Request.QueryString("pageno").Count > 0 Then
		nPageNo = Request.QueryString("pageno")
		If IsNumeric(nPageNo) Then
			nStartRec = (nPageNo-1)*nDisplayRecs+1
			If nStartRec <= 0 Then
				nStartRec = 1
			ElseIf nStartRec >= ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1 Then
				nStartRec = ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1
			End If
			Session(ewSessionTblStartRec) = nStartRec
		Else
			nStartRec = Session(ewSessionTblStartRec)
			If Not IsNumeric(nStartRec) Or nStartRec = "" Then
				nStartRec = 1 ' Reset start record counter
				Session(ewSessionTblStartRec) = nStartRec
			End If
		End If
	Else
		nStartRec = Session(ewSessionTblStartRec)
		If Not IsNumeric(nStartRec) Or nStartRec = "" Then
			nStartRec = 1 'Reset start record counter
			Session(ewSessionTblStartRec) = nStartRec
		End If
	End If
End Sub

'-------------------------------------------------------------------------------
' Function ResetCmd
' - Clear list page parameters
' - RESET: reset search parameters
' - RESETALL: reset search & master/detail parameters
' - RESETSORT: reset sort parameters

Sub ResetCmd()
	Dim sCmd

	' Get Reset Cmd
	If Request.QueryString("cmd").Count > 0 Then
		sCmd = Request.QueryString("cmd")

		' Reset Search Criteria
		If LCase(sCmd) = "reset" Then
			Call ResetSearch()

		' Reset Search Criteria & Session Keys
		ElseIf LCase(sCmd) = "resetall" Then
			Call ResetSearch()

		' Reset Sort Criteria
		ElseIf LCase(sCmd) = "resetsort" Then
			sOrderBy = ""
			Session(ewSessionTblOrderBy) = sOrderBy
			If Session(ewSessionTblSort & "_x_uid") <> "" Then Session(ewSessionTblSort & "_x_uid") = ""
			If Session(ewSessionTblSort & "_x_username") <> "" Then Session(ewSessionTblSort & "_x_username") = ""
			If Session(ewSessionTblSort & "_x_password") <> "" Then Session(ewSessionTblSort & "_x_password") = ""
			If Session(ewSessionTblSort & "_x_userLevel") <> "" Then Session(ewSessionTblSort & "_x_userLevel") = ""
			If Session(ewSessionTblSort & "_x_lastLoginTime") <> "" Then Session(ewSessionTblSort & "_x_lastLoginTime") = ""
		End If

		' Reset Start Position (Reset Command)
		nStartRec = 1
		Session(ewSessionTblStartRec) = nStartRec
	End If
End Sub
%>

