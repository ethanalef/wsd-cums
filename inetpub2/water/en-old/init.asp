<%
function PMT(ByVal Rate, nPer, PV, FV, iType)
    Rate = Rate * 0.01
    if Rate<>0 and nPer<>0 then
	    PMT  = (-PV*(1+Rate)^nPer - FV)*Rate/((1+Rate*iType)*((1+Rate)^nPer-1))
	    PMT  = round(Int(-PMT*100)/100,2)
    else
    	PMT = 0
    end if
end function

Function longDate(ByVal date)
	if not isdate(date) then
		longDate=""
	else
		longDate=day(date)&" "&ArrMonth(month(date)-1)&", "&year(date)
	end if
end Function

Function dmy(ByVal date)
	if not isdate(date) then
		dmy=""
	else
		dmy=right("0"&day(date),2)&"/"&right("0"&month(date),2)&"/"&right(year(date),4)
	end if
end Function

Function ymd(ByVal date)
	if not isdate(date) then
		ymd=""
	else
		ymd=year(date)&"/"&month(date)&"/"&day(date)
	end if
end Function

Function yymmdd(ByVal date)
	date = split(date,"/")
	if ubound(date)<>2 then
		yymmdd=""
	else
		yymmdd=date(2)&"/"&date(1)&"/"&date(0)
		if not isdate(yymmdd) then
			yymmdd=""
		end if
	end if
end Function

function iif(ByVal psdStr, trueStr, falseStr)
  if psdStr then iif = trueStr else iif = falseStr end if
end function

Function SimpleRound(ByVal number,decPoints)
	decPoints = 10^decPoints
	SimpleRound = round(number*decPoints+0.1)/decPoints
End Function

Function GetFiscalYear (ByVal x)
	If x < DateSerial(Year(x), FMonthStart, 1) Then
		GetFiscalYear = Year(x) - 1
	Else
		GetFiscalYear = Year(x)
	End If
End Function

Function GetFiscalMonth (ByVal x)
	Dim m
	m = Month(x) - FMonthStart + 1
	If m < 1 Then m = m + 12
	GetFiscalMonth = m
End Function

Function GetNormalYear (ByVal x)
	If month(x) > 13 - FMonthStart Then
		GetNormalYear = Year(x) + 1
	Else
		GetNormalYear = Year(x)
	End If
End Function

Function GetNormalMonth (ByVal x)
	Dim m
	m = x + FMonthStart - 1
	If m > 12 Then m = m - 12
	GetNormalMonth = m
End Function

Function ShareCode(ByVal x)
	select case x
          case "0A"
               ShareCode = "�Ѫ����l"
          case "A0"
               ShareCode = "�h�ٶU��"
           case "A1"
               ShareCode = "�Ȧ���b"
          case "A2"
		ShareCode ="�w����b"
          case "A3"
		ShareCode ="�{���s��"
          case "A4"
              ShareCode ="�O�I��"
          case "B0"
               ShareCode="�Ѫ��ٴ�"
          case "A7"
                  ShareCode ="�վ�"
          case "B1"
             
                   ShareCode="�h��"

          CASE "AI"
                ShareCode ="����@�@" 
          CASE "D1"
               ShareCode ="�s�U�Ȧ�"  
          CASE "B0"
                ShareCode ="�{���h��"
         case "B3"
                ShareCode ="�h�ٲ{��"
                
          case "C0"
               ShareCode="�Ѯ��@�@"
           case "CH"
               ShareCode="�Ȱ��Ѯ�"        
          case "C1"
               ShareCode="�Ѯ��Ȧ�" 
          case "C3"
             ShareCode="�Ѯ��{��" 
         case "C5"
             ShareCode="�Ѯ��ٴ�" 
          case "G0","G1","G2","G3"
                ShareCode = "�J���O"
          case "H0","H1","H2","H3"
            
              ShareCode = "��|�O" 
         case "MF"
            
              ShareCode = "�N��O" 
	end select
End Function

Function LoanCode(ByVal x)
	select case x
          case "0D"
               LoanCode="�U�ڵ��l"
          case "E1"
               LoanCode= "�Ȧ���b"
          case "E2"
		 LoanCode="�w����b"
          case "EC"
                LoanCode="�������B"
          case "E3"
		 LoanCode="�{���ٴ�"
          case "E0"
               if ms("amount") > 0 then
  		   LoanCode="�Ѫ��ٴ�"
               else
                  LoanCode="�h�٥���"
               end if 
          case "F0"
              if ms("amount") > 0 then
  		   LoanCode="�Ѫ��ٴ�"
               else
                  LoanCode="�h�٧Q��"
               end if 
                  
          case "E6"
                    LoanCode="�h��"
          case "E7"
                  LoanCode="�վ�"

          case "F1"
                 LoanCode="�Ȧ��ٮ�"
          case "F2"
                 LoanCode="�w���ٮ�"
          case "F3"
                 LoanCode="�{���ٮ�"
          case "ER"
		 LoanCode="�h�٥���"
          case "F3"
                 LoanCode="�{���ٮ�"  
          case "FR"
		 LoanCode="�h�٧Q��"
          CASE "DE"
               LoanCode="�Ȧ���" 
          CASE "DF"
               LoanCode="�Q�����" 
          CASE "NE"
               LoanCode="�w�в��"            
  
          CASE "D8"
             
                LoanCode="�U�ڲM��"
               
          case "DE","NE"
              mx = 0  
	end select
End Function

%>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="javascript" src="prototype-1.6.0.2.js"></script>
<script language="javascript" src="function.js"></script>
