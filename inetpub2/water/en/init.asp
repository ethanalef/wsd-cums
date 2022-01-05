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
               ShareCode = "股金結餘"
          case "A0"
               ShareCode = "退還貸款"
           case "A1"
               ShareCode = "銀行轉帳"
          case "A2"
		ShareCode ="庫房轉帳"
          case "A3"
		ShareCode ="現金存款"
          case "A4"
              ShareCode ="保險金"
          case "B0"
               ShareCode="股金還款"
          case "A7"
                  ShareCode ="調整"
          case "B1"
             
                   ShareCode="退股"

          CASE "AI"
                ShareCode ="脫期　　" 
          CASE "D1"
               ShareCode ="新貸銀行"  
          CASE "B0"
                ShareCode ="現金退股"
         case "B3"
                ShareCode ="退還現金"
                
          case "C0"
               ShareCode="股息　　"
           case "CH"
               ShareCode="暫停股息"        
          case "C1"
               ShareCode="股息銀行" 
          case "C3"
             ShareCode="股息現金" 
         case "C5"
             ShareCode="股息還款" 
          case "G0","G1","G2","G3"
                ShareCode = "入社費"
          case "H0","H1","H2","H3"
            
              ShareCode = "協會費" 
         case "MF"
            
              ShareCode = "冷戶費" 
	end select
End Function

Function LoanCode(ByVal x)
	select case x
          case "0D"
               LoanCode="貸款結餘"
          case "E1"
               LoanCode= "銀行轉帳"
          case "E2"
		 LoanCode="庫房轉帳"
          case "EC"
                LoanCode="劃消金額"
          case "E3"
		 LoanCode="現金還款"
          case "E0"
               if ms("amount") > 0 then
  		   LoanCode="股金還款"
               else
                  LoanCode="退還本息"
               end if 
          case "F0"
              if ms("amount") > 0 then
  		   LoanCode="股金還款"
               else
                  LoanCode="退還利息"
               end if 
                  
          case "E6"
                    LoanCode="退款"
          case "E7"
                  LoanCode="調整"

          case "F1"
                 LoanCode="銀行還息"
          case "F2"
                 LoanCode="庫房還息"
          case "F3"
                 LoanCode="現金還息"
          case "ER"
		 LoanCode="退還本金"
          case "F3"
                 LoanCode="現金還息"  
          case "FR"
		 LoanCode="退還利息"
          CASE "DE"
               LoanCode="銀行脫期" 
          CASE "DF"
               LoanCode="利息脫期" 
          CASE "NE"
               LoanCode="庫房脫期"            
  
          CASE "D8"
             
                LoanCode="貸款清數"
               
          case "DE","NE"
              mx = 0  
	end select
End Function

%>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../main.css" rel="stylesheet" type="text/css">
<script language="javascript" src="prototype-1.6.0.2.js"></script>
<script language="javascript" src="function.js"></script>
