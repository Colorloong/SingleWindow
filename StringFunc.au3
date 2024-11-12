#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <Array.au3>
#include <XML.au3>

; Strip leading and trailing whitespace as well as the double spaces (or more) in between the words.
;~ Local $aString[] = ["Say Total :  TWO PALLETS (ONE HUNDRED CARTONS) ONLY", _
;~ "Say Total :  ONE CARTONS ONLY", _
;~ "SAY TOTAL TWO(2) CTN IN ONE(1) PKG ONLY", _
;~ "Say Total : ONE PARCELS (TWO CARTONS) ONLY", _
;~ "SAY TOTAL NINETY-ONE(91) CTN IN ONE(1) PLT ONLY AND TWO(2) CTN IN ONE(1) PKG ONLY AND ONE(1) CTN ONL", _
;~ "= 7 Packages"]
;~ _ArrayDisplay($aString)
Func zgGetPkgNumType($sString)
   Local $atmp
   Local $i,$j
   $pkgNum = 0
   $sString = StringReplace($sString,"THOUSAND AND","")
   $sString = StringReplace($sString,"THOUSAND","")
   $sString = StringReplace($sString,"HUNDRED AND","")
   $sString = StringReplace($sString,"HUNDRED","")
   If StringLeft($sString,9)="Say Total" Then
	  If StringInStr($sString," AND ")>0 Then ;"SAY TOTAL NINETY-ONE(91) CTN IN ONE(1) PLT ONLY AND TWO(2) CTN IN ONE(1) PKG ONLY AND ONE(1) CTN ONL", _
;~ 	  MsgBox(0,$sString,$sString & @CRLF&StringInStr($sString," AND "))
		 $sString = $sString & " AND"
		 $aTmp = StringSplit($sString,"() ",$STR_CHRSPLIT)
		 $i = UBound($aTmp) - 1
		 While $i > 0
			If $aTmp[$i] = "AND" Then
			   $j = $i-1
			   While $aTmp[$j]<>"AND" Or $j>0
				  If StringIsDigit($aTmp[$j]) Then
					 $pkgNum = $pkgNum + $aTmp[$j]
					 ExitLoop
				  EndIf
				  $j = $j - 1
			   WEnd
			EndIf
			$i = $i - 1
		 WEnd
		 $pkgType = "'06";"PACKAGE"	;20180825注释 改用 zgFindTextlArray7
	  ElseIf IsNumInStr($sString) Then ;"SAY TOTAL TWO(2) CTN IN ONE(1) PKG ONLY", _
		 $aTmp = StringRegExp($sString,"\d+",$STR_REGEXPARRAYGLOBALMATCH)
		 $pkgNum = $aTmp[UBound($aTmp)-1]
		 $aTmp = StringSplit($sString," ",$STR_NOCOUNT)
		 $pkgType = $aTmp[UBound($aTmp)-2]	;20180825注释 改用 zgFindTextlArray7
	  Else
		 $sTmp = $sString
		 If StringLeft($sString,11)="Say Total :" Then
			$sTmp = StringTrimLeft($sString,11)
		 ElseIf StringLeft($sString,9)="Say Total" Then
			$sTmp = StringTrimLeft($sString,9)
		 EndIf
		 $sTmp = StringStripWS($sTmp,3)
		 $sTmp = TransEn2Num($sTmp)
		 $aTmp = StringSplit($sTmp," ",$STR_NOCOUNT)
		 $pkgType = $aTmp[1]	;20180825注释 改用 zgFindTextlArray7
		 $aTmp = StringRegExp($sTmp,"\d+",$STR_REGEXPARRAYGLOBALMATCH)
		 $pkgNum = $aTmp[0]
	  EndIf
   ElseIf StringLeft($sString,1)="=" Then
	  $atmp = StringSplit($sString," ",$STR_NOCOUNT)
	  $pkgNum = $aTmp[1]
	  $pkgType = $aTmp[2]	;20180825注释 改用 zgFindTextlArray7
   EndIf
   Return $pkgType
EndFunc

Func TransEn2Num($sStr)
   Local $sTmp
   $sTmp = $sStr
   Local $i
   For $i=UBound($aTargetEnNum)-1 To 0 Step -1
	  $sTmp = StringReplace($sTmp,$aTargetEnNum[$i][1],$aTargetEnNum[$i][0])
   Next
   Return $sTmp
EndFunc

Func TransPkgType($sStr)
   Local $i
   For $i=UBound($aPKGType)-1 To 0 Step -1
	  If StringInStr($sStr,$aPKGType[$i][1])>0 Then
		 return $aPKGType[$i][0]
	  EndIf
   Next
   Return 0
EndFunc

Func IsNumInStr($sStr)
   Local $NumInStr=False
   For $i = 0 To 9 Step 1
	  If StringInStr($sStr,$i) > 0 Then
		 $NumInStr=True
		 ExitLoop
	  EndIf
   Next
   return $NumInStr
EndFunc

Func TransMonNum($sStr)
   Local $i
   Local $sTmp,$Year
   For $i=UBound($aMonEn)-1 To 0 Step -1
	  If StringInStr($sStr,$aMonEn[$i][1])>0 Then
		 $sTmp = StringRegExpReplace(StringReplace($sStr,$aMonEn[$i][1],$aMonEn[$i][0]),"(\d{2})(\d{2})","$2$1")
		 If StringLeft($sTmp,2)<@MON Then
			$sTmp = @YEAR + 1 & $sTmp
		 Else
			$sTmp = @YEAR & $sTmp
		 EndIf
		 ExitLoop
	  EndIf
   Next
   If $i<0 Then
	  MsgBox(0,"月份代码未找到",$sStr)
	  Return
   EndIf
   Return $sTmp
EndFunc

;~ 指运港
Func TransTragetPort($sStr)
   Local $i
   Local $sTmp
   For $i=UBound($aTargetPort)-1 To 0 Step -1
	  If StringInStr($sStr,$aTargetPort[$i][1])>0 Then
		 $sTmp = $aTargetPort[$i][0]
		 ExitLoop
	  EndIf
   Next
   If $i<0 Then
	  MsgBox(0,"口岸代码未找到",$sStr)
	  Return
   EndIf
   Return $sTmp
EndFunc

;~ 运抵国  国家名称=>代码(字母)
Func TransTargetAdd($sValue)
   Local $i,$find
   For $i=0 To UBound($aTargetAdd)-1 Step 1
	  If StringInStr(StringStripWS($sValue,$STR_STRIPALL),StringStripWS($aTargetAdd[$i][1],$STR_STRIPALL))>0 Then
		 $find = True
		 ExitLoop
	  EndIf
   Next
   If Not $find Then
	  SetError(1)
	  Return ""
   EndIf
   Return $aTargetAdd[$i][0]
EndFunc

;~ 运抵国  代码(字母)=>代码(数字)
Func TransTargetAddCodeToNum($sValue)
   Local $i,$find
   For $i=0 To UBound($aTargetAdd)-1 Step 1
	  If StringInStr($sValue,$aTargetAdd[$i][0])>0 Then
		 $find = True
		 ExitLoop
	  EndIf
   Next
   If Not $find Then
	  SetError(1)
	  Return ""
   EndIf
   Return $aTargetAdd[$i][2]
EndFunc

Func CountItemInArray($sStr)
   Local $i
   Local $count=0
   For $i=0 To UBound($aInputInfo)-1 Step 1
	  If $aInputInfo[$i][15] = $sStr Then $count=$count+1
   Next
   Return $count
EndFunc

Func GetOverseasConsigneeEname($sStr)
   Local $i
   Local $sTmp
   For $i=UBound($aOverseasConsigneeEname)-1 To 0 Step -1
	  If $sStr = $aOverseasConsigneeEname[$i][0] Then
		 $sTmp = $aOverseasConsigneeEname[$i][1]
		 ExitLoop
	  EndIf
   Next
   If $i<0 Then
	  MsgBox($MB_ICONERROR,"配置文件缺失","指运港代码未找到【"&$sStr&"】")
	  Return
   EndIf
   Return $sTmp
EndFunc