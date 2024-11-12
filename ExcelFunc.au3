#include <Array.au3>
#include <MyExcel.au3>
#include <String.au3>
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>

Global $aTSCode[0][0],$aTargetAdd[0][0],$aTargetEnNum[0][0],$aPKGType[0][0],$aDelayDay[0][0],$aMonEn[0][0],$aTargetPort[0][0],$aOverseasConsigneeEname[0][0],$aPacketType[0][0]

Func GetConfigData()
	WriteMemo(0,"��ȡEXCEL�����ļ��У����Ե�......","",False)
    Local $oExcel = _Excel_Open(False)
   If @error Then
;~ 	  MsgBox($MB_ICONERROR,@ScriptName, "Error open EXCEL." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Return
   EndIf

   Local $sExcelFileName = @ScriptDir&"\QPDataConfig.xlsx"
   zgReadExcelToArray($aTSCode,$oExcel,$sExcelFileName, "������Ʒ����")
   zgReadExcelToArray($aTargetAdd,$oExcel,$sExcelFileName, "�˵ֹ�")
   zgReadExcelToArray($aTargetEnNum,$oExcel,$sExcelFileName, "����Ӣ��")
   zgReadExcelToArray($aPKGType,$oExcel,$sExcelFileName, "��װ����")
   zgReadExcelToArray($aDelayDay,$oExcel,$sExcelFileName, "�ڰ��ӳ�")
   zgReadExcelToArray($aMonEn,$oExcel,$sExcelFileName, "�·�Ӣ��")
   zgReadExcelToArray($aTargetPort,$oExcel,$sExcelFileName, "�걨�ڰ�")
   zgReadExcelToArray($aOverseasConsigneeEname,$oExcel,$sExcelFileName, "ָ�˸�")
   zgReadExcelToArray($aPacketType,$oExcel,$sExcelFileName, "��װ����")

   _Excel_Close($oExcel)
	WriteMemo(0,"��ȡEXCEL�����ļ����......","",False)
EndFunc

Func zgGetInfoFromFile($sDir,$sContractName)
   Local $sFullPathFileName
   ; Create application object and open an example workbook
    Local $oExcel = _Excel_Open(False)
   If @error Then
;~ 	  MsgBox($MB_ICONERROR,@ScriptName, "Error open EXCEL." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Return
   EndIf

   $sFullPathFileName = $sDir & "\" & $sContractName & "\" & $sContractName & "_EXP_Contract.xlsx"
   zgReadExcelToArray($aExcelFinal,$oExcel,$sFullPathFileName)
;~ _ArrayDisplay($aExcelFinal)

   $sFullPathFileName = $sDir & "\" & $sContractName & "\EXPCI_" & $sContractName & "*.xls"
   $sFullPathFileName = zgFindFile($sFullPathFileName)
   If @error Then
	  SetError(1,2)
	  Return
   EndIf
   $sFullPathFileName = $sDir & "\" & $sContractName & "\" & $sFullPathFileName
;~    WriteMemo(0,"|"&$sContractName&"|","2|"&$sFullPathFileName&"|")
   zgReadExcelToArray($aExcelExpci,$oExcel,$sFullPathFileName)
;~ _ArrayDisplay($aExcelExpci)

   $sFullPathFileName = $sDir & "\" & $sContractName & "\" & $sContractName & "_PACKING_LIST_EXP_FINAL.txt"
;~    MsgBox(0,"",$PdfFinal & @CRLF & $PdfExpci & @CRLF & $PdfText)
;~    $sFullPathFileName = zgFindFile($sFullPathFileName)
;~    MsgBox(0,"|"&$sContractName&"|","3|"&$sFullPathFileName&"|")
;~    If @error Then
;~ 	  SetError(1,3)
;~ 	  Return
;~    EndIf
;~    $sFullPathFileName = $sDir & "\" & $sContractName & "\" & $sFullPathFileName
   $aTextList = zgReadTextToArray($sFullPathFileName)
;~ _ArrayDisplay($aTextList)

   SetError(0)
	_Excel_Close($oExcel)

EndFunc

Func zgFindFile($sSearchString)
   Local $hSearch = FileFindFirstFile($sSearchString)

    ; Check if the search was successful, if not display a message and return False.
    If $hSearch = -1 Then
;~         MsgBox($MB_ICONERROR, "", "Error: No files/directories matched the search pattern.")
	  SetError(1,1)
      Return
    EndIf

    ; Assign a Local variable the empty string which will contain the files names found.
    Local $sFileName = "", $iResult = 0

	 $sFileName = FileFindNextFile($hSearch)
	 ; If there is no more file matching the search.
	 If @error Then
		MsgBox($MB_ICONERROR,"","��Ŀ¼" & $sDir & "\��δ�ҵ���ͬ" & $sContractName & "������ļ���")
		SetError(1,2)
;~ 		return
	 EndIf

	 ; Display the file name.
;~ 	 $iResult = MsgBox(BitOR($MB_ICONERROR, $MB_OKCANCEL), "", "File: " & $sFileName)

    ; Close the search handle.
    FileClose($hSearch)
	SetError(0)
	return $sFileName
 EndFunc

Func zgReadExcelToArray(ByRef $aExcelRead,ByRef $oExcel,$sExcelFileName,$vSheet=Default)

   Local $oWorkbook = _Excel_BookOpen($oExcel, $sExcelFileName)
   If @error Then
;~ 	  MsgBox($MB_ICONERROR,@ScriptName, "Error open the workbook." & @CRLF & $sExcelFileName & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Return
   EndIf

   ; *****************************************************************************
   ; Read data from a single cell on the active sheet of the specified workbook
   ; *****************************************************************************
   $aExcelRead = _Excel_RangeRead($oWorkbook, $vSheet, Default, 1, True)
   If @error Then
;~ 	  MsgBox($MB_ICONERROR,@ScriptName, "Error reading from workbook." & @CRLF & $sExcelFileName & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Return
   EndIf

   _Excel_BookClose($oWorkbook)
   If @error Then
;~ 	  MsgBox($MB_ICONERROR, "Excel UDF: _Excel_BookClose", "Error closing workbook." & @CRLF & $sExcelFileName & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  SetError(1)
	  Return
   EndIf

   SetError(0)
EndFunc

Func InsertIntoArray(ByRef $BaseArray,$Range,ByRef $value)
   If _ArraySearch($BaseArray, $value[0][1], 0, 0, 0, 1, 1, 1) >= 0 Then Return 0
   If $Range >= UBound($BaseArray) Then
	  _ArrayAdd($BaseArray,$value)
   Else
	  _ArrayInsert($BaseArray,$Range,$value)
   EndIf
   if @error then	  return 0
   Return 1
EndFunc

Func zgSaveOutputData($oFileName,ByRef $BaseInfoList,ByRef $DecList,$bAskOverWrite=False)
   ; Create application object and open an example workbook
   Local $oExcel = _Excel_Open(False)
   If @error Then
	  MsgBox($MB_ICONERROR,@ScriptName, "Error open EXCEL." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Return
   EndIf

   Local $oWorkbook = _Excel_BookNew($oExcel,2)
   If @error Then
	   MsgBox($MB_ICONERROR, "Excel UDF: _Excel_BookNew", "Error creating the new workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	   Return
   EndIf
;~    _Excel_SheetAdd($oWorkbook, Default, False, 2)

   ; *****************************************************************************
   ; Write a part of a 2D array to the active sheet in the active workbook
   ; *****************************************************************************
;~    �������¹���Ժ�ͬ������˫ԭ������ԭ�����й�(��),ԭ����Խ��(��)
   Local $MainTable[0][40]
   Local $Double = 0,$CN = 0,$VN = 0
   For $i = 0 To UBound($aDecList) - 1 Step 1
	  Local $aResult = _ArrayFindAll($aDecList, $aDecList[$i][0], Default, Default, Default, Default, 0)
	  If UBound($aResult) = 2 Then
		 Local $iIndex = _ArraySearch($BaseInfoList, $aDecList[$i][0], 0, 0, 0, 0, 1, 1)
		 Local $baseinfo = _ArrayExtract($BaseInfoList, $iIndex, $iIndex)
		 $Double += InsertIntoArray($MainTable,$Double,$baseinfo)
	  Else
		 If ($aDecList[$i][1] = "CN") Then
			Local $iIndex = _ArraySearch($BaseInfoList, $aDecList[$i][0], 0, 0, 0, 0, 1, 1)
			Local $baseinfo = _ArrayExtract($BaseInfoList, $iIndex, $iIndex)
			$CN += InsertIntoArray($MainTable,$Double + $CN,$baseinfo)
		 Else
			Local $iIndex = _ArraySearch($BaseInfoList, $aDecList[$i][0], 0, 0, 0, 0, 1, 1)
			Local $baseinfo = _ArrayExtract($BaseInfoList, $iIndex, $iIndex)
			$VN += InsertIntoArray($MainTable,$Double + $CN + $VN,$baseinfo)
		 EndIf
	  EndIf
   Next

   _Excel_RangeWrite($oWorkbook, "Sheet1", $MainTable)
   If @error Then
	  MsgBox($MB_ICONERROR, "Excel UDF: _Excel_RangeWrite", "Error writing to worksheet.Sheet1" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Return
   EndIf


   If UBound($DecList) = 0 Then
	  zgReadExcelToArray($aDecList,$oExcel,$oFileName, "Sheet2")
   EndIf

   _Excel_RangeWrite($oWorkbook, "Sheet2", $DecList)
	  If @error Then
		 MsgBox($MB_ICONERROR, "Excel UDF: _Excel_RangeWrite", "Error writing to worksheet.Sheet2" & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		 Return
	  EndIf


   ; *****************************************************************************
   ; Save the workbook (xls) in another format (html) to another directory and
   ; overwrite an existing version
   ; *****************************************************************************
   If $bAskOverWrite And FileExists($oFileName) Then
	  if MsgBox($MB_YESNO,"��ʾ","�ļ��Ѿ����ڣ��Ƿ񸲸ǣ�") = $IDNO Then
		 Return
	  EndIf
   EndIf
;~    FileDelete($oFileName)

   _Excel_BookSaveAs($oWorkbook,$oFileName,Default,True)
   If @error Then
	  MsgBox($MB_ICONERROR, "Excel UDF: _Excel_BookSaveAs", "Error saving workbook to '" & $oFileName & "'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Return
   EndIf

   _Excel_BookClose($oWorkbook)
   If @error Then
	  MsgBox($MB_ICONERROR, "Excel UDF: _Excel_BookClose", "Error saving workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Return
   EndIf

   _Excel_Close($oExcel)
EndFunc

Func zgReadTextToArray($sTextFileName)
   Local $aArray = FileReadToArray($sTextFileName)
   If @error Then
        MsgBox($MB_ICONERROR, "", "There was an error reading the file. @error: " & @error) ; An error occurred reading the current script file.
		Return
   EndIf
   return $aArray
EndFunc

Func zgFindExcelArray(Const ByRef $aArray,$vValue,$x=0,$y=0)
   Local $iIndex = 0, $iCol = 0
   For $iCol=0 To UBound($aArray,$UBOUND_COLUMNS )-1 Step 1
	  $iIndex = _ArraySearch($aArray, $vValue, 0, 0, 0, 3, 1, $iCol)
	  If $iIndex > 0 Then ExitLoop
   Next
   If $iIndex < 0 Then
	  MsgBox(0,"","��Excel�ļ��в�������["&$vValue&"]ʧ��","û���ҵ���Ҫ������")
	  SetError(1)
	  Return
   EndIf
   SetError(0)
   If $iIndex + $x < 0 Or $iCol + $y < 0 Or $iIndex + $x > UBound($aArray) Or $iCol + $y > UBound($aArray,$UBOUND_COLUMNS ) Then return ""
   return $aArray[$iIndex + $x][$iCol + $y]
EndFunc

Func zgFindExcelArrayMakeArea(Const ByRef $aArray,$vValue,$x=0,$y=0,$MakeCountry="",$Brand="DELL")
   Local $iRow = 0, $iCol = 0
   $iRow = _ArraySearch($aArray, $vValue,0,0,0,0,1)
   If @error Then
	  MsgBox(0,"","��Excel�ļ��в�������["&$vValue&"]ʧ��","û���ҵ���Ҫ������")
	  SetError(1)
	  Return
   EndIf
   $iCol = _ArraySearch($aArray, $vValue, 0, 0, 0, 0, 1, $iRow, True)

   Do
	  $iRow += 1
   Until (($MakeCountry = "" OR ($MakeCountry <> "" And _ArraySearch($aArray, $MakeCountry, 0, 0, 0, 3, 1, $iRow, True)>0)) _
	  And ($Brand = "" _
		 OR ($Brand = "DELL" And _ArraySearch($aArray, "Alienware", 0, 0, 0, 3, 1, $iRow, True) < 0) _
		 OR ($Brand <> "" And _ArraySearch($aArray, $Brand, 0, 0, 0, 3, 1, $iRow, True) > 0)))

   If $iRow + $x < 0 Or $iCol + $y < 0 Or $iRow + $x > UBound($aArray) Or $iCol + $y > UBound($aArray,$UBOUND_COLUMNS ) Then return ""

   return $aArray[$iRow + $x][$iCol + $y]
EndFunc

Func zgCheckMakeCountry(Const ByRef $aArray,ByRef $TotalQty,ByRef $TotalPrice,ByRef $TotalWet,$vValue,$x=0,$y=0,$ContractNo="",$MakeCountry="",$Brand="DELL")
   $TotalQty = 0
   $TotalPrice = 0
   $TotalWet = 0
   Local $iIndex = 0, $iCol = 0
   Local $count = 0
   For $iCol=0 To UBound($aArray,$UBOUND_COLUMNS )-1 Step 1
	  $iIndex = _ArraySearch($aArray, $vValue, 0, 0, 0, 3, 1, $iCol)
	  If $iIndex > 0 Then ExitLoop
   Next

   If $iIndex < 0 Then
	  MsgBox(0,"","��Excel�ļ��в�������["&$vValue&"]ʧ��","û���ҵ���Ҫ������")
	  SetError(0)
	  Return 0
   EndIf
   SetError(1)
   If $iIndex + $x < 0 Or $iCol + $y < 0 Or $iIndex + $x > UBound($aArray) Or $iCol + $y > UBound($aArray,$UBOUND_COLUMNS ) Then return 0

   Local $Serialno

   For $i = $iIndex + $x To UBound($aArray) - 1 Step 1
	  Local $aTemp = StringSplit($aArray[$i][$iCol + $y - 32]," ")
	  Local $sGetBrand = $aTemp[1]
;~ 	  MsgBox(0,$sGetBrand,$aArray[$i][$iCol + $y - 32])
	  If $Brand="DELL" Then
		 If $aArray[$i][$iCol + $y] = $MakeCountry And $sGetBrand <> "Alienware" Then
			$count = $count + 1

			$TotalQty = $TotalQty + $aArray[$i][$iCol + $y - 14]		;�����ۼ�
			If ($aArray[$i][$iCol + $y - 7] = 0) Then
			   WriteMemo($MB_ICONERROR,"���ִ��������","��"&$ContractNo&"���嵥�ļ���"&@CRLF&"ԭ����"&$MakeCountry&"�ĵ����ֵ�ڡ�"&$i&"����Ϊ0��",True)
			EndIf
			$TotalPrice = $TotalPrice + $aArray[$i][$iCol + $y - 7]	;��ֵ�ۼ�
			$Serialno = $aArray[$i][$iCol + $y - 22]
   ;~  		 MsgBox(0,$MakeCountry, $aArray[$i][$iCol + $y] & "|" &$aArray[$i][$iCol + $y - 14] & "|" & $aArray[$i][$iCol + $y - 7] & "|" & $aArray[$i][$iCol + $y - 22] & @CRLF & )
   ;~ 		 MsgBox(0,$MakeCountry,$Serialno)

			$alist = StringRegExp($Serialno,"(\d{12}) QTY (\d+)",$STR_REGEXPARRAYGLOBALMATCH)	;ȡ12λ���� �� ��Ӧ����
			For $j = 0 to UBound($alist) - 1 Step 2
   ;~ 			MsgBox(0,$Serialno,zgGetSerialAvgWet($alist[$j]))
			   $TotalWet = $TotalWet + zgGetSerialAvgWet($alist[$j]) * $alist[$j + 1]		;�����ۼ�
			Next
   ;~ 		 _ArrayDisplay($alist,$Serialno)
		 EndIf
	  Else
;~    MsgBox(0,$MakeCountry,"DELL[" & $sGetBrand & "]+++" & $count & "+++" & $TotalWet)
		 If $aArray[$i][$iCol + $y] = $MakeCountry And $sGetBrand = "Alienware" Then
			$count = $count + 1

			$TotalQty = $TotalQty + $aArray[$i][$iCol + $y - 14]		;�����ۼ�
			If ($aArray[$i][$iCol + $y - 7] = 0) Then
			   WriteMemo($MB_ICONERROR,"���ִ��������","��"&$ContractNo&"���嵥�ļ���"&@CRLF&"ԭ����"&$MakeCountry&"�ĵ����ֵ�ڡ�"&$i&"����Ϊ0��",True)
			EndIf
			$TotalPrice = $TotalPrice + $aArray[$i][$iCol + $y - 7]	;��ֵ�ۼ�
			$Serialno = $aArray[$i][$iCol + $y - 22]
   ;~  		 MsgBox(0,$MakeCountry, $aArray[$i][$iCol + $y] & "|" &$aArray[$i][$iCol + $y - 14] & "|" & $aArray[$i][$iCol + $y - 7] & "|" & $aArray[$i][$iCol + $y - 22] & @CRLF & )
   ;~ 		 MsgBox(0,$MakeCountry,$Serialno)

			$alist = StringRegExp($Serialno,"(\d{12}) QTY (\d+)",$STR_REGEXPARRAYGLOBALMATCH)	;ȡ12λ���� �� ��Ӧ����
			For $j = 0 to UBound($alist) - 1 Step 2
   ;~ 			MsgBox(0,$Serialno,zgGetSerialAvgWet($alist[$j]))
			   $TotalWet = $TotalWet + zgGetSerialAvgWet($alist[$j]) * $alist[$j + 1]		;�����ۼ�
			Next
   ;~ 		 _ArrayDisplay($alist,$Serialno)
		 EndIf
	  EndIf
   Next
   Return $count
EndFunc

Func zgGetSerialAvgWet($SerialNo)	;���ı��嵥�л�ȡ���ŵ�ƽ������
   Local $aTmp
   Local $wet = 0,$Num = 0
   For $i = 0 To UBound($aTextList) - 1 Step 1
	  If StringInStr($aTextList[$i],$SerialNo) Then
		 $aTmp = StringRegExp($aTextList[$i],"(" & $SerialNo & ")\D+(\d+)PCS\D+(\d+\.\d+)KG",$STR_REGEXPARRAYGLOBALMATCH)
		 $Num = $Num + $aTmp[1]
		 $wet = $wet + $aTmp[2]
	  EndIf
   Next

   If $wet = 0 Then
	  Return 0
   Else
	  Return $wet / $Num
   EndIf

EndFunc

Func zgFindArrayGetRowIndex(Const ByRef $aArray,$vValue)
   Local $iIndex = 0
   For $iCol=0 To UBound($aArray,$UBOUND_COLUMNS )-1 Step 1
	  $iIndex = _ArraySearch($aArray, $vValue, 0, 0, 0, 3, 1, $iCol)
	  If $iIndex > 0 Then ExitLoop
   Next
   If $iIndex < 0 Then
;~ 	  MsgBox(0,$ContractNO,"��Booking�ļ��в��Һ�ͬ��["&$vValue&"]ʧ��","û���ҵ���Ҫ������")
	  SetError(1)
	  Return
   EndIf
   SetError(0)
   return $iIndex
EndFunc

;20180123֮ǰ�İ汾��Ŀ�Ĺ����ı��ļ�β��
;~ Func zgFindTextlArray1($aArray,$ContractNO)
;~    Local $sTmp
;~    Local $aTmp
;~    If (StringLeft($ContractNO,3)="NPD") Then
;~ 	  $sTmp = $aArray[UBound($aArray)-5]
;~ 	  $aTmp = StringSplit($sTmp,":")
;~ 	  If $aTmp[1]="To" Then
;~ 		 Return $aTmp[2]
;~ 	  Else
;~ 		 Return "δ�ҵ��˵ֹ�"
;~ 	  EndIf
;~    Else
;~ 	  Return StringStripWS ($aArray[UBound($aArray)-5],3)
;~    EndIf
;~ EndFunc

;20180731�޸ģ�����ʱ����ļ�����ͬ����XN��ͷ����ʽ�б仯
Func zgFindTextlArray1(Const ByRef $aArray,$ContractNO)
   Local $sFlag
   If StringMid($ContractNO,2,2)="PD" Then
	  $sFlag = "To:"
   ElseIf (StringLeft($ContractNO,2)="XN") Then
	  $sFlag = "Ship To:"
   ElseIf (StringLeft($ContractNO,1)="L") Then
	  $sFlag = "Ship To:"
   Else
	  Return StringStripWS ($aArray[UBound($aArray)-7],3)
   EndIf
   Local $sTmp = "δ�ҵ��˵ֹ�"
   For $i = UBound($aArray)-1 To 0  Step -1
	  If StringLeft($aArray[$i],StringLen($sFlag)) = $sFlag Then
		 $sTmp = StringRight($aArray[$i],StringLen($aArray[$i]) - StringLen($sFlag))
	  EndIf
   Next
   Return $sTmp
EndFunc

Func zgFindTextlArray3(Const ByRef $aArray,$ContractNO)
   Local $i
   Local $Find = False
   For $i=0 To UBound($aArray)-1 Step 1
	  If StringInStr($aArray[$i],"Packing No") Or StringInStr($aArray[$i],"DELL ASN") Then
		 $Find = True
		 ExitLoop
	  EndIf
   Next

   If Not $Find Then
	  Return "δ�ҵ���ͬ��"
   EndIf

   Local $aTmp = StringSplit($aArray[$i],":")
   Return $aTmp[$aTmp[0]]
EndFunc

Func zgFindTextlArray4(Const ByRef $aArray,$ContractNO)
   Local $sTmp
   Local $iIndex
   $iIndex = _ArraySearch($aArray,_StringRepeat("-",100),0,0,0,0,0)
   If($iIndex<0) Then $iIndex = _ArraySearch($aArray,_StringRepeat("_",100),0,0,0,0,0)

   If($iIndex<0) Then
	  MsgBox(0,$ContractNO,"�嵥�ļ��н���������,��װ��ʧ�ܣ�����TXT�ļ���")
   Else
	  For $i=$iIndex + 1 To UBound($aArray)-1 Step 1
		If StringLen(StringStripWS($aArray[$i],$STR_STRIPLEADING + $STR_STRIPTRAILING)) > 10 Then
		   $sTmp = $aArray[$i]
		   ExitLoop
		EndIf
	  Next
   EndIf
   Return $sTmp
EndFunc

Func zgFindTextlArray6(Const ByRef $aArray,$ContractNO)
   Local $sTmp
   Local $iIndex
   $iIndex = _ArraySearch($aArray,_StringRepeat("-",100),0,0,0,0,0)
   $iIndex = _ArraySearch($aArray,_StringRepeat("-",100),0,$iIndex - 1,0,0,0)
   If($iIndex<0) Then
	  $iIndex = _ArraySearch($aArray,_StringRepeat("_",100),0,0,0,0,0)
	  $iIndex = _ArraySearch($aArray,_StringRepeat("_",100),0,$iIndex - 1,0,0,0)
   EndIf

   If($iIndex<0) Then
	  MsgBox(0,$ContractNO,"�嵥�ļ��н���������,ë��,���ء�ʧ�ܣ�����TXT�ļ���")
   Else
	  For $i=$iIndex + 1 To UBound($aArray)-1 Step 1
		If StringLen(StringStripWS($aArray[$i],$STR_STRIPLEADING + $STR_STRIPTRAILING)) > 10 Then
		   $sTmp = $aArray[$i]
		   ExitLoop
		EndIf
	  Next
;~	  $sTmp = $aArray[$iIndex-1]
;~	  If StringStripWS($sTmp,$STR_STRIPALL)="" Then $sTmp = $aArray[$iIndex-2]
   EndIf
   Return $sTmp
EndFunc

Func zgFindTextlArray7(Const ByRef $aArray,$ContractNO)
   Local $sTmp = "δ֪��װ����"
   Local $iIndex
   $iIndex = _ArraySearch($aArray,_StringRepeat("-",100),0,0,0,0,0)
   If($iIndex<0) Then $iIndex = _ArraySearch($aArray,_StringRepeat("_",100),0,0,0,0,0)

   Local $sFlag
   If StringMid($ContractNO,2,2)="PD" OR StringMid($ContractNO,6,2)="PD" Then	;NPD,PPD,CD2TNPD,CD2TPPD
	  $sFlag = "�\ݔ���b:"
	  For $i = UBound($aArray)-1 To 0  Step -1
		 If StringLeft($aArray[$i],StringLen($sFlag)) = $sFlag Then
			$sTmp = StringRight($aArray[$i],StringLen($aArray[$i]) - StringLen($sFlag))
		 EndIf
	  Next
	  $PkgType = ""
	  If StringInStr($sTmp,"�������b")>0 Then
		 $PkgType = "99"		;����
	  ElseIf StringInStr($sTmp,"��Ȼľ��")>0 Then
		 $PkgType = "93"		;��Ȼľ��
	  ElseIf StringInStr($sTmp,"���ƻ��w�S���ƺ�/��")>0 Then
		 $PkgType = "22"		;ֽ��
	  ElseIf StringInStr($sTmp,"��/��")>0 Then
		 $PkgType = "'06"		;��
	  EndIf

	  $sFlag = "�������b:"
	  For $i = UBound($aArray)-1 To 0  Step -1
		 If StringLeft($aArray[$i],StringLen($sFlag)) = $sFlag Then
			$sTmp = StringRight($aArray[$i],StringLen($aArray[$i]) - StringLen($sFlag))
		 EndIf
	  Next
	  $OtherPkgType = ""
	  If StringInStr($sTmp,"��/��")>0 Then
		 $OtherPkgType = $OtherPkgType & "/'06"		;��
	  EndIf
	  If StringInStr($sTmp,"���ƻ��w�S���ƺ�/��")>0 Then
		 $OtherPkgType = $OtherPkgType & "/22"		;ֽ��
	  EndIf
	  If StringInStr($sTmp,"��Ȼľ��")>0 Then
		 $OtherPkgType = $OtherPkgType & "/93"		;��Ȼľ��
	  EndIf
	  If StringInStr($sTmp,"�������b")>0 Then
		 $OtherPkgType = $OtherPkgType & "/99"		;����
	  EndIf
	  If StringLeft($OtherPkgType,1) = "/" Then
		 $OtherPkgType = StringRight($OtherPkgType,stringLen($OtherPkgType) - 1)
	  EndIf

   ElseIf (StringLeft($ContractNO,2)="XN") Or (StringLeft($ContractNO,1)="L") Or (StringLeft($ContractNO,6)="CD2TXN") Or (StringLeft($ContractNO,5)="CD2TL") Then
	  $PkgType = TransPkgType(zgGetPkgNumType($aArray[$iIndex]))
	  $OtherPkgType = "22"
   Else
	  Local $i=0
	  Local $Find=False
	  For $i = 0 To UBound($aArray) - 1 Step 1
		 If StringInStr($aArray[$i],"Say Total") Then
			$Find = True
			ExitLoop
		 EndIf
	  Next

	  If Not $Find Then
		 Return "δ֪��װ����"
	  EndIf

	  $sTmp = StringStripWS ($aArray[$i],3)
	  $sTmp = zgGetPkgNumType($sTmp)
	  $PkgType = TransPkgType($sTmp)
	  $OtherPkgType = ""
   EndIf
EndFunc

