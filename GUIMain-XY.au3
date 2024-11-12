#pragma compile(Out,兴亚空运报关数据自动录入(金二版)2024.exe)
#pragma compile(FileDescription, 兴亚空运报关数据处理系统 - 数据自动识别及录入)
#pragma compile(ProductName, 兴亚空运报关数据自动录入(2022分原产地版本))
#pragma compile(Icon, InputData.ico)
#pragma compile(FileVersion, 3.24.11.10 3.99.99.99)
#include-once
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiEdit.au3>
#include <EditConstants.au3>
#include <Array.au3>
#include <Date.au3>
#include <File.au3>
#include <ScrollBarsConstants.au3>

#include "ExcelFunc.au3"
#include "StringFunc.au3"
#include "StringConstants.au3"
#include "Authorization.au3"
#include "MakeSingleWindowXML_V5.1-XY.au3"
#include "SqlitFunc.au3"

Local $version="3.24.11.10"
;检查授权
If Not zgAuthorization() Then
;~    MsgBox($MB_ICONERROR,"未授权使用","请联系开发人员！")
;~    Exit
EndIf
HotKeySet("{ESC}", "Terminate")

Global $OutFileName,$LogFileName,$DBFileName,$IniFileName,$ZipFilePath
Global $DECInputFileName,$INVInputFileName,$INVApplyFileName
Global $bMakeXML=False
Global $sCustomMemo
Global $InputPath
Global $SerialNo,$sFullPathFileName

;~ Global $OutFile = FileOpen($OutFileName, 1)
Global $aInputInfo[1][40],$aDecList[0][20]
Global $aExcelBooking[0][0],$aExcelFinal[0][0],$aExcelExpci[0][0],$aTextList[0][0]
Global $PdfFinal,$PdfExpci,$PdfText
Global $ContractNO
Global $pkgNum,$pkgType,$OtherPkgType
Global $doing,$DeclistIndex=0

Global $Test,$Save,$AutoDec,$ReciveTime


Start()

Func Start()

    Local $hGUI = GUICreate("兴亚空运报关数据整理录入",620,600)

    ; Set Margins
;~     _GUICtrlEdit_SetMargins($g_idMemo, BitOR($EC_LEFTMARGIN, $EC_RIGHTMARGIN), 20, 20)

    Global $idBtnSelectBooking = GUICtrlCreateButton("Booking", 50, 10, 55, 30)
    Global $idBtnBookDirZip = GUICtrlCreateButton("ZIP", 105, 10, 24, 30)
    Global $g_idBookingData = GUICtrlCreateEdit("", 140, 10, 410, 40)
    Global $idBtnClearBooking = GUICtrlCreateButton("X", 550, 10, 20, 40)
    Global $idBtnSelectDir = GUICtrlCreateButton("数据目录", 50, 60, 70, 30)
    Global $g_idSelectDir = GUICtrlCreateEdit("", 140, 60, 350, 40)
    Global $idBtnSortDate = GUICtrlCreateButton("数据整理", 500, 60, 70, 30)
    Global $idBtnSelectFile = GUICtrlCreateButton("中介文件", 50, 110, 70, 30)
    Global $g_idOutputData = GUICtrlCreateEdit("", 140, 110, 330, 40)
    Global $idBtnClearOut = GUICtrlCreateButton("X", 470, 110, 20, 40)
;~     Global $idBtnBaoGuan = GUICtrlCreateButton("直航录入", 150, 150, 100, 30)
;~     Global $idBtnYiTiHua = GUICtrlCreateButton("一体化录入", 300, 150, 100, 30)
;~     Global $idBtnZhuanGuan = GUICtrlCreateButton("转关录入", 450, 150, 100, 30)
    Global $idBtnSaveOutput = GUICtrlCreateButton("保存", 500, 110, 70, 30)
    Global $idBtnTEST = GUICtrlCreateButton("收取回执", 50, 160, 70, 30)
    Global $idBtnINVListSave = GUICtrlCreateButton("核注[暂存]", 162, 160, 70, 30)
    Global $idBtnINVListApply = GUICtrlCreateButton("核注[申报]", 275, 160, 70, 30)
    Global $idBtnSingleWindow = GUICtrlCreateButton("报关[暂存]", 387, 160, 70, 30)
    Global $idBtnCleanData = GUICtrlCreateButton("清除过期", 500, 160, 70, 30)
;~     Global $idBtnClose = GUICtrlCreateButton("关闭", 480, 150, 70, 30)
	Global $idProgressbar = GUICtrlCreateProgress ( 50, 200, 500, 25); [, height [, style = -1 [, exStyle = -1]]]] )
    Global $idBtnClearLog = GUICtrlCreateButton("X", 550, 200, 20, 25)
    Global $g_idMemo = GUICtrlCreateEdit("", 50, 230, 520, 340)
    _GUICtrlEdit_SetReadOnly($g_idOutputData, True)
    _GUICtrlEdit_SetReadOnly($g_idSelectDir, True)
    _GUICtrlEdit_SetReadOnly($g_idBookingData, True)
    _GUICtrlEdit_SetReadOnly($g_idMemo, True)
;~     _GUICtrlEdit_SetReadOnly($idBtnINVListApply, True)

;~ 	_GUICtrlEdit_AppendText($g_idMemo,StringFormat("%s  %s" & @CRLF,_Date_Time_SystemTimeToDateTimeStr(_Date_Time_GetLocalTime(),1),"和越数据整理录入"))
    GUICtrlSetFont($g_idMemo, 9, 400, 0, "Courier New")
	WriteMemo(0,"兴亚空运报关数据整理录入v"&$version&"","",False)
    ; Display the GUI.
    GUISetState(@SW_SHOW, $hGUI)
	WriteMemo(0,"读取INI配置文件中，请稍等......","",False)
    zgSetConfigFileName()
	WriteMemo(0,"读取INI配置文件中完成......","",False)

    GetConfigData()
	WriteMemo(0,"系统准备完成！","",False)
    ; Create a GUI with various controls.

   ;~ 每5分钟检查一次收件箱  AdlibRegister("MyAdLibFunc",1000*60)
;~     AdlibRegister("zgCheckResponse",1000*300)
    Global $Timer = DllCallbackRegister ( "Timer" , "int" , "hwnd;uint;uint;dword" ) ; 创建自定义函数 Timer 的回调， API 定时器函数 SetTimer 需要
    Global $TimerDLL = DllCall ( "user32.dll" , "uint" , "SetTimer" , "hwnd" , 0, "uint" , 0, "int" , 1000*$ReciveTime , "ptr" , DllCallbackGetPtr ($Timer)) ;1000 毫秒执行一次

    ; Loop until the user exits.
    While 1

	    zgSetConfigFileName()

        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                ExitLoop
			 Case $idBtnSelectBooking
				zgSelectBooking()
			 Case $idBtnClearBooking
				_GUICtrlEdit_SetText ( $g_idBookingData, "" )
			 Case $idBtnClearOut
				_GUICtrlEdit_SetText ( $g_idOutputData, "" )
			 Case $idBtnClearLog
				_GUICtrlEdit_SetText ( $g_idMemo, "" )
			Case $idBtnSelectFile
				zgSelectFile()
			Case $idBtnBookDirZip
				zgBookDirZip()
			 Case $idBtnSelectDir
				zgSelectDir()
			Case $idBtnSortDate
				zgOutputData()
			Case $idBtnSaveOutput
			   GUICtrlSetData($g_idOutputData,$OutFileName)
			   zgSaveOutputData($OutFileName,$aInputInfo,$aDecList,True)
			Case $idBtnSingleWindow
				zgSingleWindow()
			Case $idBtnINVListSave
				zgINVSave()
			Case $idBtnINVListApply
				zgINVApply()
			Case $idBtnTEST
			   zgCheckResponse()
			Case $idBtnCleanData
			   zgCleanOldData()
        EndSwitch
    WEnd

    DllCallbackFree ($Timer) ; 关闭回调

    GUIDelete($hGUI)
 EndFunc   ;==>Start end

Func zgSetConfigFilename()
   If StringRight(@ScriptDir,1)='\' Then
	  $OutFileName = @ScriptDir & "Output\"
	  $DECInputFileName = @ScriptDir & "Input\"
	  $INVInputFileName = @ScriptDir & "Input\"
	  $INVApplyFileName = @ScriptDir & "Input\"
	  $LogFileName = @ScriptDir & "Log\"
	  $ZipFilePath = @ScriptDir & "Output\Zip\"
   Else
	  $OutFileName = @ScriptDir & "\Output\"
	  $DECInputFileName = @ScriptDir & "\Input\"
	  $INVInputFileName = @ScriptDir & "\Input\"
	  $INVApplyFileName = @ScriptDir & "\Input\"
	  $LogFileName = @ScriptDir & "\Log\"
	  $ZipFilePath = @ScriptDir & "\Output\Zip\"
   EndIf
   If NOT FileExists($OutFileName) Then DirCreate($OutFileName)
   If NOT FileExists($DECInputFileName) Then DirCreate($DECInputFileName)
   If NOT FileExists($LogFileName) Then DirCreate($LogFileName)
   If NOT FileExists($ZipFilePath) Then DirCreate($ZipFilePath)

   $OutFileName=$OutFileName&@YEAR&@MON&@MDAY&".xlsx"
   $LogFileName=$LogFileName&@YEAR&@MON&@MDAY&".Log"
   $DECInputFileName=$DECInputFileName&"报关单暂存"&@YEAR&@MON&@MDAY&".input"
   $INVInputFileName=$INVInputFileName&"核注单暂存"&@YEAR&@MON&@MDAY&".input"
   $INVApplyFileName=$INVApplyFileName&"核注单申报"&@YEAR&@MON&@MDAY&".input"
   $DBFileName = @ScriptDir & "\OutPut\SingleWindow.db"
   $IniFileName=@ScriptDir & "\QPDataInput.INI"

   $Test=IniRead ( $IniFileName, "SYSTEM", "TEST", "Real")
   $Save=IniRead ( $IniFileName, "SYSTEM", "SAVE", "True")
   $AutoDec=IniRead ( $IniFileName, "SYSTEM", "AutoDec", "True")
   $ReciveTime=IniRead ( $IniFileName, "SYSTEM", "ReciveTime", "3600")
   $sCustomMemo=IniRead ( $IniFileName, "SYSTEM", "CustomMemo", "")
EndFunc

Func zgCheckInformation($ContractNO)
   Local $sTmp1,$sTmp2,$sTmp3
   Local $aInput[1][40]
   Local $aSignature[12][3]=[ _
   ["",					"",							"TO:"], _					;1.2.10.运抵国
   ["CONTRACT NO: ",	"Invoice No.",				"PACKING NO.    :  "], _	;3.合同号
   ["",					"",							"SAY TOTAL "], _			;4.5.件数包装
   ["Quantity",			"",							"TOTAL"], _					;6.7.9.数量毛重净重
   ["Specifications",	"",							""], _						;8.规格
   ["Total Value",		"Invoice Total",			""], _						;11.总额
   ["",					"Platform Details",			""], _						;12-1  12-4
   ["",					"Engineering Description",	""], _						;12-2
   ["",					"Trade Description",		""], _						;品名
   ["BUYER：",			"",							""], _						;境外收货人
   ["",					"",							"PACKING TYPE:"], _			;包装材质
   ["",					"Total Value",				""]]						;原产国

   TrayTip ( @ScriptName, $doing & "/" & UBound($aExcelBooking) & "【" & $ContractNO & "】", 1 )
   WriteMemo(0,$doing&"/"&UBound($aExcelBooking) &"【" & $ContractNO & "】","数据处理开始...",False)
;~    检查是否有多品名，如不一致，提示后跳过此合同号
   $i=1
   $sTmp1 = zgFindExcelArray($aExcelExpci,$aSignature[8][1],$i,1)
   $TradeDestription = $sTmp1
;~    MsgBox(0,"品名" & $i,$sTmp1)
   While $sTmp1 <> ""
	  $i += 1
	  $sTmp1 = zgFindExcelArray($aExcelExpci,$aSignature[8][1],$i,1)
;~    MsgBox(0,"品名" & $i,$sTmp1)
	  If $sTmp1 <>"" AND $sTmp1 <> $TradeDestription Then
		 Return SetError(3,10,"Expci文件中有多个品名，且不一致！")
	  EndIf
   WEnd

   ;1.2.10.运抵国
   $sTmp1 = zgFindTextlArray1($aTextList,$ContractNO)
   $sTmp2 = TransTargetAdd($sTmp1)
   If @error Then
	  WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】","运抵国解析失败 清单文件中国家代码【"&$sTmp1&"】解析失败！",True)
   ;针对仁宝XN的文件格式有变化，以下代码做修改
	  $aInput[0][0]=""
   Else
	  $aInput[0][0]=$sTmp2
   EndIf
   If $aInput[0][0] = "CHN" Then
	  WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】","运抵国不应该是中国【CHN】",True)
   EndIf
   ;修改完成20180123

   ;3.合同号
   $sTmp1 = zgFindExcelArray($aExcelFinal,$aSignature[1][0],0,1)
   $sTmp2 = zgFindExcelArray($aExcelExpci,$aSignature[1][1],0,7)
   $sTmp3 = zgFindTextlArray3($aTextList,$ContractNO)
   $sTmp1 = StringStripWS($sTmp1,3)
   $sTmp2 = StringStripWS($sTmp2,3)
   $sTmp3 = StringStripWS($sTmp3,3)
   If (($sTmp1 = $sTmp2) AND ($sTmp2 = $sTmp3)) Then
	  $aInput[0][1]=$sTmp3
   Else
	  WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】", StringFormat("合同号不匹配\r\nContract文件中合同号【%s】\r\nExpci文件中合同号【%s】\r\nFinal文件中合同号【%s】",$sTmp1,$sTmp2,$sTmp3),True)
	  WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】", "关键数据校验失败，跳过此合同号，请人工检查合同及清单文件！",True)
	  Return
   EndIf
   $aInput[0][1]=$sTmp1
   ;修改完成20180123

   ;4.5.件数,包装
   $pkgNum=0
   $pkgType=""
   $OtherPkgType=""
   $sTmp1 = zgFindTextlArray4($aTextList,$ContractNO)
   zgGetPkgNumType($sTmp1)
   $aInput[0][2]=$pkgNum
   zgFindTextlArray7($aTextList,$ContractNO)
   $aInput[0][3]=$pkgType	;包装材质
   $aInput[0][28]=$OtherPkgType	;其他包装材质

   ;6.7.9.数量,毛重,净重
   $sTmp1 = zgFindTextlArray6($aTextList,$ContractNO)
   $sTmp1 = StringReplace($sTmp1,",","")
   $aTemp = StringRegExp($sTmp1,"[\d\.]+",$STR_REGEXPARRAYGLOBALMATCH )
   If UBound($aTemp)=4 Then
	  $aInput[0][4]=$aTemp[0]
	  $aInput[0][5]=Round($aTemp[2])
	  $aInput[0][6]=$aTemp[1]
   ElseIf UBound($aTemp)=5 Then
	  $aInput[0][4]=$aTemp[1]
	  $aInput[0][5]=Round($aTemp[3])
	  $aInput[0][6]=$aTemp[2]
   ElseIf UBound($aTemp)=6 Then
	  $aInput[0][4]=$aTemp[2]
	  $aInput[0][5]=Round($aTemp[4])
	  $aInput[0][6]=$aTemp[3]
   EndIf
   if $aInput[0][5]<1 Then $aInput[0][5]=1
   if $aInput[0][6]<1 Then $aInput[0][6]=1
   If $aInput[0][5] <= $aInput[0][6] Then
	  WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】", StringFormat("获取到的毛重【%s】小于净重【%s】",$aInput[0][5],$aInput[0][6]),True)
   EndIf

   ;合同中的数量
   $sTmp1 = ""
   Local $i=1
   Local $j=0
   While $sTmp1 = ""
	  for $j=0 To -4 Step -1
		 $sTmp1 = zgFindExcelArray($aExcelFinal,$aSignature[3][0],$i,$j)
		 If $sTmp1 <> "" Then
			ExitLoop
		 EndIf
	  Next
 	  $i+=1
   WEnd
   If Not IsNumber($sTmp1*1) Then
	  WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】","从发票文件中获取数量失败",True)
	  Return
   EndIf
   $ContractQty = $sTmp1*1
;~    MsgBox(0,"Qty On Contract","["&$ContractNo&"]" & $ContractQty)

   ;8.规格
   $sTmp1 = ""
   Local $i=1
   Local $j=0
   While $sTmp1 = ""
	  for $j=0 To -4 Step -1
		 $sTmp1 = zgFindExcelArray($aExcelFinal,$aSignature[4][0],$i,$j)
		 If $sTmp1 <> "" Then
			ExitLoop
		 EndIf
	  Next
 	  $i+=1
   WEnd
   $sTmp1=StringRight(StringStripWS($sTmp1,3),2)
   If Not IsNumber($sTmp1*1) Then
	  WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】","从发票文件中获取规格失败",True)
	  Return
   EndIf
   $aInput[0][7] = $sTmp1*1

   ;11.总额
   $sTmp1 = ""
   Local $i=1
   Local $j=0
   While $sTmp1 = ""
	  for $j=0 To -4 Step -1
		 $sTmp1 = zgFindExcelArray($aExcelFinal,$aSignature[5][0],$i,$j)
		 If $sTmp1 <> "" Then
			ExitLoop
		 EndIf
	  Next
 	  $i+=1
   WEnd
;~    $sTmp1=StringRight(StringStripWS($sTmp1,3),2)
   If Not IsNumber($sTmp1*1) Then
	  WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】","从发票文件中获取总额失败",True)
	  Return
   EndIf
   $sTmp2 = zgFindExcelArray($aExcelExpci,$aSignature[5][1],0,10)
   Local $AmountFinal = Round($sTmp1+0.00001,2)
   Local $AmountExpci = Round($sTmp2+0.00001,2)
   $aInput[0][8]=$AmountExpci
   If $AmountFinal = $AmountExpci Then
      ;$aInput[0][8]=$sTmp1
   Else
	  WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】",StringFormat("总额不匹配, FINAL文件中总额【%s】\r\nExpci文件中总额【%s】",$AmountFinal,$AmountExpci),True)
	  ;Return
   EndIf

   Local $index=zgFindArrayGetRowIndex($aExcelBooking,$ContractNO)
   If $aExcelBooking[$index][8]="" Then	;成都CTU或者空白，直航
	  $aExcelBooking[$index][8]="CTU"
   EndIf

   $aInput[0][13]=$aExcelBooking[$index][8]	;口岸
   $aInput[0][14]=$aExcelBooking[$index][7]	;运输方式

   $aInput[0][15]=$aExcelBooking[$index][4]	;提单号-主单
   $aInput[0][16]=$aExcelBooking[$index][2]	;提单号-分单
;~    Local $sOutput = StringStripWS(StringRegExpReplace($aExcelBooking[$index][0], '(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})', ' $1/$2/$3'),3)
;~    Local $sNewDate = _DateAdd("D",GetDelayDay($aInput[0][13]),_DateTimeFormat($sOutput,2))
   Local $sNewDate=""
   If $aInput[0][14]="4" Then;公路转关 +7D
	  $sNewDate=_DateAdd("D",15,@YEAR&"/"&@MON&"/"&@MDAY)
   ElseIf $aInput[0][14]="5" Then;航空转关 +15D
	  $sNewDate=_DateAdd("D",7,@YEAR&"/"&@MON&"/"&@MDAY)
   EndIf
   If $aInput[0][13]="CTU" Then
	  $sNewDate=@YEAR&"/"&@MON&"/"&@MDAY
   EndIf
   $aInput[0][17]=$sNewDate	;到达日期

   If $aExcelBooking[$index][8]="CTU" Then
	  $aInput[0][18]=StringReplace($aExcelBooking[$index][5]," ","")	;车牌号
   Else
	  $aTmp = StringSplit($aExcelBooking[$index][5]," ")
	  Switch $aTmp[0]
		 Case 1
			$aInput[0][18]=$aTmp[1]	;车牌号
		 case 2
			$aInput[0][18]=$aTmp[1]	;车牌号
			$aInput[0][20]=$aTmp[2]	;白卡号
		 case 3
			$aInput[0][18]=$aTmp[1]	;车牌号
			$aInput[0][19]=$aTmp[2]	;集装箱号
			$aInput[0][20]=$aTmp[3]	;白卡号
	  EndSwitch
   EndIf
   $aInput[0][21]=$aExcelBooking[$index][10]	;件数
   If $aInput[0][21]<>$aInput[0][2] Then
	  WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】","Booking文件中【件数】"&$aInput[0][21]&"与清单文件"&round($aInput[0][2],0)&"不一致！",True)
	  $aInput[0][21]=$aInput[0][2]
   EndIf

;~    $aInput[0][22]=$aExcelBooking[$index][11]	;包装类型
;~    $aInput[0][22]="92"	;	$aExcelBooking[$index][11]	;包装类型
   $aInput[0][23]=$aExcelBooking[$index][13]	;数量
   If $aInput[0][23]<>$aInput[0][4] Then
	  WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】","Booking文件中【数量】"& $aInput[0][23]&"与清单文件"&round($aInput[0][4],0)&"不一致！",True)
	  $aInput[0][23]=$aInput[0][4]
   EndIf
   If $ContractQty<>$aInput[0][4] Then
	  WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】","合同文件中【数量】"& $ContractQty &"与清单文件"&round($aInput[0][4],0)&"不一致！",True)
   EndIf
   $aInput[0][24]=$aExcelBooking[$index][12]	;毛重（四舍五入）
   If $aInput[0][24]<>$aInput[0][5] Then
	  WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】","Booking文件中【毛重】"&$aInput[0][24]&"与清单文件"&round($aInput[0][5],0)&"不一致！",True)
	  $aInput[0][24]=round($aInput[0][5],0)
   EndIf

;~ MsgBox(0,"境外收货人",zgFindExcelArray($aExcelFinal,$aSignature[9][0],0,1))
   $aInput[0][25] = zgFindExcelArray($aExcelFinal,$aSignature[9][0],0,1)	;境外收货人
   $aInput[0][26] = $aExcelBooking[$index][6]
   $aInput[0][27] = GetOverseasConsigneeEname($aExcelBooking[$index][6])
;~    $aInput[0][28] = zgFindTextlArray7($aTextList,$ContractNO)	;包装材质

;~    $sTmp1 = zgFindExcelArray($aExcelExpci,$aSignature[11][1],1,-1)	;原产国
;~ /*** 2022.02.17 修改 Sheet2需要分原产国，并需要保存到新的sheet中
;~    $sTmp1 = zgCheckMakeCountry($aExcelExpci,$aSignature[11][1],1,-1)	;原产国
;~    If ($sTmp1="ERROR") Then
;~ 	  WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】","在合同文件中发现原产地不一致的数据项",True)
;~    ElseIf $sTmp1 = "CN" Then
;~ 	  $aInput[0][33] = "142"
;~    ElseIf $sTmp1 = "VN" Then
;~ 	  $aInput[0][33] = "141"
;~    ElseIf $sTmp1 = "TWO" Then
;~ 	  WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】","原产国不唯一！请人工进行检查处理",True)
;~    EndIf
;~ *****************************************************
   Local $aTmp[1][20]
   Local $ret = 0,$TotalWet = 0
   Local $aMakeCountry = ["CN","VN","TW"]
   For $Country in $aMakeCountry
	  Local $Qty=0,$TotalValue=0,$MakeCountryWet=0
	  ;先统计Alienware品牌的数据
	  $ret = zgCheckMakeCountry($aExcelExpci,$Qty,$TotalValue,$MakeCountryWet,$aSignature[11][1],1,7,$ContractNo,$Country,"Alienware")	;返回原产国商品条数，在参数中带回数量，价值，重量
	  If $ret > 0 Then
		 $aTmp[0][0] = $ContractNO
		 $aTmp[0][1] = $Country
		 $aInput[0][9] = $Country
		 $aTmp[0][2] = $Qty
		 $aTmp[0][3] = $TotalValue
		 If (StringLeft($ContractNO,2)="XN") Or (StringLeft($ContractNO,1)="L") Then
			$aTmp[0][4] = $aInput[0][6] / $aInput[0][4] * $Qty
		 Else
			$aTmp[0][4] = $MakeCountryWet
		 EndIf
		 $aTmp[0][4] = Round($aTmp[0][4],3)
		 $TotalWet = $TotalWet + $aTmp[0][4]

		 ;12-1   12-4
		 $sTmp2 = zgFindExcelArrayMakeArea($aExcelExpci,$aSignature[6][1],0,0,$Country,"Alienware")

		 $aTemp = StringSplit($sTmp2," ")
		 $aTmp[0][5]=$aTemp[1]
		 $aTmp[0][6]=StringRight($sTmp2,StringLen($sTmp2)-StringLen($aTmp[0][5])-1)
		 $aTmp[0][9]="Alienware"

		 ;12-2
		 $sTmp2 = zgFindExcelArrayMakeArea($aExcelExpci,$aSignature[7][1],0,0,$Country,"Alienware")
		 $aTemp = StringSplit($sTmp2,",OS:",$STR_ENTIRESPLIT )
		 If ($aTemp[0]=1) Then
			$aTemp = StringRegExp($sTmp2, '(.*)\n\d{12}', $STR_REGEXPARRAYGLOBALMATCH )
			If IsArray($aTemp) Then
			   $aTmp[0][7]=$aTemp[0]
			Else
			   WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】",StringFormat("规格型号解析失败，请手工处理这个合同号",$sTmp2),True)
			EndIf
		 Else
			$aTmp[0][7]=$aTemp[1]
		 EndIf

		 If $aInput[0][7]=12 Then
			$aTmp[0][8]="Thin"
		 Elseif $aInput[0][7]=3 Then
			$aTmp[0][8]="Win 11"
		 Else;if $aInput[0][7]=1 OR $aInput[0][7]=6  OR $aInput[0][7]=10 Then
			$aTemp = StringRegExp($sTmp2,",OS:(\D+ \d+)",$STR_REGEXPARRAYGLOBALMATCH)
		    If @error Then
			  $aTmp[0][8] = "操作系统:无操作系统"
		    Else
			  $aTmp[0][8] = "操作系统:" & $aTemp[0]
		    EndIf
		 EndIf

		 If $ret > 1 Then
			$aTmp[0][5] &= "等"
			$aTmp[0][6] &= "等"
			$aTmp[0][7] &= "等"
			$aTmp[0][8] &= "等"
		 EndIf
		 _ArrayAdd($aDecList,$aTmp)
;~ 	  Else
;~ 		 WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】",StringFormat("处理原产国分类数量价值重量失败，请手工处理这个合同号"),True)
	  EndIf

;~ 	  For $i=0 To UBound($aTmp) - 1 Step 1
;~ 		 $aTmp[0][$i]=""
;~ 	  Next
	  ;下面统计DELL品牌电脑的数据
	  $ret = zgCheckMakeCountry($aExcelExpci,$Qty,$TotalValue,$MakeCountryWet,$aSignature[11][1],1,7,$ContractNo,$Country,"DELL")	;返回原产国商品条数，在参数中带回数量，价值，重量
	  If $ret > 0 Then
		 $aTmp[0][0] = $ContractNO
		 $aTmp[0][1] = $Country
		 $aInput[0][9] = $Country
		 $aTmp[0][2] = $Qty
		 $aTmp[0][3] = $TotalValue
		 If (StringLeft($ContractNO,2)="XN") Or (StringLeft($ContractNO,1)="L") Then
			$aTmp[0][4] = $aInput[0][6] / $aInput[0][4] * $Qty
		 Else
			$aTmp[0][4] = $MakeCountryWet
		 EndIf
		 $aTmp[0][4] = Round($aTmp[0][4],3)
		 $TotalWet = $TotalWet + $aTmp[0][4]

		 ;12-1   12-4
		 $sTmp2 = zgFindExcelArrayMakeArea($aExcelExpci,$aSignature[6][1],0,0,$Country,"DELL")
		 $aTemp = StringSplit($sTmp2," ")
		 $aTmp[0][5]=$aTemp[1]
		 $aTmp[0][6]=StringRight($sTmp2,StringLen($sTmp2)-StringLen($aTmp[0][5])-1)
		 $aTmp[0][9]="DELL"

		 ;12-2
		 $sTmp2 = zgFindExcelArrayMakeArea($aExcelExpci,$aSignature[7][1],0,0,$Country,"DELL")
		 $aTemp = StringSplit($sTmp2,",OS:",$STR_ENTIRESPLIT )
		 If ($aTemp[0]=1) Then
			$aTemp = StringRegExp($sTmp2, '(.*)\n\d{12}', $STR_REGEXPARRAYGLOBALMATCH )
			If IsArray($aTemp) Then
			   $aTmp[0][7]=$aTemp[0]
			Else
			   WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】",StringFormat("规格型号解析失败，请手工处理这个合同号",$sTmp2),True)
			EndIf
		 Else
			$aTmp[0][7]=$aTemp[1]
		 EndIf

		 If $aInput[0][7]=12 Then
			$aTmp[0][8]="Thin"
		 Elseif $aInput[0][7]=3 Then
			$aTmp[0][8]="Win 11"
		 Else;if $aInput[0][7]=1 OR $aInput[0][7]=6  OR $aInput[0][7]=10 Then
			$aTemp = StringRegExp($sTmp2,",OS:(\D+ \d+)",$STR_REGEXPARRAYGLOBALMATCH)
		    If @error Then
			  $aTmp[0][8] = "操作系统:无操作系统"
		    Else
			  $aTmp[0][8] = "操作系统:" & $aTemp[0]
		    EndIf
		 EndIf

		 If $ret > 1 Then
			$aTmp[0][5] &= "等"
			$aTmp[0][6] &= "等"
			$aTmp[0][7] &= "等"
			$aTmp[0][8] &= "等"
		 EndIf
		 _ArrayAdd($aDecList,$aTmp)
;~ 	  Else
;~ 		 WriteMemo($MB_ICONWARNING,$doing & "/" & UBound($aExcelBooking) & "【"&$ContractNO&"】",StringFormat("处理原产国分类数量价值重量失败，请手工处理这个合同号"),True)
	  EndIf

;~ 	  For $i=0 To UBound($aTmp) - 1 Step 1
;~ 		 $aTmp[0][$i]=""
;~ 	  Next
   Next
;~    _ArrayDisplay($aDecList)

   ;处理4舍5入带来的误差
   If Round($TotalWet,3) - Round($aInput[0][6],3) <> 0 Then
	  Local $aResult = _ArrayFindAll($aDecList,$ContractNO)
	  Local $msg = ""
	  For $i = 0 To UBound($aResult) - 1 Step 1
		 $msg &= " " & $aDecList[$aResult[$i]][1] & ":" & $aDecList[$aResult[$i]][4]
	  Next

	  WriteMemo(0,$doing&"/"&UBound($aExcelBooking) &"【" & $ContractNO & "】","净重合计值出现偏差，" _
		 & "Booking:[" & $aInput[0][6] & "] 分项合计:[" & $TotalWet & "]" & $msg,False)
	  $aDecList[UBound($aDecList)-1][4] = $aDecList[UBound($aDecList)-1][4] + $aInput[0][6] - $TotalWet
	  WriteMemo(0,$doing&"/"&UBound($aExcelBooking) &"【" & $ContractNO & "】","净重分项" _
		 & $aDecList[UBound($aDecList)-1][1] & "被修改为:"& $aDecList[UBound($aDecList)-1][4],False)
   EndIf

   _ArrayAdd($aInputInfo,$aInput)
;~  2022.01.17 修改结束*************************/

   WriteMemo(0,$doing&"/"&UBound($aExcelBooking) &"【" & $ContractNO & "】","数据处理完成！！！",False)
EndFunc

Func zgSelectBooking()
;~    $InputFile = FileOpenDialog ( "选择已经准备好的数据文件", "", "Excel文件 (*.xls)|所有文件(*.*")
   ; Create a constant variable in Local scope of the message to display in FileOpenDialog.

    ; Display an open dialog to select a list of file(s).
    Local $sFileOpenDialog = FileOpenDialog("选择Booking数据文件", "", "Excel文件 (*.xls;*.xlsx)|所有文件(*.*)", $FD_FILEMUSTEXIST + $FD_MULTISELECT)
    If @error Then
        ; Display the error message.
;~         MsgBox($MB_SYSTEMMODAL, "", "No file(s) were selected.")

        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
;~         FileChangeDir(@ScriptDir)
	  Return
    EndIf
        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
;~         FileChangeDir(@ScriptDir)

        ; Replace instances of "|" with @CRLF in the string returned by FileOpenDialog.
;~         $sFileOpenDialog = StringReplace($sFileOpenDialog, "|", @CRLF)

        ; Display the list of selected files.
	  GUICtrlSetData($g_idBookingData,$sFileOpenDialog)
EndFunc

Func zgSelectFile()
;~    $InputFile = FileOpenDialog ( "选择已经准备好的数据文件", "", "Excel文件 (*.xls)|所有文件(*.*")
   ; Create a constant variable in Local scope of the message to display in FileOpenDialog.

    ; Display an open dialog to select a list of file(s).
    Local $sFileOpenDialog = FileOpenDialog("选择已经准备好的数据文件", "", "Excel文件 (*.xls;*.xlsx)|所有文件(*.*)", $FD_FILEMUSTEXIST + $FD_MULTISELECT)
    If @error Then
        ; Display the error message.
;~         MsgBox($MB_SYSTEMMODAL, "", "No file(s) were selected.")
; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
;~         FileChangeDir(@ScriptDir)
	  Return
    EndIf
        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
;~         FileChangeDir(@ScriptDir)

        ; Replace instances of "|" with @CRLF in the string returned by FileOpenDialog.
;~         $sFileOpenDialog = StringReplace($sFileOpenDialog, "|", @CRLF)

        ; Display the list of selected files.
	  GUICtrlSetData($g_idOutputData,$sFileOpenDialog)
EndFunc

Func zgSelectDir()	;选择数据目录
   $InputPath = FileSelectFolder ( "选择数据目录", "" )

   If ($InputPath <> "") Then GUICtrlSetData($g_idSelectDir,$InputPath)
   If _GUICtrlEdit_GetText($g_idSelectDir) = "" Then
	  MsgBox($MB_ICONWARNING,"提示","请选择源数据文件目录！")
	  Return
   EndIf

EndFunc

Func zgOutputData()	;数据整理

   _ArrayDelete($aInputInfo, "0-" & UBound($aInputInfo)-1)
   _ArrayDelete($aDecList, "0-" & UBound($aDecList)-1)

   GUICtrlSetData($idProgressbar, 0)
   $doing=0

   WriteMemo(0,"数据收集整理开始。。。","",False)
   WriteMemo(0,"读取Booking文件","",False)

    Local $oExcel = _Excel_Open(False)
   If @error Then
;~ 	  MsgBox($MB_ICONERROR,@ScriptName, "Error open EXCEL." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Return
   EndIf
   zgReadExcelToArray($aExcelBooking,$oExcel,$sFullPathFileName)
   If @error Then
	  MsgBox($MB_ICONWARNING,"提示","读取Booking文件失败！")
	  Return
   EndIf
	_Excel_Close($oExcel)

   for $i = UBound($aExcelBooking)-1 To 0 step -1
	  if $aExcelBooking[$i][14]="" OR $aExcelBooking[$i][14]="ASNNumber" Then
		 _ArrayDelete($aExcelBooking,$i)
	  EndIf
   Next

   _ArrayDelete($aInputInfo,"0-"&(UBound($aInputInfo)-1))

   For $i=0 TO UBound($aExcelBooking)-1 Step 1
	  $doing += 1
	  GUICtrlSetData($idProgressbar, $doing/UBound($aExcelBooking)*100)
	  $aExcelBooking[$i][14]=StringReplace($aExcelBooking[$i][14],"@","")
	  If $aExcelBooking[$i][14]="" Then ContinueLoop
	  zgGetInfoFromFile($InputPath,$aExcelBooking[$i][14])
	  If @error Then
		 WriteMemo($MB_ICONWARNING, $i + 1 & "/" & UBound($aExcelBooking) & "【" & $aExcelBooking[$i][14] & "】","没有找到文件目录",False)
		 ContinueLoop
	  EndIf
	  Local $sReturn = zgCheckInformation($aExcelBooking[$i][14])
	  If @error Then
		 WriteMemo($MB_ICONWARNING, $i + 1 & "/" & UBound($aExcelBooking) & "【" & $aExcelBooking[$i][14] & "】","数据校验出错，" & $sReturn,False)
		 ContinueLoop
	  EndIf
   Next
   GUICtrlSetData($idProgressbar, 100)
   GUICtrlSetData($g_idOutputData,$OutFileName)
   zgSaveOutputData($OutFileName,$aInputInfo,$aDecList,True)
   WriteMemo($MB_ICONWARNING,"数据收集整理完毕！","",False)
   TrayTip ( @ScriptName, "数据收集整理完毕！", 0.5 )
   _Excel_Close($oExcel)
EndFunc

Func zgSingleWindow()	;报关单暂存
    Local $oExcel = _Excel_Open(False)
   If @error Then
;~ 	  MsgBox($MB_ICONERROR,@ScriptName, "Error open EXCEL." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Return
   EndIf

   $bMakeXML=True
   $doing=1
   WriteMemo(0,"报关单暂存数据生成开始。。。","",False)
   Local $datafilename=_GUICtrlEdit_GetText($g_idOutputData)
   If $datafilename="" Then
	  MsgBox($MB_ICONWARNING,"提示","请先选择已生成的数据文件！")
	  Return
   EndIf

    Local $oExcel = _Excel_Open(False)
   If @error Then
;~ 	  MsgBox($MB_ICONERROR,@ScriptName, "Error open EXCEL." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Return
   EndIf
   zgReadExcelToArray($aInputInfo,$oExcel,$datafilename, "Sheet1")
   zgReadExcelToArray($aDecList,$oExcel,$datafilename, "Sheet2")
   If @error Then
	  MsgBox($MB_ICONWARNING,"提示","读取数据文件失败，请联系开发人员。")
	  Return
   EndIf
   _Excel_Close($oExcel)
;~    _ArrayDisplay($aInputInfo)

   Local $SingleWindownum=0
   For $i=0 To UBound($aInputInfo)-1 Step 1
	  If StringLen($aInputInfo[$i][1])>0 And Not IsDECInputContract($aInputInfo[$i][1]) Then
		 $SingleWindownum +=1
	  EndIf
   Next

   For $i=0 To UBound($aInputInfo)-1 Step 1
	  If StringLen($aInputInfo[$i][1])>0 And Not IsDECInputContract($aInputInfo[$i][1]) Then
		 If zgMakeXML($aInputInfo,$aDecList,$i)=-1 Then ContinueLoop
		 TrayTip ( @ScriptName, "报关单暂存数据生成" & $doing &"/" & $SingleWindownum &"【" & $aInputInfo[$i][1]&"】", 10 )
		 WriteMemo(0,"报关单暂存数据生成 " & $doing &"/" & $SingleWindownum &"【" & $aInputInfo[$i][1]&"】","",False)
		 $doing+=1
		 FileWriteLine($DECInputFileName,$aInputInfo[$i][1])
	  EndIf
   Next
   WinActivate("和越数据整理录入")
   WriteMemo(0,"报关单暂存数据生成处理结束","",False)
   TrayTip(@ScriptName,"报关单暂存数据生成处理结束",0.5)
   $bMakeXML=False
   _Excel_Close($oExcel)
EndFunc

Func zgINVSave()	;核注清单暂存   参数说明:  核注清单 暂存,申报 状态标志位  0 暂存  1 申报
   $bMakeXML=True
   $DelcareFlag="0"
   $doing=1
   WriteMemo(0,"核注清单暂存数据生成开始。。。","",False)
   Local $datafilename=_GUICtrlEdit_GetText($g_idOutputData)
   If $datafilename="" Then
	  MsgBox($MB_ICONWARNING,"提示","请先选择已生成的数据文件！")
	  Return
   EndIf

    Local $oExcel = _Excel_Open(False)
   If @error Then
;~ 	  MsgBox($MB_ICONERROR,@ScriptName, "Error open EXCEL." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Return
   EndIf
   zgReadExcelToArray($aInputInfo,$oExcel,$datafilename, "Sheet1")
   zgReadExcelToArray($aDecList,$oExcel,$datafilename, "Sheet2")
   If @error Then
	  MsgBox($MB_ICONWARNING,"提示","读取数据文件失败，请联系开发人员。")
	  Return
   EndIf
   _Excel_Close($oExcel)

   Local $SingleWindownum=0
   For $i=0 To UBound($aInputInfo)-1 Step 1
	  If StringLen($aInputInfo[$i][1])>0 And Not IsINVInputContract($aInputInfo[$i][1]) Then
		 $SingleWindownum +=1
	  EndIf
   Next

;~    _ArrayDisplay($aInputInfo)
;~    _ArrayDisplay($aDecList)
   Local $aRet,$row,$col,$newContract=False
   For $i=0 To UBound($aInputInfo)-1 Step 1
	  If StringLen($aInputInfo[$i][1])>0 And Not IsINVInputContract($aInputInfo[$i][1]) Then
		 ;在数据库中查询合同号是否处理过，已经记录则弹出窗口，并跳过此号
		 $newContract=True
		 mySQLQuery($DBFileName,"SELECT * FROM INVSave WHERE ContractNo='"&$aInputInfo[$i][1]&"';",$aRet,$row,$col)
		 If $row > 0 Then
			$newContract=False
			If $aRet[1][1] <> _NowCalcDate() Then
			   WriteMemo(0,StringFormat("合同号【%s】已经于　%s　暂存处理过！\n\r请手工处理！",$aInputInfo[$i][1],$aRet[1][1]),"",False)
			   ContinueLoop
			EndIf
		 EndIf

		 If zgMakeINVXML($aInputInfo,$i,$aDecList,$DelcareFlag)=-1 Then ContinueLoop
		 TrayTip ( @ScriptName, "核注清单暂存数据生成" & $doing &"/" & $SingleWindownum &"【" & $aInputInfo[$i][1]&"】", 10 )
		 WriteMemo(0,"核注清单暂存数据生成 " & $doing &"/" & $SingleWindownum &"【" & $aInputInfo[$i][1] & "】","",False)
		 $doing+=1
		 FileWriteLine($INVInputFileName,$aInputInfo[$i][1])
		 ;将合同号记入数据库
		 If $newContract Then
			mySQLExec($DBFileName,"INSERT INTO INVSave('ContractNo','SaveDate','Flag') VALUES ('"&$aInputInfo[$i][1]&"','" & _NowCalcDate() & "','Save');")
		 EndIf
	  EndIf
   Next
   WinActivate("和越数据整理录入")
   WriteMemo(0,"核注清单暂存数据生成处理结束","",False)
   TrayTip(@ScriptName,"核注清单暂存数据生成处理结束",0.5)
   $bMakeXML=False
EndFunc

Func zgINVApply()	;核注清单申报	参数说明:  核注清单 暂存,申报 状态标志位  0 暂存  1 申报
   $bMakeXML=True
   $DelcareFlag="1"
   $doing=1
   WriteMemo(0,"核注清单申报数据生成开始。。。","",False)
   Local $datafilename=_GUICtrlEdit_GetText($g_idOutputData)
   If $datafilename="" Then
	  MsgBox($MB_ICONWARNING,"提示","请先选择已生成的数据文件！")
	  Return
   EndIf

    Local $oExcel = _Excel_Open(False)
   If @error Then
;~ 	  MsgBox($MB_ICONERROR,@ScriptName, "Error open EXCEL." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Return
   EndIf
   zgReadExcelToArray($aInputInfo,$oExcel,$datafilename, "Sheet1")
   zgReadExcelToArray($aDecList,$oExcel,$datafilename, "Sheet2")
   If @error Then
	  MsgBox($MB_ICONWARNING,"提示","读取数据文件失败，请联系开发人员。")
	  Return
   EndIf
   _excel_Close($oExcel)
;~    _ArrayDisplay($aInputInfo)

   Local $SingleWindownum=0
   For $i=0 To UBound($aInputInfo)-1 Step 1
	  If StringLen($aInputInfo[$i][1])>0 And Not IsINVApplyContract($aInputInfo[$i][1]) Then
		 $SingleWindownum +=1
	  EndIf
   Next

   For $i=0 To UBound($aInputInfo)-1 Step 1
	  If StringLen($aInputInfo[$i][1])>0 And Not IsINVApplyContract($aInputInfo[$i][1]) Then
		 If zgMakeINVXML($aInputInfo,$i,$aDecList,$DelcareFlag)=-1 Then ContinueLoop
		 TrayTip ( @ScriptName, "核注清单申报数据生成" & $doing &"/" & $SingleWindownum &"【" & $aInputInfo[$i][1]&"】", 10 )
		 WriteMemo(0,"核注清单申报数据生成 " & $doing &"/" & $SingleWindownum &"【" & $aInputInfo[$i][1] & "】","",False)
		 $doing+=1
		 FileWriteLine($INVApplyFileName,$aInputInfo[$i][1])
	  EndIf
   Next
   WinActivate("和越数据整理录入")
   WriteMemo(0,"核注清单申报数据生成处理结束","",False)
   TrayTip(@ScriptName,"核注清单申报数据生成处理结束",0.5)
   $bMakeXML=False
EndFunc

Func IsDir($sFilePath)
;~    MsgBox(0,"",FileGetAttrib($sFilePath))
    Return StringInStr(FileGetAttrib($sFilePath), "D") > 0
EndFunc   ;==>IsDir

Func IsDECInputContract($sContract)
   Local $aInputContract=FileReadToArray($DECInputFileName)
   Local $i
   For $i=0 To UBound($aInputContract)-1 Step 1
	  If $sContract = $aInputContract[$i] Then Return True
   Next
   Return False
EndFunc		;==>IsInputContract

Func IsINVInputContract($sContract)
   Local $aInputContract=FileReadToArray($INVInputFileName)
   Local $i
   For $i=0 To UBound($aInputContract)-1 Step 1
	  If $sContract = $aInputContract[$i] Then Return True
   Next
   Return False
EndFunc		;==>IsInputContract

Func IsINVApplyContract($sContract)
   Local $aInputContract=FileReadToArray($INVApplyFileName)
   Local $i
   For $i=0 To UBound($aInputContract)-1 Step 1
	  If $sContract = $aInputContract[$i] Then Return True
   Next
   Return False
EndFunc		;==>IsInputContract

Func Terminate($parm)
    return 0
 EndFunc   ;==>Terminate

 Func WriteMemo($icro,$title,$tMemo,$msg=False)

   if ($msg) Then MsgBox($icro,$title,$tMemo,3)

   $tMemo = StringFormat("%s  %s  %s" & @CRLF,_Date_Time_SystemTimeToDateTimeStr(_Date_Time_GetLocalTime(),1),$title,$tMemo)
   FileWriteLine($LogFileName,$tMemo)

   _GUICtrlEdit_BeginUpdate($g_idMemo)
   If _GUICtrlEdit_GetLineCount($g_idMemo) > 200 Then
	  $sTmp = _GUICtrlEdit_GetText($g_idMemo)
	  $sTmp = StringRight($sTmp,StringLen($sTmp) - StringLen(_GUICtrlEdit_GetLine($g_idMemo,0)) - 1 )

	  _GUICtrlEdit_SetText($g_idMemo,$sTmp)
   EndIf
   _GUICtrlEdit_AppendText($g_idMemo,$tMemo)
   _GUICtrlEdit_EndUpdate($g_idMemo)
   _GUICtrlEdit_Scroll($g_idMemo, $SB_SCROLLCARET)
EndFunc


Func Timer ($hWnd, $uiMsg, $idEvent, $dwTime)

    Switch $idEvent ; 根据定时器 ID 来进行操作

       Case $TimerDLL[0]

           zgCheckResponse()

;~        Case $Timer2DLL[0]

;~            $t2 += 1

;~            GUICtrlSetData ($Label2, $t2) ; 更新 +1

;~        Case $Timer3DLL[0]

;~            $t3 *= 2

;~            GUICtrlSetData ($Label3, $t3) ; 更新 2 的自乘

       EndSwitch

EndFunc

Func zgBookDirZIp()
   GUICtrlSetData($idProgressbar, 0)
   $doing=0

   $sFullPathFileName = _GUICtrlEdit_GetText ($g_idBookingData)
   If ($sFullPathFileName = "") Then
	  MsgBox($MB_ICONWARNING,"提示","请先选择Booking数据文件！")
	  Return
   EndIf

   Local $InputPath = FileSelectFolder ( "选择数据目录", "" )
   If ($InputPath = "")Then Return

;~    WriteMemo(0,"数据收集整理开始。。。","",False)
   WriteMemo(0,"读取Booking文件","",False)

    Local $oExcel = _Excel_Open(False)
   If @error Then
;~ 	  MsgBox($MB_ICONERROR,@ScriptName, "Error open EXCEL." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Return
   EndIf
   zgReadExcelToArray($aExcelBooking,$oExcel,$sFullPathFileName)
   If @error Then
	  MsgBox($MB_ICONWARNING,"提示","读取Booking文件失败！")
	  Return
   EndIf
	_Excel_Close($oExcel)

   for $i = UBound($aExcelBooking)-1 To 0 step -1
	  if $aExcelBooking[$i][14]="" OR $aExcelBooking[$i][14]="ASN NO" Then
		 _ArrayDelete($aExcelBooking,$i)
	  EndIf
   Next
	_excel_Close($oExcel)

   _ArrayDelete($aInputInfo,"0-"&(UBound($aInputInfo)-1))

   Local $zipfile = $ZipFilePath & 'BookingZip_' & @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC & '.zip'
   For $i=0 TO UBound($aExcelBooking)-1 Step 1
	  $doing += 1
	  GUICtrlSetData($idProgressbar, $doing/UBound($aExcelBooking)*100)
	  $aExcelBooking[$i][14]=StringReplace($aExcelBooking[$i][14],"@","")
	  If $aExcelBooking[$i][14]="" Then ContinueLoop

	  Local $sPathName = $InputPath & "\" & $aExcelBooking[$i][14]
	  Local $iDirExists = FileExists($sPathName)
	  If $iDirExists Then
;~ 		 MsgBox($MB_SYSTEMMODAL, "", "The file exists." & @CRLF & "FileExist returned: " & $iFileExists)
;~ 		 ZIP打包
		 Local $zipcmd = '"' & @ScriptDir & '\7z.exe" a "' & $zipfile & '" "' & $sPathName & '"'
		 RunWait($zipcmd,"",@SW_HIDE)
		 If @error Then return -1

	  Else
;~ 		 MsgBox($MB_SYSTEMMODAL, "", "The file doesn't exist." & @CRLF & "FileExist returned: " & $iFileExists)
		 WriteMemo($MB_ICONWARNING,"【" & $aExcelBooking[$i][14] & "】","没有找到文件目录",False)
		 ContinueLoop
	  EndIf

   Next
   GUICtrlSetData($idProgressbar, 100)
   GUICtrlSetData($g_idOutputData,$OutFileName)
;~    zgSaveOutputData($OutFileName,$aInputInfo,True)
   WriteMemo($MB_ICONWARNING,"数据ZIP完毕！" & @CRLF & $zipfile,"",False)
   TrayTip ( @ScriptName, "数据ZIP完毕！" & @CRLF & $zipfile, 0.5 )
EndFunc

Func zgCleanOldData()
   If MsgBox($MB_ICONWARNING + $MB_YESNO,"清理过期数据","点击是之后将会清理3天之前的所有数据文件，包括"&@CRLF&"DEC\[FailBox、Inbox、OutBox、SentBox]。"&@CRLF&"SAS\[FailBox、Inbox、OutBox、SentBox]。"&@CRLF&@CRLF&"仅保留最近3天的数据！") = $IDNO Then Return
;~ 	  $OutFileName = @ScriptDir & "Output\"

;~ 	   FileGetTime("

   zgDelDir(@ScriptDir & "\Output\DEC\Failbox",3)
   zgDelDir(@ScriptDir & "\Output\DEC\Inbox",3)
   zgDelDir(@ScriptDir & "\Output\DEC\Sentbox",3)
;~    zgDelDir(@ScriptDir & "Output\DEC\Inbox",3)

;~    MsgBox(0,"to do delte files","")

EndFunc

Func zgDelDir($sDir,$iDays)
   Local $dBeforDate = @YEAR&@MON&@MDAY-$iDays&"000000"
   Local $sFileName = $sDir & "\*.*"
;~    MsgBox(0,"",$sFileName)
   Local $hSearch = FileFindFirstFile($sFileName)

   ; Check if the search was successful, if not display a message and return False.
   If $hSearch = -1 Then
;~ 	 MsgBox($MB_SYSTEMMODAL, "", "Error: No files/directories matched the search pattern.")
	 Return False
   EndIf

   ; Assign a Local variable the empty string which will contain the files names found.
   Local $dFileTime = 0

   While 1
	 $sFileName = FileFindNextFile($hSearch)
	 ; If there is no more file matching the search.
	 If @error Then ExitLoop

   $sFileName = $sDir & "\" & $sFileName
	 ; Display the file name.
	 $dFileDate = FileGetTime($sFileName,$FT_CREATED ,$FT_STRING )
	 If $dFileDate < $dBeforDate Then
;~ 		MsgBox(0,$sFileName,FileGetAttrib($sFileName))
		If StringInStr(FileGetAttrib($sFileName), "D") > 0 Then
		   MsgBox(0,"目录","日期早于"&$iDays&"天前，删除！" & @CRLF & $sDir & "\" & $sFileName)
;~ 			DirRemove($sFileName)
		Else
		   MsgBox(0,"文件","日期早于"&$iDays&"天前，删除！" & @CRLF & $sDir & "\" & $sFileName)
;~ 		   FileDelete($sFileName)
		EndIf

	 EndIf
   WEnd

   ; Close the search handle.
   FileClose($hSearch)

EndFunc

