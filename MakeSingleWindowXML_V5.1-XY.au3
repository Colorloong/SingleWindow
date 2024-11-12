#makesinglewindowxml_v5.au3include-once
#include <array.au3>
#include <date.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include "XML.au3"
#include "Zip.au3"

Func zgMakeXML(ByRef $Array,ByRef $GodInfoList,$index)
   While (UBound($Array,$UBOUND_COLUMNS) < 40)
	  _ArrayColInsert($Array, UBound($Array,$UBOUND_COLUMNS))
   WEnd

   If $Array[$index][30] = "" Then Return -1	;无核注单号，不报关
;~ 	  WriteMemo(0,"生成报关单暂存报文失败","没有对应的核注单编号，不能报关！",False)
;~ 	  Return -1
;~    EndIf

	; This COM Error Hanlder will be used globally (excepting inside UDF Functions)
	Global $oErrorHandler = ObjEvent("AutoIt.Error", ErrFunc_CustomUserHandler_MAIN)
	#forceref $oErrorHandler

	; This is SetUp for the transfer UDF internal COM Error Handler to the user function
	_XML_ComErrorHandler_UserFunction(ErrFunc_CustomUserHandler_XML)

	;~ 	Local $oErrorHandler = ObjEvent("AutoIt.Error", "_ErrFunc_CustomUserHandler")
	;~  #forceref $oErrorHandler
	Local $xmlFileName = "CDHY_" & $Array[$index][1]&"_"&@YEAR&@MON&@MDAY
	Local $xmlFullPath = @ScriptDir & "\Output\Dec\OutBox\"&$xmlFileName & ".xml"
	Local $tmpFileName = @TempDir & "\XMLTemp.temp"

	; first you must create $oXmlDoc object
	Local $oXmlDoc = _XML_CreateFile($tmpFileName,"DecMessage",True)
 	FileDelete($tmpFileName)

;~ 	Local $asAttributeList[2][2] = [["version","3.1"],["xmlns","http://www.chinaport.gov.cn/dec"]]
;~ 	_XML_CreateAttribute($oXmlDoc, "DecMessage",$asAttributeList)
    _XML_SetAttrib($oXmlDoc, "DecMessage", "version", "5.0")
    _XML_SetAttrib($oXmlDoc, "DecMessage", "xmlns", "http://www.chinaport.gov.cn/dec")
;~ 	_XML_MakeNode($oXmlDoc, "ImportMessage", "SWImpHead","")
;~ 	_XML_MakeNode($oXmlDoc, "ImportMessage/SWImpHead","CopMsgId",$xmlFileName)
;~ 	_XML_MakeNode($oXmlDoc, "ImportMessage/SWImpHead","MsgType","SWDEC001")
;~ 	_XML_MakeNode($oXmlDoc, "ImportMessage/SWImpHead","FunctionCode","0")
;~ 	_XML_MakeNode($oXmlDoc, "ImportMessage/SWImpHead","Format","xml")
;~ 	_XML_MakeNode($oXmlDoc, "ImportMessage/SWImpHead","Compress","0")
;~ 	_XML_MakeNode($oXmlDoc, "ImportMessage/SWImpHead","Version","1.0")

;~ 	_XML_MakeNode($oXmlDoc, "ImportMessage", "SWImpData","")
;~ 	_XML_MakeNode($oXmlDoc, "ImportMessage/SWImpData","DecMessage","")
;~ 	_XML_MakeNode($oXmlDoc, "ImportMessage/SWImpData/DecMessage","DecHead","")

   Local $sBillNo
   Local $OutPort = TransTragetPort($Array[$index][14])
   If CountItemInArrayColumn($Array,$index,15) = 1 AND $Array[$index][14] = "CTU" Then
	  $sBillNo = StringRegExpReplace($Array[$index][15],"[^\w]","")
   Else
	  $sBillNo = StringRegExpReplace($Array[$index][15],"[^\w]","") & "_" & $Array[$index][16]
   EndIf
   Local $aTmp = StringSplit($Array[$index][18],"/")
		;$aTmp[1];运输工具名称
		;TransMonNum($aTmp[2]),"航次号

	_XML_MakeNode($oXmlDoc, "DecMessage","DecHead", "","")
    _XML_RemoveAttributeNode($oXmlDoc, "DecMessage/DecHead","xmlns")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","SeqNo","","关检关联号")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","IEFlag","E","进出口标识（海关）I：进口E：出口")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","AgentCodeScc","91310107132231740K","报关申报单位社会信用代码")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","AgentCode","3107980011","报关申报单位代码")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","AgentName","上海兴亚报关有限公司","报关申报单位名称")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","ApprNo","","批准文号")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","BillNo",$sBillNo,"提/运单号")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","ContrNo",$Array[$index][1],"合同协议号（海关）")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","CustomMaster","7922","申报地海关代码");测试环境用7900
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","Type","ML","单据类型：属地报关SD；备案清单：ML。LY：两单一审备案清单。CL:汇总征税报关单。SS:”属地申报，属地验放”报关单；MT:多式联运,Z:自报自缴。兼容第三方客户端，将该字段做成可选。")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","CutMode","","征免性质代码")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","FeeMark","",'运费标记"1"率;"2"单价;"3"总价')
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","FeeRate","","运费／率")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","GrossWet",$Array[$index][5],"报关毛重")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","IEDaIEDate","","进口日期（进口）/出口日期（出口）")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","IEPort",$OutPort,"进出口口岸代码")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","FeeCurr","","运费币制")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","InsurMark","",'保险费标记"1"率;"3"总价')
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","InsurRate","","保险费／率")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","InsurCurr","","保险费费币制")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OtherMark","",'杂费标记"1"率;"3"总价;2单价')
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OtherRate","","杂费／率")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OtherCurr","","杂费币制")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","DistinatePort",$Array[$index][27],"装货港（进口）/指运港（出口）")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","TradeCountry",$Array[$index][0],"贸易国（地区）代码（海关）")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","TradeMode","5034","监管方式代码")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","TrafMode",$Array[$index][13],"运输方式代码（海关）")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","TrafName",$aTmp[1],"运输工具代码及名称")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","TradeAreaCode","SGP","贸易国（地区）代码（海关）");132新加坡
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","DespPortCode","","启运港代码")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","LicenseNo","","许可证编号")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","ManualNo","H792218A0004","报关备案号");H79226000008
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","NetWt",$Array[$index][6],"净重")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","NoteS","","备注（海关）")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OwnerCodeScc","91510100569680768F","消费使用单位（进口）/生产销售单位（出口）社会信用代码")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OwnerCode","5101644060","消费使用单位（进口）/生产销售单位（出口）代码")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OwnerName","戴尔（成都）有限公司","消费使用单位（进口）/生产销售单位（出口）名称")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","PackNo",$Array[$index][2],"件数")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","TradeCoScc","91510100569680768F","收发货人社会信用代码（海关）")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","TradeCode","5101644060","收发货人代码（海关）")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","TradeName","戴尔（成都）有限公司","收发货人名称")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","TransMode","3","成交方式代码")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","EntryId","","海关编号")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","PreEntryId","","报关预录入编号")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","EdiId","","1：普通报关3：北方转关提前5：南方转关提前6：普通报关，运输工具名称以‘◎’开头南方H2000直转")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","Risk","","上海EDI报关单：最后10位用于存放风险评估参数RISK。")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","CopCodeScc","91310107132231740K","企业内部社会信用代码")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","CopCode","3107980011","企业内部编号")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","CopName","上海兴亚报关有限公司","企业内部名称")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","EntryType","M","报关单类型L有纸带清单报关0有纸报关D无纸带清单报关M通关无纸化")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","PDate","","首次进行暂存操作的系统时间")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","TypistNo","","录入员IC卡号")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","InputerName","","填写人员")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","PartenerID","","申报人标识，上海导入时使用")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","TgdNo","","宁波通关申请单号")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","DataSource","","数据来源，PTS代表由电子手册生成，SZJB代表由深圳集报生成，其它为空")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","BillType","1","备案清单类型1：一般备案清单2：先进区、后报关3：集报备案清单4：两单一审备案清单")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","DeclTrnRel","0","报关/转关关系标志。0：一般报关；1：转关运输提前报关")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","ChkSurety","0","担保验放标志 1:是 0否")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","PromiseItmes","999","承诺事项1-勾选0-未选第一位：特殊关系确认第二位：价格影响确认")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","CheckFlow","","查验分流")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","TaxAaminMark","","税收征管标记：0-无 1-有");???
    _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","MarkNo","N/M","标记及号码【本批货物的标记和号码】");???
    _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","EntyPortCode","511001","入境口岸代码【货物从运输工具卸离的第一个境内口岸】")
    _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","GoodsPlace","分拨仓库","货物存放地点【报检时货物的存放地点】")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","BLNo","","提货单号【本批货物的提货单或出库单号码】")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","InspOrgCode","","施检机构【对本批货物实施检验检疫的机构】")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","SpecDeclFlag","","特种业务标识：0未勾选，1选中。")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","PurpOrgCode","","目的机构代码【入境货物流向的目的地检验检疫机构】")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","DespDate","","发货日期【本批拟发货的日期】")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","CmplDschrgDt","","卸毕日期【本批货物全部卸离运输工具的日期】")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","CorrelationReasonFlag","","关联理由【关联报检号的关联理由】")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","VsaOrgCode","","领证地")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OrigBoxFlag","","入境原集装箱装载直接到目的机构，【1：是；0：否】")
    _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","DeclareName","胡琳","报关员姓名")
    _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","NoOtherPack","1","勾选 0-未选，有其他包装；1：选中，无其他包装。")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OrgCode","","检验检疫受理机关")

;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OverseasConsignorCname","","境外发货人中文名称")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OverseasConsignorEname","","境外发货人名称（外文）")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OverseasConsignorAddr","","境外发货人地址")

;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OverseasConsigneeCode","91510100569680888F","境外收货人编码")
    _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OverseasConsigneeEname",$Array[$index][25],"境外收货人名称(外文)")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","DomesticConsigneeEname","","境内收货人名称(外文) ")

;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","CorrelationNo","","关联号码")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","EdiRemark","","EDI申报备注")
;~     _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","EdiRemark2","","EDI申报备注2")


;***********************************************************************
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","PayWay","1","征税比例")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","PaymentMark","","纳税单位")
;~    Local $Memo
;~    If $OutPort = "7902" Then
;~ 	  $Memo = "全国一体化：7902，" & $Array[$index][1]
;~    Else
;~ 	  $Memo = "全国一体化：" &$OutPort&"，"&$Array[$index][1]
;~    EndIf
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","InRatio","1","征税比例")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","AgentLinkMan","","报关员编号")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","AgentLinkMAIL","","报关员联系方式EMAIL")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","AgentLinkPHONE","","报关员联系方式联系电话")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OwnerLinkMan","","消费使用单位联系人")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OwnerLinkMAIL","","消费使用单位联系人EMAIL")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","OwnerLinkPHONE","","消费使用单位联系人联系电话")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","InputerNameMAIL","","填写人员EMAIL")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecHead","InputerNamePHONE","","填写人员联系电话")
;*************************************************************************

	_XML_MakeNode($oXmlDoc, "DecMessage","DecLists","","报关单表体")
   Local $iMake = 0
   Local $GName,$CodeTS,$GUnit,$FirstUnit,$TradeCurr,$UsedTo
   For $iList = 0 To UBound($GodInfoList) - 1 Step 1
	  If $GodInfoList[$iList][0] <> $Array[$index][1] Then ContinueLoop

	  For $i = 0 To UBound($aTSCode)-1 Step 1
		 If $aTSCode[$i][8] = $Array[$index][7] Then
			$CodeTS = $aTSCode[$i][1]		;商品编码
			$INVCodeTS = $aTSCode[$i][8]	;对应核注清单表体商品序号
			$GName = $aTSCode[$i][2]		;商品名称
			$Gunit = $aTSCode[$i][3]		;成交单位
			$FirstUnit = $aTSCode[$i][4]	;法定单位
			$TradeCurr = $aTSCode[$i][5]	;成交币制
			$INVTradeCurr = $aTSCode[$i][9]	;商品核注单币制
			$GdsMtno = $aTSCode[$i][10]	;商品料号
			$UsedTo = $aTSCode[$i][7]	;用途
			ExitLoop
		 EndIf
	  Next

	  Local $GMode1
	  If $Array[$index][7] = 1 OR $Array[$index][7] = 10 Then	;项号1为笔记本，10为平板电脑，要在其他栏填写操作系统，非1,则在型号栏填入
		 $GMode1 = "4|0|" & StringregExpreplace($GodInfoList[$iList][7],"([!#+^{}])","{$1}") _;配置
			& "|" & $GodInfoList[$iList][9] _;品牌
			& "|" & StringReplace($GodInfoList[$iList][5],"等","") & " " & $GodInfoList[$iList][6] _;机型 型号
			& "|||" & $GodInfoList[$iList][8];OS
	  ElseIf $Array[$index][7] = 3 Then
		 $GMode1 = "4|0|微型机|" & StringregExpreplace($GodInfoList[$iList][7],"([!#+^{}])","{$1}") _;配置
			& "|" & $GodInfoList[$iList][9] _;品牌
			& "|" & StringReplace($GodInfoList[$iList][5],"等","") & " " & $GodInfoList[$iList][6] _;机型 型号
			& "|" & $GodInfoList[$iList][8] _;OS
			& "|" & $UsedTo	;用途
	  Else
		 $GMode1 = "4|0|" & StringregExpreplace($GodInfoList[$iList][7],"([!#+^{}])","{$1}") _;配置
			& "|" & $GodInfoList[$iList][9] _;品牌
			& "|" & StringReplace($GodInfoList[$iList][5],"等","") & " " & $aDecList[$iList][6] _;机型 型号
			& "|" & $GodInfoList[$iList][8] _;OS
			& "|" & $UsedTo	;用途
	  EndIf

	  $iMake += 1

	  _XML_MakeNode($oXmlDoc, "DecMessage/DecLists","DecList","")
;~ 	  _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","ClassMark","","归类标志")
	  _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","GNo",$imake,"货物序号")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","ContrItem","","报关单备案序号");
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","CodeTS",$CodeTS,"商品编码")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","GName",$GName,"商品名称")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","GModel",$GMode1,"规格型号")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","FirstQty",$GodInfoList[$iList][2],"法定数量")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","FirstUnit",$FirstUnit,"法定单位代码")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","GQty",$GodInfoList[$iList][2],"成交数量")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","GUnit",$Gunit,"成交计量单位代码")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","SecondQty",$GodInfoList[$iList][4],"法定第二数量")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","SecondUnit","035","法定第二计量单位代码")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","OriginCountry",$Array[$index][0],"原产国(地区)")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","TradeCurr",$TradeCurr,"成交币制（海关）")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","DeclPrice",Round($GodInfoList[$iList][3]/$GodInfoList[$iList][2],4),"商品成交单价")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","WorkUsd","","工缴费")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","UseTo","","用途/生产厂家")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","ExgNo","","货号")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","DutyMode","3","征减免税方式")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","Factor","","申报计量单位与法定单位比例因子")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","DeclTotal",Round($GodInfoList[$iList][3],2),"成交总价（海关）")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","ExgVersion","1","版本号")

	   ;对于出口报关单，表体的originalCountry需要填目的国，destinationCountry填的是原产国。
	   Local $DestCountry
	   If $GodInfoList[$iList][1] = "CN" Then
		  $DestCountry = "CHN"
	   ElseIf $GodInfoList[$iList][1] = "VN" Then
		  $DestCountry ="VNM"
	   ElseIf $GodInfoList[$iList][1] = "TW" Then
		  $DestCountry ="TWN"
	   EndIf
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","DestinationCountry",$DestCountry,"最终目的国（地区）")

   ;下面的检验检疫相关的部分
   ;~ 		 If $Array[$index][7] = 1 Then
   ;~ 			_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","CiqCode","101","填写3位检验检疫编码	")
   ;~ 			_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","CiqName","平板电脑(重量≤10千克,至少由一个中央处理器、键盘和显示器组成)(便携数字自动数据处理机)","")
   ;~ 		 ElseIf $Array[$index][7] = 5 Then
   ;~ 			_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","CiqCode","101","填写3位检验检疫编码	")
   ;~ 			_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","CiqName","其他便携式自动数据处理设备(重量≤10千克,至少由一个中央处理器、键盘和显示器组成)(便携数字自动数据处理机)","")
   ;~ 		 ElseIf $Array[$index][7] = 6 Then
   ;~ 			_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","CiqCode","999","填写3位检验检疫编码	")
   ;~ 			_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","CiqName","其他的微型机的处理部件(不论是否在同一机壳内有一或两个存储,输入或输出部件)","")
   ;~ 		 ElseIf $Array[$index][7] = 12 Then
   ;~ 			_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","CiqCode","999","填写3位检验检疫编码	")
   ;~ 			_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","CiqName","其他的微型机的处理部件(不论是否在同一机壳内有一或两个存储,输入或输出部件)","")
   ;~ 		 EndIf
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","CiqCode","101","填写3位检验检疫编码	")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","CiqName","其他便携式自动数据处理设备(重量≤10千克,至少由一个中央处理器、键盘和显示器组成)(便携数字自动数据处理机)","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","DeclGoodsEname",","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","OrigPlaceCode","","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","Purpose","","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","ProdValidDt","","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","ProdQgp","","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","GoodsAttr","","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","Stuff","","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","Uncode","","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","DangName","","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","DangPackType","","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","DangPackSpec","","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","EngManEntCnm","","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","NoDangFlag","","")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","DestCode","510116","目的地代码【货物在境内预定最终抵达的交货地】")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","GoodsSpec","","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","GoodsModel","","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","GoodsBrand","","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","ProduceDate","","")
   ;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","ProdBatchNo","","")
	   _XML_MakeNode($oXmlDoc, "DecMessage/DecLists/DecList[last()]","DistrictCode","51016","进口指境内目的地，出口指境内货源地	")
   Next
	_XML_MakeNode($oXmlDoc, "DecMessage","DecOtherPacks","")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecOtherPacks","DecOtherPack","")

;~ 20180825 注释
;~ 	运输包装	材质	包装种类	其他包装
;~     "PALLET/PLT(5)"	膠合小棧板	再生木托(92)	无
;~ 					   膠合中棧板	再生木托(92)	无
;~ 					   膠合N棧板	再生木托(92)	无
;~ 					   原木中棧板	天然木托(93)	无
;~ 					   原木直角	天然木托(93)	无
;~ 					   原木小棧板	天然木托(93)	无
;~ 					   原木斜角	天然木托(93)	无
;~ 					   塑料棧板	其他包装(99)	无
;~ 					   塑料		其他包装(99)	无
;~ 					   黑色塑膠棧板	其他包装(99)	无
;~ 	"CARTON/CTN(2)"		none	纸质或纤维板制盒/箱(22)	无
;~ 	"PKG/PARCEL(6)"		none	包/袋（6）	无
;~  "PALLET/PLT(5)"	原木小棧板/黑色塑膠棧板	天然木托(93)	其他包装(99)
;~ 					塑料棧板//原木中棧板	天然木托(93)	其他包装(99)
;~ 	"PACKAGE(7)"	原木小棧板/黑色塑膠棧板/卡通（CTN)	天然木托(93)	其他包装(99)+纸质或纤维板制盒/箱(22)

;~    Local $sWrapType,$sOtherPkg
;~    If $Array[$index][3] = "2" Then
;~ 	  $sWrapType = "22"
;~    ElseIf $Array[$index][3] = "5" Then
;~ 	  If StringInStr($Array[$index][28],"原木") > 0 And StringInStr($Array[$index][28],"塑")> 0 Then
;~ 		 $sWrapType = "93"
;~ 		 $sOtherPkg = "99"
;~ 	  ElseIf StringInStr($Array[$index][28],"原木") > 0 Then
;~ 		 $sWrapType = "93"
;~ 	  ElseIf StringInStr($Array[$index][28],"塑") > 0 Then
;~ 		 $sWrapType = "99"
;~ 	  ElseIf StringInStr($Array[$index][28],"膠") > 0 Then
;~ 		 $sWrapType = "92"
;~ 	  EndIf
;~    ElseIf $Array[$index][3] = "6" Then
;~ 	  $sWrapType = "6"
;~    ElseIf $Array[$index][3] = "7" Then
;~ 	  $sWrapType = "93"
;~ 	  If StringInStr($Array[$index][28],"原木") > 0 And StringInStr($Array[$index][28],"塑") > 0 Then
;~ 		 $sOtherPkg = "99"
;~ 	  ElseIf StringInStr($Array[$index][28],"原木") > 0 And  StringInStr($Array[$index][28],"CTN" > 0 Then
;~ 		 $sOtherPkg = "22"
;~ 	  EndIf
;~    EndIf

;~ MsgBox(0,"*"&$sWrapType&"/"&$sOtherPkg&"*",$Array[$index][3] & $Array[$index][28])
   _XML_MakeNode($oXmlDoc, "DecMessage/DecHead","WrapType",StringFormat("%02i" , $Array[$index][3]),"包装种类代码");
   If $Array[$index][28] <> "" Then
	  Local $aOtherPkg[5]
	  $aOtherPkg = StringSplit($Array[$index][28],"/")
	  For $i = 1 To $aOtherPkg[0] Step 1
		 _XML_MakeNode($oXmlDoc, "DecMessage/DecOtherPacks/DecOtherPack","PackType",$aOtherPkg[$i],"包装种类")
	  Next
   EndIf

;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecOtherPacks/DecOtherPack","PackQty",$Array[$index][4],"包装件数")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecOtherPacks/DecOtherPack","PackType","22","包装种类");纸箱 22

	_XML_MakeNode($oXmlDoc, "DecMessage","DecFreeTxt","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecFreeTxt","RelId","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecFreeTxt","RelManNo","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecFreeTxt","BonNo","","")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecFreeTxt","DecBpNo","13761541436","报关员联系方式")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecFreeTxt","CusFie","","")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecFreeTxt","DecNo","22128473","报关员号");谭承艳79000998肖静79000198胡琳22128473
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecFreeTxt","VoyNo",TransMonNum($aTmp[2]),"航次号")

;~ 	_XML_MakeNode($oXmlDoc, "DecMessage","DecGoodsLimits","许可证信息表")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits","DecGoodsLimit","","许可证信息表项")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit","GoodsNo","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit","LicTypeCode","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit","LicenceNo","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit","LicWrtofDetailNo","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit","LicWrtofQty","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit","LicWrtofQtyUnit","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit","DecGoodsLimitVin","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit/DecGoodsLimitVin","LicenceNo","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit/DecGoodsLimitVin","LicTypeCode","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit/DecGoodsLimitVin","VinNo","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit/DecGoodsLimitVin","BillLadDate","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit/DecGoodsLimitVin","QualityQgp","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit/DecGoodsLimitVin","MotorNo","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit/DecGoodsLimitVin","VinCode","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit/DecGoodsLimitVin","ChassisNo","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit/DecGoodsLimitVin","InvoiceNum","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit/DecGoodsLimitVin","ProdCnnm","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit/DecGoodsLimitVin","ProdEnnm","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit/DecGoodsLimitVin","ModeEn","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecGoodsLimits/DecGoodsLimit/DecGoodsLimitVin","PricePerUnit","","")

	_XML_MakeNode($oXmlDoc, "DecMessage","DecSign","")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecSign","OperType","G","操作类型 G：报关单暂存（转关提前报关单暂存） C:报关单申报")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecSign","ICCode","2100040041545","操作人代码")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecSign","CopCode","3107980011","操作员所属企业组织机构代码")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecSign","OperName","胡琳","操作人姓名")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecSign","Sign","","数字签名信息，兼容第三方客户端，将该字段做成可选。")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecSign","SignDate",@YEAR&@MON&@MDAY&" "&@HOUR&@MIN&@SEC&@MSEC,"签名时间。兼容第三方客户端，将该字段做成可选。yyyyMMddHH")
	_XML_MakeNode($oXmlDoc, "DecMessage/DecSign","ClientSeqNo","decnew"&$Array[$index][1],"客户端报关单编号")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecSign","Certificate","","证书号，兼容导入暂存接口，将该字段做成可选。")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecSign","HostId","","客户端邮箱的HostId，兼容导入暂存接口，将该字段做成可选。")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecSign","BillSeqNo","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecSign","DomainId","","")
;~ 	_XML_MakeNode($oXmlDoc, "DecMessage/DecSign","Note","","")

   ;FileType 	00000001:发票 00000002:装箱单  00000003:提/运单  00000004:合同
   Local $sDir = _GUICtrlEdit_GetText ($g_idSelectDir)
   zgAttachPdfFile($oXmlDoc,$outport,$sDir & "\" & $Array[$index][1] & "\",$Array[$index][1] & "_EXP_Contract.PDF","00000004")

   Local $sPathFileName = $sDir & "\" & $Array[$index][1] & "\EXPCI_" & $Array[$index][1] & "*.PDF"
   $sPathFileName = zgFindFile($sPathFileName)
   If @error Then
	  SetError(1,2)
	  Return
   EndIf
   zgAttachPdfFile($oXmlDoc,$outport,$sDir & "\" & $Array[$index][1] & "\",$sPathFileName,"00000001")


   zgAttachPdfFile($oXmlDoc,$outport,$sDir & "\" & $Array[$index][1] & "\",$Array[$index][1] & "_PACKING_LIST_EXP_FINAL.PDF","00000002")

    _XML_MakeNode($oXmlDoc, "DecMessage","DecLicenseDocus","","随附单证")
    _XML_MakeNode($oXmlDoc, "DecMessage/DecLicenseDocus","LicenseDocu","","单证")
    _XML_MakeNode($oXmlDoc, "DecMessage/DecLicenseDocus/LicenseDocu","DocuCode","a","单证代码")
    _XML_MakeNode($oXmlDoc, "DecMessage/DecLicenseDocus/LicenseDocu","CertCode",$Array[$index][30],"单证编号")	;此处应填写核注单正式编号

   ; Tidy XML
;~    MsgBox(0,"",_XML_TIDY($oXmlDoc))
   _XML_LoadXML($oXmlDoc,_XML_TIDY($oXmlDoc))

 	FileDelete($xmlFullPath)
	_XML_SaveToFile($oXmlDoc,$xmlFullPath)

    Local $hFileOpen = FileOpen($xmlFullPath,0)
	Local $sXmlRead = FileRead($hFileOpen)
	FileClose($hFileOpen)
	$sXmlRead = StringReplace($sXmlRead,' xmlns=""',"")
	$sXmlRead = StringReplace($sXmlRead,"{+}","+")
    Local $hFileOpen = FileOpen($xmlFullPath,2)
	FileWrite($hFileOpen,$sXmlRead)
	FileClose($hFileOpen)
EndFunc

Func zgMakeINVXML(ByRef $Array,$index,Byref $GodInfoList,$sDelcareFlag)
   While (UBound($Array,$UBOUND_COLUMNS) < 40)
	  _ArrayColInsert($Array, UBound($Array,$UBOUND_COLUMNS))
   WEnd
   If $sDelcareFlag = "0" And $Array[$index][29] <> "" Then Return -1	;已有暂存号，不再暂存
   If $sDelcareFlag = "1" And ($Array[$index][29] = "" Or $Array[$index][30] <> "") Then Return -1	;无暂存号，或已有编号，不申报
	; This COM Error Hanlder will be used globally (excepting inside UDF Functions)
	Global $oErrorHandler = ObjEvent("AutoIt.Error", ErrFunc_CustomUserHandler_MAIN)
	#forceref $oErrorHandler

	; This is SetUp for the transfer UDF internal COM Error Handler to the user function
	_XML_ComErrorHandler_UserFunction(ErrFunc_CustomUserHandler_XML)

	;~ 	Local $oErrorHandler = ObjEvent("AutoIt.Error", "_ErrFunc_CustomUserHandler")
	;~  #forceref $oErrorHandler

	Local $xmlFileName = "CDHY_" & $Array[$index][1]&"_"&@YEAR&@MON&@MDAY
	Local $xmlFullPath = @ScriptDir&"\Output\Sas\OutBox\" & $xmlFileName & ".xml"
	Local $tmpFileName = @TempDir&"\XMLTemp.temp"

	; first you must create $oXmlDoc object
	Local $oXmlDoc = _XML_CreateFile($tmpFileName,"Signature",True)
 	FileDelete($tmpFileName)

    _XML_SetAttrib($oXmlDoc, "Signature", "schemaLocation","http://www.chinaport.gov.cn/sas SAS101.xsd")
    _XML_SetAttrib($oXmlDoc, "Signature", "xmlns:sas","http://www.chinaport.gov.cn/sas")
    _XML_SetAttrib($oXmlDoc, "Signature", "xmlns:xsi","http://www.w3.org/2001/XMLSchema-instance")

	_XML_MakeNode($oXmlDoc, "Signature","SignedInfo", "","")
    _XML_MakeNode($oXmlDoc, "Signature/SignedInfo","CanonicalizationMethod","","")
    _XML_SetAttrib($oXmlDoc,"Signature/SignedInfo/CanonicalizationMethod","Algorithm","http://www.w3.org/TR/2001/REC-xml-c14n-20010315")
    _XML_MakeNode($oXmlDoc, "Signature/SignedInfo","SignatureMethod","","")
    _XML_SetAttrib($oXmlDoc,"Signature/SignedInfo/SignatureMethod","Algorithm","http://www.w3.org/2000/09/xmldsig#rsa-sha1")
    _XML_MakeNode($oXmlDoc, "Signature/SignedInfo","Reference","","")
    _XML_SetAttrib($oXmlDoc,"Signature/SignedInfo/Reference","URI","String")
    _XML_MakeNode($oXmlDoc, "Signature/SignedInfo/Reference","DigestMethod","","")
	_XML_SetAttrib($oXmlDoc,"Signature/SignedInfo/Reference/DigestMethod","Algorithm","http://www.w3.org/2000/09/xmldsig#sha1")
    _XML_MakeNode($oXmlDoc, "Signature/SignedInfo/Reference","DigestValue","","")


	_XML_MakeNode($oXmlDoc, "Signature","SignatureValue", "","")
	_XML_MakeNode($oXmlDoc, "Signature","KeyInfo", "","")
	_XML_MakeNode($oXmlDoc, "Signature/KeyInfo","KeyName","aa","")


	_XML_MakeNode($oXmlDoc, "Signature","Object", "","")
    _XML_SetAttrib($oXmlDoc,"Signature/Object","ID","String")
	_XML_MakeNode($oXmlDoc, "Signature/Object","Package", "","")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package","EnvelopInfo", "","")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/EnvelopInfo","version", "1.0","")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/EnvelopInfo","message_id", $ContractNO & @YEAR&@MON&@MDAY&@HOUR&@MIN&@SEC&@MSEC,"")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/EnvelopInfo","message_type", "INV101","")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/EnvelopInfo","sender_id", "DXPESW0000125111","")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/EnvelopInfo","receiver_id", "DXPDSWMFT0000001","")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/EnvelopInfo","send_time", @YEAR&"-"&@MON&"-"&@MDAY&" "&@HOUR&":"&@MIN&":"&@SEC,"")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/EnvelopInfo","Ic_Card", "2100040041545","");测试 8930000092641   正式 8930000056429

	_XML_MakeNode($oXmlDoc, "Signature/Object/Package","DataInfo", "","")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo","BussinessData", "","")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData","InvtMessage", "","")



   Local $OutPort = TransTragetPort($Array[$index][14])
   If CountItemInArrayColumn($Array,$index,15) = 1 And $Array[$index][14] = "CTU" Then
	  $sBillNo = StringRegExpReplace($Array[$index][15],"[^\w]","")
   Else
	  $sBillNo = StringRegExpReplace($Array[$index][15],"[^\w]","") & "_" & $Array[$index][16]
   EndIf
   Local $aTmp = StringSplit($Array[$index][18],"/")
		;$aTmp[1];运输工具名称
		;TransMonNum($aTmp[2]),"航次号

;~ ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
;~ 	  核注单表头
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage","InvtHeadType", "","")
    _XML_RemoveAttributeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","xmlns")

	Local $INVSeqNo = ""
	If ($sDelcareFlag = "1")	Then
	   $INVSeqNo = StringRight($Array[$index][29],StringLen($Array[$index][29])-1)
	EndIf

	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","SeqNo",$INVSeqNo,"中心统一编号");首次导入为空，导入成功后返回预录入编号
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","BondInvtNo","","保税清单编号");海关审批通过后系统自动返填
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","ChgTmsCnt","0","变更次数");有变更时填写
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","PutrecNo","H792218A0004","备案编号")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","EtpsInnerInvtNo",$Array[$index][1],"企业内部清单编号");~ 	EtpsInnerInvtNo由企业自行编写
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","BizopEtpsSccd","91510100569680768F","经营企业社会信用代码")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","BizopEtpsno","5101644060","经营企业编号")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","BizopEtpsNm","戴尔（成都）有限公司","经营企业名称")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","RvsngdEtpsSccd","91510100569680768F","收发货企业社会信用代码")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","RcvgdEtpsno","5101644060","收货企业编号")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","RcvgdEtpsNm","戴尔（成都）有限公司","收货企业名称")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","DclEtpsSccd","91310107132231740K","申报企业社会信用代码")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","DclEtpsno","3107980011","申报企业编号")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","DclEtpsNm","上海兴亚报关有限公司","申报企业名称")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","InputCreditCode","91310107132231740K","录入单位社会信用代码")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","InputCode","3107980011","录入单位代码")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","InputName","上海兴亚报关有限公司","录入单位名称")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","InputTime",@YEAR&@MON&@MDAY,"录入日期（格式：yyyyMMdd）")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","InvtDclTime","","清单申报时间（格式：yyyyMMdd）")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","EntryDclTime","","报关单申报时间（格式：yyyyMMdd）");系统自动反填
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","EntryNo","","对应报关单编号")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","CorrEntryDclEtpsSccd","91310107132231740K","对应报关单申报单位社会统一信用代码")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","CorrEntryDclEtpsNo","3107980011","对应报关单申报单位代码")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","CorrEntryDclEtpsNm","上海兴亚报关有限公司","对应报关单申报单位名称")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","RltInvtNo","","关联清单编号")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","RltPutrecNo","","关联备案编号")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","RltEntryNo","","关联报关单编号")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","RltEntryBizopEtpsSccd","","关联报关单境内收发货人社会信用代码")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","RltEntryBizopEtpsno","","关联报关单境内收发货人编号")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","RltEntryBizopEtpsNm","","关联报关单境内收发货人名称")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","RltEntryRvsngdEtpsSccd","","关联报关单收发货单位社会信用代码")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","RltEntryRcvgdEtpsno","","关联报关单海关收发货单位编码")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","RltEntryRcvgdEtpsNm","","关联报关单收发货单位名称")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","RltEntryDclEtpsSccd","","关联报关单申报单位社会统一信用代码")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","RltEntryDclEtpsno","","关联报关单海关申报单位编码")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","RltEntryDclEtpsNm","","关联报关单申报单位名称")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","ImpexpPortcd",$OutPort,"进出境关别")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","DclPlcCuscd","7922","申报地关区代码")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","ImpexpMarkcd","E","进出口标记代码（I-进口 E-出口）")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","MtpckEndprdMarkcd","I","料件成品标记代码（I-料件 E-成品）")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","SupvModecd","5034","监管方式代码")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","TrspModecd",$Array[$index][13],"运输方式代码")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","ApplyNo","","申请编号")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","ListType","","流转类型")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","DclcusFlag","1","是否报关标志（1-报关 2-非报关）")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","DclcusTypecd","2","报关类型代码1.关联报关2.对应报关")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","PrevdTime","","预核扣时间")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","FormalVrfdedTime","","正式核扣时间")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","InvtIochkptStucd","0","清单进出卡口状态代码")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","VrfdedMarkcd","","核扣标记代码（0-未核扣 1-预核扣 2-已核扣 3-已核销）")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","IcCardNo","2100040041545","申报人IC卡号")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","ListStat","1","清单状态")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","DecType","4","报关单类型")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","Rmk",$Array[$index][1] & "  ,  " & $sBillNo  & $sCustomMemo,"备注")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","StshipTrsarvNatcd",TransTargetAddCodeToNum($Array[$index][0]),"起运/运抵国(地区）")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","InvtType","0","清单类型")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","EntryStucd","","报关状态")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","PassportUsedTypeCd","","核放单生成标志代码")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","NeedEntryModified","","报关单同步修改标志")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","LevyBlAmt","","计征金额")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","DclTypecd","1","申报类型(1-备案申请 2-变更申请 3-删除申请)")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtHeadType","GenDecFlag","1","是否生成报关单(1-生成 2-不生成)")

;~ ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝


;~ 	  临时报关单表头
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage","InvtDecHeadType","","")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","SeqNo","","清单中心统一编号")
;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","DecSeqNo","","报关单统一编号")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","PutrecNo","H792218A0004","备案编号")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","BizopEtpsSccd","91510100569680768F","经营企业社会信用代码")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","BizopEtpsno","5101644060","经营企业编号")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","BizopEtpsNm","戴尔（成都）有限公司","收货企业名称")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","RvsngdEtpsSccd","91510100569680768F","收发货企业社会信用代码")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","RcvgdEtpsno","5101644060","收货企业编号")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","RcvgdEtpsNm","戴尔（成都）有限公司","收货企业名称")
	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","DclEtpsSccd","91310107132231740K","申报企业社会信用代码")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","DclEtpsno","3107980011","申报企业编号")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","DclEtpsNm","上海兴亚报关有限公司","申报企业名称")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","InputCreditCode","91310107132231740K","录入单位社会信用代码")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","InputCode","3107980011","录入单位代码")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","InputName","上海兴亚报关有限公司","申报企业名称")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","ImpexpPortcd","5110","进出口口岸代码")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","DclPlcCuscd","5110","申报地关区代码")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","ImpexpMarkcd","E","进出口标记代码（I-进口 E-出口）")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","SupvModecd","5034","监管方式代码")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","TrspModecd",$Array[$index][13],"运输方式代码")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","TradeCountry",TransTargetAddCodeToNum($Array[$index][0]),"起运国/运抵国（地区）代码")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","DecType","4","报关单类型")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","Rmk",$Array[$index][1] & "  ,  " & $sBillNo  & $sCustomMemo,"备注")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","CreateFlag","1","是否已生成报关单")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","BillNo",$sBillNo,"提单号")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","ContrNo",$Array[$index][1],"合同号")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","CutMode","","征免性质")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","DistinatePort","","经停港/指运港")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","FeeCurr","","运费币制")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","FeeMark","","运费标记")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","FeeRate","","运费／率")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","GrossWet",$Array[$index][5],"毛重")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","InsurCurr","","保险费币制")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","InsurMark","","保险费标记")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","InsurRate","","保险费／率")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","LicenseNo","","许可证编号")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","NetWt",$Array[$index][6],"净重")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","OtherCurr","","杂费币制")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","OtherMark","","杂费标记")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","OtherRate","","杂费／率")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","PackNo",$Array[$index][2],"件数")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","TrafName",$aTmp[1],"运输工具代码及名称")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","TransMode","3","成交方式")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","Type","ML","单据类型：属地报关SD；备案清单：ML。LY：两单一审备案清单。CL:汇总征税报关单。SS:”属地申报，属地验放”报关单；MT:多式联运,Z:自报自缴。兼容第三方客户端，将该字段做成可选。")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","WrapType",StringFormat("%02i",$Array[$index][3]),"包装种类代码");
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","PromiseItems","999","承诺-勾选0-未选第一位：特殊关系确认第二位：价格影响确认")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","TradeAreaCode","SGP","贸易国（地区）代码（海关）");132新加坡
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","DespPortCode","","启运港代码")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","EntryPortCode","511001","入境/离境口岸代码")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","GoodsPlace","分拨仓库","货物存放地点【报检时货物的存放地点】")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","OverseasConsignorCode","","境外发货人代码")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","OverseasConsignorCname","","境外收发货人名称")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","OverseasConsignorEname","","境外发货人名称（外文）")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","OverseasConsignorAddr","","境外收发货人地址")
;~     _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","OverseasConsigneeCode","","境外收货人编码")
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecHeadType","OverseasConsigneeEname",$Array[$index][25],"境外收货人名称(外文)")

;~ ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝

;~ 	  核注单表体
	Local $iMake = 0
   For $iList = 0 To UBound($GodInfoList) - 1 Step 1
	  If $GodInfoList[$iList][0] <> $Array[$index][1] Then ContinueLoop
;~    _ArrayDisplay($GodInfoList ,$GodInfoList[$iList][0]&$GodInfoList[$iList][1])
	  Local $GName,$CodeTS,$GUnit,$FirstUnit,$TradeCurr
	  Local $INVCodeTS,$INVGmodel,$INVTradeCurr,$GdsMtno,$UsedTo
	  Local $GMode1

	  For $i = 0 To UBound($aTSCode)-1 Step 1
		 If $aTSCode[$i][8] = $Array[$index][7] Then
			$CodeTS = $aTSCode[$i][1]		;商品编码
			$INVCodeTS = $aTSCode[$i][8]	;对应核注清单表体商品序号
			$GName = $aTSCode[$i][2]		;商品名称
			$Gunit = $aTSCode[$i][3]		;成交单位
			$FirstUnit = $aTSCode[$i][4]	;法定单位
			$TradeCurr = $aTSCode[$i][5]	;成交币制
			$INVTradeCurr = $aTSCode[$i][9]	;商品核注单币制
			$GdsMtno = $aTSCode[$i][10]	;商品料号
			$UsedTo =  $aTSCode[$i][7]		;用途
			ExitLoop
		 EndIf
	  Next
	  If $Array[$index][7] = 1 OR $Array[$index][7] = 10 Then	;项号1为笔记本，10为平板电脑，要在其他栏填写操作系统，非1,则在型号栏填入
		 $GMode1 = "4|0|" & StringregExpreplace($GodInfoList[$iList][7],"([!#+^{}])","{$1}") _;配置
			& "|" & $GodInfoList[$iList][9] _;品牌
			& "|" & StringReplace($GodInfoList[$iList][5],"等","") & " " & $GodInfoList[$iList][6] _;机型 型号
			& "|||" & $GodInfoList[$iList][8];OS
	  ElseIf $Array[$index][7] = 3 Then
		 $GMode1 = "4|0|微型机|" & StringregExpreplace($GodInfoList[$iList][7],"([!#+^{}])","{$1}") _;配置
			& "|" & $GodInfoList[$iList][9] _;品牌
			& "|" & StringReplace($GodInfoList[$iList][5],"等","") & " " & $GodInfoList[$iList][6] _;机型 型号
			& "|" & $GodInfoList[$iList][8] _;OS
			& "|" & $UsedTo	;用途
		 Else
		 $GMode1 = "4|0|" & StringregExpreplace($GodInfoList[$iList][7],"([!#+^{}])","{$1}") _;配置
			& "|" & $GodInfoList[$iList][9] _;品牌
			& "|" & StringReplace($GodInfoList[$iList][5],"等","") & " " & $GodInfoList[$iList][6] _;机型 型号
			& "|" & $GodInfoList[$iList][8] _;OS
			& "|" & $UsedTo	;用途
	  EndIf
	  $INVGMode1 = StringReplace($GodInfoList[$iList][5],"等","") & " " & $GodInfoList[$iList][6]

	  $iMake += 1

	  _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage","InvtListType", "","核注单表体")
;~ 	  _XML_RemoveAttributeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","xmlns")

;~ 	  _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","SeqNo","","中心统一编号");首次导入时自动生成并返填，非首次导入须填写
	  _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","GdsSeqno",$imake,"商品序号")

	  _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","PutrecSeqno",$INVCodeTS,"对应核注清单表体商品序号")
	  _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","GdsMtno",$GdsMtno,"商品料号");*****************************
	  _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","Gdecd",$CodeTS,"商品编码")
	  _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","GdsNm",$GName,"商品名称")
	  _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","GdsSpcfModelDesc",$INVGMode1,"商品规格型号描述")
	  _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","DclUnitcd",$Gunit,"申报计量单位代码")
	  _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","LawfUnitcd",$FirstUnit,"法定计量单位代码")
	  _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","SecdLawfUnitcd","035","第二法定计量单位代码")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","Natcd",TransTargetAddCodeToNum($GodInfoList[$iList][1]),"原产国")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","DclUprcAmt",Round($GodInfoList[$iList][3]/$GodInfoList[$iList][2],4),"申报单价金额")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","DclTotalAmt",Round($GodInfoList[$iList][3],2),"申报总金额")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","UsdStatTotalAmt","","美元统计总金额")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","DclCurrcd","502","申报币制代码")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","LawfQty",$GodInfoList[$iList][2],"法定数量")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","SecdLawfQty",$GodInfoList[$iList][4],"第二法定数量")
   ;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","WtSfVal","","重量比例因子值")
   ;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","FstSfVal","","第一比例因子值")
   ;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","SecdSfVal","","第二比例因子值")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","DclQty",$GodInfoList[$iList][2],"申报数量")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","GrossWt",$Array[$index][5],"毛重量")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","NetWt",$GodInfoList[$iList][4],"净重量")
   ;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","UseCd","","用途代码")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","LvyrlfModecd","3","征减免方式代码")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","UcnsVerno","1","单耗版本号")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","EntryGdsSeqno",$imake,"报关单商品序号")
   ;~ 	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","ApplyTbSeqno","","申请表序号")
    	_XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","ClyMarkcd","0","危化品标志");0-否、1-是	  "归类标记代码"
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","Rmk","","备注")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","DestinationNatcd",TransTargetAddCodeToNum($Array[$index][0]),"最终目的国（地区）")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtListType[last()]","ModfMarkcd","3","修改标志(0-未修改 1-修改 2-删除 3-增加)")
;~ ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
;~ 		临时报关单表体

	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage","InvtDecListType","","临时报关单表体")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","SeqNo","","清单中心统一编号")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","DecSeqNo","","报关单统一编号")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","EntryGdsSeqno",$imake,"报关单商品序号")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","PutrecSeqno","@" & $imake,"对应核注清单表体商品序号组合");格式为：“@1,@2,@3”
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","Gdecd",$CodeTS,"商品编码")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","GdsNm",$GName,"商品名称")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","GdsSpcfModelDesc",$GMode1,"规格型号")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","DclUnitcd",$Gunit,"申报计量单位代码")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","LawfUnitcd",$FirstUnit,"法定计量单位代码")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","SecdLawfUnitcd","035","第二法定计量单位代码")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","DclUprcAmt",Round($GodInfoList[$iList][3]/$GodInfoList[$iList][2],4),"申报单价")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","DclTotalAmt",Round($GodInfoList[$iList][3],2),"申报总金额")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","DclCurrCd",$TradeCurr,"申报币制代码")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","NatCd",TransTargetAddCodeToNum($Array[$index][0]),"原产国（地区）")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","DestinationNatcd","CHN","最终目的国（地区）")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","LawfQty",$GodInfoList[$iList][2],"法定数量")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","SecdLawfQty",$GodInfoList[$iList][4],"第二法定数量")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","DclQty",$GodInfoList[$iList][2],"申报数量")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","UseCd","","用途代码")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","Rmk","","备注")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","LvyrlfModecd","3","征免方式")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","CiqCode","","检验检疫编码")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","DeclGoodsEname","","商品英文名称")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","OrigPlaceCode","","原产地区代码")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","Purpose","","用途代码")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","ProdValidDt","","产品有效期")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","ProdQgp","","产品保质期")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","GoodsAttr","","货物属性代码")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","Stuff","","成份/原料/组份")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","UnCode","","UN编码")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","DangName","","危险货物名称")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","DangPackType","","危包类别")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","DangPackSpec","","危包规格")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","EngManEntCnm","","境外生产企业名称")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","NoDangFlag","","非危险化学品")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","DestCode","510116","目的地代码")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","GoodsSpec","","检验检疫货物规格")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","GoodsModel","","货物型号")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","GoodsBrand","","货物品牌")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","ProduceDate","","生产日期")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","ProdBatchNo","","生产批号")
	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","DistrictCode","51016","境内目的地/境内货源地")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","CiqName","","检验检疫名称")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","MnufctrRegno","","生产单位注册号")
;~ 	   _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage/InvtDecListType[last()]","MnufctrRegName","","生产单位名称")

	Next
;~ ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage","SysId","Z7","系统id")
;~ 	子系统ID：
;~ 95---加工贸易账册系统
;~ B1---加工贸易手册系统
;~ B2---加工贸易担保管理系统
;~ B3---保税货物流转系统二期
;~ Z7---海关特殊监管区域管理系统
;~ Z8---保税物流管理系统
    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData/InvtMessage","OperCusRegCode","3107980011","操作卡的海关十位")

    _XML_MakeNode($oXmlDoc, "Signature/Object/Package/DataInfo/BussinessData","DelcareFlag",$sDelcareFlag,"0--暂存；1--申报；");""

;~     整理 XML
    _XML_LoadXML($oXmlDoc,_XML_TIDY($oXmlDoc))
;~ 	_XML_Tidy($oXmlDoc)
	_XML_SaveToFile($oXmlDoc,$xmlFullPath)
    Local $hFileOpen = FileOpen($xmlFullPath,0)
	Local $sXmlRead = FileRead($hFileOpen)
	FileClose($hFileOpen)
	$sXmlRead = StringReplace($sXmlRead,' xmlns = ""',"")
	$sXmlRead = StringReplace($sXmlRead,"{+}","+")
    Local $hFileOpen = FileOpen($xmlFullPath,2)
 	If StringLeft($sXmlRead,5) <> "<?xml" Then $sXmlRead = '<?xml version = "1.0" encoding = "utf-8"?>' & @CRLF & $sXmlRead
	FileWrite($hFileOpen,$sXmlRead)
	FileClose($hFileOpen)

   Local $try=0
   While FileGetSize($xmlFullPath) = 0 And $try < 3
	  Sleep(1000)
	  $try += 1
   WEnd

;~     下面是ZIP打包，并删除XML文件
	  zgZipFile(@ScriptDir & "\OutPut\Sas\OutBox\",$xmlFileName,".xml",true)
    If @error Then return -1
EndFunc

Func _XML_MakeNode($oxml, $path, $node, $value, $comment = "")
;~    WriteMemo(0,$path&"/"&$node,$value&"|"&$comment,False)
	_XML_CreateChildWAttr($oxml, $path, $node, Default, $value)
;~ 	If(StringLen($comment)>0) Then
	   _XML_CreateComment($oxml, $path&"/"&$node, $comment)
;~ 	EndIf
EndFunc

Func zgZipFile($path,$filename,$filesuffex="",$deleteFile=False)
   If FileGetSize($path & $filename & $filesuffex) = 0 Then Return SetError(1,3,"File size is 0")

   Local $zipcmd = '"' & @ScriptDir & '\7z.exe" a "' & $path & $filename & '.zip" "' & $path & $filename & $filesuffex & '"'
   RunWait($zipcmd,"",@SW_HIDE)
   If $deleteFile Then FileDelete($path & $filename & $filesuffex)
EndFunc

Func CountItemInArrayColumn(Const ByRef $Array,$Row,$Column)
   Local $count = 0
   For $i=0 To UBound($Array)-1 Step 1
	  If $Array[$i][$Column] = $Array[$Row][$Column] Then
		 $count += 1
	  EndIf
   Next
   Return $count
EndFunc

Func zgCheckResponse()
   If $bMakeXML Then
	  TrayTip(@ScriptName,"正在生成报文，收取回执被推迟",2)
	  Sleep(2000)
	  Return
   EndIf

   WriteMemo(0,"开始检查回执","",False)
   TrayTip(@ScriptName,"开始检查回执",2)
   Local $datafilename=_GUICtrlEdit_GetText($g_idOutputData)
   If $datafilename = "" Then
	  MsgBox($MB_ICONWARNING,"","请先选择数据文件！")
	  Return
   EndIf

    Local $oExcel = _Excel_Open(False)
   If @error Then
;~ 	  MsgBox($MB_ICONERROR,@ScriptName, "Error open EXCEL." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	  Return
   EndIf

   Local $aDataList[0][0]
   zgReadExcelToArray($aDataList,$oExcel,$datafilename,"Sheet1")
   zgReadExcelToArray($aDecList,$oExcel,$datafilename, "Sheet2")
   If UBound($aDataList) = 0 Then
	  MsgBox($MB_ICONWARNING,"","数据文件无内容，请检查！")
	  Return
   EndIf
   _Excel_Close($oExcel)
   While (UBound($aDataList,$UBOUND_COLUMNS) < 40)
	  _ArrayColInsert($aDataList, UBound($aDataList,$UBOUND_COLUMNS))
   WEnd
;~    TrayTip(@ScriptName,"检查回执报文",2)

   #Region 检查核注单暂存回执
   WriteMemo(0,"检查核注单暂存回执","",False)
   TrayTip(@ScriptName,"检查核注单暂存回执",2)
   Local $ModiFied = False
   Local $i
   For $i = 0 To UBound($aDataList)-1 Step 1
	  If $aDataList[$i][29] <> "" Then ContinueLoop
	  Local $DecInBox = @ScriptDir & "\Output\Sas\InBox\"
	  Local $FileFlag = "Successed_" & $aDataList[$i][1] & "*.xml"
	  Local $sFileName = zgGetlastFile($DecInBox,$FileFlag)
	  If zgGetSeqNo($sFileName,"/CommonResponeMessage/DealFlag") = "0" Then
		 ; And zgGetSeqNo($sFileName,"/CommonResponeMessage/CheckInfo") = "处理成功" ;2019.12.17注释，至检查DealFlag=0即可
		 $DecSeqNo = zgGetSeqNo($sFileName,"/CommonResponeMessage/SeqNo")

		 WriteMemo(0,"收到核注单暂存成功消息"&@CR&"合同号："&$aDataList[$i][1]&@CR&"  核注单暂存号："& $DecSeqNo,"",False)
		 TrayTip(@ScriptName,"收到核注单暂存成功消息"&@CR&"合同号："&$aDataList[$i][1]&@CR&"  核注单暂存号："& $DecSeqNo,2)
		 $aDataList[$i][29] = "H" & $DecSeqNo
		 $ModiFied = True
	  EndIf
   Next
   If $ModiFied Then zgSaveOutputData($datafilename,$aDataList,$aDecList)
   #EndRegion

   #Region 检查核注单申报回执
   WriteMemo(0,"检查核注单申报回执","",False)
   TrayTip(@ScriptName,"检查核注单申报回执",2)
   Local $DecInBox = @ScriptDir & "\Output\Sas\InBox\Receipt_*.INVT"
   Local $hSearch = FileFindFirstFile($DecInBox)

    ; Check if the search was successful, if not display a message and return False.
    If $hSearch = -1 Then
;~         MsgBox($MB_SYSTEMMODAL, "", "Error: No files/directories matched the search pattern.")
;~         Return False
    EndIf

    ; Assign a Local variable the empty string which will contain the files names found.
    Local $sFileName = "", $iResult = 0
    While 1
        $sFileName = FileFindNextFile($hSearch)
        ; If there is no more file matching the search.
        If @error Then ExitLoop

        ;判断是否有收到新的xml
		 $sFileName = @ScriptDir & "\Output\Sas\InBox\" & $sFileName
		 Local $InvSaveNum
		 $InvSaveNum = zgGetSeqNo($sFileName,"/Package/DataInfo/BussinessData/INV202/InvApprResult/invPreentNo")
		 Local $i,$index
		 Local $FindContract = False
		 For $i = 0 To UBound($aDataList)-1 Step 1
			If $aDataList[$i][29] = "H" & $InvSaveNum Then
			   $FindContract = True
			   $index = $i
			EndIf
		 Next

		 If Not $FindContract Then ContinueLoop

		 If  $aDataList[$index][30] = "" _
		   And zgGetSeqNo($sFileName,"/Package/DataInfo/BussinessData/INV202/InvApprResult/manageResult") = "1" Then
			$DecSeqNo = zgGetSeqNo($sFileName,"/Package/DataInfo/BussinessData/INV202/InvApprResult//businessId")
			WriteMemo(0,"收到核注单申报成功消息"&@CR&"合同号："&$aDataList[$index][1]&@CR&"  核注单编号："& $DecSeqNo,"",False)
			TrayTip(@ScriptName,"收到核注单申报成功消息"&@CR&"合同号："&$aDataList[$index][1]&@CR&"  核注单编号："& $DecSeqNo,2)

		    ;将核注单编号写回excel
			$aDataList[$index][30] = $DecSeqNo
			zgSaveOutputData($datafilename,$aDataList,$aDecList)

			;生成对应的报关单暂存报文
			If $AutoDec = "True" Then
			   If Not IsDECInputContract($aDataList[$index][1]) Then
				  If zgMakeXML($aDataList,$aDecList,$index) = -1 Then
					 WriteMemo(0,"生成失败！报关单暂存报文"&@CR&"合同号："&$aDataList[$index][1]&@CR&"  核注单编号："& $DecSeqNo,"",False)
					 TrayTip(@ScriptName,"生成失败！报关单暂存报文"&@CR&"合同号："&$aDataList[$index][1]&@CR&"  核注单编号："& $DecSeqNo,2)
				  Else
					 WriteMemo(0,"成功生成报关单暂存报文"&@CR&"合同号："&$aDataList[$index][1]&@CR&"  核注单编号："& $DecSeqNo,"",False)
					 TrayTip(@ScriptName,"成功生成报关单暂存报文"&@CR&"合同号："&$aDataList[$index][1]&@CR&"  核注单编号："& $DecSeqNo,2)
					 FileWriteLine($DECInputFileName,$aDataList[$index][1])
				  EndIf
			   EndIf
			EndIf
		 EndIf
    WEnd

    ; Close the search handle.
    FileClose($hSearch)
   #EndRegion

   #Region 检查报关单暂存回执
   WriteMemo(0,"检查报关单暂存回执","",False)
   TrayTip(@ScriptName,"检查报关单暂存回执",2)
   $ModiFied = False
   Local $i
   For $i = 0 To UBound($aDataList)-1 Step 1
	  If $aDataList[$i][31] <> "" Then ContinueLoop
	  Local $DecInBox = @ScriptDir & "\Output\Dec\InBox\"
	  Local $FileFlag = "Successed_*"&$aDataList[$i][1]&"*.xml"
	  $sFileName = zgGetLastFile($DecInBox,$FileFlag)
	  Local $DecSeqNo = ""
	  If zgGetSeqNo($sFileName,"/DecImportResponse/ResponseCode") = "0" _
		 And zgGetSeqNo($sFileName,"/DecImportResponse/ErrorMessage") = "暂存成功" Then
		 $DecSeqNo = zgGetSeqNo($sFileName,"/DecImportResponse/SeqNo")
		 If $DecSeqNo <> "" Then
			WriteMemo(0,"收到报关单暂存成功消息"&@CR&"合同号："&$aDataList[$i][1]&@CR&"  报关单暂存号："& $DecSeqNo,"",False)
			TrayTip(@ScriptName,"收到报关单暂存成功消息"&@CR&"合同号："&$aDataList[$i][1]&@CR&"  报关单暂存号："& $DecSeqNo,2)
			$aDataList[$i][31] = $DecSeqNo
			$Modified = True
		 EndIf
	  EndIf
   Next
   If $ModiFied Then zgSaveOutputData($datafilename,$aDataList,$aDecList)
   #EndRegion

   #Region 检查报关单通关回执
   WriteMemo(0,"检查报关单通关回执","",False)
   TrayTip(@ScriptName,"检查报关单通关回执",2)
   $ModiFied = False
   Local $i
   For $i = 0 To UBound($aDataList)-1 Step 1
	  If $aDataList[$i][31] = "" Then ContinueLoop	;无报关单暂存号
	  If $aDataList[$i][32] <> "" Then ContinueLoop	;已有报关单通关号
	  Local $sFileName = zgGetLastFile(@ScriptDir&"\OutPut\Dec\Inbox\","Receipt_*"&$aDataList[$i][31]&"*.xml")
	  If $sFileName = "" Then ContinueLoop
	  Local $sENTRY_ID = zgGetSeqNo($sFileName,"/DEC_RESULT/ENTRY_ID")
	  If StringStripWS($sENTRY_ID,$STR_STRIPALL) <> "" Then
		 WriteMemo(0,"收到报关单申报成功消息"&@CR&"合同号："&$aDataList[$i][1]&@CR&"  报关单正式编号："& $sENTRY_ID,"",False)
		 TrayTip(@ScriptName,"收到报关单申报成功消息"&@CR&"合同号："&$aDataList[$i][1]&@CR&"  报关单正式编号："& $sENTRY_ID,2)
		 $aDataList[$i][32] = "B" & $sENTRY_ID
		 $Modified = True
	  EndIf
   Next
   If $ModiFied Then zgSaveOutputData($datafilename,$aDataList,$aDecList)
   #EndRegion
   WriteMemo(0,"回执收取完成","",False)
   TrayTip(@ScriptName,"回执收取完成",2)
EndFunc

Func zgGetSeqNo($sFileName,$xmlItem)
    Local $sTmpFile = @TempDir & "\tmpxmlfile.xml"
    Local $hFileOpen = FileOpen($sFileName,0)
	Local $sXmlRead = ""
    $sXmlRead = FileRead($hFileOpen)
	FileClose($hFileOpen)
	$sXmlRead = StringReplace($sXmlRead," xmlns="," xmlns:ns2 = ")
	$hFileOpen = FileOpen($sTmpFile,$FO_OVERWRITE+$FO_UTF8)
	FileWrite($hFileOpen,$sXmlRead)
	FileClose($hFileOpen)

   Local $oXmlDoc = _XML_CreateDOMDocument()
   _XML_Load($oXmlDoc, $sTmpFile)
   Local $ret = _XML_GetValue($oXmlDoc,$xmlItem)
   If $ret = 0 Then
;~ 	  MsgBox(0,"","未找到XML项" & $xmlItem)
	  Return ""
   EndIf

   If @error Then Return ""
   Return $ret[1]
EndFunc

Func zgGetLastFile($sPath,$FileFlag)
   Local $rFileName = ""
   Local $hSearch = FileFindFirstFile($sPath&$FileFlag)

	   ; Check if the search was successful, if not display a message and return False.
	   If $hSearch = -1 Then
   ;~         MsgBox($MB_SYSTEMMODAL, "", "Error: No files/directories matched the search pattern.")
		   Return ""
	   EndIf

	   ; Assign a Local variable the empty string which will contain the files names found.
	   Local $sFileName = "", $LastFileTime = ""

	   While 1
		 $sFileName = FileFindNextFile($hSearch)

		 ; If there is no more file matching the search.
		 If @error Then ExitLoop

		 $sFileName = $sPath & $sFileName
		 If FileGetTime($sFileName,$FT_MODIFIED , $FT_STRING ) > $LastFileTime Then
			$LastFileTime = FileGetTime($sFileName,$FT_MODIFIED , $FT_STRING )
			$rFileName = $sFileName
		 EndIf
	   WEnd

	  ; Close the search handle.
	  FileClose($hSearch)
	  Return $rFileName

EndFunc

Func zgAttachPdfFile(ByRef $oMainDoc,$CustomsCode,$sDir,$PdfFileName,$FileType)
;~    MsgBox(0,"","zgAttachPdfFile TO DO [" &  $PdfFileName & "]")
   If NOT FileExists($sDir & $PdfFileName) Then
		 WriteMemo(0,"报关单暂存随附单据","文件不存在" & $PdfFileName,False)
		 Return
   EndIf

;~ 	   生成XML结构
   Local $tmpPDFxmlFile = @TempDir & "\XMLPdfHead.xml"

	; first you must create $oXmlDoc object
	$oXmlPdf = _XML_CreateFile($tmpPDFxmlFile,"DataInfo",True)
    _XML_SetAttrib($oXmlPdf, "DataInfo", "version", "1.0")
    _XML_SetAttrib($oXmlPdf, "DataInfo", "xmlns", "http://www.w3.org/2000/09/xmldsig#")

	_XML_MakeNode($oXmlPdf, "DataInfo","SignedDate","","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate","EnvelopInfo","","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/EnvelopInfo","Message_Type","USF001","报文类型")	;低值货物：USF001

	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/EnvelopInfo","Master_Customs_Code",$CustomsCode,"现场关区代码")	;	现场关区代码
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate","Data","","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","PRE_ENTRY_ID","","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","ENTRY_ID","","")	;报关单海关编号
	$EDocID = "CDHY_" & $CustomsCode & $FileType & "IM" & @YEAR&@MON&@MDAY&"_"&@HOUR&@MIN&@SEC&@MSEC
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","FILE_NAME",$EDocID & ".PDF","对应主单随附单据节点中edocid")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","FORMAT_TYPE","US","")	;S:结构化  US:非结构化（pdf文件填写US）
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","TRADE_CODE","3107980011","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","MASTER_CUSTOMS_CODE","","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","GROUP_ID","","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","TRADE_FILE_NAME",$PdfFileName,"原始PDF文件名")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","DECL_TYPE","F","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","DECL_TIME",@YEAR&@MON&@MDAY&" "&@HOUR&@MIN&@SEC&@MSEC,"")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","DECL_CODE","2100040041545","上传人员代码")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","DECL_NAME","胡琳","上传人员名称")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","FILE_TYPE",$FileType,"随附单据类型")	;00000001:发票 00000002:装箱单  00000003:提/运单  00000004:合同
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","FILE_SOURCE","PORT","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","FILE_DIGEST","","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","SIGN_CERT","","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","FILE_SIGN","","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/Data","OP_NOTE","","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate","SubSignedInfo","","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/SubSignedInfo","SignerID","","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/SubSignedInfo","KeyName","","")
	_XML_MakeNode($oXmlPdf, "DataInfo/SignedDate/SubSignedInfo","HashSign","","")


   _XML_LoadXML($oXmlPdf,_XML_TIDY($oXmlPdf))

 	FileDelete($tmpPDFxmlFile)
	_XML_SaveToFile($oXmlPdf,$tmpPDFxmlFile)

;~ 		读取XML文件
    Local $hFileOpen = FileOpen($tmpPDFxmlFile,0)
	Local $sXmlRead = FileRead($hFileOpen)
	FileClose($hFileOpen)
	$sXmlRead = StringReplace($sXmlRead,' xmlns=""',"")
	$sXmlRead = StringReplace($sXmlRead,"{+}","+")
	Local $iLen = StringLen($sXmlRead)
	If $iLen<2048 Then
	   $sXmlRead = $sXmlRead & _StringRepeat(" ", 2048 - $iLen)
	EndIf

    Local $hFileOpen = FileOpen($tmpPDFxmlFile,0)
	FileWrite($hFileOpen,$sXmlRead)
	FileClose($hFileOpen)
 	FileDelete($tmpPDFxmlFile)

;~ 	   PDF 文件与XML拼接
	Local $TargetFile = @ScriptDir & "\OutPut\Dec\OutBox\" & $EDocID & ".PDF"
   FileCopy($sDir & $PdfFileName,$TargetFile)

    Local $hFileOpen = FileOpen($TargetFile,1)
	FileSetPos($hFileOpen,0,$FILE_END)
	FileWrite($hFileOpen,$sXmlRead)
	FileClose($hFileOpen)

;~ 	  表体中添加随附表单文件信息
    _XML_MakeNode($oMainDoc, "DecMessage","EdocRealation","","电子随附单据关联关系信息")
    _XML_MakeNode($oMainDoc, "DecMessage/EdocRealation","EdocID",$EDocID & ".PDF","文件名");文件名:随附单据编号（文件名命名规则是：申报口岸+随附单据类别代码+IM+18位流水号+.pdf）
    _XML_MakeNode($oMainDoc, "DecMessage/EdocRealation","EdocCode",$FileType,"随附单证类别")
    _XML_MakeNode($oMainDoc, "DecMessage/EdocRealation","EdocFomatType","US","随附单据格式类型")
    _XML_MakeNode($oMainDoc, "DecMessage/EdocRealation","OpNote","3.1随附单据上传TCS","操作说明（重传原因）")
    _XML_MakeNode($oMainDoc, "DecMessage/EdocRealation","EdocCopId",$PdfFileName,"随附单据企业内部文件名")
    _XML_MakeNode($oMainDoc, "DecMessage/EdocRealation","EdocOwnerCode","3107980011","所属单位海关编号")
    _XML_MakeNode($oMainDoc, "DecMessage/EdocRealation","SignUnit","","签名单位代码")
    _XML_MakeNode($oMainDoc, "DecMessage/EdocRealation","SignTime",@YEAR&@MON&@MDAY&" "&@HOUR&@MIN&@SEC&@MSEC,"签名时间")
    _XML_MakeNode($oMainDoc, "DecMessage/EdocRealation","EdocOwnerName","上海兴亚报关有限公司","所属单位名称")
    _XML_MakeNode($oMainDoc, "DecMessage/EdocRealation","EdocSize",10240,"随附单据文件大小")

EndFunc


#Region XMLWrapperEx__Examples.au3 - XML DOM Error/Event Handling

Func ErrFunc_CustomUserHandler_MAIN($oError)
	ConsoleWrite(@ScriptName & " (" & $oError.scriptline & ") : MainScript ==> COM Error intercepted !" & @CRLF & _
			@TAB & "err.number is: " & @TAB & @TAB & "0x" & Hex($oError.number) & @CRLF & _
			@TAB & "err.windescription:" & @TAB & $oError.windescription & @CRLF & _
			@TAB & "err.description is: " & @TAB & $oError.description & @CRLF & _
			@TAB & "err.source is: " & @TAB & @TAB & $oError.source & @CRLF & _
			@TAB & "err.helpfile is: " & @TAB & $oError.helpfile & @CRLF & _
			@TAB & "err.helpcontext is: " & @TAB & $oError.helpcontext & @CRLF & _
			@TAB & "err.lastdllerror is: " & @TAB & $oError.lastdllerror & @CRLF & _
			@TAB & "err.scriptline is: " & @TAB & $oError.scriptline & @CRLF & _
			@TAB & "err.retcode is: " & @TAB & "0x" & Hex($oError.retcode) & @CRLF & @CRLF)
EndFunc		;==>ErrFunc_CustomUserHandler_MAIN

Func ErrFunc_CustomUserHandler_XML($oError)

	; here is declared another path to UDF au3 file
	; thanks to this with using _XML_ComErrorHandler_UserFunction(ErrFunc_CustomUserHandler_XML)
	;  you get errors which after pressing F4 in SciTE4AutoIt you goes directly to the specified UDF Error Line
	ConsoleWrite(@ScriptDir & '\XMLWrapperEx.au3' & " (" & $oError.scriptline & ") : UDF ==> COM Error intercepted ! " & @CRLF & _
			@TAB & "err.number is: " & @TAB & @TAB & "0x" & Hex($oError.number) & @CRLF & _
			@TAB & "err.windescription:" & @TAB & $oError.windescription & @CRLF & _
			@TAB & "err.description is: " & @TAB & $oError.description & @CRLF & _
			@TAB & "err.source is: " & @TAB & @TAB & $oError.source & @CRLF & _
			@TAB & "err.helpfile is: " & @TAB & $oError.helpfile & @CRLF & _
			@TAB & "err.helpcontext is: " & @TAB & $oError.helpcontext & @CRLF & _
			@TAB & "err.lastdllerror is: " & @TAB & $oError.lastdllerror & @CRLF & _
			@TAB & "err.scriptline is: " & @TAB & $oError.scriptline & @CRLF & _
			@TAB & "err.retcode is: " & @TAB & "0x" & Hex($oError.retcode) & @CRLF & @CRLF)
EndFunc    ;==>ErrFunc_CustomUserHandler_XML

Func XML_DOM_EVENT_ondataavailable()
	#CS
		ondataavailable Event
		https://msdn.microsoft.com/en-us/library/ms754530(v=vs.85).aspx
	#CE
	Local $oEventObj = @COM_EventObj
	ConsoleWrite('@COM_EventObj = ' & ObjName($oEventObj, 3) & @CRLF)

	Local $sMessage = 'XML_DOM_EVENT_ fired "ondataavailable"' & @CRLF
	ConsoleWrite($sMessage)
EndFunc    ;==>XML_DOM_EVENT_ondataavailable

Func XML_DOM_EVENT_onreadystatechange()
	#CS
		onreadystatechange Event
		https://msdn.microsoft.com/en-us/library/ms759186(v=vs.85).aspx
	#CE
	Local $oEventObj = @COM_EventObj
	ConsoleWrite('@COM_EventObj = ' & ObjName($oEventObj, 3) & @CRLF)

	Local $sMessage = 'XML_DOM_EVENT_ fired "onreadystatechange" : ReadyState = ' & $oEventObj.ReadyState & @CRLF
	ConsoleWrite($sMessage)

EndFunc    ;==>XML_DOM_EVENT_onreadystatechange

Func XML_DOM_EVENT_ontransformnode($oNodeCode_XSL, $oNodeData_XML, $bBool)
	#forceref $oNodeCode_XSL, $oNodeData_XML, $bBool
	#CS
		ontransformnode Event
		https://msdn.microsoft.com/en-us/library/ms767521(v=vs.85).aspx
	#CE
	Local $oEventObj = @COM_EventObj
	ConsoleWrite('@COM_EventObj = ' & ObjName($oEventObj, 3) & @CRLF)

	Local $sMessage = 'XML_DOM_EVENT_ fired "ontransformnode"' & @CRLF
	ConsoleWrite($sMessage)

EndFunc    ;==>XML_DOM_EVENT_ontransformnode

; #FUNCTION# ====================================================================================================================
; Name ..........: XML_My_ErrorParser
; Description ...: Changing $XML_ERR_ ... to human readable description
; Syntax ........: XML_My_ErrorParser($iXMLWrapper_Error, $iXMLWrapper_Extended)
; Parameters ....: $iXMLWrapper_Error	- an integer value.
;                  $iXMLWrapper_Extended           - an integer value.
; Return values .: description as string
; Author ........: mLipok
; Modified ......:
; Remarks .......: This function is only example of how user can parse @error and @extended to human readable description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func XML_My_ErrorParser($iXMLWrapper_Error, $iXMLWrapper_Extended = 0)
	Local $sErrorInfo = ''
	Switch $iXMLWrapper_Error
		Case $XML_ERR_SUCCESS
			$sErrorInfo = '$XML_ERR_OK=' & $XML_ERR_OK & @CRLF & 'All is ok.'
		Case $XML_ERR_GENERAL
			$sErrorInfo = '$XML_ERR_GENERAL=' & $XML_ERR_GENERAL & @CRLF & 'The error which is not specifically defined.'
		Case $XML_ERR_COMERROR
			$sErrorInfo = '$XML_ERR_COMERROR=' & $XML_ERR_COMERROR & @CRLF & 'COM ERROR OCCURED. Check @extended and your own error handler function for details.'
		Case $XML_ERR_ISNOTOBJECT
			$sErrorInfo = '$XML_ERR_ISNOTOBJECT=' & $XML_ERR_ISNOTOBJECT & @CRLF & 'No object passed to function'
		Case $XML_ERR_INVALIDDOMDOC
			$sErrorInfo = '$XML_ERR_INVALIDDOMDOC=' & $XML_ERR_INVALIDDOMDOC & @CRLF & 'Invalid object passed to function'
		Case $XML_ERR_INVALIDATTRIB
			$sErrorInfo = '$XML_ERR_INVALIDATTRIB=' & $XML_ERR_INVALIDATTRIB & @CRLF & 'Invalid object passed to function.'
		Case $XML_ERR_INVALIDNODETYPE
			$sErrorInfo = '$XML_ERR_INVALIDNODETYPE=' & $XML_ERR_INVALIDNODETYPE & @CRLF & 'Invalid object passed to function.'
		Case $XML_ERR_OBJCREATE
			$sErrorInfo = '$XML_ERR_OBJCREATE=' & $XML_ERR_OBJCREATE & @CRLF & 'Object can not be created.'
		Case $XML_ERR_NODECREATE
			$sErrorInfo = '$XML_ERR_NODECREATE=' & $XML_ERR_NODECREATE & @CRLF & 'Can not create Node - check also COM Error Handler'
		Case $XML_ERR_NODEAPPEND
			$sErrorInfo = '$XML_ERR_NODEAPPEND=' & $XML_ERR_NODEAPPEND & @CRLF & 'Can not append Node - check also COM Error Handler'
		Case $XML_ERR_PARSE
			$sErrorInfo = '$XML_ERR_PARSE=' & $XML_ERR_PARSE & @CRLF & 'Error: with Parsing objects, .parseError.errorCode=' & $iXMLWrapper_Extended & ' Use _XML_ErrorParser_GetDescription() for get details.'
		Case $XML_ERR_PARSE_XSL
			$sErrorInfo = '$XML_ERR_PARSE_XSL=' & $XML_ERR_PARSE_XSL & @CRLF & 'Error with Parsing XSL objects .parseError.errorCode=' & $iXMLWrapper_Extended & ' Use _XML_ErrorParser_GetDescription() for get details.'
		Case $XML_ERR_LOAD
			$sErrorInfo = '$XML_ERR_LOAD=' & $XML_ERR_LOAD & @CRLF & 'Error opening specified file.'
		Case $XML_ERR_SAVE
			$sErrorInfo = '$XML_ERR_SAVE=' & $XML_ERR_SAVE & @CRLF & 'Error saving file.'
		Case $XML_ERR_PARAMETER
			$sErrorInfo = '$XML_ERR_PARAMETER=' & $XML_ERR_PARAMETER & @CRLF & 'Wrong parameter passed to function.'
		Case $XML_ERR_ARRAY
			$sErrorInfo = '$XML_ERR_ARRAY=' & $XML_ERR_ARRAY & @CRLF & 'Wrong array parameter passed to function. Check array dimension and conent.'
		Case $XML_ERR_XPATH
			$sErrorInfo = '$XML_ERR_XPATH=' & $XML_ERR_XPATH & @CRLF & 'XPath syntax error - check also COM Error Handler.'
		Case $XML_ERR_NONODESMATCH
			$sErrorInfo = '$XML_ERR_NONODESMATCH=' & $XML_ERR_NONODESMATCH & @CRLF & 'No nodes match the XPath expression'
		Case $XML_ERR_NOCHILDMATCH
			$sErrorInfo = '$XML_ERR_NOCHILDMATCH=' & $XML_ERR_NOCHILDMATCH & @CRLF & 'There is no Child in nodes matched by XPath expression.'
		Case $XML_ERR_NOATTRMATCH
			$sErrorInfo = '$XML_ERR_NOATTRMATCH=' & $XML_ERR_NOATTRMATCH & @CRLF & 'There is no such attribute in selected node.'
		Case $XML_ERR_DOMVERSION
			$sErrorInfo = '$XML_ERR_DOMVERSION=' & $XML_ERR_DOMVERSION & @CRLF & 'DOM Version: ' & 'MSXML Version ' & $iXMLWrapper_Extended & ' or greater required for this function'
		Case $XML_ERR_EMPTYCOLLECTION
			$sErrorInfo = '$XML_ERR_EMPTYCOLLECTION=' & $XML_ERR_EMPTYCOLLECTION & @CRLF & 'Collections of objects was empty'
		Case $XML_ERR_EMPTYOBJECT
			$sErrorInfo = '$XML_ERR_EMPTYOBJECT=' & $XML_ERR_EMPTYOBJECT & @CRLF & 'Object is empty'
		Case Else
			$sErrorInfo = '=' & $iXMLWrapper_Error & @CRLF & 'NO ERROR DESCRIPTION FOR THIS @error'
	EndSwitch

	Local $sExtendedInfo = ''
	Switch $iXMLWrapper_Error
		Case $XML_ERR_COMERROR, $XML_ERR_NODEAPPEND, $XML_ERR_NODECREATE
			$sExtendedInfo = 'COM ERROR NUMBER (@error returned via @extended) =' & $iXMLWrapper_Extended
		Case $XML_ERR_PARAMETER
			$sExtendedInfo = 'This @error was fired by parameter: #' & $iXMLWrapper_Extended
		Case Else
			Switch $iXMLWrapper_Extended
				Case $XML_EXT_DEFAULT
					$sExtendedInfo = '$XML_EXT_DEFAULT=' & $XML_EXT_DEFAULT & @CRLF & 'Default - Do not return any additional information'
				Case $XML_EXT_XMLDOM
					$sExtendedInfo = '$XML_EXT_XMLDOM=' & $XML_EXT_XMLDOM & @CRLF & '"Microsoft.XMLDOM" related Error'
				Case $XML_EXT_DOMDOCUMENT
					$sExtendedInfo = '$XML_EXT_DOMDOCUMENT=' & $XML_EXT_DOMDOCUMENT & @CRLF & '"Msxml2.DOMDocument" related Error'
				Case $XML_EXT_XSLTEMPLATE
					$sExtendedInfo = '$XML_EXT_XSLTEMPLATE=' & $XML_EXT_XSLTEMPLATE & @CRLF & '"Msxml2.XSLTemplate" related Error'
				Case $XML_EXT_SAXXMLREADER
					$sExtendedInfo = '$XML_EXT_SAXXMLREADER=' & $XML_EXT_SAXXMLREADER & @CRLF & '"MSXML2.SAXXMLReader" related Error'
				Case $XML_EXT_MXXMLWRITER
					$sExtendedInfo = '$XML_EXT_MXXMLWRITER=' & $XML_EXT_MXXMLWRITER & @CRLF & '"MSXML2.MXXMLWriter" related Error'
				Case $XML_EXT_FREETHREADEDDOMDOCUMENT
					$sExtendedInfo = '$XML_EXT_FREETHREADEDDOMDOCUMENT=' & $XML_EXT_FREETHREADEDDOMDOCUMENT & @CRLF & '"Msxml2.FreeThreadedDOMDocument" related Error'
				Case $XML_EXT_XMLSCHEMACACHE
					$sExtendedInfo = '$XML_EXT_XMLSCHEMACACHE=' & $XML_EXT_XMLSCHEMACACHE & @CRLF & '"Msxml2.XMLSchemaCache." related Error'
				Case $XML_EXT_STREAM
					$sExtendedInfo = '$XML_EXT_STREAM=' & $XML_EXT_STREAM & @CRLF & '"ADODB.STREAM" related Error'
				Case $XML_EXT_ENCODING
					$sExtendedInfo = '$XML_EXT_ENCODING=' & $XML_EXT_ENCODING & @CRLF & 'Encoding related Error'
				Case Else
					$sExtendedInfo = '$iXMLWrapper_Extended=' & $iXMLWrapper_Extended & @CRLF & 'NO ERROR DESCRIPTION FOR THIS @extened'
			EndSwitch
	EndSwitch
	; return back @error and @extended for further debuging
	Return SetError($iXMLWrapper_Error, $iXMLWrapper_Extended, _
			'@error description:' & @CRLF & _
			$sErrorInfo & @CRLF & _
			@CRLF & _
			'@extended description:' & @CRLF & _
			$sExtendedInfo & @CRLF & _
			'')

EndFunc    ;==>XML_My_ErrorParser
#EndRegion  XMLWrapperEx__Examples.au3 - XML DOM Error/Event Handling

#Region XMLWrapperEx__Examples.au3 - Help, Documentation, Manuals
#CS
	How Do I Use XML?
	https://msdn.microsoft.com/en-us/library/ms759092(v=vs.85).aspx
	A Beginner's Guide to the XML DOM
	https://msdn.microsoft.com/en-us/library/aa468547.aspx
	DOM Reference
	https://msdn.microsoft.com/en-us/library/ms764730(v=vs.85).aspx
	XML DOM Properties
	https://msdn.microsoft.com/en-us/library/ms763798(v=vs.85).aspx
	XML DOM Methods
	https://msdn.microsoft.com/en-us/library/ms757828(v=vs.85).aspx
	XML Glossary
	https://msdn.microsoft.com/en-us/library/ms256452(v=vs.85).aspx
	MSXML API History
	https://msdn.microsoft.com/en-us/library/ms762314(v=vs.85).aspx
	XML Standards Reference
	https://msdn.microsoft.com/en-us/library/ms256177(v=vs.85).aspx
	XSLT Reference
	https://msdn.microsoft.com/en-us/library/ms256069(v=vs.85).aspx
	XPath Reference
	https://msdn.microsoft.com/en-us/library/ms256115(v=vs.85).aspx
	XPath Syntax
	https://msdn.microsoft.com/en-us/library/ms256471(v=vs.85).aspx
	XPath Examples
	https://msdn.microsoft.com/en-us/library/ms256086(v=vs.110).aspx
	XML DOM Objects/Interfaces
	https://msdn.microsoft.com/en-us/library/ms760218(v=vs.85).aspx
	Location Path Examples
	https://msdn.microsoft.com/en-us/library/ms256236(v=vs.110).aspx
	Microsoft XML DOM output has no CR/LF
	http://www.experts-exchange.com/Web_Development/Web_Languages-Standards/Q_20135543.html
	Formatting of XML file
	http://www.visualbasicscript.com/Formatting-of-XML-file-m77414.aspx
	Stream Object (ADO)
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms675032(v=vs.85).aspx
	TODO:
	IXMLDOMAttribute Members  ($oAttribute)
	https://msdn.microsoft.com/en-us/library/ms767677(v=vs.85).aspx
	XML Editing: A WYSIWYG XML Document Editor
	https://msdn.microsoft.com/en-us/library/ms977865.aspx
	Understanding XML Namespaces
	https://msdn.microsoft.com/en-us/library/aa468565.aspx
	Understanding XML
	https://msdn.microsoft.com/en-us/library/aa468558.aspx
#CE
#CS ON LINE XPath Tools
	http://www.xpathtester.com/xpath
	http://codebeautify.org/Xpath-Tester
	http://xpath.online-toolz.com/tools/xpath-editor.php
	http://www.whitebeam.org/library/guide/TechNotes/xpathtestbed.rhtm
	http://www.qutoric.com/xslt/analyser/xpathtool.html
	https://addons.mozilla.org/en-US/firefox/addon/xpather/
	http://www.xmlme.com/XpathTool.aspx
#CE
#CS OFF LINE XPath Tools
	http://qutoric.com/xmlquire/
	https://xpathvisualizer.codeplex.com/
#CE
#EndRegion  XMLWrapperEx__Examples.au3 - Help, Documentation, Manuals
