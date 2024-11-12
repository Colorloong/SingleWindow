#include <MsgBoxConstants.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <Date.au3>

;~ Local $aRet,$row,$col
;~ mySQLCreateDB("SingleWindow.db","INVSave")
;~ mySQLExec("SingleWindow.db","Delete From INVSave;")
;~ mySQLExec("SingleWindow.db","INSERT INTO INVSave('ContractNo','SaveDate','Flag') VALUES ('NPD1983','" & _NowCalcDate() & "','1');")
;~ Local $row = mySQLQueryCount("SingleWindow.db","SELECT count(*) FROM INVSave WHERE ContractNo='"&"NPD1983"&"';")
;~ ConsoleWrite(stringformat("get return %i\n",$row))
;~ _ArrayDisplay($array)
;~ Exit

Func mySQLCreateDB($sDbName,$TableName)
   _SQLite_Startup()
   Local $hDskDb = _SQLite_Open($sDbName) ; Open a permanent disk database
   If @error Then
	   MsgBox($MB_SYSTEMMODAL, "SQLite Error", "Can't open or create a permanent Database!")
	   Return -1
	EndIf

   _SQLite_Exec($hDskDb,"Drop Table INVSave;")
   _SQLite_Exec($hDskDb,"Create table "&$TableName&"(ContractNo TEXT PRIMARY KEY ASC,SaveDate DATETIME,Flag TEXT);")
;~    _SQLite_Exec($hDskDb, "INSERT INTO INVSave('ContractNo','SaveDate','Flag') VALUES ('NPD1980','" & _NowCalcDate() & "','1');") ; INSERT Data
;~    _SQLite_Exec($hDskDb, "INSERT INTO INVSave('ContractNo','SaveDate','Flag') VALUES ('NPD1981','" & _NowCalcDate() & "','1');") ; INSERT Data
;~    _SQLite_Exec($hDskDb, "INSERT INTO INVSave('ContractNo','SaveDate','Flag') VALUES ('NPD1982','" & _NowCalcDate() & "','1');") ; INSERT Data

   If @error Then
	   MsgBox($MB_SYSTEMMODAL, "SQLite Error", "Can't create a permanent Database!" & @CRLF & "Error Code: " & _SQLite_ErrCode() & @CRLF & "Error Message: " & _SQLite_ErrMsg())
	   Return -1
	EndIf
   ConsoleWrite(StringFormat("DataBase is Created:[%s],Table Name [%s],_SQLite_LibVersion=%s%s" ,$sDbName,$TableName,_SQLite_LibVersion(),@CRLF))

   _SQLite_Close($sDbName)
   _SQLite_Shutdown()
EndFunc

Func mySQLExec($sDbName,$sqlComment)
   _SQLite_Startup()
   Local $hDskDb = _SQLite_Open($sDbName) ; Open a permanent disk database
   If @error Then
	   MsgBox($MB_SYSTEMMODAL, "SQLite Error", "Can't open or create a permanent Database!")
	   Return -1
	EndIf

   Local $ret = _SQLite_Exec($hDskDb,$sqlComment)
   If $ret <> $SQLITE_OK Then
	   MsgBox($MB_SYSTEMMODAL, "SQLite Error", "SQL Exec error!" & @CRLF & "Error Code: " & _SQLite_ErrCode() & @CRLF & "Error Message: " & _SQLite_ErrMsg())
	   Return -1
	EndIf

   _SQLite_Close($sDbName)
   _SQLite_Shutdown()
EndFunc

Func mySQLQuery($sDbName,$sqlComment,ByRef $aRet,ByRef $iRow,ByRef $iCol)
   _SQLite_Startup()

   _SQLite_Open($sDbName) ; Open a permanent disk database
   If @error Then
	   MsgBox($MB_SYSTEMMODAL, "SQLite Error", "Can't open or create a permanent Database!")
	   Return -1
	EndIf

   Local $aRow
   Local $iRval = _SQLite_GetTable2d(-1, $sqlComment, $aRet, $iRow, $iCol)
   If $iRval <> $SQLITE_OK Then
	   MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
	   Return -1
   EndIf

   _SQLite_Close($sDbName)
   _SQLite_Shutdown()
EndFunc
