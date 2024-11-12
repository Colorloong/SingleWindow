Func zgAuthorization()
   Local $AllSerials="812202004,419818,512940,3707357416,1153680252,1219813213,15742245,470926915,3875245373,787799,785030" & _
					 "697601,697617,808537517,137841"
   Local $sSerial=DriveGetSerial(@HomeDrive & "\")
   If StringInStr($AllSerials,$sSerial)>0 Then Return True
   Return False
EndFunc