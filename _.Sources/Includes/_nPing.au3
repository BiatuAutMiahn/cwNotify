#include-once
#include "advping.au3"
;#include "_Common.au3"

#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         BiatuAutMiahn[@outlook.com]

This function makes use of Windows' internal IcmpSendEcho API, just like the ping.exe utility.
The key difference is that ping.exe uses resources and adds neccessary overhead to a simple call.
It also takes at lease 1 second to return. With the direct API call we can easily return in 250ms.
In the OhioHealth network most ping times are less than 50ms, so we shouldnt have to wait 1000ms+
to check if a host is online.

#ce ----------------------------------------------------------------------------

Func _isAlive($sHost,$iTimeout=250)
    Local $vRet=_nPing($sHost,$iTimeout)
    Return $vRet[1] ? True : False
EndFunc

Func _nPing($sDest,$iTimeout=250)
	Local $vRet[3]
	Local $ECHORet
	$vRet[0]=1
	Local $bRet = _IcmpCustomPayload($sDest, "abcdefghijklmnopqrstuvwabcdefghi", $ECHORet, $iTimeout)
    ;ConsoleWrite("icmp:"&$sDest&':'&_IcmpCustomPayload($sDest, "abcdefghijklmnopqrstuvwabcdefghi", $ECHORet, $iTimeout)&@CRLF)
	If @error Then
	  $vRet[1]=0
	  $vRet[2]=null
	  Switch @error
	   Case $IP_REQ_TIMED_OUT
		Return SetError(1, 1,$vRet)
	   Case $IP_DEST_HOST_UNREACHABLE
		Return SetError(1, 2,$vRet)
	   Case $IP_BAD_DESTINATION
		Return SetError(1, 3,$vRet)
	   Case Else
		Return SetError(1, 4,$vRet)
	  EndSwitch
	Else
	  $vRet[1]=1
	  $vRet[2]=DllStructGetData($ECHORet, "RoundTripTime")
	  Return $vRet
	EndIf
EndFunc
