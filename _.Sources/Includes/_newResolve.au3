#include-once
#include <Array.au3>
#include "Network.au3"
#include "_nPing.au3"
#include "_Dig.au3"

#cs ----------------------------------------------------------------------------
  AutoIt Version: 3.3.14.5
  Author:         BiatuAutMiahn[@outlook.com]

    The current resolver functionality works, but has a big flaw. The DNS server itself does not update
immediately. And with the current resolver only considering a single response from the DNS server, and
only querying a single DNS server. We often run into issues resolving hosts that switch between WLAN
and LAN connections. This script will aim to rewrite _Dig and _Resolve to a _NewResolve that will
look on all DNS servers and consider all replys.

#ce ----------------------------------------------------------------------------

; Need to detect when adapter states change and update availabel DNS servers.

Global $g_aResolveSrvDNS[0][2]
Global $g_aResolveQueryDNS[0]
Global $g_aResolveAdapters
;_nResolveUpdateInfoDNS()
;_ArrayDisplay()

Func _nResolve($sQuery,$bReverse=False,$bFullSearch=False,$iTimeout=1000)
    Local $aResolve[0][2], $sHost, $aAnswer, $sDigQuery, $bHaveResolve=False
    If UBound($g_aResolveQueryDNS,1)==0 Or UBound($g_aResolveSrvDNS,1)==0 Then
        _nResolveUpdateInfoDNS()
        If UBound($g_aResolveSrvDNS,1)==0 Then Return SetError(1,0,False)
        If UBound($g_aResolveQueryDNS,1)==0 Then Return SetError(2,0,False)
    EndIf
    For $i=0 To UBound($g_aResolveSrvDNS,1)-1
        ; is DNS Server Online?
        If not $g_aResolveSrvDNS[$i][1] Then
            ; If all available DNS servers are offline, then Return Error
            If UBound($g_aResolveSrvDNS,1)-1==$i Then Return SetError(3,0,False)
            ContinueLoop
        EndIf
        ; Query each DNS Suffix
        If Not $bReverse Then
            For $j=0 To UBound($g_aResolveQueryDNS,1)-1
                If StringInStr($sQuery,'.'&$g_aResolveQueryDNS[$j]) Then $sQuery=StringReplace($sQuery,'.'&$g_aResolveQueryDNS[$j],"")
                If StringRight($sQuery,1)=='.' Then $sQuery=StringTrimRight($sQuery,1)
            Next
        EndIf
        For $j=0 To UBound($g_aResolveQueryDNS,1)-1
            ;If not performing full search and we already have a result, dont query next suffix.
            If Not $bFullSearch And $bHaveResolve Then ExitLoop
            If $bReverse Then
                ; if performing a reverse query, we need to look for an ARP PTR record.
                $sDigQuery=StringRegExpReplace($sQuery,"([0-9]+)\.([0-9]+)\.([0-9]+)\.([0-9]+)","\4\.\3\.\2\.\1") & ".in-addr.arpa"
            Else
                ; Append suffix for forward search.
                $sDigQuery=$sQuery&'.'&$g_aResolveQueryDNS[$j]
            EndIf
            $sbRev=$bReverse ? "PTR" : "A"
            ConsoleWrite('nDig('&$sDigQuery&','&$g_aResolveSrvDNS[$i][0]&','&$sbRev&','&$iTimeout&')'&@CRLF)
            $aRet=_nDig($sDigQuery,$g_aResolveSrvDNS[$i][0],$bReverse ? "PTR" : "A", $iTimeout)
            If @error Then
                ContinueLoop
            EndIf
            For $k=0 To UBound($aRet,1)-1
                If $aRet[$k][0]<>'A' Then ContinueLoop
                If $bReverse Then
                    $sHost=$sQuery
                    If StringRight($aRet[$k][5],1)=='.' Then $aRet[$k][5]=StringTrimRight($aRet[$k][5],1)
                Else
                    If StringRight($aRet[$k][1],1)=='.' Then $aRet[$k][1]=StringTrimRight($aRet[$k][1],1)
                    $sHost=$aRet[$k][1]
                EndIf
                $iMax=UBound($aResolve,1)
                If $iMax>0 Then
                    For $l=0 To $iMax-1
                        If $sHost==$aResolve[$l][0] Then
                            If $aRet[$k][5]==$aResolve[$l][1] Then ContinueLoop 2
                            If IsArray($aResolve[$l][1]) Then
                                $aAnswer=$aResolve[$l][1]
                                $iMaxA=UBound($aAnswer,1)
                                ReDim $aAnswer[$iMaxA+1]
                                $aAnswer[$iMaxA]=$aRet[$k][5]
                            Else
                                Dim $aAnswer[2]
                                $aAnswer[0]=$aResolve[$l][1]
                                $aAnswer[1]=$aRet[$k][5]
                            EndIf
                            $aResolve[$l][1]=$aAnswer
                            ContinueLoop 2
                        EndIf
                    Next
                EndIf
                ReDim $aResolve[$iMax+1][2]
                $aResolve[$iMax][0]=$sHost
                $aResolve[$iMax][1]=$aRet[$k][5]
                $bHaveResolve=True
            Next
        Next
    Next
    If UBound($aResolve,1)==0 Then
        Return SetError(4,0,False)
    EndIf
    Return SetError(0,0,$aResolve)
EndFunc

; Get DNS Suffix Search Order, Available DNS Servers, and check if DNS servers are pingable.
Func _nResolveUpdateInfoDNS()
    $g_aResolveAdapters=_nResolveGetAdapterInfo()
    ;_nResolveUpdateAStates()
    Local $iMax, $bNew=True, $aSearchDNS
    Dim $g_aResolveQueryDNS[0]
    Dim $g_aResolveSrvDNS[0][2]
    For $i=0 To UBound($g_aResolveAdapters)-1
        ; If the adapter is not connected or enabled then skip it.
        ; Get DNS Search Order.
        $aSearchDNS=StringSplit($g_aResolveAdapters[$i][0],',')
        For $j=1 To $aSearchDNS[0]
            $iMax=UBound($g_aResolveQueryDNS,1)
            If $iMax==0 Then
                For $k=0 To $iMax-1
                    If $g_aResolveQueryDNS[$i]==$aSearchDNS[$j] Then ContinueLoop 2
                Next
            EndIf
            ReDim $g_aResolveQueryDNS[$iMax+1]
            $g_aResolveQueryDNS[$iMax]=$aSearchDNS[$j]
        Next
        $aAddrDNS=StringSplit($g_aResolveAdapters[$i][1],',')
        For $j=1 To $aAddrDNS[0]
            $iMax=UBound($g_aResolveSrvDNS,1)
            If $iMax==0 Then
                For $k=0 To $iMax-1
                    If $g_aResolveSrvDNS[$i][0]==$aAddrDNS[$j] Then ContinueLoop 2
                Next
            EndIf
            ReDim $g_aResolveSrvDNS[$iMax+1][2]
            $g_aResolveSrvDNS[$iMax][0]=$aAddrDNS[$j]
            $g_aResolveSrvDNS[$iMax][1]=_isAlive($aAddrDNS[$j])
        Next
    Next
EndFunc

; =======================================
; 3rd party Funcs
; =======================================
;
; #FUNCTION# ====================================================================================================================
; Name...........: _GetNetworkAdapterInfos
; Author.........: JGUINCH
; Source.........: network.au3
; Modified By....: BiatuAutMiahn
; Description....: Retrieve informations for the specified network card.
;                  If no network adapter is specified (default), the function returns informations for all network adapters.
; Syntax.........: _GetNetworkAdapterInfos([$sNetAdapter])
; Parameters.....: $sNetAdapter        - Name of the network adapter or network ID
;                                        The Windows network connection name can be used instead of network adapter.
; Return values..: Success  - Returns a 2 dimensional array containing informations about the adapter configuration :
;                   - element[n][0] = AdapterType                    - Network adapter type.
;                       "Ethernet 802.3"
;                       "Token Ring 802.5"
;                       "Fiber Distributed Data Interface (FDDI)"
;                       "Wide Area Network (WAN)"
;                       "LocalTalk"
;                       "Ethernet using DIX header format"
;                       "ARCNET"
;                       "ARCNET (878.2)"
;                       "ATM"
;                       "Wireless"
;                       "Infrared Wireless"
;                       "Bpc"
;                       "CoWan"
;                       "1394"
;                   - element[n][1] = MACAddress                     - MAC address for this network adapter.
;                   - element[n][2] = Name                           - Label by which the object is known.
;                   - element[n][3] = NetConnectionStatus            - State of the network adapter connection to the network.
;                       0 (0x0)  Disconnected
;                       1 (0x1)  Connecting
;                       2 (0x2)  Connected
;                       3 (0x3)  Disconnecting
;                       4 (0x4)  Hardware not present
;                       5 (0x5)  Hardware disabled
;                       6 (0x6)  Hardware malfunction
;                       7 (0x7)  Media disconnected
;                       8 (0x8)  Authenticating
;                       9 (0x9)  Authentication succeeded
;                       10 (0xA)  Authentication failed
;                       11 (0xB)  Invalid address
;                       12 (0xC)  Credentials required
;                   - element[n][4] = NetEnabled                    - Indicates whether the adapter is enabled or not.
;                   - element[n][5] = DNSDomain                     - Organization name followed by a period and an extension that indicates the type of organization.
;                   - element[n][6] = DNSDomainSuffixSearchOrder    - List of DNS domain suffixes to be appended to the end of host names during name resolution (comma-separated values).
;                   - element[n][7] = DNSServerSearchOrder          - List of server IP addresses to be used in querying for DNS servers (comma-separated values).
;                  Failure  - 0
; ===============================================================================================================================
Func _nResolveGetAdapterInfo($sNetAdapter = "")
    ;element[n][7]=="PANGP Virtual Ethernet Adapter"
    ;element[n][9]==2 (0x2)  Connected
    ;element[n][10]==True
	Local $objWMIService = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
	If $objWMIService = 0 Then Return 0
	Local $aAdaptersList = _GetNetworkAdapterList()
	If Not IsArray($aAdaptersList) Then Return 0
    ;_ArrayDisplay($aAdaptersList)
	Local $filter, $aInfos[1][1], $sQueryNetAdapterConfig, $colNetAdapterConfig, $objNetAdapter, $objNetAdapterConfig, $DeviceID, $n = 0, $adapterIndex
	;Local $sAdapterName = _GetNetworkAdapterFromID($sNetAdapter)
	;If $sAdapterName Then $sNetAdapter = $sAdapterName
    For $i = 0 To UBound($aAdaptersList) - 1
        $filter &= " OR Description = '" & $aAdaptersList[$i][0] & "'"
    Next
    $filter = StringTrimLeft($filter, 3)
	Local $objWMIService = ObjGet("winmgmts:\\"&@ComputerName&"\root\CIMV2")
	If $objWMIService = 0 Then Return 0
	Local $sQueryNetAdapter = 'select * from Win32_NetworkAdapter WHERE NetConnectionStatus = 2 And NetEnabled = True and ('&$filter&')'
	Local $colNetAdapter = $objWMIService.ExecQuery($sQueryNetAdapter, "WQL", $wbemFlagReturnImmediately)
	If NOT IsObj($colNetAdapter) Then Return 0
	For $objNetAdapter In $colNetAdapter
		$adapterIndex = $objNetAdapter.Index
		$sQueryNetAdapterConfig = "select * from Win32_NetworkAdapterConfiguration Where Index = " & $adapterIndex
		$colNetAdapterConfig = $objWMIService.ExecQuery($sQueryNetAdapterConfig, "WQL", $wbemFlagReturnImmediately)
		If IsObj($colNetAdapterConfig) Then
			For $objNetAdapterConfig In $colNetAdapterConfig
				$n += 1
				Redim $aInfos[$n][2]
				$aInfos[$n-1][0] = _Array2String(($objNetAdapterConfig.DNSDomainSuffixSearchOrder),",")
				$aInfos[$n-1][1] = _Array2String(($objNetAdapterConfig.DNSServerSearchOrder),",")
                For $m=0 To 1
                    $aInfos[$n-1][$m]=StringReplace($aInfos[$n-1][$m],Chr(0),'')
                Next
			Next
		EndIf
	Next
	If $n = 0 Then Return 0
	Return $aInfos
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _nDig
; Modified By....: BiatuAutMiahn
; Description ...: Queries a DNS server and returns the result in a 'dig' like format (not exactly, but near)
;                  to be parsed by yourself :-)
; Syntax.........: _nDig($sDig_domain[, $sDig_server = ""[, $iDig_port = 53[, $sDig_type = "A"[, $sDig_class = "IN" _
;                     [, $sDig_proto = "UDP"[, $sDig_timeout = 1000]]]]]])
; Parameters ....: $sDig_domain  - Domain to query
;                  $sDig_server  - DNS server to query
;                  $sDig_type    - [optional] ressource record type to be queried (default: "A")
;                  $sDig_timeout - [optional] Timeout in seconds to wait for a response (default: 1)
; Return values .: Success - a string formatted similar to a 'dig' query
;                  Failure - "", sets @error:
;                  |1 - something went wrong (sorry, no specific error messages at the moment)
; Author ........: Andreas Börner (mail@andreas-boerner.de)
; Modified.......: Andreas Börner (mail@andreas-boerner.de)
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; No
; ===============================================================================================================================
Func _nDig($sDig_domain,$sDig_server,$sDig_type=Default,$sDig_timeout=Default)
    Local $iDig_port=53
    Local $sDig_proto="UDP"
    Local $sDig_class="IN"
    Local $iErr
    if $sDig_type=Default Then $sDig_type="A"
    if $sDig_timeout=Default Then $sDig_timeout=1000
    Local $dig_id = Hex(Random(0, 65535, 1), 4)
    Local $dig_flags="0100"
    $dig_counters="0001000000000000"
    Local $sDig_domain_binary=EncodeName($sDig_domain)
    Local $sDig_type_binary = EncodeType($sDig_type)
    Local $sDig_class_binary = EncodeClass($sDig_class)
    Local $dig_request = $dig_id & $dig_flags & $dig_counters & $sDig_domain_binary & $sDig_type_binary & $sDig_class_binary
    $dig_request="0x" & $dig_request
    Local $dig_sock
    UDPStartup()
    $dig_sock = UDPOpen($sDig_server, $iDig_port)
    If @error Then
        $iErr=@Error
        UDPShutdown()
        Return SetError(1,$iErr,False)
    EndIf
    Local $query_time=TimerInit()
    UDPSend($dig_sock, $dig_request)
    Local $tik = 0
    Do
        $bDig_amsg = UDPRecv($dig_sock, 512,1)
        Sleep(100)
    Until $bDig_amsg <> "" Or TimerDiff($query_time)>=$sDig_timeout
    If TimerDiff($query_time)>=$sDig_timeout Then Return SetError(2,@error)
    $query_time=Round(TimerDiff($query_time))
    UDPShutdown()
    If $bDig_amsg = "" Then Return SetError(3,0,False)
    If StringMid(BinaryMid($bDig_amsg, 1, 2), 3) <> $dig_id Then Return SetError(4,0,False)
    $iDig_ptr=1
    $dig_id=ReadHex2Int(2)
    $dig_flags=ReadHex2Int(2)
    $iDig_q_count=ReadHex2Int(2)
    $iDig_a_count=ReadHex2Int(2)
    $iDig_au_count=ReadHex2Int(2)
    $iDig_ar_count=ReadHex2Int(2)
    $dig_flags_flags=""
    $dig_flags_opcode=BitShift(BitAND($dig_flags,30720),11)
    $dig_flags_rcode=BitAND($dig_flags,15)
	Local $aRet[0][6]
    if $iDig_q_count>0 Then
        $vRet=_nReadResourceRecords("q")
        for $i = 0 to UBound($vRet,1)-1
            $iMax=UBound($aRet,1)
            ReDim $aRet[$iMax+1][6]
            $aRet[$iMax][0]="Q"
            For $j=1 To 5
                $aRet[$iMax][$j]=$vRet[$i][$j-1]
            Next
        Next
;~         _ArrayDisplay($vRet)
;~         _ArrayDisplay($aRet)
    EndIf
    if $iDig_a_count>0 Then
        $vRet=_nReadResourceRecords("a")
        for $i = 0 to UBound($vRet,1)-1
            $iMax=UBound($aRet,1)
            ReDim $aRet[$iMax+1][6]
            $aRet[$iMax][0]="A"
            For $j=1 To 5
                $aRet[$iMax][$j]=$vRet[$i][$j-1]
            Next
        Next

    EndIf
    if $iDig_au_count>0 Then
        $vRet=_nReadResourceRecords("au")
        for $i = 0 to UBound($vRet,1)-1
            $iMax=UBound($aRet,1)
            ReDim $aRet[$iMax+1][6]
            $aRet[$iMax][0]="AU"
            For $j=1 To 5
                $aRet[$iMax][$j]=$vRet[$i][$j-1]
            Next
        Next
    EndIf
    if $iDig_ar_count>0 Then
        $vRet=_nReadResourceRecords("ar")
        for $i = 0 to UBound($vRet,1)-1
            $iMax=UBound($aRet,1)
            ReDim $aRet[$iMax+1][6]
            $aRet[$iMax][0]="AR"
            For $j=1 To 5
                $aRet[$iMax][$j]=$vRet[$i][$j-1]
            Next
        Next
    EndIf
    return SetError(0,0,$aRet)
EndFunc

; read a ressource record section from the response message
Func _nReadResourceRecords($section_id)
    Local $i,$count
    Local $name_dec,$type_dec,$class_dec,$ttl_dec,$rd_len,$data_dec
    Switch $section_id
        case "q"
            $count=$iDig_q_count
        case "a"
            $count=$iDig_a_count
        case "au"
            $count=$iDig_au_count
        case "ar"
            $count=$iDig_ar_count
        case Else
            Return SetError(1)
    EndSwitch
	Local $aRet[0][5]
    for $i=1 to $count
        $name_dec=DecodeName()
        $type_dec=DecodeType()
        $class_dec=DecodeClass()
        $ttl_dec=""
        $data_dec=""
        if $section_id<>"q" Then
            $ttl_dec=ReadHex2Int(4)
            $rd_len=ReadHex2Int(2)
            $iDig_ptr_end=$iDig_ptr+$rd_len
            $data_dec=DecodeRData($rd_len,$type_dec)
        EndIf
        $iMax=UBound($aRet,1)
        ReDim $aRet[$iMax+1][5]
        $aRet[$iMax][0]=$name_dec
        $aRet[$iMax][1]=$ttl_dec
        $aRet[$iMax][2]=$class_dec
        $aRet[$iMax][3]=$type_dec
        $aRet[$iMax][4]=$data_dec
    Next
    return $aRet
EndFunc


;~ _ArrayDisplay(_nResolve('LT202165'),@Error&','&@extended)
;~ MsgBox(64,@Error,@extended)
