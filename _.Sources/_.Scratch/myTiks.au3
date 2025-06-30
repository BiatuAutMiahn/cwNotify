#include <Debug.au3>
#include <Array.au3>
#include <Date.au3>
#include "..\Includes\JSON_Dictionary.au3"
#include "..\Includes\Base64.au3"
#include "..\Includes\ArrayMultiColSort.au3"

Global $g_cwm_sEpoch="1970/01/01 00:00:00"
Global $aTiks[][4]=[[0,'','','']]
Global $gsStateFile=@LocalAppDataDir&"\InfinitySys\cwNotifier\state.rc.ini"
;Global $sModFields="_info.dateEntered,_info.lastUpdated,id,status.name,owner.name,summary,company.name,contact.name,subType.name,item.name,priority.name,severity.name,type.name,_info.enteredBy,_info.updatedBy"
Global $sModFields="id,type.name,priority.name,status.name,item.name,subType.name,company.name,summary,_info.lastUpdated,_info.dateEntered,owner.name,slaStatus"

Global $aFieldsDesc[][2]=[ _
    [0,0], _
    ["id","id"], _
    ["status.name","Status"], _
    ["owner.name","Owner"], _
    ["summary","Summary"], _
    ["company.name","Company"], _
    ["contact.name","Contact"], _
    ["subType.name","subType"], _
    ["item.name","Item"], _
    ["priority.name","Priority"], _
    ["severity.name","Severity"], _
    ["impact.name","Impact"], _
    ["type.name","Type"], _
    ["_info.dateEntered","Created"], _
    ["_info.enteredBy","Creator"], _
    ["_info.updatedBy","Modified By"], _
    ["_info.lastUpdated","Updated"], _
    ["slaStatus","SLA Status"] _
]

$aFieldsDesc[0][0]=UBound($aFieldsDesc,1)-1
For $i=1 To $aFieldsDesc[0][0]
  $iLen=StringLen($aFieldsDesc[$i][1])
  If $iLen>$aFieldsDesc[0][1] Then $aFieldsDesc[0][1]=$iLen
Next

StringReplace($sModFields,",","")
Local $iFields=@Extended+2
Global $aUserTik[1][$iFields]
_Log(_loadState()&','&@Error&','&@Extended)

;~ _DebugArrayDisplay($aTiks)
For $i=1 To $aTiks[0][0]
  If $aTiks[$i][0]="2049912" Then
    ClipPut(_JSON_Generate($aTiks[$i][3]))
    exit
  EndIf
Next
Exit
For $i=1 To $aTiks[0][0]
  Dim $aFields
  If StringInStr(_JSON_Get($aTiks[$i][2],"closedFlag"),"true") Then ContinueLoop
  _tikGetFields($aTiks[$i][2],$aFields,$sModFields)
  $iMax=UBound($aUserTik,1)
  ReDim $aUserTik[$iMax+1][$iFields]
  For $y=1 To $aFields[0][0]

    If $aUserTik[0][$y-1]="" Then $aUserTik[0][$y-1]=_getFieldDesc($aFields[$y][0])
    $aUserTik[$iMax][$y-1]=$aFields[$y][1]
  Next
Next
_ArrayDisplay($aTiks)
$aUserTik[0][12]="Age"
For $i=1 To UBound($aUserTik,1)-1
  $aUserTik[$i][8]=_cwmConvDate2Read($aUserTik[$i][8])
  ;$iCreate=_cwmConvDate2Sec($aUserTik[$i][9])
  Local $dtCreate=$aUserTik[$i][9]
  $aUserTik[$i][9]=_cwmConvDate2Read($dtCreate)
  $iNow=_cwmConvDate2Sec(_NowCalc())
  $iCreate=_cwmConvDate2Sec($dtCreate)
  $iDiff=$iNow-$iCreate
  $iAge=Round($iDiff/(24*60*60),1)
  $aUserTik[$i][12]=$iAge
  ;ConsoleWrite($iCreate&','&$iNow&','&$iDiff&','&$iAge&@CRLF)
Next
Global $aSortData[][] = [[2, 0],[3, 0],[12,1],[1, 0],[10, 1]]
_ArrayMultiColSort($aUserTik, $aSortData,1)
_ArrayDisplay($aUserTik)
; SortBy: SLA,Priority,Status,Type
;LastMod,Age,SLA,Board,Tik,Type,Priority,Status,Owner,Company,Desc,Contact

; Retireve and Format Fields.
Func _tikGetFields(ByRef $tData, ByRef $aFields, ByRef $sFields)
  Local $vData
  $aFields=StringSplit($sFields,',')
  _ArrayColInsert($aFields,1)
  For $j=1 To $aFields[0][0]
    $vData=_JSON_Get($tData,$aFields[$j][0])
    Switch $aFields[$j][0]
      Case "slaStatus"
        $vData=StringRegExpReplace($vData,"Resolve by [a-zA-z]{3} (\d{2})/(\d{2}) (\d{1,2}):(\d{2}) ([a-zA-Z]{2}) (UTC[-+]\d{1,2})","9 - $6 "&@YEAR&".$1.$2 $3:$4 $5")
        If $vData="Waiting" Then $vData="1 - Waiting"
        If $vData="" Then $vData="0 - No SLA"
      Case "priority.name"
        Switch $vData
          Case "Critical"
            $vData="0, "&$vData
          Case "Urgent"
            $vData="1, "&$vData
          Case "Standard"
            $vData="2, "&$vData
          Case "Planned"
            $vData="3, "&$vData
          Case Else
            $vData="9, "&$vData
        EndSwitch
      Case "status.name"
        Switch $vData
          Case "In Progress"
            $vData="0, "&$vData
          Case "Needs Escalation"
            $vData="1, "&$vData
          Case "Needs Followup"
            $vData="2, "&$vData
          Case "Waiting On Client"
            $vData="3, "&$vData
          Case "Scheduled"
            $vData="4, "&$vData
          Case Else
            $vData="9, "&$vData
        EndSwitch
      Case "type.name"
        Switch $vData
          Case "Alert"
            $vData="0, "&$vData
          Case "Incident"
            $vData="1, "&$vData
          Case "Problem"
            $vData="2, "&$vData
          Case "No Type"
            $vData="3, "&$vData
          Case "Request"
            $vData="4, "&$vData
          Case Else
            $vData="9, "&$vData
        EndSwitch
      Case "owner.name"
        If $vData='' Then $vData="(Unassigned)"
      ;Case "_info.dateEntered"
      ;  $vData=_cwmConvDate2Read($vData)
      ;Case "_info.lastUpdated"
      ;  $vData=_cwmConvDate2Read($vData)
    EndSwitch
    If $vData='' Then $vData='-'
    $aFields[$j][1]=$vData
  Next
EndFunc   ;==>_tikGetFields

Func _getFieldDesc($sDesc)
  For $i=1 To $aFieldsDesc[0][0]
    If $aFieldsDesc[$i][0]<>$sDesc Then ContinueLoop
    Return $aFieldsDesc[$i][1]
  Next
  Return $sDesc
EndFunc   ;==>_getFieldDesc


Func _getFieldCol(ByRef $aArr,$sDesc)
  Local $iMaxY=UBound($aArr,2)
  For $i=0 To $iMaxY-1
    If $aArr[0][$i]=$sDesc Or $aArr[0][$i]=_getFieldDesc($sDesc) Then Return $i
  Next
  Return SetError(1,1,0)
EndFunc   ;==>_getFieldDesc


Func _cwmConvDate2Sec($sDate)
  If Not _DateIsValid($sDate) Then Return SetError(1,0,0)
  Local $sRet=_DateDiff('s',$g_cwm_sEpoch,$sDate)
  If @error Then Return SetError(1,@error,0)
  Return SetError(0,0,$sRet)
EndFunc   ;==>_cwmConvDate2Sec

Func _cwmConvSec2Date($iSec)
  If Not IsInt($iSec) Then Return SetError(1,0,0)
  Local $sDate=_DateAdd('s',$iSec,$g_cwm_sEpoch)
  If @error Then Return SetError(1,@error,0)
  Return SetError(0,0,$sDate)
EndFunc   ;==>_cwmConvSec2Date

Func _cwmConvDate2Read($sDate)
  If Not _DateIsValid($sDate) Then Return SetError(1,(@error*10)+1,0)
  ; Convert TimeZone
  Local $aDate,$aTime,$iHour,$sMeridiem
  _DateTimeSplit($sDate,$aDate,$aTime)
  If @error Then Return SetError(1,(@error*10)+2,0)
  $tSysTime=_Date_Time_EncodeSystemTime($aDate[2],$aDate[3],$aDate[1],$aTime[1],$aTime[2],$aTime[3])
  If @error Then Return SetError(1,(@error*10)+3,0)
  $tLocal=_Date_Time_SystemTimeToTzSpecificLocalTime($tSysTime)
  If @error Then Return SetError(1,(@error*10)+4,0)
  $sDate=_Date_Time_SystemTimeToDateTimeStr($tLocal)
  If @error Then Return SetError(1,(@error*10)+5,0)
  _DateTimeSplit($sDate,$aDate,$aTime)
  If @error Then Return SetError(1,(@error*10)+6,0)
  _ArrayDelete($aDate,0)
  If @error Then Return SetError(1,(@error*10)+7,0)
  _ArrayConcatenate($aDate,$aTime,1)
  If @error Then Return SetError(1,(@error*10)+8,0)
  $iHour=$aDate[3]
  If $aDate[3]>12 Then
    $iHour=$aDate[3]-12
    $sMeridiem="p"
  Else
    $sMeridiem="a"
  EndIf
  $sDate=StringFormat("%04d.%02d.%02d@%d%02d%s",$aDate[2],$aDate[0],$aDate[1],$iHour,$aDate[4],$sMeridiem)
  Return SetError(0,0,$sDate)
EndFunc   ;==>_cwmConvDate2Read


Func _loadState()
  If Not FileExists($gsStateFile) Then Return SetError(1,1,0)
  $aConfig=IniReadSectionNames($gsStateFile)
  Dim $aConfigNew[$aConfig[0]+1][6]
  $aConfigNew[0][0]=$aConfig[0]
  For $i=1 To $aConfig[0]
    $aConfigNew[$i][0]=$aConfig[$i]
    $aConfigNew[$i][1]=IniRead($gsStateFile,$aConfig[$i],"LastMod","")
    If $aConfigNew[$i][1]=="" Then
      _Log("LastState Corrupt ("&$aConfig[$i]&')'&@CRLF)
      ContinueLoop
    EndIf
    $vData=IniRead($gsStateFile,$aConfig[$i],"jTik","")
    If $vData=="" Then
      _Log("LastState Corrupt ("&$aConfig[$i]&')'&@CRLF)
      ContinueLoop
    EndIf
    $vData=BinaryToString(_Base64Decode($vData))
    $aConfigNew[$i][2]=_JSON_Parse($vData)
    $vData=IniRead($gsStateFile,$aConfig[$i],"jLastNote","")
    If $vData=="" Then
      _Log("LastState Corrupt ("&$aConfig[$i]&')'&@CRLF)
      ContinueLoop
    EndIf
    $vData=BinaryToString(_Base64Decode($vData))
    $aConfigNew[$i][3]=_JSON_Parse($vData)
    $aConfigNew[$i][4]=IniRead($gsStateFile,$aConfig[$i],"Type","")
  Next
  ;_ArrayDisplay($aConfigNew)
  $aTiks=$aConfigNew
EndFunc   ;==>_loadState

Func _Log($sLine)
  ConsoleWrite($sLIne)
EndFunc
