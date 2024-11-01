
#pragma compile(AutoItExecuteAllowed,True)
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Res\cwdgs.ico
#AutoIt3Wrapper_Outfile_x64=..\_.rc\cwNotify.exe
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Res_Description=ConnectWise Notifier
#AutoIt3Wrapper_Res_ProductName=
#AutoIt3Wrapper_Res_Fileversion=1.1.0.1017
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_Fileversion_First_Increment=y
#AutoIt3Wrapper_Run_After=echo %fileversion%>..\VERSION.rc
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/tl /debug /mo /rsln
#AutoIt3Wrapper_Change2CUI=n
#AutoIt3Wrapper_Run_Tidy=n
#Tidy_Parameters=/kv 0 /reel /tc 2 /tcb 1
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#pragma compile(AutoItExecuteAllowed,False)
Opt("TrayAutoPause",0)
Opt("TrayIconHide",1)
Opt("TrayMenuMode",3)

#cs ----------------------------------------------------------------------------

  AutoIt Version: 3.3.16.1
  Source: https://github.com/BiatuAutMiahn/cwNotifyAu3

  TODO :
  -Implement Ticket Notify Queue.
  -Make Notify UI Non-Blocking
  -Implement UI Scaling
  -Snap notify To cursor's active monitor. (Pull from ctOverlay)
  -Implement update functionality.
  -Implement Ticket Detail/History UI.(ListView/w DateTime,Summary,Company/Contact,Show audit trail Or history of notes.)
  -Tray Context Menu :
  -UI Scaling Option
  -Implemented best server detection(DNS resolution/ping A records,Not used)
  -Implement change highlights.
  -Implement Notify With each line In a multidim array[n]. include color info In element[n+1].
  -Implement History UI,a list view With columns ordered by last ticket recieved. 2 Panes,1 pane For list[lastUpdated|id|summary] other pane containing the notify content.
  -Summary acts like the dismiss all button In that it will Not alert you For the rest of that batch.

#ce ----------------------------------------------------------------------------

#include <Misc.au3>
#include <Array.au3>
#include <Debug.au3>
#include <Date.au3>
#include <Timers.au3>
#include <WinAPISys.au3>
#include <StaticConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiImageList.au3>
#include <GuiTreeView.au3>
#include <MsgBoxConstants.au3>
#include <Math.au3>
#include <WinAPISys.au3>
#include <GuiEdit.au3>
#include <WinAPIProc.au3>

#include "Includes\JSON_Dictionary.au3"
#include "Includes\Base64.au3"
#include "Includes\WinHTTP.au3"
#include "Includes\Toast.au3"
#include "Includes\GuiFlatButton.au3"
#include "Includes\CryptProtect.au3"
#include "Includes\_WinAPI_DPI.au3"
#include "Includes\_newResolve.au3"
#include "Includes\_nPing.au3"
Global Const $giLineMain=@ScriptLineNumber
;#include "..\Includes\_StringInPixels.au3"
Global Const $sAlias="cwNotify"
Global Const $VERSION = "1.1.0.1017"
Global $sTitle=$sAlias&" v"&$VERSION

; Logging,Purge log >=1MB
Global $gsDataDir=@LocalAppDataDir&"\InfinitySys\cwNotifier"
Global $gsLogFile=$gsDataDir&"\cwNotifier.log"

Func _Log($sLog)
  DirCreate($gsDataDir)
  If FileGetSize($gsLogFile)>=1024*1024 Then FileDelete($gsLogFile)
  $sStamp=StringFormat("%s.%02d.%02d,%02d:%02d:%02d.%03d",@YEAR,@MON,@MDAY,@HOUR,@MIN,@SEC,@MSEC)
  FileWriteLine($gsLogFile,'['&$sStamp&'] '&$sLog)
  ConsoleWrite('+>['&$sStamp&'] '&$sLog&@CRLF)
EndFunc   ;==>_Log

$g_oErrorCbDef=ObjEvent("AutoIt.Error")
$g_oErrorCb=ObjEvent("AutoIt.Error","_AutErrorFunc")
Func _AutErrorFunc()
  If Not IsObj($g_oErrorCb) Then Return
  $g_iAutError=1
  $g_iAutErrorExt=$g_oErrorCb.number
  $g_sAutError=$g_oErrorCb.windescription&" (0x"&Hex($g_iAutErrorExt)&")"
  _Log($g_sAutError)
EndFunc   ;==>_AutErrorFunc

; Error Handling (Execute self with /ErrorStdOut flag and write output to _Log)
;~ If Not StringInStr($CmdLineRaw,"/ErrorStdOut",0) Then
;~     ; Prevent multiple Instances of the wrapper.
;~     If _Singleton("Infinity."&$sAlias&"_Wrap",1)=0 Then
;~         MsgBox(32,$sTitle,"Another instance is already running.")
;~         Exit
;~     EndIf

;~     ; Purge last log
;~     FileDelete($gsLogFile)

;~     ; Run and Wait for errors.
;~     Local $iExit=0,$sErr=''
;~     Local $tProcess=DllStructCreate($tagPROCESS_INFORMATION)
;~     Local $tStartup=DllStructCreate($tagSTARTUPINFO)
;~     ;If _WinAPI_CreateProcess($sTitle,@AutoItExe
;~     $iPid=Run(@AutoItExe,"/ErrorStdOut "&$CmdLineRaw,@SW_SHOW,$STDERR_CHILD)
;~     Local $hProcess=_WinAPI_OpenProcess($PROCESS_QUERY_LIMITED_INFORMATION,0,$iPid,1)
;~     If Not $hProcess Then
;~         _Log("Failed to create/open child process. Error: "&@Error)
;~         Exit 1
;~     EndIf
;~     Do
;~         $vStdErr=StderrRead($iPid,1,0)
;~         If @error Then
;~             ExitLoop
;~         EndIf
;~         If $vStdErr<>'' Then
;~             $sErr&=$vStdErr
;~             ConsoleWriteError($vStdErr)
;~             If StringInStr($vStdErr,@CR) Then
;~                 _Log($sErr)
;~                 $sErr=''
;~             EndIf
;~         EndIf
;~         Sleep(1)
;~     Until Not ProcessExists($iPid)
;~     Exit 1 ;_WinAPI_GetExitCodeProcess($hProcess)
;~ EndIf

;~ ; Owner,Note
;~ Local $aLastNote[]=[ _
;~     "Tara Butts", _
;~     "Hi,I tried my best to trouble shoot the computers at YMCA Brc and have no luck."&@CRLF& _
;~     ""&@CRLF& _
;~     "we have like over 10 computers down."&@CRLF& _
;~     ""&@CRLF& _
;~     "Best,"&@CRLF& _
;~     "Tara"&@CRLF& _
;~     ""&@CRLF& _
;~     "Tara N Butts,"&@CRLF& _
;~     "Administrator Assistant "&@CRLF& _
;~     "BRC Vanderbilt Stabilization Program-YMCA "&@CRLF& _
;~     "224 47th St "&@CRLF& _
;~     "New York,NY 10017  "&@CRLF& _
;~     "WP 646-841-4323 "&@CRLF& _
;~     "E: Tbutts@brc.org   "&@CRLF& _
;~     "Please note I am in the office Monday Through Friday,"&@CRLF& _
;~     "9am-5:30pm "&@CRLF& _
;~     "You can have anything you want in life if you dress for it... "&@CRLF& _
;~     " "&@CRLF& _
;~     "This e-mail message and any attachments may contain confidential information meant solely for the intended recipient. If you have received this message in error,please notify the sender immediately by replying to this message,then delete the e-mail and any attachments from your system. Any use of this message or its attachments that is not in keeping with the confidential nature of the information,including but not limited to disclosing information to others,dissemination,distribution,or copying is strictly prohibited. "&@CRLF _
;~ ]
;~ ;",,,,,,,,,,,,"
;~ Local $aFields[][2]=[[13,''], _
;~     ["_info.lastUpdated","2024.07.23@1151a"], _
;~     ["id","1937290"], _
;~     ["status.name","Needs Escalation"], _
;~     ["owner.name","Luis Perez"], _
;~     ["summary","CrowdStrike issue remediation"], _
;~     ["company.name","BRC-ADMINISTRATION"], _
;~     ["contact.name","Thomas Wyse"], _
;~     ["subType.name","Workstation"], _
;~     ["item.name","BSOD"], _
;~     ["priority.name","Urgent"], _
;~     ["severity.name","Medium"], _
;~     ["type.name","Problem"], _
;~     ["_info.updatedBy","Corsica"] _
;~ ]
;~ ; $iFlag is for Ignoring or other modifiers.
;~ ; $hwndNotify,Ticket#,$iFlag,tSnoozeTimer,Title,$aFields,$aLastNote
;~ Local $aN[]=[ _
;~     -1, _
;~     1937290, _
;~     0, _
;~     0, _
;~     "Ticket Updated", _
;~     $aFields, _
;~     $aLastNote _
;~ ]
;~ Local $aNotify[]=[1,$aN]
;~ ;ConsoleWrite(($aNotify[1])[0]&@CRLF)
;~ _ShowNotify($aNotify[1])
;~ ;ConsoleWrite(($aNotify[1])[0]&@CRLF)
;~ ;For $i=1 To $aNotify[0]
;~     ;ConsoleWrite(($aNotify[$i])[0]&@CRLF)
;~ ;Next

;~ ;Next
;~ Exit
If StringInStr($CmdLineRaw,"~!Install") Then
  cwInstall()
  Exit 0
EndIf
If Not StringInStr($CmdLineRaw,"~!PostInstall") And _Singleton("Infinity."&$sAlias,1)=0 Then
  ; Prevent multiple instances of the App.
  MsgBox(32,$sTitle,"Another instance is already running.")
  _Log("_Singleton,Exit")
  Exit
EndIf

Global $aNotify[]=[0]
Global $aTiksLast[][4]=[[0,'','','']]
Global $aTiks[][4]=[[0,'','','']]
Global $bExit
Global $tDelay
Global $g_cwm_oHttp
Global $g_cwm_sEpoch="1970/01/01 00:00:00"
Global $g_cwm_sCI,$g_cwm_sCompany,$g_cwm_sCodeBase,$g_cwm_sSiteUrl,$g_cwm_sApiUri,$g_cwm_sClientId,$g_cwm_sAuthToken
Global $g_cwm_sClientId,$g_cwm_sPrivKey,$g_cwm_sPubKey,$g_cwm_sUser,$g_cwm_jLastRet
Global $g_cwm_hHttp,$g_cwm_hConnect
Global $bFieldMod

Global $sNewFields="_info.dateEntered,_info.lastUpdated,id,status.name,owner.name,summary,company.name,contact.name,subType.name,item.name,priority.name,severity.name,type.name,_info.enteredBy"
Global $sModFields="_info.dateEntered,_info.lastUpdated,id,status.name,owner.name,summary,company.name,contact.name,subType.name,item.name,priority.name,severity.name,type.name,_info.updatedBy"
Global $sNoModFields="_info.lastUpdated"

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
    ["_info.lastUpdated","Updated"] _
]

$aFieldsDesc[0][0]=UBound($aFieldsDesc,1)-1
For $i=1 To $aFieldsDesc[0][0]
  $iLen=StringLen($aFieldsDesc[$i][1])
  If $iLen>$aFieldsDesc[0][1] Then $aFieldsDesc[0][1]=$iLen
Next
Global $aNoModFields
If StringInStr($aNoModFields,',') Then
  $aNoModFields=StringSplit($sNoModFields,',')
Else
  Dim $aNoModFields[]=[1,$sNoModFields]
EndIf

;_DebugArrayDisplay($aFieldsDesc)
Global $idTrayExit

; cwmAuth Vars
Global $g_cwmAuth_idDoneBtn
Global $g_cwmAuth_aidInput
Global $g_cwmAuth_abValidate
Global $g_cwm_sClientId_Crypt

; Config
Global $gsConfigFile=$gsDataDir&"\config.ini"

; DPI
_WinAPI_SetProcessDpiAwarenessContext($DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)
$iDpi=_WinAPI_GetDpiForPrimaryMonitor()/96

_Log("======================================================"&@CRLF)

Global $gsStateFile=$gsDataDir&"\state.ini"
Global $g_cwm_aEndPoints,$g_cwm_sSite="na.myconnectwise.net"
Global $g_cwm_aApiEndPoints
Global $bOffline=False,$bOfflineLast=False,$iOfflineThresh=3

;
; Main
;
_Toast_Set(0,-1,-1,-1,-1,-1,"Consolas",125,125)

;~ $sTikTitle="[New Ticket]"
;~ $sNotify="    Created: 2024.07.30@459p"&@CRLF&_
;~     "    Updated: 2024.08.02@1213a"&@CRLF&_
;~     "         id: 1948137"&@CRLF&_
;~     "     Status: Needs Followup"&@CRLF&_
;~     "      Owner: <TechnicianA>"&@CRLF&_
;~     "    Summary: Needs to download all emails from user@domain.tld"&@CRLF&_
;~     "    Company: <Client Name>"&@CRLF&_
;~     "    Contact: <Contact Name>"&@CRLF&_
;~     "    subType: Email"&@CRLF&_
;~     "       Item: Information"&@CRLF&_
;~     "   Priority: Standard"&@CRLF&_
;~     "   Severity: Medium"&@CRLF&_
;~     "       Type: Request"&@CRLF&_
;~     "    Creator: <TikCreator>"&@CRLF&_
;~     "  Last Note: [<TechnicianA>] Assigned/<TechnicianB>/"

;~ $sNotify=BinaryToString(_Base64Decode("ICAgIENyZWF0ZWQ6IDIwMjQuM;kuMzBAMTIyNGENCiAgICBVcGRhdGVkOiAyMDI0LjA5LjMwQDMwMnANCiAgICAgICAgIGlkOiAyMDIyOTM1DQogICAgIFN0YXR1czogTmVlZHMgRXNjYWxhdGlvbg0KICAgICAgT3duZXI6IENocmlzdG9waGVyIEdvcmRvbg0KICAgIFN1bW1hcnk6IE5ldyBWb2ljZSBNZXNzYWdlIGZyb20gU2VydmljZSBEZXNrIC0gRUlLTyBHTE9CQUwgKDkxMykgNjY3LTg1MzUgb24gMDkvMzAvMjAyNCAxMDoyMyBBTQ0KICAgIENvbXBhbnk6IEVpS08gR2xvYmFsIExMQw0KICAgIENvbnRhY3Q6IEVpS08gU3VwcG9ydA0KICAgIHN1YlR5cGU6ICpNVVNUIENIQU5HRSoNCiAgICAgICBJdGVtOiAqTVVTVCBDSEFOR0UqDQogICBQcmlvcml0eTogVXJnZW50DQogICBTZXZlcml0eTogTWVkaXVtDQogICAgICAgVHlwZTogSW5jaWRlbnQNCk1vZGlmaWVkIEJ5OiBqaHV0dG8gLT4gY2dvcmRvbg0KICBMYXN0IE5vdGU6IFtSaW5nQ2VudHJhbF0gIVtcW0xvZ29cXV0oaHR0cHM6Ly9uZXRzdG9yYWdlLnJpbmdjZW50cmFsLmNvbS9lbWFpbC8xeDEuZ2lmKSFbXFtMb2dvXF1dKGh0dHBzOi8vbmV0c3RvcmFnZS5yaW5nY2VudHJhbC5jb20vZW1haWwvMXgxLmdpZikKIVtcW0xvZ29cXV0oaHR0cHM6Ly9uZXRzdG9yYWdlLnJpbmdjZW50cmFsLmNvbS9pbWFnZXMvdW5zL3JpbmdjZW50cmFsL2xvZ28vZGVmYXVsdC8yMDIzL2xvZ28tZW5fVVMucG5nKVZvaWNlIE1lc3NhZ2UKCkRlYXIgU2VydmljZSBWb2ljZW1haWwsICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKWW91IGhhdmUgYSBuZXcgdm9pY2UgbWVzc2FnZToKKipGcm9tOioqU2VydmljZSBEZXNrIC0gRUlLTyBHTE9CQUwgXCg5MTNcKSA2NjctODUzNQoqKlJlY2VpdmVkOioqTW9uZGF5LCBTZXB0ZW1iZXIgMzAsIDIwMjQgYXQgMTA6MjMgQU0KKipMZW5ndGg6KiowMDozNAoqKlRvOioqXCg4NTVcKSA0MTEtMzM4NyBcKiA5OTk4IFNlcnZpY2UgVm9pY2VtYWlsCgoqKlZvaWNlbWFpbCBQcmV2aWV3OioqCgoiWWVhaCwgbXkgbmFtZSBpcyBCb2JieSBSb3NzLiBUZWxlcGhvbmUgbnVtYmVycyA5MTM5MTUwMzQ3SSB3YXMganVzdCBnaXZpbmcgeW91IGEgY2FsbC4gSSBoYWQgYSBraW5kIG9mIGFuIG9kZCBlbWFpbCBzaXR1YXRpb24sIG5vdCBzdXJlIGhvdyBpdCBoYXBwZW5lZCwgYnV0IEkganVzdCB3YW50ZWQgdG8gdGFsayB0byBzb21lYm9keSBpbiB0aGUgU09DIHRvIHNlZSBpZiB0aGV5IGNhbiBsZW5kIHNvbWUgYXNzaXN0YW5jZSBhbmQgZmlndXJpbmcgb3V0IGhvdyBpdCBob3cgaXQgaGFwcGVuZWQuIE15IHRlbGVwaG9uZSBudW1iZXIgYWdhaW4gaXMgOTEzLTkxNS0wMzQ3LCB0aGFuayB5b3UuIiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCgpMaXN0ZW4gdG8gdGhpcyB2b2ljZW1haWwgb3ZlciB5b3VyIHBob25lIG9yIGJ5IG9wZW5pbmcgdGhlIGF0dGFjaGVkIHNvdW5kIGZpbGUuIFlvdSBjYW4gYWxzbyBzaWduIGluIHRvIHlvdXIgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBbUmluZ0NlbnRyYWwgYWNjb3VudF0oaHR0cHM6Ly9zZXJ2aWNlLnJpbmdjZW50cmFsLmNvbSkgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIHdpdGggeW91ciBtYWluIG51bWJlciwgZXh0ZW5zaW9uIG51bWJlciwgYW5kIHBhc3N3b3JkIHRvIG1hbmFnZSBhbmQgbGlzdGVuIHRvIHZvaWNlbWFpbHMuICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAoKVGhhbmsgeW91IGZvciB1c2luZyBSaW5nQ2VudHJhbCEKCldvcmsgZnJvbSBhbnl3aGVyZSB3aXRoIHRoZSBSaW5nQ2VudHJhbCBhcHAuIEl0J3MgZ290IGV2ZXJ5dGhpbmcgeW91IG5lZWQgdG8gc3RheSBjb25uZWN0ZWQ6IHRlYW0gbWVzc2FnaW5nLCB2aWRlbyBtZWV0aW5ncyBhbmQgcGhvbmUgLSBhbGwgaW4gb25lIGFwcC4gW0dldCBzdGFydGVkXShodHRwczovL2FwcC5yaW5nY2VudHJhbC5jb20pCgpCeSBzdWJzY3JpYmluZyB0byBhbmQvb3IgICAgICAgICAgICAgICAgICAgICAgICB1c2luZyBSaW5nQ2VudHJhbCwgeW91IGFja25vd2xlZGdlIGFncmVlbWVudCB0byAgICAgICAgICAgICAgICAgICAgICAgIG91ciBbVGVybXMgb2YgVXNlXShodHRwczovL3d3dy5yaW5nY2VudHJhbC5jb20vbGVnYWwvZXVsYXRvcy5odG1sKS4gCgpDb3B5cmlnaHQgMjAyNCBSaW5nQ2VudHJhbCwgSW5jLiBBbGwgcmlnaHRzIHJlc2VydmVkLiBSaW5nQ2VudHJhbCBhbmQgdGhlIFJpbmdDZW50cmFsIGxvZ28gYXJlIHRyYWRlbWFya3Mgb2YgUmluZ0NlbnRyYWwsIEluYy4sICAgICAgICAgICAgICAgICAgICAgICAgMjDCoERhdmlzwqBEcml2ZSwgQmVsbW9udCwgQ0HCoDk0MDAyLCBVU0Eu"))
;~ $aRet=_Toast_ShowMod(0,$sTikTitle,$sNotify,Null,True,True)
;~ If @error Then _Log(StringFormat("~!Error@%s,_Toast_ShowMod:%s:%s",@ScriptLineNumber,@Error,@extended))
;~ MsgBox(64,"",$fToast_bDismissAll)
;~ Exit

; Tray Config
Opt("TrayIconHide",0)
TraySetToolTip($sAlias)
_nResolveUpdateInfoDNS()

;_getIdealSrv($g_cwm_sSite)
Local $bExit=False

$idTrayExit=TrayCreateItem("Exit")
AdlibRegister("_TrayEvent",20)

_Toast_Hide()

Global $iWatchTimer=TimerInit()
Global $iWatchInterval=10
Global $bFirstRun=True
Global $iMaxUrlLen=1745

Global $iIdx
Global $bTikNew=False
Global $bBatch
Global $aWatchFields=StringSplit("summary,type,subType,item,status,company,priority,severity,impact,owner",',')
Local $vTikId,$vTikLastUpdate
Global $iUserIdle,$bUserIdle,$bUserIdleLast,$bWsLock,$bWsLockLast
Global $fToast_OpenTik=False,$hToast_OpenTik
Global $fToast_bDismissAll=False,$hToast_DismissAll

;Global $aQueue[][3]=[[0,0,0]

_Log($sTitle)
_Log("Debug.StartLine:"&$giLineMain)
_Log("iDPI: "&$iDpi&@CRLF)
If FileExists($gsConfigFile) Then $bFirstRun=False
If Not $bFirstRun Then
  $tDelay=TimerInit()
  _Log("Welcome Back...")
  $aRet=_Toast_ShowMod(0,$sTitle,"Welcome Back!                    ",Null,0)  ; Don't Wait
  _Log("Loading State...")
  _loadState()
  If Not @error Then $bFirstRun=False
  _Log("Loading State...Done")
  While TimerDiff($tDelay)<=5000
    Sleep(125)
  WEnd
  _Toast_Hide()
Else
  $aRet=_Toast_ShowMod(0,$sTitle,"Welcome to "&$sTitle,-5)
  _Toast_Hide()
EndIf
While doOfflineCheck()
  Sleep(125)
WEnd

If StringInStr($CmdLineRaw,"~!PrePurgeOldTiks") And Not $bFirstRun Then
  _Log("Purging Old Tiks...")
  If _PurgeOldTiks() Then
    FileDelete($gsStateFile)
    _saveState()
  EndIf
  _Log("Purging Old Tiks...Done")
EndIf

_Log("Initializing WinHttp...")
$iTimer=TimerInit()
_InitHttp()
If @error Then
  _Log((@error*1000)+@extended)
  _Exit()
EndIf
_Log("Initializing WinHttp...Done")
_Log("Loading Config...")
_loadConfig()

If $g_cwm_sCompany=='' Or $g_cwm_sClientId=='' Or $g_cwm_sPrivKey=='' Or $g_cwm_sPubKey=='' Or StringInStr($CmdLineRaw,"~!cwmAuth") Then
  _cwmAuth()
  _loadConfig()
EndIf
_Log("Loading Config...Done")
_Log("Initializing cwm...")
_cwmInit()
If @error Then
  _Log((@error*2000)+@extended)
  _Exit()
EndIf
_Log("Initializing cwm...Done")

$iTimer=TimerInit()
Local $bDev=@Compiled ? False : True


;1961209
;$jFetch=_cwmCall("/service/tickets?conditions=id=1961209&pageSize=1000")
;ClipPut(_JSON_Generate($jFetch))
;Exit
ConsoleWrite("Watching Tickets..."&@CRLF)
While Sleep(125)
  ; Test Connectivity
  While doOfflineCheck()
    Sleep(1000)
  WEnd
  If (Not $bFirstRun And Not $bUserIdle) And TimerDiff($iWatchTimer)<=($iWatchInterval*1000) Then ContinueLoop
  $iWatchTimer=TimerInit()
  $tMainLoop=TimerInit()
  ;
  ; Idle detection
  ;
  $iUserIdle=_Timer_GetIdleTime()
  $bUserIdle=$iUserIdle>=300000 ? True : False
  $bWsLock=_isWindowsLocked()
  If $bWsLockLast<>$bWsLock Then
    If $bWsLock Then
      $bUserIdle=True
    EndIf
    $bWsLockLast=$bWsLock
  EndIf
  If $bUserIdleLast<>$bUserIdle Then
    $bUserIdleLast=$bUserIdle
    If Not $bUserIdle Then
      _Log("Welcome Back"&@CRLF)
      _Toast_ShowMod(0,$sTitle,"Welcome Back!                    ",-30)
      _Toast_Hide()
      Sleep(5000)
    Else
      _Log("Workstation Idle."&@CRLF)
    EndIf
  EndIf
  If $bUserIdle Then ContinueLoop
  ;
  ; Get Ticket Info
  ;
  $bBatch=True
  $iTimer=TimerInit()
  ;_Log(TimerDiff($iTimer)&@CRLF)
  ;_Log("Checking Service Tickets...")
  _cwmGetTickets($aTiks,0,$g_cwm_sUser)
  If @error Then
    _Log("Error "&@extended&",cannot check for Service tickets.")
    $iWatchTimer=TimerInit()
    ContinueLoop
  EndIf
  ;_Log("Checking Service Tickets...Done")
  ;_Log("Checking Project Tickets...")
  _cwmGetTickets($aTiks,1,$g_cwm_sUser)
  If @error Then
    _Log("Error "&@extended&",cannot check for Project tickets.")
    $iWatchTimer=TimerInit()
    ContinueLoop
  EndIf
  ;_Log("Checking Project Tickets...Done")
  ;_Log(TimerDiff($iTimer)&@CRLF)
  ;
  ; Compare fields
  ;
  Local $bNotify
  Local $aOldFields
  Local $aNewFields
  Local $aModFields
  Global $bCommit=False

  ;_DebugArrayDisplay($aTiks)
  $fToast_bDismissAll=False
  For $i=1 To $aTiks[0][0]
    $bNewTik=False
    $bNotify=False
    $bFieldMod=False
    $sNotify=''
    $tNew=$aTiks[$i][2]
    $iIdxLast=_cwmInArray($aTiksLast,$aTiks[$i][0])
    If $bExit Then
      _Exit()
    EndIf
    If @error Then
      $sTikTitle="[New Ticket]"
      _tikGetFields($tNew,$aNewFields,$sNewFields)
      For $j=1 To $aNewFields[0][0]
        $sNotify&=StringFormat("%"&$aFieldsDesc[0][1]&"s: %s",_getFieldDesc($aNewFields[$j][0]),$aNewFields[$j][1])&@CRLF
      Next
    Else
      ; If tik exists and date is not newer,then skip.
      If $aTiks[$i][1]<=$aTiksLast[$iIdxLast][1] Then ContinueLoop
      $sUpdater=_JSON_Get($tNew,'_info.updatedBy')
      If $sUpdater=$g_cwm_sUser Then ; Skip Updates by ourselves.
        _Log("Skipped Tik Mod By: "&$sUpdater)
        $bCommit=True
        ContinueLoop
      EndIf
      $sTikTitle="[Ticket Updated]"
      _tikGetFields($aTiksLast[$i][2],$aOldFields,$sModFields)
      _tikGetFields($tNew,$aModFields,$sModFields)
      For $j=1 To $aModFields[0][0]
        If _isFieldNoMod($aModFields[$j][0]) Or StringCompare($aOldFields[$j][1],$aModFields[$j][1])==0 Then
          $sNotify&=StringFormat("%"&$aFieldsDesc[0][1]&"s: %s",_getFieldDesc($aModFields[$j][0]),$aModFields[$j][1])&@CRLF
        Else
          $bFieldMod=True
          $sNotify&=StringFormat("%"&$aFieldsDesc[0][1]&"s: %s -> %s",_getFieldDesc($aModFields[$j][0]),$aOldFields[$j][1],$aModFields[$j][1])&@CRLF
        EndIf
      Next
    EndIf
    $bCommit=True
    $tLastNote=$aTiks[$i][3]
    $sName=''
    If $tLastNote.Exists("member") Then $sName=_JSON_Get($tLastNote,"member.name")
    If $tLastNote.Exists("contact") Then $sName=_JSON_Get($tLastNote,"contact.name")
    $sNotify&=StringFormat("%"&$aFieldsDesc[0][1]&"s: [%s] %s","Last Note",$sName,_JSON_Get($tLastNote,"text"))&@CRLF
    If $bBatch Then
      $bBatch=False
      _Log("================"&@CRLF)
    EndIf
    _Log($sTikTitle&@CRLF&$sNotify&@CRLF)
    If $fToast_bDismissAll=False Then
      $aRet=_Toast_ShowMod(0,$sTikTitle,$sNotify,Null,True,True)
      If @error Then _Log(StringFormat("~!Error@%s,_Toast_ShowMod:%s:%s",@ScriptLineNumber,@Error,@extended))
      If $fToast_OpenTik Then
        ShellExecute("https://na.myconnectwise.net/v4_6_release/services/system_io/Service/fv_sr100_request.rails?service_recid="&$aTiks[$i][0]&"&companyName="&$g_cwm_sCompany)
      EndIf
      $fToast_OpenTik=False
      _Toast_Hide()
    EndIf
  Next
  If Not $bBatch Then
    $bBatch=True
    _Log("================"&@CRLF)
  EndIf
  If $bFirstRun Then
    $bFirstRun=False
  EndIf
  If $bCommit Then
    ; Update aTikLast
    ;
    ;_DebugArrayDisplay($aTiksLast)
    For $i=1 To $aTiks[0][0]
      If $bExit Then
        _Exit()
      EndIf
      $iIdxLast=_cwmInArray($aTiksLast,$aTiks[$i][0])
      If @error Then
        $iMax=UBound($aTiksLast,1)
        ReDim $aTiksLast[$iMax+1][6]
        $iIdxLast=$iMax
        $aTiksLast[0][0]=$iMax
      EndIf
      For $j=0 To UBound($aTiks,2)-1
        $aTiksLast[$iIdxLast][$j]=$aTiks[$i][$j]
      Next
    Next
    ;
    ; Purge Resolved/Closed>7 days.
    ;
    If _PurgeOldTiks() Then
      FileDelete($gsStateFile)
    EndIf
    ;
    ; Finally,Serialize $aTiksLast,and save.
    ;
    _saveState()
  EndIf

  If $bExit Then
    _Exit()
  EndIf
  $iWatchTimer=TimerInit()
  _WinAPI_EmptyWorkingSet()
  _Log("MainLoop Took "&TimerDiff($tMainLoop))
WEnd
_Log("DropLoop"&@CRLF)

Func _CreateBorderLabel($sText,$iX,$iY,$iW,$iH,$iColor,$iPenSize=1,$iStyle=-1,$iStyleEx=0)
  $internalLabelID1=GUICtrlCreateLabel("",$iX,$iY,$iW,$iH,BitOR($SS_CENTER,$SS_CENTERIMAGE))
  GUICtrlSetBkColor(-1,$COLOR_GRAY)
  GUICtrlSetState(-1,$GUI_SHOW)
  $internalLabelID2=GUICtrlCreateLabel($sText,($iX-$iPenSize),($iY-$iPenSize),($iW+2*$iPenSize)-1,($iH+2*$iPenSize)-1,BitOR($SS_CENTER,$SS_CENTERIMAGE),$iStyleEx)
  GUICtrlSetBkColor($internalLabelID2,$COLOR_BLACK)   ; $GUI_BKCOLOR_TRANSPARENT
  GUICtrlSetColor(-1,$COLOR_WHITE)
  GUICtrlSetState(-1,$GUI_SHOW)

  Return $internalLabelID2
EndFunc   ;==>_CreateBorderLabel

;
; Generate Random 16 digit Alphanumeric String
; UEZ,modified by Biatu
;
Func _RandStr()
  Local $sRet="",$aTmp[3],$iLen=16
  For $i=1 To $iLen
    $aTmp[0]=Chr(Random(65,90,1))     ;A-Z
    $aTmp[1]=Chr(Random(97,122,1))     ;a-z
    $aTmp[2]=Chr(Random(48,57,1))     ;0-9
    $sRet&=$aTmp[Random(0,2,1)]
  Next
  Return $sRet
EndFunc   ;==>_RandStr

;
; cwmAuth
;

Global $g_cwmAuth_idDoneBtn
Global $g_cwmAuth_aidInput
Global $g_cwmAuth_abValidate

Func _cwmAuth()

  Local $bFatal
  Local $iuiWidth=320*$iDpi
  Local $iuiHeight=(128+64+28+4)*$iDpi
  Local $iuiMargin=4*$iDpi
  Local $iuiCtrlWidth=$iuiWidth-($iuiMargin*2)
  ConsoleWrite($iuiMargin&@CRLF)
  Local $iuiBtnWidth=75*$iDpi
  Local $iuiCtrlHeight=24*$iDpi
  Local $iuiBtnLeft=($iuiWidth/2)-$iuiBtnWidth-$iuiMargin
  $hAuthWnd=GUICreate($sTitle,$iuiWidth,$iuiHeight,-1,-1)
  GUISetFont(8,400,0,"Consolas")
  GUICtrlCreateLabel("Please provide API details to continue",$iuiMargin,$iuiMargin,$iuiCtrlWidth,13*$iDpi)
  $idCompany=GUICtrlCreateInput("",$iuiMargin,$iuiCtrlHeight,$iuiCtrlWidth,$iuiCtrlHeight)
  $idClientId=GUICtrlCreateInput("",$iuiMargin,($iuiMargin+$iuiCtrlHeight)*2,$iuiCtrlWidth,$iuiCtrlHeight)
  $idPubKey=GUICtrlCreateInput("",$iuiMargin,($iuiMargin+$iuiCtrlHeight)*3,$iuiCtrlWidth,$iuiCtrlHeight)
  $idPrivKey=GUICtrlCreateInput("",$iuiMargin,($iuiMargin+$iuiCtrlHeight)*4,$iuiCtrlWidth,$iuiCtrlHeight)
  $idUser=GUICtrlCreateInput("",$iuiMargin,($iuiMargin+$iuiCtrlHeight)*5,$iuiCtrlWidth,$iuiCtrlHeight)
  $idCancelBtn=GUICtrlCreateButton("Cancel",$iuiBtnLeft,($iuiMargin+$iuiCtrlHeight)*6,$iuiBtnWidth,$iuiCtrlHeight)
  $idDoneBtn=GUICtrlCreateButton("Continue",$iuiBtnLeft+$iuiBtnWidth+$iuiMargin,($iuiMargin+$iuiCtrlHeight)*6,$iuiBtnWidth,$iuiCtrlHeight)
  $hUser=ControlGetHandle($hAuthWnd,"",$idUser)
  $hPrivKey=ControlGetHandle($hAuthWnd,"",$idPrivKey)
  $hCompany=ControlGetHandle($hAuthWnd,"",$idCompany)
  $hClientId=ControlGetHandle($hAuthWnd,"",$idClientId)
  $hPubKey=ControlGetHandle($hAuthWnd,"",$idPubKey)
  _GUICtrlEdit_SetCueBanner($hCompany,"Company",True)
  _GUICtrlEdit_SetCueBanner($hClientId,"ClientId",True)
  _GUICtrlEdit_SetCueBanner($hPubKey,"PubKey",True)
  _GUICtrlEdit_SetCueBanner($hPrivKey,"PrivKey",True)
  _GUICtrlEdit_SetCueBanner($hUser,"User Identifier",True)
  $hStatus=_GUICtrlStatusBar_Create($hAuthWnd,-1,"Initializing...")
  GUICtrlSetState($idDoneBtn,$GUI_DISABLE)
  $g_cwmAuth_idDoneBtn=$idDoneBtn
  Dim $g_cwmAuth_aidInput[]=[5,$idCompany,$idClientId,$idPubKey,$idPrivKey,$idUser]
  Dim $g_cwmAuth_abValidate[$g_cwmAuth_aidInput[0]]
  GUIRegisterMsg($WM_COMMAND,"_cwmAuth_WM_COMMAND")
  GUISetState(@SW_SHOW)
  _GUICtrlStatusBar_SetText($hStatus,"Ready")
  While 1
    $nMsg=GUIGetMsg()
    Switch $nMsg
      Case $GUI_EVENT_CLOSE,$idCancelBtn
        GUIRegisterMsg($WM_COMMAND,'')
        GUIDelete($hAuthWnd)
        If $bFatal Then
          _Log("~!bFatal@_cwmAuth")
          _Exit()
        EndIf
        Return SetError(0,1,0)
      Case $idDoneBtn
        GUIRegisterMsg($WM_COMMAND,'')
        GUICtrlSetState($idDoneBtn,$GUI_DISABLE)
        _GUICtrlStatusBar_SetText($hStatus,"Please Wait...")
        $g_cwm_sCompany=GUICtrlRead($idCompany)
        $g_cwm_sClientId=_Base64Encode(_CryptProtectData(GUICtrlRead($idClientId)))
        $g_cwm_sClientId_Crypt=$g_cwm_sClientId
        $g_cwm_sPrivKey=_Base64Encode(_CryptProtectData(GUICtrlRead($idPrivKey)))
        $g_cwm_sPubKey=_Base64Encode(_CryptProtectData(GUICtrlRead($idPubKey)))
        $g_cwm_sUser=GUICtrlRead($idUser)
        _cwmInit()
        If @error Then
          If @extended=72 Then
            _GUICtrlStatusBar_SetText($hStatus,"Error: Invalid Company!")
            GUICtrlSetState($idDoneBtn,$GUI_ENABLE)
            GUIRegisterMsg($WM_COMMAND,"_cwmAuth_WM_COMMAND")
            ContinueLoop
          Else
            _GUICtrlStatusBar_SetText($hStatus,"Error: "&(@extended*10)+1&",cannot continue!")
            $bFatal=True
            For $i=1 To $g_cwmAuth_aidInput[0]
              GUICtrlSetState($g_cwmAuth_aidInput[$i],$GUI_DISABLE)
            Next
            GUICtrlSetData($idCancelBtn,"Exit")
            ContinueLoop
          EndIf
        EndIf
        $jRet=_cwmCall("/system/info/?fields=version",True)
        If @error Then
          If @extended==4 Then
            _GUICtrlStatusBar_SetText($hStatus,"Error: "&_JSON_Get($g_cwm_jLastRet,"code")&","&_JSON_Get($g_cwm_jLastRet,"message"))
            GUICtrlSetState($idDoneBtn,$GUI_ENABLE)
            GUIRegisterMsg($WM_COMMAND,"_cwmAuth_WM_COMMAND")
            ContinueLoop
          Else
            _GUICtrlStatusBar_SetText($hStatus,"Error: "&(@extended*10)+2&",cannot continue!")
            $bFatal=True
            For $i=1 To $g_cwmAuth_aidInput[0]
              GUICtrlSetState($g_cwmAuth_aidInput[$i],$GUI_DISABLE)
            Next
            GUICtrlSetData($idCancelBtn,"Exit")
            ContinueLoop
          EndIf
        EndIf
        _GUICtrlStatusBar_SetText($hStatus,"API Keys Accepted.")
        Sleep(1000)
        _GUICtrlStatusBar_SetText($hStatus,"Encrypting/Saving Keys...")
        Sleep(1000)
        _saveConfig()
        _GUICtrlStatusBar_SetText($hStatus,"Encrypting/Saving Keys...Done")
        Sleep(1000)
        _GUICtrlStatusBar_SetText($hStatus,"Configuring...")
        cwInstall()
        GUIDelete($hAuthWnd)
        Return SetError(0,0,0)
    EndSwitch
  WEnd

EndFunc   ;==>_cwmAuth

Func _cwmAuth_WM_COMMAND($hWnd,$iMsg,$wParam,$lParam)
  Local $iCode,$inID,$bMod=False
  $iId=BitAND($wParam,0xFFFF)
  $iCode=BitShift($wParam,16)
  If $iCode==$EN_CHANGE Then
    For $i=1 To $g_cwmAuth_aidInput[0]
      If $iId<>$g_cwmAuth_aidInput[$i] Then ContinueLoop
      $bMod=True
      $g_cwmAuth_abValidate[$i-1]=GUICtrlRead($g_cwmAuth_aidInput[$i])<>""
    Next
  EndIf
  If $bMod Then
    $iEn=0
    For $i=1 To $g_cwmAuth_aidInput[0]
      If $g_cwmAuth_abValidate[$i-1] Then $iEn+=1
    Next
    ConsoleWrite($iEn&@CRLF)
    If $iEn=$g_cwmAuth_aidInput[0] Then
      GUICtrlSetState($g_cwmAuth_idDoneBtn,$GUI_ENABLE)
    Else
      GUICtrlSetState($g_cwmAuth_idDoneBtn,$GUI_DISABLE)
    EndIf
  EndIf
  Return $GUI_RUNDEFMSG
EndFunc   ;==>_cwmAuth_WM_COMMAND

Func _loadConfig()
  DirCreate($gsDataDir)
  $g_cwm_sCompany=IniRead($gsConfigFile,"API","Company","")
  $g_cwm_sClientId=IniRead($gsConfigFile,"API","ClientId","")
  $g_cwm_sPrivKey=IniRead($gsConfigFile,"API","PrivKey","")
  $g_cwm_sPubKey=IniRead($gsConfigFile,"API","PubKey","")
  $g_cwm_sUser=IniRead($gsConfigFile,"Notifier","User","")
EndFunc   ;==>_loadConfig

Func _saveConfig()
  DirCreate($gsDataDir)
  IniWrite($gsConfigFile,"API","Company",$g_cwm_sCompany)
  IniWrite($gsConfigFile,"API","ClientId",$g_cwm_sClientId_Crypt)
  IniWrite($gsConfigFile,"API","PrivKey",$g_cwm_sPrivKey)
  IniWrite($gsConfigFile,"API","PubKey",$g_cwm_sPubKey)
  IniWrite($gsConfigFile,"Notifier","User",$g_cwm_sUser)
  $g_cwm_sPrivKey=_RandStr()
  $g_cwm_sPubKey=_RandStr()
EndFunc   ;==>_saveConfig

Func _loadState()
  DirCreate($gsDataDir)
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
  $aTiksLast=$aConfigNew
EndFunc   ;==>_loadState

Func _saveState()
  _Log("Saving State...")
  DirCreate($gsDataDir)
  For $i=1 To $aTiksLast[0][0]
    IniWrite($gsStateFile,$aTiksLast[$i][0],"LastMod",$aTiksLast[$i][1])
    IniWrite($gsStateFile,$aTiksLast[$i][0],"jTik",_Base64Encode(_JSON_Generate($aTiksLast[$i][2])))
    IniWrite($gsStateFile,$aTiksLast[$i][0],"jLastNote",_Base64Encode(_JSON_Generate($aTiksLast[$i][3])))
    IniWrite($gsStateFile,$aTiksLast[$i][0],"Type",$aTiksLast[$i][4])
  Next
  _Log("Saving State...Done")
EndFunc   ;==>_saveState

Func _InitHttp()
  $g_cwm_hHttp=_WinHttpOpen()
  If @error Then Return SetError(1,1,0)
  _WinHttpSetTimeouts(10000,60000,60000,60000)
  Return SetError(0,0,1)
EndFunc   ;==>_InitHttp

Func _HttpGet($sUrl,$sDomain,$aHeaders=Null)
  If Not IsHWnd($g_cwm_hHttp) Then
    _InitHttp()
    If @error Then Return SetError(1,1,0)
  EndIf
  $g_cwm_hConnect=_WinHttpConnect($g_cwm_hHttp,$sDomain,$INTERNET_DEFAULT_HTTPS_PORT)
  If @error Then Return SetError(1,2,0)
  $sHeaders=''
  If $aHeaders<>Null And IsArray($aHeaders) Then
    If UBound($aHeaders,0)<>2 Then Return SetError(1,3,0)
    If UBound($aHeaders,1)<2 Then Return SetError(1,4,0)
    For $i=1 To $aHeaders[0][0]
      $sHeaders&=StringFormat("%s: %s",$aHeaders[$i][0],$aHeaders[$i][1])&@CRLF
    Next
  EndIf
  $vRet=_WinHttpSimpleSSLRequest($g_cwm_hConnect,"GET",$sUrl,Default,Default,$sHeaders,True)
  If @error Then
    _Log($vRet&@CRLF)
    Return SetError(1,(6*1000)+@error,0)
  EndIf
  If StringInStr($vRet[0],"404 Not Found") Then Return SetError(1,7,0)
  Return SetError(0,0,$vRet[1])
EndFunc   ;==>_HttpGet

Func _cwmInit()
  $sCI=_HttpGet("/login/companyinfo/"&$g_cwm_sCompany,"na.myconnectwise.net")
  If @error Then Return SetError(1,(@extended*10)+1,0)
  $g_cwm_sCI=$sCI
  $jCI=_JSON_Parse($sCI)
  If @error Then Return SetError(1,2,0)
  $sCompany=_JSON_Get($jCI,"CompanyName")
  If @error Then Return SetError(1,3,0)
  $sCodeBase=_JSON_Get($jCI,"Codebase")
  If @error Then Return SetError(1,4,0)
  $sSiteUrl=_JSON_Get($jCI,"SiteUrl")
  If @error Then Return SetError(1,5,0)
  $g_cwm_sCodeBase=$sCodeBase
  $g_cwm_sSiteUrl=$sSiteUrl
  $g_cwm_sApiUri='/'&$sCodeBase&"apis/3.0"
  $sPubKey=_CryptUnprotectData(_Base64Decode($g_cwm_sPubKey)) ;$g_cwm_sPubKey;
  $sPrivKey=_CryptUnprotectData(_Base64Decode($g_cwm_sPrivKey)) ;$g_cwm_sPrivKey;
  $g_cwm_sAuthToken=_Base64Encode($sCompany&'+'&$sPubKey&':'&$sPrivKey)
  $g_cwm_sClientId=_CryptUnprotectData(_Base64Decode($g_cwm_sClientId))
EndFunc   ;==>_cwmInit

Func _cwmCall($sCall,$bApi=False)
  If $g_cwm_sAuthToken=='' Then Return SetError(1,1,0)
  While doOfflineCheck()
    Sleep(1000)
  WEnd
  If StringLen($sCall)>=1000 Then ; 2000 is the API limit.
    _Log("Warn: _cwmGetTickets $sCall>=1000")
  EndIf
  $iTimerA=TimerInit()
  Local $sRet,$jRet,$aHeader[][2]=[ _
      [2,''], _
      ["Authorization","Basic "&$g_cwm_sAuthToken], _
      ["clientid",$g_cwm_sClientId] _
  ]
  $iTimerB=TimerInit()
  $sRet=_HttpGet($g_cwm_sApiUri&$sCall,$g_cwm_sSiteUrl,$aHeader)
  If @error Then Return SetError(1,(@extended*10)+2,0)
  ;_Log("_cwmCall._HttpGet took "&TimerDiff($iTimerB)&@CRLF)
  $iTimerC=TimerInit()
  $jRet=_JSON_Parse($sRet)
  If @error Then Return SetError(1,3,$sRet)
  If IsObj($jRet) Then
    If $jRet.Exists("code") Then
      _Log($sRet&@CRLF)
      $g_cwm_jLastRet=$jRet
      If $bApi Then Return SetError(1,4,0)
      _Log("The Server Returned an Error..."&@LF&@LF&"Code:    "&@TAB&_JSON_Get($jRet,"code")&@LF&"Message: "&_JSON_Get($jRet,"message"))
      Return SetError(1,4,0)
    EndIf
  EndIf
  ;_Log("_cwmCall._JSON_Parse took "&TimerDiff($iTimerC)&@CRLF)
  ;_Log("_cwmCall took "&TimerDiff($iTimerA)&@CRLF)
  Return SetError(0,0,$jRet)
EndFunc   ;==>_cwmCall

Func _cwmGetTicketList(ByRef $aTik,$sType)
  Local $jTik,$sQuery=''
  For $i=1 To $aTik[0][0]
    If $sType<>$aTik[$i][1] Then ContinueLoop
    $sQuery&="id="&$aTik[$i][0]
    If $i<$aTik[0][0] Then $sQuery&=" Or "
  Next
  If StringRight($sQuery,4)==" Or " Then $sQuery=StringTrimRight($sQuery,4)
  ;_Log($sQuery&@CRLF)
  $jTik=_cwmCall("/"&$sType&"/tickets?conditions="&$sQuery&"&pageSize=1000")
  If @error Then Return SetError(1,(@extended*10)+1,0)
  ;_Log(_JSON_Generate($jTik)&@CRLF)
  Return SetError(0,0,$jTik)
EndFunc   ;==>_cwmGetTicketList

Func _cwmGetTikNfo($iType,$sId)
  Local $sType=$iType==0 ? "service" : "project"
  $jTik=_cwmCall('/'&$sType&'/tickets?conditions=id='&$sId&'&fields=id,_info/lastUpdated&pageSize=1')
  If @error Then Return SetError(1,(@extended*10+1),0)
  If IsArray($jTik) Then $jTik=$jTik[0]
  Return SetError(0,0,$jTik)
EndFunc   ;==>_cwmGetTikNfo

Func _cwmGetTiks(ByRef $aTikNfo,$iType,$sUser)
  Local $sType=$iType==0 ? "service" : "project"
  $jTik=_cwmCall('/'&$sType&'/tickets?conditions=closedFlag=False and resources contains "'&$sUser&'" or closedFlag=False and owner/identifier="'&$sUser&'"'&'&fields=id,_info/lastUpdated&pageSize=1000')
  If @error Then Return SetError(1,(@extended*10+1),0)
EndFunc   ;==>_cwmGetTiks

Func _cwmProcTik(ByRef $aTikNfo,$iType,$t)
  Local $sType=$iType==0 ? "service" : "project"
  ;_Log(_JSON_Generate($t))
  $vTikId=_JSON_Get($t,"id")
  If @error Then
    _Log("Failed to get ticket id"&@CRLF)
    _Log(_JSON_Generate($t)&@CRLF)
    If @error Then Return    ; SetError(1,2,$t)
  EndIf
  $vTikLastUpdate=_JSON_Get($t,"_info.lastUpdated")
  If @error Then
    _Log("Failed to get ticket _info.lastUpdated"&@CRLF)
    _Log(_JSON_Generate($t)&@CRLF)
    If @error Then Return    ; SetError(1,2,$t)
  EndIf
  $vTikLastUpdate=_cwmConvDate2Sec($vTikLastUpdate)
  $iIdx=_cwmInArray($aTiks,$vTikId)
  If @error Then
    $iIdx=UBound($aTikNfo,1)
    ReDim $aTikNfo[$iIdx+1][6]
    _Log("New Ticket:"&$vTikId&@CRLF)
    $aTikNfo[$iIdx][0]=$vTikId
    $aTikNfo[$iIdx][1]=$vTikLastUpdate
    $aTikNfo[$iIdx][4]=$iType
    $aTikNfo[0][0]=$iIdx
  Else
    If ($aTikNfo[$iIdx][1])==$vTikLastUpdate Then Return  ;ContinueLoop
    _Log('+Ticket Updated: '&$vTikId&"("&($aTikNfo[$iIdx][1])&','&$vTikLastUpdate&")"&@CRLF)
  EndIf
  _Log("Fetch Ticket:"&$vTikId&@CRLF)
  ;Sleep(50)
  $jFetch=_cwmCall('/'&$sType&"/tickets?conditions=id="&$vTikId&"&pageSize=1")
  If @error Then
    Return SetError(1,(@extended*10+3),0)
  Else
    If IsArray($jFetch) Then
      $aTikNfo[$iIdx][2]=$jFetch[0]
    Else
      $aTikNfo[$iIdx][2]=$jFetch
    EndIf
  EndIf
  ;ConsoleWrite(_JSON_Generate($jFetch)&@CRLF)
  _Log("Get Ticket Notes:"&$vTikId&@CRLF)
  ;Sleep(50)
  $jTikNotes=_cwmCall('/'&$sType&"/tickets/"&$vTikId&"/allNotes?orderBy=_info/sortByDate desc&pageSize=1")
  If @error Then
    _Log("Warn: Cannot retrieve notes for: "&$vTikId&" (Error: "&@extended&')'&@CRLF)
  Else
    If IsArray($jTikNotes) Then
      If UBound($jTikNotes,1)<>1 Then
        _Log("Warn: Cannot retrieve notes for: "&$vTikId&" (Error: "&@extended&')'&@CRLF)
      Else
        $aTikNfo[$iIdx][3]=$jTikNotes[0]
      EndIf
    Else
      $aTikNfo[$iIdx][3]=$jTikNotes
    EndIf
  EndIf
  $aTikNfo[$iIdx][1]=$vTikLastUpdate
  ;Sleep(50)
EndFunc   ;==>_cwmProcTik

Func _cwmGetTickets(ByRef $aTikNfo,$iType,$sUser)
  Local $iTimer=TimerInit()
  Local $sType=$iType==0 ? "service" : "project"
  Local $aPastIds[]=[0]
  For $i=1 To $aTikNfo[0][0]
    If $iType<>$aTikNfo[$i][4] Then ContinueLoop
    $iMax=UBound($aPastIds,1)
    ReDim $aPastIds[$iMax+1]
    $aPastIds[$iMax]=_cwmGetTikNfo($iType,$aTikNfo[$i][0])
    If @error Then
      _Log("_cwmGetTickets took "&TimerDiff($iTimer))
      Return SetError(1,(@extended*10+1),0)
    EndIf
    $aPastIds[0]=$iMax
  Next
  $jTik=_cwmCall('/'&$sType&'/tickets?conditions=closedFlag=False and resources contains "'&$sUser&'" or closedFlag=False and owner/identifier="'&$sUser&'"&fields=id,_info/lastUpdated&pageSize=1000')
  If @error Then
    _Log("_cwmGetTickets took "&TimerDiff($iTimer))
    Return SetError(1,(@extended*10+2),0)
  EndIf
  For $t In $jTik
    If $bExit Then
      _Exit()
    EndIf
    _cwmProcTik($aTikNfo,$iType,$t)
  Next
  For $i=1 To $aPastIds[0]
    If $bExit Then
      _Exit()
    EndIf
    _cwmProcTik($aTikNfo,$iType,$aPastIds[$i])
  Next
  _Log("_cwmGetTickets took "&TimerDiff($iTimer))
EndFunc   ;==>_cwmGetTickets

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

Func _cwmIsInArray(ByRef $aArr,$vObj,$iDim=0,$iStart=1)
  Local $iIdx=_cwmInArray($aArr,$vObj,$iDim,$iStart)
  If @error Then Return SetError(1,0,0)
  Return SetError(0,$iIdx,1)
EndFunc   ;==>_cwmIsInArray

Func _cwmInArray(ByRef $aArr,$vObj,$iDim=0,$iStart=1)
  Switch UBound($aArr,0)
    Case 1
      For $i=$iStart To $aArr[0]
        If $aArr[$i]==$vObj Then Return SetError(0,0,$i)
      Next
    Case 2
      For $i=$iStart To $aArr[0][$iDim]
        If $aArr[$i][$iDim]==$vObj Then Return SetError(0,0,$i)
      Next
  EndSwitch
  Return SetError(1,0,0)
EndFunc   ;==>_cwmInArray

Func _isWindowsLocked()
  If _WinAPI_OpenInputDesktop() Then Return False
  Return True
EndFunc   ;==>_isWindowsLocked

Func _TrayEvent()
  $iTrayMsg=TrayGetMsg()
  Switch $iTrayMsg
    Case $idTrayExit
      AdlibUnRegister("_TrayEvent")
      $bExit=True
      _Toast_Hide()
      _Exit()
  EndSwitch
EndFunc   ;==>_TrayEvent

Func _Exit()
  $bExit=True
  _Toast_Set(0,-1,-1,-1,-1,-1,"Consolas",125,125)
  $aRet=_Toast_ShowMod(0,$sTitle,"Exiting...                        ",-5)
  _Toast_Hide()
  _Log("_Exit()")
  Exit
EndFunc   ;==>_Exit

Func _getIdealSrv($sDom)
  ; Resolve
  Local $aRet[][3]=[[0,Null]]
  Local $iMax,$iRet
  $aDig=_nDig($sDom,$g_aResolveSrvDNS[0][0],"A",250)
  $iMax=UBound($aDig,1)
  If $iMax==1 Then Return SetError(1,1,0) ;DNS didnt reply
  If $iMax==2 And $aDig[1][4]=="SOA" Then Return SetError(1,2,0) ;Got Site of Authority,bad subdomain?
  ;_DebugArrayDisplay($aDig)
  For $i=1 To $iMax-1
    If $aDig[$i][4]=="A" Then
      ; Attempt Ping.
      $aPing=_nPing($aDig[$i][5])
      If @error Then
        _Log(StringFormat("Failed to ping %s,Error: %d",$aDig[1][5],@extended)&@CRLF)
        ContinueLoop
      EndIf
      If $aPing[1]==0 Then ContinueLoop
      $iRet=UBound($aRet,1)
      ReDim $aRet[$iRet+1][3]
      $aRet[$iRet][0]=$aDig[$i][5]
      $aRet[$iRet][1]=$aPing[2]
    EndIf
  Next
  $aRet[0][0]=$iRet
  _ArraySort($aRet,0,0,0,1)
  ;_DebugArrayDisplay($aRet)
  If $iRet=0 Then Return SetError(1,3,0) ;No Results.
  Return SetError(0,0,$aRet)
EndFunc   ;==>_getIdealSrv

Func doOfflineCheck()
  If $bExit Then
    While Sleep(1000)
    WEnd
  EndIf
  $bOffline=False
  $iOffline=0
  While $iOffline<$iOfflineThresh
    $aPing=_nPing($g_cwm_sSite)
    If @error Then
      $bOffline=True
      $iOffline+=1
    Else
      If $aPing[1]==0 Then
        $bOffline=True
        $iOffline+=1
      Else
        $bOffline=False
        ExitLoop
      EndIf
    EndIf
  WEnd
  If $bOfflineLast<>$bOffline Then
    $bOfflineLast=$bOffline
    If $bOffline Then
      _Log("We're Offline")
      _Toast_ShowMod(0,$sTitle,"We're Offline                    ",-5)
      _Toast_Hide()
    Else
      _Log("We're Online")
      _Toast_ShowMod(0,$sTitle,"We're back Online                    ",-5)
      _Toast_Hide()
    EndIf
  EndIf
  Return $bOffline
EndFunc   ;==>doOfflineCheck

; Retireve and Format Fields.
Func _tikGetFields(ByRef $tData, ByRef $aFields, ByRef $sFields)
  Local $vData
  $aFields=StringSplit($sFields,',')
  _ArrayColInsert($aFields,1)
  For $j=1 To $aFields[0][0]
    $vData=_JSON_Get($tData,$aFields[$j][0])
    Switch $aFields[$j][0]
      Case "owner.name"
        If $vData='' Then $vData="(Unassigned)"
      Case "_info.dateEntered"
        $vData=_cwmConvDate2Read($vData)
      Case "_info.lastUpdated"
        $vData=_cwmConvDate2Read($vData)
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

Func _PurgeOldTiks()
  If $bFirstRun Then Return SetError(1,1,0)
  Local $iMax=0
  $iyMax=UBound($aTiksLast,2)
  $iOld=$aTiksLast[0][0]
  Dim $aKeepTiks[1][$iyMax]
  $aKeepTiks[0][0]=0
  $iLastWeek=_DateDiff('s',$g_cwm_sEpoch,_DateAdd('D',-7,_NowCalc()))
  For $i=1 To $aTiksLast[0][0]
    $jTik=$aTiksLast[$i][2]
    ;_Log($aTiksLast[$i][0]&':'&_JSON_Get($jTik,"closedFlag")&@CRLF)
    If StringInStr(_JSON_Get($jTik,"closedFlag"),"true") And $aTiksLast[$i][1]<$iLastWeek Then
      _Log("Dropping Old/Closed Ticket: "&$aTiksLast[$i][0]&' ( '&_cwmConvDate2Read(_JSON_Get($jTik,"_info.lastUpdated"))&' <<< '&_cwmConvDate2Read(_cwmConvSec2Date($iLastWeek))&' )')
      ContinueLoop
    EndIf
    $iMax=UBound($aKeepTiks,1)
    ReDim $aKeepTiks[$iMax+1][$iyMax]
    For $y=0 To $iyMax-1
      $aKeepTiks[$iMax][$y]=$aTiksLast[$i][$y]
    Next
  Next
  Dim $aTiksLast[$iMax+1][$iyMax]
  For $x=0 To $iMax
    For $y=0 To $iyMax-1
      $aTiksLast[$x][$y]=$aKeepTiks[$x][$y]
    Next
  Next
  If $iMax=0 Then Return SetError(1,2,0)
  $aTiksLast[0][0]=$iMax
  $aTiks=$aTiksLast
  If $iOld<>$iMax Then Return 1
  Return 0
  ;_DebugArrayDisplay($aTiksLast)
EndFunc   ;==>_PurgeOldTiks

Func _isFieldNoMod($s)
  For $i=1 To $aNoModFields[0]
    If StringCompare($s,$aNoModFields[$i])==0 Then Return True
  Next
  Return False
EndFunc   ;==>_isFieldNoMod


;~ Global $iUi_MonRect[4],$iUi_WidthMax,$iUi_WidthMin,$iUi_HeightMax,$iUi_FontSize,$iUi_LabelMaxWidth
;~ Global $iUi_FontName="Consolas"
;~ Opt("GuiOnEventMode",1)
;~ Func _uiCalc()
;~     $iUi_FontSize=18
;~     ; Get Monitor Info
;~     Local $tPos=_WinAPI_GetMousePos()
;~     Local $hMonitor=_WinAPI_MonitorFromPoint($tPos)
;~     Local $aMonitor=_WinAPI_GetMonitorInfo($hMonitor)
;~     If @error Then Return SetError(1,0,0)
;~     For $i=0 To 3
;~         ;LTRM
;~         $iUi_MonRect[$i]=DllStructGetData($aMonitor[1],$i+1)
;~     Next
;~     Local $iWkSpWidth=($iUi_MonRect[0]-$iUi_MonRect[2])
;~     Local $iWkSpHeight=($iUi_MonRect[1]-$iUi_MonRect[3])
;~     $iUi_WidthMax=$iWkSpWidth/3
;~     $iUi_HeightMax=$iWkSpHeight/3
;~     $iUi_WidthMin=$iWkSpWidth/6
;~     $iUi_LabelMaxWidth=$iUi_WidthMax

;~ EndFunc

;~ Dim $aLastNote[]=[ _
;~     "Tara Butts", _
;~     "Hi,I tried my best to trouble shoot the computers at YMCA Brc and have no luck."&@CRLF& _
;~     ""&@CRLF& _
;~     "we have like over 10 computers down."&@CRLF& _
;~     ""&@CRLF& _
;~     "Best,"&@CRLF& _
;~     "Tara"&@CRLF& _
;~     ""&@CRLF& _
;~     "Tara N Butts,"&@CRLF& _
;~     "Administrator Assistant "&@CRLF& _
;~     "BRC Vanderbilt Stabilization Program-YMCA "&@CRLF& _
;~     "224 47th St "&@CRLF& _
;~     "New York,NY 10017  "&@CRLF& _
;~     "WP 646-841-4323 "&@CRLF& _
;~     "E: Tbutts@brc.org   "&@CRLF& _
;~     "Please note I am in the office Monday Through Friday,"&@CRLF& _
;~     "9am-5:30pm "&@CRLF& _
;~     "You can have anything you want in life if you dress for it... "&@CRLF& _
;~     " "&@CRLF& _
;~     "This e-mail message and any attachments may contain confidential information meant solely for the intended recipient. If you have received this message in error,please notify the sender immediately by replying to this message,then delete the e-mail and any attachments from your system. Any use of this message or its attachments that is not in keeping with the confidential nature of the information,including but not limited to disclosing information to others,dissemination,distribution,or copying is strictly prohibited. "&@CRLF _
;~ ]
;~ ;",,,,,,,,,,,,"
;~ Dim $aFields[][2]=[[13,''], _
;~     ["_info.lastUpdated","2024.07.23@1151a"], _
;~     ["id","1937290"], _
;~     ["status.name","Needs Escalation"], _
;~     ["owner.name","Luis Perez"], _
;~     ["summary","CrowdStrike issue remediation"], _
;~     ["company.name","BRC-ADMINISTRATION"], _
;~     ["contact.name","Thomas Wyse"], _
;~     ["subType.name","Workstation"], _
;~     ["item.name","BSOD"], _
;~     ["priority.name","Urgent"], _
;~     ["severity.name","Medium"], _
;~     ["type.name","Problem"], _
;~     ["_info.updatedBy","Corsica"] _
;~ ]
;~ ; $iFlag is for Ignoring or other modifiers.
;~ ; $hwndNotify,Ticket#,$iFlag,tSnoozeTimer,Title,$aFields,$aLastNote
;~ Dim $aN[]=[ _
;~     -1, _
;~     1937290, _
;~     0, _
;~     0, _
;~     "Ticket Updated", _
;~     $aFields, _
;~     $aLastNote _
;~ ]
;~ Func _ShowNotify(ByRef $aNotify)
;~     Local $iWidth,$iHeight
;~     $aNotify[0]="NaN"
;~     $hWnd=GUICreate("",$iToast_Width,$iToast_Height,$aToast_Data[0],$aToast_Data[1]-2,$WS_POPUPWINDOW,BitOR($WS_EX_TOOLWINDOW,$WS_EX_TOPMOST))
;~     $aNotify[0]=$hWnd

;~     ; Label Generation
;~     ;Local $iLabelWidth=_StringSize($aNotify[1],$iUi_FontSize,Default,Default,$iUi_FontName,$iUi_LabelMaxWidth)
;~     ;If @error Then
;~     ;    _Log(StringFormat("~!Error@_StringSize(%s,%s,%s,%s),",$aNotify[1],$iUi_FontSize,$iUi_FontName,$iUi_LabelMaxWidth,@extended))
;~     ;    Return
;~     ;EndIf

;~ EndFunc

; Author ........: Melba23,based on some original code by GioVit for the Toast
; Modified By ...: BiatuAutMiahn
Func _Toast_ShowMod($vIcon,$sTitle,$sMessage,$iDelay=0,$fWait=True,$bisTicket=False,$fRaw=False)
  $bInfinite=False
  If $iDelay=Null Then
    $bInfinite=True
    $iDelay+=4294967295
  EndIf
  $iToast_Font_Size=10
  $sToast_Font_Name="Consolas"
  $fToast_OpenTik=False
  ; If previous Toast retracting must wait until process is completed
  If $fToast_Retracting Then
    ; Store parameters
    $vIcon_Retraction=$vIcon
    $sTitle_Retraction=$sTitle
    $sMessage_Retraction=$sMessage
    $iDelay_Retraction=$iDelay
    $fWait_Retraction=$fWait
    $fRaw_Retraction=$fRaw
    ; Keep looking to see if previous Toast retracted
    AdlibRegister("__Toast_Retraction_Check",100)
    ; Explain situation to user
    Return SetError(5,0,-1)
  EndIf
  ; Store current GUI mode and set Message mode
  Local $nOldOpt=Opt('GUIOnEventMode',0)
  ; Retract any Toast already in place
  If $hToast_Handle<>0 Then _Toast_Hide()
  ; Reset non-reacting Close [X] ControlID
  $hToast_Close_X=9999
  ; Set default auto-sizing Toast widths
  Local $iToast_Width_max=500
  Local $iToast_Width_min=150
  ; Check for icon
  Local $iIcon_Style=0
  Local $iIcon_Reduction=36
  Local $sDLL="user32.dll"
  Local $sImg=""
  If StringIsDigit($vIcon) Then
    Switch $vIcon
      Case 0
        $iIcon_Reduction=0
      Case 8
        $sDLL="imageres.dll"
        $iIcon_Style=78
      Case 16 ; Stop
        $iIcon_Style=-4
      Case 32 ; Query
        $iIcon_Style=-3
      Case 48 ; Exclam
        $iIcon_Style=-2
      Case 64 ; Info
        $iIcon_Style=-5
      Case Else
        Return SetError(1,0,-1)
    EndSwitch
  Else
    If StringInStr($vIcon,"|") Then
      $iIcon_Style=StringRegExpReplace($vIcon,"(.*)\|","")
      $sDLL=StringRegExpReplace($vIcon,"\|.*$","")
    Else
      Switch StringLower(StringRight($vIcon,3))
        Case "exe","ico"
          $sDLL=$vIcon
        Case "bmp","jpg","gif","png"
          $sImg=$vIcon
      EndSwitch
    EndIf
  EndIf
  ; Determine max message width
  Local $iMax_Label_Width=$iToast_Width_max-20-$iIcon_Reduction
  If $fRaw=True Then $iMax_Label_Width=0
  ; Get message label size
  Local $aLabel_Pos=_StringSize($sMessage,$iToast_Font_Size,Default,Default,$sToast_Font_Name,$iMax_Label_Width,$hToast_Handle)
  If @error Then
    If @error=3 Then
      Local $iScale=$iMax_Label_Width
      Do
        $aLabel_Pos=_StringSize($sMessage,$iToast_Font_Size,Default,Default,$sToast_Font_Name,$iScale,$hToast_Handle)
        $iScale+=2
      Until @error<>3
      _Log("iScaleFix="&$iScale)
    Else
      $nOldOpt=Opt('GUIOnEventMode',$nOldOpt)
      Return SetError((@Error*10)+3,0,-1)
    EndIf
  EndIf
  ; Reset text to match rectangle
  $sMessage=$aLabel_Pos[0]
  ;Set line height for this font
  Local $iLine_Height=$aLabel_Pos[1]
  ; Set label size
  Local $iLabelwidth=$aLabel_Pos[2]
  Local $iLabelheight=$aLabel_Pos[3]
  ; Set Toast size
  Local $iToast_Width=$iLabelwidth+20+$iIcon_Reduction
  ; Check if Toast will fit on screen
  If $iToast_Width>@DesktopWidth-20 Then
    $nOldOpt=Opt('GUIOnEventMode',$nOldOpt)
    Return SetError(4,0,-1)
  EndIf
  ; Increase if below min size
  If $iToast_Width<$iToast_Width_min+$iIcon_Reduction Then
    $iToast_Width=$iToast_Width_min+$iIcon_Reduction
    $iLabelwidth=$iToast_Width_min-20
  EndIf
  ; Set title bar height-with minimum for [X]
  Local $iTitle_Height=0
  If $sTitle="" Then
    If $iDelay<0 Then $iTitle_Height=6
  Else
    $iTitle_Height=$iLine_Height+2
    If $iDelay<0 Then
      If $iTitle_Height<17 Then $iTitle_Height=17
    EndIf
  EndIf
  ; Set Toast height as label height+title bar+bottom margin
  Local $iToast_Height=$iLabelheight+$iTitle_Height+20
  If $bisTicket Then $iToast_Height+=$iTitle_Height
  ; Ensure enough room for icon if displayed
  If $iIcon_Reduction Then
    If $iToast_Height<$iTitle_Height+42 Then $iToast_Height=$iTitle_Height+20
  EndIf
  $iTitle_Height+=2
  ; Get Toast starting position and direction
  Local $aToast_Data=__Toast_Locate($iToast_Width,$iToast_Height)
  ; Create Toast slice with $WS_POPUPWINDOW,$WS_EX_TOOLWINDOW style and $WS_EX_TOPMOST extended style
  ;
  $hToast_Handle=GUICreate("",$iToast_Width,$iToast_Height,$aToast_Data[0],$aToast_Data[1]-2,$WS_POPUPWINDOW,BitOR($WS_EX_TOOLWINDOW,$WS_EX_TOPMOST))
  If @error Then
    $nOldOpt=Opt('GUIOnEventMode',$nOldOpt)
    Return SetError(1,0,-1)
  EndIf
  GUISetFont($iToast_Font_Size,Default,Default,$sToast_Font_Name)
  GUISetBkColor($iToast_Message_BkCol)
  ; Set centring parameter
  Local $iLabel_Style=0 ;$SS_LEFT
  If BitAND($iToast_Style,1)=1 Then
    $iLabel_Style=1 ;$SS_CENTER
  ElseIf BitAND($iToast_Style,2)=2 Then
    $iLabel_Style=2 ;$SS_RIGHT
  EndIf
  ; Check installed fonts
  Local $sX_Font="WingDings"
  Local $sX_Char="x"
  Local $i=1
  While 1
    Local $sInstalled_Font=RegEnumVal("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts",$i)
    If @error Then ExitLoop
    If StringInStr($sInstalled_Font,"WingDings 2") Then
      $sX_Font="WingDings 2"
      $sX_Char="T"
    EndIf
    $i+=1
  WEnd
  ; Create title bar if required
  If $sTitle<>"" Then
    ; Create disabled background strip
    GUICtrlCreateLabel("",0,0,$iToast_Width,$iTitle_Height)
    GUICtrlSetBkColor(-1,$iToast_Header_BkCol)
    GUICtrlSetState(-1,128) ;$GUI_DISABLE
    ;If $bisTicket Then
    ;	GUICtrlCreateLabel("",0,$iToast_Height-$iTitle_Height,$iToast_Width,$iTitle_Height,$WS_CLIPSIBLINGS)
    ;	GUICtrlSetBkColor(-1,$iToast_Header_BkCol)
    ;	GUICtrlSetState(-1,128) ;$GUI_DISABLE
    ;EndIf
    ; Set title bar width to offset text
    Local $iTitle_Width=$iToast_Width-10
    ; Create closure [X] if needed
    If $iDelay<0 Then
      ; Create [X]
      Local $iX_YCoord=Int(($iTitle_Height-17)/2)
      $hToast_Close_X=GUICtrlCreateLabel("",$iToast_Width-18,$iX_YCoord,17,17)
      GUICtrlSetFont(-1,14,Default,Default,$sX_Font)
      GUICtrlSetBkColor(-1,-2) ;$GUI_BKCOLOR_TRANSPARENT
      GUICtrlSetColor(-1,$iToast_Header_Col)
      ; Reduce title bar width to allow [X] to activate
      $iTitle_Width -= 18
    EndIf
    ; Create Title label with bold text,centred vertically in case bar is higher than line
    GUICtrlCreateLabel($sTitle,10,0,$iTitle_Width,$iTitle_Height,0x0200) ;$SS_CENTERIMAGE
    GUICtrlSetBkColor(-1,$iToast_Header_BkCol)
    GUICtrlSetColor(-1,$iToast_Header_Col)
    If BitAND($iToast_Style,4)=4 Then GUICtrlSetFont(-1,$iToast_Font_Size,600)
  Else
    If $iDelay<0 Then
      ; Only need [X]
      $hToast_Close_X=GUICtrlCreateLabel($sX_Char,$iToast_Width-18,0,17,17)
      GUICtrlSetFont(-1,14,Default,Default,$sX_Font)
      GUICtrlSetBkColor(-1,-2) ;$GUI_BKCOLOR_TRANSPARENT
      GUICtrlSetColor(-1,$iToast_Message_Col)
    EndIf
  EndIf
  ; Create icon
  If $iIcon_Reduction Then
    Switch StringLower(StringRight($sImg,3))
      Case "bmp","jpg","gif"
        GUICtrlCreatePic($sImg,10,10+$iTitle_Height,32,32)
      Case "png"
        __Toast_ShowPNG($sImg,$iTitle_Height)
      Case Else
        GUICtrlCreateIcon($sDLL,$iIcon_Style,10,10+$iTitle_Height)
    EndSwitch
  EndIf
  ; Create Message label
  GUICtrlCreateLabel($sMessage,10+$iIcon_Reduction,10+$iTitle_Height,$iLabelwidth,$iLabelheight)
  GUICtrlSetStyle(-1,$iLabel_Style)
  If $iToast_Message_Col<>Default Then GUICtrlSetColor(-1,$iToast_Message_Col)
  $hToast_OpenTik=Null
  If $bisTicket Then
    Local $aColorsEx=[0x000000,0xFFFFFF,0xFFFFFF,0x000000,0xFFFFFF,0xFFFFFF,0x333333,0xFFFFFF,0xFFFFFF,0x666666,0xFFFFFF,0xFFFFFF]
    GuiFlatButton_SetDefaultColorsEx($aColorsEx)
    Local $iBtnWidth,$iBtnMax=3
    $iBtnWidth=($iToast_Width/$iBtnMax)
    Local $iBtnLeft=($iToast_Width/2)-(($iBtnWidth/2)*$iBtnMax)+1
    Local $sBtnA="Open Ticket"
    Local $aBtn_PosA=_StringSize($sBtnA,$iToast_Font_Size,Default,Default,$sToast_Font_Name,$iToast_Width/$iBtnMax)
    $hToast_OpenTik=GuiFlatButton_Create($sBtnA,$iBtnLeft,$iToast_Height-$iTitle_Height+1,$iBtnWidth-2,$iTitle_Height-2)
    GUICtrlSetFont($hToast_OpenTik,$iToast_Font_Size,Default,Default,"Consolas")
    Local $sBtnB="Dismiss"
    Local $aBtn_PosB=_StringSize($sBtnB,$iToast_Font_Size,Default,Default,$sToast_Font_Name,$iToast_Width/$iBtnMax)
    $hToast_Close_X=GuiFlatButton_Create($sBtnB,$iBtnLeft+$iBtnWidth,$iToast_Height-$iTitle_Height+1,$iBtnWidth-2,$iTitle_Height-2)
    GUICtrlSetFont($hToast_Close_X,$iToast_Font_Size,Default,Default,"Consolas")
    Local $sBtnB="Dismiss All"
    Local $aBtn_PosB=_StringSize($sBtnB,$iToast_Font_Size,Default,Default,$sToast_Font_Name,$iToast_Width/$iBtnMax)
    $hToast_DismissAll=GuiFlatButton_Create($sBtnB,$iBtnLeft+($iBtnWidth*2),$iToast_Height-$iTitle_Height+1,$iBtnWidth-2,$iTitle_Height-2)
    GUICtrlSetFont($hToast_DismissAll,$iToast_Font_Size,Default,Default,"Consolas")
    GuiFlatButton_SetState($hToast_OpenTik,$GUI_SHOW)
    GuiFlatButton_SetState($hToast_Close_X,$GUI_SHOW)
    GuiFlatButton_SetState($hToast_DismissAll,$GUI_SHOW)
  EndIf
  ; Slide Toast Slice into view from behind systray and activate
  DllCall("user32.dll","int","AnimateWindow","hwnd",$hToast_Handle,"int",$iToast_Time_Out,"long",$aToast_Data[2])
  ; Activate Toast without stealing focus
  GUISetState(@SW_SHOWNOACTIVATE,$hToast_Handle)
  _WinAPI_RedrawWindow($hToast_Handle)
  ; If script is to pause
  If $fWait=True Then
    ; Clear message queue
    Do
    Until GUIGetMsg()=0
    ; Begin timeout counter
    Local $iTimeout_Begin=TimerInit()
    ; Wait for timeout or closure
    Local $iMsg
    While Sleep(10)
      $iMsg=GUIGetMsg()
      If $iMsg=$hToast_Close_X Or (Not $bInfinite And TimerDiff($iTimeout_Begin)/1000>=Abs($iDelay)) Then
        ExitLoop
      ElseIf $iMsg=$hToast_DismissAll Then
        $fToast_bDismissAll=True
        ExitLoop
      ElseIf $iMsg=$hToast_OpenTik Then
        $fToast_OpenTik=True
        ExitLoop
      EndIf
    WEnd
    ; If script is to continue and delay has been set
  ElseIf (Not $bInfinite And Abs($iDelay)>0) Then
    ; Store timer info
    $iToast_Timer=Abs($iDelay*1000)
    $iToast_Start=TimerInit()
    ; Register Adlib function to run timer
    AdlibRegister("__Toast_Timer_Check",100)
    ; Register message handler to check for [X] click
    GUIRegisterMsg(0x0021,"__Toast_WM_EVENTSMod") ;$WM_MOUSEACTIVATE
  EndIf
  ; Reset original mode
  $nOldOpt=Opt('GUIOnEventMode',$nOldOpt)
  ; Create array to return Toast dimensions
  Local $aToast_Data[3]=[$iToast_Width,$iToast_Height,$iLine_Height]
  Return $aToast_Data
EndFunc   ;==>_Toast_ShowMod

Func __Toast_WM_EVENTSMod($hWnd,$Msg,$wParam,$lParam)
  #forceref $wParam,$lParam
  If $hWnd=$hToast_Handle Then
    If $Msg=0x0021 Then ; $WM_MOUSEACTIVATE
      ; Check mouse position
      Local $aPos=GUIGetCursorInfo($hToast_Handle)
      If $aPos[4]=$hToast_Close_X Then $fToast_Close=True
      If $aPos[4]=$hToast_OpenTik Then
        $fToast_OpenTik=True
        $fToast_Close=True
      EndIf
    EndIf
  EndIf
  Return 'GUI_RUNDEFMSG'
EndFunc   ;==>__Toast_WM_EVENTSMod


Func cwInstall()
  If @Compiled Then
    Local $bStartup,$bDesktop,$bStartMenu
    If MsgBox(32+4,$sTitle,"Would you like to run at startup?")==6 Then $bStartup=1
    If MsgBox(32+4,$sTitle,"Would you like to add to the desktop shortcut?")==6 Then $bDesktop=1
    If MsgBox(32+4,$sTitle,"Would you like to add to the Start Menu?")==6 Then $bStartMenu=1
    If $bStartup Or $bDesktop Or $bStartMenu Then
      FileCopy(@AutoItExe,$gsDataDir&"\cwNotify.exe",1)
    EndIf
    If $bStartup Then
      RegWrite("HKCU\Software\Microsoft\Windows\CurrentVersion\Run","cwNotify","REG_SZ",$gsDataDir&"\cwNotify.exe")
    EndIf
    If $bDesktop Then
      FileCreateShortcut($gsDataDir&"\cwNotify.exe",@DesktopDir&"\cwNotify.lnk",$gsDataDir)
    EndIf
    If $bStartMenu Then
      FileCreateShortcut($gsDataDir&"\cwNotify.exe",@ProgramsDir&"\cwNotify.lnk",$gsDataDir)
    EndIf
    If MsgBox(32+4,$sTitle,"Would you like to run now?")==6 Then
      Run($gsDataDir&"\cwNotify.exe ~!PostInstall",$gsDataDir,@SW_SHOW)
      Exit 0
    EndIf
  EndIf
EndFunc
