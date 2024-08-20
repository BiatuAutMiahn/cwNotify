; https://www.autoitscript.com/forum/topic/108969-subrogation/
#include <WinAPI.au3>
#include <Memory.au3>
#include <WinAPISys.au3>

Func _DllFromMemory($bBinaryImage,$sSubrogor="explorer.exe")
	Local $tBinary=DllStructCreate("byte["&BinaryLen($bBinaryImage)&"]")
	DllStructSetData($tBinary,1,$bBinaryImage)
	Local $pPointer=DllStructGetPtr($tBinary)
	Local $tIMAGE_DOS_HEADER=DllStructCreate("char Magic[2];"& _
			"word BytesOnLastPage;"& _
			"word Pages;"& _
			"word Relocations;"& _
			"word SizeofHeader;"& _
			"word MinimumExtra;"& _
			"word MaximumExtra;"& _
			"word SS;"& _
			"word SP;"& _
			"word Checksum;"& _
			"word IP;"& _
			"word CS;"& _
			"word Relocation;"& _
			"word Overlay;"& _
			"char Reserved[8];"& _
			"word OEMIdentifier;"& _
			"word OEMInformation;"& _
			"char Reserved2[20];"& _
			"dword AddressOfNewExeHeader", _
			$pPointer)
	; Move pointer
	$pPointer+=DllStructGetData($tIMAGE_DOS_HEADER,"AddressOfNewExeHeader");move to PE file header
	Local $sMagic=DllStructGetData($tIMAGE_DOS_HEADER,"Magic")
	; Check if it's valid format
	If Not ($sMagic=="MZ") Then
		Return SetError(1,0,0);MS-DOS header missing. Btw 'MZ' are the initials of Mark Zbikowski in case you didn't know.
	EndIf
	Local $tIMAGE_NT_SIGNATURE=DllStructCreate("dword Signature",$pPointer)
	; Move pointer
	$pPointer+=4;size of $tIMAGE_NT_SIGNATURE structure
	; Check signature
	If DllStructGetData($tIMAGE_NT_SIGNATURE,"Signature")<>17744 Then;IMAGE_NT_SIGNATURE
		Return SetError(2,0,0);wrong signature. For PE image should be "PE\0\0" or 17744 dword.
	EndIf
	; In place of IMAGE_FILE_HEADER structure
	Local $tIMAGE_FILE_HEADER=DllStructCreate("word Machine;"& _
			"word NumberOfSections;"& _
			"dword TimeDateStamp;"& _
			"dword PointerToSymbolTable;"& _
			"dword NumberOfSymbols;"& _
			"word SizeOfOptionalHeader;"& _
			"word Characteristics", _
			$pPointer)
	; Get number of sections
	Local $iNumberOfSections=DllStructGetData($tIMAGE_FILE_HEADER,"NumberOfSections")
	; Move pointer
	$pPointer+=20;size of $tIMAGE_FILE_HEADER structure
	; Determine the type
	Local $tMagic=DllStructCreate("word Magic;",$pPointer)
	Local $iMagic=DllStructGetData($tMagic,1)
	Local $tIMAGE_OPTIONAL_HEADER
	If $iMagic=267 Then;x86 version
		If @AutoItX64 Then Return SetError(3,0,0);uncompatible versions
		$tIMAGE_OPTIONAL_HEADER=DllStructCreate("word Magic;"& _
				"byte MajorLinkerVersion;"& _
				"byte MinorLinkerVersion;"& _
				"dword SizeOfCode;"& _
				"dword SizeOfInitializedData;"& _
				"dword SizeOfUninitializedData;"& _
				"dword AddressOfEntryPoint;"& _
				"dword BaseOfCode;"& _
				"dword BaseOfData;"& _
				"dword ImageBase;"& _
				"dword SectionAlignment;"& _
				"dword FileAlignment;"& _
				"word MajorOperatingSystemVersion;"& _
				"word MinorOperatingSystemVersion;"& _
				"word MajorImageVersion;"& _
				"word MinorImageVersion;"& _
				"word MajorSubsystemVersion;"& _
				"word MinorSubsystemVersion;"& _
				"dword Win32VersionValue;"& _
				"dword SizeOfImage;"& _
				"dword SizeOfHeaders;"& _
				"dword CheckSum;"& _
				"word Subsystem;"& _
				"word DllCharacteristics;"& _
				"dword SizeOfStackReserve;"& _
				"dword SizeOfStackCommit;"& _
				"dword SizeOfHeapReserve;"& _
				"dword SizeOfHeapCommit;"& _
				"dword LoaderFlags;"& _
				"dword NumberOfRvaAndSizes", _
				$pPointer)
		; Move pointer
		$pPointer+=96;size of $tIMAGE_OPTIONAL_HEADER
	ElseIf $iMagic=523 Then;x64 version
		If Not @AutoItX64 Then Return SetError(3,0,0);uncompatible versions
		$tIMAGE_OPTIONAL_HEADER=DllStructCreate("word Magic;"& _
				"byte MajorLinkerVersion;"& _
				"byte MinorLinkerVersion;"& _
				"dword SizeOfCode;"& _
				"dword SizeOfInitializedData;"& _
				"dword SizeOfUninitializedData;"& _
				"dword AddressOfEntryPoint;"& _
				"dword BaseOfCode;"& _
				"uint64 ImageBase;"& _
				"dword SectionAlignment;"& _
				"dword FileAlignment;"& _
				"word MajorOperatingSystemVersion;"& _
				"word MinorOperatingSystemVersion;"& _
				"word MajorImageVersion;"& _
				"word MinorImageVersion;"& _
				"word MajorSubsystemVersion;"& _
				"word MinorSubsystemVersion;"& _
				"dword Win32VersionValue;"& _
				"dword SizeOfImage;"& _
				"dword SizeOfHeaders;"& _
				"dword CheckSum;"& _
				"word Subsystem;"& _
				"word DllCharacteristics;"& _
				"uint64 SizeOfStackReserve;"& _
				"uint64 SizeOfStackCommit;"& _
				"uint64 SizeOfHeapReserve;"& _
				"uint64 SizeOfHeapCommit;"& _
				"dword LoaderFlags;"& _
				"dword NumberOfRvaAndSizes", _
				$pPointer)
		; Move pointer
		$pPointer+=112;size of $tIMAGE_OPTIONAL_HEADER
	Else
		Return SetError(3,0,0);uncompatible versions
	EndIf
	; Extract data
	Local $iSizeOfImage=DllStructGetData($tIMAGE_OPTIONAL_HEADER,"SizeOfImage");if loaded binary image would start executing at this address
	Local $iEntryPoint=DllStructGetData($tIMAGE_OPTIONAL_HEADER,"AddressOfEntryPoint");if loaded binary image would start executing at this address
	Local $pOptionalHeaderImageBase=DllStructGetData($tIMAGE_OPTIONAL_HEADER,"ImageBase");address of the first byte of the image when it's loaded in memory
	$pPointer+=8;skipping IMAGE_DIRECTORY_ENTRY_EXPORT
	; Import Directory
	Local $tIMAGE_DIRECTORY_ENTRY_IMPORT=DllStructCreate("dword VirtualAddress; dword Size",$pPointer)
	; Collect data
	Local $pAddressImport=DllStructGetData($tIMAGE_DIRECTORY_ENTRY_IMPORT,"VirtualAddress")
	Local $iSizeImport=DllStructGetData($tIMAGE_DIRECTORY_ENTRY_IMPORT,"Size")
	$pPointer+=8;size of $tIMAGE_DIRECTORY_ENTRY_IMPORT
	$pPointer+=24;skipping IMAGE_DIRECTORY_ENTRY_RESOURCE,IMAGE_DIRECTORY_ENTRY_EXCEPTION,IMAGE_DIRECTORY_ENTRY_SECURITY
	; Base Relocation Directory
	Local $tIMAGE_DIRECTORY_ENTRY_BASERELOC=DllStructCreate("dword VirtualAddress; dword Size",$pPointer)
	; Collect data
	Local $pAddressNewBaseReloc=DllStructGetData($tIMAGE_DIRECTORY_ENTRY_BASERELOC,"VirtualAddress")
	Local $iSizeBaseReloc=DllStructGetData($tIMAGE_DIRECTORY_ENTRY_BASERELOC,"Size")
	$pPointer+=8;size of IMAGE_DIRECTORY_ENTRY_BASERELOC
	$pPointer+=40;skipping IMAGE_DIRECTORY_ENTRY_DEBUG,IMAGE_DIRECTORY_ENTRY_COPYRIGHT,IMAGE_DIRECTORY_ENTRY_GLOBALPTR,IMAGE_DIRECTORY_ENTRY_TLS,IMAGE_DIRECTORY_ENTRY_LOAD_CONFIG
	$pPointer+=40;skipping five more unused data directories
	; Load the subrogor
	Local $pBaseAddress_OR=_WinAPI_LoadLibraryEx($sSubrogor,1);"lighter" loading
	Local $pBaseAddress=$pBaseAddress_OR
	If @error Then Return SetError(4,0,0);Couldn't load subrogor
	Local $bCleanLoad=_WinAPI_UnmapViewOfSection($pBaseAddress)
	If $bCleanLoad Then $pBaseAddress=_MemVirtualAlloc($pBaseAddress,$iSizeOfImage,$MEM_RESERVE+$MEM_COMMIT,$PAGE_READWRITE)
	; In case of any error make corrections
	If @error Or $pBaseAddress=0 Then
		$pBaseAddress=$pBaseAddress_OR
		$bCleanLoad=False
	EndIf
	Local $pHeadersNew=DllStructGetPtr($tIMAGE_DOS_HEADER);starting address of binary image headers
	Local $iOptionalHeaderSizeOfHeaders=DllStructGetData($tIMAGE_OPTIONAL_HEADER,"SizeOfHeaders");the size of the MS-DOS stub,the PE header,and the section headers
	; Set proper memory protection for writing
	_WinAPI_VirtualProtect($pBaseAddress,$iOptionalHeaderSizeOfHeaders,$PAGE_READWRITE)
	If @error Then
		_WinAPI_FreeLibrary($pBaseAddress)
		Return SetError(6,0,0);Couldn't set writing right for the headers
	EndIf
	; Write NEW headers
	DllStructSetData(DllStructCreate("byte["&$iOptionalHeaderSizeOfHeaders&"]",$pBaseAddress),1,DllStructGetData(DllStructCreate("byte["&$iOptionalHeaderSizeOfHeaders&"]",$pHeadersNew),1))
	; Dealing with sections. Will write them.
	Local $tIMAGE_SECTION_HEADER
	Local $iSizeOfRawData,$pPointerToRawData
	Local $iVirtualSize,$iVirtualAddress
	Local $tImpRaw,$tRelocRaw
	For $i=1 To $iNumberOfSections
		$tIMAGE_SECTION_HEADER=DllStructCreate("char Name[8];"& _
				"dword VirtualSize;"& _
				"dword VirtualAddress;"& _
				"dword SizeOfRawData;"& _
				"dword PointerToRawData;"& _
				"dword PointerToRelocations;"& _
				"dword PointerToLinenumbers;"& _
				"word NumberOfRelocations;"& _
				"word NumberOfLinenumbers;"& _
				"dword Characteristics", _
				$pPointer)
		; Collect data
		$iSizeOfRawData=DllStructGetData($tIMAGE_SECTION_HEADER,"SizeOfRawData")
		$pPointerToRawData=$pHeadersNew+DllStructGetData($tIMAGE_SECTION_HEADER,"PointerToRawData")
		$iVirtualAddress=DllStructGetData($tIMAGE_SECTION_HEADER,"VirtualAddress")
		$iVirtualSize=DllStructGetData($tIMAGE_SECTION_HEADER,"VirtualSize")
		If $iVirtualSize And $iVirtualSize < $iSizeOfRawData Then $iSizeOfRawData=$iVirtualSize
		; Set MEM_EXECUTE_READWRITE for sections (PAGE_EXECUTE_READWRITE for all for simplicity)
		_WinAPI_VirtualProtect($pBaseAddress+$iVirtualAddress,$iVirtualSize,$PAGE_EXECUTE_READWRITE)
		If @error Then;unavailable space. Will skip it and see what happens.
			$pPointer+=40;size of $tIMAGE_SECTION_HEADER structure
			ContinueLoop
		EndIf
		; Clean the space
		DllStructSetData(DllStructCreate("byte["&$iVirtualSize&"]",$pBaseAddress+$iVirtualAddress),1,DllStructGetData(DllStructCreate("byte["&$iVirtualSize&"]"),1))
		; If there is data to write,write it
		If $iSizeOfRawData Then
			DllStructSetData(DllStructCreate("byte["&$iSizeOfRawData&"]",$pBaseAddress+$iVirtualAddress),1,DllStructGetData(DllStructCreate("byte["&$iSizeOfRawData&"]",$pPointerToRawData),1))
		EndIf
		; Imports
		If $iVirtualAddress<=$pAddressImport And $iVirtualAddress+$iSizeOfRawData > $pAddressImport Then
			$tImpRaw=DllStructCreate("byte["&$iSizeImport&"]",$pPointerToRawData+($pAddressImport - $iVirtualAddress))
			__FixImports($tImpRaw,$pBaseAddress)
		EndIf
		; Relocations
		If $iVirtualAddress<=$pAddressNewBaseReloc And $iVirtualAddress+$iSizeOfRawData > $pAddressNewBaseReloc Then
			$tRelocRaw=DllStructCreate("byte["&$iSizeBaseReloc&"]",$pPointerToRawData+($pAddressNewBaseReloc - $iVirtualAddress))
		EndIf
		; Move pointer
		$pPointer+=40;size of $tIMAGE_SECTION_HEADER structure
	Next
	; Fix relocations
	If $pAddressNewBaseReloc And $iSizeBaseReloc Then __FixReloc($tRelocRaw,$pBaseAddress,$pOptionalHeaderImageBase,$iMagic=523)
	; Entry point address
	Local $pEntryFunc=$pBaseAddress+$iEntryPoint
	; DllMain simulation
	If $iEntryPoint Then DllCallAddress("bool",$pEntryFunc,"ptr",$pBaseAddress,"dword",1,"ptr",0);DLL_PROCESS_ATTACH
	; Get pseudo-handle
	Local $hPseudo=DllOpen($sSubrogor)
	If $bCleanLoad Then _WinAPI_FreeLibrary($pBaseAddress);decrement reference count
	; Remove subrogor from list of loaded modules so it can be reused if necessary
	__DeattachSubrogor($sSubrogor);call with second argument if you know what you're doing
	Return $hPseudo
EndFunc

Func __FixReloc($tData,$pAddressNew,$pAddressOld,$fImageX64)
	Local $iDelta=$pAddressNew - $pAddressOld;dislocation value
	Local $iSize=DllStructGetSize($tData);size of data
	Local $pData=DllStructGetPtr($tData);addres of the data structure
	Local $tIMAGE_BASE_RELOCATION,$iRelativeMove
	Local $iVirtualAddress,$iSizeofBlock,$iNumberOfEntries
	Local $tEnries,$iData,$tAddress
	Local $iFlag=3+7*$fImageX64;IMAGE_REL_BASED_HIGHLOW=3 or IMAGE_REL_BASED_DIR64=10
	While $iRelativeMove < $iSize;for all data available
		$tIMAGE_BASE_RELOCATION=DllStructCreate("dword VirtualAddress; dword SizeOfBlock",$pData+$iRelativeMove)
		$iVirtualAddress=DllStructGetData($tIMAGE_BASE_RELOCATION,"VirtualAddress")
		$iSizeofBlock=DllStructGetData($tIMAGE_BASE_RELOCATION,"SizeOfBlock")
		$iNumberOfEntries=($iSizeofBlock - 8)/2
		$tEnries=DllStructCreate("word["&$iNumberOfEntries&"]",DllStructGetPtr($tIMAGE_BASE_RELOCATION)+8)
		; Go through all entries
		For $i=1 To $iNumberOfEntries
			$iData=DllStructGetData($tEnries,1,$i)
			If BitShift($iData,12)=$iFlag Then;check type
				$tAddress=DllStructCreate("ptr",$pAddressNew+$iVirtualAddress+BitAND($iData,0xFFF));the rest of $iData is offset
				DllStructSetData($tAddress,1,DllStructGetData($tAddress,1)+$iDelta);this is what's this all about
			EndIf
		Next
		$iRelativeMove+=$iSizeofBlock
	WEnd
	Return 1;all OK!
EndFunc

Func __FixImports($tData,$hInstance)
	Local $pImportDirectory=DllStructGetPtr($tData)
	Local $hModule,$pFuncName,$tFuncName,$sFuncName,$pFuncAddress
	Local $tIMAGE_IMPORT_MODULE_DIRECTORY,$pModuleName,$tModuleName
	Local $tBufferOffset2,$iBufferOffset2
	Local $iInitialOffset,$iInitialOffset2,$iOffset
	Local Const $iPtrSize=DllStructGetSize(DllStructCreate("ptr"))
	While 1
		$tIMAGE_IMPORT_MODULE_DIRECTORY=DllStructCreate("dword RVAOriginalFirstThunk;"& _
				"dword TimeDateStamp;"& _
				"dword ForwarderChain;"& _
				"dword RVAModuleName;"& _
				"dword RVAFirstThunk", _
				$pImportDirectory)
		If Not DllStructGetData($tIMAGE_IMPORT_MODULE_DIRECTORY,"RVAFirstThunk") Then ExitLoop;the end
		$pModuleName=$hInstance+DllStructGetData($tIMAGE_IMPORT_MODULE_DIRECTORY,"RVAModuleName")
		$tModuleName=DllStructCreate("char Name["& _WinAPI_StringLenA($pModuleName)&"]",$hInstance+DllStructGetData($tIMAGE_IMPORT_MODULE_DIRECTORY,"RVAModuleName"))
		$hModule=_WinAPI_LoadLibrary(DllStructGetData($tModuleName,"Name"));load the module
		$iInitialOffset=$hInstance+DllStructGetData($tIMAGE_IMPORT_MODULE_DIRECTORY,"RVAFirstThunk")
		$iInitialOffset2=$hInstance+DllStructGetData($tIMAGE_IMPORT_MODULE_DIRECTORY,"RVAOriginalFirstThunk")
		If $iInitialOffset2=$hInstance Then $iInitialOffset2=$iInitialOffset
		$iOffset=0;back to 0
		While 1
			$tBufferOffset2=DllStructCreate("ptr",$iInitialOffset2+$iOffset)
			$iBufferOffset2=DllStructGetData($tBufferOffset2,1);value at that address
			If Not $iBufferOffset2 Then ExitLoop;zero value is the end
			If BitShift(BinaryMid($iBufferOffset2,$iPtrSize,1),7) Then;MSB is set for imports by ordinal,otherwise not
				$pFuncAddress=_WinAPI_GetProcAddress($hModule,BitAND($iBufferOffset2,0xFFFF));the rest is ordinal value
			Else
				$pFuncName=$hInstance+$iBufferOffset2+2
				$tFuncName=DllStructCreate("word Ordinal; char Name["& _WinAPI_StringLenA($pFuncName)&"]",$hInstance+$iBufferOffset2)
				$sFuncName=DllStructGetData($tFuncName,"Name")
				$pFuncAddress=_WinAPI_GetProcAddress($hModule,$sFuncName)
			EndIf
			DllStructSetData(DllStructCreate("ptr",$iInitialOffset+$iOffset),1,$pFuncAddress);and this is what's this all about
			$iOffset+=$iPtrSize;size of $tBufferOffset2
		WEnd
		$pImportDirectory+=20;size of $tIMAGE_IMPORT_MODULE_DIRECTORY
	WEnd
	Return 1;all OK!
EndFunc

Func _WinAPI_UnmapViewOfSection($pAddress)
	Local $aCall=DllCall("ntdll.dll","int","NtUnmapViewOfSection", _
			"handle", _WinAPI_GetCurrentProcess(), _
			"ptr",$pAddress)
	If @error Or $aCall[0] Then Return SetError(1,0,False);Failure
	Return True
EndFunc

Func _WinAPI_VirtualProtect($pAddress,$iSize,$iProtection)
	Local $aCall=DllCall("kernel32.dll","bool","VirtualProtect","ptr",$pAddress,"dword_ptr",$iSize,"dword",$iProtection,"dword*",0)
	If @error Or Not $aCall[0] Then Return SetError(1,0,0)
	Return 1
EndFunc

;~ Func _WinAPI_GetProcAddress($hModule,$vName)
;~ 	Local $sType="str"
;~ 	If IsNumber($vName) Then $sType="word"
;~ 	Local $aCall=DllCall("kernel32.dll","ptr","GetProcAddress","handle",$hModule,$sType,$vName)
;~ 	If @error Or Not $aCall[0] Then Return SetError(1,0,0)
;~ 	Return $aCall[0]
;~ EndFunc

Func __DeattachSubrogor($sSubrogor,$sNewName=Default)
	If $sNewName=Default Then $sNewName="~L"&Random(100,999,1)&".img"
	Local $tPROCESS_BASIC_INFORMATION=DllStructCreate("dword_ptr ExitStatus;"& _
			"ptr PebBaseAddress;"& _
			"dword_ptr AffinityMask;"& _
			"dword_ptr BasePriority;"& _
			"dword_ptr UniqueProcessId;"& _
			"dword_ptr InheritedFromUniqueProcessId")
	Local $aCall=DllCall("ntdll.dll","long","NtQueryInformationProcess", _
			"handle", _WinAPI_GetCurrentProcess(), _
			"dword",0, _
			"ptr",DllStructGetPtr($tPROCESS_BASIC_INFORMATION), _
			"dword",DllStructGetSize($tPROCESS_BASIC_INFORMATION), _
			"dword*",0)
	If @error Then Return SetError(1,0,0)
	Local $pPEB=DllStructGetData($tPROCESS_BASIC_INFORMATION,"PebBaseAddress")
	Local $tPEB_Small=DllStructCreate("byte InheritedAddressSpace;"& _
			"byte ReadImageFileExecOptions;"& _
			"byte BeingDebugged;"& _
			"byte Spare;"& _
			"ptr Mutant;"& _
			"ptr ImageBaseAddress;"& _
			"ptr LoaderData;", _
			$pPEB)
	Local $pPEB_LDR_DATA=DllStructGetData($tPEB_Small,"LoaderData")
	Local $tPEB_LDR_DATA=DllStructCreate("byte Reserved1[8];"& _
			"ptr Reserved2;"& _
			"ptr InLoadOrderModuleList[2];"& _
			"ptr InMemoryOrderModuleList[2];"& _
			"ptr InInitializationOrderModuleList[2];", _
			$pPEB_LDR_DATA)
	Local $pPointer=DllStructGetData($tPEB_LDR_DATA,"InMemoryOrderModuleList",2)
	Local $pEnd=$pPointer
	Local $tLDR_DATA_TABLE_ENTRY,$sModule
	While 1
		$tLDR_DATA_TABLE_ENTRY=DllStructCreate("ptr InMemoryOrderModuleList[2];"& _
				"ptr InInitializationOrderModuleList[2];"& _
				"ptr DllBase;"& _
				"ptr EntryPoint;"& _
				"ptr Reserved3;"& _
				"word LengthFullDllName;"& _
				"word LengtMaxFullDllName;"& _
				"ptr FullDllName;"& _
				"word LengthBaseDllName;"& _
				"word LengtMaxBaseDllName;"& _
				"ptr BaseDllName;", _
				$pPointer)
		$pPointer=DllStructGetData($tLDR_DATA_TABLE_ENTRY,"InMemoryOrderModuleList",2)
		If $pPointer=$pEnd Then ExitLoop
		$sModule=DllStructGetData(DllStructCreate("wchar["&DllStructGetData($tLDR_DATA_TABLE_ENTRY,"LengthBaseDllName")/2&"]",DllStructGetData($tLDR_DATA_TABLE_ENTRY,"BaseDllName")),1)
		If $sModule=$sSubrogor Then
			DllStructSetData(DllStructCreate("wchar["&DllStructGetData($tLDR_DATA_TABLE_ENTRY,"LengthBaseDllName")/2&"]",DllStructGetData($tLDR_DATA_TABLE_ENTRY,"BaseDllName")),1,$sNewName)
			Return 1
		EndIf
	WEnd
	Return SetError(2,0,0)
EndFunc
