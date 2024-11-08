#include-once

;#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w- 4 -w 5 -w 6 -w- 7

; #INCLUDES# =========================================================================================================
#include <Array.au3>
; ===============================================================================================================================

; #INDEX# =======================================================================================================================
; Title .........: ArrayMultiColSort
; AutoIt Version : v3.3.8.1 or higher
; Language ......: English
; Description ...: Sorts 2D arrays on several columns
; Note ..........:
; Author(s) .....: Melba23
; Remarks .......:
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _ArrayMultiColSort : Sort 2D arrays on several columns
; ===============================================================================================================================

; #INTERNAL_USE_ONLY#=================================================================================================
; __AMCS_SortChunk : Sorts array section
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name...........: _ArrayMultiColSort
; Description ...: Sort 2D arrays on several columns
; Syntax.........: _ArrayMultiColSort(ByRef $aArray, $aSortData[, $iStart = 0[, $iEnd = 0]])
; Parameters ....: $aArray    - The 2D array to be sorted
;                  $aSortData - 2D array holding details of the sort format
;                               Format: [Column to be sorted, Sort order]
;                                   Sort order can be either numeric (0/1 = ascending/descending) or a ordered string of items
;                                   Any elements not matched in string are left unsorted after all sorted elements
;                  $iStart    - Element of array at which sort starts (default = 0)
;                  $iEnd      - Element of array at which sort endd (default = 0 - converted to end of array)
; Requirement(s).: v3.3.8.1 or higher
; Return values .: Success: No error
;                  Failure: @error set as follows
;                            @error = 1 with @extended set as follows (all refer to $sIn_Date):
;                                1 = Array to be sorted not 2D
;                                2 = Sort data array not 2D
;                                3 = More data rows in $aSortData than columns in $aArray
;                                4 = Start beyond end of array
;                                5 = Start beyond End
;                            @error = 2 with @extended set as follows:
;                                1 = Invalid string parameter in $aSortData
;                                2 = Invalid sort direction parameter in $aSortData
;                                3 = Invalid column index in $aSortData
; Author ........: Melba23
; Remarks .......: Columns can be sorted in any order
; Example .......; Yes
; ===============================================================================================================================
Func _ArrayMultiColSort(ByRef $aArray, $aSortData, $iStart = 0, $iEnd = 0)

	; Errorchecking
	; 2D array to be sorted
	If UBound($aArray, 2) = 0 Then
		Return SetError(1, 1, "")
	EndIf
	; 2D sort data
	If UBound($aSortData, 2) <> 2 Then
		Return SetError(1, 2, "")
	EndIf
	If UBound($aSortData) > UBound($aArray) Then
		Return SetError(1, 3)
	EndIf
	For $i = 0 To UBound($aSortData) - 1
		If $aSortData[$i][0] < 0 Or $aSortData[$i][0] > UBound($aArray, 2) -1 Then
			Return SetError(2, 3, "")
		EndIf
	Next
	; Start element
	If $iStart < 0 Then
		$iStart = 0
	EndIf
	If $iStart >= UBound($aArray) - 1 Then
		Return SetError(1, 4, "")
	EndIf
	; End element
	If $iEnd <= 0 Or $iEnd >= UBound($aArray) - 1 Then
		$iEnd = UBound($aArray) - 1
	EndIf
	; Sanity check
	If $iEnd <= $iStart Then
		Return SetError(1, 5, "")
	EndIf

	Local $iCurrCol, $iChunk_Start, $iMatchCol

	; Sort first column
	__AMCS_SortChunk($aArray, $aSortData, 0, $aSortData[0][0], $iStart, $iEnd)
	If @error Then
		Return SetError(2, @extended, "")
	EndIf
	; Now sort within other columns
	For $iSortData_Row = 1 To UBound($aSortData) - 1
		; Determine column to sort
		$iCurrCol = $aSortData[$iSortData_Row][0]
		; Create arrays to hold data from previous columns
		Local $aBaseValue[$iSortData_Row]
		; Set base values
		For $i = 0 To $iSortData_Row - 1
			$aBaseValue[$i] = $aArray[$iStart][$aSortData[$i][0]]
		Next
		; Set start of this chunk
		$iChunk_Start = $iStart
		; Now work down through array
		For $iRow = $iStart + 1 To $iEnd
			; Match each column
			For $k = 0 To $iSortData_Row - 1
				$iMatchCol = $aSortData[$k][0]
				; See if value in each has changed
				If $aArray[$iRow][$iMatchCol] <> $aBaseValue[$k] Then
					; If so and row has advanced
					If $iChunk_Start < $iRow - 1 Then
						; Sort this chunk
						__AMCS_SortChunk($aArray, $aSortData, $iSortData_Row, $iCurrCol, $iChunk_Start, $iRow - 1)
						If @error Then
							Return SetError(2, @extended, "")
						EndIf
					EndIf
					; Set new base value
					$aBaseValue[$k] = $aArray[$iRow][$iMatchCol]
					; Set new chunk start
					$iChunk_Start = $iRow
				EndIf
			Next
		Next
		; Sort final section
		If $iChunk_Start < $iRow - 1 Then
			__AMCS_SortChunk($aArray, $aSortData, $iSortData_Row, $iCurrCol, $iChunk_Start, $iRow - 1)
			If @error Then
				Return SetError(2, @extended, "")
			EndIf
		EndIf
	Next

EndFunc   ;==>_ArrayMultiColSort

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __AMCS_SortChunk
; Description ...: Sorts array section
; Author ........: Melba23
; Remarks .......:
; ===============================================================================================================================
Func __AMCS_SortChunk(ByRef $aArray, $aSortData, $iRow, $iColumn, $iChunkStart, $iChunkEnd)

	Local $aSortOrder

	; Set default sort direction
	Local $iSortDirn = 1
	; Need to prefix elements?
	If IsString($aSortData[$iRow][1]) Then
		; Split elements
		$aSortOrder = StringSplit($aSortData[$iRow][1], ",")
		If @error Then
			Return SetError(1, 1, "")
		EndIf
		; Add prefix to each element
		For $i = $iChunkStart To $iChunkEnd
			For $j = 1 To $aSortOrder[0]
				If $aArray[$i][$iColumn] = $aSortOrder[$j] Then
					$aArray[$i][$iColumn] = StringFormat("%02i-", $j) & $aArray[$i][$iColumn]
					ExitLoop
				EndIf
			Next
			; Deal with anything that does not match
			If $j > $aSortOrder[0] Then
				$aArray[$i][$iColumn] = StringFormat("%02i-", $j) & $aArray[$i][$iColumn]
			EndIf
		Next
	Else
		Switch $aSortData[$iRow][1]
			Case 0, 1
				; Set required sort direction if no list
				If $aSortData[$iRow][1] Then
					$iSortDirn = -1
				Else
					$iSortDirn = 1
				EndIf
			Case Else
				Return SetError(1, 2, "")
		EndSwitch
	EndIf

	; Sort the chunk
	Local $iSubMax = UBound($aArray, 2) - 1
	__ArrayQuickSort2D($aArray, $iSortDirn, $iChunkStart, $iChunkEnd, $iColumn, $iSubMax)

	; Remove any prefixes
	If IsString($aSortData[$iRow][1]) Then
		For $i = $iChunkStart To $iChunkEnd
			$aArray[$i][$iColumn] = StringTrimLeft($aArray[$i][$iColumn], 3)
		Next
	EndIf

EndFunc   ;==>__AMCS_SortChunk
