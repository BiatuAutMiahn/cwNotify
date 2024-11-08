
#include "ArrayMultiColSort.au3"
#include <String.au3> ; Only used to fill array

; Create and display array
Global $aArray[100][4]
For $i = 0 To 99
	$aArray[$i][0] = _StringRepeat(Chr(Random(65, 68, 1)), 5)
	$aArray[$i][1] = _StringRepeat(Chr(Random(74, 77, 1)), 5)
	$aArray[$i][2] = _StringRepeat(Chr(Random(80, 83, 1)), 5)
	$aArray[$i][3] = _StringRepeat(Chr(Random(87, 90, 1)), 5)
Next
_ArrayDisplay($aArray, "Unsorted")

; Copy arrays for separate examples below
$aArray_1 = $aArray
$aArray_2 = $aArray

; This sorts columns in ascending order - probably the most common requirement
; Sort requirement:
;  Col 0 = Decending
;  Col 1 = Ascending
;  Col 2 = Required order of elements (note not alphabetic PQRS nor reverse SRQP)
;  Col 3 = Ascending
Global $aSortData[][] = [ _
		[0, 1], _
		[1, 0], _
		[2, "SSSSS,QQQQQ,PPPPP,RRRRR"], _
		[3, 0]]

; Sort and display array
_ArrayMultiColSort($aArray_1, $aSortData)
; Display any errors encountered
If @error Then ConsoleWrite("Oops: " & @error & " - " & @extended & @CRLF)
_ArrayDisplay($aArray_1, "Sorted")

; But the UDF can sort columns in any order
; Sort requirement:
;  Col 2 = Decending
;  Col 0 = Ascending
Global $aSortData[][] = [ _
		[2, 1], _
		[0, 0]]

; Sort and display array
_ArrayMultiColSort($aArray_2, $aSortData)
; Display any errors encountered
If @error Then ConsoleWrite("Oops: " & @error & " - " & @extended & @CRLF)
_ArrayDisplay($aArray_2, "Sorted")