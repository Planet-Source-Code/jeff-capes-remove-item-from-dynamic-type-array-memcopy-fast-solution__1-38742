<div align="center">

## Remove item from dynamic type array \(MemCopy \- fast solution\)


</div>

### Description

Erase a specified (mIndex) item in a Dynamic Type Array. When the index is valid it shrinks the Array, so an item will not hold any 'empty' variable/data (1,2,3,4, 0 ,6,7,8 OR "a","b","c","d", "" ,"f","g")

This is the fastest possible way I know.

Please comment anything!
 
### More Info
 
Private Sub Command1_Click()

' NOTE: I didn't use ArrayOfType(0)

Dim arrayItems As Long

arrayItems = 10                       ' array size (10 items)

ReDim ArrayOfType(arrayItems) As ArrayOfType        ' set array

For i = 1 To arrayItems                   ' 10 items - i didn't use item (0)

ArrayOfType(i).item_01 = i

ArrayOfType(i).item_02 = i

ArrayOfType(i).item_03 = i

ArrayOfType(i).item_04 = i               ' fill array items with data

ArrayOfType(i).item_05 = i

ArrayOfType(i).item_06 = i

ArrayOfType(i).item_07 = i

Next i

' remove item #7 in array

If RemoveArrayItem(7) = True Then

MsgBox "item #7 in array removed..." & vbCrLf & vbCrLf & "look in you debugwindow!", vbInformation, "info"

End If

' check your debug-window for the resized array

For i = 1 To UBound(ArrayOfType)

Debug.Print "ArrayOfType(" & i & ")"

Debug.Print vbTab & "|___ item_01 = " & ArrayOfType(i).item_01

Debug.Print vbTab & "|___ item_02 = " & ArrayOfType(i).item_02

Debug.Print vbTab & "|___ item_03 = " & ArrayOfType(i).item_03

Debug.Print vbTab & "|___ item_04 = " & ArrayOfType(i).item_04

Debug.Print vbTab & "|___ item_05 = " & ArrayOfType(i).item_05

Debug.Print vbTab & "|___ item_06 = " & ArrayOfType(i).item_06

Debug.Print vbTab & "|___ item_07 = " & ArrayOfType(i).item_07

Debug.Print

Next i

Debug.Print "UBound(ArrayOfType) = " & UBound(ArrayOfType)

Debug.Print String(50, "-")

Debug.Print

End Sub

d5mn' fast


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[\~:\. Jeff 'Capes' \.:\~](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeff-capes.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeff-capes-remove-item-from-dynamic-type-array-memcopy-fast-solution__1-38742/archive/master.zip)





### Source Code

```
'---------------------------- MODULE --------------------------
Public Type ArrayOfType
 item_01     As Long
 item_02     As Long
 item_03     As Long
 item_04     As Long
 item_05     As Long
 item_06     As Long
 item_07     As Long
End Type
Public ArrayOfType()  As ArrayOfType  ' declare type as array
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Public Function RemoveArrayItem(ByVal mIndex As Long) As Boolean
' Erase a specified (mIndex) item in a Dynamic Type Array.
' When the index is valid it shrinks the Array, so an item
' will not hold any 'empty' variable (1,2,3,4, 0 ,6,7,8 OR "a","b","c","d", "" ,"f","g")
' NOTE: I don't use ArrayOfType(0)
'  if we use (as below) UBound(ArrayOfType), and the ArrayOfType() isn't
'  holding any data ( = Nothing) we get an error! :(
On Error GoTo dspErr
 Dim i   As Long    ' counter
 Dim hMatrix As Long    ' size of array
 hMatrix = UBound(ArrayOfType)
  If hMatrix = 1 Then           ' size of array is 1 (1 item hold data)
   Erase ArrayOfType          ' clear complete array (size was 1)
   RemoveArrayItem = True         ' return function
   Exit Function           ' done...
  ElseIf mIndex = hMatrix Then        ' last item in matrix?
   ReDim Preserve ArrayOfType(hMatrix - 1) As ArrayOfType ' hold data and resize array and delete last item
   RemoveArrayItem = True         ' return function
   Exit Function           ' done...
  End If
    For i = mIndex + 1 To hMatrix          ' start with item mIndex
     MemCopy ArrayOfType(i - 1), ArrayOfType(i), Len(ArrayOfType(i)) ' copy all items into the items 1 step down in the array (overwrites)
    Next i
     ReDim Preserve ArrayOfType(hMatrix - 1) As ArrayOfType   ' resize array [removes last item -> we copied it, remember?!]
     RemoveArrayItem = True           ' return function
     Exit Function             ' done...
dspErr:
 MsgBox Err.Number & " - " & Err.Description
End Function
```

