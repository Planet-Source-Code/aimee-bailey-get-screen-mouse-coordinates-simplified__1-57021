<div align="center">

## Get Screen Mouse Coordinates \(Simplified\)


</div>

### Description

See code name.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Aimee Bailey](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aimee-bailey.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/aimee-bailey-get-screen-mouse-coordinates-simplified__1-57021/archive/master.zip)





### Source Code

```
'Hide the following code in a module someware!
Private Declare Function GetCursorPos Lib _
  "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
  x As Long
  y As Long
End Type
'------------------------------
Public Function GetPos(Optional x As Single _
  = 0, Optional y As Single = 0)
Dim Pos As POINTAPI
Dim retVal As Boolean
 retVal = GetCursorPos(Pos)
 x = Pos.x
 y = Pos.y
End Function
' Put the following into the form of your
' choice and then create a timer called
' 'Timer1' and remember to set the interval
' to something like '10'
'-----------------------------
Private Sub Timer1_Timer()
Dim x As Single
Dim y As Single
 GetPos x, y
 Me.Caption = x & "x" & y
End Sub
```

