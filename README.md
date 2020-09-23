<div align="center">

## Round, Ceiling, Floor Function


</div>

### Description

I found the round function from the planet source code before and I grouped it with my own ceil and floor function together. I hope these could help someone who don't want to use the round and format function to handle the numeric information.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jacky Wong](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jacky-wong.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jacky-wong-round-ceiling-floor-function__1-32172/archive/master.zip)





### Source Code

Public Function AdvRound(InValue As Double, InDecimal As Integer) As Double <br>
  Dim lDblProcess As Double <br>
 <br>
  lDblProcess = InValue * (10 ^ InDecimal) <br>
  AdvRound = Int(lDblProcess + 0.5) / (10 ^ InDecimal) <br>
End Function <br>
 <br>
Public Function AdvCeil(InValue As Double, InDecimal As Integer) As Double <br>
  Dim lDblProcess As Double <br>
  lDblProcess = InValue * (10 ^ InDecimal) <br>
  If Int(lDblProcess) < lDblProcess Then <br>
    lDblProcess = Int(lDblProcess) + 1 <br>
  Else <br>
    lDblProcess = Int(lDblProcess) <br>
  End If <br>
  AdvCeil = lDblProcess / (10 ^ InDecimal) <br>
End Function <br>
 <br>
Public Function AdvFloor(InValue As Double, InDecimal As Integer) As Double <br>
  Dim lDblProcess As Double <br>
  lDblProcess = InValue * (10 ^ InDecimal) <br>
  AdvFloor = Int(lDblProcess) / (10 ^ InDecimal) <br>
End Function <br>

