<div align="center">

## Get Current Winamp Song


</div>

### Description

This Function gets the current winamp song. I got tired of people iming me asking for this and me redoing it each time.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[\-=Pc=\-](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pc.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/pc-get-current-winamp-song__1-55280/archive/master.zip)





### Source Code

Option Explicit<br><br>
'To Use Just Do<br>
'Call MsgBox(GetSong)<br>
'And it will popup a message box with the song.
<br><br>
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long<br>
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long<br>
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long<br>
<br>
Public Function GetSong() As String<br>
Dim WinampCaption As String<br>
Dim CaptionLength As Long<br>
Dim Winamp As Long<br>
 Winamp& = FindWindow("Winamp v1.x", vbNullString)<br>
 CaptionLength& = GetWindowTextLength(Winamp&)<br>
 WinampCaption$ = String$(CaptionLength&, 0)<br>
 Call GetWindowText(Winamp&, WinampCaption$, (CaptionLength& + 1&))<br>
 <br>
 If WinampCaption <> "" Then<br>
 GetSong = WinampCaption<br>
 Else<br>
 GetSong = "Error"<br>
 End If<br>
End Function<br>

