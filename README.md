<div align="center">

## Form Drag / Move


</div>

### Description

This code will make it so you dont have to use the title bar to move your forms/projects.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kevin Still](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kevin-still.md)
**Level**          |Advanced
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kevin-still-form-drag-move__1-14658/archive/master.zip)

### API Declarations

```
'Declarations
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
```


### Source Code

```
'Put This In Form_MouseDown
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, &H112, &HF012, 0
End Sub
```

