<div align="center">

## Disable 'X' on Forms \(Including MDI Child Forms\)


</div>

### Description

This is shorter way to disable the 'X' or close button on a form. It also works on MDI Child forms also which I have found most other code does not.
 
### More Info
 
Just copy and paste it as stated below!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mark Troutwine](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mark-troutwine.md)
**Level**          |Unknown
**User Rating**    |4.6 (65 globes from 14 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mark-troutwine-disable-x-on-forms-including-mdi-child-forms__1-3335/archive/master.zip)

### API Declarations

```
' Place this in a module:
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Const SC_CLOSE = &HF060&
Public Const MF_BYCOMMAND = &H0&
```


### Source Code

```
' Place this in the Form Load event of the form you want to disable the 'X':
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hwnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND
```

