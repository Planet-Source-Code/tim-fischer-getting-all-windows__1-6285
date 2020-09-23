<div align="center">

## Getting all windows


</div>

### Description

This code explains how to use the API call EnumWindows to change captions, minimize all windows, or anything you need to do with a handle.

I explain many things you can do with the base of the code.
 
### More Info
 
All explanation is in the module with the base function. Use the following to call get windows:

Call EnumWindows(AddressOf EnumWindowProc, &H0)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tim Fischer](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tim-fischer.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tim-fischer-getting-all-windows__1-6285/archive/master.zip)

### API Declarations

```
Public Const MAX_PATH = 260
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_RESTORE = 9
Public Const WM_CLOSE = &H10
Public Declare Function SendMessage Lib "user32" _
 Alias "SendMessageA" (ByVal hwnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Integer, _
  ByVal lParam As Long) As Long
Public Declare Function ShowWindow Lib "user32" _
 (ByVal hwnd As Long, _
 ByVal nCmdShow As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" _
 (ByVal hwnd As Long) As Long
Public Declare Function EnumWindows _
 Lib "user32" (ByVal lpEnumFunc As Long, _
  ByVal lParam As Long) As Long
Public Declare Function GetClassName Lib "user32" _
 Alias "GetClassNameA" _
 (ByVal hwnd As Long, _
 ByVal lpClassName As String, _
 ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" _
 Alias "GetWindowTextA" _
 (ByVal hwnd As Long, _
 ByVal lpString As String, _
 ByVal cch As Long) As Long
Public Function EnumWindowProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
 Dim sTitle As String
 Dim sClass As String
 Dim sLoc As String
 sTitle = Space$(MAX_PATH)
 sClass = Space$(MAX_PATH)
 sLoc = Space$(MAX_PATH)
 Call GetClassName(hwnd, sClass, MAX_PATH) 'sClass
 Call GetWindowText(hwnd, sTitle, MAX_PATH) 'and sTitle
  'are given
  'their values through
  'these functions
 'You now have a handle, caption, and class name.
 'If you wanted to minimize all windows you could do
 'the following:
 'ShowWindow hwnd, SW_MINIMIZE
 'In Windows it is not this simple though.
 'There are many programs not meant to be hidden
 'like tooltips and OLE programs. You can implement
 'the IsWindowVisible function to check if it is
 'supposed to be minimized:
 'If IsWindowVisible(hwnd) Then
 'ShowWindow hwnd, SW_MINIMIZE
 'End if
 'You can change the SW_MINIMZE to any other constant
 'to make it maximize, restore to normal, or any other
 'combo. To close the windows you can do this:
 'If IsWindowVisible(hwnd) Then
 'SendMessage hwnd, WM_CLOSE, 0, 0
 'End if
 'You can also change captions and many other things
 'knowing the handle. The one thing I have come across
 'is getting the path of an application. This is not
 'possible through VB but can be achieved through DLLs.
 'I have come across a very good one created by
 'Jürgen Thümmler <thue@gmx.de>.
 EnumWindowProc = 1 'To keep EnumWindows from
  'continuing it's loop, have
  'your function return 1.
End Function
```


### Source Code

```
'You can place the function in any event.
'Call function like this:
Call EnumWindows(AddressOf EnumWindowProc, &H0)
```

