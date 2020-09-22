<div align="center">

## Change FileTime


</div>

### Description

This is a compact code that changes the date of an at the moment hard coded file to the time at the moment.
 
### More Info
 
I put everything (declarations etc.) in the Form so no extra modul etc. is needed.

One command button named Command1 is needed to execute the FileTime-changes. I don't like to put actions in the Form_Load-event.

It's now possible to include a select-file dialogue to easyly change the FileTimes of multiple files.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bernhard ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bernhard.md)
**Level**          |Intermediate
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bernhard-change-filetime__1-9831/archive/master.zip)





### Source Code

```
Option Explicit
Private Const GENERIC_WRITE As Long = &H40000000
Private Const GENERIC_READ As Long = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const CREATE_ALWAYS As Long = 2
Private Const OPEN_ALWAYS As Long = 4
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, _
 ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, _
 ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, _
 lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
 lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CreateFile Lib "kernel32" _
 Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
 ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, _
 ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
 ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, _
 lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, _
 lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" _
 (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" _
 (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Type FILETIME
 dwLowDateTime As Long
 dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
 wYear As Integer
 wMonth As Integer
 wDayOfWeek As Integer
 wDay As Integer
 wHour As Integer
 wMinute As Integer
 wSecond As Integer
 wMilliseconds As Integer
End Type
Private Sub Command1_Click()
Dim fHandle As Long
Dim FILE_NAME As String
FILE_NAME = "c:\test.txt" 'File with the dates to change
Dim FTime As FILETIME
fHandle = CreateFile(FILE_NAME, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
If fHandle <> INVALID_HANDLE_VALUE Then
 FTime = GetSysTimeAsFILETIME
 SetFileTime fHandle, FTime, FTime, FTime
 CloseHandle fHandle
End If
End Sub
Private Function GetSysTimeAsFILETIME() As FILETIME
Dim SysTime As SYSTEMTIME
Dim FTime As FILETIME
Dim erg As Long
GetSystemTime SysTime
erg = SystemTimeToFileTime(SysTime, FTime)
GetSysTimeAsFILETIME = FTime
End Function
```

