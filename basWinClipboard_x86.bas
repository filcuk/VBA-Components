Attribute VB_Name = "basWinClipboard_x86"
Option Explicit
'Module from http://stackoverflow.com/questions/18668928/excel-2013-64-bit-vba-clipboard-api-doesnt-work

'Found 64-bit API declarations here: http://spreadsheet1.com/uploads/3/0/6/6/3066620/win32api_ptrsafe.txt
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Sub ClipBoard_SetData(MyString As String)
'32-bit code by Microsoft: http://msdn.microsoft.com/en-us/library/office/ff192913.aspx
    Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr
    Dim hClipMemory As LongPtr, X As Long

    ' Allocate moveable global memory.
    hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)
 
    ' Lock the block to get a far pointer to this memory.
    lpGlobalMemory = GlobalLock(hGlobalMemory)

    ' Copy the string to this global memory.
    lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

    ' Unlock the memory.
    If GlobalUnlock(hGlobalMemory) <> 0 Then
       MsgBox "Could not unlock memory location. Copy aborted."
       'Debug.Print "GlobalFree returned: " & CStr(GlobalFree(hGlobalMemory))
       GoTo OutOfHere
    End If

    ' Open the Clipboard to copy data to.
    If OpenClipboard(0&) = 0 Then
       MsgBox "Could not open the Clipboard. Copy aborted."
       Exit Sub
    End If

    ' Clear the Clipboard.
    X = EmptyClipboard()

    ' Copy the data to the Clipboard.
    hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere:
    If CloseClipboard() = 0 Then
       MsgBox "Could not close Clipboard."
    End If
End Sub
