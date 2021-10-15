Attribute VB_Name = "QuoteFixWithPAR"

' SPDX-License-Identifier: BSD-3-Clause

' Tries to fix quotes using the "par" tool

' For information on QuoteFixMacro heat to: https://macros4outlook.github.io/quotefixmacro/

Option Explicit

Private Const PAR_OPTIONS As String = "75q"                                   'DEFAULT=rTbgqR B=.,?_A_a Q=_s>|
Private Const PAR_CMD As String = "C:\cygwin\bin\bash.exe --login -c 'export PARINIT=""rTbgq B=.,?_A_a Q=_s>|"" ; par " & PAR_OPTIONS & "'"

' clipboard interaction in win32
' Provided by Allen Browne, allen@allenbrowne.com
Private Declare PtrSafe Function abOpenClipboard Lib "User32" Alias "OpenClipboard" (ByVal Hwnd As Long) As Long
Private Declare PtrSafe Function abCloseClipboard Lib "User32" Alias "CloseClipboard" () As Long
Private Declare PtrSafe Function abEmptyClipboard Lib "User32" Alias "EmptyClipboard" () As Long
Private Declare PtrSafe Function abIsClipboardFormatAvailable Lib "User32" Alias "IsClipboardFormatAvailable" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function abSetClipboardData Lib "User32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare PtrSafe Function abGetClipboardData Lib "User32" Alias "GetClipboardData" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function abGlobalAlloc Lib "Kernel32" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare PtrSafe Function abGlobalLock Lib "Kernel32" Alias "GlobalLock" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function abGlobalUnlock Lib "Kernel32" Alias "GlobalUnlock" (ByVal hMem As Long) As Boolean
Private Declare PtrSafe Function abLstrcpy Lib "Kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare PtrSafe Function abGlobalFree Lib "Kernel32" Alias "GlobalFree" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function abGlobalSize Lib "Kernel32" Alias "GlobalSize" (ByVal hMem As Long) As Long

Private Const GHND As Long = &H42
Private Const CF_TEXT As Long = 1
Private Const APINULL As Long = 0


Private Function ExecPar(ByVal mailtext As String) As String
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Debug.Print PAR_CMD
    Dim pipe As Object
    Set pipe = shell.Exec(PAR_CMD)
    Debug.Print "END PAR"

    pipe.StdIn.Write (mailtext)
    pipe.StdIn.Close

    Debug.Print "READING..."
    Do While (pipe.StdOut.AtEndOfStream = False)
        Dim line As String
        line = pipe.StdOut.ReadLine()
        If (Left$(line, 1) = ">") Then
            Dim ret As String
            ret = ret & ">" & line & vbCrLf
        Else
            ret = ret & "> " & line & vbCrLf
        End If
    Loop
    'ret = pipe.StdOut.ReadAll()
    Debug.Print ret

    Set pipe = Nothing
    Set shell = Nothing

    ExecPar = ret
End Function


Public Sub ReformatSelectedText()
    'copy selection to clipboard
    SendKeys "^c", True 'ctrl-c, wait until done

    'get text from clipboard
    Dim ret As Variant
    ret = Clipboard2Text
    If (IsNull(ret)) Then Exit Sub 'error or no text in clipboard
    Dim text As String
    text = CStr(ret)
    Debug.Print "FROM CLIPBOARD: " & vbCrLf & text

    'reformat
    text = ExecPar(text)
    Debug.Print "AFTER PAR: " & vbCrLf & text

    'write back to clipboard
    Text2Clipboard (text)


    'finally, replace selected text
    SendKeys "^v", True 'ctrl-v, wait until done
End Sub


'TODO: 2: add `ByVal` or `ByRef` (default is `ByRef`). `szText` is changed
'         in the marked line and thus goes (maybe) changed to the calling
'         procedure. (Is that intended?)
Private Sub Text2Clipboard(szText As String)
    ' Get the length, including one extra for a CHR$(0) at the end.
    Dim wLen As Long
    wLen = Len(szText) + 1
    szText = szText & Chr$(0)       '<-- {2}
    Dim hMemory As Long
    hMemory = abGlobalAlloc(GHND, wLen + 1)
    If hMemory = APINULL Then
        MsgBox "Unable to allocate memory."
        Exit Sub
    End If
    Dim wFreeMemory As Boolean
    wFreeMemory = True
    Dim lpMemory As Long
    lpMemory = abGlobalLock(hMemory)
    If lpMemory = APINULL Then
        MsgBox "Unable to lock memory."
        GoTo T2CB_Free
    End If

    ' Copy our string into the locked memory.
    Dim retval As Variant
    retval = abLstrcpy(lpMemory, szText)
    ' Don't send clipboard locked memory.
    retval = abGlobalUnlock(hMemory)

    If abOpenClipboard(0&) = APINULL Then
        MsgBox "Unable to open Clipboard.  Perhaps some other application is using it."
        GoTo T2CB_Free
    End If
    If abEmptyClipboard() = APINULL Then
        MsgBox "Unable to empty the clipboard."
        GoTo T2CB_Close
    End If
    If abSetClipboardData(CF_TEXT, hMemory) = APINULL Then
        MsgBox "Unable to set the clipboard data."
        GoTo T2CB_Close
    End If
    wFreeMemory = False

T2CB_Close:
    If abCloseClipboard() = APINULL Then
        MsgBox "Unable to close the Clipboard."
    End If
    If wFreeMemory Then GoTo T2CB_Free
    Exit Sub

T2CB_Free:
    If abGlobalFree(hMemory) <> APINULL Then
        MsgBox "Unable to free global memory."
    End If
End Sub


Private Sub Clipboard2Text()
    If abIsClipboardFormatAvailable(CF_TEXT) = APINULL Then
        Clipboard2Text = Null
        Exit Sub
    End If

    If abOpenClipboard(0&) = APINULL Then
        MsgBox "Unable to open Clipboard.  Perhaps some other application is using it."
        GoTo CB2T_Free
    End If

    Dim hMemory As Long
    hMemory = abGetClipboardData(CF_TEXT)
    If hMemory = APINULL Then
        MsgBox "Unable to retrieve text from the Clipboard."
        Exit Sub
    End If
    Dim wSize As Long
    wSize = abGlobalSize(hMemory)
    Dim szText As String
    szText = Space$(wSize)

    Dim wFreeMemory As Boolean
    wFreeMemory = True

    Dim lpMemory As Long
    lpMemory = abGlobalLock(hMemory)
    If lpMemory = APINULL Then
        MsgBox "Unable to lock clipboard memory."
        GoTo CB2T_Free
    End If

    ' Copy our string into the locked memory.
    Dim retval As Variant
    retval = abLstrcpy(szText, lpMemory)
    ' Get rid of trailing stuff.
    szText = Trim$(szText)
    ' Get rid of trailing 0.
    Clipboard2Text = Left$(szText, Len(szText) - 1)
    wFreeMemory = False

CB2T_Close:
    If abCloseClipboard() = APINULL Then
        MsgBox "Unable to close the Clipboard."
    End If
    If wFreeMemory Then GoTo CB2T_Free
    Exit Sub

CB2T_Free:
    If abGlobalFree(hMemory) <> APINULL Then
        MsgBox "Unable to free global clipboard memory."
    End If
End Sub
