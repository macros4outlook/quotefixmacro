Attribute VB_Name = "QuoteFixWithPAR"
'$Id$
'
'QuoteFix with PAR TRUNK
'
'QuoteFix with PAR is part of the macros4outlook project
'see http://sourceforge.net/projects/macros4outlook/ for more information
'
'For more information on Outlook see http://www.microsoft.com/outlook
'Outlook is (C) by Microsoft

'****************************************************************************
'License:
'
'QuoteFix with PAR
'  copyright 2008-2009 Daniel Martin. All rights reserved.
'
'
'Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
'
'   1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
'   2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
'   3. The name of the author may not be used to endorse or promote products derived from this software without specific prior written permission.
'
'THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'****************************************************************************

'Changelog
'
'$Revision$ - not released

Option Explicit
                                                                                          
Private Const PAR_OPTIONS As String = "75q"                                             'DEFAULT=rTbgqR B=.,?_A_a Q=_s>|
Private Const PAR_CMD As String = "C:\cygwin\bin\bash.exe --login -c 'export PARINIT=""rTbgq B=.,?_A_a Q=_s>|"" ; par " & PAR_OPTIONS & "'"

' clipboard interaction in win32
' Provided by Allen Browne, allen@allenbrowne.com
Declare Function abOpenClipboard Lib "User32" Alias "OpenClipboard" (ByVal Hwnd As Long) As Long
Declare Function abCloseClipboard Lib "User32" Alias "CloseClipboard" () As Long
Declare Function abEmptyClipboard Lib "User32" Alias "EmptyClipboard" () As Long
Declare Function abIsClipboardFormatAvailable Lib "User32" Alias "IsClipboardFormatAvailable" (ByVal wFormat As Long) As Long
Declare Function abSetClipboardData Lib "User32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function abGetClipboardData Lib "User32" Alias "GetClipboardData" (ByVal wFormat As Long) As Long
Declare Function abGlobalAlloc Lib "Kernel32" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function abGlobalLock Lib "Kernel32" Alias "GlobalLock" (ByVal hMem As Long) As Long
Declare Function abGlobalUnlock Lib "Kernel32" Alias "GlobalUnlock" (ByVal hMem As Long) As Boolean
Declare Function abLstrcpy Lib "Kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function abGlobalFree Lib "Kernel32" Alias "GlobalFree" (ByVal hMem As Long) As Long
Declare Function abGlobalSize Lib "Kernel32" Alias "GlobalSize" (ByVal hMem As Long) As Long
Const GHND = &H42
Const CF_TEXT = 1
Const APINULL = 0



Function ExecPar(mailtext As String) As String
    Dim ret As String
    Dim line As String
        
    Dim shell As Object
    Dim pipe As Object
    Set shell = CreateObject("WScript.Shell")
    
    Debug.Print PAR_CMD
    Set pipe = shell.Exec(PAR_CMD)
    Debug.Print "END PAR"
    
    pipe.StdIn.Write (mailtext)
    pipe.StdIn.Close
    
    Debug.Print "READING..."
    While (pipe.StdOut.AtEndOfStream = False)
        line = pipe.StdOut.ReadLine()
        If (Left(line, 1) = ">") Then
            ret = ret & ">" & line & vbCrLf
        Else
            ret = ret & "> " & line & vbCrLf
        End If
    Wend
    'ret = pipe.StdOut.ReadAll()
    Debug.Print ret
    
    Set pipe = Nothing
    Set shell = Nothing
    
    ExecPar = ret
End Function


Public Sub ReformatSelectedText()
    Dim text As String
    Dim ret As Variant

    'copy selection to clipboard
    SendKeys "^c", True 'ctrl-c, wait until done
    
    'get text from clipboard
    ret = Clipboard2Text
    If (IsNull(ret)) Then Exit Sub 'error or no text in clipboard
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


Function Text2Clipboard(szText As String)
    Dim wLen As Integer
    Dim hMemory As Long
    Dim lpMemory As Long
    Dim retval As Variant
    Dim wFreeMemory As Boolean

    ' Get the length, including one extra for a CHR$(0) at the end.
    wLen = Len(szText) + 1
    szText = szText & Chr$(0)
    hMemory = abGlobalAlloc(GHND, wLen + 1)
    If hMemory = APINULL Then
        MsgBox "Unable to allocate memory."
        Exit Function
    End If
    wFreeMemory = True
    lpMemory = abGlobalLock(hMemory)
    If lpMemory = APINULL Then
        MsgBox "Unable to lock memory."
        GoTo T2CB_Free
    End If

    ' Copy our string into the locked memory.
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
    Exit Function

T2CB_Free:
    If abGlobalFree(hMemory) <> APINULL Then
        MsgBox "Unable to free global memory."
    End If
End Function



Function Clipboard2Text()
    Dim wLen As Integer
    Dim hMemory As Long
    Dim hMyMemory As Long

    Dim lpMemory As Long
    Dim lpMyMemory As Long

    Dim retval As Variant
    Dim wFreeMemory As Boolean
    Dim wClipAvail As Integer
    Dim szText As String
    Dim wSize As Long

    If abIsClipboardFormatAvailable(CF_TEXT) = APINULL Then
        Clipboard2Text = Null
        Exit Function
    End If

    If abOpenClipboard(0&) = APINULL Then
        MsgBox "Unable to open Clipboard.  Perhaps some other application is using it."
        GoTo CB2T_Free
    End If

    hMemory = abGetClipboardData(CF_TEXT)
    If hMemory = APINULL Then
        MsgBox "Unable to retrieve text from the Clipboard."
        Exit Function
    End If
    wSize = abGlobalSize(hMemory)
    szText = Space(wSize)

    wFreeMemory = True

    lpMemory = abGlobalLock(hMemory)
    If lpMemory = APINULL Then
        MsgBox "Unable to lock clipboard memory."
        GoTo CB2T_Free
    End If

    ' Copy our string into the locked memory.
    retval = abLstrcpy(szText, lpMemory)
    ' Get rid of trailing stuff.
    szText = Trim(szText)
    ' Get rid of trailing 0.
    Clipboard2Text = Left(szText, Len(szText) - 1)
    wFreeMemory = False

CB2T_Close:
    If abCloseClipboard() = APINULL Then
        MsgBox "Unable to close the Clipboard."
    End If
    If wFreeMemory Then GoTo CB2T_Free
    Exit Function

CB2T_Free:
    If abGlobalFree(hMemory) <> APINULL Then
        MsgBox "Unable to free global clipboard memory."
    End If
End Function
