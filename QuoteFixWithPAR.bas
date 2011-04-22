Attribute VB_Name = "QuoteFixWithPAR"
'$Id$
'
'QuoteFix with PAR - branch "no clipboard"
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
'  copyright 2011 Oliver Kopp. All rights reserved.
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
' * Removed dependency on clipboard. Currently, par does not work with certain quotes (see Tools.bas).

Option Explicit
                                                                                          
Private Const PAR_OPTIONS As String = "75q"                                             'DEFAULT=rTbgqR B=.,?_A_a Q=_s>|
Private Const PAR_CMD As String = "C:\cygwin\bin\bash.exe --login -c 'export PARINIT=""rTbgq B=.,?_A_a Q=_s>|"" ; par " & PAR_OPTIONS & "'"

'Automatically convert HTML/RTF-Mails to plain text?
Private Const CONVERT_TO_PLAIN As Boolean = False

Private Enum ReplyType
    TypeReply = 1
    TypeReplyAll = 2
    TypeForward = 3
End Enum

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
    
    'Debug.Print "READING..."
    While (pipe.StdOut.AtEndOfStream = False)
        line = pipe.StdOut.ReadLine()
        If (Left(line, 1) = ">") Then
            ret = ret & ">" & line & vbCrLf
        Else
            ret = ret & "> " & line & vbCrLf
        End If
    Wend
    ret = pipe.StdOut.ReadAll()
    'Debug.Print ret
    
    Set pipe = Nothing
    Set shell = Nothing
    
    ExecPar = ret
End Function



Private Sub FixMailText(SelectedObject As Object, MailMode As ReplyType)
    Dim TempObj As Object
    
    'we only understand mail items, no PostItems, NoteItems, ...
    If Not (TypeName(SelectedObject) = "MailItem") Then
        On Error GoTo catch:   'try, catch replacement
        Dim HadError As Boolean
        HadError = True
                          
        Select Case MailMode
            Case TypeReply:
                Set TempObj = SelectedObject.reply
                TempObj.Display
                HadError = False
                Exit Sub
            Case TypeReplyAll:
                Set TempObj = SelectedObject.ReplyAll
                TempObj.Display
                HadError = False
                Exit Sub
            Case TypeForward:
                Set TempObj = SelectedObject.Forward
                TempObj.Display
                HadError = False
                Exit Sub
        End Select
        
catch:
        On Error GoTo 0  'deactivate errorhandling
        
        If (HadError = True) Then
            'reply / replyall / forward caused error
            ' -->  just display it
            SelectedObject.Display
            Exit Sub
        End If
    End If

    Dim OriginalMail As MailItem
    Set OriginalMail = SelectedObject  'cast!!!
    
    
    'mails that have not been sent can´t be replied to (draft mails)
    If Not OriginalMail.Sent Then
        MsgBox "This mail seems to be a draft, so it cannot be replied to.", vbExclamation
        Exit Sub
    End If
    
    'we don´t understand HTML mails!!!
    If Not (OriginalMail.BodyFormat = olFormatPlain) Then
        If CONVERT_TO_PLAIN Then
            'Unfortunately, it´s only possible to convert the original mail as there is
            'no easy way to create a clone. Therefore, you cannot go back to the original format!
            'If you e.g. would decide that you need to forward the mail in HTML format,
            'this will not be possible anymore.
            SelectedObject.BodyFormat = olFormatPlain
        Else
            Dim ReplyObj As MailItem
            
            Select Case MailMode
                Case TypeReply:
                    Set ReplyObj = OriginalMail.reply
                Case TypeReplyAll:
                    Set ReplyObj = OriginalMail.ReplyAll
                Case TypeForward:
                    Set ReplyObj = OriginalMail.Forward
            End Select
            
            ReplyObj.Display
            Exit Sub
        End If
    End If
    
    'create reply --> outlook style!
    Dim NewMail As MailItem
    Select Case MailMode
        Case TypeReply:
            Set NewMail = OriginalMail.reply
        Case TypeReplyAll:
            Set NewMail = OriginalMail.ReplyAll
        Case TypeForward:
            Set NewMail = OriginalMail.Forward
    End Select
    
    'if the mail is marked as a possible phishing mail, a warning will be shown and
    'the reply methods will return null (forward method is ok)
    If NewMail Is Nothing Then Exit Sub
    
    'put the whole mail as composed by Outlook into an array
    Dim BodyLines() As String
    BodyLines = Split(NewMail.Body, vbCrLf)

    'reformat
    Dim text As String
    text = NewMail.Body
    Debug.Print "BEFORE PAR: " & vbCrLf & text
    text = ExecPar(text)
    Debug.Print "AFTER PAR: " & vbCrLf & text
    NewMail.Body = text
    
    NewMail.Display

    'mark original mail as read
    OriginalMail.UnRead = False
End Sub

'these are the macros called by the custom buttons
Sub FixedReply()
    Dim m As Object
    Set m = GetCurrentItem()

    Call FixMailText(m, TypeReply)
End Sub


Sub FixedReplyAll()
    Dim m As Object
    Set m = GetCurrentItem()

    Call FixMailText(m, TypeReplyAll)
End Sub


Sub FixedForward()
    Dim m As Object
    Set m = GetCurrentItem()

    Call FixMailText(m, TypeForward)
End Sub

