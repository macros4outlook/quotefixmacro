Attribute VB_Name = "QuoteFixMacro"
'QuoteFix Macro 1.2b
'QuoteFix Macro is part of the Outlook Theurgists - tools for ms Outlook.
'see http://www.flupp.de/OutlookTheurgists for more information
'
'For more information on Outlook see http://www.microsoft.com/outlook
'Outlook is (C) by Microsoft


'If you like this software, please write a post card to
'
'Oliver Kopp
'Schwabstrasse 70a
'70193 Stuttgart
'Germany
'
'If you don't have money (or don't like the software that much, but
'appreciate the development), please send an email to
'daniel309 [at] users [dot] sourceforge [dot] net  or theurgists [at] flupp [dot] de
'
'Thank you :-)


'****************************************************************************
'License:
'
'QuoteFix Macro 1.2b copyright 2006 Oliver Kopp and Daniel Martin. All rights reserved.
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
'Version 1.0a - 2006-09-14
' * first public release
'
'Version 1.1 - 2006-09-15
' * Macro %OH introduced
' * Outlook header contains "> " at the end
' * If no macros are in the signature, the default behavior of outlook (insert header and quoted text) text is used. (1.0a removed the header)
'
'Version 1.2 - 2006-09-25
' * QuoteFix now also fixes newly introduced first-level-quotes ("> text")
' * Header matching now matches the English header
'
'Version 1.2a - 2006-09-26
' * quick fix of bug introduced by reformating first-level-quotes
'   (it was reformated too often)
'
'Version 1.2b - 2007-01-24
' * included on-behalf-of handling written by Per Soderlind (per [at] soderlind [dot] no)

'Ideas were taken from
'  * Daniele Bochicchio
'    Button integration and sample code - http://lab.aspitalia.com/35/Outlook-2007-2003-Reply-With-Quoting-Macro.aspx
'  * Dominik Jain
'    Outlook Quotefix. An excellent program working up to Outlook 2003: http://home.in.tum.de/~jain/software/outlook-quotefix/

'Precondition:
' * The received mail has to contain the right quotes. Wrong original quotes can not always be fixed
'   > > > w1
'   > >
'   > > w2
'   > >
'   > > > w3
'   won't be fixed to w1 w2 w3. How can it be known, that w2 belongs to w1 and w3?

'Todo:
' * Implement own wrap algorithm instead of relying on the bad output of the Outlook wrap algorithm

Option Explicit

'Private Const Outlook_OriginalMessage = "> -----Urspr?ngliche Nachricht-----"
Private Const Outlook_OriginalMessage = "> -----Original Message-----"

Private Const Outlook_Headerfinish = "> "

Private Const PATTERN_QUOTED_TEXT = "%Q"
Private Const PATTERN_CURSOR_POSITION = "%C"
Private Const PATTERN_SENDER_NAME = "%SN"
Private Const PATTERN_FIRST_NAME = "%FN"
Private Const PATTERN_SENT_DATE = "%D"
Private Const PATTERN_OUTLOOK_HEADER = "%OH"

Private Const DATE_FORMAT = "yyyy-mm-dd"


'At which column should the text be wrapped?
Public Const LINE_WRAP_AFTER = 75

Private Enum ReplyType
    TypeReply = 1
    TypeReplyAll = 2
    TypeForward = 3
End Enum

Type NestingType
    level As Integer
    additionalSpacesCount As Integer
    
    'the sum + 1 - +1 because of the tailing space
    total As Integer
End Type

'Global Variables to make code more readable (-> parameter passing gets easier)
Private result As String
Private unformatedBlock As String
Private curBlock As String
Private curBlockNeedsToBeReFormated As Boolean
Private curPrefix As String
Private lastLineWasParagraph As Boolean
Private lastNesting As NestingType

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





Function CalcNesting(line As String) As NestingType 'changed to default scope
    Dim lastQuoteSignPos As Integer
    Dim i As Integer
    Dim count As Integer
    Dim curChar As String
    Dim res As NestingType
  
    count = 0
    i = 1
  
    Do While i <= Len(line)
        curChar = mid(line, i, 1)
        If curChar = ">" Then
            count = count + 1
            lastQuoteSignPos = i
        ElseIf curChar <> " " Then
            'Char is neither ">" nor " " - Quote intro ended
            'leave function
            Exit Do
        End If
        i = i + 1
    Loop
    
    res.level = count
  
    If i <= Len(line) Then
        'i contains the pos of the first character
        
        'if there is no space i = lastQuoteSignPos + 1
        'One space is normal, the others are nesting
        '  It could be, that there is no space
        
        res.additionalSpacesCount = i - lastQuoteSignPos - 2
        If res.additionalSpacesCount < 0 Then
            res.additionalSpacesCount = 0
        End If
    Else
        res.additionalSpacesCount = 0
    End If
    
    res.total = res.level + res.additionalSpacesCount + 1 '+1 = tailing space
    
    CalcNesting = res
End Function

'Description:
'   Strips away ">" and " " at the beginning to have the plain text
Private Function StripLine(line As String) As String
    Dim res As String
    res = line
    
    Do While (Len(res) > 0) And (InStr("> ", Left(res, 1)) <> 0)
        'First character is a space or a quote
        res = mid(res, 2)
    Loop
    
    'Remove the spaces at the end of res
    res = Trim(res)
    
    StripLine = res
End Function

Private Function CalcPrefix(ByRef nesting As NestingType) As String
    Dim res As String
    Dim i As Integer
    
    For i = 1 To nesting.level
        res = res & ">"
    Next i
    For i = 1 To nesting.additionalSpacesCount
        res = res & " "
    Next i
    
    CalcPrefix = res & " "
End Function

'Description:
'   Adds the current line to unfomatedBlock and to curBlock
Private Sub AppendCurLine(ByRef curLine As String)
    If unformatedBlock = "" Then
        'unformatedBlock has to be used here, because it might be the case that the first
        '  line is "". Therefore curBlock remains "", while unformatedBlock gets <> ""
        
        curBlock = curLine
        unformatedBlock = curPrefix & curLine & vbCrLf
    Else
        curBlock = curBlock & " " & curLine
        unformatedBlock = unformatedBlock & curPrefix & curLine & vbCrLf
    End If
End Sub

Private Sub HandleParagraph(ByRef prefix As String)
    If Not lastLineWasParagraph Then
        FinishBlock lastNesting
        lastLineWasParagraph = True
    Else
        'lastline was already a paragraph. No further action required
    End If
    
    'Add a <br> in all cases...
    result = result & prefix & vbCrLf
End Sub

'Description:
'   Finishes the current Block
'
'   Also resets
'       curBlockNeedsToBeReFormated
'       curBlock
'       unformatedBlock
Private Sub FinishBlock(ByRef nesting As NestingType)
    If Not curBlockNeedsToBeReFormated Then
        result = result & unformatedBlock
    Else
        'reformat curBlock and append it
        Dim prefix As String
        Dim curLine As String
        Dim maxLength As Integer
        Dim i As Integer
    
        prefix = CalcPrefix(nesting)
    
        maxLength = LINE_WRAP_AFTER - nesting.total
    
        Do While Len(curBlock) > maxLength
            i = maxLength
            Do While (i > 0) And (mid(curBlock, i, 1) <> " ")
                i = i - 1
            Loop
    
            If i = 0 Then
                'No space found -> use the full line
                curLine = Left(curBlock, maxLength)
                curBlock = mid(curBlock, maxLength + 1)
            Else
                curLine = Left(curBlock, i - 1)
                curBlock = mid(curBlock, i + 1)
            End If
    
            result = result & prefix & curLine & vbCrLf
        Loop
    
        If Len(curBlock) > 0 Then
            result = result & prefix & curBlock & vbCrLf
        End If
    End If
    
    'Resetting
    curBlockNeedsToBeReFormated = False
    curBlock = ""
    unformatedBlock = ""
    'lastLineWasParagraph = False
End Sub

Private Function ReFormatText(text As String) As String
    Dim curLine As String
    Dim rows() As String
    
    Dim lastPrefix As String
    
    Dim i As Integer
    Dim curNesting As NestingType
    Dim nextNesting As NestingType

    'Reset (partially global) variables
    result = ""
    curBlock = ""
    unformatedBlock = ""
    curNesting.level = 0
    lastNesting.level = 0
    curBlockNeedsToBeReFormated = False
    
    rows = Split(text, vbCrLf)
    
    For i = LBound(rows) To UBound(rows)
        curLine = StripLine(rows(i))
        lastNesting = curNesting
        curNesting = CalcNesting(rows(i))
        If curNesting.total <> lastNesting.total Then
            lastPrefix = curPrefix
            curPrefix = CalcPrefix(curNesting)
        End If
        If curNesting.total = lastNesting.total Then
            'Quote continues
            If curLine = "" Then
                'new paragraph has started
                HandleParagraph curPrefix
            Else
                AppendCurLine curLine
                lastLineWasParagraph = False
            
                If (curNesting.level = 1) And (i < UBound(rows)) Then
                    'possibly a wrong break is found
                    nextNesting = CalcNesting(rows(i + 1))
                    If (CountOccurencesOfStringInString(curLine, " ") = 0) And (curNesting.total = nextNesting.total) _
                        And (Len(rows(i - 1)) > LINE_WRAP_AFTER - Len(curLine) - 10) Then '10 is only a rough heuristics... - should be improved
                        'Yes, it is a wrong Wrap (same recognition as below)
                        curBlockNeedsToBeReFormated = True
                    End If
                End If
            End If
        ElseIf curNesting.total < lastNesting.total Then 'curNesting.level = lastNesting.level - 1 doesn't work, because ">>", ">>>", ... are also killed by Office
            lastLineWasParagraph = False
            
            'Quote is idented less. Maybe it 's a wrong line wrap of outlook?
            
            If (i < UBound(rows)) Then
                nextNesting = CalcNesting(rows(i + 1))
                If nextNesting.total = lastNesting.total Then
                    'Yeah. Wrong line wrap found
                    
                    If curLine = "" Then
                        'The linebreak has to be interpreted as paragraph
                        'new Paragraph has started. No joining of quotes is necessary
                        HandleParagraph lastPrefix
                    Else
                        curBlockNeedsToBeReFormated = True
                    
                        'nesting and prefix have to be adjusted
                        curNesting = lastNesting
                        curPrefix = lastPrefix
                    
                        AppendCurLine curLine
                    End If
                Else
                    'No wrong line wrap found. Last block is finished
                    FinishBlock lastNesting ', unformatedBlock, curBlock, curBlockNeedsToBeReFormated, result
                    
                    'next block starts with curLine
                    AppendCurLine curLine
                End If
            Else
                'Quote is the last one - just use it
                AppendCurLine curLine
            End If
        Else
            lastLineWasParagraph = False
            
            'it's nested one level deeper. Current block is finished
            FinishBlock lastNesting
        
            'next block starts with curLine
            'Debug.Assert(curBlock == "")
            AppendCurLine curLine
        End If
    Next i
    
    'Finish current Block
    FinishBlock curNesting
    
    'strip last (unnecessary) line feeds and spaces
    Do While ((Len(result) > 0) And (InStr(vbCr & vbLf & " ", Right(result, 1)) <> 0))
        result = Left(result, Len(result) - 1)
    Loop
    
    ReFormatText = result
End Function


Private Sub FixMailText(SelectedObject As Object, MailMode As ReplyType)

    Dim OriginalMail As MailItem
    Dim TempObj As Object
    

    'wir verstehen nur mail items, keine PostItems, NoteItems, ...
    If Not (TypeName(SelectedObject) = "MailItem") Then
        On Error GoTo catch:   'try, catch ersatz
        Dim HadError As Boolean
        HadError = True
                          
                          
        Select Case MailMode
            Case TypeReply:
                    Set TempObj = SelectedObject.reply
                    TempObj.Display
                    HadError = False
                    Exit Sub  'ende, wir koennen nix mehr machen ausser anzeigen
            Case TypeReplyAll:
                    Set TempObj = SelectedObject.ReplyAll
                    TempObj.Display
                    HadError = False
                    Exit Sub 'ende, wir koennen nix mehr machen ausser anzeigen
            Case TypeForward:
                    Set TempObj = SelectedObject.Forward
                    TempObj.Display
                    HadError = False
                    Exit Sub 'ende, wir koennen nix mehr machen ausser anzeigen
        End Select
        
        
catch:
        On Error GoTo 0  'fehlerbehandlung wieder auschalten
        
        If (HadError = True) Then
            'reply / replyall / forward hat fehler erzeugt
            ' --> einfach nur anzeigen
            SelectedObject.Display
            Exit Sub 'ENDE
        End If
    
    Else
        Set OriginalMail = SelectedObject  'cast machen!!!
    End If


    'wir verstehen keine HTML mails!!!   ...noch nicht, Olly, magst du da mal ran?
    If Not (OriginalMail.BodyFormat = olFormatPlain) Then
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
        Exit Sub   'ENDE
    End If
    
    
    'erzeuge reply --> outlook style!
    Dim NewMail As MailItem
    Select Case MailMode
        Case TypeReply: Set NewMail = OriginalMail.reply
        Case TypeReplyAll: Set NewMail = OriginalMail.ReplyAll
        Case TypeForward: Set NewMail = OriginalMail.Forward
    End Select
    
    
    Dim text As String
    text = NewMail.Body 'diese zeile erzeugt die warnmeldung !!! --> jetzt nichtmehr --> session.application
    'MsgBox text
    
    Dim lines() As String
    lines = Split(text, vbCrLf)
    
    Dim QuotedText As String
    Dim MySignature As String
    Dim OutlookHeader As String
    Dim i As Integer
    Dim curLine As String
   ' Dim NewHeader As String
   
    'This information is useless at most of the time, isn't it?
    'NewHeader = OriginalMail.SenderName & " wrote on " & Format(OriginalMail.SentOn, "yyyy-mm-dd") & ":"
    
    'Die ersten beiden Leerzeilen ?berspringen
    i = 2
    
    'Als erstes kommt die Signatur
    'In VBA werden beide Teile eines ifs gleichzeitig gepr?ft, deshalb "i < UBound(lines)"
    Do While i < UBound(lines) And ((InStr(lines(i), Outlook_OriginalMessage) = 0))
        MySignature = MySignature & lines(i) & vbCrLf
        i = i + 1
    Loop
    
    
    
    'Wildcard replaces
    Dim fromName As String
    Dim firstName As String
    Dim pos As Integer
    
    
    If OriginalMail.SentOnBehalfOfName = "" Then
      fromName = OriginalMail.SenderName
    Else
      fromName = OriginalMail.SentOnBehalfOfName
    End If
    
    firstName = fromName
    pos = InStr(firstName, " ")
    If pos = 0 Then
      'No first name could be parsed
      firstName = ""
    Else
      firstName = Left(firstName, pos - 1)
    End If
    
    MySignature = Replace(MySignature, PATTERN_FIRST_NAME, firstName)
    MySignature = Replace(MySignature, PATTERN_SENT_DATE, Format(OriginalMail.SentOn, DATE_FORMAT))
    MySignature = Replace(MySignature, PATTERN_SENDER_NAME, fromName)
    
    
    'Dann kommt der Header von Outlook. Abgeschlossen durch "> " (Outlook_Headerfinish)
    Do While (i < UBound(lines)) And (lines(i) <> Outlook_Headerfinish)
        'F?r die Freaks, die beim Forwarden auch ">" brauchen...
        OutlookHeader = OutlookHeader & lines(i) & vbCrLf
        i = i + 1
    Loop
    OutlookHeader = OutlookHeader & Outlook_Headerfinish & vbCrLf
    
    i = i + 1
    
    'Jetzt kommt der eigentliche Text:
    Do While i <= UBound(lines)
        QuotedText = QuotedText & lines(i) & vbCrLf
        i = i + 1
    Loop
    
    QuotedText = ReFormatText(QuotedText)
    
    Dim NewText As String
        
    'Mail je nach Knopf einzusetzenden Text zusammenbauen
    Select Case MailMode
        Case TypeReply:
                NewText = QuotedText
        Case TypeReplyAll:
                NewText = QuotedText
        Case TypeForward:
                NewText = OutlookHeader & QuotedText
    End Select
    
    'Calculate number of downs to sent
    Dim downCount As Integer
    downCount = -1
    
    If (InStr(MySignature, PATTERN_CURSOR_POSITION) <> 0) Then
        downCount = CalcDownCount(PATTERN_CURSOR_POSITION, MySignature)
    ElseIf InStr(MySignature, PATTERN_QUOTED_TEXT) <> 0 Then
        downCount = CalcDownCount(PATTERN_QUOTED_TEXT, MySignature)
    End If
    
    'Put text in signature (=Template for text)
    
    MySignature = Replace(MySignature, "PATTERN_OUTLOOK_HEADER" & vbCrLf, OutlookHeader)
    MySignature = Replace(MySignature, PATTERN_CURSOR_POSITION, "")
    
    If InStr(MySignature, PATTERN_QUOTED_TEXT) <> 0 Then
        NewMail.Body = Replace(MySignature, PATTERN_QUOTED_TEXT, NewText)
    Else
        'There's no placeholder. Fall back to outlook behavior
        NewMail.Body = vbCrLf & vbCrLf & MySignature & OutlookHeader & NewText
        
        'If there was "%C" used (downcount is set and therefore <> -1), adjust ist (because of the two newlines above)
        If downCount <> -1 Then
            downCount = downCount + 2
        End If
    End If
   
    'Display window
    Dim mid As String
    
    'Extensions, if Colorize and SoftWrap is activated
    'mid = QuoteColorizerMacro.ColorizeMailItem(NewMail)
    'If (Trim("" & mid) <> "") Then  'no error occured or quotefix macro not there...
    '    Call QuoteColorizerMacro.DisplayMailItemByID(mid)
    '    Call SoftWrapMacro.ResizeWindowForSoftWrap
    'Else
        NewMail.Display
    'End If
    
    'jump to the right place
    For i = 1 To downCount
        SendKeys "{DOWN}"
    Next i
End Sub


Private Function CalcDownCount(pattern As String, textToSearch As String)
    Dim PosOfPattern As Integer
    Dim TextBeforePattern As String
    
    PosOfPattern = InStr(textToSearch, pattern)
    TextBeforePattern = Left(textToSearch, PosOfPattern - 1)
    CalcDownCount = CountOccurencesOfStringInString(TextBeforePattern, vbCrLf)
End Function



Function GetCurrentItem() As Object  'changed to default scope
        Dim objApp As Application
        Set objApp = Session.Application
        
        'Dim MailObj As Object
                
        Select Case TypeName(objApp.ActiveWindow)
                Case "Explorer":  'Wenn einfach reply in der ?bersicht ged?ckt wird!
                        Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
                Case "Inspector": 'Dr?cke reply in der mail in eigenem fenster!
                        Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
        End Select
        
End Function

'Parameters:
'  InString: String to count in
'  What:     What to count
'Note:
'  * Order of parameters taken from "InStr"
Public Function CountOccurencesOfStringInString(InString As String, What As String) As Integer

    Dim count As Integer
    Dim lastPos As Integer
    Dim curPos As Integer
    
    count = 0
    lastPos = 0
    curPos = InStr(InString, What)
    Do While curPos <> 0
        lastPos = curPos + 1
        count = count + 1
        curPos = InStr(lastPos, InString, What)
    Loop
        
    CountOccurencesOfStringInString = count
End Function


