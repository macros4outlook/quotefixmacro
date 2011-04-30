Attribute VB_Name = "QuoteFixMacro"
'$Id$
'
'QuoteFix Macro TRUNK
'
'QuoteFix Macro is part of the macros4outlook project
'see http://sourceforge.net/projects/macros4outlook/ for more information
'
'For more information on Outlook see http://www.microsoft.com/outlook
'Outlook is (C) by Microsoft


'If you like this software, please write a post card to
'
'Oliver Kopp
'Schwabstr. 70a
'70193 Stuttgart
'Germany
'
'If you don't have money (or don't like the software that much, but
'appreciate the development), please send an email to
'macros4outlook-users@lists.sourceforge.net.
'
'For bug reports please go to our sourceforge bugtracker: http://sourceforge.net/projects/macros4outlook/support
'
'Thank you :-)


'****************************************************************************
'License:
'
'QuoteFix Macro
'  copyright 2006-2009 Oliver Kopp and Daniel Martin. All rights reserved.
'  copyright 2010-2011 Oliver Kopp and Lars Monsees. All rights reserved.
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
'
'Version 1.3 - 2011-04-22
' * included %C patch 2778722 by Karsten Heimrich
' * included %SE patch 2807638 by Peter Lindgren
' * check for beginning of quote is now language independent
' * added support to strip quotes of level N and greater
' * more support of alternative name formatting
'   * added support of reversed name format ("Lastname, Firstname" instead of "Firstname Lastname")
'   * added support of "LASTNAME firstname" format
'   * if no firstname is found, then the destination is used
'     * "firstname.lastname@domain" is supported
'   * firstName always starts with an uppercase letter
'   * Added support for "Dr."
' * added USE_COLORIZER and USE_SOFTWRAP conditional compiling flags.
'     They enable QuoteColorizerMacro and SoftWrapMacro
' * splitted code for parsing mailtext from FixMailText() into smaller functions
' * added support of removing the sender´s signature
' * bugfix: FinishBlock() would in some cases throw error 5
' * bugfix: Prevent error 91 when mail is marked as possible phishing mail
' * Original mail is marked as read
' * Added CONVERT_TO_PLAIN flag to enable viewing mails as HTML first.
' * renamed "fromName" to "senderName" in order to reflect real content of the variable
' * fixed cursor position in the case of absence of "%C", but presence of "%Q"
'
'$Revision$ - not released
'  * Added CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS, which condenses quoted outlook headers
'    The format of the condensed header is configured at CONDENSED_HEADER_FORMAT
'  * Added CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER
'  * Fixed compile time constants to work with Outlook 2007 and 2010
'  * Added support for custom template configured in the macro (QUOTING_TEMPLATE) - this can be used instead of the signature configuration
'  * Merged SoftWrap into QuoteFixMacro.bas

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

Option Explicit

'--------------------------------------------------------
'*** Constants for conditional compiling ***
'
'Enter these constants in the VBA project properties. The lines here only document the
'available constants.
'--------------------------------------------------------

'Should mails be colorized? (needs QuoteColorizerMacro.bas)
'(Different configuration formats for Outlook 2010 and older outlooks. Please choose the right variant)
'#Const USE_COLORIZER = True 'Outlook 2010
'USE_COLORIZER = -1 'Outlook 2003 and 2007


'--------------------------------------------------------
'*** Feature configuration ***
'--------------------------------------------------------
'Enable SoftWrap
'resize window so that the text editor wraps the text automatically
'after N charaters. Outlook wraps text automatically after sending it,
'but doesn't display the wrap when editing
'you can edit the auto wrap setting at "Tools / Options / Email Format / Internet Format"
Private Const USE_SOFTWRAP = False

'put as much characters as set in Outlook at "Tools / Options / Email Format / Internet Format"
Private Const SEVENTY_SIX_CHARS As String = "123456789x123456789x123456789x123456789x123456789x123456789x123456789x123456"

'This constant has to be adapted to fit your needs (incoprating the used font, display size, ...)
Private Const PIXEL_PER_CHARACTER As Double = 8.61842105263158


'--------------------------------------------------------
'*** Configuration constants ***
'--------------------------------------------------------
'If <> -1, strip quotes with level > INCLUDE_QUOTES_TO_LEVEL
Private Const INCLUDE_QUOTES_TO_LEVEL As Integer = -1

'At which column should the text be wrapped?
Public Const LINE_WRAP_AFTER As Integer = 75

Private Const DATE_FORMAT As String = "yyyy-mm-dd"
'alternative date format
'Private Const DATE_FORMAT As String = "ddd, d MMM yyyy at HH:mm:ss"

'Strip the sender´s signature?
Private Const STRIP_SIGNATURE As Boolean = True

'Automatically convert HTML/RTF-Mails to plain text?
Private Const CONVERT_TO_PLAIN As Boolean = False

'Enable QUOTING_TEMPLATE
Private Const USE_QUOTING_TEMPLATE As Boolean = False

'If the constant USE_QUOTING_TEMPLATE is set, this template is used instead of the signature
Private Const QUOTING_TEMPLATE As String = _
"%SN wrote on %D:" & vbCr & _
"%Q"


'--------------------------------------------------------
'*** Configuration of condensing ***
'--------------------------------------------------------

'Condense embedded quoted Outlook headers?
Private Const CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS = True

'Should the first header also be condensed?
'In case you use a custom header, (e.g., "You wrote on %D:", this should be set to false)
Private Const CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER = False

'Format of condensed header
Private Const CONDENSED_HEADER_FORMAT = "%SN wrote on %D:"

'--------------------------------------------------------


Private Const OUTLOOK_PLAIN_ORIGINALMESSAGE = "-----"
'Private Const OUTLOOK_PLAIN_ORIGINALMESSAGE = "-----Ursprüngliche Nachricht-----"
'Private Const OUTLOOK_PLAIN_ORIGINALMESSAGE = "-----Original Message-----"
Private Const OUTLOOK_ORIGINALMESSAGE   As String = "> " & OUTLOOK_PLAIN_ORIGINALMESSAGE
Private Const OUTLOOK_HEADERFINISH      As String = "> "
Private Const SIGNATURE_SEPARATOR       As String = "> --"

Private Const PATTERN_QUOTED_TEXT       As String = "%Q"
Private Const PATTERN_CURSOR_POSITION   As String = "%C"
Private Const PATTERN_SENDER_NAME       As String = "%SN"
Private Const PATTERN_SENDER_EMAIL      As String = "%SE"
Private Const PATTERN_FIRST_NAME        As String = "%FN"
Private Const PATTERN_SENT_DATE         As String = "%D"
Private Const PATTERN_OUTLOOK_HEADER    As String = "%OH"

Private Enum ReplyType
    TypeReply = 1
    TypeReplyAll = 2
    TypeForward = 3
End Enum

Public Type NestingType
    'the level of the current quote plus
    level As Integer
    
    'the amount of spaces until the next word
    'needed as outlook sometimes inserts more than one space to separate the quoteprefix and the actual quote
    'we use that information to fix the quote
    additionalSpacesCount As Integer
    
    'total = level + additionalSpacesCount + 1
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
    
    res.total = res.level + res.additionalSpacesCount + 1 '+1 = trailing space
    
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
    
    res = String(nesting.level, ">")
    res = res & String(nesting.additionalSpacesCount, " ")
    
    CalcPrefix = res & " "
End Function

'Description:
'   Adds the current line to unfomatedBlock and to curBlock
Private Sub AppendCurLine(ByRef curLine As String)
    If unformatedBlock = "" Then
        'unformatedBlock has to be used here, because it might be the case that the first
        '  line is "". Therefore curBlock remains "", whereas unformatedBlock gets <> ""
        
        If curLine = "" Then Exit Sub
        
        curBlock = curLine
        unformatedBlock = curPrefix & curLine & vbCrLf
    Else
        curBlock = curBlock & IIf(curBlock = "", "", " ") & curLine
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
    
    'Add a new line in all cases...
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
            'go through block from maxLength to beginning to find a space
            i = maxLength
            If i > 0 Then
                Do While (mid(curBlock, i, 1) <> " ")
                    i = i - 1
                    If i = 0 Then Exit Do
                Loop
            End If
    
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

'Reformat text to correct broken wrap inserted by Outlook.
'Needs to be public so the test cases can run this function.
Public Function ReFormatText(text As String) As String
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
                    'check if the next line contains a wrong break
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
            
            'Quote is indented less. Maybe it's a wrong line wrap of outlook?
            
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
                    FinishBlock lastNesting
                    
                    If curLine = "" Then
                        If curNesting.level <> lastNesting.level Then
                            lastLineWasParagraph = True
                            HandleParagraph curPrefix
                        End If
                    End If
                    
                    'next block starts with curLine
                    AppendCurLine curLine
                End If
            Else
                'Quote is the last one - just use it
                AppendCurLine curLine
            End If
        
        Else
            'curNesting.total > lastNesting.total
            
            lastLineWasParagraph = False
            
            'it's nested one level deeper. Current block is finished
            FinishBlock lastNesting
        
            If curLine = "" Then
                If curNesting.level <> lastNesting.level Then
                    lastLineWasParagraph = True
                    HandleParagraph curPrefix
                End If
            End If
            
            If CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS Then
                If Left(curLine, Len(OUTLOOK_PLAIN_ORIGINALMESSAGE)) = OUTLOOK_PLAIN_ORIGINALMESSAGE Then
                    'We found a header
                    
                    Dim posColon As Integer
                    
                    'Name and Email
                    i = i + 1
                    Dim sName As String
                    Dim sEmail As String
                    curLine = StripLine(rows(i))
                    posColon = InStr(curLine, ":")
                    Dim posLeftBracket As String
                    Dim posRightBracket As Integer
                    posLeftBracket = InStr(curLine, "[") '[ is the indication of the beginning of the E-Mail-Adress
                    posRightBracket = InStr(curLine, "]")
                    If (posLeftBracket) > 0 Then
                        sName = mid(curLine, posColon + 2, posLeftBracket - posColon - 3)
                        If posRightBracket = 0 Then
                            sEmail = mid(curLine, posLeftBracket + 8) '8 = Len("mailto: ")
                        Else
                            sEmail = mid(curLine, posLeftBracket + 8, posRightBracket - posLeftBracket - 8) '8 = Len("mailto: ")
                        End If
                    Else
                        sName = mid(curLine, posColon + 2)
                        sEmail = ""
                    End If
                    
                    i = i + 1
                    curLine = StripLine(rows(i))
                    If InStr(curLine, ":") = 0 Then
                        'There is a wrap in the email-Adress
                        posRightBracket = InStr(curLine, "]")
                        If posRightBracket > 0 Then
                            sEmail = sEmail + Left(curLine, posRightBracket - 1)
                        Else
                            'something wrent wrong, do nothing
                        End If
                        'go to next line
                        i = i + 1
                        curLine = StripLine(rows(i))
                    End If
                    
                    'Date
                    'We assume that there is always a weekday present before the date
                    Dim sDate As String
                    sDate = StripLine(rows(i))
                    'posColon = InStr(sDate, ":")
                    'sDate = mid(sDate, posColon + 2)
                    Dim posFirstComma As Integer
                    posFirstComma = InStr(sDate, ",")
                    sDate = mid(sDate, posFirstComma + 2)
                    Dim dDate As Date
                    If IsDate(sDate) Then
                        dDate = DateValue(sDate)
                        'there is no function "IsTime", therefore try with brute force
                        dDate = dDate + TimeValue(sDate)
                    End If
                    If dDate <> CDate("00:00:00") Then
                        sDate = Format(dDate, DATE_FORMAT)
                    Else
                        'leave sDate as is -> date is output as found in email
                    End If
                    
                    i = i + 3 'skip next three lines (To, [possibly CC], Subject, empty line)
                    'if CC exists, then i points to the empty line
                    'if CC does not exist, then i points to the first non-empty line
                    
                    'Strip empty lines
                    Do
                        i = i + 1
                        curLine = StripLine(rows(i))
                    Loop Until (curLine <> "") Or (i = UBound(rows))
                    i = i - 1 'i now points to the last empty line
                    
                    Dim condensedHeader As String
                    condensedHeader = CONDENSED_HEADER_FORMAT
                    condensedHeader = Replace(condensedHeader, PATTERN_SENDER_NAME, sName)
                    condensedHeader = Replace(condensedHeader, PATTERN_SENT_DATE, sDate)
                    condensedHeader = Replace(condensedHeader, PATTERN_SENDER_EMAIL, sEmail)
                    
                    Dim prefix As String
                    'the prefix for the result has to be one level shorter as it is the quoted text from the sender
                    If (curNesting.level = 1) Then
                        prefix = ""
                    Else
                        prefix = mid(curPrefix, 2)
                    End If
                    
                    result = result & prefix & condensedHeader & vbCrLf
                Else
                    'fall back to default behavior
                    'next block starts with curLine
                    AppendCurLine curLine
                End If
            Else
                'next block starts with curLine
                AppendCurLine curLine
            End If
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
    Dim TempObj As Object
    
    'we only understand mail items, no PostItems, NoteItems, ...
    If Not (TypeName(SelectedObject) = "MailItem") Then
        On Error GoTo catch:   'try, catch replacement
        Dim HadError As Boolean
        HadError = True
                          
        Select Case MailMode
            Case TypeReply:
                Set TempObj = SelectedObject.Reply
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
                    Set ReplyObj = OriginalMail.Reply
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
            Set NewMail = OriginalMail.Reply
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
    
    'lineCounter is used to provide information about how many lines we already parsed.
    'This variable is always passed to the various parser functions by reference to get
    'back the new value.
    Dim lineCounter As Long

    ' A new mail starts with signature -if- set, try to parse until we find the the
    ' original message separator - might loop until the end of the whole message since
    ' this depends on the International Option settings (english), even worse it might
    ' find some separator in-between and mess up the whole reply, so check the nesting too.
    Dim MySignature As String
    MySignature = getSignature(BodyLines, lineCounter)
    ' lineCounter now indicates the line after the signature
   
    If USE_QUOTING_TEMPLATE Then
        'Override MySignature in case the QUOTING_TEMPLATE should be used
        MySignature = QUOTING_TEMPLATE
    End If

    Dim senderName As String
    Dim firstName As String
    Call getNames(OriginalMail, senderName, firstName)
    
    If InStr(MySignature, PATTERN_SENDER_EMAIL) <> 0 Then
        Dim senderEmail As String
        senderEmail = getSenderEmailAdress(OriginalMail)
        MySignature = Replace(MySignature, PATTERN_SENDER_EMAIL, senderEmail)
    End If
    
    MySignature = Replace(MySignature, PATTERN_FIRST_NAME, firstName)
    MySignature = Replace(MySignature, PATTERN_SENT_DATE, Format(OriginalMail.SentOn, DATE_FORMAT))
    MySignature = Replace(MySignature, PATTERN_SENDER_NAME, senderName)
    
        
    Dim OutlookHeader As String
    If CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER Then
        OutlookHeader = ""
        'The real condensing is made below, where CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS is checked
        'Disabling getOutlookHeader leads to an unmodified lineCounter, which in turn gets the header included in "quotedText"
    Else
        OutlookHeader = getOutlookHeader(BodyLines, lineCounter)
    End If


    Dim quotedText As String
    quotedText = getQuotedText(BodyLines, lineCounter)
    
    
    Dim NewText As String
    'create mail according to reply mode
    Select Case MailMode
        Case TypeReply:
            NewText = quotedText
        Case TypeReplyAll:
            NewText = quotedText
        Case TypeForward:
            NewText = OutlookHeader & quotedText
    End Select
    
    'Put text in signature (=Template for text)
    MySignature = Replace(MySignature, PATTERN_OUTLOOK_HEADER & vbCrLf, OutlookHeader)
    
    'Stores number of downs to send
    Dim downCount As Long
    downCount = -1
    
    If InStr(MySignature, PATTERN_QUOTED_TEXT) <> 0 Then
        If InStr(MySignature, PATTERN_CURSOR_POSITION) = 0 Then
            'if PATTERN_CURSOR_POSITION is not set, but PATTERN_QUOTED_TEXT is, then the cursor is moved to the quote
            downCount = CalcDownCount(PATTERN_QUOTED_TEXT, MySignature)
        End If
        MySignature = Replace(MySignature, PATTERN_QUOTED_TEXT, NewText)
    Else
        'There's no placeholder. Fall back to outlook behavior
        MySignature = vbCrLf & vbCrLf & MySignature & OutlookHeader & NewText
    End If

    If (InStr(MySignature, PATTERN_CURSOR_POSITION) <> 0) Then
        downCount = CalcDownCount(PATTERN_CURSOR_POSITION, MySignature)
        'remove cursor_position pattern from mail text
        MySignature = Replace(MySignature, PATTERN_CURSOR_POSITION, "")
    End If
    
    NewMail.Body = MySignature
    
    'Extensions, in case Colorize is activated
    #If USE_COLORIZER Then
        Dim mailID As String
        mailID = QuoteColorizerMacro.ColorizeMailItem(NewMail)
        If (Trim("" & mailID) <> "") Then  'no error occured or quotefix macro not there...
            Call QuoteColorizerMacro.DisplayMailItemByID(mailID)
        Else
            'Display window
            NewMail.Display
        End If
    #Else
        'Display window
        NewMail.Display
    #End If

    'jump to the right place
    Dim i As Integer
    For i = 1 To downCount
        SendKeys "{DOWN}"
    Next i

    If USE_SOFTWRAP Then
           Call ResizeWindowForSoftWrap
    End If

    'mark original mail as read
    OriginalMail.UnRead = False
End Sub


Private Function getSignature(ByRef BodyLines() As String, ByRef lineCounter As Long) As String
    
    ' drop the first two lines, they're empty
    For lineCounter = 2 To UBound(BodyLines)
        If (InStr(BodyLines(lineCounter), OUTLOOK_ORIGINALMESSAGE) <> 0) Then
            If (CalcNesting(BodyLines(lineCounter)).level = 1) Then
                Exit For
            End If
        End If
        getSignature = getSignature & BodyLines(lineCounter) & vbCrLf
    Next lineCounter

End Function

Private Function getSenderEmailAdress(ByRef OriginalMail As MailItem) As String
    Dim senderEmail As String
    
    If OriginalMail.SenderEmailType = "SMTP" Then
        senderEmail = OriginalMail.SenderEmailAddress
    
    ElseIf OriginalMail.SenderEmailType = "EX" Then
        Dim gal As Outlook.AddressList
        Dim exchAddressEntries As Outlook.AddressEntries
        Dim exchAddressEntry As Outlook.AddressEntry
        Dim i As Integer, found As Boolean
        
        'FIXME: This seems only to work in Outlook 2007
        Set gal = OriginalMail.Session.GetGlobalAddressList
        Set exchAddressEntries = gal.AddressEntries
        
        'check if we can get the correct item by sendername
        Set exchAddressEntry = exchAddressEntries.Item(OriginalMail.senderName)
        If exchAddressEntry.name <> OriginalMail.senderName Then Set exchAddressEntry = exchAddressEntries.GetFirst

        found = False
        While (Not found) And (Not exchAddressEntry Is Nothing)
            found = (LCase(exchAddressEntry.Address) = LCase(OriginalMail.SenderEmailAddress))
            If Not found Then Set exchAddressEntry = exchAddressEntries.GetNext
        Wend
        
        If Not exchAddressEntry Is Nothing Then
            senderEmail = exchAddressEntry.GetExchangeUser.PrimarySmtpAddress
        Else
            senderEmail = ""
        End If
    End If
    
    getSenderEmailAdress = senderEmail
    
End Function

'Extracts the name of the sender from the sender's name provided in the E-Mail.
'
'In:
'  originalName - name as presented by Outlook
'Out:
'  senderName - complete name of sender
'  firstName - first name of sender
'Notes:
'  * Public to enable testing
'  * Names are returned by reference
Public Sub getNamesOutOfString(ByVal originalName, ByRef senderName As String, ByRef firstName As String)
    'Find out firstName
    
    Dim tmpName As String
    tmpName = originalName
    senderName = originalName
    
    'default: fullname
    firstName = tmpName
    
    Dim title As String
    title = ""
    'Has to be later used for extracting the last name
    
    Dim pos As Integer
    
    If (Left(tmpName, 3) = "Dr.") Then
        tmpName = mid(tmpName, 5)
        title = "Dr. "
    End If
        
    pos = InStr(tmpName, ",")
    If pos > 0 Then
        'Firstname is separated by comma and positioned behind the lastname
        firstName = Trim(mid(tmpName, pos + 1))
    Else
        pos = InStr(tmpName, " ")
        If pos > 0 Then
            'first name separated by space
            firstName = Trim(Left(tmpName, pos - 1))
            If firstName = UCase(firstName) Then
                'in case the firstName is written in uppercase letters,
                'we assume that the lastName is the real firstName
                firstName = Trim(mid(tmpName, pos + 1))
            End If
        Else
            pos = InStr(tmpName, "@")
            If pos > 0 Then
                'first name is (currenty) an eMail-Adress. Just take the prefix
                tmpName = Left(tmpName, pos - 1)
            End If
            pos = InStr(tmpName, ".")
            If pos > 0 Then
                'first name is separated by a dot
                tmpName = Left(tmpName, pos - 1)
            End If
            firstName = tmpName
        End If
    End If
    
    'fix casing of names
    firstName = UCase(Left(firstName, 1)) + mid(firstName, 2)
End Sub


'Extracts the name of the sender from the sender's name provided in the E-Mail.
'Future work is to extract the first name out of the stored Outlook contacts (if that contact exists)
'
'Notes:
'  * Names are returned by reference
Private Sub getNames(ByRef OriginalMail As MailItem, ByRef senderName As String, ByRef firstName As String)
    
    'Wildcard replaces
    senderName = OriginalMail.SentOnBehalfOfName
    
    If senderName = "" Then
        senderName = OriginalMail.senderName
    End If
    
    Call getNamesOutOfString(senderName, senderName, firstName)
End Sub


Private Function getOutlookHeader(ByRef BodyLines() As String, ByRef lineCounter As Long) As String

    ' parse until we find the header finish "> " (Outlook_Headerfinish)
    
    For lineCounter = lineCounter To UBound(BodyLines)
        If (BodyLines(lineCounter) = OUTLOOK_HEADERFINISH) Then
            Exit For
        End If
        getOutlookHeader = getOutlookHeader & BodyLines(lineCounter) & vbCrLf
    Next lineCounter
    
    'skip OUTLOOK_HEADERFINISH
    lineCounter = lineCounter + 1

End Function


Private Function getQuotedText(ByRef BodyLines() As String, ByRef lineCounter As Long) As String

    ' parse the rest of the message
    For lineCounter = lineCounter To UBound(BodyLines)
        If STRIP_SIGNATURE And (BodyLines(lineCounter) = SIGNATURE_SEPARATOR) Then
            'beginning of signature reached
            Exit For
        End If
        
        getQuotedText = getQuotedText & BodyLines(lineCounter) & vbCrLf
    Next lineCounter
    
    getQuotedText = ReFormatText(getQuotedText)

    If INCLUDE_QUOTES_TO_LEVEL <> -1 Then
        getQuotedText = StripQuotes(getQuotedText, INCLUDE_QUOTES_TO_LEVEL)
    End If

End Function


Private Function CalcDownCount(pattern As String, textToSearch As String) As Long
    Dim PosOfPattern As Long
    Dim TextBeforePattern As String
    
    PosOfPattern = InStr(textToSearch, pattern)
    TextBeforePattern = Left(textToSearch, PosOfPattern - 1)
    CalcDownCount = CountOccurencesOfStringInString(TextBeforePattern, vbCrLf)
End Function



Function GetCurrentItem() As Object  'changed to default scope
        Dim objApp As Application
        Set objApp = Session.Application
        
        Select Case TypeName(objApp.ActiveWindow)
            Case "Explorer":  'on clicking reply in the main window
                Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
            Case "Inspector": 'on clicking reply when mail is shown in separate window
                Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
        End Select
        
End Function

'Parameters:
'  InString: String to count in
'  What:     What to count
'Note:
'  * Order of parameters taken from "InStr"
Public Function CountOccurencesOfStringInString(InString As String, What As String) As Long
    Dim count As Long
    Dim lastPos As Long
    Dim curPos As Long
    
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



Private Function StripQuotes(quotedText As String, stripLevel As Integer) As String
    Dim quoteLines() As String
    Dim level As Integer
    Dim curLine As String
    Dim res As String
    Dim i As Integer
    
    quoteLines = Split(quotedText, vbCrLf)
    
    For i = 1 To UBound(quoteLines)
        level = InStr(quoteLines(i), " ") - 1
        If level <= stripLevel Then
            res = res + quoteLines(i) + vbCrLf
        End If
    Next i
    
    StripQuotes = res
End Function


'resize window so that the text editor wraps the text automatically
'after N charaters. Outlook wraps text automatically after sending it,
'but doesn't display the wrap when editing
'you can edit the auto wrap setting at "Tools / Options / Email Format / Internet Format"
Public Sub ResizeWindowForSoftWrap()
    'Application.ActiveInspector.CurrentItem.Body = SEVENTY_SIX_CHARS
    If (TypeName(Application.ActiveWindow) = "Inspector") And Not _
        (Application.ActiveInspector.WindowState = olMaximized) Then
            
        Application.ActiveInspector.Width = (LINE_WRAP_AFTER + 2) * PIXEL_PER_CHARACTER
    End If
End Sub
