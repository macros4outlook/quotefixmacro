Attribute VB_Name = "QuoteFixMacro"

' SPDX-License-Identifier: BSD-3-Clause

' Precondition:
'
' The received mail has to contain the "right" quotes. Wrong original quotes cannot always be fixed
'
'   > > > w1
'   > >
'   > > w2
'   > >
'   > > > w3
'
'   won't be fixed to w1 w2 w3. How can it be known, that w2 belongs to w1 and w3?

' For information on configuration head to QuoteFix Macro's homepage: https://macros4outlook.github.io/quotefixmacro/

Option Explicit


'----- DEFAULT CONFIGURATION ------------------------------------------------------------------------------------------

'The configuration is now stored in the registry
'Below, the DEFAULT values are provided (if no registry setting is found)
'
'The macro NEVER stores entries in the registry by itself
'
'You can store the default configuration in the registry by executing
'  StoreDefaultConfiguration()
'or by writing a routing executing commands similar to the following:
'   Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "CONVERT_TO_PLAIN", "true")
'Finally, or by manually creating entries in this registry hive:
'    HKEY_CURRENT_USER\Software\VB and VBA Program Settings\QuoteFixMacro
Private Const APPNAME As String = "QuoteFixMacro"
Private Const REG_GROUP_CONFIG As String = "Config"
Private Const REG_GROUP_FIRSTNAMES As String = "Firstnames" 'stores replacements for firstnames


'--------------------------------------------------------
'*** Feature QuoteColorizer ***
'--------------------------------------------------------
Private Const DEFAULT_USE_COLORIZER As Boolean = False
'If you enable it, you need MAPIRTF.DLL in C:\Windows\System32
'Does NOT work at Windows 7/64bit Outlook 2010/32bit
'
'Please enable convert RTF-to-Text at sending. Otherwise, the recipients will always receive HTML E-Mails

'How many different colors should be used for colorizing the quotes?
Private Const DEFAULT_NUM_RTF_COLORS As Integer = 4


'--------------------------------------------------------
'*** Feature SoftWrap ***
'--------------------------------------------------------
'Enable SoftWrap
'resize window so that the text editor wraps the text automatically
'after N characters. Outlook wraps text automatically after sending it,
'but doesn't display the wrap when editing
'you can edit the auto wrap setting at "Tools / Options / Email Format / Internet Format"
Private Const DEFAULT_USE_SOFTWRAP As Boolean = False

'put as much characters as set in Outlook at "Tools / Options / Email Format / Internet Format"
'default: 76 characters
Private Const DEFAULT_SEVENTY_SIX_CHARS As String = "123456789x123456789x123456789x123456789x123456789x123456789x123456789x123456"

'This constant has to be adapted to fit your needs (incoprating the used font, display size, ...)
Private Const DEFAULT_PIXEL_PER_CHARACTER As Double = 8.61842105263158


'--------------------------------------------------------
'*** Configuration constants ***
'--------------------------------------------------------
'If <> -1, strip quotes with level > INCLUDE_QUOTES_TO_LEVEL
Private Const DEFAULT_INCLUDE_QUOTES_TO_LEVEL As Integer = -1

'At which column should the text be wrapped?
Private Const DEFAULT_LINE_WRAP_AFTER As Integer = 75

Private Const DEFAULT_DATE_FORMAT As String = "yyyy-mm-dd HH:MM"
'alternative date format
'Private Const DEFAULT_DATE_FORMAT As String = "ddd, d MMM yyyy at HH:mm:ss"

'Strip the sender's signature?
Private Const DEFAULT_STRIP_SIGNATURE As Boolean = True

'Automatically convert HTML/RTF-Mails to plain text?
Private Const DEFAULT_CONVERT_TO_PLAIN As Boolean = False

'Enable QUOTING_TEMPLATE
Private Const DEFAULT_USE_QUOTING_TEMPLATE As Boolean = False

'If the constant USE_QUOTING_TEMPLATE is set, this template is used instead of the signature
Private Const DEFAULT_QUOTING_TEMPLATE As String = "(Reply inline - powered by https://macros4outlook.github.io/quotefixmacro/)\n\n%SN wrote on %D:\n\n%Q\n\nCheers"

'English quote template
Private Const DEFAULT_QUOTING_TEMPLATE_EN As String = "(Reply inline - powered by https://macros4outlook.github.io/quotefixmacro/)\n\n%SN wrote on %D:\n\n%Q\n\nCheers"

'Last names may contain formal suffixes.  If they are included here, they can be stripped.
Private Const LASTNAME_SUFFIXES As String = _
"II/III/IV/Jr/Jr./Sr/Sr./Esq."
'--------------------------------------------------------
'*** Configuration of condensing ***
'--------------------------------------------------------

'Condense embedded quoted Outlook headers?
Private Const DEFAULT_CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS As Boolean = True

'Should the first header also be condensed?
'In case you use a custom header, (e.g., "You wrote on %D:", this should be set to false)
Private Const DEFAULT_CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER As Boolean = False

'Format of condensed header
Private Const DEFAULT_CONDENSED_HEADER_FORMAT As String = "%SN wrote on %D:"

'----- END OF DEFAULT CONFIGURATION -----------------------------------------------------------------------------------


Private Const OUTLOOK_PLAIN_ORIGINALMESSAGE As String = "-----"
'Private Const OUTLOOK_PLAIN_ORIGINALMESSAGE = "-----Ursprüngliche Nachricht-----"
'Private Const OUTLOOK_PLAIN_ORIGINALMESSAGE = "-----Original Message-----"
Private Const OUTLOOK_ORIGINALMESSAGE   As String = "> " & OUTLOOK_PLAIN_ORIGINALMESSAGE
Private Const PGP_MARKER                As String = "-----BEGIN PGP"
Private Const OUTLOOK_HEADERFINISH      As String = "> "
Private Const SIGNATURE_SEPARATOR       As String = "> --"

Private Const PATTERN_QUOTED_TEXT       As String = "%Q"
Private Const PATTERN_CURSOR_POSITION   As String = "%C"
Private Const PATTERN_SENDER_NAME       As String = "%SN"
Private Const PATTERN_SENDER_EMAIL      As String = "%SE"
Private Const PATTERN_FIRST_NAME        As String = "%FN"
Private Const PATTERN_LAST_NAME         As String = "%LN"
Private Const PATTERN_SENT_DATE         As String = "%D"
Private Const PATTERN_OUTLOOK_HEADER    As String = "%OH"


'Variables storing the configuration
'They are set in LoadConfiguration()
Private USE_COLORIZER As Boolean
Private NUM_RTF_COLORS As Integer
Private USE_SOFTWRAP As Boolean
Private SEVENTY_SIX_CHARS As String
Private PIXEL_PER_CHARACTER As Double
Private INCLUDE_QUOTES_TO_LEVEL As Integer
Private LINE_WRAP_AFTER As Integer
Private DATE_FORMAT As String
Private STRIP_SIGNATURE As Boolean
Private CONVERT_TO_PLAIN As Boolean
Private USE_QUOTING_TEMPLATE As Boolean
Private QUOTING_TEMPLATE As String
Private QUOTING_TEMPLATE_EN As String
Private CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS As Boolean
Private CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER As Boolean
Private CONDENSED_HEADER_FORMAT As String

'These are fetched from the registry (LoadConfiguration), but not saved by StoreDefaultConfiguration
Private FIRSTNAME_REPLACEMENT__EMAIL() As String
Private FIRSTNAME_REPLACEMENT__FIRSTNAME() As String


'For QuoteColorizer
Public Declare PtrSafe Function WriteRTF _
        Lib "mapirtf.dll" _
        Alias "writertf" (ByVal ProfileName As String, _
                          ByVal MessageID As String, _
                          ByVal StoreID As String, _
                          ByVal cText As String) _
        As Integer

'For QuoteColorizer
Public Declare PtrSafe Function ReadRTF _
        Lib "mapirtf.dll" _
        Alias "readrtf" (ByVal ProfileName As String, _
                         ByVal SrcMsgID As String, _
                         ByVal SrcStoreID As String, _
                         ByRef MsgRTF As String) _
        As Integer


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
Private unformattedBlock As String
Private curBlock As String
Private curBlockNeedsToBeReFormatted As Boolean
Private curPrefix As String
Private lastLineWasParagraph As Boolean
Private lastNesting As NestingType

'"Fixed Reply" functionality - has to be made available as shortcut in Outlook
Sub FixedReply()
    Dim m As Object
    Set m = GetCurrentItem()

    Call FixMailText(m, TypeReply)
End Sub

'"Fixed Reply All" functionality - has to be made available as shortcut in Outlook
Sub FixedReplyAll()
    Dim m As Object
    Set m = GetCurrentItem()

    Call FixMailText(m, TypeReplyAll)
End Sub

'"Fixed Reply All" functionality with English template
Sub FixedReplyAllEnglish()
    Dim m As Object
    Set m = GetCurrentItem()

    Call FixMailText(m, TypeReplyAll, True)
End Sub

'"Fixed Forward" functionality - has to be made available as shortcut in Outlook
Sub FixedForward()
    Dim m As Object
    Set m = GetCurrentItem()

    Call FixMailText(m, TypeForward)
End Sub

Function CalcNesting(line As String) As NestingType
    Dim lastQuoteSignPos As Integer
    Dim i As Integer
    Dim count As Integer
    Dim curChar As String
    Dim res As NestingType

    count = 0
    i = 1

    Do While i <= Len(line)
        curChar = Mid(line, i, 1)
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

'Stores the default values in the system registry
Public Sub StoreDefaultConfiguration()
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "USE_COLORIZER", DEFAULT_USE_COLORIZER)
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "NUM_RTF_COLORS", DEFAULT_NUM_RTF_COLORS)
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "USE_SOFTWRAP", DEFAULT_USE_SOFTWRAP)
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "SEVENTY_SIX_CHARS", DEFAULT_SEVENTY_SIX_CHARS)
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "PIXEL_PER_CHARACTER", DEFAULT_PIXEL_PER_CHARACTER)
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "INCLUDE_QUOTES_TO_LEVEL", DEFAULT_INCLUDE_QUOTES_TO_LEVEL)
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "LINE_WRAP_AFTER", DEFAULT_LINE_WRAP_AFTER)
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "DATE_FORMAT", DEFAULT_DATE_FORMAT)
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "STRIP_SIGNATURE", DEFAULT_STRIP_SIGNATURE)
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "CONVERT_TO_PLAIN", DEFAULT_CONVERT_TO_PLAIN)
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "USE_QUOTING_TEMPLATE", DEFAULT_USE_QUOTING_TEMPLATE)
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "QUOTING_TEMPLATE", DEFAULT_QUOTING_TEMPLATE)
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "QUOTING_TEMPLATE_EN", DEFAULT_QUOTING_TEMPLATE_EN)
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS", DEFAULT_CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS)
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER", DEFAULT_CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER)
    Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "CONDENSED_HEADER_FORMAT", DEFAULT_CONDENSED_HEADER_FORMAT)
End Sub

'Loads the personal settings from the registry.
Public Sub LoadConfiguration()
    USE_COLORIZER = CBool(GetSetting(APPNAME, REG_GROUP_CONFIG, "USE_COLORIZER", DEFAULT_USE_COLORIZER))
    NUM_RTF_COLORS = Val(GetSetting(APPNAME, REG_GROUP_CONFIG, "NUM_RTF_COLORS", DEFAULT_NUM_RTF_COLORS))
    USE_SOFTWRAP = CBool(GetSetting(APPNAME, REG_GROUP_CONFIG, "USE_SOFTWRAP", DEFAULT_USE_SOFTWRAP))
    SEVENTY_SIX_CHARS = GetSetting(APPNAME, REG_GROUP_CONFIG, "SEVENTY_SIX_CHARS", DEFAULT_SEVENTY_SIX_CHARS)
    PIXEL_PER_CHARACTER = CDbl(GetSetting(APPNAME, REG_GROUP_CONFIG, "PIXEL_PER_CHARACTER", DEFAULT_PIXEL_PER_CHARACTER))
    INCLUDE_QUOTES_TO_LEVEL = Val(GetSetting(APPNAME, REG_GROUP_CONFIG, "INCLUDE_QUOTES_TO_LEVEL", DEFAULT_INCLUDE_QUOTES_TO_LEVEL))
    LINE_WRAP_AFTER = Val(GetSetting(APPNAME, REG_GROUP_CONFIG, "LINE_WRAP_AFTER", DEFAULT_LINE_WRAP_AFTER))
    DATE_FORMAT = GetSetting(APPNAME, REG_GROUP_CONFIG, "DATE_FORMAT", DEFAULT_DATE_FORMAT)
    STRIP_SIGNATURE = CBool(GetSetting(APPNAME, REG_GROUP_CONFIG, "STRIP_SIGNATURE", DEFAULT_STRIP_SIGNATURE))
    CONVERT_TO_PLAIN = CBool(GetSetting(APPNAME, REG_GROUP_CONFIG, "CONVERT_TO_PLAIN", DEFAULT_CONVERT_TO_PLAIN))
    USE_QUOTING_TEMPLATE = CBool(GetSetting(APPNAME, REG_GROUP_CONFIG, "USE_QUOTING_TEMPLATE", DEFAULT_USE_QUOTING_TEMPLATE))
    CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS = CBool(GetSetting(APPNAME, REG_GROUP_CONFIG, "CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS", DEFAULT_CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS))
    CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER = CBool(GetSetting(APPNAME, REG_GROUP_CONFIG, "CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER", DEFAULT_CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER))
    CONDENSED_HEADER_FORMAT = GetSetting(APPNAME, REG_GROUP_CONFIG, "CONDENSED_HEADER_FORMAT", DEFAULT_CONDENSED_HEADER_FORMAT)

    QUOTING_TEMPLATE = GetSetting(APPNAME, REG_GROUP_CONFIG, "QUOTING_TEMPLATE", DEFAULT_QUOTING_TEMPLATE)
    QUOTING_TEMPLATE = Replace(QUOTING_TEMPLATE, "\n", vbCrLf)

    QUOTING_TEMPLATE_EN = GetSetting(APPNAME, REG_GROUP_CONFIG, "QUOTING_TEMPLATE_EN", DEFAULT_QUOTING_TEMPLATE_EN)
    QUOTING_TEMPLATE_EN = Replace(QUOTING_TEMPLATE_EN, "\n", vbCrLf)

    Dim count As Variant
    count = CDbl(GetSetting(APPNAME, REG_GROUP_FIRSTNAMES, "Count", 0))
    ReDim FIRSTNAME_REPLACEMENT__EMAIL(count)
    ReDim FIRSTNAME_REPLACEMENT__FIRSTNAME(count)
    Dim i As Integer
    For i = 1 To count
        Dim group As String
        group = REG_GROUP_FIRSTNAMES & "\" & i
        FIRSTNAME_REPLACEMENT__EMAIL(i) = GetSetting(APPNAME, group, "email", vbNullString)
        FIRSTNAME_REPLACEMENT__FIRSTNAME(i) = GetSetting(APPNAME, group, "firstName", vbNullString)
    Next i
End Sub

'Description:
'   Strips away ">" and " " at the beginning to have the plain text
Private Function StripLine(line As String) As String
    Dim res As String
    res = line

    Do While (Len(res) > 0) And (InStr("> ", Left(res, 1)) <> 0)
        'First character is a space or a quote
        res = Mid(res, 2)
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
'   Adds the current line to unformattedBlock and to curBlock
Private Sub AppendCurLine(ByRef curLine As String)
    If unformattedBlock = "" Then
        'unformattedBlock has to be used here, because it might be the case that the first
        '  line is "". Therefore curBlock remains "", whereas unformattedBlock gets <> ""

        If curLine = "" Then Exit Sub

        curBlock = curLine
        unformattedBlock = curPrefix & curLine & vbCrLf
    Else
        curBlock = curBlock & IIf(curBlock = "", "", " ") & curLine
        unformattedBlock = unformattedBlock & curPrefix & curLine & vbCrLf
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
'       curBlockNeedsToBeReFormatted
'       curBlock
'       unformattedBlock
Private Sub FinishBlock(ByRef nesting As NestingType)
    If Not curBlockNeedsToBeReFormatted Then
        result = result & unformattedBlock
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
                Do While (Mid(curBlock, i, 1) <> " ")
                    i = i - 1
                    If i = 0 Then Exit Do
                Loop
            End If

            If i = 0 Then
                'No space found -> use the full line
                curLine = Left(curBlock, maxLength)
                curBlock = Mid(curBlock, maxLength + 1)
            Else
                curLine = Left(curBlock, i - 1)
                curBlock = Mid(curBlock, i + 1)
            End If

            result = result & prefix & curLine & vbCrLf
        Loop

        If Len(curBlock) > 0 Then
            result = result & prefix & curBlock & vbCrLf
        End If
    End If

    'Resetting
    curBlockNeedsToBeReFormatted = False
    curBlock = ""
    unformattedBlock = ""
    'lastLineWasParagraph = False
End Sub

'Reformat text to correct broken wrap inserted by Outlook.
'Needs to be public so the test cases can run this function.
Public Function ReFormatText(text As String) As String
    Dim curLine As String
    Dim rows() As String
    Dim lastPrefix As String
    Dim i As Long
    Dim curNesting As NestingType
    Dim nextNesting As NestingType

    'Reset (partially global) variables
    result = ""
    curBlock = ""
    unformattedBlock = ""
    curNesting.level = 0
    lastNesting.level = 0
    curBlockNeedsToBeReFormatted = False

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
                        curBlockNeedsToBeReFormatted = True
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
                        curBlockNeedsToBeReFormatted = True

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
                If Left(curLine, Len(OUTLOOK_PLAIN_ORIGINALMESSAGE)) = OUTLOOK_PLAIN_ORIGINALMESSAGE _
                And Not Left(curLine, Len(PGP_MARKER)) = PGP_MARKER _
                Then
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
                        Dim lengthName As Integer
                        lengthName = posLeftBracket - posColon - 3
                        If lengthName > 0 Then
                            sName = Mid(curLine, posColon + 2, lengthName)
                        Else
                            Debug.Print "Could not get name. Is the header formatted correctly?"
                        End If

                        If posRightBracket = 0 Then
                            sEmail = Mid(curLine, posLeftBracket + 8) '8 = Len("mailto: ")
                        Else
                            sEmail = Mid(curLine, posLeftBracket + 8, posRightBracket - posLeftBracket - 8) '8 = Len("mailto: ")
                        End If
                    Else
                        sName = Mid(curLine, posColon + 2)
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
                    sDate = Mid(sDate, posFirstComma + 2)
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
                        prefix = Mid(curPrefix, 2)
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
    Do While ((Len(result) > 0) And (InStr(vbCrLf & " ", Right(result, 1)) <> 0))
        result = Left(result, Len(result) - 1)
    Loop

    ReFormatText = result
End Function

' @param UseEnglishTemplate In case USE_QUOTING_TEMPLATE is True, should the default or the English template be used?
Private Sub FixMailText(SelectedObject As Object, MailMode As ReplyType, Optional ByVal UseEnglishTemplate As Boolean = False)
    Dim TempObj As Object

    Call LoadConfiguration

    'we only understand mail items and meeting items , no PostItems, NoteItems, ...
    If Not (TypeName(SelectedObject) = "MailItem") And _
    Not (TypeName(SelectedObject) = "MeetingItem") Then
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

    Dim isMail As Boolean
    isMail = (TypeName(SelectedObject) = "MailItem")

    Dim OriginalMail As MailItem
    Dim OriginalMeeting As MeetingItem
    If isMail Then
        Set OriginalMail = SelectedObject 'cast!
    Else
       Set OriginalMeeting = SelectedObject 'cast!
    End If

    Dim sent As Boolean
    If isMail Then
        sent = OriginalMail.sent
    Else
        sent = OriginalMeeting.sent
    End If

    'mails that have not been sent cannot be replied to (draft mails)
    If Not sent Then
        MsgBox "This mail seems to be a draft, so it cannot be replied to.", vbExclamation
        Exit Sub
    End If

    Dim bodyFormat As olBodyFormat
    If isMail Then
        bodyFormat = OriginalMail.bodyFormat
    Else
        bodyFormat = OriginalMeeting.bodyFormat
    End If

    'basically, we do not understand HTML mails
    If Not (bodyFormat = olFormatPlain) Then
        If CONVERT_TO_PLAIN Then
            'Unfortunately, it is only possible to convert the original mail as there is
            'no easy way to create a clone. Therefore, you cannot go back to the original format!
            'If you e.g. would decide that you need to forward the mail in HTML format,
            'this will not be possible anymore.
            SelectedObject.bodyFormat = olFormatPlain
        Else
            Dim ReplyObj As MailItem

            Select Case MailMode
                Case TypeReply:
                    If isMail Then
                        Set ReplyObj = OriginalMail.Reply
                    Else
                        Set ReplyObj = OriginalMeeting.Reply
                    End If
                Case TypeReplyAll:
                    If isMail Then
                        Set ReplyObj = OriginalMail.ReplyAll
                    Else
                        Set ReplyObj = OriginalMeeting.ReplyAll
                    End If
                Case TypeForward:
                    If isMail Then
                        Set ReplyObj = OriginalMail.Forward
                    Else
                        Set ReplyObj = OriginalMeeting.Forward
                    End If
            End Select

            ReplyObj.Display
            Exit Sub
        End If
    End If

    'create reply --> outlook style!
    Dim NewMail As MailItem
    Select Case MailMode
        Case TypeReply:
            If isMail Then
                Set NewMail = OriginalMail.Reply
            Else
                Set NewMail = OriginalMeeting.Reply
            End If
        Case TypeReplyAll:
            If isMail Then
                Set NewMail = OriginalMail.ReplyAll
            Else
                Set NewMail = OriginalMeeting.ReplyAll
            End If
        Case TypeForward:
            If isMail Then
                Set NewMail = OriginalMail.Forward
            Else
                Set NewMail = OriginalMeeting.Forward
            End If
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
    '
    ' We need to call getSignature in all cases as it sets "lineCounter" as side effect
    Dim MySignature As String
    MySignature = getSignature(BodyLines, lineCounter)
    ' lineCounter now indicates the line after the signature

    If USE_QUOTING_TEMPLATE Then
        'Override MySignature in case the QUOTING_TEMPLATE should be used
        'lineCounter is still valid, because lineCounter is based on the current message whereas QUOTING_TEMPLATE is a general setting
        If UseEnglishTemplate Then
            MySignature = QUOTING_TEMPLATE_EN
        Else
            MySignature = QUOTING_TEMPLATE
        End If
    End If

    Dim senderName As String
    Dim firstName As String
    Dim lastName As String
    If isMail Then
        Call getNamesFromMail(OriginalMail, senderName, firstName, lastName)
    Else
        Call getNamesFromMeeting(OriginalMeeting, senderName, firstName, lastName)
    End If

    If (UBound(FIRSTNAME_REPLACEMENT__EMAIL) > 0) Or (InStr(MySignature, PATTERN_SENDER_EMAIL) <> 0) Then
        Dim senderEmail As String
        If isMail Then
            senderEmail = getSenderEmailAdress(OriginalMail.senderEmailType, senderName, OriginalMail.senderEmailAddress, OriginalMail.session)
        Else
            senderEmail = getSenderEmailAdress(OriginalMeeting.senderEmailType, senderName, OriginalMeeting.senderEmailAddress, OriginalMeeting.session)
        End If
        MySignature = Replace(MySignature, PATTERN_SENDER_EMAIL, senderEmail)
    End If

    If (UBound(FIRSTNAME_REPLACEMENT__EMAIL) > 0) Then
        'replace firstName by email stored in registry
        Dim rEmail As Variant
        Dim curIndex As Integer
        For curIndex = 1 To UBound(FIRSTNAME_REPLACEMENT__EMAIL)
            rEmail = FIRSTNAME_REPLACEMENT__EMAIL(curIndex)
            If (StrComp(LCase(senderEmail), LCase(rEmail)) = 0) Then
                firstName = FIRSTNAME_REPLACEMENT__FIRSTNAME(curIndex)
                Exit For
            End If
        Next curIndex
    End If

    MySignature = Replace(MySignature, PATTERN_FIRST_NAME, firstName)
    MySignature = Replace(MySignature, PATTERN_LAST_NAME, lastName)
    If isMail Then
        MySignature = Replace(MySignature, PATTERN_SENT_DATE, Format(OriginalMail.SentOn, DATE_FORMAT))
    Else
        MySignature = Replace(MySignature, PATTERN_SENT_DATE, Format(OriginalMeeting.SentOn, DATE_FORMAT))
    End If
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

    MySignature = cleanUpDoubleLines(MySignature)

    NewMail.Body = MySignature

    'Extensions, in case Colorize is activated
    If USE_COLORIZER Then
        Dim mailID As String
        mailID = ColorizeMailItem(NewMail)
        If (Trim("" & mailID) <> "") Then  'no error occured or quotefix macro not there...
            Call DisplayMailItemByID(mailID)
        Else
            'Display window
            NewMail.Display
        End If
    Else
        'Display window
        NewMail.Display
    End If

    'jump to the right place
    Dim i As Integer
    For i = 1 To downCount
        SendKeys "{DOWN}"
    Next i

    If USE_SOFTWRAP Then
           Call ResizeWindowForSoftWrap
    End If

    'mark original mail as read
    If isMail Then
        OriginalMail.UnRead = False
    Else
        OriginalMeeting.UnRead = False
    End If
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

Private Function getSenderEmailAdress(senderEmailType As String, senderName As String, senderEmailAddress As String, session As NameSpace) As String
    Dim senderEmail As String

    If senderEmailType = "SMTP" Then
        senderEmail = senderEmailAddress

    ElseIf senderEmailType = "EX" Then
        Dim gal As Outlook.AddressList
        Dim exchAddressEntries As Outlook.AddressEntries
        Dim exchAddressEntry As Outlook.AddressEntry
        Dim i As Integer, found As Boolean

        'FIXME: This seems only to work in Outlook 2007
        Set gal = session.GetGlobalAddressList
        Set exchAddressEntries = gal.AddressEntries

        'check if we can get the correct item by sendername
        Set exchAddressEntry = exchAddressEntries.item(senderName)
        If exchAddressEntry.name <> senderName Then Set exchAddressEntry = exchAddressEntries.GetFirst

        found = False
        While (Not found) And (Not exchAddressEntry Is Nothing)
            found = (LCase(exchAddressEntry.Address) = LCase(senderEmailAddress))
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

Private Function IsUpperCaseChar(ByVal c As String) As Boolean
    IsUpperCaseChar = c Like "[A-Z]"
End Function

Public Function IsUpperCaseWord(ByVal word As String) As Boolean
    IsUpperCaseWord = word Like "[A-Z][A-Z]*"
End Function

Private Function IsWordCased(ByVal word As String) As Boolean
    IsWordCased = (word Like "[A-Z][a-z]*") Or (word Like "[A-Z][a-z]*-[A-Z][a-z]*")
End Function

Private Function FixCase(ByRef word As String) As String
    If word = "" Then
        FixCase = word
        Exit Function
    End If
    Dim parts() As String
    parts = Split(word, "-")
    Dim result As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
       result = result + UCase(Left(parts(i), 1)) + LCase(Mid(parts(i), 2)) + "-"
    Next i
    FixCase = Left(result, Len(result) - 1)
End Function


'Attempts to extract the name of the sender from the sender's name provided in the E-Mail.
'
'Comment:
'   This is very difficult to do definitively.  Consider how many different variations
'   and arrangements a name like "Dr. John James Walker Smith III" could entail including
'   questions of first vs. last names, initials, titles, suffixes, etc..
'
'In:
'  originalName - name as presented by Outlook
'Out:
'  senderName - complete name of sender
'  firstName - first name of sender
'  lastName - last name of sender
'  senderEmailAddress - sender email address (optional because of tests)
'Notes:
'  * Public to enable testing
'  * Names are returned by reference
Public Sub getNamesOutOfString(ByVal originalName, ByRef senderName As String, ByRef firstName As String, ByRef lastName As String, Optional ByRef senderEmailAddress As String = "")
    'Find out firstName

    Dim tmpName As String
    tmpName = originalName

    'cleanup quotes: if name is encloded in quotes, just remove them
    If (Left(tmpName, 1) = """" And Right(tmpName, 1) = """") Then
        tmpName = Mid(tmpName, 2, Len(tmpName) - 2)
    End If

    'default full senderName: originalName without quotes
    senderName = tmpName

    'default firstName: fullname
    firstName = tmpName

    Dim title As String
    title = ""
    'Has to be later used for extracting the last name

    Dim fpos As Integer
    Dim lpos As Integer

    tmpName = removeDepartment(tmpName)

    If (Left(tmpName, 3) = "Dr.") Then
        tmpName = Mid(tmpName, 5)
        title = "Dr. "
    End If

    'Some companies have "(Text)" at the end of their name.
    'We strip that
    If (Right(tmpName, 1) = ")") Then
        fpos = InStrRev(tmpName, "(")
        If fpos > 0 Then
            tmpName = Trim(Left(tmpName, fpos - 1))
        End If
    End If

    fpos = InStr(tmpName, ",")
    If fpos > 0 Then
        'Firstname is separated by comma and positioned behind the lastname
        firstName = Trim(Mid(tmpName, fpos + 1))
        'Firstname field may include middle initial(s)
        Do While (UCase(Right(firstName, 2)) Like " [A-Z]" Or UCase(Right(firstName, 2)) Like "[A-Z].")
            firstName = Trim(Left(firstName, Len(firstName) - 2))
        Loop
        lastName = Trim(Left(tmpName, fpos - 1))
        'lastName field may have a formal suffix
        lastName = StripSuffixes(lastName)
    Else
        'Determining first and last name is really hard unless
        'there are only two names, or there is a middle initial(s)
        fpos = InStr(Trim(tmpName), " ")
        If fpos > 0 Then
            'First strip any possible, (single,) formal suffix on the name
            tmpName = StripSuffixes(tmpName)
            lpos = InStrRev(Trim(tmpName), " ")
            If fpos = lpos Then
                'single first name and last name separated by space
                firstName = Trim(Left(tmpName, fpos - 1))
                lastName = Trim(Mid(tmpName, lpos + 1))
                If firstName = UCase(firstName) And Not lastName = UCase(lastName) Then
                    'in case the firstName is written in uppercase letters (and not everything in capital letters),
                    'we assume that the sender's last name is the firstName (in the string)
                    lastName = firstName
                    firstName = Trim(Mid(tmpName, lpos + 1))
                End If
            Else
                'middle section could be a single/multiple name/initial (or both)
                Dim midName As String
                midName = Trim(Mid(Left(tmpName, lpos), fpos))
                Dim i, j As Integer

                'One or two initials are easy
                Do While Len(midName) = 1 Or _
                        Left(midName, 1) = "." Or _
                        Left(midName, 2) Like "[A-Z] " Or _
                        Left(midName, 2) Like "[A-Z]."
                    midName = Trim(Mid(midName, 2))
                    i = i + 1
                Loop
                Do While Right(midName, 2) Like " [A-Z]" Or _
                        Right(midName, 2) Like "[A-Z]."
                    midName = Trim(Left(midName, Len(midName) - 2))
                    j = j + 1
                Loop

                If Len(midName) = 0 Then
                    'initials only
                    firstName = Trim(Left(tmpName, fpos - 1))
                    lastName = Trim(Mid(tmpName, lpos + 1))
                ElseIf i <> 0 And j = 0 Then
                    'initials before double last name
                    lastName = midName + Trim(Mid(tmpName, lpos + 1))
                    firstName = Trim(Left(tmpName, fpos - 1))
                ElseIf i = 0 And j <> 0 Then
                    'initials after double first name
                    lastName = Trim(Mid(tmpName, lpos + 1))
                    firstName = Trim(Left(tmpName, fpos - 1)) + midName
                ElseIf Left(midName, 1) = LCase(Left(midName, 1)) Then
                    'Midname starts with a lower case letter
                    'We assume "correct" casing. Thus, we hit a name such as Firstname von Lastname
                    firstName = Trim(Left(tmpName, fpos - 1))
                    lastName = Trim(Mid(tmpName, fpos + 1))
                Else
                    'anything else can't be definitively identified as a first, middle or last name
                    firstName = tmpName
                    lastName = ""
                End If
            End If
        Else
            fpos = InStr(tmpName, "@")
            If fpos > 0 Then
                'first name is (currenty) an eMail-Adress. Just take the prefix
                tmpName = Left(tmpName, fpos - 1)
            End If
            fpos = InStr(tmpName, ".")
            If fpos > 0 Then
                'first name is separated by a dot
                lastName = Mid(tmpName, fpos + 1)
                tmpName = Left(tmpName, fpos - 1)
            Else
                'name is a single string, without "." or " "
                'final guess: LastnameFirstname
                If (IsUpperCaseChar(Left(tmpName, 1))) Then
                    i = 2
                    Dim UpperCaseCharCount As Integer
                    UpperCaseCharCount = 0
                    Dim LastUpperCaseCharPos As Integer
                    LastUpperCaseCharPos = 0
                    Do While (i < Len(tmpName) And (UpperCaseCharCount < 2))
                        If (IsUpperCaseChar(Mid(tmpName, i, 1))) Then
                            LastUpperCaseCharPos = i
                            UpperCaseCharCount = UpperCaseCharCount + 1
                        End If
                        i = i + 1
                    Loop
                    If (UpperCaseCharCount = 1) Then
                        'LastnameFirstname format found
                        tmpName = Mid(tmpName, LastUpperCaseCharPos)
                    End If
                End If
            End If
            firstName = tmpName
        End If
    End If

    Dim fnEmail As String
    Dim lnEmail As String
    Call getFirstNameLastNameOutOfEmail(senderEmailAddress, fnEmail, lnEmail)
    If (LCase(firstName) = lnEmail) And (LCase(lastName) = fnEmail) Then
        ' in case firstname and lastname are reversed in the email address, we assume that email format is firstname.lastname and reverse the names here
        Dim tmp As String
        tmp = firstName
        firstName = lastName
        lastName = tmp
    End If

    'fix casing of names
    If InStr(firstName, " ") = 0 Then
        firstName = FixCase(firstName)
    End If
    If InStr(lastName, " ") = 0 Then
        lastName = FixCase(lastName)
    End If
    senderName = Trim(firstName + " " + lastName)
End Sub

Public Sub getFirstNameLastNameOutOfEmail(ByRef email As String, ByRef firstName As String, ByRef lastName As String)
    If email = "" Then
        firstName = ""
        lastName = ""
        Exit Sub
    End If
    Dim parts() As String
    parts = Split(email, "@")
    Dim addressee As String
    addressee = parts(0)
    parts = Split(addressee, ".")

    Dim length As Long
    length = UBound(parts) - LBound(parts) + 1

    If (length <> 2) Then
        firstName = addressee
        lastName = ""
        Exit Sub
    End If
    firstName = parts(LBound(parts))
    lastName = parts(UBound(parts))
End Sub

Public Function removeDepartment(ByVal tmpName) As String
    Dim parts() As String
    parts = Split(tmpName, " ")

    Dim length As Long
    length = UBound(parts) - LBound(parts) + 1

    If length <= 2 Then
        removeDepartment = tmpName
        Exit Function
    End If

    Dim indexWordBeforeLastUppercasedWord As Long
    indexWordBeforeLastUppercasedWord = UBound(parts)
    Do While (indexWordBeforeLastUppercasedWord >= LBound(parts) + 2)
        If Not IsUpperCaseWord(parts(indexWordBeforeLastUppercasedWord)) Then
            Exit Do
        End If
        indexWordBeforeLastUppercasedWord = indexWordBeforeLastUppercasedWord - 1
    Loop

    If indexWordBeforeLastUppercasedWord < LBound(parts) + 1 Then
        removeDepartment = tmpName
        Exit Function
    End If

    Dim result As String
    Dim i As Long
    For i = LBound(parts) To indexWordBeforeLastUppercasedWord
       result = result + parts(i) + " "
    Next i
    removeDepartment = Left(result, Len(result) - 1)
End Function


'Extracts the name of the sender from the sender's name provided in the E-Mail.
'TODO: Future work is to extract the first name out of the stored Outlook contacts (if that contact exists)
'
'Notes:
'  * Names are returned by reference
Private Sub getNamesFromMail(ByRef item As MailItem, ByRef senderName As String, ByRef firstName As String, ByRef lastName As String)

    'Wildcard replacements
    senderName = item.SentOnBehalfOfName

    If senderName = "" Then
        senderName = item.senderName
    End If

    Call getNamesOutOfString(senderName, senderName, firstName, lastName, item.senderEmailAddress)
End Sub

'Code duplication of getNamesFromMail, because there is no common ancestor of MailItem and MeetingItem
Private Sub getNamesFromMeeting(ByRef item As MeetingItem, ByRef senderName As String, ByRef firstName As String, ByRef lastName As String)

    'Wildcard replacements
    senderName = item.SentOnBehalfOfName

    If senderName = "" Then
        senderName = item.senderName
    End If

    Call getNamesOutOfString(senderName, senderName, firstName, lastName, item.senderEmailAddress)
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
        Set objApp = session.Application

        Select Case TypeName(objApp.ActiveWindow)
            Case "Explorer":  'on clicking reply in the main window
                Set GetCurrentItem = objApp.ActiveExplorer.Selection.item(1)
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


'Changes
'
' >
' >
'
'To
'
' >
'
Private Function cleanUpDoubleLines(quotedText As String) As String
    Dim quoteLines() As String
    Dim curLine As String
    Dim res As String
    Dim i As Integer
    Dim previousLineWasEmptyQuote As Boolean

    previousLineWasEmptyQuote = False
    quoteLines = Split(quotedText, vbCrLf)

    For i = 0 To UBound(quoteLines)
        If (quoteLines(i) = "> ") Then
            If Not previousLineWasEmptyQuote Then
                previousLineWasEmptyQuote = True
                res = res + quoteLines(i) + vbCrLf
            End If
        Else
            previousLineWasEmptyQuote = False
            res = res + quoteLines(i) + vbCrLf
        End If
    Next i

    cleanUpDoubleLines = res
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


Public Function ColorizeMailItem(MyMailItem As MailItem) As String
    Dim folder As MAPIFolder
    Dim rtf  As String, lines() As String, resRTF As String
    Dim i As Integer, n As Integer, ret As Integer


    'save the mailitem to get an entry id, then forget reference to that rtf gets commited.
    'display mailitem by id later on.
    If ((Not MyMailItem.bodyFormat = olFormatPlain)) Then 'we just understand Plain Mails
        ColorizeMailItem = ""
        Exit Function
    End If

    'richt text it
    MyMailItem.bodyFormat = olFormatRichText
    MyMailItem.Save  'need to save to be able to access rtf via EntryID (.save creates ExtryID if not saved before)!

    Set folder = session.GetDefaultFolder(olFolderInbox)

    rtf = Space(99999)  'init rtf to max length of message!
    ret = ReadRTF(session.CurrentProfileName, MyMailItem.EntryID, folder.StoreID, rtf)
    If (ret = 0) Then
        'ole call success!!!
        rtf = Trim(rtf)  'kill unnecessary spaces (from rtf var init with Space(rtf))
        Debug.Print rtf & vbCrLf & "*************************************************************" & vbCrLf

        'we have our own rtf haeder, remove generated one
        Dim PosHeaderEnd As Integer
        Dim sTestString As String
        PosHeaderEnd = InStr(rtf, "\uc1\pard\plain\deftab360")
        If (PosHeaderEnd = 0) Then
            sTestString = "\uc1\pard\f0\fs20\lang1031"
            PosHeaderEnd = InStr(rtf, sTestString)
        End If
        If (PosHeaderEnd = 0) Then
            sTestString = "\viewkind4\uc1\pard\f0\fs20"
            PosHeaderEnd = InStr(rtf, sTestString)
        End If
        If (PosHeaderEnd = 0) Then
            sTestString = "\pard\f0\fs20\lang1031"
            PosHeaderEnd = InStr(rtf, sTestString)
        End If

        rtf = Mid(rtf, PosHeaderEnd + Len(sTestString))

        rtf = "{\rtf1\ansi\ansicpg1252 \deff0{\fonttbl" & vbCrLf & _
                "{\f0\fswiss\fcharset0 Courier New;}}" & vbCrLf & _
                "{\colortbl\red0\green0\blue0;\red106\green44\blue44;\red44\green106\blue44;\red44\green44\blue106;}" & vbCrLf & _
                rtf

        lines = Split(rtf, vbCrLf)
        Dim s As String
        For i = LBound(lines) To UBound(lines)
            n = QuoteFixMacro.CalcNesting(lines(i)).level
            If (n = 0) Then
                resRTF = resRTF & lines(i) & vbCrLf
            Else
                If (Right(lines(i), 4) = "\par") Then
                    s = Left(lines(i), Len(lines(i)) - Len("\par"))
                    resRTF = resRTF & "\cf" & n Mod NUM_RTF_COLORS & " " & s & "\cf0  " & "\par" & vbCrLf
                Else
                    resRTF = resRTF & "\cf" & n Mod NUM_RTF_COLORS & " " & lines(i) & "\cf0  " & vbCrLf
                End If
            End If
        Next i
    Else
        Debug.Print "error while reading rtf! " & ret
        ColorizeMailItem = ""
        Exit Function
    End If

    'remove some rtf commands
    resRTF = Replace(resRTF, "\viewkind4\uc1", "")
    resRTF = Replace(resRTF, "\uc1", "")
    'VERY IMPORTANT, outlook will change the message back to PlainText otherwise!!!
    resRTF = Replace(resRTF, "\fromtext", "")
    Debug.Print resRTF

    'write RTF back to form
    ret = WriteRTF(session.CurrentProfileName, MyMailItem.EntryID, folder.StoreID, resRTF)
    If (ret = 0) Then
        Debug.Print "rtf write okay"
    Else
        Debug.Print "rtf write FAILURE"
        ColorizeMailItem = ""
        Exit Function
    End If

    'dereference all objects! otherwise, rtf isn't going to be updated!
    Set folder = Nothing
    'save return value
    ColorizeMailItem = MyMailItem.EntryID
    Set MyMailItem = Nothing
End Function


Public Sub DisplayMailItemByID(id As String)
    Dim it As MailItem
    Set it = session.GetItemFromID(id, session.GetDefaultFolder(olFolderInbox).StoreID)
    it.Display
    Set it = Nothing
End Sub

Private Function StripSuffixes(ByRef tempName As String) As String
    'Create array of possible suffixes
    Dim NameSuffixesArr() As String
    NameSuffixesArr = Split(LASTNAME_SUFFIXES, "/")
    Dim i As Integer

    'Strip the last suffix (is it ever the case that someone has multiple suffixes?)
    For i = LBound(NameSuffixesArr) To UBound(NameSuffixesArr)
        If (Right(tempName, Len(NameSuffixesArr(i)) + 1)) = " " & NameSuffixesArr(i) Then
            StripSuffixes = Trim(Left(tempName, Len(tempName) - Len(NameSuffixesArr(i))))
        End If
    Next
    StripSuffixes = tempName
End Function
