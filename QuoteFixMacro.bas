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

'@Folder("QuoteFixMacro")
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
'   SaveSetting APPNAME, REG_GROUP_CONFIG, "CONVERT_TO_PLAIN", "true"
'Finally, or by manually creating entries in this registry hive:
'    HKEY_CURRENT_USER\Software\VB and VBA Program Settings\QuoteFixMacro
Private Const APPNAME As String = "QuoteFixMacro"
Private Const REG_GROUP_CONFIG As String = "Config"
Private Const REG_GROUP_FIRSTNAMES As String = "Firstnames" 'stores replacements for firstnames


'--------------------------------------------------------
'*** Feature QuoteColorizer ***
'--------------------------------------------------------
Private Const DEFAULT_USE_COLORIZER As Boolean = False
'TODO: add note where to get the DLL from. I couldn't find it on my system
'If you enable it, you need MAPIRTF.DLL in C:\Windows\System32
'Does NOT work at Windows 7/64bit Outlook 2010/32bit
'
'Please enable convert RTF-to-Text at sending. Otherwise, the recipients will always receive HTML emails

'How many different colors should be used for colorizing the quotes?
Private Const DEFAULT_NUM_RTF_COLORS As Long = 4


'--------------------------------------------------------
'*** Feature SoftWrap ***
'--------------------------------------------------------
'Enable SoftWrap
'resize window so that the text editor wraps the text automatically
'after N characters. Outlook wraps text automatically after sending it,
'but doesn't display the wrap when editing
'you can edit the auto wrap setting at "File > Options > Mail > Message format > Remove extra line breaks in plain text messages
Private Const DEFAULT_USE_SOFTWRAP As Boolean = False

'put as much characters as set in Outlook at "File > Options > Mail > Message format > Automatically wrap text at character"
'default: 76 characters
Private Const DEFAULT_SEVENTY_SIX_CHARS As String = "123456789x123456789x123456789x123456789x123456789x123456789x123456789x123456"

'This constant has to be adapted to fit your needs (incorporating the used font, display size, ...)
Private Const DEFAULT_PIXEL_PER_CHARACTER As Double = 8.61842105263158


'--------------------------------------------------------
'*** Configuration constants ***
'--------------------------------------------------------
'If <> -1, strip quotes with level > INCLUDE_QUOTES_TO_LEVEL
Private Const DEFAULT_INCLUDE_QUOTES_TO_LEVEL As Long = -1

'At which column should the text be wrapped?
Private Const DEFAULT_LINE_WRAP_AFTER As Long = 75

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
Private Const DEFAULT_QUOTING_TEMPLATE As String = "Dear %FN,\n\n(reply inline)\n\nYou wrote on %D:\n\n%Q\n\nCheers,\n\n{Name}\n\n(Reply inline - powered by https://macros4outlook.github.io/quotefixmacro/)"

'English quote template
Private Const DEFAULT_QUOTING_TEMPLATE_EN As String = "Dear %FN,\n\n(reply inline)\n\nYou wrote on %D:\n\n%Q\n\nCheers,\n\n{Name}\n\n(Reply inline - powered by https://macros4outlook.github.io/quotefixmacro/)"

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
'Private Const OUTLOOK_PLAIN_ORIGINALMESSAGE As String = "-----Ursprüngliche Nachricht-----"
'Private Const OUTLOOK_PLAIN_ORIGINALMESSAGE As String = "-----Original Message-----"
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
Private NUM_RTF_COLORS As Long
Private USE_SOFTWRAP As Boolean
Private SEVENTY_SIX_CHARS As String
Private PIXEL_PER_CHARACTER As Double
Private INCLUDE_QUOTES_TO_LEVEL As Long
Private LINE_WRAP_AFTER As Long
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


'TODO: 1: check if these can also be changed into `Long`s. Unfortunately I
'         don't have the DLL and therefore can't test it myself
'For QuoteColorizer
Public Declare PtrSafe Function WriteRTF _
        Lib "mapirtf.dll" _
        Alias "writertf" (ByVal ProfileName As String, _
                          ByVal MessageID As String, _
                          ByVal StoreID As String, _
                          ByVal cText As String) _
        As Integer      ' <-- {1}

'For QuoteColorizer
Public Declare PtrSafe Function ReadRTF _
        Lib "mapirtf.dll" _
        Alias "readrtf" (ByVal ProfileName As String, _
                         ByVal SrcMsgID As String, _
                         ByVal SrcStoreID As String, _
                         ByRef MsgRTF As String) _
        As Integer      '<-- {1}


Private Enum ReplyType
    TypeReply = 1
    TypeReplyAll = 2
    TypeForward = 3
End Enum

Public Type NestingType
    'the level of the current quote plus
    level As Long

    'the amount of spaces until the next word
    'needed as outlook sometimes inserts more than one space to separate the quoteprefix and the actual quote
    'we use that information to fix the quote
    additionalSpacesCount As Long

    'total = level + additionalSpacesCount + 1
    total As Long
End Type

'Module Variables to make code more readable (-> parameter passing gets easier)
Private result As String
Private unformattedBlock As String
Private curBlock As String
Private curBlockNeedsToBeReFormatted As Boolean
Private curPrefix As String
Private lastLineWasParagraph As Boolean
Private lastNesting As NestingType

'"Fixed Reply" functionality - has to be made available as shortcut in Outlook
Public Sub FixedReply()
    Dim m As Object
    Set m = GetCurrentItem()

    FixMailText m, TypeReply
End Sub

'"Fixed Reply All" functionality - has to be made available as shortcut in Outlook
Public Sub FixedReplyAll()
    Dim m As Object
    Set m = GetCurrentItem()

    FixMailText m, TypeReplyAll
End Sub

'"Fixed Reply All" functionality with English template
Public Sub FixedReplyAllEnglish()
    Dim m As Object
    Set m = GetCurrentItem()

    FixMailText m, TypeReplyAll, True
End Sub

'"Fixed Forward" functionality - has to be made available as shortcut in Outlook
Public Sub FixedForward()
    Dim m As Object
    Set m = GetCurrentItem()

    FixMailText m, TypeForward
End Sub

Private Function CalcNesting(ByVal line As String) As NestingType

    Dim count As Long
    count = 0

    Dim i As Long
    i = 1

    Do While i <= Len(line)
        Dim curChar As String
        curChar = Mid$(line, i, 1)
        If curChar = ">" Then
            count = count + 1
            Dim lastQuoteSignPos As Long
            lastQuoteSignPos = i
        ElseIf curChar <> " " Then
            'Char is neither ">" nor " " - Quote intro ended
            'leave function
            Exit Do
        End If
        i = i + 1
    Loop

    Dim res As NestingType
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
    SaveSetting APPNAME, REG_GROUP_CONFIG, "USE_COLORIZER", DEFAULT_USE_COLORIZER
    SaveSetting APPNAME, REG_GROUP_CONFIG, "NUM_RTF_COLORS", DEFAULT_NUM_RTF_COLORS
    SaveSetting APPNAME, REG_GROUP_CONFIG, "USE_SOFTWRAP", DEFAULT_USE_SOFTWRAP
    SaveSetting APPNAME, REG_GROUP_CONFIG, "SEVENTY_SIX_CHARS", DEFAULT_SEVENTY_SIX_CHARS
    SaveSetting APPNAME, REG_GROUP_CONFIG, "PIXEL_PER_CHARACTER", DEFAULT_PIXEL_PER_CHARACTER
    SaveSetting APPNAME, REG_GROUP_CONFIG, "INCLUDE_QUOTES_TO_LEVEL", DEFAULT_INCLUDE_QUOTES_TO_LEVEL
    SaveSetting APPNAME, REG_GROUP_CONFIG, "LINE_WRAP_AFTER", DEFAULT_LINE_WRAP_AFTER
    SaveSetting APPNAME, REG_GROUP_CONFIG, "DATE_FORMAT", DEFAULT_DATE_FORMAT
    SaveSetting APPNAME, REG_GROUP_CONFIG, "STRIP_SIGNATURE", DEFAULT_STRIP_SIGNATURE
    SaveSetting APPNAME, REG_GROUP_CONFIG, "CONVERT_TO_PLAIN", DEFAULT_CONVERT_TO_PLAIN
    SaveSetting APPNAME, REG_GROUP_CONFIG, "USE_QUOTING_TEMPLATE", DEFAULT_USE_QUOTING_TEMPLATE
    SaveSetting APPNAME, REG_GROUP_CONFIG, "QUOTING_TEMPLATE", DEFAULT_QUOTING_TEMPLATE
    SaveSetting APPNAME, REG_GROUP_CONFIG, "QUOTING_TEMPLATE_EN", DEFAULT_QUOTING_TEMPLATE_EN
    SaveSetting APPNAME, REG_GROUP_CONFIG, "CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS", DEFAULT_CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS
    SaveSetting APPNAME, REG_GROUP_CONFIG, "CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER", DEFAULT_CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER
    SaveSetting APPNAME, REG_GROUP_CONFIG, "CONDENSED_HEADER_FORMAT", DEFAULT_CONDENSED_HEADER_FORMAT
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
    QUOTING_TEMPLATE = Replace$(QUOTING_TEMPLATE, "\n", vbCrLf)

    QUOTING_TEMPLATE_EN = GetSetting(APPNAME, REG_GROUP_CONFIG, "QUOTING_TEMPLATE_EN", DEFAULT_QUOTING_TEMPLATE_EN)
    QUOTING_TEMPLATE_EN = Replace$(QUOTING_TEMPLATE_EN, "\n", vbCrLf)

    Dim count As Variant
    count = CDbl(GetSetting(APPNAME, REG_GROUP_FIRSTNAMES, "Count", 0))
    ReDim FIRSTNAME_REPLACEMENT__EMAIL(count)
    ReDim FIRSTNAME_REPLACEMENT__FIRSTNAME(count)

    Dim i As Long
    For i = 1 To count
        Dim group As String
        group = REG_GROUP_FIRSTNAMES & "\" & i
        FIRSTNAME_REPLACEMENT__EMAIL(i) = GetSetting(APPNAME, group, "email", vbNullString)
        FIRSTNAME_REPLACEMENT__FIRSTNAME(i) = GetSetting(APPNAME, group, "firstName", vbNullString)
    Next
End Sub

'Description:
'   Strips away ">" and " " at the beginning to have the plain text
Private Function StripLine(ByVal line As String) As String
    Dim res As String
    res = line

    Do While (Len(res) > 0) And (InStr("> ", Left$(res, 1)) <> 0)
        'First character is a space or a quote
        res = Mid$(res, 2)
    Loop

    'Remove the spaces at the end of res
    res = Trim$(res)

    StripLine = res
End Function

Private Function CalcPrefix(ByRef nesting As NestingType) As String
    Dim res As String

    res = String$(nesting.level, ">")
    res = res & String$(nesting.additionalSpacesCount, " ")

    CalcPrefix = res & " "
End Function

'Description:
'   Adds the current line to unformattedBlock and to curBlock
Private Sub AppendCurLine(ByVal curLine As String)
    If Len(unformattedBlock) = 0 Then
        'unformattedBlock has to be used here, because it might be the case that the first
        '  line is "". Therefore curBlock remains "", whereas unformattedBlock gets <> ""

        If Len(curLine) = 0 Then Exit Sub

        curBlock = curLine
        unformattedBlock = curPrefix & curLine & vbCrLf
    Else
        curBlock = curBlock & IIf(Len(curBlock) = 0, vbNullString, " ") & curLine
        unformattedBlock = unformattedBlock & curPrefix & curLine & vbCrLf
    End If
End Sub

Private Sub HandleParagraph(ByVal prefix As String)
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
        prefix = CalcPrefix(nesting)

        Dim maxLength As Long
        maxLength = LINE_WRAP_AFTER - nesting.total

        Do While Len(curBlock) > maxLength
            'go through block from maxLength to beginning to find a space
            Dim i As Long
            i = maxLength
            If i > 0 Then
                Do While (Mid$(curBlock, i, 1) <> " ")
                    i = i - 1
                    If i = 0 Then Exit Do
                Loop
            End If

            If i = 0 Then
                'No space found -> use the full line
                Dim curLine As String
                curLine = Left$(curBlock, maxLength)
                curBlock = Mid$(curBlock, maxLength + 1)
            Else
                curLine = Left$(curBlock, i - 1)
                curBlock = Mid$(curBlock, i + 1)
            End If

            result = result & prefix & curLine & vbCrLf
        Loop

        If Len(curBlock) > 0 Then
            result = result & prefix & curBlock & vbCrLf
        End If
    End If

    'Resetting
    curBlockNeedsToBeReFormatted = False
    curBlock = vbNullString
    unformattedBlock = vbNullString
    'lastLineWasParagraph = False
End Sub

'Reformat text to correct broken wrap inserted by Outlook.
'Needs to be public so the test cases can run this function.
Public Function ReFormatText(ByVal text As String) As String
    'Reset (partially global) variables
    result = vbNullString
    curBlock = vbNullString
    unformattedBlock = vbNullString
    Dim curNesting As NestingType
    curNesting.level = 0
    lastNesting.level = 0
    curBlockNeedsToBeReFormatted = False

    Dim rows() As String
    rows = Split(text, vbCrLf)

    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        Dim curLine As String
        curLine = StripLine(rows(i))
        lastNesting = curNesting
        curNesting = CalcNesting(rows(i))

        If curNesting.total <> lastNesting.total Then
            Dim lastPrefix As String
            lastPrefix = curPrefix
            curPrefix = CalcPrefix(curNesting)
        End If

        If curNesting.total = lastNesting.total Then
            'Quote continues
            If Len(curLine) = 0 Then
                'new paragraph has started
                HandleParagraph curPrefix
            Else
                AppendCurLine curLine
                lastLineWasParagraph = False

                If (curNesting.level = 1) And (i < UBound(rows)) Then
                    'check if the next line contains a wrong break
                    Dim nextNesting As NestingType
                    nextNesting = CalcNesting(rows(i + 1))
                    If (CountOccurrencesOfStringInString(curLine, " ") = 0) And (curNesting.total = nextNesting.total) _
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

                    If Len(curLine) = 0 Then
                        'The line break has to be interpreted as paragraph
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

                    If Len(curLine) = 0 Then
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
                FinishBlock lastNesting
            End If

        Else
            'curNesting.total > lastNesting.total

            lastLineWasParagraph = False

            'it's nested one level deeper. Current block is finished
            FinishBlock lastNesting

            If Len(curLine) = 0 Then
                If curNesting.level <> lastNesting.level Then
                    lastLineWasParagraph = True
                    HandleParagraph curPrefix
                End If
            End If

            If CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS Then
                If Left$(curLine, Len(OUTLOOK_PLAIN_ORIGINALMESSAGE)) = OUTLOOK_PLAIN_ORIGINALMESSAGE _
                And Not Left$(curLine, Len(PGP_MARKER)) = PGP_MARKER _
                Then
                    'We found a header

                    'Name and Email
                    i = i + 1
                    curLine = StripLine(rows(i))

                    Dim posColon As Long
                    posColon = InStr(curLine, ":")

                    Dim posLeftBracket As String
                    posLeftBracket = InStr(curLine, "[")  '[ is the indication of the beginning of the email address

                    Dim posRightBracket As Long
                    posRightBracket = InStr(curLine, "]")

                    If (posLeftBracket) > 0 Then
                        Dim lengthName As Long
                        lengthName = posLeftBracket - posColon - 3

                        If lengthName > 0 Then
                            Dim sName As String
                            sName = Mid$(curLine, posColon + 2, lengthName)
                        Else
                            Debug.Print "Could not get name. Is the header formatted correctly?"
                        End If

                        If posRightBracket = 0 Then
                            Dim sEmail As String
                            sEmail = Mid$(curLine, posLeftBracket + 8) '8 = Len("mailto: ")
                        Else
                            sEmail = Mid$(curLine, posLeftBracket + 8, posRightBracket - posLeftBracket - 8) '8 = Len("mailto: ")
                        End If
                    Else
                        sName = Mid$(curLine, posColon + 2)
                        sEmail = vbNullString
                    End If

                    i = i + 1
                    curLine = StripLine(rows(i))
                    If InStr(curLine, ":") = 0 Then
                        'There is a wrap in the email address
                        posRightBracket = InStr(curLine, "]")
                        If posRightBracket > 0 Then
                            sEmail = sEmail & Left$(curLine, posRightBracket - 1)
                        Else
                            'something went wrong, do nothing
                        End If
                        'go to next line
                        i = i + 1
                        curLine = StripLine(rows(i))
                    End If

                    'Date
                    'We assume that there is always a weekday present before the date
                    Dim sDate As String
                    sDate = StripLine(rows(i))
                    'posColon = InStr$(sDate, ":")
                    'sDate = Mid$(sDate, posColon + 2)
                    Dim posFirstComma As Long
                    posFirstComma = InStr(sDate, ",")
                    sDate = Mid$(sDate, posFirstComma + 2)
                    Dim dDate As Date
                    If IsDate(sDate) Then
                        dDate = DateValue(sDate)
                        'there is no function "IsTime", therefore try with brute force
                        dDate = dDate + TimeValue(sDate)
                    End If
                    If dDate <> CDate("00:00:00") Then
                        sDate = Format$(dDate, DATE_FORMAT)
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
                    Loop Until (Len(curLine) > 0) Or (i = UBound(rows))
                    i = i - 1 'i now points to the last empty line

                    Dim condensedHeader As String
                    condensedHeader = CONDENSED_HEADER_FORMAT
                    condensedHeader = Replace$(condensedHeader, PATTERN_SENDER_NAME, sName)
                    condensedHeader = Replace$(condensedHeader, PATTERN_SENT_DATE, sDate)
                    condensedHeader = Replace$(condensedHeader, PATTERN_SENDER_EMAIL, sEmail)

                    Dim prefix As String
                    'the prefix for the result has to be one level shorter as it is the quoted text from the sender
                    If (curNesting.level = 1) Then
                        prefix = vbNullString
                    Else
                        prefix = Mid$(curPrefix, 2)
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
    Next

    'Finish current Block
    FinishBlock curNesting

    'strip last (unnecessary) line feeds and spaces
    Do While ((Len(result) > 0) And (InStr(vbCrLf & " ", Right$(result, 1)) <> 0))
        result = Left$(result, Len(result) - 1)
    Loop

    ReFormatText = result
End Function

' @param UseEnglishTemplate In case USE_QUOTING_TEMPLATE is True, should the default or the English template be used?
Private Sub FixMailText(ByVal SelectedObject As Object, ByRef MailMode As ReplyType, Optional ByVal UseEnglishTemplate As Boolean = False)
    LoadConfiguration

    'we only understand mail items and meeting items , no PostItems, NoteItems, ...
    If Not (TypeName(SelectedObject) = "MailItem") And _
    Not (TypeName(SelectedObject) = "MeetingItem") Then
        On Error GoTo catch   'try, catch replacement
        Dim HadError As Boolean
        HadError = True

        Select Case MailMode
            Case TypeReply
                Dim TempObj As Object
                Set TempObj = SelectedObject.Reply
                TempObj.Display
                HadError = False
                Exit Sub
            Case TypeReplyAll
                Set TempObj = SelectedObject.ReplyAll
                TempObj.Display
                HadError = False
                Exit Sub
            Case TypeForward
                Set TempObj = SelectedObject.Forward
                TempObj.Display
                HadError = False
                Exit Sub
        End Select

catch:
        On Error GoTo 0  'deactivate error handling

        If (HadError = True) Then
            'reply / reply all / forward caused error
            ' -->  just display it
            SelectedObject.Display
            Exit Sub
        End If
    End If

    Dim isMail As Boolean
    isMail = (TypeName(SelectedObject) = "MailItem")

    If isMail Then
        Dim OriginalMail As MailItem
        Set OriginalMail = SelectedObject 'cast!
    Else
        Dim OriginalMeeting As MeetingItem
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
        ' `MeetingItem.BodyFormat` doesn't exist in Outlook 2016 and causes a runtime error --> skip it
        On Error Resume Next
        bodyFormat = OriginalMeeting.bodyFormat
        On Error GoTo 0
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
                Case TypeReply
                    If isMail Then
                        Set ReplyObj = OriginalMail.Reply
                    Else
                        Set ReplyObj = OriginalMeeting.Reply
                    End If
                Case TypeReplyAll
                    If isMail Then
                        Set ReplyObj = OriginalMail.ReplyAll
                    Else
                        Set ReplyObj = OriginalMeeting.ReplyAll
                    End If
                Case TypeForward
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

    '''create reply --> outlook style!
    ''Actions(1) = Actions("Reply")' or 'Actions("Antworten")' respectively, etc.
    If isMail Then
        With OriginalMail.Actions(MailMode)
            Dim OriginalReplyStyle As OlActionReplyStyle
            OriginalReplyStyle = .ReplyStyle
            .ReplyStyle = olReplyTickOriginalText

            Dim NewMail As MailItem
            Set NewMail = .Execute

            .ReplyStyle = OriginalReplyStyle
        End With
    Else
        With OriginalMeeting.Actions(MailMode)
            OriginalReplyStyle = .ReplyStyle
            .ReplyStyle = olReplyTickOriginalText

            Set NewMail = .Execute

            .ReplyStyle = OriginalReplyStyle
        End With
    End If

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
        getNamesFromMail OriginalMail, senderName, firstName, lastName
    Else
        getNamesFromMeeting OriginalMeeting, senderName, firstName, lastName
    End If

    If (UBound(FIRSTNAME_REPLACEMENT__EMAIL) > 0) Or (InStr(MySignature, PATTERN_SENDER_EMAIL) <> 0) Then
        Dim senderEmail As String
        If isMail Then
            senderEmail = getSenderEmailAddress(OriginalMail.senderEmailType, senderName, OriginalMail.senderEmailAddress, OriginalMail.session)
        Else
            senderEmail = getSenderEmailAddress(OriginalMeeting.senderEmailType, senderName, OriginalMeeting.senderEmailAddress, OriginalMeeting.session)
        End If
        MySignature = Replace$(MySignature, PATTERN_SENDER_EMAIL, senderEmail)
    End If

    If (UBound(FIRSTNAME_REPLACEMENT__EMAIL) > 0) Then
        'replace firstName by email stored in registry
        Dim curIndex As Long
        For curIndex = 1 To UBound(FIRSTNAME_REPLACEMENT__EMAIL)
            Dim rEmail As Variant
            rEmail = FIRSTNAME_REPLACEMENT__EMAIL(curIndex)
            If (StrComp(LCase$(senderEmail), LCase$(rEmail)) = 0) Then
                firstName = FIRSTNAME_REPLACEMENT__FIRSTNAME(curIndex)
                Exit For
            End If
        Next
    End If

    MySignature = Replace$(MySignature, PATTERN_FIRST_NAME, firstName)
    MySignature = Replace$(MySignature, PATTERN_LAST_NAME, lastName)
    If isMail Then
        MySignature = Replace$(MySignature, PATTERN_SENT_DATE, Format$(OriginalMail.SentOn, DATE_FORMAT))
    Else
        MySignature = Replace$(MySignature, PATTERN_SENT_DATE, Format$(OriginalMeeting.SentOn, DATE_FORMAT))
    End If
    MySignature = Replace$(MySignature, PATTERN_SENDER_NAME, senderName)

    Dim OutlookHeader As String
    If CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER Then
        OutlookHeader = vbNullString
        'The real condensing is made below, where CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS is checked
        'Disabling getOutlookHeader leads to an unmodified lineCounter, which in turn gets the header included in "quotedText"
    Else
        OutlookHeader = getOutlookHeader(BodyLines, lineCounter, MailMode)
    End If

    Dim quotedText As String
    quotedText = getQuotedText(BodyLines, lineCounter)

    Dim NewText As String
    'create mail according to reply mode
    Select Case MailMode
        Case TypeReply
            NewText = quotedText
        Case TypeReplyAll
            NewText = quotedText
        Case TypeForward
            NewText = OutlookHeader & quotedText
    End Select

    'Put text in signature (=Template for text)
    MySignature = Replace$(MySignature, PATTERN_OUTLOOK_HEADER & vbCrLf, OutlookHeader)

    'Stores number of downs to send
    Dim downCount As Long
    downCount = -1

    If InStr(MySignature, PATTERN_QUOTED_TEXT) <> 0 Then
        If InStr(MySignature, PATTERN_CURSOR_POSITION) = 0 Then
            'if PATTERN_CURSOR_POSITION is not set, but PATTERN_QUOTED_TEXT is, then the cursor is moved to the quote
            downCount = CalcDownCount(PATTERN_QUOTED_TEXT, MySignature)
        End If
        MySignature = Replace$(MySignature, PATTERN_QUOTED_TEXT, NewText)
    Else
        'There's no placeholder. Fall back to outlook behavior
        MySignature = vbCrLf & vbCrLf & MySignature & OutlookHeader & NewText
    End If

    If (InStr(MySignature, PATTERN_CURSOR_POSITION) <> 0) Then
        downCount = CalcDownCount(PATTERN_CURSOR_POSITION, MySignature)
        'remove cursor_position pattern from mail text
        MySignature = Replace$(MySignature, PATTERN_CURSOR_POSITION, vbNullString)
    End If

    MySignature = cleanUpDoubleLines(MySignature)

    NewMail.Body = MySignature

    'Extensions, in case Colorize is activated
    If USE_COLORIZER Then
        Dim mailID As String
        mailID = ColorizeMailItem(NewMail)
        If (Len(Trim$(vbNullString & mailID)) > 0) Then  'no error occurred or quotefix macro not there...
            DisplayMailItemByID mailID
        Else
            'Display window
            NewMail.Display
        End If
    Else
        'Display window
        NewMail.Display
    End If

    'jump to the right place
    Dim i As Long
    For i = 1 To downCount
        SendKeys "{DOWN}"
    Next

    If USE_SOFTWRAP Then
        ResizeWindowForSoftWrap
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
    Next
End Function

Private Function getSenderEmailAddress(ByVal senderEmailType As String, ByVal senderName As String, ByVal senderEmailAddress As String, ByVal session As NameSpace) As String
    Dim senderEmail As String

    If senderEmailType = "SMTP" Then
        senderEmail = senderEmailAddress

    ElseIf senderEmailType = "EX" Then
        'FIXME: This seems only to work in Outlook 2007
        Dim gal As Outlook.AddressList
        Set gal = session.GetGlobalAddressList
        Dim exchAddressEntries As Outlook.AddressEntries
        Set exchAddressEntries = gal.AddressEntries

        'check if we can get the correct item by sendername
        Dim exchAddressEntry As Outlook.AddressEntry
        Set exchAddressEntry = exchAddressEntries.item(senderName)
        If exchAddressEntry.Name <> senderName Then Set exchAddressEntry = exchAddressEntries.GetFirst

        Dim found As Boolean
        found = False
        Do While (Not found) And (Not exchAddressEntry Is Nothing)
            found = (LCase$(exchAddressEntry.Address) = LCase$(senderEmailAddress))
            If Not found Then Set exchAddressEntry = exchAddressEntries.GetNext
        Loop

        If Not exchAddressEntry Is Nothing Then
            senderEmail = exchAddressEntry.GetExchangeUser.PrimarySmtpAddress
        Else
            senderEmail = vbNullString
        End If
    End If

    getSenderEmailAddress = senderEmail
End Function

'NOTE: not used --> delete it?
Private Function IsWordCased(ByVal word As String) As Boolean
    IsWordCased = (word Like "[A-Z][a-z]*") Or (word Like "[A-Z][a-z]*-[A-Z][a-z]*")
End Function

Private Function getOutlookHeader(ByRef BodyLines() As String, ByRef lineCounter As Long, ByRef MailMode As ReplyType) As String
    ' parse until we find the header finish "> " (Outlook_Headerfinish)

    For lineCounter = lineCounter To UBound(BodyLines)
        If (BodyLines(lineCounter) = OUTLOOK_HEADERFINISH) Then
            Exit For
        End If
        getOutlookHeader = getOutlookHeader & BodyLines(lineCounter) & vbCrLf
    Next

    'skip OUTLOOK_HEADERFINISH for replies
    If Not MailMode = TypeForward Then
        lineCounter = lineCounter + 1
    End If

End Function


Private Function getQuotedText(ByRef BodyLines() As String, ByRef lineCounter As Long) As String
    ' parse the rest of the message
    For lineCounter = lineCounter To UBound(BodyLines)
        If STRIP_SIGNATURE And (BodyLines(lineCounter) = SIGNATURE_SEPARATOR) Then
            'beginning of signature reached
            Exit For
        End If

        getQuotedText = getQuotedText & BodyLines(lineCounter) & vbCrLf
    Next

    getQuotedText = ReFormatText(getQuotedText)

    If INCLUDE_QUOTES_TO_LEVEL <> -1 Then
        getQuotedText = StripQuotes(getQuotedText, INCLUDE_QUOTES_TO_LEVEL)
    End If
End Function


Private Function CalcDownCount(ByVal pattern As String, ByVal textToSearch As String) As Long
    Dim PosOfPattern As Long
    PosOfPattern = InStr(textToSearch, pattern)

    Dim TextBeforePattern As String
    TextBeforePattern = Left$(textToSearch, PosOfPattern - 1)

    CalcDownCount = CountOccurrencesOfStringInString(TextBeforePattern, vbCrLf)
End Function


Private Function GetCurrentItem() As Object  'changed to default scope
        Dim objApp As Application
        Set objApp = session.Application

        Select Case TypeName(objApp.ActiveWindow)
            Case "Explorer"  'on clicking reply in the main window
                Set GetCurrentItem = objApp.ActiveExplorer.Selection.item(1)
            Case "Inspector" 'on clicking reply when mail is shown in separate window
                Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
        End Select

End Function

'Parameters:
'  InString: String to count in
'  What:     What to count
'Note:
'  * Order of parameters taken from "InStr"
Public Function CountOccurrencesOfStringInString(ByVal InString As String, ByVal What As String) As Long
    Dim count As Long
    count = 0

    Dim lastPos As Long
    lastPos = 0

    Dim curPos As Long
    curPos = InStr(InString, What)

    Do While curPos <> 0
        lastPos = curPos + 1
        count = count + 1
        curPos = InStr(lastPos, InString, What)
    Loop

    CountOccurrencesOfStringInString = count
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
Private Function cleanUpDoubleLines(ByVal quotedText As String) As String
    Dim previousLineWasEmptyQuote As Boolean
    previousLineWasEmptyQuote = False

    Dim quoteLines() As String
    quoteLines = Split(quotedText, vbCrLf)

    Dim i As Long
    For i = 0 To UBound(quoteLines)
        If (quoteLines(i) = "> ") Then
            If Not previousLineWasEmptyQuote Then
                previousLineWasEmptyQuote = True
                Dim res As String
                res = res & quoteLines(i) & vbCrLf
            End If
        Else
            previousLineWasEmptyQuote = False
            res = res & quoteLines(i) & vbCrLf
        End If
    Next

    cleanUpDoubleLines = res
End Function


Private Function StripQuotes(ByVal quotedText As String, ByVal stripLevel As Long) As String
    Dim quoteLines() As String
    quoteLines = Split(quotedText, vbCrLf)

    Dim i As Long
    For i = 1 To UBound(quoteLines)
        Dim level As Long
        level = InStr(quoteLines(i), " ") - 1
        If level <= stripLevel Then
            Dim res As String
            res = res & quoteLines(i) & vbCrLf
        End If
    Next

    StripQuotes = res
End Function


'resize window so that the text editor wraps the text automatically
'after N characters. Outlook wraps text automatically after sending it,
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
    'save the mailitem to get an entry id, then forget reference to that rtf gets committed.
    'display mailitem by id later on.
    If ((Not MyMailItem.bodyFormat = olFormatPlain)) Then 'we just understand Plain Mails
        ColorizeMailItem = vbNullString
        Exit Function
    End If

    'rich text it
    MyMailItem.bodyFormat = olFormatRichText
    MyMailItem.Save  'need to save to be able to access rtf via EntryID (.save creates EntryID if not saved before)!

    Dim folder As MAPIFolder
    Set folder = session.GetDefaultFolder(olFolderInbox)

    Dim rtf  As String
    rtf = Space$(99999)  'init rtf to max length of message!

    Dim ret As Integer      '<-- {1}
    ret = ReadRTF(session.CurrentProfileName, MyMailItem.EntryID, folder.StoreID, rtf)
    If (ret = 0) Then
        'ole call success!!!
        rtf = Trim$(rtf)  'kill unnecessary spaces (from rtf var init with Space$(rtf))
        Debug.Print rtf & vbCrLf & "*************************************************************" & vbCrLf

        'we have our own rtf header, remove generated one
        Dim PosHeaderEnd As Long
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

        rtf = Mid$(rtf, PosHeaderEnd + Len(sTestString))

        rtf = "{\rtf1\ansi\ansicpg1252 \deff0{\fonttbl" & vbCrLf & _
                "{\f0\fswiss\fcharset0 Courier New;}}" & vbCrLf & _
                "{\colortbl\red0\green0\blue0;\red106\green44\blue44;\red44\green106\blue44;\red44\green44\blue106;}" & vbCrLf & _
                rtf

        Dim lines() As String
        lines = Split(rtf, vbCrLf)

        Dim i As Long
        For i = LBound(lines) To UBound(lines)
            Dim n As Long
            n = QuoteFixMacro.CalcNesting(lines(i)).level

            Dim resRTF As String
            If (n = 0) Then
                resRTF = resRTF & lines(i) & vbCrLf
            Else
                If (Right$(lines(i), 4) = "\par") Then
                    Dim s As String
                    s = Left$(lines(i), Len(lines(i)) - Len("\par"))
                    resRTF = resRTF & "\cf" & n Mod NUM_RTF_COLORS & " " & s & "\cf0  " & "\par" & vbCrLf
                Else
                    resRTF = resRTF & "\cf" & n Mod NUM_RTF_COLORS & " " & lines(i) & "\cf0  " & vbCrLf
                End If
            End If
        Next
    Else
        Debug.Print "error while reading rtf! " & ret
        ColorizeMailItem = vbNullString
        Exit Function
    End If

    'remove some rtf commands
    resRTF = Replace$(resRTF, "\viewkind4\uc1", vbNullString)
    resRTF = Replace$(resRTF, "\uc1", vbNullString)
    'VERY IMPORTANT, outlook will change the message back to PlainText otherwise!!!
    resRTF = Replace$(resRTF, "\fromtext", vbNullString)
    Debug.Print resRTF

    'write RTF back to form
    ret = WriteRTF(session.CurrentProfileName, MyMailItem.EntryID, folder.StoreID, resRTF)
    If (ret = 0) Then
        Debug.Print "rtf write okay"
    Else
        Debug.Print "rtf write FAILURE"
        ColorizeMailItem = vbNullString
        Exit Function
    End If

    'dereference all objects! otherwise, rtf isn't going to be updated!
    Set folder = Nothing
    'save return value
    ColorizeMailItem = MyMailItem.EntryID
    Set MyMailItem = Nothing
End Function


Public Sub DisplayMailItemByID(ByVal id As String)
    Dim it As MailItem
    Set it = session.GetItemFromID(id, session.GetDefaultFolder(olFolderInbox).StoreID)
    it.Display
    Set it = Nothing
End Sub
