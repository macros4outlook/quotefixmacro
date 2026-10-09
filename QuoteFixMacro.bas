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
'   SaveSetting APPNAME, REG_GROUP_CONFIG, "STRIP_SIGNATURE", "false"
'Finally, or by manually creating entries in this registry hive:
'    HKEY_CURRENT_USER\Software\VB and VBA Program Settings\QuoteFixMacro
Private Const APPNAME As String = "QuoteFixMacro"
Private Const REG_GROUP_CONFIG As String = "Config"
Private Const REG_GROUP_FIRSTNAMES As String = "Firstnames" 'stores replacements for firstnames


'--------------------------------------------------------
'*** Feature QuoteColorizer ***
'--------------------------------------------------------
'Colored mode: the reply is an HTML mail in which each quote level has its own color.
'Before the mail is sent, ThisOutlookSession converts it to plain text (COLORIZER_SEND_AS_PLAIN).
'Without ThisOutlookSession, the mail is sent as HTML mail.
Private Const DEFAULT_USE_COLORIZER As Boolean = False

'How many different colors should be used? (at most the number of QUOTE_COLORS)
'Each author gets a color (an author is known from the condensed header "X wrote on ..."), otherwise each quote level
Private Const DEFAULT_NUM_QUOTE_COLORS As Long = 6

'Send a colored reply as plain text mail?
Private Const DEFAULT_COLORIZER_SEND_AS_PLAIN As Boolean = True

'Recipients who get the colored reply as HTML mail nevertheless: addresses or domains ("@example.org"), separated by ";"
'A mail is sent as HTML mail if all of its recipients are listed
Private Const DEFAULT_COLORIZER_HTML_RECIPIENTS As String = ""


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

'Enable QUOTING_TEMPLATE
Private Const DEFAULT_USE_QUOTING_TEMPLATE As Boolean = False

'If the constant USE_QUOTING_TEMPLATE is set, this template is used instead of the signature
Private Const DEFAULT_QUOTING_TEMPLATE As String = "Hallo %FN,\n\n(Antwort inline)\n\n%Q\n\nMit freundlichen Grüßen\n\n%MN\n\n(Antwort inline - powered by https://macros4outlook.github.io/quotefixmacro/)"

'English quote template
Private Const DEFAULT_QUOTING_TEMPLATE_EN As String = "Dear %FN,\n\n(reply inline)\n\n%Q\n\nCheers,\n\n%MFN\n\n(Reply inline - powered by https://macros4outlook.github.io/quotefixmacro/)"

'If USE_QUOTING_TEMPLATE is set: keep the signature Outlook puts into the reply below the template?
'A colored reply (USE_COLORIZER) to an HTML mail keeps it as HTML, with its pictures
Private Const DEFAULT_KEEP_SIGNATURE As Boolean = False

'--------------------------------------------------------
'*** Configuration of condensing ***
'--------------------------------------------------------

'Condense the headers of the older mails within the quoted text (From, Sent, To, Subject) to one line each?
'In a reply, the text below such a header gets one quote level more
Private Const DEFAULT_CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS As Boolean = True

'Should the header of the mail being replied to also be condensed?
'In case you use a custom header (e.g., "You wrote on %D:" in QUOTING_TEMPLATE), this should be set to False
Private Const DEFAULT_CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER As Boolean = True

'Format of a condensed header: %SN sender, %SE sender's address, %D date (DATE_FORMAT), %TO recipients
Private Const DEFAULT_CONDENSED_HEADER_FORMAT As String = "%SN wrote on %D:"

'----- END OF DEFAULT CONFIGURATION -----------------------------------------------------------------------------------


Private Const OUTLOOK_PLAIN_ORIGINALMESSAGE As String = "-----"
'Private Const OUTLOOK_PLAIN_ORIGINALMESSAGE As String = "-----Ursprüngliche Nachricht-----"
'Private Const OUTLOOK_PLAIN_ORIGINALMESSAGE As String = "-----Original Message-----"
Private Const OUTLOOK_ORIGINALMESSAGE   As String = "> " & OUTLOOK_PLAIN_ORIGINALMESSAGE
Private Const PGP_MARKER                As String = "-----BEGIN PGP"
Private Const OUTLOOK_HEADERFINISH      As String = "> "
Private Const SIGNATURE_SEPARATOR       As String = "> --"

'A line of a plain text mail which is not longer is regarded as wrapped by the sender
Private Const MAX_HARD_WRAP_WIDTH       As Long = 80

'Application.LanguageSettings: the language of the user interface (msoLanguageIDUI)
Private Const LANGUAGE_ID_UI            As Long = 2
'the primary languages of a language id
Private Const LANGUAGE_GERMAN           As Long = 7
Private Const LANGUAGE_ENGLISH          As Long = 9

'Frequent words which are used to detect the language of a mail
'Words existing in both languages (e.g., "in", "was", "will") are left out
Private Const WORDS_GERMAN              As String = " der die das und ist nicht ich wir sie ein eine einen mit auf den dem des zu von es auch noch hallo danke bitte wenn kann haben wird sind dass sich aber oder wie bei nach aus mir dir uns zum zur nur schon viele habe hat bis im "
Private Const WORDS_ENGLISH             As String = " the and is are you we to of for on that this it with have be not hello thanks regards please can would but or as at by from your our i my me if there what which do does they their been has "
'The language is detected using the first words of a mail
Private Const MAX_WORDS_FOR_DETECTION   As Long = 300

Private Const PATTERN_QUOTED_TEXT       As String = "%Q"
Private Const PATTERN_CURSOR_POSITION   As String = "%C"
Private Const PATTERN_SENDER_NAME       As String = "%SN"
Private Const PATTERN_SENDER_EMAIL      As String = "%SE"
Private Const PATTERN_FIRST_NAME        As String = "%FN"
Private Const PATTERN_LAST_NAME         As String = "%LN"
Private Const PATTERN_SENT_DATE         As String = "%D"
Private Const PATTERN_OUTLOOK_HEADER    As String = "%OH"
'recipients of a condensed header
Private Const PATTERN_RECIPIENTS        As String = "%TO"
'the user's own name and first name
Private Const PATTERN_MY_NAME           As String = "%MN"
Private Const PATTERN_MY_FIRST_NAME     As String = "%MFN"

'Labels of the first line of the header of an older mail ("From:"), lower case, several languages
Private Const LABELS_FROM               As String = " from von de da van fra od af "
'Labels of the last line of such a header ("Subject:")
Private Const LABELS_SUBJECT            As String = " subject betreff objet oggetto onderwerp asunto assunto emne temat aihe "
'Labels of the date line of such a header, which some programs put behind the subject
Private Const LABELS_DATE               As String = " sent date datum gesendet envoye inviato enviado verzonden sendt skickat "
'The last word of a one-line attribution ("01.10.2026 16:15 - Firstname Lastname schrieb:")
Private Const WORDS_WROTE               As String = " wrote schrieb schreef skrev scrisse kirjoitti "
'A header has at most this number of lines
Private Const MAX_HEADER_LINES          As Long = 8


'Variables storing the configuration
'They are set in LoadConfiguration()
Private USE_COLORIZER As Boolean
Private NUM_QUOTE_COLORS As Long
Private COLORIZER_SEND_AS_PLAIN As Boolean
Private COLORIZER_HTML_RECIPIENTS As String
Private USE_SOFTWRAP As Boolean
Private SEVENTY_SIX_CHARS As String
Private PIXEL_PER_CHARACTER As Double
Private INCLUDE_QUOTES_TO_LEVEL As Long
Private LINE_WRAP_AFTER As Long
Private DATE_FORMAT As String
Private STRIP_SIGNATURE As Boolean
Private USE_QUOTING_TEMPLATE As Boolean
Private QUOTING_TEMPLATE As String
Private QUOTING_TEMPLATE_EN As String
Private KEEP_SIGNATURE As Boolean
Private CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS As Boolean
Private CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER As Boolean
Private CONDENSED_HEADER_FORMAT As String

'These are fetched from the registry (LoadConfiguration), but not saved by StoreDefaultConfiguration
Private FIRSTNAME_REPLACEMENT__EMAIL() As String
Private FIRSTNAME_REPLACEMENT__FIRSTNAME() As String


'Colors of the colored mode (RGB in hexadecimal, separated by ";"): blue, green, purple, amber, teal, brown
Private Const QUOTE_COLORS As String = "1F6FB2;2E8B57;7B4FA0;B8731B;008B8B;8B5A2B"
'The color of the quoted text written by the user (dark gray)
Private Const OWN_TEXT_COLOR As String = "555555"

'The property marking a colored reply. ThisOutlookSession converts such a mail to plain text before it is sent
Private Const COLORED_MAIL_PROPERTY As String = "QuoteFixMacroColored"


Private Enum ReplyType
    TypeReply = 1
    TypeReplyAll = 2
    TypeForward = 3
End Enum

'Where the text handed to ReFormatText comes from
Public Enum QuoteSource
    'Outlook prefixed the text and thereby wrapped it: the broken wraps are repaired
    SourceOutlookReply = 0
    'A plain text mail prefixed by QuoteText: paragraphs with lines being too long are wrapped anew
    SourcePlainText = 1
    'A mail converted by HtmlToPlainText and prefixed by QuoteText: each line is a paragraph and wrapped on its own
    SourceHtml = 2
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
Private curSource As QuoteSource

'"Fixed Reply" functionality - has to be made available as shortcut in Outlook
Public Sub FixedReply()
    Dim m As Object
    Set m = GetCurrentItem()

    FixMailText m, TypeReply
End Sub

'"Fixed Reply" with colored quotes, whatever USE_COLORIZER says
Public Sub FixedReplyColored()
    Dim m As Object
    Set m = GetCurrentItem()

    FixMailText m, TypeReply, False, True
End Sub

'"Fixed Reply" as plain text, whatever USE_COLORIZER says
Public Sub FixedReplyPlain()
    Dim m As Object
    Set m = GetCurrentItem()

    FixMailText m, TypeReply, False, False
End Sub

'"Fixed Reply All" functionality - has to be made available as shortcut in Outlook
Public Sub FixedReplyAll()
    Dim m As Object
    Set m = GetCurrentItem()

    FixMailText m, TypeReplyAll
End Sub

'"Fixed Reply All" with colored quotes, whatever USE_COLORIZER says
Public Sub FixedReplyAllColored()
    Dim m As Object
    Set m = GetCurrentItem()

    FixMailText m, TypeReplyAll, False, True
End Sub

'"Fixed Reply All" as plain text, whatever USE_COLORIZER says
Public Sub FixedReplyAllPlain()
    Dim m As Object
    Set m = GetCurrentItem()

    FixMailText m, TypeReplyAll, False, False
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

        If count = 0 Then
            'not quoted: the spaces at the beginning are indentation
            res.additionalSpacesCount = i - 1
        Else
            res.additionalSpacesCount = i - lastQuoteSignPos - 2
        End If
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
    SaveSetting APPNAME, REG_GROUP_CONFIG, "NUM_QUOTE_COLORS", DEFAULT_NUM_QUOTE_COLORS
    SaveSetting APPNAME, REG_GROUP_CONFIG, "COLORIZER_SEND_AS_PLAIN", DEFAULT_COLORIZER_SEND_AS_PLAIN
    SaveSetting APPNAME, REG_GROUP_CONFIG, "COLORIZER_HTML_RECIPIENTS", DEFAULT_COLORIZER_HTML_RECIPIENTS
    SaveSetting APPNAME, REG_GROUP_CONFIG, "USE_SOFTWRAP", DEFAULT_USE_SOFTWRAP
    SaveSetting APPNAME, REG_GROUP_CONFIG, "SEVENTY_SIX_CHARS", DEFAULT_SEVENTY_SIX_CHARS
    SaveSetting APPNAME, REG_GROUP_CONFIG, "PIXEL_PER_CHARACTER", DEFAULT_PIXEL_PER_CHARACTER
    SaveSetting APPNAME, REG_GROUP_CONFIG, "INCLUDE_QUOTES_TO_LEVEL", DEFAULT_INCLUDE_QUOTES_TO_LEVEL
    SaveSetting APPNAME, REG_GROUP_CONFIG, "LINE_WRAP_AFTER", DEFAULT_LINE_WRAP_AFTER
    SaveSetting APPNAME, REG_GROUP_CONFIG, "DATE_FORMAT", DEFAULT_DATE_FORMAT
    SaveSetting APPNAME, REG_GROUP_CONFIG, "STRIP_SIGNATURE", DEFAULT_STRIP_SIGNATURE
    SaveSetting APPNAME, REG_GROUP_CONFIG, "USE_QUOTING_TEMPLATE", DEFAULT_USE_QUOTING_TEMPLATE
    SaveSetting APPNAME, REG_GROUP_CONFIG, "QUOTING_TEMPLATE", DEFAULT_QUOTING_TEMPLATE
    SaveSetting APPNAME, REG_GROUP_CONFIG, "QUOTING_TEMPLATE_EN", DEFAULT_QUOTING_TEMPLATE_EN
    SaveSetting APPNAME, REG_GROUP_CONFIG, "KEEP_SIGNATURE", DEFAULT_KEEP_SIGNATURE
    SaveSetting APPNAME, REG_GROUP_CONFIG, "CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS", DEFAULT_CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS
    SaveSetting APPNAME, REG_GROUP_CONFIG, "CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER", DEFAULT_CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER
    SaveSetting APPNAME, REG_GROUP_CONFIG, "CONDENSED_HEADER_FORMAT", DEFAULT_CONDENSED_HEADER_FORMAT
End Sub

'Loads the personal settings from the registry.
Public Sub LoadConfiguration()
    USE_COLORIZER = CBool(GetSetting(APPNAME, REG_GROUP_CONFIG, "USE_COLORIZER", DEFAULT_USE_COLORIZER))
    'NUM_RTF_COLORS is the name of the setting in versions up to 1.9
    NUM_QUOTE_COLORS = Val(GetSetting(APPNAME, REG_GROUP_CONFIG, "NUM_QUOTE_COLORS", GetSetting(APPNAME, REG_GROUP_CONFIG, "NUM_RTF_COLORS", DEFAULT_NUM_QUOTE_COLORS)))
    COLORIZER_SEND_AS_PLAIN = CBool(GetSetting(APPNAME, REG_GROUP_CONFIG, "COLORIZER_SEND_AS_PLAIN", DEFAULT_COLORIZER_SEND_AS_PLAIN))
    COLORIZER_HTML_RECIPIENTS = GetSetting(APPNAME, REG_GROUP_CONFIG, "COLORIZER_HTML_RECIPIENTS", DEFAULT_COLORIZER_HTML_RECIPIENTS)
    USE_SOFTWRAP = CBool(GetSetting(APPNAME, REG_GROUP_CONFIG, "USE_SOFTWRAP", DEFAULT_USE_SOFTWRAP))
    SEVENTY_SIX_CHARS = GetSetting(APPNAME, REG_GROUP_CONFIG, "SEVENTY_SIX_CHARS", DEFAULT_SEVENTY_SIX_CHARS)
    PIXEL_PER_CHARACTER = CDbl(GetSetting(APPNAME, REG_GROUP_CONFIG, "PIXEL_PER_CHARACTER", DEFAULT_PIXEL_PER_CHARACTER))
    INCLUDE_QUOTES_TO_LEVEL = Val(GetSetting(APPNAME, REG_GROUP_CONFIG, "INCLUDE_QUOTES_TO_LEVEL", DEFAULT_INCLUDE_QUOTES_TO_LEVEL))
    LINE_WRAP_AFTER = Val(GetSetting(APPNAME, REG_GROUP_CONFIG, "LINE_WRAP_AFTER", DEFAULT_LINE_WRAP_AFTER))
    DATE_FORMAT = GetSetting(APPNAME, REG_GROUP_CONFIG, "DATE_FORMAT", DEFAULT_DATE_FORMAT)
    STRIP_SIGNATURE = CBool(GetSetting(APPNAME, REG_GROUP_CONFIG, "STRIP_SIGNATURE", DEFAULT_STRIP_SIGNATURE))
    USE_QUOTING_TEMPLATE = CBool(GetSetting(APPNAME, REG_GROUP_CONFIG, "USE_QUOTING_TEMPLATE", DEFAULT_USE_QUOTING_TEMPLATE))
    CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS = CBool(GetSetting(APPNAME, REG_GROUP_CONFIG, "CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS", DEFAULT_CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS))
    CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER = CBool(GetSetting(APPNAME, REG_GROUP_CONFIG, "CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER", DEFAULT_CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER))
    CONDENSED_HEADER_FORMAT = GetSetting(APPNAME, REG_GROUP_CONFIG, "CONDENSED_HEADER_FORMAT", DEFAULT_CONDENSED_HEADER_FORMAT)

    QUOTING_TEMPLATE = GetSetting(APPNAME, REG_GROUP_CONFIG, "QUOTING_TEMPLATE", DEFAULT_QUOTING_TEMPLATE)
    QUOTING_TEMPLATE = Replace$(QUOTING_TEMPLATE, "\n", vbCrLf)

    QUOTING_TEMPLATE_EN = GetSetting(APPNAME, REG_GROUP_CONFIG, "QUOTING_TEMPLATE_EN", DEFAULT_QUOTING_TEMPLATE_EN)
    QUOTING_TEMPLATE_EN = Replace$(QUOTING_TEMPLATE_EN, "\n", vbCrLf)

    KEEP_SIGNATURE = CBool(GetSetting(APPNAME, REG_GROUP_CONFIG, "KEEP_SIGNATURE", DEFAULT_KEEP_SIGNATURE))

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

    If nesting.level = 0 Then
        'not quoted: no space separating the quote characters from the text
        CalcPrefix = res
    Else
        CalcPrefix = res & " "
    End If
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

    If curSource = SourcePlainText Then
        'The line is too long with its prefix, but short enough to be a line wrapped by the sender
        If (Len(curPrefix) + Len(curLine) > LINE_WRAP_AFTER) And (Len(curLine) <= MAX_HARD_WRAP_WIDTH) Then
            curBlockNeedsToBeReFormatted = True
        End If
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

'Description:
'   Formats the date of a quoted Outlook header line (e.g., "Sent: Thursday, April 07, 2011 9:36 AM") using DATE_FORMAT
'   The label and the weekday are optional
'   If the date cannot be parsed, it is returned as found in the email
Private Function FormatHeaderDate(ByVal headerLine As String) As String
    Dim formatted As String
    TryFormatHeaderDate headerLine, formatted
    FormatHeaderDate = formatted
End Function

'Description:
'   The same as FormatHeaderDate, but tells whether a date was found (formatted is the text as found otherwise)
Private Function TryFormatHeaderDate(ByVal headerLine As String, ByRef formatted As String) As Boolean
    Dim sDate As String
    sDate = headerLine

    'strip the label
    Dim posLabelEnd As Long
    posLabelEnd = InStr(sDate, ": ")
    If posLabelEnd > 0 Then
        sDate = Mid$(sDate, posLabelEnd + 2)
    End If

    'strip the weekday: the text before the first comma is a weekday if it does not contain a digit
    '("Thursday, April 07, 2011" has a weekday, "April 7, 2011" has none)
    Dim posFirstComma As Long
    posFirstComma = InStr(sDate, ",")
    If posFirstComma > 0 Then
        If Not (Left$(sDate, posFirstComma - 1) Like "*#*") Then
            sDate = Trim$(Mid$(sDate, posFirstComma + 1))
        End If
    End If

    If IsDate(sDate) Then
        formatted = Format$(CDate(sDate), DATE_FORMAT)
        TryFormatHeaderDate = True
    Else
        'leave sDate as is -> date is output as found in email
        formatted = sDate
        TryFormatHeaderDate = False
    End If
End Function

'Description:
'   Condenses the header of each older mail within the text to one line (CONDENSED_HEADER_FORMAT)
'   A header at the same quote level as the text above it starts an older mail: the text below it gets one level more
'   A header which is deeper than the text above it is quoted already: its condensed line gets one level less
'Notes:
'   * Public to enable testing
Public Function CondenseHeaders(ByVal text As String) As String
    Dim rows() As String
    rows = Split(text, vbCrLf)

    'the quote levels at which older mails start: a line at such a level or deeper gets one level more per entry
    Dim startLevels() As Long
    ReDim startLevels(1 To 1)
    Dim startCount As Long
    startCount = 0

    'the quote level (including the additional levels) of the last line with text
    Dim previousTextLevel As Long
    previousTextLevel = 0

    Dim res As String
    Dim i As Long
    i = LBound(rows)
    Do While i <= UBound(rows)
        Dim level As Long
        level = CalcNesting(rows(i)).level
        Dim extraLevels As Long
        extraLevels = CountStartLevelsUpTo(startLevels, startCount, level)

        Dim endIndex As Long
        Dim condensedHeader As String
        Dim isHeader As Boolean
        isHeader = TryCondenseHeader(rows, i, endIndex, condensedHeader)
        If Not isHeader Then
            'a one-line attribution above text which is not quoted deeper: the older mail starts here
            If TryCondenseAttribution(rows(i), condensedHeader) Then
                If NextTextLevel(rows, i) <= level Then
                    isHeader = True
                    endIndex = i
                End If
            End If
        End If

        If isHeader Then
            Dim condensedLevel As Long
            If level + extraLevels > previousTextLevel Then
                'the older mail is quoted deeper already
                condensedLevel = level + extraLevels - 1
            Else
                'the older mail starts here
                condensedLevel = level + extraLevels
                startCount = startCount + 1
                ReDim Preserve startLevels(1 To startCount)
                startLevels(startCount) = level
            End If

            If condensedLevel > 0 Then
                res = res & String$(condensedLevel, ">") & " "
            End If
            res = res & condensedHeader & vbCrLf

            'the text of the older mail follows directly
            i = endIndex + 1
            Do While i <= UBound(rows)
                If Len(StripLine(rows(i))) > 0 Then Exit Do
                i = i + 1
            Loop
        Else
            If extraLevels > 0 Then
                res = res & String$(extraLevels, ">") & " "
            End If
            res = res & rows(i) & vbCrLf

            If Len(StripLine(rows(i))) > 0 Then
                previousTextLevel = level + extraLevels
            End If
            i = i + 1
        End If
    Loop

    If Len(res) > 0 Then
        res = Left$(res, Len(res) - 2)
    End If
    CondenseHeaders = res
End Function

Private Function CountStartLevelsUpTo(ByRef startLevels() As Long, ByVal startCount As Long, ByVal level As Long) As Long
    Dim i As Long
    For i = 1 To startCount
        If startLevels(i) <= level Then
            CountStartLevelsUpTo = CountStartLevelsUpTo + 1
        End If
    Next
End Function

'Description:
'   Checks whether a header of an older mail starts at rows(start): an optional marker line ("-----Original Message-----"),
'   a "From:" line, and further "Label: value" lines up to the "Subject:" line or an empty line, all at the same quote level.
'   One of the lines has to contain a date.
'Returns:
'   True if a header was found. Then, endIndex is its last line, and condensedHeader the line replacing it
Private Function TryCondenseHeader(ByRef rows() As String, ByVal start As Long, ByRef endIndex As Long, ByRef condensedHeader As String) As Boolean
    Dim level As Long
    level = CalcNesting(rows(start)).level

    Dim i As Long
    i = start
    Dim line As String
    line = HeaderLine(rows(i))

    Dim hasMarker As Boolean
    hasMarker = IsHeaderMarker(line)
    If hasMarker Then
        'the "From:" line follows the marker, possibly behind empty lines
        Do
            i = i + 1
            If i > UBound(rows) Then Exit Function
            If CalcNesting(rows(i)).level <> level Then Exit Function
            line = HeaderLine(rows(i))
        Loop While Len(line) = 0
    End If

    If Not IsListedLabel(GetLabel(line), LABELS_FROM) Then Exit Function

    'the lines of the header: "Label: value"
    Dim values() As String
    ReDim values(1 To MAX_HEADER_LINES)
    Dim valueCount As Long
    valueCount = 0
    Dim dateIndex As Long
    dateIndex = 0

    Do While i <= UBound(rows)
        If CalcNesting(rows(i)).level <> level Then Exit Do
        line = HeaderLine(rows(i))
        If Len(line) = 0 Then Exit Do

        Dim label As String
        label = GetLabel(line)
        If Len(label) > 0 Then
            If valueCount = MAX_HEADER_LINES Then Exit Do
            valueCount = valueCount + 1
            values(valueCount) = Trim$(Mid$(line, Len(label) + 2))
            If dateIndex = 0 And valueCount > 1 Then
                If IsDateLike(values(valueCount)) Then dateIndex = valueCount
            End If
            If IsListedLabel(label, LABELS_SUBJECT) Then
                'the subject is the last line, unless the date follows it
                i = i + 1
                If i <= UBound(rows) Then
                    If CalcNesting(rows(i)).level = level And dateIndex = 0 And valueCount < MAX_HEADER_LINES Then
                        line = HeaderLine(rows(i))
                        If IsListedLabel(GetLabel(line), LABELS_DATE) Then
                            valueCount = valueCount + 1
                            values(valueCount) = Trim$(Mid$(line, Len(GetLabel(line)) + 2))
                            If IsDateLike(values(valueCount)) Then dateIndex = valueCount
                            i = i + 1
                        End If
                    End If
                End If
                Exit Do
            End If
        ElseIf valueCount > 0 And Len(values(valueCount)) >= 50 Then
            'a long value was wrapped by a mail program
            values(valueCount) = values(valueCount) & " " & line
        Else
            Exit Do
        End If
        i = i + 1
    Loop

    'a header needs sender, date, and at least one more line
    If valueCount < 3 Or dateIndex = 0 Then Exit Function

    Dim senderRaw As String
    Dim senderEmail As String
    SplitSender values(1), senderRaw, senderEmail
    Dim senderName As String
    Dim firstName As String
    Dim lastName As String
    getNamesOutOfString senderRaw, senderName, firstName, lastName, senderEmail

    'the recipients follow the date
    Dim recipients As String
    If dateIndex < valueCount Then
        recipients = values(dateIndex + 1)
    End If

    condensedHeader = CONDENSED_HEADER_FORMAT
    condensedHeader = Replace$(condensedHeader, PATTERN_SENDER_NAME, senderName)
    condensedHeader = Replace$(condensedHeader, PATTERN_SENT_DATE, FormatHeaderDate(values(dateIndex)))
    condensedHeader = Replace$(condensedHeader, PATTERN_SENDER_EMAIL, senderEmail)
    condensedHeader = Replace$(condensedHeader, PATTERN_RECIPIENTS, recipients)

    endIndex = i - 1
    TryCondenseHeader = True
End Function

'Description:
'   Returns the text of a row of a header without quote prefix and emphasis markers:
'   HTML mails show the labels (or the sender) in bold, "*From:* Art Ross" is read as "From: Art Ross"
Private Function HeaderLine(ByVal row As String) As String
    Dim line As String
    line = StripLine(row)
    If (InStr(line, "*") = 0) And (InStr(line, "_") = 0) Then
        HeaderLine = line
        Exit Function
    End If

    Dim rows(0 To 0) As String
    rows(0) = line
    Dim boldRoles() As String
    boldRoles = MatchEmphasisMarkers(rows, "*")
    Dim underlineRoles() As String
    underlineRoles = MatchEmphasisMarkers(rows, "_")

    Dim res As String
    Dim pos As Long
    For pos = 1 To Len(line)
        If (Mid$(boldRoles(0), pos, 1) = " ") And (Mid$(underlineRoles(0), pos, 1) = " ") Then
            res = res & Mid$(line, pos, 1)
        End If
    Next
    HeaderLine = res
End Function

'Description:
'   Checks whether the row is the one-line attribution which ticket systems write above a quoted mail
'   without prefixing it: "01.10.2026 16:15 - Firstname Lastname schrieb:" (or "wrote:")
'Returns:
'   True if it is one. Then, condensedHeader is the line replacing it
Private Function TryCondenseAttribution(ByVal row As String, ByRef condensedHeader As String) As Boolean
    Dim line As String
    line = HeaderLine(row)
    If Right$(line, 1) <> ":" Then Exit Function

    Dim posSeparator As Long
    posSeparator = InStr(line, " - ")
    If posSeparator = 0 Then Exit Function

    Dim sDate As String
    sDate = Left$(line, posSeparator - 1)
    If Not IsDateLike(sDate) Then Exit Function

    'the rest: "Firstname Lastname schrieb"
    Dim rest As String
    rest = Mid$(line, posSeparator + 3, Len(line) - posSeparator - 3)
    Dim posLastWord As Long
    posLastWord = InStrRev(rest, " ")
    If posLastWord = 0 Then Exit Function
    If Not IsListedLabel(Mid$(rest, posLastWord + 1), WORDS_WROTE) Then Exit Function

    Dim senderRaw As String
    Dim senderEmail As String
    SplitSender Left$(rest, posLastWord - 1), senderRaw, senderEmail
    If Len(senderRaw) = 0 Then Exit Function
    Dim senderName As String
    Dim firstName As String
    Dim lastName As String
    getNamesOutOfString senderRaw, senderName, firstName, lastName, senderEmail

    condensedHeader = CONDENSED_HEADER_FORMAT
    condensedHeader = Replace$(condensedHeader, PATTERN_SENDER_NAME, senderName)
    condensedHeader = Replace$(condensedHeader, PATTERN_SENT_DATE, FormatHeaderDate(sDate))
    condensedHeader = Replace$(condensedHeader, PATTERN_SENDER_EMAIL, senderEmail)
    condensedHeader = Replace$(condensedHeader, PATTERN_RECIPIENTS, vbNullString)

    TryCondenseAttribution = True
End Function

'Returns the quote level of the next row with text behind start, or the level of start if there is none
Private Function NextTextLevel(ByRef rows() As String, ByVal start As Long) As Long
    Dim i As Long
    For i = start + 1 To UBound(rows)
        If Len(StripLine(rows(i))) > 0 Then
            NextTextLevel = CalcNesting(rows(i)).level
            Exit Function
        End If
    Next
    NextTextLevel = CalcNesting(rows(start)).level
End Function

'"-----Original Message-----" (any language), "---- Forwarded message ----", or the line Outlook on the web puts above a header
Private Function IsHeaderMarker(ByVal line As String) As Boolean
    If Left$(line, Len(PGP_MARKER)) = PGP_MARKER Then Exit Function
    IsHeaderMarker = (Left$(line, 4) = "----") Or (Left$(line, 10) = "__________")
End Function

'Description:
'   Returns the label of a "Label: value" line (without the colon), vbNullString if the line has none
'   A label consists of 2 to 20 letters or spaces and is followed by ": " or ":" at the end of the line
Private Function GetLabel(ByVal line As String) As String
    Dim posColon As Long
    posColon = InStr(line, ":")
    If posColon < 3 Or posColon > 21 Then Exit Function
    If posColon < Len(line) Then
        If Mid$(line, posColon + 1, 1) <> " " Then Exit Function
    End If

    Dim i As Long
    For i = 1 To posColon - 1
        Dim c As String
        c = Mid$(line, i, 1)
        If Not (c Like "[A-Za-z ]") And AscW(c) < 128 And AscW(c) >= 0 Then Exit Function
    Next

    GetLabel = Left$(line, posColon - 1)
End Function

Private Function IsListedLabel(ByVal label As String, ByVal labels As String) As Boolean
    If Len(label) = 0 Then Exit Function
    IsListedLabel = (InStr(labels, " " & LCase$(label) & " ") > 0)
End Function

'True if the text contains a date, or at least a number of four digits (e.g., a year within a date which cannot be parsed)
Private Function IsDateLike(ByVal value As String) As Boolean
    Dim formatted As String
    If TryFormatHeaderDate(value, formatted) Then
        IsDateLike = True
        Exit Function
    End If
    IsDateLike = (value Like "*#???#*") And (value Like "*####*")
End Function

'Description:
'   Splits "Name <address>" or "Name [mailto:address]" into name and address
Private Sub SplitSender(ByVal fromValue As String, ByRef senderRaw As String, ByRef senderEmail As String)
    Dim posStart As Long
    Dim addressOffset As Long
    posStart = InStr(fromValue, "<")
    addressOffset = 1
    If posStart = 0 Then
        posStart = InStr(fromValue, "[mailto:")
        addressOffset = 8
    End If

    If posStart = 0 Then
        senderRaw = Trim$(fromValue)
        senderEmail = vbNullString
    Else
        senderRaw = Trim$(Left$(fromValue, posStart - 1))
        senderEmail = Mid$(fromValue, posStart + addressOffset)
        Dim posEnd As Long
        posEnd = InStr(senderEmail, ">")
        If posEnd = 0 Then posEnd = InStr(senderEmail, "]")
        If posEnd > 0 Then senderEmail = Left$(senderEmail, posEnd - 1)
        senderEmail = Trim$(senderEmail)
    End If

    If Len(senderRaw) = 0 Then
        senderRaw = senderEmail
    End If
End Sub

'Description:
'   Prefixes each line of the text with "> ", the way Outlook does it for a plain text mail
'   Empty lines at the end are dropped
'Notes:
'   * Public to enable testing
Public Function QuoteText(ByVal text As String) As String
    Dim lines() As String
    lines = Split(Replace$(text, vbCrLf, vbLf), vbLf)

    Dim lastLine As Long
    lastLine = UBound(lines)
    Do While lastLine >= LBound(lines)
        If Len(Trim$(lines(lastLine))) > 0 Then Exit Do
        lastLine = lastLine - 1
    Loop

    Dim i As Long
    For i = LBound(lines) To lastLine
        If i > LBound(lines) Then
            QuoteText = QuoteText & vbCrLf
        End If
        QuoteText = QuoteText & "> " & lines(i)
    Next
End Function

'Description:
'   Detects the language of a mail by counting frequent words
'   Only the newest part of the mail is regarded: the text above the first quote or header of an older mail
'   Returns the primary language id (7 for German, 9 for English), or 0 if there is no clear result
'Notes:
'   * Public to enable testing
Public Function DetectLanguage(ByVal text As String) As Long
    Dim rows() As String
    rows = Split(Replace$(text, vbCrLf, vbLf), vbLf)

    Dim countGerman As Long
    Dim countEnglish As Long
    Dim countWords As Long

    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        Dim row As String
        row = LCase$(Trim$(rows(i)))
        If IsStartOfOlderMail(row) Or countWords >= MAX_WORDS_FOR_DETECTION Then Exit For

        'everything but a letter separates words
        Dim word As String
        word = vbNullString
        Dim j As Long
        For j = 1 To Len(row) + 1
            Dim c As String
            c = Mid$(row, j, 1)
            If IsLetter(c) Then
                word = word & c
            ElseIf Len(word) > 0 Then
                If InStr(WORDS_GERMAN, " " & word & " ") > 0 Then
                    countGerman = countGerman + 1
                End If
                If InStr(WORDS_ENGLISH, " " & word & " ") > 0 Then
                    countEnglish = countEnglish + 1
                End If
                countWords = countWords + 1
                word = vbNullString
            End If
        Next
    Next

    'a clear result: at least two words of a language, and at least twice as many as of the other language
    If countGerman >= 2 And countGerman >= 2 * countEnglish Then
        DetectLanguage = LANGUAGE_GERMAN
    ElseIf countEnglish >= 2 And countEnglish >= 2 * countGerman Then
        DetectLanguage = LANGUAGE_ENGLISH
    End If
End Function

'row has to be trimmed and in lower case
Private Function IsStartOfOlderMail(ByVal row As String) As Boolean
    IsStartOfOlderMail = (Left$(row, 1) = ">") Or (Left$(row, 5) = "-----") Or (Left$(row, 5) = "_____") _
        Or (Left$(row, 6) = "from: ") Or (Left$(row, 5) = "von: ")
End Function

'c has to be in lower case
Private Function IsLetter(ByVal c As String) As Boolean
    If Len(c) = 0 Then Exit Function

    'letters outside of a-z (e.g., umlauts) have a code above 127, or a negative one
    IsLetter = (c Like "[a-z]") Or (AscW(c) > 127) Or (AscW(c) < 0)
End Function

'Description:
'   Builds the header of the original mail the way Outlook puts it above the original text of a plain text mail
'   The header is German if languageId is German, English otherwise
'   Each line ends with a line break
'Notes:
'   * Public to enable testing
Public Function BuildOutlookHeader(ByVal languageId As Long, ByVal senderName As String, ByVal senderEmail As String, ByVal sentDate As String, ByVal recipientsTo As String, ByVal recipientsCc As String, ByVal mailSubject As String) As String
    'the lower ten bits of a language id are the primary language
    Dim isGerman As Boolean
    isGerman = ((languageId And 1023) = LANGUAGE_GERMAN)

    Dim fromValue As String
    fromValue = senderName
    If Len(senderEmail) > 0 Then
        fromValue = fromValue & " [mailto:" & senderEmail & "]"
    End If

    Dim header As String
    If isGerman Then
        header = "-----Urspr" & ChrW$(252) & "ngliche Nachricht-----" & vbCrLf
        header = header & "Von: " & fromValue & vbCrLf
        header = header & "Gesendet: " & sentDate & vbCrLf
        header = header & "An: " & recipientsTo & vbCrLf
    Else
        header = "-----Original Message-----" & vbCrLf
        header = header & "From: " & fromValue & vbCrLf
        header = header & "Sent: " & sentDate & vbCrLf
        header = header & "To: " & recipientsTo & vbCrLf
    End If

    If Len(recipientsCc) > 0 Then
        header = header & "Cc: " & recipientsCc & vbCrLf
    End If

    If isGerman Then
        header = header & "Betreff: " & mailSubject & vbCrLf
    Else
        header = header & "Subject: " & mailSubject & vbCrLf
    End If

    BuildOutlookHeader = header
End Function

Private Function FirstWord(ByVal text As String) As String
    Dim posSpace As Long
    posSpace = InStr(text, " ")
    If posSpace = 0 Then
        FirstWord = text
    Else
        FirstWord = Left$(text, posSpace - 1)
    End If
End Function

'Description:
'   Wraps each line being longer than LINE_WRAP_AFTER
'   Each line is wrapped on its own (it is not joined with the next line), the quote prefix is repeated
'   Words being too long (e.g., links) are not broken
Private Function WrapLongLines(ByVal text As String) As String
    Dim rows() As String
    rows = Split(text, vbCrLf)

    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        If i > LBound(rows) Then
            WrapLongLines = WrapLongLines & vbCrLf
        End If

        If Len(rows(i)) <= LINE_WRAP_AFTER Then
            WrapLongLines = WrapLongLines & rows(i)
        Else
            Dim nesting As NestingType
            nesting = CalcNesting(rows(i))

            Dim prefix As String
            If nesting.level = 0 Then
                prefix = vbNullString
            Else
                prefix = CalcPrefix(nesting)
            End If

            Dim maxLength As Long
            maxLength = LINE_WRAP_AFTER - Len(prefix)
            If maxLength < 20 Then
                'very deep nesting: keep some room for the text
                maxLength = 20
            End If

            Dim remaining As String
            remaining = StripLine(rows(i))
            Do While Len(remaining) > maxLength
                'wrap at the last space which still fits
                Dim breakPos As Long
                breakPos = InStrRev(remaining, " ", maxLength + 1)
                If breakPos = 0 Then
                    'the first word is too long: keep it complete
                    breakPos = InStr(remaining, " ")
                    If breakPos = 0 Then Exit Do
                End If

                WrapLongLines = WrapLongLines & prefix & RTrim$(Left$(remaining, breakPos - 1)) & vbCrLf
                remaining = LTrim$(Mid$(remaining, breakPos + 1))
            Loop
            WrapLongLines = WrapLongLines & prefix & remaining
        End If
    Next
End Function

'Reformat text to correct broken wrap inserted by Outlook.
'Needs to be public so the test cases can run this function.
'
'textSource tells where the text comes from, see QuoteSource
Public Function ReFormatText(ByVal text As String, Optional ByVal textSource As QuoteSource = SourceOutlookReply) As String
    'Reset (partially global) variables
    curSource = textSource

    If CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS And (textSource <> SourceOutlookReply) Then
        text = CondenseHeaders(text)
    End If
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

                'Only Outlook breaks lines of the first level (when it prefixes them)
                If (curSource = SourceOutlookReply) And (curNesting.level = 1) And (i < UBound(rows)) Then
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
                'The nesting of text converted from HTML is given by the HTML: there are no broken wraps
                Dim isBrokenWrap As Boolean
                isBrokenWrap = (curSource <> SourceHtml) And (nextNesting.total = lastNesting.total)
                If isBrokenWrap And (curSource = SourcePlainText) And (Len(curLine) > 0) Then
                    'Outlook only broke the line above if the first word of this line did not fit into it.
                    'Otherwise, this line is an answer between two quotes
                    isBrokenWrap = (Len(rows(i - 1)) > LINE_WRAP_AFTER - Len(FirstWord(curLine)) - 10) '10: the same rough heuristics as above
                End If

                If isBrokenWrap Then
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
                AppendCurLine curLine
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

            'Text not prefixed by Outlook has its headers condensed by CondenseHeaders already
            If CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS And (curSource = SourceOutlookReply) Then
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
                            sName = vbNullString
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
                    Dim sDate As String
                    sDate = FormatHeaderDate(StripLine(rows(i)))

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

    If curSource <> SourceOutlookReply Then
        'Outlook did not wrap the text
        result = WrapLongLines(result)
    End If

    ReFormatText = result
End Function

' @param UseEnglishTemplate In case USE_QUOTING_TEMPLATE is True, should the default or the English template be used?
' @param Colorize If given, overrides USE_COLORIZER for this reply
Private Sub FixMailText(ByVal SelectedObject As Object, ByRef MailMode As ReplyType, Optional ByVal UseEnglishTemplate As Boolean = False, Optional ByVal Colorize As Variant)
    LoadConfiguration
    If Not IsMissing(Colorize) Then
        USE_COLORIZER = CBool(Colorize)
    End If

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

    Dim originalIsPlain As Boolean
    originalIsPlain = (bodyFormat = olFormatPlain)

    'Forwarding is left as it was before replies to HTML mails were supported:
    'A mail which is not a plain text mail is forwarded by Outlook itself,
    'the text of a plain text mail is prefixed by Outlook
    Dim isForward As Boolean
    isForward = (MailMode = TypeForward)
    If isForward And Not originalIsPlain Then
        Dim ForwardObj As Object
        If isMail Then
            Set ForwardObj = OriginalMail.Forward
        Else
            Set ForwardObj = OriginalMeeting.Forward
        End If
        ForwardObj.Display
        Exit Sub
    End If

    'The text of the original mail: it decides the language, and it is quoted in replies
    Dim originalText As String
    If isMail Then
        originalText = getOriginalText(OriginalMail, originalIsPlain)
    Else
        'MeetingItem does not offer HTMLBody
        originalText = OriginalMeeting.Body
    End If

    'Reply: Outlook creates the reply without the original text, header and text of the original are added below.
    'The original mail is never modified. The reply is always a plain text mail.
    Dim NewReplyStyle As OlActionReplyStyle
    If isForward Then
        NewReplyStyle = olReplyTickOriginalText
    Else
        NewReplyStyle = olOmitOriginalText
    End If

    '''create reply --> outlook style!
    ''Actions(1) = Actions("Reply")' or 'Actions("Antworten")' respectively, etc.
    If isMail Then
        With OriginalMail.Actions(MailMode)
            Dim OriginalReplyStyle As OlActionReplyStyle
            OriginalReplyStyle = .ReplyStyle
            .ReplyStyle = NewReplyStyle

            Dim NewMail As MailItem
            Set NewMail = .Execute

            .ReplyStyle = OriginalReplyStyle
        End With
    Else
        With OriginalMeeting.Actions(MailMode)
            OriginalReplyStyle = .ReplyStyle
            .ReplyStyle = NewReplyStyle

            Set NewMail = .Execute

            .ReplyStyle = OriginalReplyStyle
        End With
    End If

    'if the mail is marked as a possible phishing mail, a warning will be shown and
    'the reply methods will return null (forward method is ok)
    If NewMail Is Nothing Then Exit Sub

    'A colored reply keeps the HTML of Outlook's signature, with its pictures (KEEP_SIGNATURE)
    Dim signatureHtml As String
    If USE_QUOTING_TEMPLATE And KEEP_SIGNATURE And USE_COLORIZER Then
        If NewMail.bodyFormat = olFormatHTML Then
            signatureHtml = NewMail.HTMLBody
        End If
    End If

    If Not originalIsPlain And Len(signatureHtml) = 0 Then
        NewMail.bodyFormat = olFormatPlain
    End If

    'put the whole mail as composed by Outlook into an array
    Dim BodyLines() As String
    BodyLines = Split(NewMail.Body, vbCrLf)

    'lineCounter is used to provide information about how many lines we already parsed.
    'This variable is always passed to the various parser functions by reference to get
    'back the new value.
    Dim lineCounter As Long

    Dim MySignature As String
    Dim textSource As QuoteSource
    If isForward Then
        ' A new mail starts with signature -if- set, try to parse until we find the the
        ' original message separator - might loop until the end of the whole message since
        ' this depends on the International Option settings (english), even worse it might
        ' find some separator in-between and mess up the whole reply, so check the nesting too.
        '
        ' We need to call getSignature in all cases as it sets "lineCounter" as side effect
        MySignature = getSignature(BodyLines, lineCounter)
        ' lineCounter now indicates the line after the signature

        textSource = SourceOutlookReply
    Else
        'The reply consists of the signature only
        MySignature = getSignatureOfEmptyReply(BodyLines)

        'Header and text of the original are put into BodyLines the way Outlook does it when it prefixes a plain text mail
        If isMail Then
            BodyLines = Split(getQuotedOriginalOfMail(OriginalMail, originalText), vbCrLf)
        Else
            BodyLines = Split(getQuotedOriginalOfMeeting(OriginalMeeting, originalText), vbCrLf)
        End If
        lineCounter = 0

        If originalIsPlain Then
            textSource = SourcePlainText
        Else
            textSource = SourceHtml
        End If
    End If

    Dim outlookSignature As String
    If USE_QUOTING_TEMPLATE Then
        outlookSignature = MySignature
        'Override MySignature in case the QUOTING_TEMPLATE should be used
        'lineCounter is still valid, because lineCounter is based on the current message whereas QUOTING_TEMPLATE is a general setting
        'The English template is used for a mail written in English, or on request (FixedReplyAllEnglish)
        If UseEnglishTemplate Or (DetectLanguage(originalText) = LANGUAGE_ENGLISH) Then
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

    If InStr(MySignature, PATTERN_MY_NAME) > 0 Then
        Dim ownName As String
        Dim ownFirstName As String
        getOwnNames ownName, ownFirstName
        'the first name first: %MFN starts with %MN
        MySignature = Replace$(MySignature, PATTERN_MY_FIRST_NAME, ownFirstName)
        MySignature = Replace$(MySignature, PATTERN_MY_NAME, ownName)
    End If

    Dim OutlookHeader As String
    If CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER Then
        OutlookHeader = vbNullString
        'The real condensing is made below, where CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS is checked
        'Disabling getOutlookHeader leads to an unmodified lineCounter, which in turn gets the header included in "quotedText"
    Else
        OutlookHeader = getOutlookHeader(BodyLines, lineCounter, MailMode)
    End If

    Dim quotedText As String
    quotedText = getQuotedText(BodyLines, lineCounter, textSource)

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

    'Outlook's signature below the template; the HTML one is added below
    If USE_QUOTING_TEMPLATE And KEEP_SIGNATURE And Len(signatureHtml) = 0 Then
        Do While Right$(outlookSignature, 2) = vbCrLf
            outlookSignature = Left$(outlookSignature, Len(outlookSignature) - 2)
        Loop
        If Len(outlookSignature) > 0 Then
            MySignature = MySignature & vbCrLf & vbCrLf & outlookSignature
        End If
    End If

    MySignature = cleanUpDoubleLines(MySignature)

    If USE_COLORIZER Then
        NewMail.bodyFormat = olFormatHTML
        If Len(signatureHtml) > 0 Then
            NewMail.HTMLBody = InsertColoredHtml(TextToColoredHtml(MySignature, getOwnName()), signatureHtml)
        Else
            NewMail.HTMLBody = TextToColoredHtml(MySignature, getOwnName())
        End If
        'the mark lets ThisOutlookSession convert the mail to plain text before it is sent
        NewMail.UserProperties.Add(COLORED_MAIL_PROPERTY, olText, False).Value = "yes"
    Else
        NewMail.Body = MySignature
    End If

    'Display window
    NewMail.Display

    'jump to the right place
    If downCount > 0 Then
        MoveCursorDown NewMail, downCount
    End If

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

'Description:
'   Returns the signature of a reply Outlook created without the original text
Private Function getSignatureOfEmptyReply(ByRef BodyLines() As String) As String
    'drop the empty lines at the beginning
    Dim i As Long
    i = LBound(BodyLines)
    Do While i <= UBound(BodyLines)
        If Len(Trim$(Replace$(BodyLines(i), ChrW$(160), " "))) > 0 Then Exit Do
        i = i + 1
    Loop

    For i = i To UBound(BodyLines)
        getSignatureOfEmptyReply = getSignatureOfEmptyReply & BodyLines(i) & vbCrLf
    Next
End Function

'Description:
'   Returns the text of a mail: the plain text, or the text made of the HTML
Private Function getOriginalText(ByVal item As MailItem, ByVal isPlainText As Boolean) As String
    If isPlainText Then
        getOriginalText = item.Body
    Else
        'Outlook offers HTML for Rich Text mails, too
        getOriginalText = HtmlToPlainText(item.HTMLBody)
    End If
End Function

'Description:
'   Returns header and text of a mail the way Outlook puts them into the reply to a plain text mail:
'   each line is prefixed, header and text are separated by an empty line
Private Function getQuotedOriginalOfMail(ByVal item As MailItem, ByVal originalText As String) As String
    Dim senderEmail As String
    If item.senderEmailType = "SMTP" Then
        senderEmail = item.senderEmailAddress
    End If

    Dim header As String
    header = BuildOutlookHeader(getHeaderLanguageId(originalText), item.senderName, senderEmail, Format$(item.SentOn, DATE_FORMAT), item.To, item.CC, item.Subject)

    getQuotedOriginalOfMail = QuoteText(header & vbCrLf & originalText)
End Function

'Code duplication of getQuotedOriginalOfMail, because there is no common ancestor of MailItem and MeetingItem
Private Function getQuotedOriginalOfMeeting(ByVal item As MeetingItem, ByVal originalText As String) As String
    Dim senderEmail As String
    If item.senderEmailType = "SMTP" Then
        senderEmail = item.senderEmailAddress
    End If

    'MeetingItem does not offer "To"
    Dim recipientNames As String
    Dim curRecipient As Recipient
    For Each curRecipient In item.Recipients
        If Len(recipientNames) > 0 Then
            recipientNames = recipientNames & "; "
        End If
        recipientNames = recipientNames & curRecipient.Name
    Next

    Dim header As String
    header = BuildOutlookHeader(getHeaderLanguageId(originalText), item.senderName, senderEmail, Format$(item.SentOn, DATE_FORMAT), recipientNames, vbNullString, item.Subject)

    getQuotedOriginalOfMeeting = QuoteText(header & vbCrLf & originalText)
End Function

'Description:
'   Returns the language for the header of the original mail: the language of the mail.
'   If that cannot be detected, it is the language of Outlook's user interface
Private Function getHeaderLanguageId(ByVal originalText As String) As Long
    getHeaderLanguageId = DetectLanguage(originalText)
    If getHeaderLanguageId = 0 Then
        getHeaderLanguageId = getUiLanguageId()
    End If
End Function

'Description:
'   Returns the name of the user ("Firstname Lastname"), vbNullString if it cannot be determined
Private Function getOwnName() As String
    Dim ownFirstName As String
    getOwnNames getOwnName, ownFirstName
End Function

'Description:
'   Returns name ("Firstname Lastname") and first name of the user, vbNullString if they cannot be determined
'Notes:
'   * Names are returned by reference
Private Sub getOwnNames(ByRef ownName As String, ByRef ownFirstName As String)
    ownName = vbNullString
    ownFirstName = vbNullString

    Dim rawName As String
    On Error Resume Next
    rawName = session.CurrentUser.Name
    On Error GoTo 0
    If Len(rawName) = 0 Then Exit Sub

    Dim lastName As String
    getNamesOutOfString rawName, ownName, ownFirstName, lastName
End Sub

'Description:
'   Returns the language of Outlook's user interface (e.g., 1031 for German, 1033 for English)
'   If it cannot be determined, 0 is returned
Private Function getUiLanguageId() As Long
    On Error Resume Next
    getUiLanguageId = Application.LanguageSettings.LanguageID(LANGUAGE_ID_UI)
    On Error GoTo 0
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


Private Function getQuotedText(ByRef BodyLines() As String, ByRef lineCounter As Long, ByVal textSource As QuoteSource) As String
    ' parse the rest of the message
    For lineCounter = lineCounter To UBound(BodyLines)
        'the separator of a signature is "-- ": the space at its end is ignored
        If STRIP_SIGNATURE And (RTrim$(BodyLines(lineCounter)) = SIGNATURE_SEPARATOR) Then
            'beginning of signature reached
            Exit For
        End If

        getQuotedText = getQuotedText & BodyLines(lineCounter) & vbCrLf
    Next

    getQuotedText = ReFormatText(getQuotedText, textSource)

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


'Description:
'   Moves the cursor of the displayed mail from the beginning down by the given number of lines (paragraphs)
'   The Word editor of the inspector is used. SendKeys is only the fallback: called repeatedly, it may switch off NumLock
Private Sub MoveCursorDown(ByVal mail As MailItem, ByVal lineCount As Long)
    Const wdParagraph As Long = 4
    Const wdStory As Long = 6

    On Error GoTo fallback
    Dim editor As Object
    Set editor = mail.GetInspector.WordEditor
    With editor.Windows(1).Selection
        .HomeKey wdStory
        .MoveDown wdParagraph, lineCount
    End With
    Exit Sub

fallback:
    On Error GoTo 0
    SendKeys "{DOWN " & lineCount & "}", True
    DoEvents
End Sub

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


'Description:
'   Converts the text of the reply into HTML in which each author has its own color (colored mode)
'   The author of a quote level is known from the condensed header above it ("X wrote on ...:"), which is shown
'   as heading in the color of the author. A level without known author gets the color of the level.
'   The text of the user (ownName, "Firstname Lastname") is dark gray.
'   Each line becomes a paragraph without space around it, so that the conversion back to plain text yields the lines again.
'   The color is set on a span within the paragraph, not on the paragraph: this way, a new paragraph
'   created by pressing Enter at the end of a quoted line gets the default color (black) for the answer.
'Notes:
'   * Public to enable testing
Public Function TextToColoredHtml(ByVal text As String, ByVal ownName As String) As String
    Dim html As String
    html = "<html><body><div style=""font-family:Consolas,'Courier New',monospace;font-size:10pt"">" & vbCrLf

    Dim headerPattern As String
    headerPattern = CondensedHeaderPattern()

    'the authors of the quote levels, as far as known from the condensed headers
    Dim authorOfLevel(1 To 50) As String
    'the authors in the order of their first appearance, separated by ";" (each one gets the next color)
    Dim authors As String
    authors = ";"

    Dim rows() As String
    rows = Split(Replace$(text, vbCrLf, vbLf), vbLf)

    'the emphasis markers (*bold*, _underlined_), matched over the lines
    Dim boldRoles() As String
    boldRoles = MatchEmphasisMarkers(rows, "*")
    Dim underlineRoles() As String
    underlineRoles = MatchEmphasisMarkers(rows, "_")
    Dim isBold As Boolean
    Dim isUnderlined As Boolean

    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        Dim level As Long
        level = CalcNesting(rows(i)).level

        Dim line As String
        line = StripLine(rows(i))

        'the style of the span holding the text of the line (empty: no span)
        Dim spanStyle As String
        spanStyle = vbNullString

        If (Len(line) > 0) And (line Like headerPattern) Then
            'the header of an older mail: the lines below it (one level deeper) are written by its author
            Dim author As String
            author = HeaderAuthor(line)
            If level + 1 <= UBound(authorOfLevel) Then
                authorOfLevel(level + 1) = author
                Dim deeperLevel As Long
                For deeperLevel = level + 2 To UBound(authorOfLevel)
                    authorOfLevel(deeperLevel) = vbNullString
                Next
            End If
            spanStyle = "color:#" & ColorOfAuthor(author, ownName, authors) & ";font-weight:bold"
        ElseIf level > 0 Then
            If HasKnownAuthor(authorOfLevel, level) Then
                spanStyle = "color:#" & ColorOfAuthor(authorOfLevel(level), ownName, authors)
            Else
                spanStyle = "color:#" & QuoteColor(level)
            End If
        End If

        Dim content As String
        content = EmphasizedHtml(rows(i), boldRoles(i), underlineRoles(i), isBold, isUnderlined)
        If Len(content) = 0 Then
            'an empty paragraph would be dropped
            content = "&nbsp;"
        ElseIf Len(spanStyle) > 0 Then
            content = "<span style=""" & spanStyle & """>" & content & "</span>"
        End If
        If Right$(content, 7) = "</span>" Then
            'Word gives the cursor at the end of the line the format of the character before it, and the text typed after Enter gets it, too.
            'An invisible character without format behind the span keeps that text black.
            content = content & "&#8203;"
        End If

        html = html & "<p style=""margin:0"">" & content & "</p>" & vbCrLf
    Next

    TextToColoredHtml = html & "</div></body></html>"
End Function

'Description:
'   Finds the pairs of emphasis markers (e.g., *bold*) in the lines
'   A pair may span several lines of the same quote level, but no empty line
'Returns:
'   For each line a string as long as the line: "o" at an opening marker, "c" at a closing marker, " " elsewhere
Private Function MatchEmphasisMarkers(ByRef rows() As String, ByVal marker As String) As String()
    If UBound(rows) < LBound(rows) Then
        MatchEmphasisMarkers = rows
        Exit Function
    End If

    Dim roles() As String
    ReDim roles(LBound(rows) To UBound(rows))

    Dim isOpen As Boolean
    Dim openRow As Long
    Dim openPos As Long
    Dim openLevel As Long

    Dim i As Long
    For i = LBound(rows) To UBound(rows)
        roles(i) = Space$(Len(rows(i)))

        Dim level As Long
        level = CalcNesting(rows(i)).level
        If (Len(StripLine(rows(i))) = 0) Or (level <> openLevel) Then
            isOpen = False
        End If

        Dim pos As Long
        For pos = 1 To Len(rows(i))
            If Mid$(rows(i), pos, 1) = marker Then
                If isOpen Then
                    If IsClosingMarker(rows(i), pos) Then
                        Mid$(roles(openRow), openPos, 1) = "o"
                        Mid$(roles(i), pos, 1) = "c"
                        isOpen = False
                    End If
                ElseIf IsOpeningMarker(rows(i), pos) Then
                    isOpen = True
                    openRow = i
                    openPos = pos
                    openLevel = level
                End If
            End If
        Next
    Next

    MatchEmphasisMarkers = roles
End Function

'An opening marker is followed by text and stands at the beginning of a word
Private Function IsOpeningMarker(ByVal row As String, ByVal pos As Long) As Boolean
    If pos >= Len(row) Then Exit Function

    Dim nextChar As String
    nextChar = Mid$(row, pos + 1, 1)
    If (nextChar = " ") Or (nextChar = Mid$(row, pos, 1)) Then Exit Function

    If pos = 1 Then
        IsOpeningMarker = True
    Else
        'space, quote prefix, opening bracket, quotation mark (also the typographic ones), or another marker
        IsOpeningMarker = (InStr(" >([""'*_" & ChrW$(8222) & ChrW$(8220) & ChrW$(8218) & ChrW$(8216) & ChrW$(171) & ChrW$(187), Mid$(row, pos - 1, 1)) > 0)
    End If
End Function

'A closing marker follows text and stands at the end of a word
Private Function IsClosingMarker(ByVal row As String, ByVal pos As Long) As Boolean
    If pos = 1 Then Exit Function

    Dim previousChar As String
    previousChar = Mid$(row, pos - 1, 1)
    If (previousChar = " ") Or (previousChar = Mid$(row, pos, 1)) Then Exit Function

    If pos = Len(row) Then
        IsClosingMarker = True
    Else
        'space, punctuation, closing bracket, quotation mark (also the typographic ones), or another marker
        IsClosingMarker = (InStr(" .,;:!?)]""'*_" & ChrW$(8220) & ChrW$(8221) & ChrW$(8217) & ChrW$(171) & ChrW$(187), Mid$(row, pos + 1, 1)) > 0)
    End If
End Function

'Description:
'   Escapes a line for HTML: text between emphasis markers (roles from MatchEmphasisMarkers) is bold or underlined
'   The markers stay part of the text, so that the conversion back to plain text yields them again
'Parameters:
'   isBold, isUnderlined: the emphasis at the beginning of the line, changed to the one at its end
Private Function EmphasizedHtml(ByVal row As String, ByVal boldRoles As String, ByVal underlineRoles As String, ByRef isBold As Boolean, ByRef isUnderlined As Boolean) As String
    Dim res As String

    'an emphasis continued from the line above starts behind the quote prefix
    Dim contentStart As Long
    contentStart = 1
    Do While contentStart <= Len(row)
        If InStr("> ", Mid$(row, contentStart, 1)) = 0 Then Exit Do
        contentStart = contentStart + 1
    Loop
    Dim continuedBold As Boolean
    continuedBold = isBold
    Dim continuedUnderlined As Boolean
    continuedUnderlined = isUnderlined
    isBold = False
    isUnderlined = False

    Dim segmentStart As Long
    segmentStart = 1
    Dim segmentIsBold As Boolean
    Dim segmentIsUnderlined As Boolean

    Dim pos As Long
    For pos = 1 To Len(row)
        If pos = contentStart Then
            isBold = continuedBold
            isUnderlined = continuedUnderlined
        End If
        If Mid$(boldRoles, pos, 1) = "o" Then isBold = True
        If Mid$(underlineRoles, pos, 1) = "o" Then isUnderlined = True

        If (isBold <> segmentIsBold) Or (isUnderlined <> segmentIsUnderlined) Then
            res = res & StyledHtml(Mid$(row, segmentStart, pos - segmentStart), segmentIsBold, segmentIsUnderlined, segmentStart = 1)
            segmentStart = pos
            segmentIsBold = isBold
            segmentIsUnderlined = isUnderlined
        End If

        'the closing marker is still emphasized, the emphasis ends behind it
        If Mid$(boldRoles, pos, 1) = "c" Then isBold = False
        If Mid$(underlineRoles, pos, 1) = "c" Then isUnderlined = False
    Next
    res = res & StyledHtml(Mid$(row, segmentStart), segmentIsBold, segmentIsUnderlined, segmentStart = 1)

    EmphasizedHtml = res
End Function

Private Function StyledHtml(ByVal text As String, ByVal isBold As Boolean, ByVal isUnderlined As Boolean, ByVal atLineStart As Boolean) As String
    Dim res As String
    res = EscapeHtml(text, atLineStart)
    If Len(res) = 0 Then Exit Function

    'styles instead of <b> and <u>: the conversion back to plain text would add markers for these
    Dim style As String
    If isBold Then
        style = "font-weight:bold"
    End If
    If isUnderlined Then
        If Len(style) > 0 Then
            style = style & ";"
        End If
        style = style & "text-decoration:underline"
    End If
    If Len(style) > 0 Then
        res = "<span style=""" & style & """>" & res & "</span>"
    End If

    StyledHtml = res
End Function

Private Function HasKnownAuthor(ByRef authorOfLevel() As String, ByVal level As Long) As Boolean
    If level < LBound(authorOfLevel) Or level > UBound(authorOfLevel) Then Exit Function
    HasKnownAuthor = (Len(authorOfLevel(level)) > 0)
End Function

'The pattern (for Like) matching a condensed header, built from CONDENSED_HEADER_FORMAT
Private Function CondensedHeaderPattern() As String
    Dim pattern As String
    pattern = CONDENSED_HEADER_FORMAT
    'characters having a meaning in a pattern
    pattern = Replace$(pattern, "[", "[[]")
    pattern = Replace$(pattern, "#", "[#]")
    pattern = Replace$(pattern, "?", "[?]")
    pattern = Replace$(pattern, PATTERN_SENDER_NAME, "*")
    pattern = Replace$(pattern, PATTERN_SENDER_EMAIL, "*")
    pattern = Replace$(pattern, PATTERN_SENT_DATE, "*")
    pattern = Replace$(pattern, PATTERN_RECIPIENTS, "*")
    CondensedHeaderPattern = pattern
End Function

'Description:
'   Returns the name of the author within a condensed header: the text at the place of %SN in CONDENSED_HEADER_FORMAT
Private Function HeaderAuthor(ByVal condensedHeader As String) As String
    Dim posName As Long
    posName = InStr(CONDENSED_HEADER_FORMAT, PATTERN_SENDER_NAME)
    If posName = 0 Then
        HeaderAuthor = condensedHeader
        Exit Function
    End If

    'the text behind %SN up to the next placeholder
    Dim textBehind As String
    textBehind = Mid$(CONDENSED_HEADER_FORMAT, posName + Len(PATTERN_SENDER_NAME))
    Dim posPlaceholder As Long
    posPlaceholder = InStr(textBehind, "%")
    If posPlaceholder > 0 Then
        textBehind = Left$(textBehind, posPlaceholder - 1)
    End If

    Dim nameStart As Long
    nameStart = posName
    Dim nameEnd As Long
    If Len(textBehind) > 0 Then
        nameEnd = InStr(nameStart, condensedHeader, textBehind)
    End If
    If nameEnd = 0 Then
        nameEnd = Len(condensedHeader) + 1
    End If

    HeaderAuthor = Trim$(Mid$(condensedHeader, nameStart, nameEnd - nameStart))
End Function

'Returns the color (RGB in hexadecimal) of an author. A new author gets the next color, the user gets OWN_TEXT_COLOR
Private Function ColorOfAuthor(ByVal author As String, ByVal ownName As String, ByRef authors As String) As String
    If Len(ownName) > 0 Then
        If LCase$(author) = LCase$(ownName) Then
            ColorOfAuthor = OWN_TEXT_COLOR
            Exit Function
        End If
    End If

    Dim key As String
    key = ";" & LCase$(author) & ";"
    If InStr(authors, key) = 0 Then
        authors = authors & LCase$(author) & ";"
    End If

    'the number of the author: the number of ";" in front of its name
    Dim authorIndex As Long
    authorIndex = CountOccurrencesOfStringInString(Left$(authors, InStr(authors, key)), ";") - 1

    ColorOfAuthor = QuoteColor(authorIndex + 1)
End Function

'Returns the color (RGB in hexadecimal) of a quote level (or the n-th author)
Private Function QuoteColor(ByVal level As Long) As String
    Dim colors() As String
    colors = Split(QUOTE_COLORS, ";")

    Dim numColors As Long
    numColors = NUM_QUOTE_COLORS
    If numColors < 1 Or numColors > UBound(colors) + 1 Then
        numColors = UBound(colors) + 1
    End If

    QuoteColor = colors((level - 1) Mod numColors)
End Function

'Description:
'   Escapes the characters having a meaning in HTML; spaces at the beginning and multiple spaces are kept
'Parameters:
'   atLineStart: a space at the beginning is kept as a non-breaking space
Private Function EscapeHtml(ByVal text As String, Optional ByVal atLineStart As Boolean = True) As String
    Dim res As String
    res = Replace$(text, "&", "&amp;")
    res = Replace$(res, "<", "&lt;")
    res = Replace$(res, ">", "&gt;")
    res = Replace$(res, """", "&quot;")

    If atLineStart And (Left$(res, 1) = " ") Then
        res = "&nbsp;" & Mid$(res, 2)
    End If
    Do While InStr(res, "  ") > 0
        res = Replace$(res, "  ", " &nbsp;")
    Loop

    EscapeHtml = res
End Function

'Description:
'   Puts the content of a colored HTML (TextToColoredHtml) at the beginning of the body of another HTML,
'   e.g., in front of Outlook's signature in a reply. Head and pictures of the other HTML are kept.
'Notes:
'   * Public to enable testing
Public Function InsertColoredHtml(ByVal coloredHtml As String, ByVal html As String) As String
    Dim fragmentStart As Long
    Dim fragmentEnd As Long
    fragmentStart = InStr(coloredHtml, "<body>") + Len("<body>")
    fragmentEnd = InStrRev(coloredHtml, "</body>")

    Dim bodyStart As Long
    bodyStart = InStr(LCase$(html), "<body")
    If bodyStart > 0 Then
        bodyStart = InStr(bodyStart, html, ">")
    End If
    If bodyStart = 0 Or fragmentStart <= Len("<body>") Or fragmentEnd = 0 Then
        InsertColoredHtml = coloredHtml
        Exit Function
    End If

    InsertColoredHtml = Left$(html, bodyStart) & Mid$(coloredHtml, fragmentStart, fragmentEnd - fragmentStart) & Mid$(html, bodyStart + 1)
End Function

'Description:
'   Deletes the pictures embedded in the HTML of a mail (e.g., the logo of a signature):
'   in a plain text mail, they would show up as attachments
Private Sub deleteEmbeddedPictures(ByVal mailToSend As Object)
    Const PR_ATTACH_CONTENT_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

    Dim html As String
    html = LCase$(mailToSend.HTMLBody)

    Dim i As Long
    For i = mailToSend.Attachments.count To 1 Step -1
        Dim contentId As String
        contentId = vbNullString
        On Error Resume Next
        contentId = mailToSend.Attachments.item(i).PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID)
        On Error GoTo 0
        If Len(contentId) > 0 Then
            If InStr(html, "cid:" & LCase$(contentId)) > 0 Then
                mailToSend.Attachments.item(i).Delete
            End If
        End If
    Next
End Sub

'Description:
'   Has to be called by ThisOutlookSession before a mail is sent:
'   a colored reply is converted to plain text (COLORIZER_SEND_AS_PLAIN)
Public Sub BeforeSend(ByVal mailToSend As Object)
    If TypeName(mailToSend) <> "MailItem" Then Exit Sub

    Dim mark As UserProperty
    Set mark = mailToSend.UserProperties.Find(COLORED_MAIL_PROPERTY)
    If mark Is Nothing Then Exit Sub

    LoadConfiguration
    Dim sendAsPlain As Boolean
    sendAsPlain = COLORIZER_SEND_AS_PLAIN
    If sendAsPlain And Len(COLORIZER_HTML_RECIPIENTS) > 0 Then
        'the recipients listed get the colors
        sendAsPlain = Not AllRecipientsListed(getRecipientAddresses(mailToSend), COLORIZER_HTML_RECIPIENTS)
    End If

    If sendAsPlain Then
        'Outlook's own conversion to plain text wraps the lines anew and puts an empty line behind each paragraph,
        'therefore the text is taken from the HTML by QuoteFixHtml
        Dim plainText As String
        plainText = HtmlToPlainText(mailToSend.HTMLBody)
        'the separator of the signature needs its space at the end
        plainText = Replace$(vbCrLf & plainText & vbCrLf, vbCrLf & "--" & vbCrLf, vbCrLf & "-- " & vbCrLf)
        plainText = Mid$(plainText, 3, Len(plainText) - 4)

        deleteEmbeddedPictures mailToSend
        mailToSend.bodyFormat = olFormatPlain
        mailToSend.Body = plainText
    End If
    mark.Delete
End Sub

'Description:
'   Sends the mail shown in the active window with its colors (as HTML mail), whatever the configuration says.
'   For a button of the message window.
Public Sub SendWithColors()
    Dim mailToSend As Object
    Set mailToSend = Application.ActiveInspector.CurrentItem
    If TypeName(mailToSend) <> "MailItem" Then Exit Sub

    'without the mark, BeforeSend leaves the mail alone
    Dim mark As UserProperty
    Set mark = mailToSend.UserProperties.Find(COLORED_MAIL_PROPERTY)
    If Not mark Is Nothing Then
        mark.Delete
    End If

    mailToSend.Send
End Sub

'Description:
'   Returns the addresses of the recipients of a mail, separated by ";"
Private Function getRecipientAddresses(ByVal mailToSend As Object) As String
    Dim curRecipient As Recipient
    For Each curRecipient In mailToSend.Recipients
        Dim address As String
        address = vbNullString

        On Error Resume Next
        'a recipient of the own Exchange organization has an X.500 address, its SMTP address is in the directory
        If curRecipient.AddressEntry.Type = "EX" Then
            address = curRecipient.AddressEntry.GetExchangeUser.PrimarySmtpAddress
        End If
        If Len(address) = 0 Then
            address = curRecipient.Address
        End If
        On Error GoTo 0

        If Len(getRecipientAddresses) > 0 Then
            getRecipientAddresses = getRecipientAddresses & ";"
        End If
        getRecipientAddresses = getRecipientAddresses & address
    Next
End Function

'Description:
'   True if every recipient (addresses separated by ";") is listed, by its address or by its domain given as "@domain"
'   The comparison ignores the case. Nothing is listed if one of the two lists is empty.
'Notes:
'   * Public to enable testing
Public Function AllRecipientsListed(ByVal recipients As String, ByVal listed As String) As Boolean
    Dim recipientList() As String
    recipientList = Split(LCase$(recipients), ";")
    Dim entries() As String
    entries = Split(LCase$(listed), ";")

    Dim recipientCount As Long
    Dim i As Long
    For i = LBound(recipientList) To UBound(recipientList)
        Dim address As String
        address = Trim$(recipientList(i))
        If Len(address) > 0 Then
            recipientCount = recipientCount + 1
            If Not IsAddressListed(address, entries) Then Exit Function
        End If
    Next

    AllRecipientsListed = (recipientCount > 0)
End Function

Private Function IsAddressListed(ByVal address As String, ByRef entries() As String) As Boolean
    Dim i As Long
    For i = LBound(entries) To UBound(entries)
        Dim entry As String
        entry = Trim$(entries(i))
        If Len(entry) > 0 Then
            If Left$(entry, 1) = "@" Then
                If Right$(address, Len(entry)) = entry Then
                    IsAddressListed = True
                    Exit Function
                End If
            ElseIf address = entry Then
                IsAddressListed = True
                Exit Function
            End If
        End If
    Next
End Function
