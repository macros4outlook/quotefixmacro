Attribute VB_Name = "QuoteFixHtml"

' SPDX-License-Identifier: BSD-3-Clause

'Converts the HTML of a mail into plain text
'
'The HTML is scanned as text. It is not handed to a browser component,
'because it comes from arbitrary senders: nothing of it may be executed or loaded.
'
'  * Quotes (<blockquote>) are marked with ">", one for each level
'  * Paragraphs, line breaks, list items and table rows become lines
'  * Bold text is marked as *text*, underlined text as _text_
'  * Scripts, styles and comments are dropped

'@Folder("QuoteFixMacro")
Option Explicit
Option Private Module

'Outlook shows a horizontal line this way in plain text
Private Const HORIZONTAL_RULE As String = "________________________________"

'name=code of the supported named character references (the names are case-sensitive)
Private Const NAMED_ENTITIES As String = _
    ";amp=38;lt=60;gt=62;quot=34;apos=39;nbsp=160" & _
    ";auml=228;ouml=246;uuml=252;Auml=196;Ouml=214;Uuml=220;szlig=223" & _
    ";agrave=224;aacute=225;acirc=226;egrave=232;eacute=233;ecirc=234" & _
    ";iacute=237;oacute=243;ocirc=244;uacute=250;ccedil=231;ntilde=241" & _
    ";euro=8364;copy=169;reg=174;trade=8482;sect=167;deg=176;middot=183" & _
    ";bull=8226;hellip=8230;ndash=8211;mdash=8212;laquo=171;raquo=187" & _
    ";lsquo=8216;rsquo=8217;sbquo=8218;ldquo=8220;rdquo=8221;bdquo=8222;shy=173;"

Private Const CODE_NON_BREAKING_SPACE As Long = 160
Private Const CODE_SOFT_HYPHEN As Long = 173

'Module Variables to make code more readable (-> parameter passing gets easier)

'the lines converted so far
Private resultText As String
Private lengthBeforeLastLine As Long
'the quote level of the last line if that line is empty, -1 otherwise
Private lastEmptyLineLevel As Long
'False at the beginning of the text and at the beginning of a quote: empty lines are dropped there
Private blockHasLines As Boolean

'the line being built
Private curLine As String
'True if the line has to be output even if it has no text (e.g., it consists of a non-breaking space)
Private curLineHasContent As Boolean
'the marker of the list item the line belongs to
Private curMarker As String
Private pendingSpace As Boolean

Private quoteLevel As Long
Private preLevel As Long
Private atPreStart As Boolean
Private paragraphIsCompact As Boolean

'for each open list: -1 for an unordered list, otherwise the number of its last item
Private listCounters(1 To 20) As Long
Private listLevel As Long

'the emphasis markers (* bold, _ underlined) waiting for the first character of their text
Private pendingMarkers As String
'the emphasis markers whose text has begun, written by this module: they get a closing marker
Private openMarkers As String
'the emphasis markers whose text has begun with the marker itself (e.g., *text* in bold): the text has its closing marker
Private carriedMarkers As String
'the markers the text has opened, but not closed yet (e.g., a text in bold running over several paragraphs)
Private textOpenMarkers As String
Private boldDepth As Long
Private underlineDepth As Long

Private inAnchor As Boolean
Private anchorHref As String
Private anchorText As String

'Converts HTML into plain text
'Quoted text is prefixed by ">" (one for each level of <blockquote>)
'
'Notes:
'  * Public to enable testing
Public Function HtmlToPlainText(ByVal html As String) As String
    resultText = vbNullString
    lengthBeforeLastLine = 0
    lastEmptyLineLevel = -1
    blockHasLines = False
    curLine = vbNullString
    curLineHasContent = False
    curMarker = vbNullString
    pendingSpace = False
    quoteLevel = 0
    preLevel = 0
    atPreStart = False
    paragraphIsCompact = False
    listLevel = 0
    pendingMarkers = vbNullString
    openMarkers = vbNullString
    carriedMarkers = vbNullString
    textOpenMarkers = vbNullString
    boldDepth = 0
    underlineDepth = 0
    inAnchor = False

    Dim pos As Long
    pos = 1
    Do While pos <= Len(html)
        Dim tagStart As Long
        tagStart = InStr(pos, html, "<")
        If tagStart = 0 Then
            AppendText Mid$(html, pos)
            Exit Do
        End If
        If tagStart > pos Then
            AppendText Mid$(html, pos, tagStart - pos)
        End If

        If Mid$(html, tagStart, 4) = "<!--" Then
            pos = PosBehind(html, tagStart + 4, "-->")
        ElseIf Not IsTagStart(Mid$(html, tagStart + 1, 1)) Then
            'a single "<" is text
            AppendDecodedText "<"
            pos = tagStart + 1
        Else
            Dim tagEnd As Long
            tagEnd = FindTagEnd(html, tagStart)
            pos = tagEnd + 1
            HandleTag Mid$(html, tagStart + 1, tagEnd - tagStart - 1), html, pos
        End If
    Loop
    FinishLine False

    'no empty line at the end
    If lastEmptyLineLevel >= 0 Then
        DropLastLine
    End If
    'no line break at the end
    If Right$(resultText, 2) = vbCrLf Then
        resultText = Left$(resultText, Len(resultText) - 2)
    End If

    HtmlToPlainText = resultText
End Function

Private Function IsTagStart(ByVal c As String) As Boolean
    IsTagStart = (c Like "[A-Za-z/!?]")
End Function

Private Function IsWhiteSpace(ByVal c As String) As Boolean
    IsWhiteSpace = (c = " ") Or (c = vbTab) Or (c = vbCr) Or (c = vbLf)
End Function

'Returns the position of the ">" finishing the tag which starts at tagStart
'A ">" within a quoted attribute value does not finish the tag
'
'Notes:
'  * html is passed by reference to avoid copying it
Private Function FindTagEnd(ByRef html As String, ByVal tagStart As Long) As Long
    Dim quoteChar As String
    Dim afterEquals As Boolean

    Dim i As Long
    For i = tagStart + 1 To Len(html)
        Dim c As String
        c = Mid$(html, i, 1)
        If Len(quoteChar) > 0 Then
            If c = quoteChar Then
                quoteChar = vbNullString
            End If
        ElseIf c = ">" Then
            FindTagEnd = i
            Exit Function
        ElseIf afterEquals And (c = """" Or c = "'") Then
            quoteChar = c
            afterEquals = False
        ElseIf c = "=" Then
            afterEquals = True
        ElseIf Not IsWhiteSpace(c) Then
            afterEquals = False
        End If
    Next

    'the tag is not finished
    FindTagEnd = Len(html) + 1
End Function

'Returns the position behind the next occurrence of terminator
Private Function PosBehind(ByRef html As String, ByVal start As Long, ByVal terminator As String) As Long
    Dim found As Long
    found = InStr(start, html, terminator)
    If found = 0 Then
        PosBehind = Len(html) + 1
    Else
        PosBehind = found + Len(terminator)
    End If
End Function

'Returns the position behind the closing tag of an element (e.g., behind "</style>")
Private Function PosBehindClosingTag(ByRef html As String, ByVal start As Long, ByVal tagName As String) As Long
    Dim found As Long
    found = InStr(start, html, "</")
    Do While found > 0
        If LCase$(Mid$(html, found + 2, Len(tagName))) = tagName Then
            PosBehindClosingTag = FindTagEnd(html, found) + 1
            Exit Function
        End If
        found = InStr(found + 2, html, "</")
    Loop

    'the element is not closed
    PosBehindClosingTag = Len(html) + 1
End Function

'Returns the lower-cased name of a tag ("p" for "p class=x" and for "/p")
Private Function GetTagName(ByVal tag As String) As String
    Dim start As Long
    start = 1
    If Left$(tag, 1) = "/" Then
        start = 2
    End If

    Dim i As Long
    i = start
    Do While i <= Len(tag)
        Dim c As String
        c = Mid$(tag, i, 1)
        If IsWhiteSpace(c) Or c = "/" Then Exit Do
        i = i + 1
    Loop

    GetTagName = LCase$(Mid$(tag, start, i - start))
End Function

Private Function SkipWhiteSpace(ByVal text As String, ByVal start As Long) As Long
    Dim i As Long
    i = start
    Do While i <= Len(text)
        If Not IsWhiteSpace(Mid$(text, i, 1)) Then Exit Do
        i = i + 1
    Loop
    SkipWhiteSpace = i
End Function

'Returns the value of an attribute of a tag (vbNullString if the tag does not have it)
Private Function GetAttribute(ByVal tag As String, ByVal attributeName As String) As String
    Dim pos As Long
    pos = InStr(1, tag, attributeName, vbTextCompare)
    Do While pos > 0
        Dim valueStart As Long
        valueStart = SkipWhiteSpace(tag, pos + Len(attributeName))

        'the name of an attribute follows white space and is followed by "="
        If pos > 1 And Mid$(tag, valueStart, 1) = "=" Then
            If IsWhiteSpace(Mid$(tag, pos - 1, 1)) Then
                valueStart = SkipWhiteSpace(tag, valueStart + 1)

                Dim quoteChar As String
                quoteChar = Mid$(tag, valueStart, 1)

                Dim valueEnd As Long
                If quoteChar = """" Or quoteChar = "'" Then
                    valueStart = valueStart + 1
                    valueEnd = InStr(valueStart, tag, quoteChar)
                    If valueEnd = 0 Then
                        valueEnd = Len(tag) + 1
                    End If
                Else
                    valueEnd = valueStart
                    Do While valueEnd <= Len(tag)
                        If IsWhiteSpace(Mid$(tag, valueEnd, 1)) Then Exit Do
                        valueEnd = valueEnd + 1
                    Loop
                End If

                GetAttribute = DecodeEntities(Mid$(tag, valueStart, valueEnd - valueStart))
                Exit Function
            End If
        End If

        pos = InStr(pos + 1, tag, attributeName, vbTextCompare)
    Loop
End Function

'Parameters:
'  tag:  the tag without "<" and ">"
'  html: the complete HTML
'  pos:  the position behind the tag. It is moved if the content of the element has to be skipped
Private Sub HandleTag(ByVal tag As String, ByRef html As String, ByRef pos As Long)
    Dim isClosingTag As Boolean
    isClosingTag = (Left$(tag, 1) = "/")

    Dim tagName As String
    tagName = GetTagName(tag)

    Select Case tagName
        Case "script", "style", "title"
            'the content is not part of the text
            If Not isClosingTag Then
                pos = PosBehindClosingTag(html, pos, tagName)
            End If

        Case "br"
            FinishLine True

        Case "p", "h1", "h2", "h3", "h4", "h5", "h6"
            FinishLine False
            If Not isClosingTag Then
                paragraphIsCompact = (listLevel > 0) Or IsCompactParagraph(tag)
            End If
            If Not paragraphIsCompact Then
                'there is space between paragraphs
                EmitLine vbNullString
            End If

        Case "div", "table", "tr", "dl", "dt", "dd", "center", "address", "section", "article", "header", "footer"
            FinishLine False

        Case "td", "th"
            'the cells of a table row are put into one line
            pendingSpace = True

        Case "blockquote"
            FinishLine False
            If isClosingTag Then
                'no empty line at the end of a quote
                If quoteLevel > 0 And lastEmptyLineLevel = quoteLevel Then
                    DropLastLine
                End If
                If quoteLevel > 0 Then
                    quoteLevel = quoteLevel - 1
                End If
                blockHasLines = (Len(resultText) > 0)
            Else
                quoteLevel = quoteLevel + 1
                blockHasLines = False
            End If

        Case "pre"
            FinishLine False
            If isClosingTag Then
                If preLevel > 0 Then
                    preLevel = preLevel - 1
                End If
            Else
                preLevel = preLevel + 1
                atPreStart = True
            End If

        Case "hr"
            If Not isClosingTag Then
                FinishLine False
                EmitLine HORIZONTAL_RULE
            End If

        Case "ul", "ol"
            FinishLine False
            If isClosingTag Then
                If listLevel > 0 Then
                    listLevel = listLevel - 1
                End If
            Else
                listLevel = listLevel + 1
                If listLevel <= UBound(listCounters) Then
                    If tagName = "ol" Then
                        listCounters(listLevel) = 0
                    Else
                        listCounters(listLevel) = -1
                    End If
                End If
            End If

        Case "li"
            FinishLine False
            If isClosingTag Then
                curMarker = vbNullString
            Else
                curMarker = GetListMarker()
            End If

        Case "b", "strong"
            HandleEmphasis "*", isClosingTag, boldDepth

        Case "u"
            HandleEmphasis "_", isClosingTag, underlineDepth

        Case "img"
            If Not isClosingTag Then
                Dim altText As String
                altText = GetAttribute(tag, "alt")
                If Len(altText) > 0 Then
                    AppendDecodedText "[" & altText & "]"
                End If
            End If

        Case "a"
            If isClosingTag Then
                FinishAnchor
            Else
                anchorHref = GetAttribute(tag, "href")
                anchorText = vbNullString
                inAnchor = True
            End If
    End Select
End Sub

'Description:
'   Marks bold text as *text* and underlined text as _text_
'   Nested elements of the same kind get one pair of markers, elements without text get none
'   The markers of a text running over several lines are put at its beginning and at its end
Private Sub HandleEmphasis(ByVal marker As String, ByVal isClosingTag As Boolean, ByRef depth As Long)
    'preformatted text is kept as it is
    If preLevel > 0 Then Exit Sub

    If Not isClosingTag Then
        depth = depth + 1
        If depth = 1 Then
            pendingMarkers = pendingMarkers & marker
        End If
        Exit Sub
    End If

    If depth = 0 Then Exit Sub
    depth = depth - 1
    If depth > 0 Then Exit Sub

    If InStr(pendingMarkers, marker) > 0 Then
        'no text in between
        pendingMarkers = Replace$(pendingMarkers, marker, vbNullString)
    ElseIf InStr(openMarkers, marker) > 0 Then
        openMarkers = Replace$(openMarkers, marker, vbNullString)
        WriteClosingMarker marker
    Else
        carriedMarkers = Replace$(carriedMarkers, marker, vbNullString)
    End If
End Sub

'Description:
'   Puts the pending opening markers in front of the first character of the text
'   A marker is not doubled if the text starts with it already (e.g., *text* in bold)
Private Sub WritePendingMarkers(ByVal firstChar As String)
    Dim i As Long
    For i = 1 To Len(pendingMarkers)
        Dim marker As String
        marker = Mid$(pendingMarkers, i, 1)
        If marker = firstChar Then
            carriedMarkers = carriedMarkers & marker
            textOpenMarkers = textOpenMarkers & marker
        ElseIf InStr(textOpenMarkers, marker) > 0 Then
            'the text continues one which has its opening marker already (Word puts each paragraph into an element of its own)
            carriedMarkers = carriedMarkers & marker
        Else
            curLine = curLine & marker
            openMarkers = openMarkers & marker
        End If
    Next
    pendingMarkers = vbNullString
End Sub

'Description:
'   Puts a closing marker behind the last character of the text (of the last line if the current one is empty)
'   A marker is not doubled if the text ends with it already
Private Sub WriteClosingMarker(ByVal marker As String)
    If Len(curLine) > 0 Then
        If Right$(curLine, 1) <> marker Then
            curLine = curLine & marker
        End If
        Exit Sub
    End If

    'the end of the last line with text: an empty line may follow it
    Dim lineEnd As Long
    If lastEmptyLineLevel >= 0 Then
        lineEnd = lengthBeforeLastLine
    Else
        lineEnd = Len(resultText)
    End If
    If lineEnd < 3 Then Exit Sub
    If Mid$(resultText, lineEnd - 2, 1) = marker Then Exit Sub

    resultText = Left$(resultText, lineEnd - 2) & marker & Mid$(resultText, lineEnd - 1)
    If lastEmptyLineLevel >= 0 Then
        lengthBeforeLastLine = lengthBeforeLastLine + 1
    End If
End Sub

'Paragraphs written by Word (Outlook) and paragraphs without margin are shown without space in between
Private Function IsCompactParagraph(ByVal tag As String) As Boolean
    If LCase$(Left$(GetAttribute(tag, "class"), 3)) = "mso" Then
        IsCompactParagraph = True
        Exit Function
    End If

    Dim style As String
    style = Replace$(LCase$(GetAttribute(tag, "style")), " ", vbNullString)
    IsCompactParagraph = HasZeroValue(style, "margin:") Or HasZeroValue(style, "margin-bottom:")
End Function

'True for "margin:0" and "margin:0cm", False for "margin:0.5em"
Private Function HasZeroValue(ByVal style As String, ByVal cssProperty As String) As Boolean
    Dim pos As Long
    pos = InStr(style, cssProperty & "0")
    If pos > 0 Then
        HasZeroValue = (Mid$(style, pos + Len(cssProperty) + 1, 1) <> ".")
    End If
End Function

Private Function GetListMarker() As String
    Dim marker As String
    marker = "* "

    If listLevel >= 1 And listLevel <= UBound(listCounters) Then
        If listCounters(listLevel) >= 0 Then
            listCounters(listLevel) = listCounters(listLevel) + 1
            marker = listCounters(listLevel) & ". "
        End If
    End If

    'nested lists are indented
    If listLevel > 1 Then
        marker = String$((listLevel - 1) * 2, " ") & marker
    End If

    GetListMarker = marker
End Function

'Description:
'   Adds the target of a link behind its text: text <target>
'   Nothing is added if the text already shows the target
Private Sub FinishAnchor()
    If Not inAnchor Then Exit Sub
    inAnchor = False

    Dim target As String
    target = Trim$(anchorHref)
    If LCase$(Left$(target, 7)) = "mailto:" Then
        'a mention ("@Name") is kept as it is
        If Left$(Trim$(anchorText), 1) = "@" Then Exit Sub

        target = Mid$(target, 8)
        'drop parameters such as "?subject="
        If InStr(target, "?") > 0 Then
            target = Left$(target, InStr(target, "?") - 1)
        End If
    ElseIf LCase$(Left$(target, 7)) <> "http://" And LCase$(Left$(target, 8)) <> "https://" Then
        'other targets (e.g., positions within the mail) are of no use in plain text
        Exit Sub
    End If

    If Len(target) = 0 Or Len(Trim$(anchorText)) = 0 Then Exit Sub
    If NormalizeLink(anchorText) = NormalizeLink(target) Then Exit Sub
    'a text being a link itself is kept as it is: the target is the same link wrapped by a link checker
    If IsLink(anchorText) Then Exit Sub

    AppendDecodedText " <" & target & ">"
End Sub

Private Function IsLink(ByVal text As String) As Boolean
    Dim res As String
    res = LCase$(Trim$(text))
    IsLink = (Left$(res, 7) = "http://") Or (Left$(res, 8) = "https://") Or (Left$(res, 4) = "www.")
End Function

Private Function NormalizeLink(ByVal link As String) As String
    Dim res As String
    res = LCase$(Trim$(link))

    If Left$(res, 7) = "mailto:" Then
        res = Mid$(res, 8)
    ElseIf Left$(res, 8) = "https://" Then
        res = Mid$(res, 9)
    ElseIf Left$(res, 7) = "http://" Then
        res = Mid$(res, 8)
    End If
    If Right$(res, 1) = "/" Then
        res = Left$(res, Len(res) - 1)
    End If

    NormalizeLink = res
End Function

'Description:
'   Adds text of the HTML (still containing character references) to the current line
Private Sub AppendText(ByVal text As String)
    AppendDecodedText DecodeEntities(text)
End Sub

'Description:
'   Adds text to the current line
'   A sequence of white space is a single space, white space at the beginning of a line is dropped
Private Sub AppendDecodedText(ByVal text As String)
    If inAnchor Then
        anchorText = anchorText & text
    End If

    If preLevel > 0 Then
        AppendPreformattedText text
        Exit Sub
    End If

    Dim nonBreakingSpace As String
    nonBreakingSpace = ChrW$(CODE_NON_BREAKING_SPACE)

    'characters without width: zero width space, zero width non-joiner, zero width joiner, byte order mark
    Dim zeroWidthSpace As String
    zeroWidthSpace = ChrW$(8203)
    Dim zeroWidthNonJoiner As String
    zeroWidthNonJoiner = ChrW$(8204)
    Dim zeroWidthJoiner As String
    zeroWidthJoiner = ChrW$(8205)
    Dim byteOrderMark As String
    byteOrderMark = ChrW$(65279)

    Dim i As Long
    For i = 1 To Len(text)
        Dim c As String
        c = Mid$(text, i, 1)
        Select Case c
            Case " ", vbTab, vbCr, vbLf
                pendingSpace = True
            Case nonBreakingSpace
                'kept as a space of its own: it is not collapsed with the white space around it
                If pendingSpace And Len(curLine) > 0 Then
                    curLine = curLine & " "
                End If
                pendingSpace = False
                curLine = curLine & " "
                curLineHasContent = True
            Case zeroWidthSpace, zeroWidthNonJoiner, zeroWidthJoiner, byteOrderMark
                'invisible: dropped
            Case Else
                If pendingSpace And Len(curLine) > 0 Then
                    curLine = curLine & " "
                End If
                pendingSpace = False
                If InStr(textOpenMarkers, c) > 0 Then
                    'the text closes its marker
                    textOpenMarkers = Replace$(textOpenMarkers, c, vbNullString)
                End If
                If Len(pendingMarkers) > 0 Then
                    WritePendingMarkers c
                End If
                curLine = curLine & c
                curLineHasContent = True
        End Select
    Next
End Sub

'Description:
'   Adds text of a <pre> element to the current line: white space and line breaks are kept
Private Sub AppendPreformattedText(ByVal text As String)
    Dim normalizedText As String
    normalizedText = Replace$(Replace$(text, vbCrLf, vbLf), vbCr, vbLf)
    normalizedText = Replace$(normalizedText, ChrW$(CODE_NON_BREAKING_SPACE), " ")

    'a line break directly behind <pre> is not part of the text
    If atPreStart And Left$(normalizedText, 1) = vbLf Then
        normalizedText = Mid$(normalizedText, 2)
    End If
    atPreStart = False

    Dim lines() As String
    lines = Split(normalizedText, vbLf)

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If i > LBound(lines) Then
            FinishLine True
        End If
        curLine = curLine & lines(i)
        If Len(lines(i)) > 0 Then
            curLineHasContent = True
        End If
    Next
End Sub

'Description:
'   Finishes the current line
'   A line without text is only output if force is set or if it has content (non-breaking space)
Private Sub FinishLine(ByVal force As Boolean)
    Dim text As String
    text = RTrim$(curLine)
    Dim hasContent As Boolean
    hasContent = curLineHasContent

    curLine = vbNullString
    curLineHasContent = False
    pendingSpace = False

    If Len(text) = 0 And Not hasContent And Not force Then
        'nothing to output. A list marker is kept for the text to come
        Exit Sub
    End If

    EmitLine RTrim$(curMarker & text)
    curMarker = vbNullString
End Sub

'Description:
'   Adds a line to the result, prefixed according to the current quote level
'   An empty line is neither output at the beginning (of the text or of a quote) nor after an empty line
Private Sub EmitLine(ByVal text As String)
    If Len(text) = 0 Then
        If Not blockHasLines Or lastEmptyLineLevel = quoteLevel Then Exit Sub
    End If

    lengthBeforeLastLine = Len(resultText)
    If quoteLevel = 0 Then
        resultText = resultText & text & vbCrLf
    ElseIf Len(text) = 0 Then
        resultText = resultText & String$(quoteLevel, ">") & vbCrLf
    Else
        resultText = resultText & String$(quoteLevel, ">") & " " & text & vbCrLf
    End If

    If Len(text) = 0 Then
        lastEmptyLineLevel = quoteLevel
        'a text marked by the text itself does not continue over an empty line
        textOpenMarkers = vbNullString
    Else
        lastEmptyLineLevel = -1
    End If
    blockHasLines = True
End Sub

Private Sub DropLastLine()
    resultText = Left$(resultText, lengthBeforeLastLine)
    lastEmptyLineLevel = -1
End Sub

'Replaces the character references ("&amp;", "&#228;", "&#xE4;") by their characters
'Unknown references are kept
Private Function DecodeEntities(ByVal text As String) As String
    Dim pos As Long
    pos = InStr(text, "&")
    If pos = 0 Then
        DecodeEntities = text
        Exit Function
    End If

    Dim res As String
    'position of the first character which is not yet part of res
    Dim copyStart As Long
    copyStart = 1

    Do While pos > 0
        Dim isEntity As Boolean
        isEntity = False

        'the longest reference is "&#1114111;"
        Dim entityEnd As Long
        entityEnd = InStr(pos, text, ";")
        If entityEnd > 0 And entityEnd - pos <= 9 Then
            Dim decoded As String
            isEntity = DecodeEntity(Mid$(text, pos + 1, entityEnd - pos - 1), decoded)
        End If

        If isEntity Then
            res = res & Mid$(text, copyStart, pos - copyStart) & decoded
            copyStart = entityEnd + 1
            pos = InStr(copyStart, text, "&")
        Else
            pos = InStr(pos + 1, text, "&")
        End If
    Loop

    DecodeEntities = res & Mid$(text, copyStart)
End Function

'Parameters:
'  entity:  the reference without "&" and ";"
'  decoded: the character (returned by reference)
'Returns False if the reference is not known
Private Function DecodeEntity(ByVal entity As String, ByRef decoded As String) As Boolean
    Dim code As Long
    code = -1

    If Left$(entity, 1) = "#" Then
        If LCase$(Mid$(entity, 2, 1)) = "x" Then
            code = ParseNumber(Mid$(entity, 3), 16)
        Else
            code = ParseNumber(Mid$(entity, 2), 10)
        End If
    ElseIf Len(entity) > 0 Then
        Dim entry As Long
        entry = InStr(NAMED_ENTITIES, ";" & entity & "=")
        If entry > 0 Then
            Dim codeStart As Long
            codeStart = entry + Len(entity) + 2
            code = ParseNumber(Mid$(NAMED_ENTITIES, codeStart, InStr(codeStart, NAMED_ENTITIES, ";") - codeStart), 10)
        End If
    End If

    If code < 1 Or code > 1114111 Then
        DecodeEntity = False
        Exit Function
    End If

    If code = CODE_SOFT_HYPHEN Then
        'only marks a position where a word may be broken
        decoded = vbNullString
    ElseIf code < 65536 Then
        decoded = ChrW$(code)
    Else
        'characters behind U+FFFF consist of two code units (surrogate pair)
        code = code - 65536
        decoded = ChrW$(55296 + code \ 1024) & ChrW$(56320 + (code And 1023))
    End If
    DecodeEntity = True
End Function

'Returns the value of a number given in the radix 10 or 16, -1 if it is not a number
Private Function ParseNumber(ByVal digits As String, ByVal radix As Long) As Long
    ParseNumber = -1
    'more digits are not needed for character codes and would cause an overflow
    If Len(digits) = 0 Or Len(digits) > 7 Then Exit Function

    Dim res As Long
    Dim i As Long
    For i = 1 To Len(digits)
        Dim digit As Long
        digit = InStr(Left$("0123456789abcdef", radix), LCase$(Mid$(digits, i, 1))) - 1
        If digit < 0 Then Exit Function
        res = res * radix + digit
    Next

    ParseNumber = res
End Function
