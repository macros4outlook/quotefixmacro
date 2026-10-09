Attribute VB_Name = "QuoteFixMacroTest"

Option Explicit
Option Private Module

'@TestModule
'@Folder("QuoteFixMacro.Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Private outlookOutput As String
Private expectedResult As String

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..

    'Currently required for reformat only
    QuoteFixMacro.LoadConfiguration
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'Required settings:
'
'USE_COLORIZER unset
'INCLUDE_QUOTES_TO_LEVEL = -1
'LINE_WRAP_AFTER = 75

'@TestMethod("reformat")
Private Sub reformatTest1()
    On Error GoTo TestFail

    outlookOutput = vbNullString & _
        "> >>" & vbNewLine & _
        "> >> I have a Win 2k3 SBS and I want to replicate the users into my" & vbNewLine & _
        "> OpenLDAP" & vbNewLine & _
        "> >> 2.4.11." & vbNewLine & _
        "> >" & vbNewLine & _
        "> > This is not possible. You could however implement your own sync" & vbNewLine & _
        "> process" & vbNewLine & _
        "> > in your favourite scripting/programming language." & vbNewLine & _
        "> " & vbNewLine & _
        "> Actually we have done some preliminary work..."
    expectedResult = vbNullString & _
        ">>> " & vbNewLine & _
        ">>> I have a Win 2k3 SBS and I want to replicate the users into my" & vbNewLine & _
        ">>> OpenLDAP 2.4.11." & vbNewLine & _
        ">> " & vbNewLine & _
        ">> This is not possible. You could however implement your own sync process" & vbNewLine & _
        ">> in your favourite scripting/programming language." & vbNewLine & _
        "> " & vbNewLine & _
        "> Actually we have done some preliminary work..."

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("reformat")
Private Sub reformatNoReformat()
    On Error GoTo TestFail

    outlookOutput = vbNullString & _
        "> Moin," & vbNewLine & _
        "> " & vbNewLine & _
        "> Kurzanleitung """"Deckel öffnen"""":" & vbNewLine & _
        "> 1. Unten rechts die Kunststoff-Abdeckung mit einem Schraubendreher" & vbNewLine & _
        "> nach rechts schieben." & vbNewLine & _
        "> 2. Das Blech nach links schieben." & vbNewLine & _
        "> 3. Kreuzschlitzschraube lösen." & vbNewLine & _
        "> " & vbNewLine & _
        "> " & vbNewLine & _
        "> Mit freundlichen Grüßen" & vbNewLine & _
        "> " & vbNewLine & _
        "> company" & vbNewLine & _
        "> Jon Doe"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput)

    Assert.AreEqual outlookOutput, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("reformat")
Private Sub reformatGreetingsKept()
    On Error GoTo TestFail

    outlookOutput = vbNullString & _
        "> Hallo Jon, ich hatte mal von xxxxxx ein Anti-Virus Programm, aber ich" & vbNewLine & _
        "> habe" & vbNewLine & _
        "> so viele Spams trotzdem erhalten, dass ich das nicht mehr abonniert" & vbNewLine & _
        "> habe." & vbNewLine & _
        "> xxx xxxxx? Haste eine Lösung für mein Virenprogramm, kann ich was" & vbNewLine & _
        "> runterladen?" & vbNewLine & _
        "> Lieben Gruß Jane"
    expectedResult = vbNullString & _
        "> Hallo Jon, ich hatte mal von xxxxxx ein Anti-Virus Programm, aber ich" & vbNewLine & _
        "> habe so viele Spams trotzdem erhalten, dass ich das nicht mehr abonniert" & vbNewLine & _
        "> habe. xxx xxxxx? Haste eine Lösung für mein Virenprogramm, kann" & vbNewLine & _
        "> ich was runterladen?" & vbNewLine & _
        "> Lieben Gruß Jane"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput)

    'TODO: Keeping the greeting unformatted is currently not implemented
    'Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'Required settings:
'
'CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS = True
'CONDENSED_HEADER_FORMAT = "%SN wrote on %D:"
'DATE_FORMAT = "yyyy-mm-dd HH:MM"
'
'The dates in the headers are written as yyyy-mm-dd, because the parsing of other formats depends on the regional settings of Windows

'@TestMethod("condense")
Private Sub condenseHeaderDateWithWeekday()
    On Error GoTo TestFail

    outlookOutput = vbNullString & _
        "> answer of Adam" & vbNewLine & _
        "> " & vbNewLine & _
        "> > -----Original Message-----" & vbNewLine & _
        "> > From: Art Ross" & vbNewLine & _
        "> > Sent: Thursday, 2011-04-07 09:36" & vbNewLine & _
        "> > To: Adam Swift" & vbNewLine & _
        "> > Subject: RE: Testing" & vbNewLine & _
        "> > " & vbNewLine & _
        "> > Hi Adam,"
    expectedResult = vbNullString & _
        "> answer of Adam" & vbNewLine & _
        "> " & vbNewLine & _
        "> Art Ross wrote on 2011-04-07 09:36:" & vbNewLine & _
        ">> Hi Adam,"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("condense")
Private Sub condenseHeaderDateWithoutWeekday()
    On Error GoTo TestFail

    outlookOutput = vbNullString & _
        "> answer of Adam" & vbNewLine & _
        "> " & vbNewLine & _
        "> > -----Original Message-----" & vbNewLine & _
        "> > From: Art Ross" & vbNewLine & _
        "> > Sent: 2011-04-07 09:36" & vbNewLine & _
        "> > To: Adam Swift" & vbNewLine & _
        "> > Subject: RE: Testing" & vbNewLine & _
        "> > " & vbNewLine & _
        "> > Hi Adam,"
    expectedResult = vbNullString & _
        "> answer of Adam" & vbNewLine & _
        "> " & vbNewLine & _
        "> Art Ross wrote on 2011-04-07 09:36:" & vbNewLine & _
        ">> Hi Adam,"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("condense")
Private Sub condenseHeaderWithBoldLabels()
    On Error GoTo TestFail

    'HTML mails show the labels in bold, HtmlToPlainText marks them
    outlookOutput = vbNullString & _
        "> answer of Adam" & vbNewLine & _
        "> " & vbNewLine & _
        "> > *From:* Art Ross" & vbNewLine & _
        "> > *Sent:* 2011-04-07 09:36" & vbNewLine & _
        "> > *To:* Adam Swift" & vbNewLine & _
        "> > *Subject:* RE: Testing" & vbNewLine & _
        "> > " & vbNewLine & _
        "> > Hi Adam,"
    expectedResult = vbNullString & _
        "> answer of Adam" & vbNewLine & _
        "> " & vbNewLine & _
        "> Art Ross wrote on 2011-04-07 09:36:" & vbNewLine & _
        ">> Hi Adam,"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("condense")
Private Sub condenseHeaderUnparsableDateIsNotTakenFromPreviousHeader()
    On Error GoTo TestFail

    outlookOutput = vbNullString & _
        "> answer of Adam" & vbNewLine & _
        "> " & vbNewLine & _
        "> > -----Original Message-----" & vbNewLine & _
        "> > From: Art Ross" & vbNewLine & _
        "> > Sent: Thursday, 2011-04-07 09:36" & vbNewLine & _
        "> > To: Adam Swift" & vbNewLine & _
        "> > Subject: RE: Testing" & vbNewLine & _
        "> > " & vbNewLine & _
        "> > Hi Adam," & vbNewLine & _
        "> > " & vbNewLine & _
        "> > > -----Original Message-----" & vbNewLine & _
        "> > > From: Adam Swift" & vbNewLine & _
        "> > > Sent: some day" & vbNewLine & _
        "> > > To: Art Ross" & vbNewLine & _
        "> > > Subject: Testing" & vbNewLine & _
        "> > > " & vbNewLine & _
        "> > > Hi Art,"
    expectedResult = vbNullString & _
        "> answer of Adam" & vbNewLine & _
        "> " & vbNewLine & _
        "> Art Ross wrote on 2011-04-07 09:36:" & vbNewLine & _
        ">> Hi Adam," & vbNewLine & _
        ">> " & vbNewLine & _
        ">> Adam Swift wrote on some day:" & vbNewLine & _
        ">>> Hi Art,"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("reformat")
Private Sub reformatLastLineWithLowerNestingIsKept()
    On Error GoTo TestFail

    outlookOutput = vbNullString & _
        "> > quoted line" & vbNewLine & _
        "> last line"
    expectedResult = vbNullString & _
        ">> quoted line" & vbNewLine & _
        "> last line"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'The following tests are on text which has been converted from HTML and prefixed by QuoteText (SourceHtml)

'@TestMethod("reformat")
Private Sub reformatWithoutRepairWrapsEachLongLineOnItsOwn()
    On Error GoTo TestFail

    outlookOutput = vbNullString & _
        "> Hi Adam," & vbNewLine & _
        "> this is a very long paragraph line that was never wrapped by the sender because it comes from an HTML mail where paragraphs are just one line each." & vbNewLine & _
        "> Thanks" & vbNewLine & _
        "> Art"
    expectedResult = vbNullString & _
        "> Hi Adam," & vbNewLine & _
        "> this is a very long paragraph line that was never wrapped by the sender" & vbNewLine & _
        "> because it comes from an HTML mail where paragraphs are just one line" & vbNewLine & _
        "> each." & vbNewLine & _
        "> Thanks" & vbNewLine & _
        "> Art"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourceHtml)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("reformat")
Private Sub reformatWithoutRepairKeepsInterleavedLines()
    On Error GoTo TestFail

    outlookOutput = vbNullString & _
        "> > question one?" & vbNewLine & _
        "> Yes." & vbNewLine & _
        "> > question two?" & vbNewLine & _
        "> No."
    expectedResult = vbNullString & _
        ">> question one?" & vbNewLine & _
        "> Yes." & vbNewLine & _
        ">> question two?" & vbNewLine & _
        "> No."

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourceHtml)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("reformat")
Private Sub reformatWithoutRepairDoesNotBreakLongWords()
    On Error GoTo TestFail

    outlookOutput = vbNullString & _
        "> > see https://example.com/a/very/long/link/which/is/longer/than/the/line/width/of/seventy-five/characters for details"
    expectedResult = vbNullString & _
        ">> see" & vbNewLine & _
        ">> https://example.com/a/very/long/link/which/is/longer/than/the/line/width/of/seventy-five/characters" & vbNewLine & _
        ">> for details"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourceHtml)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("quote")
Private Sub quoteTextPrefixesEachLine()
    On Error GoTo TestFail

    outlookOutput = vbNullString & _
        "line 1" & vbNewLine & _
        vbNewLine & _
        "> quoted" & vbNewLine & _
        vbNewLine
    expectedResult = vbNullString & _
        "> line 1" & vbNewLine & _
        "> " & vbNewLine & _
        "> > quoted"

    Assert.AreEqual expectedResult, QuoteFixMacro.QuoteText(outlookOutput)
    Assert.AreEqual vbNullString, QuoteFixMacro.QuoteText(vbNullString)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("quote")
Private Sub buildOutlookHeaderEnglish()
    On Error GoTo TestFail

    expectedResult = vbNullString & _
        "-----Original Message-----" & vbNewLine & _
        "From: Art Ross [mailto:art@example.com]" & vbNewLine & _
        "Sent: 2011-04-07 09:36" & vbNewLine & _
        "To: Adam Swift" & vbNewLine & _
        "Cc: Drift Over" & vbNewLine & _
        "Subject: Testing" & vbNewLine

    Assert.AreEqual expectedResult, QuoteFixMacro.BuildOutlookHeader(1033, "Art Ross", "art@example.com", "2011-04-07 09:36", "Adam Swift", "Drift Over", "Testing")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("quote")
Private Sub buildOutlookHeaderWithoutEmailAndCc()
    On Error GoTo TestFail

    expectedResult = vbNullString & _
        "-----Original Message-----" & vbNewLine & _
        "From: Art Ross" & vbNewLine & _
        "Sent: 2011-04-07 09:36" & vbNewLine & _
        "To: Adam Swift" & vbNewLine & _
        "Subject: Testing" & vbNewLine

    Assert.AreEqual expectedResult, QuoteFixMacro.BuildOutlookHeader(1033, "Art Ross", vbNullString, "2011-04-07 09:36", "Adam Swift", vbNullString, "Testing")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'The way FixMailText quotes a mail which is not a plain text mail
'Requires the settings of the "condense" tests
'@TestMethod("quote")
Private Sub quoteOfHtmlMailWithCondensedHeader()
    On Error GoTo TestFail

    Dim header As String
    header = QuoteFixMacro.BuildOutlookHeader(1033, "Art Ross", "art@example.com", "2011-04-07 09:36", "Adam Swift", vbNullString, "Testing")

    Dim originalText As String
    originalText = HtmlToPlainText("<p class=MsoNormal>Hi Adam,</p><p class=MsoNormal>&nbsp;</p><p class=MsoNormal>is it ok?</p>" & _
        "<blockquote><p class=MsoNormal>older text</p></blockquote>")

    outlookOutput = QuoteFixMacro.QuoteText(header & vbNewLine & originalText)
    expectedResult = vbNullString & _
        "Art Ross wrote on 2011-04-07 09:36:" & vbNewLine & _
        "> Hi Adam," & vbNewLine & _
        "> " & vbNewLine & _
        "> is it ok?" & vbNewLine & _
        ">> older text"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourceHtml)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("quote")
Private Sub buildOutlookHeaderGerman()
    On Error GoTo TestFail

    expectedResult = vbNullString & _
        "-----Urspr" & ChrW$(252) & "ngliche Nachricht-----" & vbNewLine & _
        "Von: Art Ross [mailto:art@example.com]" & vbNewLine & _
        "Gesendet: 2011-04-07 09:36" & vbNewLine & _
        "An: Adam Swift" & vbNewLine & _
        "Cc: Drift Over" & vbNewLine & _
        "Betreff: Testing" & vbNewLine

    'German (Germany) and German (Switzerland)
    Assert.AreEqual expectedResult, QuoteFixMacro.BuildOutlookHeader(1031, "Art Ross", "art@example.com", "2011-04-07 09:36", "Adam Swift", "Drift Over", "Testing")
    Assert.AreEqual expectedResult, QuoteFixMacro.BuildOutlookHeader(2055, "Art Ross", "art@example.com", "2011-04-07 09:36", "Adam Swift", "Drift Over", "Testing")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'The way FixMailText quotes a plain text mail, with a German header
'Requires the settings of the "condense" tests
'@TestMethod("quote")
Private Sub quoteOfPlainTextMailWithCondensedGermanHeader()
    On Error GoTo TestFail

    Dim header As String
    header = QuoteFixMacro.BuildOutlookHeader(1031, "Art Ross", "art@example.com", "2011-04-07 09:36", "Adam Swift", vbNullString, "Testing")

    outlookOutput = QuoteFixMacro.QuoteText(header & vbNewLine & _
        "Hi Adam," & vbNewLine & _
        vbNewLine & _
        "is it ok?" & vbNewLine & _
        "> older text" & vbNewLine)
    expectedResult = vbNullString & _
        "Art Ross wrote on 2011-04-07 09:36:" & vbNewLine & _
        "> Hi Adam," & vbNewLine & _
        "> " & vbNewLine & _
        "> is it ok?" & vbNewLine & _
        ">> older text"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourcePlainText)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'The following tests are on plain text mails which have been prefixed by QuoteText (SourcePlainText)

'@TestMethod("reformat")
Private Sub reformatPlainTextWrapsParagraphOfSenderAnew()
    On Error GoTo TestFail

    'the sender wrapped at 76 characters: the lines do not fit anymore once they are prefixed
    outlookOutput = vbNullString & _
        "> Hi Adam," & vbNewLine & _
        "> " & vbNewLine & _
        "> this is a paragraph which the sender wrapped at seventy-six characters, so" & vbNewLine & _
        "> that every full line is too long once the quote prefix has been added to" & vbNewLine & _
        "> it." & vbNewLine & _
        "> " & vbNewLine & _
        "> Thanks" & vbNewLine & _
        "> Art"
    expectedResult = vbNullString & _
        "> Hi Adam," & vbNewLine & _
        "> " & vbNewLine & _
        "> this is a paragraph which the sender wrapped at seventy-six characters," & vbNewLine & _
        "> so that every full line is too long once the quote prefix has been added" & vbNewLine & _
        "> to it." & vbNewLine & _
        "> " & vbNewLine & _
        "> Thanks" & vbNewLine & _
        "> Art"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourcePlainText)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("reformat")
Private Sub reformatPlainTextWrapsLongLineOnItsOwn()
    On Error GoTo TestFail

    'a line being longer than any line wrapped by a sender is a paragraph
    outlookOutput = vbNullString & _
        "> Hi Adam," & vbNewLine & _
        "> this is a very long paragraph line that was never wrapped by the sender because the mail program of the sender does not wrap." & vbNewLine & _
        "> Thanks" & vbNewLine & _
        "> Art"
    expectedResult = vbNullString & _
        "> Hi Adam," & vbNewLine & _
        "> this is a very long paragraph line that was never wrapped by the sender" & vbNewLine & _
        "> because the mail program of the sender does not wrap." & vbNewLine & _
        "> Thanks" & vbNewLine & _
        "> Art"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourcePlainText)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("reformat")
Private Sub reformatPlainTextKeepsShortLines()
    On Error GoTo TestFail

    'in text prefixed by Outlook, "Thanks" would be regarded as the broken end of the line above
    outlookOutput = vbNullString & _
        "> this line is nearly as long as the limit allows, but it still fits." & vbNewLine & _
        "> Thanks" & vbNewLine & _
        "> Art"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourcePlainText)

    Assert.AreEqual outlookOutput, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("reformat")
Private Sub reformatPlainTextRepairsWrapsOfOlderReplies()
    On Error GoTo TestFail

    'the quote of the second level has been broken by Outlook in an older reply
    outlookOutput = vbNullString & _
        "> > I have a Win 2k3 SBS and I want to replicate the users into my" & vbNewLine & _
        "> OpenLDAP" & vbNewLine & _
        "> > 2.4.11." & vbNewLine & _
        "> " & vbNewLine & _
        "> This is not possible."
    expectedResult = vbNullString & _
        ">> I have a Win 2k3 SBS and I want to replicate the users into my OpenLDAP" & vbNewLine & _
        ">> 2.4.11." & vbNewLine & _
        "> " & vbNewLine & _
        "> This is not possible."

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourcePlainText)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'DetectLanguage returns 7 for German, 9 for English, and 0 if there is no clear result

'@TestMethod("language")
Private Sub detectLanguageGerman()
    On Error GoTo TestFail

    outlookOutput = vbNullString & _
        "Hallo Adam," & vbNewLine & _
        vbNewLine & _
        "kannst du mir bitte die Folien schicken? Ich bin bis 11 Uhr in der Vorlesung und danach im Raum." & vbNewLine & _
        vbNewLine & _
        "Danke und viele Gr" & ChrW$(252) & ChrW$(223) & "e" & vbNewLine & _
        "Art"

    Assert.AreEqual 7&, QuoteFixMacro.DetectLanguage(outlookOutput)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("language")
Private Sub detectLanguageEnglish()
    On Error GoTo TestFail

    outlookOutput = vbNullString & _
        "Hello Adam," & vbNewLine & _
        vbNewLine & _
        "could you please send me the slides? I am in the lecture until 11 and in the room after that." & vbNewLine & _
        vbNewLine & _
        "Thanks and regards" & vbNewLine & _
        "Art"

    Assert.AreEqual 9&, QuoteFixMacro.DetectLanguage(outlookOutput)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("language")
Private Sub detectLanguageRegardsNewestPartOnly()
    On Error GoTo TestFail

    'a German answer to an English mail, which follows behind a header
    outlookOutput = vbNullString & _
        "Hallo Art," & vbNewLine & _
        "das ist kein Problem, wir sehen uns morgen." & vbNewLine & _
        vbNewLine & _
        "From: Art Ross <art@example.com>" & vbNewLine & _
        "Sent: Thursday, April 7, 2011 9:36 AM" & vbNewLine & _
        "Subject: Testing" & vbNewLine & _
        vbNewLine & _
        "Hello Adam, could you please send me the slides? I am in the lecture and have no access to them. Thanks and regards"
    Assert.AreEqual 7&, QuoteFixMacro.DetectLanguage(outlookOutput)

    'an English answer to a German mail, which is quoted
    outlookOutput = vbNullString & _
        "Hello Art," & vbNewLine & _
        "this is not a problem, see you tomorrow." & vbNewLine & _
        vbNewLine & _
        "> Hallo Adam, kannst du mir bitte die Folien schicken? Ich bin in der Vorlesung und habe sie nicht. Danke und viele Gr" & ChrW$(252) & ChrW$(223) & "e"
    Assert.AreEqual 9&, QuoteFixMacro.DetectLanguage(outlookOutput)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("language")
Private Sub detectLanguageNoClearResult()
    On Error GoTo TestFail

    Assert.AreEqual 0&, QuoteFixMacro.DetectLanguage(vbNullString)
    Assert.AreEqual 0&, QuoteFixMacro.DetectLanguage("Danke!")
    Assert.AreEqual 0&, QuoteFixMacro.DetectLanguage("ok")
    'as many German as English words
    Assert.AreEqual 0&, QuoteFixMacro.DetectLanguage("Hallo und danke, hello and thanks")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("reformat")
Private Sub reformatPlainTextKeepsAnswersBetweenQuotes()
    On Error GoTo TestFail

    'the answers are no broken wraps: they would have fit into the line above
    outlookOutput = vbNullString & _
        "> > > Should we meet on Monday?" & vbNewLine & _
        "> > Tuesday fits better." & vbNewLine & _
        "> > > At ten?" & vbNewLine & _
        "> > Yes." & vbNewLine & _
        "> Fine with me."
    expectedResult = vbNullString & _
        ">>> Should we meet on Monday?" & vbNewLine & _
        ">> Tuesday fits better." & vbNewLine & _
        ">>> At ten?" & vbNewLine & _
        ">> Yes." & vbNewLine & _
        "> Fine with me."

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourcePlainText)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'The following tests are on headers of older mails within text prefixed by QuoteText (SourcePlainText, SourceHtml).
'The headers do not need the line "-----Original Message-----": a mail written in HTML does not have it.

'@TestMethod("condense")
Private Sub condenseHeadersOfThreadStartDeeperLevels()
    On Error GoTo TestFail

    'two older mails, as Outlook writes them in HTML (German and English). The date of the first one cannot be parsed.
    outlookOutput = vbNullString & _
        "> Hi Art," & vbNewLine & _
        "> fine with me." & vbNewLine & _
        "> " & vbNewLine & _
        "> Von: Swift, Adam <adam@example.com>" & vbNewLine & _
        "> Datum: Dienstag, 29. September 2026 um 19:40" & vbNewLine & _
        "> An: Art Ross <art@example.com>; Drift Over <drift@example.com>; Zimmer, Zoe" & vbNewLine & _
        "> <zoe@example.com>" & vbNewLine & _
        "> Betreff: RE: Testing" & vbNewLine & _
        "> " & vbNewLine & _
        "> Hallo," & vbNewLine & _
        "> " & vbNewLine & _
        "> is Tuesday ok?" & vbNewLine & _
        "> " & vbNewLine & _
        "> From: Art Ross <art@example.com>" & vbNewLine & _
        "> Sent: 2026-09-29 19:01" & vbNewLine & _
        "> To: Adam Swift <adam@example.com>" & vbNewLine & _
        "> Subject: Testing" & vbNewLine & _
        "> " & vbNewLine & _
        "> Hi Adam," & vbNewLine & _
        "> " & vbNewLine & _
        "> when do we meet?" & vbNewLine & _
        "> > an older quote"
    expectedResult = vbNullString & _
        "> Hi Art," & vbNewLine & _
        "> fine with me." & vbNewLine & _
        "> " & vbNewLine & _
        "> Adam Swift wrote on 29. September 2026 um 19:40:" & vbNewLine & _
        ">> Hallo," & vbNewLine & _
        ">> " & vbNewLine & _
        ">> is Tuesday ok?" & vbNewLine & _
        ">> " & vbNewLine & _
        ">> Art Ross wrote on 2026-09-29 19:01:" & vbNewLine & _
        ">>> Hi Adam," & vbNewLine & _
        ">>> " & vbNewLine & _
        ">>> when do we meet?" & vbNewLine & _
        ">>>> an older quote"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourceHtml)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("condense")
Private Sub condenseHeaderQuotedDeeperAlready()
    On Error GoTo TestFail

    'the older mail was prefixed by Outlook already: its condensed line gets one level less, the text keeps its level
    outlookOutput = vbNullString & _
        "> answer" & vbNewLine & _
        "> " & vbNewLine & _
        "> > -----Original Message-----" & vbNewLine & _
        "> > From: Art Ross [mailto:art@example.com]" & vbNewLine & _
        "> > Sent: 2011-04-07 09:36" & vbNewLine & _
        "> > To: Adam Swift" & vbNewLine & _
        "> > Subject: Testing" & vbNewLine & _
        "> > " & vbNewLine & _
        "> > Hi Adam,"
    expectedResult = vbNullString & _
        "> answer" & vbNewLine & _
        "> " & vbNewLine & _
        "> Art Ross wrote on 2011-04-07 09:36:" & vbNewLine & _
        ">> Hi Adam,"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourcePlainText)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("condense")
Private Sub condenseHeadersLeavesOtherLabelsAlone()
    On Error GoTo TestFail

    'labelled lines which are no header: the first label is not "From", or there is no date
    outlookOutput = vbNullString & _
        "> Note: the meeting is at ten." & vbNewLine & _
        "> Date: 2026-09-29" & vbNewLine & _
        "> Room: Q107" & vbNewLine & _
        "> " & vbNewLine & _
        "> From: Art Ross" & vbNewLine & _
        "> To: Adam Swift" & vbNewLine & _
        "> Subject: no date in this one" & vbNewLine & _
        "> " & vbNewLine & _
        "> From: the beginning, it was clear."

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourceHtml)

    Assert.AreEqual outlookOutput, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'Required settings:
'
'NUM_QUOTE_COLORS = 6
'CONDENSED_HEADER_FORMAT = "%SN wrote on %D:"

'@TestMethod("color")
Private Sub coloredHtmlHasOneParagraphPerLineAndColorsPerLevel()
    On Error GoTo TestFail

    'without condensed headers, the authors are unknown: each level gets a color
    outlookOutput = vbNullString & _
        "Hi <you> & ""me""" & vbNewLine & _
        "> level one" & vbNewLine & _
        ">> level two" & vbNewLine & _
        vbNewLine & _
        ">>>>>>> level seven has the first color again" & vbNewLine & _
        "  * two spaces  in front and within"
    expectedResult = vbNullString & _
        "<html><body><div style=""font-family:Consolas,'Courier New',monospace;font-size:10pt"">" & vbNewLine & _
        "<p style=""margin:0"">Hi &lt;you&gt; &amp; &quot;me&quot;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#1F6FB2"">&gt; level one</span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#2E8B57"">&gt;&gt; level two</span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0"">&nbsp;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#1F6FB2"">&gt;&gt;&gt;&gt;&gt;&gt;&gt; level seven has the first color again</span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0"">&nbsp; * two spaces &nbsp;in front and within</p>" & vbNewLine & _
        "</div></body></html>"

    Assert.AreEqual expectedResult, QuoteFixMacro.TextToColoredHtml(outlookOutput, vbNullString)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("color")
Private Sub coloredHtmlColorsPerAuthor()
    On Error GoTo TestFail

    'the condensed headers tell the authors: each author gets a color, the header is a heading in that color
    outlookOutput = vbNullString & _
        "Art Ross wrote on 2011-04-07 09:36:" & vbNewLine & _
        "> Hi Adam," & vbNewLine & _
        "> " & vbNewLine & _
        "> Adam Swift wrote on 2011-04-06 15:12:" & vbNewLine & _
        ">> is it ok?" & vbNewLine & _
        ">> " & vbNewLine & _
        ">> Art Ross wrote on 2011-04-05 10:00:" & vbNewLine & _
        ">>> when do we meet?"
    expectedResult = vbNullString & _
        "<html><body><div style=""font-family:Consolas,'Courier New',monospace;font-size:10pt"">" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#1F6FB2;font-weight:bold"">Art Ross wrote on 2011-04-07 09:36:</span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#1F6FB2"">&gt; Hi Adam,</span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#1F6FB2"">&gt; </span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#2E8B57;font-weight:bold"">&gt; Adam Swift wrote on 2011-04-06 15:12:</span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#2E8B57"">&gt;&gt; is it ok?</span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#2E8B57"">&gt;&gt; </span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#1F6FB2;font-weight:bold"">&gt;&gt; Art Ross wrote on 2011-04-05 10:00:</span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#1F6FB2"">&gt;&gt;&gt; when do we meet?</span>&#8203;</p>" & vbNewLine & _
        "</div></body></html>"

    Assert.AreEqual expectedResult, QuoteFixMacro.TextToColoredHtml(outlookOutput, vbNullString)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("color")
Private Sub coloredHtmlOwnTextIsGray()
    On Error GoTo TestFail

    'the user (Adam Swift) gets the neutral color, the others get the palette
    outlookOutput = vbNullString & _
        "Art Ross wrote on 2011-04-07 09:36:" & vbNewLine & _
        "> Hi Adam," & vbNewLine & _
        "> Adam Swift wrote on 2011-04-06 15:12:" & vbNewLine & _
        ">> is it ok?" & vbNewLine & _
        ">> Drift Over wrote on 2011-04-05 10:00:" & vbNewLine & _
        ">>> when do we meet?"
    expectedResult = vbNullString & _
        "<html><body><div style=""font-family:Consolas,'Courier New',monospace;font-size:10pt"">" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#1F6FB2;font-weight:bold"">Art Ross wrote on 2011-04-07 09:36:</span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#1F6FB2"">&gt; Hi Adam,</span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#555555;font-weight:bold"">&gt; Adam Swift wrote on 2011-04-06 15:12:</span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#555555"">&gt;&gt; is it ok?</span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#2E8B57;font-weight:bold"">&gt;&gt; Drift Over wrote on 2011-04-05 10:00:</span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#2E8B57"">&gt;&gt;&gt; when do we meet?</span>&#8203;</p>" & vbNewLine & _
        "</div></body></html>"

    Assert.AreEqual expectedResult, QuoteFixMacro.TextToColoredHtml(outlookOutput, "Adam Swift")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("color")
Private Sub allRecipientsListedByAddressOrDomain()
    On Error GoTo TestFail

    Assert.AreEqual True, QuoteFixMacro.AllRecipientsListed("a@example.org;b@example.org", "@example.org")
    Assert.AreEqual False, QuoteFixMacro.AllRecipientsListed("a@example.org;x@other.org", "@example.org")
    Assert.AreEqual True, QuoteFixMacro.AllRecipientsListed("Boss@Example.org", "boss@example.org")
    Assert.AreEqual True, QuoteFixMacro.AllRecipientsListed("a@example.org;boss@example.org", "@example.org; boss@example.org")
    Assert.AreEqual False, QuoteFixMacro.AllRecipientsListed("a@example.org", vbNullString)
    Assert.AreEqual False, QuoteFixMacro.AllRecipientsListed(vbNullString, "@example.org")
    'the domain has to match completely
    Assert.AreEqual False, QuoteFixMacro.AllRecipientsListed("a@noexample.org", "@example.org")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("color")
Private Sub coloredHtmlShowsEmphasis()
    On Error GoTo TestFail

    'the text between the markers is bold or underlined, also over two lines; the markers are left out
    outlookOutput = vbNullString & _
        "> werden _neuartige, auch" & vbNewLine & _
        "> risikoreiche Ansaetze_, deren" & vbNewLine & _
        "*own* text, 5 * 3 and snake_case_name"
    expectedResult = vbNullString & _
        "<html><body><div style=""font-family:Consolas,'Courier New',monospace;font-size:10pt"">" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#1F6FB2"">&gt; werden <u>neuartige, auch</u></span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0""><span style=""color:#1F6FB2"">&gt; <u>risikoreiche Ansaetze</u>, deren</span>&#8203;</p>" & vbNewLine & _
        "<p style=""margin:0""><b>own</b> text, 5 * 3 and snake_case_name</p>" & vbNewLine & _
        "</div></body></html>"

    Assert.AreEqual expectedResult, QuoteFixMacro.TextToColoredHtml(outlookOutput, vbNullString)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("color")
Private Sub coloredHtmlIsPutInFrontOfOutlooksSignature()
    On Error GoTo TestFail

    Dim signatureHtml As String
    signatureHtml = "<html><head><style>p.MsoNormal {margin:0}</style></head>" & vbNewLine & _
        "<body lang=DE link=""#0563C1""><div class=WordSection1><p class=MsoNormal>Regards<img src=""cid:image001.png@01DC0000.00000000""></p></div></body></html>"
    expectedResult = "<html><head><style>p.MsoNormal {margin:0}</style></head>" & vbNewLine & _
        "<body lang=DE link=""#0563C1""><div style=""margin:0"">text</div><div class=WordSection1><p class=MsoNormal>Regards<img src=""cid:image001.png@01DC0000.00000000""></p></div></body></html>"

    Assert.AreEqual expectedResult, QuoteFixMacro.InsertColoredHtml("<html><body><div style=""margin:0"">text</div></body></html>", signatureHtml)

    'without a body, the colored HTML is used alone
    Assert.AreEqual "<html><body>x</body></html>", QuoteFixMacro.InsertColoredHtml("<html><body>x</body></html>", "Regards")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("condense")
Private Sub condenseAttributionOfTicketSystem()
    On Error GoTo TestFail

    'a ticket system quotes the older mail below a one-line attribution, without prefixing it
    outlookOutput = vbNullString & _
        "> Hallo Herr Ross," & vbNewLine & _
        "> " & vbNewLine & _
        "> die Kapazitaet ist erhoeht." & vbNewLine & _
        "> " & vbNewLine & _
        "> 2026-10-01 16:15 - Art Ross schrieb: " & vbNewLine & _
        "> " & vbNewLine & _
        "> Hallo Adam," & vbNewLine & _
        "> " & vbNewLine & _
        "> bitte pruefen." & vbNewLine & _
        "> " & vbNewLine & _
        "> Von: Swift, Adam <adam@example.org>" & vbNewLine & _
        "> Datum: Donnerstag, 1. Oktober 2026 um 16:02" & vbNewLine & _
        "> An: Ross, Art <art@example.org>" & vbNewLine & _
        "> Betreff: Testing" & vbNewLine & _
        "> " & vbNewLine & _
        "> Hallo Art,"
    expectedResult = vbNullString & _
        "> Hallo Herr Ross," & vbNewLine & _
        "> " & vbNewLine & _
        "> die Kapazitaet ist erhoeht." & vbNewLine & _
        "> " & vbNewLine & _
        "> Art Ross wrote on 2026-10-01 16:15:" & vbNewLine & _
        ">> Hallo Adam," & vbNewLine & _
        ">> " & vbNewLine & _
        ">> bitte pruefen." & vbNewLine & _
        ">> " & vbNewLine & _
        ">> Adam Swift wrote on 1. Oktober 2026 um 16:02:" & vbNewLine & _
        ">>> Hallo Art,"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourceHtml)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("condense")
Private Sub condenseAttributionAboveQuoteIsKept()
    On Error GoTo TestFail

    'the attribution of a mail program whose quote is prefixed already stays as it is
    outlookOutput = vbNullString & _
        "> Fine with me." & vbNewLine & _
        "> " & vbNewLine & _
        "> 2026-10-01 16:15 - Art Ross wrote:" & vbNewLine & _
        "> > Hallo Adam," & vbNewLine & _
        "> > bitte pruefen."
    expectedResult = vbNullString & _
        "> Fine with me." & vbNewLine & _
        "> " & vbNewLine & _
        "> 2026-10-01 16:15 - Art Ross wrote:" & vbNewLine & _
        ">> Hallo Adam," & vbNewLine & _
        ">> bitte pruefen."

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourceHtml)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("condense")
Private Sub condenseForwardedMessageWithDateBehindSubject()
    On Error GoTo TestFail

    'a ticket system forwards a mail with a marker of four dashes and the date behind the subject
    outlookOutput = vbNullString & _
        "> please have a look." & vbNewLine & _
        "> ---- Weitergeleitete Nachricht von ""Ross, Art"" <art@example.com> ---" & vbNewLine & _
        "> " & vbNewLine & _
        "> Von: ""Ross, Art"" <art@example.com>" & vbNewLine & _
        "> An: ""support@example.com"" <support@example.com>" & vbNewLine & _
        "> Betreff: Problem" & vbNewLine & _
        "> Datum: 2026-09-30 15:10:02" & vbNewLine & _
        "> " & vbNewLine & _
        "> Guten Tag,"
    expectedResult = vbNullString & _
        "> please have a look." & vbNewLine & _
        "> Art Ross wrote on 2026-09-30 15:10:" & vbNewLine & _
        ">> Guten Tag,"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput, SourceHtml)

    Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
