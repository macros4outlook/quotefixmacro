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
