Attribute VB_Name = "QuoteFixMacro"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private originalName As String
Private senderName As String
Private firstName As String
Private lastName As String

Private outlookOutput As String
Private expectedResult As String

'@ModuleInitialize
Private Sub ModuleInitialize()
'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module.

    'Currently required for reformat only
    Call QuoteFixMacro.LoadConfiguration
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("getNamesOutOfString")
Private Sub FirstnameLastname()
    On Error GoTo TestFail

    originalName = "Firstname Lastname"

    Call getNamesOutOfString(originalName, senderName, firstName, lastName)

    Assert.AreEqual "Firstname Lastname", senderName
    Assert.AreEqual "Firstname", firstName
    Assert.AreEqual "Lastname", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("getNamesOutOfString")
Private Sub LASTNAMEfirstname()
    On Error GoTo TestFail

    originalName = "Lastname, Firstname"

    Call getNamesOutOfString(originalName, senderName, firstName, lastName)

    Assert.AreEqual "Firstname Lastname", senderName
    Assert.AreEqual "Firstname", firstName
    Assert.AreEqual "Lastname", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("getNamesOutOfString")
Private Sub FirstnameVanLastname()
    On Error GoTo TestFail

    originalName = "Firstname van Lastname"

    Call getNamesOutOfString(originalName, senderName, firstName, lastName)

    Assert.AreEqual "Firstname van Lastname", senderName
    Assert.AreEqual "Firstname", firstName
    Assert.AreEqual "van Lastname", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("IsUpperCaseWord")
Private Sub IsUpperCaseWordTests()
    On Error GoTo TestFail

    Assert.AreEqual False, IsUpperCaseWord("van")
    Assert.AreEqual False, IsUpperCaseWord("Lastname")
    Assert.AreEqual False, IsUpperCaseWord("LastName")
    Assert.AreEqual True, IsUpperCaseWord("LASTNAME")
    Assert.AreEqual False, IsUpperCaseWord("")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'@TestMethod("getNamesOutOfString")
Private Sub FirstMiddleLast()
    On Error GoTo TestFail

    originalName = "First Middle Last"

    Call getNamesOutOfString(originalName, senderName, firstName, lastName)

    'The function cannot know where "Middle" belong to.
    'Safe fallback: put it as first name
    Assert.AreEqual "First Middle Last", firstName
    Assert.AreEqual "", lastName
    Assert.AreEqual "First Middle Last", senderName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("getNamesOutOfString")
Private Sub firstAtExampleCom()
    On Error GoTo TestFail

    originalName = "first@example.com"

    Call getNamesOutOfString(originalName, senderName, firstName, lastName)

    Assert.AreEqual "First", senderName
    Assert.AreEqual "First", firstName
    Assert.AreEqual "", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("getNamesOutOfString")
Private Sub firstDotLastAtExampleCom()
    On Error GoTo TestFail

    originalName = "first.last@example.com"

    Call getNamesOutOfString(originalName, senderName, firstName, lastName)

    Assert.AreEqual "First Last", senderName
    Assert.AreEqual "First", firstName
    Assert.AreEqual "Last", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("getNamesOutOfString")
Private Sub DrFirstLast()
    On Error GoTo TestFail

    originalName = "Dr. First Last"

    Call getNamesOutOfString(originalName, senderName, firstName, lastName)

    Assert.AreEqual "First Last", senderName
    Assert.AreEqual "First", firstName
    Assert.AreEqual "Last", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("getNamesOutOfString")
Private Sub UppercaseLASTNAMEfirstname()
    On Error GoTo TestFail

    originalName = "LAST first"

    Call getNamesOutOfString(originalName, senderName, firstName, lastName)

    Assert.AreEqual "First Last", senderName
    Assert.AreEqual "First", firstName
    Assert.AreEqual "Last", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("getNamesOutOfString")
Private Sub UppercaseFIRSTNAMEandLASTNAME()
    On Error GoTo TestFail

    originalName = "FIRST LAST"

    Call getNamesOutOfString(originalName, senderName, firstName, lastName)

    Assert.AreEqual "First Last", senderName
    Assert.AreEqual "First", firstName
    Assert.AreEqual "Last", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("getNamesOutOfString")
Private Sub FirstnameWithDashCorrectlyCased()
    On Error GoTo TestFail

    originalName = "First-First Last"

    Call getNamesOutOfString(originalName, senderName, firstName, lastName)

    Assert.AreEqual "First-First Last", senderName
    Assert.AreEqual "First-First", firstName
    Assert.AreEqual "Last", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("getNamesOutOfString")
Private Sub FirstnameLastnameDepartment()
    On Error GoTo TestFail

    originalName = "First Last DEPT DEPT"

    Call getNamesOutOfString(originalName, senderName, firstName, lastName)

    Assert.AreEqual "First Last", senderName
    Assert.AreEqual "First", firstName
    Assert.AreEqual "Last", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("removeDepartmentName")
Private Sub FirstnameLastnameDepartmentFunction()
    On Error GoTo TestFail

    Dim result As String

    result = removeDepartment("First Last DEPT DEPT")

    Assert.AreEqual "First Last", result

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'Required settings:
'
'USE_COLORIZER unset
'INCLUDE_QUOTES_TO_LEVEL = -1
'LINE_WRAP_AFTER = 75

'@TestMethod("reformat")
Private Sub reformatTest1()
    On Error GoTo TestFail

    outlookOutput = "" + _
        "> >>" + vbNewLine + _
        "> >> I have a Win 2k3 SBS and I want to replicate the users into my" + vbNewLine + _
        "> OpenLDAP" + vbNewLine + _
        "> >> 2.4.11." + vbNewLine + _
        "> >" + vbNewLine + _
        "> > This is not possible. You could however implement your own sync" + vbNewLine + _
        "> process" + vbNewLine + _
        "> > in your favourite scripting/programming language." + vbNewLine + _
        "> " + vbNewLine + _
        "> Actually we have done some preliminary work..."
    expectedResult = "" + _
        ">>> " + vbNewLine + _
        ">>> I have a Win 2k3 SBS and I want to replicate the users into my" + vbNewLine + _
        ">>> OpenLDAP 2.4.11." + vbNewLine + _
        ">> " + vbNewLine + _
        ">> This is not possible. You could however implement your own sync process" + vbNewLine + _
        ">> in your favourite scripting/programming language." + vbNewLine + _
        "> " + vbNewLine + _
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

    outlookOutput = "" + _
        "> Moin," + vbNewLine + _
        "> " + vbNewLine + _
        "> Kurzanleitung """"Deckel öffnen"""":" + vbNewLine + _
        "> 1. Unten rechts die Kunststoff-Abdeckung mit einem Schraubendreher" + vbNewLine + _
        "> nach rechts schieben." + vbNewLine + _
        "> 2. Das Blech nach links schieben." + vbNewLine + _
        "> 3. Kreuzschlitzschraube lösen." + vbNewLine + _
        "> " + vbNewLine + _
        "> " + vbNewLine + _
        "> Mit freundlichen Grüßen" + vbNewLine + _
        "> " + vbNewLine + _
        "> company" + vbNewLine + _
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

    outlookOutput = "" + _
        "> Hallo Jon, ich hatte mal von xxxxxx ein Anti-Virus Programm, aber ich" + vbNewLine + _
        "> habe" + vbNewLine + _
        "> so viele Spams trotzdem erhalten, dass ich das nicht mehr abonniert" + vbNewLine + _
        "> habe." + vbNewLine + _
        "> xxx xxxxx? Haste eine Lösung für mein Virenprogramm, kann ich was" + vbNewLine + _
        "> runterladen?" + vbNewLine + _
        "> Lieben Gruß Jane"
    expectedResult = "" + _
        "> Hallo Jon, ich hatte mal von xxxxxx ein Anti-Virus Programm, aber ich" + vbNewLine + _
        "> habe so viele Spams trotzdem erhalten, dass ich das nicht mehr abonniert" + vbNewLine + _
        "> habe. xxx xxxxx? Haste eine Lösung für mein Virenprogramm, kann" + vbNewLine + _
        "> ich was runterladen?" + vbNewLine + _
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
