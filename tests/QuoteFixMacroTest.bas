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
    'This method runs before every test in the module..
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
