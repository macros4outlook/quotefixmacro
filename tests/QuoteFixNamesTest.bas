Attribute VB_Name = "QuoteFixNamesTest"

Option Explicit
Option Private Module

'@TestModule
'@Folder("QuoteFixMacro.Tests")

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
    Private Fakes As Object
#Else
    Private Assert As Rubberduck.AssertClass
    Private Fakes As Rubberduck.FakesProvider
#End If

Private originalName As String
Private senderName As String
Private firstName As String
Private lastName As String

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
#If LateBind Then
        Set Assert = CreateObject("Rubberduck.AssertClass")
        Set Fakes = CreateObject("Rubberduck.FakesProvider")
#Else
        Set Assert = New Rubberduck.AssertClass
        Set Fakes = New Rubberduck.FakesProvider
#End If
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

    getNamesOutOfString originalName, senderName, firstName, lastName

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

    getNamesOutOfString originalName, senderName, firstName, lastName

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

    getNamesOutOfString originalName, senderName, firstName, lastName

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
    Assert.AreEqual False, IsUpperCaseWord(vbNullString)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'@TestMethod("getNamesOutOfString")
Private Sub FirstMiddleLast()
    On Error GoTo TestFail

    originalName = "First Middle Last"

    getNamesOutOfString originalName, senderName, firstName, lastName

    'The function cannot know where "Middle" belong to.
    'Safe fallback: put it as first name
    Assert.AreEqual "First Middle Last", firstName
    Assert.AreEqual vbNullString, lastName
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

    getNamesOutOfString originalName, senderName, firstName, lastName

    Assert.AreEqual "First", senderName
    Assert.AreEqual "First", firstName
    Assert.AreEqual vbNullString, lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("getNamesOutOfString")
Private Sub firstDotLastAtExampleCom()
    On Error GoTo TestFail

    originalName = "first.last@example.com"

    getNamesOutOfString originalName, senderName, firstName, lastName

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

    getNamesOutOfString originalName, senderName, firstName, lastName

    Assert.AreEqual "Dr. First Last", senderName
    Assert.AreEqual "First", firstName
    Assert.AreEqual "Dr. Last", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("getNamesOutOfString")
Private Sub UppercaseLASTNAMEfirstname()
    On Error GoTo TestFail

    originalName = "LAST first"

    getNamesOutOfString originalName, senderName, firstName, lastName

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

    getNamesOutOfString originalName, senderName, firstName, lastName

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

    getNamesOutOfString originalName, senderName, firstName, lastName

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

    getNamesOutOfString originalName, senderName, firstName, lastName

    Assert.AreEqual "First Last", senderName
    Assert.AreEqual "First", firstName
    Assert.AreEqual "Last", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("getNamesOutOfString")
Private Sub LastnameFirstnameDrAllCommaSeparated()
    On Error GoTo TestFail

    originalName = "Last, First, Dr."

    getNamesOutOfString originalName, senderName, firstName, lastName, "First.Last3@example.com"

    Assert.AreEqual "Dr. First Last", senderName
    Assert.AreEqual "First", firstName
    Assert.AreEqual "Dr. Last", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("getNamesOutOfString")
Private Sub FirstnameLastnameDrDepartmentEmailWithNumberReverse()
    On Error GoTo TestFail

    originalName = "Last First Dr. DEP DEP2"

    getNamesOutOfString originalName, senderName, firstName, lastName, "First.Last3@example.com"

    Assert.AreEqual "Dr. First Last", senderName
    Assert.AreEqual "First", firstName
    Assert.AreEqual "Dr. Last", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("getNamesOutOfString")
Private Sub FirstnameLastnameDepartmentEmailWithNumberReverse()
    On Error GoTo TestFail

    originalName = "Last First DEP DEP2"

    getNamesOutOfString originalName, senderName, firstName, lastName, "First.Last3@example.com"

    Assert.AreEqual "First Last", senderName
    Assert.AreEqual "First", firstName
    Assert.AreEqual "Last", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("getNamesOutOfString")
Private Sub FirstnameLastnameDepartmentEmailReverse()
    On Error GoTo TestFail

    originalName = "Last First (xy/z)"

    getNamesOutOfString originalName, senderName, firstName, lastName, "first.last@example.com"

    Assert.AreEqual "First Last", senderName
    Assert.AreEqual "First", firstName
    Assert.AreEqual "Last", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("getNamesOutOfString")
Private Sub UppercaseLastnameFirstnameReversedEmail()
    On Error GoTo TestFail

    originalName = "last first"

    getNamesOutOfString originalName, senderName, firstName, lastName, "first.last@example.com"

    Assert.AreEqual "First Last", senderName
    Assert.AreEqual "First", firstName
    Assert.AreEqual "Last", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("getNamesOutOfString")
Private Sub LowerCaseNamesDEPReversedEmail()
    On Error GoTo TestFail

    originalName = "last first DEP DEP"

    getNamesOutOfString originalName, senderName, firstName, lastName, "first.last@example.com"

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

'@TestMethod("removeDepartmentName")
Private Sub LowerCaseFirstnameLastnameDepartmentFunction()
    On Error GoTo TestFail

    Dim result As String

    result = removeDepartment("first last DEPT DEPT")

    Assert.AreEqual "first last", result

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("getNamesOutOfEmail")
Private Sub getNamesOutOfEmailNormalCase()
    On Error GoTo TestFail

    getFirstNameLastNameOutOfEmail "firstname.lastname@example.com", firstName, lastName

    Assert.AreEqual "firstname", firstName
    Assert.AreEqual "lastname", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("getNamesOutOfEmail")
Private Sub getNamesOutOfEmailTwoDots()
    On Error GoTo TestFail

    getFirstNameLastNameOutOfEmail "firstname.lastname.something@example.com", firstName, lastName

    Assert.AreEqual "firstname.lastname.something", firstName
    Assert.AreEqual vbNullString, lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("getNamesOutOfEmail")
Private Sub getNamesOutOfEmailNoDot()
    On Error GoTo TestFail

    getFirstNameLastNameOutOfEmail "thing@example.com", firstName, lastName

    Assert.AreEqual "thing", firstName
    Assert.AreEqual vbNullString, lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("getNamesOutOfEmail")
Private Sub getNamesOutOfEmailNumberAtEnd()
    On Error GoTo TestFail

    getFirstNameLastNameOutOfEmail "First.Last3@example.com", firstName, lastName

    Assert.AreEqual "First", firstName
    Assert.AreEqual "Last", lastName

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

