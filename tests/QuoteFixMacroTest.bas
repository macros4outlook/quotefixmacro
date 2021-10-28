Attribute VB_Name = "QuoteFixMacroTest"

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

Private outlookOutput As String
Private expectedResult As String

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
        "> Kurzanleitung """"Deckel ˆffnen"""":" & vbNewLine & _
        "> 1. Unten rechts die Kunststoff-Abdeckung mit einem Schraubendreher" & vbNewLine & _
        "> nach rechts schieben." & vbNewLine & _
        "> 2. Das Blech nach links schieben." & vbNewLine & _
        "> 3. Kreuzschlitzschraube lˆsen." & vbNewLine & _
        "> " & vbNewLine & _
        "> " & vbNewLine & _
        "> Mit freundlichen Gr¸ﬂen" & vbNewLine & _
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
        "> xxx xxxxx? Haste eine Lˆsung f¸r mein Virenprogramm, kann ich was" & vbNewLine & _
        "> runterladen?" & vbNewLine & _
        "> Lieben Gruﬂ Jane"
    expectedResult = vbNullString & _
        "> Hallo Jon, ich hatte mal von xxxxxx ein Anti-Virus Programm, aber ich" & vbNewLine & _
        "> habe so viele Spams trotzdem erhalten, dass ich das nicht mehr abonniert" & vbNewLine & _
        "> habe. xxx xxxxx? Haste eine Lˆsung f¸r mein Virenprogramm, kann" & vbNewLine & _
        "> ich was runterladen?" & vbNewLine & _
        "> Lieben Gruﬂ Jane"

    Dim processedText As String
    processedText = QuoteFixMacro.ReFormatText(outlookOutput)

    'TODO: Keeping the greeting unformatted is currently not implemented
    'Assert.AreEqual expectedResult, processedText

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


