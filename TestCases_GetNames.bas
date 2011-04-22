Attribute VB_Name = "TestCases_GetNames"
'$Id$
'
'These test cases part of the macros4outlook project
'see http://sourceforge.net/projects/macros4outlook/ for more information
'
'For more information on Outlook see http://www.microsoft.com/outlook
'Outlook is (C) by Microsoft

'****************************************************************************
'License:
'
'QuoteFixMacro testcases for getNames sub
'  copyright 2011 Oliver Kopp and Lars Monsees. All rights reserved.
'
'
'Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
'
'   1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
'   2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
'   3. The name of the author may not be used to endorse or promote products derived from this software without specific prior written permission.
'
'THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'****************************************************************************

'Changelog
'
'$Revision$ - not released

Option Explicit

Private Type typeTestCase
    originalName As String
    ExpectedFirstName As String
    ExpectedSenderName As String
End Type

Private mTestCases() As typeTestCase
Private Sub addTestCaseToArray(ByRef testcase As typeTestCase)
    
    ReDim Preserve mTestCases(UBound(mTestCases) + 1)
    mTestCases(UBound(mTestCases)) = testcase
End Sub

'Puts all test cases into the passed array.
Private Sub initTestCases()
    
    Dim testcase As typeTestCase
    
    ReDim mTestCases(0)
    'dummy - will never be called as testcases are called from 1 on
    '  Alternative: Use "Option Base 1" and add first testcase by a direct assignment and not by addTestCaseToArray
    mTestCases(0) = testcase
    
    testcase.originalName = "First Last"
    testcase.ExpectedFirstName = "First"
    testcase.ExpectedSenderName = "First Last"
    Call addTestCaseToArray(testcase)
    
    testcase.originalName = "Last, First"
    testcase.ExpectedFirstName = "First"
    testcase.ExpectedSenderName = testcase.originalName
    Call addTestCaseToArray(testcase)
    
    testcase.originalName = "First Middle Last"
    testcase.ExpectedFirstName = "First"
    testcase.ExpectedSenderName = testcase.originalName
    Call addTestCaseToArray(testcase)
    
    testcase.originalName = "first@example.com"
    testcase.ExpectedFirstName = "First"
    testcase.ExpectedSenderName = testcase.originalName
    Call addTestCaseToArray(testcase)
    
    testcase.originalName = "first.last@example.com"
    testcase.ExpectedFirstName = "First"
    testcase.ExpectedSenderName = testcase.originalName
    Call addTestCaseToArray(testcase)
    
    testcase.originalName = "Dr. First Last"
    testcase.ExpectedFirstName = "First"
    testcase.ExpectedSenderName = testcase.originalName
    Call addTestCaseToArray(testcase)
    
'    testcase.OriginalName = ""
'    testcase.ExpectedFirstName = ""
'    testcase.ExpectedSenderName = testcase.originalName
'    Call addTestCaseToArray(testcase)
    
End Sub

'Runs a single test case
Private Function runTestCase(ByRef testcase As typeTestCase, ByRef curNum As Integer) As Boolean
    Dim firstName As String
    Dim senderName As String
    Call getNamesOutOfString(testcase.originalName, senderName, firstName)
    
    Dim firstNameDiffers As Boolean
    Dim senderNameDiffers As Boolean
    firstNameDiffers = (testcase.ExpectedFirstName <> firstName)
    senderNameDiffers = (testcase.ExpectedSenderName <> senderName)
    
    If firstNameDiffers Or senderNameDiffers Then
        Debug.Print "TestCase " + CStr(curNum) + " failed:"
        
        Dim fiS As String
        If firstNameDiffers Then
          fiS = " <> "
        Else
          fiS = " = "
        End If
        
        Dim srS As String
        If senderNameDiffers Then
          srS = " <> "
        Else
          srS = " = "
        End If
        
        Debug.Print testcase.originalName + ":"
        Debug.Print firstName + fiS + testcase.ExpectedFirstName
        Debug.Print senderName + srS + testcase.ExpectedSenderName
        Debug.Print
        
        'MsgBox "TestCase " + CStr(curNum) + " failed", vbExclamation
        runTestCase = False
    Else
        runTestCase = True
    End If
End Function


Public Sub runTestCaseNo_GetNames(ByVal nIndex As Integer)
    Call initTestCases
    
    ' Array runs from 0 to UBound, but we use entries from 1 to UBound
    If nIndex >= 1 And nIndex <= UBound(mTestCases) Then
         Call runTestCase(mTestCases(nIndex), nIndex)
    End If
End Sub


Public Sub runTests_GetNames()
    Dim i As Integer
    
    Call initTestCases
    
    Dim allSuccessful As Boolean
    allSuccessful = True
    
    For i = 1 To UBound(mTestCases)
        allSuccessful = allSuccessful And runTestCase(mTestCases(i), i)
    Next i

    If Not allSuccessful Then
        MsgBox "At least one testcase failed. See debug output for details", vbExclamation
    End If
End Sub


