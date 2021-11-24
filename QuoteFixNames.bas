Attribute VB_Name = "QuoteFixNames"

'@Folder("QuoteFixMacro")
Option Explicit
Option Private Module

'Last names may contain formal suffixes.  If they are included here, they can be stripped.
Private Const LASTNAME_SUFFIXES As String = _
"II/III/IV/Jr/Jr./Sr/Sr./Esq."

'Extracts the name of the sender from the sender's name provided in the email.
'TODO: Future work is to extract the first name out of the stored Outlook contacts (if that contact exists)
'
'Notes:
'  * Names are returned by reference
Public Sub getNamesFromMail(ByVal item As MailItem, ByRef senderName As String, ByRef firstName As String, ByRef lastName As String)

    'Wildcard replacements
    senderName = item.SentOnBehalfOfName

    If Len(senderName) = 0 Then
        senderName = item.senderName
    End If

    getNamesOutOfString senderName, senderName, firstName, lastName, item.senderEmailAddress
End Sub

'Code duplication of getNamesFromMail, because there is no common ancestor of MailItem and MeetingItem
Public Sub getNamesFromMeeting(ByVal item As MeetingItem, ByRef senderName As String, ByRef firstName As String, ByRef lastName As String)

    'Wildcard replacements
    ' `MeetingItem.SentOnBehalfOfName` not supported in Outlook 2019.
    ' Thus, we fall back to `senderName`.
    senderName = item.senderName

    If Len(senderName) = 0 Then
        senderName = item.senderName
    End If

    getNamesOutOfString senderName, senderName, firstName, lastName, item.senderEmailAddress
End Sub

'Attempts to extract the name of the sender from the sender's name provided in the email.
'
'Comment:
'   This is very difficult to do definitively.  Consider how many different variations
'   and arrangements a name like "Dr. John James Walker Smith III" could entail including
'   questions of first vs. last names, initials, titles, suffixes, etc..
'
'In:
'  originalName - name as presented by Outlook
'Out:
'  senderName - complete name of sender
'  firstName - first name of sender
'  lastName - last name of sender
'  senderEmailAddress - sender email address (optional because of tests)
'Notes:
'  * Public to enable testing
'  * Names are returned by reference
Public Sub getNamesOutOfString(ByVal originalName As String, ByRef senderName As String, ByRef firstName As String, ByRef lastName As String, Optional ByRef senderEmailAddress As String = vbNullString)
    'Find out firstName

    Dim tmpName As String
    tmpName = originalName

    'cleanup quotes: if name is enclosed in quotes, just remove them
    If (Left$(tmpName, 1) = """" And Right$(tmpName, 1) = """") Then
        tmpName = Mid$(tmpName, 2, Len(tmpName) - 2)
    End If

    'default full senderName: originalName without quotes
    senderName = tmpName

    'default firstName: fullname
    firstName = tmpName

    Dim title As String
    title = vbNullString
    'Has to be later used for extracting the last name

    tmpName = removeDepartment(tmpName)

    If (Left$(tmpName, 4) = "Dr. ") Then
        tmpName = Mid$(tmpName, 5)
        title = "Dr. "
    ElseIf (Right$(tmpName, 5) = ", Dr.") Then
        tmpName = Left$(tmpName, Len(tmpName) - 5)
        title = "Dr. "
    ElseIf (Right$(tmpName, 3) = "Dr.") Then
        tmpName = Left$(tmpName, Len(tmpName) - 3)
        title = "Dr. "
    End If

    'Some companies have "(Text)" at the end of their name.
    'We strip that
    If (Right$(tmpName, 1) = ")") Then
        Dim fPos As Long
        fPos = InStrRev(tmpName, "(")
        If fPos > 0 Then
            tmpName = Trim$(Left$(tmpName, fPos - 1))
        End If
    End If

    fPos = InStr(tmpName, ",")
    If fPos > 0 Then
        'Firstname is separated by comma and positioned behind the lastname
        firstName = Trim$(Mid$(tmpName, fPos + 1))
        'Firstname field may include middle initial(s)
        Do While (UCase$(Right$(firstName, 2)) Like " [A-Z]" Or UCase$(Right$(firstName, 2)) Like "[A-Z].")
            firstName = Trim$(Left$(firstName, Len(firstName) - 2))
        Loop
        lastName = Trim$(Left$(tmpName, fPos - 1))
        'lastName field may have a formal suffix
        lastName = StripSuffixes(lastName)
    Else
        'Determining first and last name is really hard unless
        'there are only two names, or there is a middle initial(s)
        fPos = InStr(Trim$(tmpName), " ")
        If fPos > 0 Then
            'First strip any possible, (single,) formal suffix on the name
            tmpName = StripSuffixes(tmpName)
            Dim lPos As Long
            lPos = InStrRev(Trim$(tmpName), " ")
            If fPos = lPos Then
                'single first name and last name separated by space
                firstName = Trim$(Left$(tmpName, fPos - 1))
                lastName = Trim$(Mid$(tmpName, lPos + 1))
                If firstName = UCase$(firstName) And Not lastName = UCase$(lastName) Then
                    'in case the firstName is written in uppercase letters (and not everything in capital letters),
                    'we assume that the sender's last name is the firstName (in the string)
                    lastName = firstName
                    firstName = Trim$(Mid$(tmpName, lPos + 1))
                End If
            Else
                'middle section could be a single/multiple name/initial (or both)
                Dim midName As String
                midName = Trim$(Mid$(Left$(tmpName, lPos), fPos))

                'One or two initials are easy
                Do While Len(midName) = 1 Or _
                        Left$(midName, 1) = "." Or _
                        Left$(midName, 2) Like "[A-Z] " Or _
                        Left$(midName, 2) Like "[A-Z]."
                    midName = Trim$(Mid$(midName, 2))
                    Dim i As Long
                    i = i + 1
                Loop
                Do While Right$(midName, 2) Like " [A-Z]" Or _
                        Right$(midName, 2) Like "[A-Z]."
                    midName = Trim$(Left$(midName, Len(midName) - 2))
                    Dim j As Long
                    j = j + 1
                Loop

                If Len(midName) = 0 Then
                    'initials only
                    firstName = Trim$(Left$(tmpName, fPos - 1))
                    lastName = Trim$(Mid$(tmpName, lPos + 1))
                ElseIf i <> 0 And j = 0 Then
                    'initials before double last name
                    lastName = midName & Trim$(Mid$(tmpName, lPos + 1))
                    firstName = Trim$(Left$(tmpName, fPos - 1))
                ElseIf i = 0 And j <> 0 Then
                    'initials after double first name
                    lastName = Trim$(Mid$(tmpName, lPos + 1))
                    firstName = Trim$(Left$(tmpName, fPos - 1)) & midName
                ElseIf Left$(midName, 1) = LCase$(Left$(midName, 1)) Then
                    'Midname starts with a lower case letter
                    'We assume "correct" casing. Thus, we hit a name such as Firstname von Lastname
                    firstName = Trim$(Left$(tmpName, fPos - 1))
                    lastName = Trim$(Mid$(tmpName, fPos + 1))
                Else
                    'anything else can't be definitively identified as a first, middle or last name
                    firstName = tmpName
                    lastName = vbNullString
                End If
            End If
        Else
            fPos = InStr(tmpName, "@")
            If fPos > 0 Then
                'first name is (currently) an email address. Just take the prefix
                tmpName = Left$(tmpName, fPos - 1)
            End If
            fPos = InStr(tmpName, ".")
            If fPos > 0 Then
                'first name is separated by a dot
                lastName = Mid$(tmpName, fPos + 1)
                tmpName = Left$(tmpName, fPos - 1)
            Else
                'name is a single string, without "." or " "
                'final guess: LastnameFirstname
                If (IsUpperCaseChar(Left$(tmpName, 1))) Then
                    i = 2
                    Dim UpperCaseCharCount As Long
                    UpperCaseCharCount = 0
                    Dim LastUpperCaseCharPos As Long
                    LastUpperCaseCharPos = 0
                    Do While (i < Len(tmpName) And (UpperCaseCharCount < 2))
                        If (IsUpperCaseChar(Mid$(tmpName, i, 1))) Then
                            LastUpperCaseCharPos = i
                            UpperCaseCharCount = UpperCaseCharCount + 1
                        End If
                        i = i + 1
                    Loop
                    If (UpperCaseCharCount = 1) Then
                        'LastnameFirstname format found
                        tmpName = Mid$(tmpName, LastUpperCaseCharPos)
                    End If
                End If
            End If
            firstName = tmpName
        End If
    End If

    Dim fnEmail As String
    Dim lnEmail As String
    getFirstNameLastNameOutOfEmail senderEmailAddress, fnEmail, lnEmail
    If (LCase$(firstName) = LCase$(lnEmail)) And (LCase$(lastName) = LCase$(fnEmail)) Then
        ' in case firstname and lastname are reversed in the email address, we assume that email format is firstname.lastname and reverse the names here
        Dim tmp As String
        tmp = firstName
        firstName = lastName
        lastName = tmp
    End If

    'fix casing of names
    If InStr(firstName, " ") = 0 Then
        firstName = FixCase(firstName)
    End If
    If InStr(lastName, " ") = 0 Then
        lastName = FixCase(lastName)
    End If
    senderName = title & Trim$(firstName & " " & lastName)
    lastName = title & lastName
End Sub

Public Function removeDepartment(ByVal tmpName As String) As String
    Dim parts() As String
    parts = Split(tmpName, " ")

    Dim length As Long
    length = UBound(parts) - LBound(parts) + 1

    If length <= 2 Then
        removeDepartment = tmpName
        Exit Function
    End If

    Dim indexWordBeforeLastUppercasedWord As Long
    indexWordBeforeLastUppercasedWord = UBound(parts)
    Do While (indexWordBeforeLastUppercasedWord >= LBound(parts) + 2)
        If Not IsUpperCaseWord(parts(indexWordBeforeLastUppercasedWord)) Then
            Exit Do
        End If
        indexWordBeforeLastUppercasedWord = indexWordBeforeLastUppercasedWord - 1
    Loop

    If indexWordBeforeLastUppercasedWord < LBound(parts) + 1 Then
        removeDepartment = tmpName
        Exit Function
    End If

    Dim i As Long
    For i = LBound(parts) To indexWordBeforeLastUppercasedWord
        Dim result As String
        result = result & parts(i) & " "
    Next
    removeDepartment = Left$(result, Len(result) - 1)
End Function

Public Function IsUpperCaseWord(ByVal word As String) As Boolean
    IsUpperCaseWord = word Like "[A-Z][A-Z]*"
End Function

Private Function StripSuffixes(ByVal tempName As String) As String
    'Create array of possible suffixes
    Dim NameSuffixesArr() As String
    NameSuffixesArr = Split(LASTNAME_SUFFIXES, "/")

    'Strip the last suffix (is it ever the case that someone has multiple suffixes?)
    Dim i As Long
    For i = LBound(NameSuffixesArr) To UBound(NameSuffixesArr)
        If (Right$(tempName, Len(NameSuffixesArr(i)) + 1)) = " " & NameSuffixesArr(i) Then
            StripSuffixes = Trim$(Left$(tempName, Len(tempName) - Len(NameSuffixesArr(i))))
        End If
    Next
    StripSuffixes = tempName
End Function

Private Function IsUpperCaseChar(ByVal c As String) As Boolean
    IsUpperCaseChar = c Like "[A-Z]"
End Function

Public Sub getFirstNameLastNameOutOfEmail(ByVal email As String, ByRef firstName As String, ByRef lastName As String)
    If Len(email) = 0 Then
        firstName = vbNullString
        lastName = vbNullString
        Exit Sub
    End If
    Dim parts() As String
    parts = Split(email, "@")
    Dim addressee As String
    addressee = parts(0)
    parts = Split(addressee, ".")

    Dim length As Long
    length = UBound(parts) - LBound(parts) + 1

    If (length <> 2) Then
        firstName = addressee
        lastName = vbNullString
        Exit Sub
    End If
    firstName = parts(LBound(parts))
    firstName = stripNumbers(firstName)
    lastName = parts(UBound(parts))
    lastName = stripNumbers(lastName)
End Sub

Private Function FixCase(ByVal word As String) As String
    If Len(word) = 0 Then
        FixCase = word
        Exit Function
    End If

    Dim parts() As String
    parts = Split(word, "-")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim result As String
        result = result & UCase$(Left$(parts(i), 1)) & LCase$(Mid$(parts(i), 2)) & "-"
    Next

    FixCase = Left$(result, Len(result) - 1)
End Function

Private Function stripNumbers(ByVal s As String) As String
    Dim result As String
    result = s
    Do While (Right$(result, 1) Like "#")
        result = Left$(result, Len(result) - 1)
    Loop
    stripNumbers = result
End Function

