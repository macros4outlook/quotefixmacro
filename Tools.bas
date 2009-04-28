Attribute VB_Name = "Tools"
Option Explicit
   
Global InterceptorCollection As New Collection




Public Sub MarkMailAsUnread(MyMail As MailItem)
    MyMail.UnRead = True
End Sub

Public Sub ReadCurrentMailItemRTF()
    Dim rtf As String, ret As Integer
    rtf = Space(99999)
    ret = ReadRTF("MAPI", GetCurrentItem.EntryID, Session.GetDefaultFolder(olFolderInbox).StoreID, rtf)
    rtf = Trim(rtf)
    
    Debug.Print "RTF READ:" & ret & vbCrLf & rtf
End Sub

Public Sub TestColors()
    Dim mi As MailItem
    'Set mi = Session.GetDefaultFolder(olFolderInbox).Items(99)
    Set mi = GetCurrentItem()
    'mi.Display
    
    Dim answer As MailItem
    Set answer = mi.reply
    Set mi = Nothing
    
    answer.BodyFormat = olFormatRichText
    
    Dim mid As String
    'mid = QuoteColorizerMacro.ColorizeMailItem(answer)
    answer.Display
    Set answer = Nothing 'answer bodyformat changes here to 1 for some stupid reason...
    
    'Call Tools.DisplayMailItemByID(mid)
End Sub


Public Sub FranksMacro(CurrentItem As MailItem)
    'put mails with me as the ONLY recipient into one folder, all others into another
    
    'declare mapifolders to move to here...
    
    
    If (CurrentItem.Recipients.count > 1) Then
        'put into "uninteresting" folder...
        'CurrentItem.Move(...)
    Else
        'put into "interesting" folder
        'CurrentItem.Move
    End If
    
End Sub


Public Sub TestPar()
    Dim s As String
    Dim ret As String
    Dim cmd As String
    
    Dim shell As Object
    Dim pipe As Object
    Set shell = CreateObject("WScript.Shell")
    
    s = "test daniel 23e " & vbCrLf & _
        "> asd asd sad " & vbCrLf & _
        "> sad asdad as " & vbCrLf & _
        ">> sa asddsa asd aas kj kj kj k jlkjhlkjhsda asdf asdf adsf as df asdf ads fa dsfa dsf " & vbCrLf & _
        ">> aasd asdaasdf asd fasdf asd f asd fa sdf adsf asdf saas " & vbCrLf & _
        "> sasad asda  sasd asd asd asd asd aasdf asdf as df asdf a sd f asd f as df asd fasdf a sdf asdf sdasdasd "
  
    cmd = "C:\Programme\cygwin\bin\bash.exe --login -c 'export PARINIT=""rTbgqR B=.,?_A_a Q=_s>|"" ; par 60q'"
  
    Debug.Print cmd
    Set pipe = shell.Exec(cmd)
    Debug.Print "END PAR"
    
    pipe.StdIn.Write (s)
    pipe.StdIn.Close
    
    
    Debug.Print "READING..."
    'While (pipe.StdOut.AtEndOfStream = False)
    '    ret = ret + pipe.StdOut.ReadLine() + vbCrLf
    'Wend
    ret = pipe.StdOut.ReadAll()
    Debug.Print ret
    
    Set pipe = Nothing
    Set shell = Nothing
End Sub

