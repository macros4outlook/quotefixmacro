Attribute VB_Name = "Tools"
'$Id$
'
'Tools TRUNK
'contains test routines
'
'Tools is part of the macros4outlook project
'see http://sourceforge.net/projects/macros4outlook/ for more information
'
'For more information on Outlook see http://www.microsoft.com/outlook
'Outlook is (C) by Microsoft

'****************************************************************************
'License:
'
'Tools
'  copyright 2006-2009 Daniel Martin. All rights reserved.
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

Public InterceptorCollection As New Collection


Public Sub MarkMailAsUnread(ByVal MyMail As MailItem)
    MyMail.UnRead = True
End Sub


Public Sub ReadCurrentMailItemRTF()
    Dim rtf As String
    rtf = Space$(99999)
    Dim ret As Long
    ret = ReadRTF("MAPI", GetCurrentItem.EntryID, session.GetDefaultFolder(olFolderInbox).StoreID, rtf)
    rtf = Trim$(rtf)
    
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
    
    answer.bodyFormat = olFormatRichText
    
    Dim ColoredAnswer As String
    'ColoredAnswer = QuoteColorizerMacro.ColorizeMailItem(answer)
    answer.Display
    Set answer = Nothing 'answer bodyformat changes here to 1 for some stupid reason...
    
    'Tools.DisplayMailItemByID mid
End Sub


Public Sub FranksMacro(ByVal CurrentItem As MailItem)
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
    
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    
    Dim s As String
'    s = "test daniel 23e " & vbCrLf & _
'        "> asd asd sad " & vbCrLf & _
'        "> sad asdad as " & vbCrLf & _
'        ">> sa asddsa asd aas kj kj kj k jlkjhlkjhsda asdf asdf adsf as df asdf ads fa dsfa dsf " & vbCrLf & _
'        ">> aasd asdaasdf asd fasdf asd f asd fa sdf adsf asdf saas " & vbCrLf & _
'        "> sasad asda  sasd asd asd asd asd aasdf asdf as df asdf a sd f asd f as df asd fasdf a sdf asdf sdasdasd "

    'par does not work as expected
    '  --> par combines all the lines together and seems to completely ignore the quoting characters
    s = "test daniel 23e " & vbCrLf & _
        "> asd asd sad " & vbCrLf & _
        "> sad asdad as " & vbCrLf & _
        "> > sa asddsa asd aas kj kj kj k jlkjhlkjhsda asdf asdf adsf as df asdf ads fa dsfa dsf " & vbCrLf & _
        "> > aasd asdaasdf asd fasdf asd f asd fa sdf adsf asdf saas " & vbCrLf & _
        "> Nur mit einem Quote " & vbCrLf & _
        "Testtext."
        
    'From the manual of par
    'Result is fine
    Dim s2 As String
    s2 = "Joe Public writes:" & vbCrLf & _
          "> Jane Doe writes:" & vbCrLf & _
          "> >" & vbCrLf & _
          "> >" & vbCrLf & _
          "> > I can't find the source for uncompress." & vbCrLf & _
          "> Oh no, not again!!!" & vbCrLf & _
          ">" & vbCrLf & _
          ">" & vbCrLf & _
          "> Isn't there a FAQ for this?" & vbCrLf & _
          ">" & vbCrLf & _
          ">" & vbCrLf & _
          "That wasn't very helpful, Joe. Jane," & vbCrLf & _
          "just make a link from uncompress to compress."

    Dim cmd As String
    cmd = "C:\cygwin\bin\bash.exe --login -c 'export PARINIT=""rTbgqR B=.,?_A_a Q=_s>|"" ; par 60q'"
    'cmd = "C:\cygwin\bin\bash.exe --login -c 'par 60q'"
    
    Debug.Print cmd
    Dim pipe As Object
    Set pipe = shell.Exec(cmd)
    Debug.Print "END PAR"
    
    pipe.StdIn.Write (s)
    pipe.StdIn.Close
    
    
    Debug.Print "READING..."
    Dim ret As String
    'Do While (pipe.StdOut.AtEndOfStream = False)
    '    ret = ret & pipe.StdOut.ReadLine() & vbCrLf
    'Loop
    ret = pipe.StdOut.ReadAll()
    Debug.Print ret
    
    Set pipe = Nothing
    Set shell = Nothing
End Sub
