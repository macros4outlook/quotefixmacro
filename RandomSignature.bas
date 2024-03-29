Attribute VB_Name = "RandomSignature"

'@Folder("QuoteFixMacro")
Option Explicit

Public Sub NewMailMessage()
    ' Creates a new mail message and tacks a random signature onto the end.
    Dim Msg As Outlook.MailItem
    Set Msg = Application.CreateItem(olMailItem)
    
    MakeSig Msg
    Msg.Display
    Set Msg = Nothing
End Sub

Public Sub SwapSig()
    ' Replaces the existing signature with a new randomly chosen one.
    ' Assumes the active window is a compose window.
    If TypeName(Application.ActiveWindow) = "Inspector" Then
        Dim Msg As Outlook.MailItem
        Set Msg = Application.ActiveWindow.CurrentItem
    End If
    
    ' Find the last (if existing) signature delimiter and
    '   remove it and everything below it.
    ' See:  http://en.wikipedia.org/wiki/Signature_block
    Dim strSigStart As String
    strSigStart = InStrRev(Msg.Body, ("--" & vbCrLf))
    If strSigStart <> 0 Then
        Msg.Body = Left$(Msg.Body, strSigStart - 3)
    End If
    
    ' Put a new signature on the message.
    MakeSig Msg
End Sub

Private Sub MakeSig(ByVal Msg As MailItem)
    ' Parses a signature "Fortune-Cookie" file for a fixed, informational
    ' piece that is included with every signature and a quote to be
    ' randomly selected from a list of quotes.  Adds the two pieces
    ' to the end of the passed mail item.
    ' Inspiration from:  http://mozext.achimonline.de/signatureswitch_cookie_signature.php
    '
    ' Fortune-Cookie file location:
    ' %AppData%\Microsoft\Outlook\EmailSigs.txt
    '
    ' Fortune-Cookie Syntax:
    ' Lines are "recorded" from the start of the file.  Delimiters indicate
    '   the end of a quote (or fixed informational line):
    '   $ on a line alone indicates the end of the fixed, informational lines.
    '      Only the last one encountered will be used.
    '   % on a line alone indicates the end of an individual quote.  Any text after the
    '      last "%" (and last "$") will not be included in any signature.
    Dim strFilePath As String
    strFilePath = Environ$("AppData") & "\Microsoft\Outlook\EmailSigs.txt"
    Dim numQuotes As Long
    numQuotes = 0
    Dim strQuote As String
    strQuote = vbNullString
    
    If Len(Dir$(strFilePath)) > 0 Then
        ' Open the file for reading
        Open strFilePath For Input As #1
        
        ' Parse each line in the file
        Dim strLine As String
        Line Input #1, strLine
        
        Do Until EOF(1)
            If Trim$(strLine) = "$" Then
                ' Complete the fixed, informational string.
                Dim strFixedSigPart As String
                strFixedSigPart = vbCrLf & vbCrLf & "--" & strQuote
                strQuote = vbNullString
            ElseIf Trim$(strLine) = "%" Then
                ' Complete a quote and increment the count
                Dim arrQuotes() As String
                ReDim Preserve arrQuotes(0 To numQuotes + 1) As String
                arrQuotes(numQuotes) = strQuote
                numQuotes = numQuotes + 1
                strQuote = vbNullString
            Else
                ' Add another line to the current quote.
                strQuote = strQuote & vbCrLf & strLine
            End If
            Line Input #1, strLine
        Loop
        
        Close #1
    Else
        MsgBox ("Quotes file wasn't found!")
    End If
    
    If numQuotes <> 0 Then
        ' Initialize the RNG seed based on system clock
        Randomize
        
        ' Get the random line number
        Dim intRandom As Long
        intRandom = Int(numQuotes * Rnd())
        
        ' Insert the random quote
        Msg.Body = Msg.Body & strFixedSigPart & arrQuotes(intRandom)
    End If

End Sub
