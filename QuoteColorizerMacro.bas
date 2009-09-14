Attribute VB_Name = "QuoteColorizerMacro"
Option Explicit


Public Declare Function WriteRTF _
        Lib "mapirtf.dll" _
        Alias "writertf" (ByVal ProfileName As String, _
                          ByVal MessageID As String, _
                          ByVal StoreID As String, _
                          ByVal cText As String) _
        As Integer

Public Declare Function ReadRTF _
        Lib "mapirtf.dll" _
        Alias "readrtf" (ByVal ProfileName As String, _
                         ByVal SrcMsgID As String, _
                         ByVal SrcStoreID As String, _
                         ByRef MsgRTF As String) _
        As Integer


Private Const NUM_RTF_COLORS As Integer = 4

Private Const ENABLE_MACRO As Boolean = True


Public Function ColorizeMailItem(MyMailItem As MailItem) As String
    Dim folder As MAPIFolder
    Dim rtf  As String, lines() As String, resRTF As String
    Dim i As Integer, n As Integer, ret As Integer
  
    
    'save the mailitem to get an entry id, then forget reference to that rtf gets commited.
    'display mailitem by id later on.
    If ((Not MyMailItem.BodyFormat = olFormatPlain) Or (ENABLE_MACRO = False)) Then 'we just understand Plain Mails
        ColorizeMailItem = ""
        Exit Function
    End If
       
    'richt text it
    MyMailItem.BodyFormat = olFormatRichText
    MyMailItem.Save  'need to save to be able to access rtf via EntryID (.save creates ExtryID if not saved before)!
        
    Set folder = Session.GetDefaultFolder(olFolderInbox)
    
    rtf = Space(99999)  'init rtf to max length of message!
    ret = ReadRTF("MAPI", MyMailItem.EntryID, folder.StoreID, rtf)
    If (ret = 0) Then
        'ole call success!!!
        rtf = Trim(rtf)  'kill unnecessary spaces (from rtf var init with Space(rtf))
        Debug.Print rtf & vbCrLf & "*************************************************************" & vbCrLf
        
        'we have out own rtf haeder, remove generated one
        Dim PosHeaderEnd As Integer
        Dim sTestString As String
        PosHeaderEnd = InStr(rtf, "\uc1\pard\plain\deftab360")
        If (PosHeaderEnd = 0) Then
            sTestString = "\uc1\pard\f0\fs20\lang1031"
            PosHeaderEnd = InStr(rtf, sTestString)
        End If
        If (PosHeaderEnd = 0) Then
            sTestString = "\viewkind4\uc1\pard\f0\fs20"
            PosHeaderEnd = InStr(rtf, sTestString)
        End If
        If (PosHeaderEnd = 0) Then
            sTestString = "\pard\f0\fs20\lang1031"
            PosHeaderEnd = InStr(rtf, sTestString)
        End If
        
        rtf = mid(rtf, PosHeaderEnd + Len(sTestString))
        
        rtf = "{\rtf1\ansi\ansicpg1252 \deff0{\fonttbl" & vbCrLf & _
                "{\f0\fswiss\fcharset0 Courier New;}}" & vbCrLf & _
                "{\colortbl\red0\green0\blue0;\red106\green44\blue44;\red44\green106\blue44;\red44\green44\blue106;}" & vbCrLf & _
                rtf
                
        lines = Split(rtf, vbCrLf)
        Dim s As String
        For i = LBound(lines) To UBound(lines)
            n = QuoteFixMacro.CalcNesting(lines(i)).level
            If (n = 0) Then
                resRTF = resRTF & lines(i) & vbCrLf
            Else
                If (Right(lines(i), 4) = "\par") Then
                    s = Left(lines(i), Len(lines(i)) - Len("\par"))
                    resRTF = resRTF & "\cf" & n Mod NUM_RTF_COLORS & " " & s & "\cf0  " & "\par" & vbCrLf
                Else
                    resRTF = resRTF & "\cf" & n Mod NUM_RTF_COLORS & " " & lines(i) & "\cf0  " & vbCrLf
                End If
            End If
        Next i
    Else
        Debug.Print "error while reading rtf! " & ret
        ColorizeMailItem = ""
        Exit Function
    End If
    
    'remove some rtf commands
    resRTF = Replace(resRTF, "\viewkind4\uc1", "")
    resRTF = Replace(resRTF, "\uc1", "")
    'VERY IMPORTANT, outlook will change the message back to PlainText otherwise!!!
    resRTF = Replace(resRTF, "\fromtext", "")
    Debug.Print resRTF
    
       
    'write RTF back to form
    ret = WriteRTF("MAPI", MyMailItem.EntryID, folder.StoreID, resRTF)
    If (ret = 0) Then
        Debug.Print "rtf write okay"
    Else
        Debug.Print "rtf write FAILURE"
        ColorizeMailItem = ""
        Exit Function
    End If
    
    
    'dereference all objects! otherwise, rtf isnt going to be updated!
    Set folder = Nothing
    'save return value
    ColorizeMailItem = MyMailItem.EntryID
    Set MyMailItem = Nothing
End Function


Public Sub DisplayMailItemByID(id As String)
    Dim it As MailItem
    Set it = Session.GetItemFromID(id, Session.GetDefaultFolder(olFolderInbox).StoreID)
    it.Display
    Set it = Nothing
End Sub
