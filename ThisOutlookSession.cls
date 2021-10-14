Option Explicit
Private WithEvents oExpl As Explorer
Private WithEvents oItem As mailitem

Private Sub Application_Startup()
  Set oExpl = Application.ActiveExplorer
End Sub

Private Sub oExpl_SelectionChange()
  On Error Resume Next
  Set oItem = oExpl.Selection.item(1)
End Sub

Private Sub oItem_Reply(ByVal Response As Object, Cancel As Boolean)
   'QuoteFixMacro will open _a new_ reply, therefore the original one will be _always_ cancelled
   Cancel = True
   Call QuoteFixMacro.FixedReply
End Sub

Private Sub oItem_ReplyAll(ByVal Response As Object, Cancel As Boolean)
   'QuoteFixMacro will open _a new_ reply, therefore the original one will be _always_ cancelled
   Cancel = True
   Call QuoteFixMacro.FixedReplyAll
   'Call QuoteFixMacro.FixedReplyAllEnglish
End Sub

Private Sub oItem_Forward(ByVal Response As Object, Cancel As Boolean)
   'QuoteFixMacro will open _a new_ reply, therefore the original one will be _always_ cancelled
   Cancel = True
   Call QuoteFixMacro.FixedForward
End Sub
