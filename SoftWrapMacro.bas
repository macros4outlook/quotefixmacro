Attribute VB_Name = "SoftWrapMacro"
Option Explicit


Private Const SEVENTY_SIX_CHARS As String = "123456789x123456789x123456789x123456789x123456789x123456789x123456789x123456"
Private Const PIXEL_PER_CHARACTER As Double = 8.61842105263158

Private Const ENABLE_MACRO As Boolean = True


'resize window so that the text editor wraps the text automatically
'after N charaters. Outlook wraps text automatically after sending it,
'but doesnt display the wrap when editing
'you can edit the auto wrap setting at "Tools / Options / Email Format / Internet Format"
Public Sub ResizeWindowForSoftWrap()
    'Application.ActiveInspector.CurrentItem.Body = SEVENTY_SIX_CHARS
    If ((ENABLE_MACRO = True) And _
            (TypeName(Application.ActiveWindow) = "Inspector") And Not _
            (Application.ActiveInspector.WindowState = olMaximized)) Then
            
        Application.ActiveInspector.Width = (LINE_WRAP_AFTER + 2) * PIXEL_PER_CHARACTER
    End If
End Sub
