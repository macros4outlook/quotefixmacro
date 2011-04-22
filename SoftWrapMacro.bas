Attribute VB_Name = "SoftWrapMacro"
'$Id$
'
'SoftWrapMacro TRUNK
'
'SoftWrapMacro is part of the macros4outlook project
'see http://sourceforge.net/projects/macros4outlook/ for more information
'
'For more information on Outlook see http://www.microsoft.com/outlook
'Outlook is (C) by Microsoft

'****************************************************************************
'License:
'
'SoftWrapMacro
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
'Version 1.0 - 2011-04-22
' * first public relese
'
'$Revision$ - not released

Option Explicit

Private Const SEVENTY_SIX_CHARS As String = "123456789x123456789x123456789x123456789x123456789x123456789x123456789x123456"

'This constant has to be adapted to fit your needs (incoprating the used font, display size, ...)
Private Const PIXEL_PER_CHARACTER As Double = 8.61842105263158

'resize window so that the text editor wraps the text automatically
'after N charaters. Outlook wraps text automatically after sending it,
'but doesn't display the wrap when editing
'you can edit the auto wrap setting at "Tools / Options / Email Format / Internet Format"
Public Sub ResizeWindowForSoftWrap()
    'Application.ActiveInspector.CurrentItem.Body = SEVENTY_SIX_CHARS
    If (TypeName(Application.ActiveWindow) = "Inspector") And Not _
        (Application.ActiveInspector.WindowState = olMaximized) Then
            
        Application.ActiveInspector.Width = (LINE_WRAP_AFTER + 2) * PIXEL_PER_CHARACTER
    End If
End Sub
