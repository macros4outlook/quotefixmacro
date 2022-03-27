---
nav_order: 2
has_children: true
---
# Setup

## Download

Download the latest version from the [GitHub release page](https://github.com/macros4outlook/quotefixmacro/releases).
You can also download the latest development version using <https://github.com/macros4outlook/quotefixmacro/archive/refs/heads/main.zip>.

## Import macros

1. Extract the downloaded zip-file
2. Open Outlook's VBA editor (<kbd>Alt</kbd>+<kbd>F11</kbd> or "Tools > Macro > Visual Basic-Editor")
3. File > Import File ... > Select `QuoteFixMacro.bas` > Open  
   If you don't want to get a security warning when you use the macros, go to "Tools > Macro > Security" and disable the security check. A better solution is to sign the macro. See "Signing a Macro" below.
4. File > Import File ... > Select `QuoteFixNames.bas` > Open

Note: You can easily import all files at once by dragging them from the Explorer into the VBA editor and dropping them onto the project tree.

## Assign macros to buttons

After importing the module, you need to replace the original "Reply" and "ReplyAll" buttons with buttons that trigger the macros defined in the file you just imported. Remember, these buttons are in Outlook's main window, and also in the message window that pops up when you double click on an email.

1. Right-click on the toolbar and select "Customize..."
2. Go to the "Commands" tab and navigate to the "Macro" category
3. Drag the "FixedReply" and "FixedReplyAll" entries and drop it onto the toolbar

You can also change the name and image of the newly created buttons using the customization dialog. If you use "Fixed&Reply" as the name, <kbd>Alt</kbd>+<kbd>R</kbd> is kept as a shortcut for reply. Since Outlook does not support custom keybindings, you cannot map the shortcut <kbd>Ctrl</kbd>+<kbd>R</kbd> to the new FixedReply macro. Nevertheless, the mapping can be done by using AutoHotkey (see below).

## Set up email

1. Tools > Options > Preferences > E-mail Options... > On replies and forwards

   * Change the value "When replying to a message" to "Prefix each line of the original message"
   <!-- markdownlint-disable-next-line MD038 -->
   * Set "Prefix each line with" to "`> `"
   * Change the value "When replying to a message" back to to "Include original message text"

2. Tools > Options > Mail Format > Internet Format...

   * Automatic wordwrap after: 76 characters (which is the default when you did not touch that setting)

## Read all email as text

QuoteFixMacro requires plain text to work. One can either read all emails as plain text from the beginning.
In case all texts should be read as plain text, see Microsoft [KB 831607](https://support.microsoft.com/en-us/office/change-the-message-format-to-html-rich-text-format-or-plain-text-338a389d-11da-47fe-b693-cf41f792fefa?ui=en-us&rs=en-us&ad=us) for an explanation how to turn on this feature.
For Outlook 2010 and later (also described at ["Read email messages in plain text"](https://support.microsoft.com/en-us/office/read-email-messages-in-plain-text-16dfe54a-fadc-4261-b2ce-19ad072ed7e3?ui=en-US&rs=en-US&ad=US)):

File > Options > Security Center > Options for the Security Center > E-Mail Security > "Read as Plain Text" >  
`[X]` Read all standard mail in plain text,  
`[X]` Read all digitally signed mail in plain text".

Note that one can also have QuoteFixMacro converting all emails automatically to text.
See [Advanced Features](https://macros4outlook.github.io/quotefixmacro/advanced-features.html#auto-conversion-to-plain-format) for details.
This setting, however, has issues with Outlook 2019.

## Persist settings across updates

An update of QuoteFixMacro happens by replacing the content of the `.bas` file.
Thus, any settings are overwritten during an update.
QuoteFixMacro can read settings from the registry.
The macro **NEVER** stores entries in the registry by itself.

You can store the default configuration in the registry:

1. by executing `StoreDefaultConfiguration()`
2. by writing a routing executing command similar to the following: `Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "CONVERT_TO_PLAIN", "true")`
3. by manually creating entries in this registry hive: `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\QuoteFixMacro`
