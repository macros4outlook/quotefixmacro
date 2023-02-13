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
   If you don't want to get a security warning when you use the macros, go to "Tools > Macro > Security" and disable the security check.
   A better solution is to sign the macro. See "Signing a Macro" below.
4. File > Import File ... > Select `QuoteFixNames.bas` > Open

Note: You can easily import all files at once by dragging them from the Explorer into the VBA editor and dropping them onto the project tree.

## Configure Outlook to prepare the messages for QuoteFixMacro

1. File > Options > Mail > (scroll down to) Replies and forwards

   * Change the value of "When replying to a message" to "Prefix each line of the original message"
   <!-- markdownlint-disable-next-line MD038 -->
   * Ensure that "Prefix each line in a plain-text message with" contains "`> `"
   * Change the value "When replying to a message" back to "Include original message text"

2. Tools > Options > Mail > (scroll up to) Compose messages

   * Change the value of "Compose messages in this format:" to "Plain Text"

3. Tools > Options > Mail > (scroll down to) Message format

   * "Automatic wrap text at character": 76 characters (which is the default when you did not touch that setting)

4. QuoteFixMacro requires plain text to work.
   It is possible, to read all emails as plain text right from the start.\
   🇺🇸: Navigate to Tools > Options > Trust Center > Trust Center Settings... > Email Security > "Read as Plain Text"
   🇩🇪: Datei > Optionen > Trust Center > E-Mail-Sicherheit > Als Nur-Text lesen\
   🇺🇸: `[X]` Read all standard mail in plain text,\
   🇩🇪: Standardnachrichten im Nur-Text-Format lesen\
   🇺🇸: `[X]` Read all digitally signed mail in plain text".\
   See also Microsoft [KB 831607](https://support.microsoft.com/en-us/office/change-the-message-format-to-html-rich-text-format-or-plain-text-338a389d-11da-47fe-b693-cf41f792fefa?ui=en-us&rs=en-us&ad=us) and ["Read email messages in plain text"](https://support.microsoft.com/en-us/office/read-email-messages-in-plain-text-16dfe54a-fadc-4261-b2ce-19ad072ed7e3?ui=en-US&rs=en-US&ad=US) for another explanation.\
   \
   Note that one can also have QuoteFixMacro converting all emails automatically to text.
   See [Advanced Features](https://macros4outlook.github.io/quotefixmacro/advanced-features.html#auto-conversion-to-plain-format) for details.
   This setting, however, has issues with Outlook 2019.

## Set email signature

QuoteFixMacro needs to know where to put the fixed quote.
For that, it uses the placeholder `%Q` in the signature.

Go to File > Options > Mail > section "Compose messages" > Signatures...

* Create a signature that is only used for reply and forward. You have to insert at least `%Q` to get the quoted original mail.
* Assign this signature to every mail account you want to use.

Alternatively, one can configure `QUOTING_TEMPLATE` in the code (see [Advanced Features](https://macros4outlook.github.io/quotefixmacro/advanced-features.html#configure-the-template-inside-the-code)).

## Assign macros to buttons

After importing the module, you need to replace the original "Reply" and "ReplyAll" buttons with buttons that trigger the macros defined in the file you just imported.
Remember, these buttons are in Outlook's main window, and also in the message window that pops up when you double click on an email.

1. Right-click on the toolbar and select "Customize..."
2. Go to the "Quick Access Toolbar" tab
3. Choose "Macros" at "Choose commands from"
3. Drag the "FixedReply" and "FixedReplyAll" entries and drop it onto the toolbar

You can also change the name and image of the newly created buttons using the customization dialog.
If you use "Fixed&Reply" as the name, <kbd>Alt</kbd>+<kbd>R</kbd> is kept as a shortcut for reply.
Since Outlook does not support custom keybindings, you cannot map the shortcut <kbd>Ctrl</kbd>+<kbd>R</kbd> to the new FixedReply macro.
Nevertheless, the mapping can be done by using AutoHotkey (see below).

## Persist settings across updates

An update of QuoteFixMacro happens by replacing the content of the `.bas` file.
Thus, any settings are overwritten during an update.
QuoteFixMacro can read settings from the registry.
The macro **NEVER** stores entries in the registry by itself.

You can store the default configuration in the registry:

1. by executing `StoreDefaultConfiguration()`
2. by writing a routing executing command similar to the following: `Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "CONVERT_TO_PLAIN", "true")`
3. by manually creating entries in this registry hive: `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\QuoteFixMacro`

<!-- markdownlint-disable-file MD033 -->
